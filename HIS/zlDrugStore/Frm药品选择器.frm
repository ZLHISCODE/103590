VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form Frm药品选择器 
   Caption         =   "药品选择器"
   ClientHeight    =   6270
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9465
   Icon            =   "Frm药品选择器.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   6270
   ScaleWidth      =   9465
   StartUpPosition =   1  '所有者中心
   Begin VB.CommandButton Cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   8250
      TabIndex        =   5
      Top             =   5850
      Width           =   1100
   End
   Begin VB.CommandButton Cmd确定 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   350
      Left            =   7020
      TabIndex        =   4
      Top             =   5850
      Width           =   1100
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf药品规格 
      Height          =   2805
      Left            =   2640
      TabIndex        =   2
      Top             =   1290
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   4948
      _Version        =   393216
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
   Begin MSComctlLib.ImageList ImgTvw 
      Left            =   2010
      Top             =   1320
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm药品选择器.frx":0E42
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm药品选择器.frx":249C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView Tvw药品用途分类 
      Height          =   4485
      Left            =   0
      TabIndex        =   1
      Top             =   1290
      Width           =   2595
      _ExtentX        =   4577
      _ExtentY        =   7911
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImgTvw"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ImgLvwSmall 
      Left            =   8820
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm药品选择器.frx":41A6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView Lvw 
      Height          =   1275
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9435
      _ExtentX        =   16642
      _ExtentY        =   2249
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      _Version        =   393217
      Icons           =   "ImgLvwSmall"
      SmallIcons      =   "ImgLvwSmall"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   1
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "名称"
         Object.Width           =   2540
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid Msf批次 
      Height          =   1635
      Left            =   2640
      TabIndex        =   3
      Top             =   4140
      Width           =   6795
      _ExtentX        =   11986
      _ExtentY        =   2884
      _Version        =   393216
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
   Begin VB.Image ImgUpDownLvw_S 
      Height          =   45
      Left            =   30
      MousePointer    =   7  'Size N S
      Top             =   1290
      Width           =   9435
   End
   Begin VB.Image ImgLeftRight_S 
      Height          =   4485
      Left            =   2580
      MousePointer    =   9  'Size W E
      Top             =   1290
      Width           =   45
   End
   Begin VB.Image ImgUpDown_S 
      Height          =   45
      Left            =   2640
      MousePointer    =   7  'Size N S
      Top             =   4080
      Width           =   6765
   End
End
Attribute VB_Name = "Frm药品选择器"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--输入参数--
Private IntEditState As Integer                 '编辑状态(1-入库;2-出库)
Private Lng源库房ID As Long                     '源库房ID
Private Lng目库房ID As Long                     '目库房ID
Private Lng使用部门ID As Long                   '使用部门ID
Private lng供应商ID As Long                     '供应商ID
Private OutObj As Form                          '使用本程序的窗体（必须提供一个公共记录集，用以返回）

Private BlnStartUp As Boolean                   '启动成功
Private BlnFirstStart As Boolean                '第一次启动
Private RecUnit As New ADODB.Recordset          '单位
Private StrUnitString As String                 'SQL字串
Private IntStockCheck As Integer                '库存检测
Private bln盘点单 As Boolean                    '盘点单据标志
Private bln空批次 As Boolean                    '是否增加空批次供输入
Private blnCheck As Boolean                     '是否检测库存(盘点用)
Private blnPrice As Boolean                     '是否允许时价或批次药品零出库
Private mstrPreNode As String
Private mstrCaption As String
    
'本程序使用记录集
Private RecData As New ADODB.Recordset          '药品用途分类
Private RecPhysic As New ADODB.Recordset        '药品卡片
Private RecStock As New ADODB.Recordset         '药品规格

'返回记录集
Private RecReturn As ADODB.Recordset            '返回记录集(药品信息所有列,药品目录所有列,药品库存所有列)
Private int库房 As Integer                      '1-药库;2-药房;3-制剂室
Private int分批 As Integer                      '0-不分批;1-药库分批;2-药房分批;3-药库药房分批
Private bln时价 As Boolean                      '时价
Private blnStock  As Boolean
Private StrCardSortBy As String                 '药品卡片排序列
Private StrPhysicSortBy As String               '药品规格排序列
Private LngCardRow As Long
Private LngPhysicRow As Long
Private LngLastSelect药品ID As Long             '上次选择的药品ID（用于是否刷新）
Private mbln中药库房 As Boolean
Private mblnNoStock As Boolean                  '本地参数：是否允许盘点没有设置存储库房的药品
Private int领用方式 As Integer                  '0-向库房领药;1-向科室留存领药
Private mbln包含停用药品 As Boolean

'调用get可用库存后，返回的可用数量，实际数量，实际金额及实际差价
Private mdbl可用数量 As Double
Private mdbl实际数量 As Double
Private mdbl实际金额 As Double
Private mdbl实际差价 As Double
Private mdbl库存数量 As Double

'--公共--
Private Const StrFormat As String = "'999999999990.99999'"

Private mstrCostFormat As String
Private mstrPriceFormat As String
Private mstrNumberFormat As String
Private mstrMoneyFormat As String

Private mintUnit As Integer             '单位系数：1-售价;2-门诊;3-住院;4-药库

'从参数表中取药品价格、数量、金额小数位数
Private mintCostDigit As Integer        '成本价小数位数
Private mintPriceDigit As Integer       '售价小数位数
Private mintNumberDigit As Integer      '数量小数位数
Private mintMoneyDigit As Integer       '金额小数位数

Private Const mconint售价单位 As Integer = 1
Private Const mconint门诊单位 As Integer = 2
Private Const mconint住院单位 As Integer = 3
Private Const mconint药库单位 As Integer = 4

'批次列表
Private Const mconIntCol列数 As Integer = 17
Private Const mconIntColRID As Integer = 0
Private Const mconIntCol库房 As Integer = 1
Private Const mconIntCol批次 As Integer = 2
Private Const mconIntCol入库日期 As Integer = 3
Private Const mconIntCol批号 As Integer = 4
Private Const mconIntCol生产日期 As Integer = 5
Private Const mconIntCol失效期 As Integer = 6
Private Const mconIntCol产地 As Integer = 7
Private Const mconintCol成本价 As Integer = 8
Private Const mconIntCol售价 As Integer = 9
Private Const mconIntCol可用数量 As Integer = 10
Private Const mconintCol库存数量 As Integer = 11
Private Const mconIntCol库存金额 As Integer = 12
Private Const mconIntCol库存差价 As Integer = 13
Private Const mconIntCol上次供应商ID As Integer = 14
Private Const mconIntCol实际数量 As Integer = 15
Private Const mconIntCol批准文号 As Integer = 16

'规格列表
Private Const mconIntColSpec列数 As Integer = 37
Private Const mconIntColSpec剂型 As Integer = 0
Private Const mconIntColSpec药名编码 As Integer = 1
Private Const mconIntColSpec通用名称 As Integer = 2
Private Const mconIntColSpec药品来源 As Integer = 3
Private Const mconIntColSpec药典ID As Integer = 4
Private Const mconIntColSpec用途分类ID As Integer = 5
Private Const mconIntColSpec剂量单位 As Integer = 6
Private Const mconIntColSpec药品编码 As Integer = 7
Private Const mconIntColSpec商品名 As Integer = 8
Private Const mconIntColSpec规格 As Integer = 9
Private Const mconIntColSpec产地 As Integer = 10
Private Const mconIntColSpec药名ID As Integer = 11
Private Const mconIntColSpec药品ID As Integer = 12
Private Const mconIntColSpec上次采购价 As Integer = 13
Private Const mconIntColSpec售价 As Integer = 14
Private Const mconIntColSpec售价单位 As Integer = 15
Private Const mconIntColSpec售价包装 As Integer = 16
Private Const mconIntColSpec门诊单位 As Integer = 17
Private Const mconIntColSpec门诊包装 As Integer = 18
Private Const mconIntColSpec住院单位 As Integer = 19
Private Const mconIntColSpec住院包装 As Integer = 20
Private Const mconIntColSpec药库单位 As Integer = 21
Private Const mconIntColSpec药库包装 As Integer = 22
Private Const mconIntColSpec可用数量 As Integer = 23
Private Const mconIntColSpec库存数量 As Integer = 24
Private Const mconIntColSpec库存金额 As Integer = 25
Private Const mconIntColSpec库存差价 As Integer = 26
Private Const mconIntColSpec有效期 As Integer = 27
Private Const mconIntColSpec药库分批 As Integer = 28
Private Const mconIntColSpec药房分批 As Integer = 29
Private Const mconIntColSpec时价 As Integer = 30
Private Const mconIntColSpec指导批发价 As Integer = 31
Private Const mconIntColSpec指导差价率 As Integer = 32
Private Const mconIntColSpec库房货位 As Integer = 33
Private Const mconIntColSpec批准文号 As Integer = 34
Private Const mconIntColSpec实际数量 As Integer = 35
Private Const mconIntColSpec留存数量 As Integer = 36
Private Sub RestoreColWidth()
    '功能：恢复列宽度
    Dim strType As String
    
    If Val(GetSetting("ZLSOFT", "私有全局\" & gstrDbUser, "使用个性化风格", 1)) = 0 Then Exit Sub
    
    strType = Tvw药品用途分类.SelectedItem.Tag
    
    If strType = "5" Then
        strType = "1"
    ElseIf strType = "6" Then
        strType = "2"
    ElseIf strType = "7" Then
        strType = "3"
    End If
    
    Call RestoreFlexState(Msf药品规格, App.ProductName & Me.Name & strType)
    Call RestoreFlexState(Msf批次, App.ProductName & Me.Name & strType)
    Msf批次.ColWidth(mconIntCol入库日期) = IIf(mstrCaption = "药品移库管理" Or mstrCaption = "药品申领管理", 1000, 0)
End Sub

Private Sub SaveColWidth(ByVal strType As String)
'功能：保存列宽度
'说明：应放在SaveWinState之前,以在不使用个性化时从注册表清除
            
    If Val(GetSetting("ZLSOFT", "私有全局\" & gstrDbUser, "使用个性化风格", 1)) = 0 Then Exit Sub
    If strType = "" And Not Tvw药品用途分类.SelectedItem Is Nothing Then strType = Tvw药品用途分类.SelectedItem.Tag
    
    If strType = "5" Then
        strType = "1"
    ElseIf strType = "6" Then
        strType = "2"
    ElseIf strType = "7" Then
        strType = "3"
    End If
    
    Call SaveFlexState(Msf药品规格, App.ProductName & Me.Name & strType)
    Call SaveFlexState(Msf批次, App.ProductName & Me.Name & strType)
End Sub

Public Property Get In_编辑状态() As Integer
    In_编辑状态 = IntEditState
End Property

Public Property Let In_编辑状态(ByVal vNewValue As Integer)
    IntEditState = vNewValue
End Property

Public Property Get In_源库房() As Long
    In_源库房 = Lng源库房ID
End Property

Public Property Let In_源库房(ByVal vNewValue As Long)
    Lng源库房ID = vNewValue
End Property

Public Property Get In_目库房() As Long
    In_目库房 = Lng目库房ID
End Property

Public Property Let In_目库房(ByVal vNewValue As Long)
    Lng目库房ID = vNewValue
End Property

Public Property Get In_部门() As Long
    In_部门 = Lng使用部门ID
End Property

Public Property Let In_部门(ByVal vNewValue As Long)
    Lng使用部门ID = vNewValue
End Property

Public Property Let In_MainFrm(ByVal vNewValue As Form)
    Set OutObj = vNewValue
End Property

Private Sub SetFormat(Optional ByVal IntMain As Integer = 1, Optional ByVal BlnSetHeader As Boolean = False)
    Dim intCol As Integer
    
    '设置各列表控件的格式
    Select Case IntMain
    Case 1
        With Msf药品规格
            
            If BlnSetHeader Then
                .Cols = IIf(int领用方式 = 0, mconIntColSpec列数 - 1, mconIntColSpec列数)
                '卡片
                .TextMatrix(0, mconIntColSpec剂型) = "剂型"
                .TextMatrix(0, mconIntColSpec药名编码) = "药名编码"
                .TextMatrix(0, mconIntColSpec通用名称) = "通用名称"
                .TextMatrix(0, mconIntColSpec药品来源) = "药品来源"
                .TextMatrix(0, mconIntColSpec药典ID) = "药典ID"
                .TextMatrix(0, mconIntColSpec用途分类ID) = "用途分类ID"
                .TextMatrix(0, mconIntColSpec剂量单位) = "剂量单位"
                
                '规格
                .TextMatrix(0, mconIntColSpec药品编码) = "药品编码"
                .TextMatrix(0, mconIntColSpec商品名) = "商品名"
                .TextMatrix(0, mconIntColSpec规格) = "规格"
                .TextMatrix(0, mconIntColSpec产地) = "产地"
                .TextMatrix(0, mconIntColSpec药名ID) = "药名ID"
                .TextMatrix(0, mconIntColSpec药品ID) = "药品ID"
                .TextMatrix(0, mconIntColSpec上次采购价) = "上次采购价"
                .TextMatrix(0, mconIntColSpec售价) = "售价"
                .TextMatrix(0, mconIntColSpec售价单位) = "售价单位"
                .TextMatrix(0, mconIntColSpec售价包装) = "售价包装"
                .TextMatrix(0, mconIntColSpec门诊单位) = "门诊单位"
                .TextMatrix(0, mconIntColSpec门诊包装) = "门诊包装"
                .TextMatrix(0, mconIntColSpec住院单位) = "住院单位"
                .TextMatrix(0, mconIntColSpec住院包装) = "住院包装"
                .TextMatrix(0, mconIntColSpec药库单位) = "药库单位"
                .TextMatrix(0, mconIntColSpec药库包装) = "药库包装"
                .TextMatrix(0, mconIntColSpec可用数量) = "可用数量"
                .TextMatrix(0, mconIntColSpec库存数量) = "库存数量"
                .TextMatrix(0, mconIntColSpec库存金额) = "库存金额"
                .TextMatrix(0, mconIntColSpec库存差价) = "库存差价"
                .TextMatrix(0, mconIntColSpec有效期) = "有效期"
                .TextMatrix(0, mconIntColSpec药库分批) = "药库分批"
                .TextMatrix(0, mconIntColSpec药房分批) = "药房分批"
                .TextMatrix(0, mconIntColSpec时价) = "时价"
                .TextMatrix(0, mconIntColSpec指导批发价) = "指导批发价"
                .TextMatrix(0, mconIntColSpec指导差价率) = "指导差价率"
                .TextMatrix(0, mconIntColSpec库房货位) = "库房货位"
                .TextMatrix(0, mconIntColSpec批准文号) = "批准文号"
                .TextMatrix(0, mconIntColSpec实际数量) = "实际数量"
                If int领用方式 = 1 Then
                    .TextMatrix(0, mconIntColSpec留存数量) = "留存数量"
                End If
            End If
            
            For intCol = 0 To .Cols - 1
                .ColAlignmentFixed(intCol) = 4
            Next
            
            .ColAlignment(mconIntColSpec上次采购价) = 7
            .ColAlignment(mconIntColSpec售价) = 7
            .ColAlignment(mconIntColSpec售价包装) = 7
            .ColAlignment(mconIntColSpec门诊包装) = 7
            .ColAlignment(mconIntColSpec住院包装) = 7
            .ColAlignment(mconIntColSpec药库包装) = 7
            .ColAlignment(mconIntColSpec可用数量) = 7
            .ColAlignment(mconIntColSpec库存数量) = 7
            .ColAlignment(mconIntColSpec库存金额) = 7
            .ColAlignment(mconIntColSpec库存差价) = 7
            .ColAlignment(mconIntColSpec有效期) = 7
            .ColAlignment(mconIntColSpec时价) = 7
            .ColAlignment(mconIntColSpec指导批发价) = 7
            .ColAlignment(mconIntColSpec指导差价率) = 7
            .ColAlignment(mconIntColSpec实际数量) = 7
            If int领用方式 = 1 Then
                .ColAlignment(mconIntColSpec留存数量) = 7
            End If
            
            If BlnStartUp = False Then
                .ColWidth(mconIntColSpec剂型) = 500
                .ColWidth(mconIntColSpec药名编码) = 0
                .ColWidth(mconIntColSpec通用名称) = 0
                .ColWidth(mconIntColSpec药典ID) = 0
                .ColWidth(mconIntColSpec用途分类ID) = 0
                .ColWidth(mconIntColSpec剂量单位) = 0
                '规格
                .ColWidth(mconIntColSpec药品编码) = 1000
                .ColWidth(mconIntColSpec商品名) = 1800
                .ColWidth(mconIntColSpec规格) = 1000
                .ColWidth(mconIntColSpec产地) = 1200
                .ColWidth(mconIntColSpec药名ID) = 0
                .ColWidth(mconIntColSpec药品ID) = 0
                .ColWidth(mconIntColSpec售价) = 1200
                .ColWidth(mconIntColSpec可用数量) = 1200
                .ColWidth(mconIntColSpec库存数量) = 0
                .ColWidth(mconIntColSpec库存金额) = 0
                .ColWidth(mconIntColSpec库存差价) = 0
                .ColWidth(mconIntColSpec有效期) = 900
                .ColWidth(mconIntColSpec药库分批) = 900
                .ColWidth(mconIntColSpec药房分批) = 900
                .ColWidth(mconIntColSpec时价) = 900
                .ColWidth(mconIntColSpec指导批发价) = 0
                .ColWidth(mconIntColSpec指导差价率) = 0
                .ColWidth(mconIntColSpec库房货位) = 1500
                .ColWidth(mconIntColSpec批准文号) = 1000
                .ColWidth(mconIntColSpec实际数量) = 0
                If int领用方式 = 1 Then
                    .ColWidth(mconIntColSpec留存数量) = 1000
                End If
                .Row = 1
                
                .ColWidth(mconIntColSpec售价单位) = IIf(mintUnit = mconint售价单位, 900, 0)
                .ColWidth(mconIntColSpec售价包装) = IIf(mintUnit = mconint售价单位, 900, 0)
                .ColWidth(mconIntColSpec门诊单位) = IIf(mintUnit = mconint门诊单位, 900, 0)
                .ColWidth(mconIntColSpec门诊包装) = IIf(mintUnit = mconint门诊单位, 900, 0)
                .ColWidth(mconIntColSpec住院单位) = IIf(mintUnit = mconint住院单位, 900, 0)
                .ColWidth(mconIntColSpec住院包装) = IIf(mintUnit = mconint住院单位, 900, 0)
                .ColWidth(mconIntColSpec药库单位) = IIf(mintUnit = mconint药库单位, 900, 0)
                .ColWidth(mconIntColSpec药库包装) = IIf(mintUnit = mconint药库单位, 900, 0)
                .ColWidth(mconIntColSpec上次采购价) = IIf(mstrCaption = "药品外购入库管理", 1200, 0)
            End If
        End With
    Case 0
        With Msf批次
            
            If BlnSetHeader Then
                .Cols = mconIntCol列数
                .TextMatrix(0, mconIntColRID) = "RID"
                .TextMatrix(0, mconIntCol库房) = "库房"
                .TextMatrix(0, mconIntCol批次) = "批次"
                .TextMatrix(0, mconIntCol入库日期) = "入库日期"
                .TextMatrix(0, mconIntCol批号) = "批号"
                .TextMatrix(0, mconIntCol生产日期) = "生产日期"
                .TextMatrix(0, mconIntCol失效期) = IIf(gtype_UserSysParms.P149_效期显示方式 = 1, "有效期至", "失效期")
                .TextMatrix(0, mconIntCol产地) = "产地"
                .TextMatrix(0, mconintCol成本价) = "成本价"
                .TextMatrix(0, mconIntCol售价) = "售价"
                .TextMatrix(0, mconIntCol可用数量) = "可用数量"
                .TextMatrix(0, mconintCol库存数量) = "库存数量"
                .TextMatrix(0, mconIntCol库存金额) = "库存金额"
                .TextMatrix(0, mconIntCol库存差价) = "库存差价"
                .TextMatrix(0, mconIntCol上次供应商ID) = "上次供应商ID"
                .TextMatrix(0, mconIntCol实际数量) = "实际数量"
                .TextMatrix(0, mconIntCol批准文号) = "批准文号"
            End If
            
            For intCol = 0 To .Cols - 1
                .ColAlignmentFixed(intCol) = 4
            Next
            .ColWidth(mconIntColRID) = 0
            .ColAlignment(mconIntCol产地) = 7
            .ColAlignment(mconintCol成本价) = 7
            .ColAlignment(mconIntCol售价) = 7
            .ColAlignment(mconIntCol可用数量) = 7
            .ColAlignment(mconintCol库存数量) = 7
            .ColAlignment(mconIntCol库存金额) = 7
            .ColAlignment(mconIntCol库存差价) = 7
            
            If BlnStartUp = False Then
                .ColWidth(mconIntColRID) = 0
                .ColWidth(mconIntCol库房) = 1200
                .ColWidth(mconIntCol批次) = 0
                .ColWidth(mconIntCol批号) = 1000
                .ColWidth(mconIntCol入库日期) = IIf(mstrCaption = "药品移库管理" Or mstrCaption = "药品申领管理", 1000, 0)
                .ColWidth(mconIntCol生产日期) = 1000
                .ColWidth(mconIntCol失效期) = 1000
                .ColWidth(mconIntCol产地) = 1200
                .ColWidth(mconintCol成本价) = 1200
                .ColWidth(mconIntCol售价) = 1200
                .ColWidth(mconIntCol可用数量) = 1200
                .ColWidth(mconintCol库存数量) = 1200
                .ColWidth(mconIntCol库存金额) = 1200
                .ColWidth(mconIntCol库存差价) = 1200
                .ColWidth(mconIntCol上次供应商ID) = 0
                .ColWidth(mconIntCol实际数量) = 0
                .ColWidth(mconIntCol批准文号) = 1000
                .Row = 1
            End If
        End With
    End Select
End Sub

Private Sub Cmd取消_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub Cmd确定_Click()
    Dim blnValid As Boolean
    On Error Resume Next
    
    If In_编辑状态 = 2 Then If CheckData = False Then Exit Sub
    '检查分批属性与库存数据是否一致
    If In_编辑状态 = 2 Then
        blnValid = 检查库存数据(Lng源库房ID, LngLastSelect药品ID)
    Else
        blnValid = 检查库存数据(Lng目库房ID, LngLastSelect药品ID)
    End If
    If Not blnValid Then
        MsgBox "发现该药品在当前库房中的库存记录存在错误（可能是基础数据设置错误，请检查当前库房的部门性质及该药品的分批属性）！", vbInformation, gstrSysName
        Exit Sub
    End If
    '组装记录集
    If CombinateRec = False Then Exit Sub
    
    Unload Me
    Exit Sub
End Sub

Private Sub Form_Activate()
    If BlnStartUp = False Then
        Unload Me
        Exit Sub
    End If

End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    
    BlnStartUp = False
    BlnFirstStart = False
    mstrPreNode = ""
    
    '取售价单位
    StrUnitString = ""
    IntStockCheck = 0
    LngLastSelect药品ID = 0
    Msf批次.Visible = (In_编辑状态 = 2)
    
    '初始化记录集
    InitRec
    
    If OutObj Is Nothing Then
        MsgBox "请指定主窗体！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '初始化并检测相关数据完整性
    If ReadAndSendDataToTvw() = False Then Exit Sub
    
    '提取当前库存控制参数
    gstrSQL = "Select Nvl(检查方式,0) 库存检查 From 药品出库检查 Where 库房ID=[1]"
    Set RecUnit = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Lng源库房ID)
    
    With RecUnit
        If Not .EOF Then
            IntStockCheck = !库存检查
        End If
    End With
    
    '检查源库房是否为药库
    If Lng源库房ID <> 0 Then
        int库房 = 3
        
        gstrSQL = "select 部门ID from 部门性质说明 where (工作性质 like '%药房' Or 工作性质 like '%制剂室') And 部门id=[1]"
        Set RecUnit = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Lng源库房ID)
        
        If RecUnit.EOF Then
            RecUnit.Close
            
            gstrSQL = "select 部门ID from 部门性质说明 where 工作性质 like '%药库' And 部门id=[1]"
            Set RecUnit = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Lng源库房ID)
            
            If Not RecUnit.EOF Then int库房 = 1
        Else
            int库房 = 2
        End If

    End If
    
    mstrCaption = GetText(GetParentWindow(OutObj.hWnd))
    If mstrCaption = "药品申领管理" Then
        Call GetDrugDigit(Lng使用部门ID, mstrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
        Lng使用部门ID = 0
    ElseIf mstrCaption = "药品移库管理" Then
        Call GetDrugDigit(Lng使用部门ID, mstrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
        Lng使用部门ID = 0
    Else
        Call GetDrugDigit(IIf(Lng源库房ID = 0, Lng目库房ID, Lng源库房ID), mstrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    End If
    
    mstrCostFormat = "'999999999990." & String(mintCostDigit, "0") & "'"
    mstrPriceFormat = "'999999999990." & String(mintPriceDigit, "0") & "'"
    mstrNumberFormat = "'999999999990." & String(mintNumberDigit, "0") & "'"
    mstrMoneyFormat = "'999999999990." & String(mintMoneyDigit, "0") & "'"

    Select Case mintUnit
        Case mconint门诊单位
            StrUnitString = "/nvl(门诊包装,1)"
        Case mconint住院单位
            StrUnitString = "/nvl(住院包装,1)"
        Case mconint药库单位
            StrUnitString = "/nvl(药库包装,1)"
    End Select
    
    Tvw药品用途分类_NodeClick Tvw药品用途分类.SelectedItem
    
    BlnStartUp = True
End Sub

Private Function ReadAndSendDataToTvw() As Boolean
    Dim NodeThis As node, ItemThis As ListItem
    Dim Int末级 As Integer
    Dim lng库房ID As Long
    Dim rs材质分类 As New ADODB.Recordset
    
    '药品用途分类是否有数据
    ReadAndSendDataToTvw = False

    gstrSQL = " Select 编码,名称 From 诊疗项目类别 " & _
              " Where Instr([1],编码,1) > 0 " & _
              " Order by 编码"
    Set rs材质分类 = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, "567")
    
    With Lvw
        .ListItems.Clear
    End With
    With Tvw药品用途分类
        .Nodes.Clear
        Do While Not rs材质分类.EOF
            .Nodes.Add , , "Root" & rs材质分类!名称, rs材质分类!名称, 1, 1
            .Nodes("Root" & rs材质分类!名称).Tag = rs材质分类!编码
            rs材质分类.MoveNext
        Loop
    End With
    
    '如果是入库，以入库库房为准，否则以出库库房为准
    If IntEditState = 1 Then
        lng库房ID = Lng目库房ID
    Else
        lng库房ID = Lng源库房ID
    End If
    
    mbln中药库房 = False
    If lng库房ID <> 0 Then
        '提取该库房现有剂型，供用户选择
        gstrSQL = "Select 1 From 部门性质说明 " & _
                 " Where 工作性质 Like '中药%' And 部门ID=[1]"
        Set RecData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[检查部门性质]", lng库房ID)
        
        If Not RecData.EOF Then mbln中药库房 = True
        gstrSQL = "Select Distinct J.编码,J.名称 " & _
                 " From 诊疗执行科室 A,药品特性 B,药品剂型 J " & _
                 " Where A.诊疗项目ID=B.药名ID And B.药品剂型=J.名称" & _
                 " And A.执行科室ID=[1]"
        Set RecData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[提取该库房现在剂型]", lng库房ID)
    Else
        gstrSQL = "Select 编码,名称 From 药品剂型"
        Call zlDatabase.OpenRecordset(RecData, gstrSQL, "提取所有药品剂型")
    End If
    
    With RecData
        Lvw.ListItems.Clear
        Do While Not .EOF
            Lvw.ListItems.Add , "K" & !编码, !名称, 1, 1
            .MoveNext
        Loop
        
        If .State = 1 Then .Close
        gstrSQL = "Select ID,上级ID,名称,1 as 末级,decode(类型,1,'西成药',2,'中成药','中草药') as 材质,类型  " & _
                 " From 诊疗分类目录  " & _
                 " where 类型 in (1,2,3)  " & _
                 " Start With 上级ID IS NULL Connect By Prior ID=上级ID Order by level,ID"
        
        Call SQLTest(App.Title, Me.Caption, gstrSQL)
        .Open gstrSQL, gcnOracle
        Call SQLTest
        
        If .EOF Then
            MsgBox "请初始化药品用途分类（药品用途分类）！", vbInformation, gstrSysName
            Exit Function
        End If
        
        '将药品用途分类数据装入
        Do While Not .EOF
            Int末级 = IIf(!末级 = 1, 2, 1)
            If IsNull(!上级ID) Then
                Set NodeThis = Tvw药品用途分类.Nodes.Add("Root" & !材质, 4, "K_" & !Id, !名称, Int末级, Int末级)
            Else
                Set NodeThis = Tvw药品用途分类.Nodes.Add("K_" & !上级ID, 4, "K_" & !Id, !名称, Int末级, Int末级)
            End If
            NodeThis.Tag = !类型   '存放分类类型:1-西成药,2-中成药,3-中草药
            .MoveNext
        Loop
    End With
    
    With Tvw药品用途分类
        .Nodes(1).Selected = True
        If .Nodes(1).Children <> 0 Then
            Int末级 = 1
            .Nodes(Int末级).Child.Selected = True
            .SelectedItem.Selected = True
        ElseIf .Nodes(2).Children <> 0 Then
            Int末级 = 2
            .Nodes(Int末级).Child.Selected = True
            .SelectedItem.Selected = True
        ElseIf .Nodes(3).Children <> 0 Then
            Int末级 = 3
            .Nodes(Int末级).Child.Selected = True
            .SelectedItem.Selected = True
        Else
            Int末级 = 0
            .Nodes(1).Selected = True
            .SelectedItem.Selected = True
        End If
        If Int末级 <> 0 Then .Nodes(Int末级).Expanded = True
    End With
    
    ReadAndSendDataToTvw = True
End Function

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    
    BlnFirstStart = True
    If Me.Height < 5835 Then Me.Height = 5835
    If Me.Width < 8415 Then Me.Width = 8415
    
    With ImgUpDownLvw_S
        .Left = 0
        .Width = Me.ScaleWidth
    End With
    With Me.Lvw
        .Width = Me.ScaleWidth
        .Height = ImgUpDownLvw_S.Top
    End With
    
    With ImgLeftRight_S
        .Top = ImgUpDownLvw_S.Top + ImgUpDownLvw_S.Height
        .Height = Me.ScaleHeight - 200 - Cmd取消.Height - .Top
    End With
    With Tvw药品用途分类
        .Top = ImgUpDownLvw_S.Top + ImgUpDownLvw_S.Height
        .Height = ImgLeftRight_S.Height
        .Width = ImgLeftRight_S.Left
    End With
    
    With ImgUpDown_S
        .Left = ImgLeftRight_S.Left + ImgLeftRight_S.Width
        .Width = Me.ScaleWidth - .Left
    End With
    With Msf药品规格
        .Left = ImgUpDown_S.Left
        .Top = ImgLeftRight_S.Top
        .Width = ImgUpDown_S.Width
    End With
    With Msf批次
        If .Visible Then
            .Top = ImgUpDown_S.Top + ImgUpDown_S.Height
            .Height = ImgLeftRight_S.Top + ImgLeftRight_S.Height - .Top
            .Left = Msf药品规格.Left
            .Width = Msf药品规格.Width
        End If
    End With
    
    With Cmd取消
        .Top = Tvw药品用途分类.Top + Tvw药品用途分类.Height + 150
        .Left = Me.ScaleWidth - .Width - 150
    End With
    With Cmd确定
        .Top = Cmd取消.Top
        .Left = Cmd取消.Left - .Width - 100
    End With
    
    With Msf药品规格
        .Height = IIf(Msf批次.Visible = False, Tvw药品用途分类.Top + Tvw药品用途分类.Height - .Top, Msf批次.Top - 45 - .Top)
        If .RowIsVisible(.Row) = False Then .TopRow = .TopRow + IIf(.Row > .TopRow, 1, -1)
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveColWidth("")
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub ImgLeftRight_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    
    With ImgLeftRight_S
        If .Left + x < 2500 Then Exit Sub
        If .Left + x > Me.ScaleWidth - 4500 Then Exit Sub
        
        .Move .Left + x
    End With
    
    Form_Resize
End Sub

Private Sub ImgUpDown_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    
    With ImgUpDown_S
        If .Top + y < 2500 Then Exit Sub
        If .Top + y > Me.ScaleHeight - 2500 Then Exit Sub
        
        .Move .Left, .Top + y
    End With
    
    Form_Resize
End Sub

Private Sub ImgUpDownLvw_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    
    With ImgUpDownLvw_S
        If .Top + y > 2500 Or .Top + y < 1300 Then Exit Sub
        .Move .Left, .Top + y
    End With
    
    Form_Resize
End Sub

Private Sub Lvw_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Tvw药品用途分类_NodeClick Tvw药品用途分类.SelectedItem
End Sub

Private Sub Msf批次_Click()
    Dim StrHeader As String
    Dim intCol As Integer
    '实现列排序
    On Error Resume Next
    With Msf批次
        If .MouseRow <> 0 Then Exit Sub
        If RecStock.EOF Then Exit Sub
        
        StrHeader = .TextMatrix(0, .MouseCol)
        Set .DataSource = Nothing
        If Mid(StrPhysicSortBy, 2) = StrHeader Then
            StrPhysicSortBy = IIf(Mid(StrPhysicSortBy, 1, 1) = "A", "D", "A") & .TextMatrix(0, .MouseCol)
            RecStock.Sort = .TextMatrix(0, .MouseCol) & IIf(Mid(StrPhysicSortBy, 1, 1) = "A", " Desc", " Asc")
        Else
            StrPhysicSortBy = "A" & .TextMatrix(0, .MouseCol)
            RecStock.Sort = .TextMatrix(0, .MouseCol) & " Asc"
        End If
        Set .DataSource = RecStock

        For intCol = 0 To .Cols - 1
            .ColAlignmentFixed(intCol) = 4
        Next
        
        Call SetFormat(0, False)
    End With
End Sub

Private Sub Msf批次_DblClick()
    On Error Resume Next
    If Cmd确定.Enabled = False Then Exit Sub
    
    With RecStock
        If .RecordCount <> 0 Then .MoveFirst
        If .EOF Then Exit Sub
        If .RecordCount = 0 Then Exit Sub
    End With
    Call Cmd确定_Click
End Sub

Private Sub Msf批次_EnterCell()
    Dim intCol As Integer, LngSelectRow As Long
    Dim RecGetPrice As New ADODB.Recordset
    Dim Lng收费细目ID As Long
    On Error Resume Next
    
    With Msf批次
        .Redraw = False
        
        LngSelectRow = .Row     '保存当前选中行
        If LngPhysicRow <> 0 Then
            .Row = IIf(LngPhysicRow > .Rows - 1, 0, LngPhysicRow)       '清除上次选中行
            For intCol = 0 To .Cols - 1
                    .Col = intCol
                    .CellBackColor = &H80000005
                    .CellForeColor = &H80000008
            Next
            .Col = 0
        End If
        
        LngPhysicRow = LngSelectRow
        .Row = LngPhysicRow     '设置当前选中行
        For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = &H8000000D
                .CellForeColor = &H80000005
        Next
        .Col = 0
        
        .Redraw = True
    End With
End Sub

Private Sub Msf批次_GotFocus()
    With Msf批次
        .GridColorFixed = &H80000008
        .GridColor = &H80000008
    End With
End Sub

Private Sub Msf批次_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Msf批次_DblClick
End Sub

Private Sub Msf批次_LostFocus()
    With Msf批次
        .GridColorFixed = &H80000011
        .GridColor = &H80000011
    End With
End Sub

Private Sub Msf药品规格_Click()
    Dim StrHeader As String
    Dim intCol As Integer
    
    '实现列排序
    On Error Resume Next
    With Msf药品规格
        If .MouseRow <> 0 Then Exit Sub
        If RecPhysic.EOF Then Exit Sub
        
        StrHeader = .TextMatrix(0, .MouseCol)
        Set .DataSource = Nothing
        If Mid(StrCardSortBy, 2) = StrHeader Then
            StrCardSortBy = IIf(Mid(StrCardSortBy, 1, 1) = "A", "D", "A") & .TextMatrix(0, .MouseCol)
            RecPhysic.Sort = .TextMatrix(0, .MouseCol) & IIf(Mid(StrCardSortBy, 1, 1) = "A", " Desc", " Asc")
        Else
            StrCardSortBy = "A" & .TextMatrix(0, .MouseCol)
            RecPhysic.Sort = .TextMatrix(0, .MouseCol) & " Asc"
        End If
        Set .DataSource = RecPhysic
        
        For intCol = 0 To .Cols - 1
            .ColAlignmentFixed(intCol) = 4
        Next
        
        Call SetFormat(1, False)
    End With
End Sub

Private Sub Msf药品规格_DblClick()
    If RecPhysic.EOF Then Exit Sub
    If RecPhysic.RecordCount = 0 Then Exit Sub
    
    If Cmd确定.Enabled Then
        Cmd确定_Click
    Else
        MsgBox "该药品没有库存，不能继续操作！", vbInformation, gstrSysName
    End If
End Sub

Private Sub Msf药品规格_EnterCell()
    Dim Lng收费细目ID As Long, intCol As Integer, LngSelectRow As Long
    Dim StrTmp As String, RecGetPrice As New ADODB.Recordset
    Dim strSql效期 As String
    Dim str售价 As String
     
    With Msf药品规格
        .Redraw = False
        
        LngSelectRow = .Row     '保存当前选中行
        If LngCardRow <> 0 Then
            .Row = IIf(LngCardRow > .Rows - 1, 0, LngCardRow)       '清除上次选中行
            For intCol = 0 To .Cols - 1
                    .Col = intCol
                    .CellBackColor = &H80000005
                    .CellForeColor = &H80000008
            Next
            .Col = 0
        End If
        
        LngCardRow = LngSelectRow
        .Row = LngCardRow       '设置当前选中行
        For intCol = 0 To .Cols - 1
                .Col = intCol
                .CellBackColor = &H8000000D
                .CellForeColor = &H80000005
        Next
        .Col = 0
        .Redraw = True
        
        '如果该规格药品的价格到执行时间还未执行,则触发
        Lng收费细目ID = Val(.TextMatrix(.Row, mconIntColSpec药品ID))
        If Lng收费细目ID = 0 Then
            If Msf批次.Visible Then
                Msf批次.Clear
                Msf批次.Rows = 2
                Call SetFormat(0, True)
                Msf批次_EnterCell
            Else
                Call SetFormat(0, True)
            End If
            Exit Sub
        End If
        
        If LngLastSelect药品ID = Lng收费细目ID Then Exit Sub
        LngLastSelect药品ID = Lng收费细目ID
        
        '如果已到执行日期而价格未执行，执行计算过程
        gstrSQL = " Select ID From 收费价目 Where 收费细目ID=[1] And 变动原因=0"
        Set RecGetPrice = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Lng收费细目ID)
        
        With RecGetPrice
            If Not .EOF Then
                If Not IsNull(!Id) Then
                    Lng收费细目ID = !Id
                    gstrSQL = "zl_药品收发记录_Adjust(" & Lng收费细目ID & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption & "-产生药品价格调整记录")
                End If
            End If
        End With
    End With
    
    If In_编辑状态 = 2 Then
        Msf批次.Visible = False
        '读出该药品规格下所有的药品批次库存信息
        bln时价 = (Msf药品规格.TextMatrix(Msf药品规格.Row, mconIntColSpec时价) = "是")
        int分批 = 0
        str售价 = Msf药品规格.TextMatrix(Msf药品规格.Row, mconIntColSpec售价)
        If Msf药品规格.TextMatrix(Msf药品规格.Row, mconIntColSpec药库分批) = "是" Or Msf药品规格.TextMatrix(Msf药品规格.Row, mconIntColSpec药房分批) = "是" Then
            If Msf药品规格.TextMatrix(Msf药品规格.Row, mconIntColSpec药库分批) = "是" And Msf药品规格.TextMatrix(Msf药品规格.Row, mconIntColSpec药房分批) = "是" Then
                int分批 = 3
            ElseIf Msf药品规格.TextMatrix(Msf药品规格.Row, mconIntColSpec药库分批) = "是" Then
                int分批 = 1
            Else
                int分批 = 2
            End If
        End If
        If Not ((int分批 = 3 And int库房 <> 3) Or (int分批 = 1 And int库房 = 1) Or (int分批 = 2 And int库房 = 2)) Then '如果该药品不分批
            Msf批次.Visible = False
            Form_Resize
        Else
            If Msf批次.Visible = False Then Msf批次.Visible = True
        End If
        Form_Resize
        
        With RecStock
            If .State = 1 Then .Close
            gstrSQL = ""
            If bln空批次 Then
                strSql效期 = IIf(gtype_UserSysParms.P149_效期显示方式 = 1, "有效期至", "失效期")
                gstrSQL = "Select 1 RID,名称,0 批次,'' 入库日期,'新增批次药品' 批号,NULL 生产日期,sysdate " & strSql效期 & ",'' 产地,'' 成本价,''售价," & _
                          "'' 可用数量,'' 库存数量,'' 库存金额,'' 库存差价,0 上次供应商ID,'' 实际数量,'' 批准文号 " & _
                          " From 部门表" & _
                          " Where ID=[1] " & _
                          " Union "
            End If
            
            strSql效期 = IIf(gtype_UserSysParms.P149_效期显示方式 = 1, "K.效期-1 As 有效期至", "K.效期 As 失效期")
            gstrSQL = gstrSQL & " Select 2 RID,P.名称 库房,K.批次,TO_CHAR(S.审核日期, 'YYYY-MM-DD') As 入库日期,K.上次批号 批号,To_Char(K.上次生产日期,'YYYY-MM-DD') 生产日期," & strSql效期 & ",K.上次产地 产地,"
            If blnStock Then
                Select Case mintUnit
                Case mconint售价单位
                    StrTmp = " To_Char(K.上次采购价," & mstrNumberFormat & ") 成本价, " & _
                             IIf(bln时价 = True, " Decode(Sign(K.实际数量),1,To_Char(K.实际金额/K.实际数量," & mstrNumberFormat & "),'" & str售价 & "') 售价, ", " '" & str售价 & "' 售价, ") & _
                             " To_Char(K.可用数量," & mstrNumberFormat & ") 可用数量," & _
                             " To_Char(K.实际数量," & mstrNumberFormat & ") 库存数量,"
                Case mconint门诊单位
                    StrTmp = " To_Char(K.上次采购价*nvl(D.门诊包装,1)," & mstrNumberFormat & ") 成本价, " & _
                             IIf(bln时价 = True, " Decode(Sign(K.实际数量),1,To_Char(K.实际金额/K.实际数量*nvl(D.门诊包装,1)," & mstrNumberFormat & "),'" & str售价 & "') 售价, ", " '" & str售价 & "' 售价, ") & _
                             " To_Char(K.可用数量" & StrUnitString & "," & mstrNumberFormat & ") 可用数量," & _
                             " To_Char(K.实际数量" & StrUnitString & "," & mstrNumberFormat & ") 库存数量,"
                Case mconint住院单位
                    StrTmp = " To_Char(K.上次采购价*nvl(D.住院包装,1)," & mstrNumberFormat & ") 成本价, " & _
                             IIf(bln时价 = True, " Decode(Sign(K.实际数量),1,To_Char(K.实际金额/K.实际数量*nvl(D.住院包装,1)," & mstrNumberFormat & "),'" & str售价 & "') 售价, ", " '" & str售价 & "' 售价, ") & _
                             " To_Char(K.可用数量" & StrUnitString & "," & mstrNumberFormat & ") 可用数量," & _
                             " To_Char(K.实际数量" & StrUnitString & "," & mstrNumberFormat & ") 库存数量,"
                Case mconint药库单位
                    StrTmp = " To_Char(K.上次采购价*nvl(D.药库包装,1)," & mstrNumberFormat & ") 成本价, " & _
                             IIf(bln时价 = True, " Decode(Sign(K.实际数量),1,To_Char(K.实际金额/K.实际数量*nvl(D.药库包装,1)," & mstrNumberFormat & "),'" & str售价 & "') 售价, ", " '" & str售价 & "' 售价, ") & _
                             " To_Char(K.可用数量" & StrUnitString & "," & mstrNumberFormat & ") 可用数量," & _
                             " To_Char(K.实际数量" & StrUnitString & "," & mstrNumberFormat & ") 库存数量,"
                End Select
            Else
                StrTmp = "'' 可用数量,'' 库存数量,"
            End If
            
            gstrSQL = gstrSQL & StrTmp & IIf(blnStock, " To_Char(K.实际金额," & mstrMoneyFormat & ") 库存金额,", "'' 库存金额,") & _
                     IIf(blnStock, " To_Char(K.实际差价," & mstrMoneyFormat & ") 库存差价", "'' 库存差价") & _
                     " ,NVL(K.上次供应商id,0) 上次供应商id,To_Char(K.实际数量," & mstrNumberFormat & ") AS 实际数量,K.批准文号 " & _
                     " From 部门表 P,药品规格 D,药品库存 K,药品收发记录 S" & _
                     " Where K.库房ID = P.ID And D.药品ID = K.药品ID And K.库房ID=[2] " & _
                     " And K.药品ID=[3] And K.性质=1 And Decode(Nvl(K.批次,0),0,-999,K.批次)=S.Id(+) "
            If bln盘点单 Then
                gstrSQL = gstrSQL & " And (K.实际数量<>0 Or K.实际金额<>0 Or K.实际差价<>0)"
            ElseIf glngModul <> 1303 Then   '如果是库存差价调整模块，则允许过滤库存数量为0的药品记录
                gstrSQL = gstrSQL & " And K.实际数量<>0 "
            End If
            
            If gtype_UserSysParms.P150_药品出库优先算法 = 0 Then
                gstrSQL = gstrSQL & " Order by RID,批次"
            Else
                gstrSQL = gstrSQL & " Order by RID," & IIf(gtype_UserSysParms.P149_效期显示方式 = 1, "有效期至", "失效期") & ",批次"
            End If
        End With
       
        Set RecStock = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Lng目库房ID, IIf(Lng源库房ID = 0, Lng目库房ID, Lng源库房ID), LngLastSelect药品ID)
        
        Dim BlnState As Boolean
        With Msf批次
            If Not RecStock.EOF Then
                Set .DataSource = RecStock
                .ColWidth(mconIntColRID) = 0
            Else
                .Clear
                .Rows = 2
            End If
            
            Call SetFormat(0, RecStock.EOF)
            If bln空批次 And RecStock.RecordCount <> 0 Then .Row = IIf(.Rows > 2, 2, 1)
            BlnState = ((int分批 = 3 And int库房 <> 3) Or (int分批 = 1 And int库房 = 1) Or (int分批 = 2 And int库房 = 2)) And Not RecStock.EOF
            .Visible = BlnState
            Msf批次_EnterCell
        End With
        Form_Resize
    End If
    
    '设置按钮状态
    With RecPhysic
        If .RecordCount <> 0 Then .MoveFirst
        .Find "药品ID=" & Val(Msf药品规格.TextMatrix(Msf药品规格.Row, mconIntColSpec药品ID))
        If .EOF Then
            MsgBox "发生内部错误！", vbInformation, gstrSysName
            Exit Sub
        End If
        
        If In_编辑状态 = 2 And ((int分批 = 3 And int库房 <> 3) Or (int分批 = 1 And int库房 = 1) Or (int分批 = 2 And int库房 = 2)) And blnPrice Then
            Cmd确定.Enabled = BlnState
        Else
            Cmd确定.Enabled = True
        End If
    End With
End Sub

Private Sub Msf药品规格_GotFocus()
    With Msf药品规格
        .GridColorFixed = &H80000008
        .GridColor = &H80000008
    End With
End Sub

Private Sub Msf药品规格_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Msf药品规格_DblClick
End Sub

Private Sub Msf药品规格_LostFocus()
    With Msf药品规格
        .GridColorFixed = &H80000011
        .GridColor = &H80000011
    End With
End Sub

Private Sub Tvw药品用途分类_NodeClick(ByVal node As MSComctlLib.node)
    Dim StrTmp As String, StrGroupBy As String
    Dim str单位转换串 As String
    Dim str显示留存 As String
    
    On Error GoTo ErrHand
    '读出该药品用途分类、属于指定剂型的规格药品
'    如果目标库房不明确（如其他出库和领用）或是制剂室，则不限制药品材质
'    如果目标库房是西药库或西药房，则西成药可以进入；
'    如果目标库房是成药库或成药房，则中成药可以进入；
'    如果目标库房是中药库或中药房，则中草药可以进入；
'
'    如果目标库房不明确（如其他出库和领用）或是药库、制剂室，则不限制服务对象
'    如果目标库房是服务于门诊病人，则门诊用药可以进入；
'    如果目标库房是服务于住院病人，则住院用药可以进入；
    
            
    If node.Key = mstrPreNode Then Exit Sub
    If Visible Then
        Call SaveColWidth(Tvw药品用途分类.Nodes(mstrPreNode).Tag)
    End If
    mstrPreNode = node.Key
    
    str显示留存 = IIf(int领用方式 = 1, ",To_Char(S.留存数量 ," & mstrNumberFormat & ") 留存数量", "")

    Select Case mintUnit
        Case mconint售价单位
            str单位转换串 = "*1"
        Case mconint门诊单位
            str单位转换串 = "*D.门诊包装"
        Case mconint住院单位
            str单位转换串 = "*D.住院包装"
        Case mconint药库单位
            str单位转换串 = "*D.药库包装"
    End Select

    With RecPhysic
        If .State = 1 Then .Close
        
        '对列头排顺序
        gstrSQL = " Select D.剂型,D.药名编码,D.通用名称,D.药品来源,D.药典ID,D.用途分类ID,D.剂量单位,D.药品编码,D.商品名,D.规格," & IIf(IntEditState = 1, "D.产地", "Nvl(D.产地,S.产地)") & " AS 产地," & _
                " D.药名ID,D.药品ID,trim(to_char(D.初始成本价" & str单位转换串 & "," & mstrCostFormat & ")) As 上次采购价,trim(to_char(P.售价" & str单位转换串 & ", " & mstrPriceFormat & ")) As 售价," & _
                " D.售价单位,D.剂量系数,D.门诊单位,D.门诊包装,D.住院单位,D.住院包装,D.药库单位,D.药库包装," & _
                IIf(blnStock, " To_Char(S.可用数量 " & StrUnitString & " ," & mstrNumberFormat & ") 可用数量,To_Char(S.库存数量 " & StrUnitString & "," & mstrNumberFormat & ") 库存数量,S.库存金额,S.库存差价,", "'' 可用数量,'' 库存数量,'' 库存金额,'' 库存差价,") & _
                " D.最大效期 有效期,D.药库分批,D.药房分批,D.时价,D.初始成本价,D.指导批发价,D.指导差价率,E.库房货位,D.批准文号,To_Char(S.库存数量 ," & mstrNumberFormat & ") 实际数量" & str显示留存 & _
                " From"
        '药品信息，药品目录
        gstrSQL = gstrSQL & " (SELECT DISTINCT J.名称 剂型,C.编码 药名编码,C.名称 AS 通用名称,0 AS 药典ID,M.分类ID AS 用途分类ID,M.计算单位 AS 剂量单位,C.编码 AS 药品编码," & _
                " " & IIf(mblnTradeName, "NVL(A.名称,C.名称)", "C.名称") & " 商品名,C.规格,C.产地,D.药品来源,D.批准文号, D.药名ID,D.药品ID, C.计算单位 AS 售价单位," & _
                " To_Char(D.剂量系数," & StrFormat & " ) 剂量系数,nvl(To_Char(D.最大效期,'9999990'),0) 最大效期," & _
                " DECODE(D.药库分批,1,'是','否') 药库分批,DECODE(D.药房分批,1,'是','否') 药房分批,DECODE(C.是否变价,1,'是','否') 时价," & _
                " D.门诊单位,To_Char(D.门诊包装," & StrFormat & " ) 门诊包装,D.住院单位," & _
                " To_Char(D.住院包装," & StrFormat & " ) 住院包装,D.药库单位,To_Char(D.药库包装," & StrFormat & " ) 药库包装," & _
                " To_Char(D.指导批发价," & mstrCostFormat & ") 指导批发价,nvl(D.成本价,0) 初始成本价,To_Char(D.指导差价率," & StrFormat & " ) 指导差价率" & _
                " FROM 收费项目目录 C,药品规格 D,收费项目别名 A,药品剂型 J,药品特性 T,诊疗项目目录 M," & _
                "             (Select 执行科室ID,收费细目ID From 收费执行科室 Where 执行科室ID" & IIf(Lng源库房ID <> 0, "=[1]", IIf(Lng目库房ID <> 0, "=[2]", " Is Not NULL")) & " Group By 执行科室ID,收费细目ID) K," & _
                "             (Select 执行科室ID,收费细目ID From 收费执行科室 Where 执行科室ID" & IIf(Lng目库房ID <> 0, "=[2]", IIf(Lng源库房ID <> 0, "=[1]", " Is Not NULL")) & " Group By 执行科室ID,收费细目ID) I " & _
                " WHERE C.ID=D.药品ID AND D.药名ID=T.药名ID AND T.药名ID=M.ID AND M.类别 IN ('5','6','7')" & _
                " AND D.药品ID=K.收费细目ID" & IIf(mblnNoStock, "(+)", "") & " " & _
                " And D.药品ID=I.收费细目ID" & IIf(mblnNoStock, "(+)", "") & " " & _
                " AND D.药品ID=A.收费细目ID(+) AND A.性质(+)=3 And (C.站点 = '" & gstrNodeNo & "' Or C.站点 is Null) " & _
                " AND T.药品剂型=J.名称(+)"
                'IIf(Lng使用部门ID <> 0, " And K.开单科室ID=I.开单科室ID And K.开单科室ID=" & Lng使用部门ID, "")
        gstrSQL = gstrSQL & "" & _
            " and ([2] is null" & _
                " or exists(select 1 from 部门性质说明 where 工作性质='制剂室' and 部门id=[2])" & _
                " or C.类别=(select distinct '5' from 部门性质说明 where 工作性质 like '西药%' and 部门id=[2])" & _
                " or C.类别=(select distinct '6' from 部门性质说明 where 工作性质 like '成药%' and 部门id=[2])" & _
                " or C.类别=(select distinct '7' from 部门性质说明 where 工作性质 like '中药%' and 部门id=[2]) Or [2]=0)" & _
            " and ([2] is null" & _
                " or exists(select 1 from 部门性质说明 where 工作性质 like '%药库' and 部门id=[2])" & _
                " or exists(select 1 from 部门性质说明 where 工作性质='制剂室' and 部门id=[2])" & _
                " or decode(C.服务对象,1,1,3,1,0)=(select distinct '1' from 部门性质说明 where 工作性质 like '%药房' and 部门id=[2] and 服务对象 in(1,3))" & _
                " or decode(C.服务对象,2,1,3,1,0)=(select distinct '1' from 部门性质说明 where 工作性质 like '%药房' and 部门id=[2] and 服务对象 in(2,3)) Or [2]=0)"
        
        '查找指定药品用途分类的规格药品
        If Not (node.Key Like "Root*") Then
            gstrSQL = gstrSQL & _
                    " And M.分类ID IN " & _
                    "     (Select ID from 诊疗分类目录 " & _
                    "     Where 类型 In (1,2,3)" & _
                    "     Start With ID=" & Mid(node.Key, 3) & _
                    "     Connect By Prior ID=上级ID)"
        Else
            gstrSQL = gstrSQL & " And M.类别='" & node.Tag & "' "
        End If
        '只查找未停用的规格药品（需要根据传入参数决定，暂时只有盘点时该参数才可能为True）
        If mbln包含停用药品 = False Then
            gstrSQL = gstrSQL & " And (C.撤档时间 Is Null Or To_char(C.撤档时间,'yyyy-MM-dd')='3000-01-01')"
        End If
        '只查找指定剂型的规格药品
        gstrSQL = gstrSQL & Get剂型SQL
        '只查找指定材质分类的规格药品
        gstrSQL = gstrSQL & " ) D,"
        
        '提取所有药品的当前售价
        gstrSQL = gstrSQL & " (Select 收费细目id,To_Char(现价," & mstrPriceFormat & ") 售价 From 收费价目 Where (Sysdate Between 执行日期 And 终止日期 or Sysdate>=执行日期 And 终止日期 Is Null)) P,"
        '提取所有药品的当前售价
        If blnStock Then
            If int领用方式 = 1 Then
                gstrSQL = gstrSQL & " (Select a.药品id,Max(上次产地) AS 产地,To_Char(Sum(a.可用数量),'99999999999990.99999') 可用数量," & _
                        " To_Char(Sum(a.实际数量),'99999999999990.99999') 库存数量," & _
                        " To_Char(Sum(a.实际金额),'99999999999990.99999') 库存金额," & _
                        " To_Char(Sum(a.实际差价),'99999999999990.99999') 库存差价," & _
                        " To_Char(Sum(b.实际数量),'99999999999990.99') 留存数量" & _
                        " From 药品库存 a ,药品留存 b Where a.性质=1 and a.药品id=b.药品id And a.库房id =b.库房id and b.科室id=[3] and b.期间=[4] "
            Else
                gstrSQL = gstrSQL & " (Select a.药品id,Max(上次产地) AS 产地,To_Char(Sum(a.可用数量),'99999999999990.99999') 可用数量," & _
                        " To_Char(Sum(a.实际数量),'99999999999990.99999') 库存数量," & _
                        " To_Char(Sum(a.实际金额),'99999999999990.99999') 库存金额," & _
                        " To_Char(Sum(a.实际差价),'99999999999990.99999') 库存差价" & _
                        " From 药品库存 a Where a.性质=1 "
            End If
        Else
            gstrSQL = gstrSQL & " (Select 药品id,' ' 产地, '' 可用数量," & _
                    " '' 库存数量,'' 库存金额,'' 库存差价" & _
                    " From 药品库存 a Where 性质=1 "
        End If
'        If lng供应商ID <> 0 Then gstrSQL = gstrSQL & " And (上次供应商ID Is Null Or 上次供应商ID=" & lng供应商ID & ")"
        If Lng源库房ID <> 0 Or Lng目库房ID <> 0 Then
            gstrSQL = gstrSQL & " And a.库房ID=" & IIf(Lng源库房ID = 0, "[2]", "[1]") & "  Group By a.药品id) S"
        Else
            gstrSQL = gstrSQL & " Group By a.药品id) S"
        End If
        gstrSQL = gstrSQL & ",(Select 药品ID,库房ID,库房货位 From 药品储备限额 " & _
                  " Where 库房ID=" & IIf(IntEditState = 2, "[1]", "[2]") & ") E"
        
        '总条件
        gstrSQL = gstrSQL & " Where D.药品ID=P.收费细目ID And D.药品ID=S.药品ID"
        '当系统参数“药品出库库存检查”为不足禁止时，不提库存为零
        If Not (IntStockCheck = 2 And In_编辑状态 = 2) Or bln盘点单 Or Not blnCheck Then gstrSQL = gstrSQL & "(+) "
        'If In_编辑状态 = 2 Then gstrSQL = gstrSQL & " And S.可用数量<>0"
        gstrSQL = gstrSQL & " And D.药品ID=E.药品ID(+) Order By D.药名编码,D.药品编码"
    End With
    
    Set RecPhysic = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Lng源库房ID, Lng目库房ID, Lng使用部门ID, Format(zlDatabase.Currentdate(), "yyyy"))
    
    With Msf药品规格
        If Not RecPhysic.EOF Then
            Set .DataSource = RecPhysic
        Else
            .Clear
            .Rows = 2
            LngLastSelect药品ID = 0
        End If
        Call SetFormat(1, RecPhysic.EOF)
    End With
    Cmd确定.Enabled = (RecPhysic.EOF <> True)
    
    Call Msf药品规格_EnterCell
    Call RestoreColWidth
     
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then Resume
End Sub

Private Function Get剂型SQL() As String
    Dim ItemThis As ListItem, strReturn As String
    '返回获取剂型的SQL
    strReturn = ""
    Get剂型SQL = ""
    
    For Each ItemThis In Lvw.ListItems
        If ItemThis.Checked Then
            strReturn = strReturn & ",'" & ItemThis.Text & "'"
        End If
    Next
    If mbln中药库房 And strReturn <> "" Then strReturn = strReturn & ",'方剂'"
    
    If strReturn = "" Then Exit Function
    strReturn = Mid(strReturn, 2)
    Get剂型SQL = " And T.药品剂型 In (" & strReturn & ")"
End Function

Private Function InitRec()
        '编制人:朱玉宝
        '编制日期:2000-11-02
        '初始化记录集
        
        Set RecReturn = New ADODB.Recordset
        With RecReturn
            If .State = 1 Then .Close
            .Fields.Append "剂型", adLongVarChar, 20, adFldIsNullable
            .Fields.Append "药名编码", adLongVarChar, 8, adFldIsNullable
            .Fields.Append "药品来源", adLongVarChar, 8, adFldIsNullable
            .Fields.Append "通用名称", adLongVarChar, 40, adFldIsNullable
            .Fields.Append "药典ID", adDouble, 18, adFldIsNullable
            .Fields.Append "用途分类ID", adDouble, 18, adFldIsNullable
            .Fields.Append "剂量单位", adLongVarChar, 40, adFldIsNullable
            .Fields.Append "药品编码", adLongVarChar, 10, adFldIsNullable
            .Fields.Append "商品名", adLongVarChar, 40, adFldIsNullable
            .Fields.Append "规格", adLongVarChar, 30, adFldIsNullable
            .Fields.Append "产地", adLongVarChar, 40, adFldIsNullable
            .Fields.Append "药名ID", adDouble, 18, adFldIsNullable
            .Fields.Append "药品ID", adDouble, 18, adFldIsNullable
            .Fields.Append "售价", adDouble, 18, adFldIsNullable
            .Fields.Append "售价单位", adLongVarChar, 8, adFldIsNullable
            .Fields.Append "剂量系数", adDouble, 11, adFldIsNullable
            .Fields.Append "最大效期", adDouble, 5, adFldIsNullable
            .Fields.Append "门诊单位", adLongVarChar, 8, adFldIsNullable
            .Fields.Append "门诊包装", adDouble, 11, adFldIsNullable
            .Fields.Append "住院单位", adLongVarChar, 8, adFldIsNullable
            .Fields.Append "住院包装", adDouble, 11, adFldIsNullable
            .Fields.Append "药库单位", adLongVarChar, 8, adFldIsNullable
            .Fields.Append "药库包装", adDouble, 11, adFldIsNullable
            .Fields.Append "药库分批", adDouble, 2, adFldIsNullable
            .Fields.Append "药房分批", adDouble, 2, adFldIsNullable
            .Fields.Append "时价", adDouble, 2, adFldIsNullable
            .Fields.Append "批次", adLongVarChar, 8, adFldIsNullable
            .Fields.Append "批号", adLongVarChar, 8, adFldIsNullable
            .Fields.Append "生产日期", adDate, , adFldIsNullable
            .Fields.Append "效期", adDate, , adFldIsNullable
            .Fields.Append "可用数量", adLongVarChar, 8, adFldIsNullable
            .Fields.Append "实际数量", adLongVarChar, 8, adFldIsNullable
            .Fields.Append "实际金额", adLongVarChar, 8, adFldIsNullable
            .Fields.Append "实际差价", adLongVarChar, 8, adFldIsNullable
            .Fields.Append "指导批发价", adDouble, 11, adFldIsNullable
            .Fields.Append "指导差价率", adDouble, 11, adFldIsNullable
            .Fields.Append "上次供应商ID", adDouble, 18, adFldIsNullable
            .Fields.Append "库存数量", adLongVarChar, 8, adFldIsNullable
            .Fields.Append "批准文号", adLongVarChar, 40, adFldIsNullable
            .Fields.Append "成本价", adDouble, 11, adFldIsNullable
            
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open
        End With
End Function

Private Function CombinateRec() As Boolean
    '组装记录集
    '定位记录集
    Dim blnEof As Boolean               '是否存在批次库存
    Dim dblPrice As Double
    Dim rsTemp As New ADODB.Recordset
    Dim strMsg As String
    
    CombinateRec = False
    With RecPhysic
        If .RecordCount <> 0 Then .MoveFirst
        .Find "药品ID=" & Val(Msf药品规格.TextMatrix(Msf药品规格.Row, mconIntColSpec药品ID))
        If .EOF Then
            MsgBox "发生内部错误！", vbInformation, gstrSysName
            Exit Function
        End If
        
        If ((int分批 = 3 And int库房 <> 3) Or (int分批 = 1 And int库房 = 1) Or (int分批 = 2 And int库房 = 2)) And In_编辑状态 = 2 Then
            With RecStock
                If .RecordCount <> 0 Then .MoveFirst
                .Find "批次=" & Val(Msf批次.TextMatrix(Msf批次.Row, mconIntCol批次))
                If .EOF Then
                    blnEof = True
                    If blnPrice Then
                        MsgBox "发生内部错误！", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
            End With
        End If
    End With
    
    '提取该药品的零售单位价格
    gstrSQL = "Select 现价, B.指导批发价, B.指导零售价 " & _
            " From 收费价目 A, 药品规格 B " & _
            " Where A.收费细目id = B.药品id And A.收费细目ID=[1] And Sysdate Between A.执行日期 And Nvl(A.终止日期,Sysdate) "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[提取该药品的零售单位价格]", CLng(RecPhysic!药品ID))
    
    dblPrice = 0
    If Not rsTemp.EOF Then
        dblPrice = Nvl(rsTemp!现价, 0)
    End If
    
    '检查指导批发价，指导零售价，为0时不允许对该药品操作
    strMsg = ""
    If Not rsTemp.EOF Then
        If rsTemp!指导批发价 = 0 And rsTemp!指导零售价 = 0 Then
            strMsg = "[" & RecPhysic!药名编码 & RecPhysic!通用名称 & "]采购限价和指导售价为0，请先设置价格。"
        ElseIf rsTemp!指导批发价 = 0 Then
            strMsg = "[" & RecPhysic!药名编码 & RecPhysic!通用名称 & "]采购限价为0，请先设置价格。"
        ElseIf rsTemp!指导零售价 = 0 Then
            strMsg = "[" & RecPhysic!药名编码 & RecPhysic!通用名称 & "]指导售价为0，请先设置价格。"
        End If
    End If
    If strMsg <> "" Then
        MsgBox strMsg, vbInformation, gstrSysName
        CombinateRec = False
        Exit Function
    End If
    
'    '如果是定价药品，则现价必须大于0，否则不允许对该药品操作
'    If IIf(RecPhysic!时价 = "是", 1, 0) = 0 And dblPrice = 0 Then
'        MsgBox "[" & RecPhysic!药名编码 & RecPhysic!通用名称 & "]是定价药品，请先设置零售价。", vbInformation, gstrSysName
'        CombinateRec = False
'        Exit Function
'    End If
    
    '装数据写入记录集，供其它窗体使用
    With RecReturn
        If .EOF Then .AddNew
        !剂型 = RecPhysic!剂型
        !药名编码 = RecPhysic!药名编码
        !药品来源 = RecPhysic!药品来源
        !通用名称 = RecPhysic!通用名称
        !药典ID = RecPhysic!药典ID
        !用途分类ID = RecPhysic!用途分类ID
        !剂量单位 = RecPhysic!剂量单位
        !药品编码 = RecPhysic!药品编码
        !商品名 = RecPhysic!商品名
        !规格 = RecPhysic!规格
        !产地 = RecPhysic!产地
        !药名ID = RecPhysic!药名ID
        !药品ID = RecPhysic!药品ID
        !售价 = dblPrice
        !售价单位 = RecPhysic!售价单位
        !剂量系数 = RecPhysic!剂量系数
        !最大效期 = RecPhysic!有效期
        !门诊单位 = RecPhysic!门诊单位
        !门诊包装 = RecPhysic!门诊包装
        !住院单位 = RecPhysic!住院单位
        !住院包装 = RecPhysic!住院包装
        !药库单位 = RecPhysic!药库单位
        !药库包装 = RecPhysic!药库包装
        !药库分批 = IIf(RecPhysic!药库分批 = "是", 1, 0)
        !药房分批 = IIf(RecPhysic!药房分批 = "是", 1, 0)
        !时价 = IIf(RecPhysic!时价 = "是", 1, 0)
        !上次供应商ID = 0
        !批准文号 = IIf(IsNull(RecPhysic!批准文号), "", RecPhysic!批准文号)
        If In_编辑状态 = 2 And ((int分批 = 3 And int库房 <> 3) Or (int分批 = 1 And int库房 = 1) Or (int分批 = 2 And int库房 = 2)) Then
            If Msf批次.TextMatrix(Msf批次.Row, mconIntCol批号) = "新增批次药品" Then
                !批次 = -1
            Else
                If Not blnEof Then
                    !批次 = Val(RecStock!批次)
                    !批号 = RecStock!批号
                    !生产日期 = RecStock!生产日期
                    If gtype_UserSysParms.P149_效期显示方式 = 0 Then
                        !效期 = RecStock!失效期
                    Else
                        !效期 = RecStock!有效期至
                    End If
                    !产地 = Nvl(RecStock!产地)
                    !上次供应商ID = Nvl(RecStock!上次供应商ID, 0)
                    !可用数量 = IIf(IsNull(RecStock!可用数量), 0, RecStock!可用数量)
                    !实际数量 = IIf(IsNull(RecStock!库存数量), 0, RecStock!库存数量)
                    !实际金额 = IIf(IsNull(RecStock!库存金额), 0, RecStock!库存金额)
                    !实际差价 = IIf(IsNull(RecStock!库存差价), 0, RecStock!库存差价)
                    !库存数量 = IIf(IsNull(RecStock!实际数量), 0, RecStock!实际数量)
                    !批准文号 = IIf(IsNull(RecStock!批准文号), "", RecStock!批准文号)
                    If Not blnStock Then Call Get可用库存(!药品ID, !批次)
                End If
            End If
        Else
            !可用数量 = IIf(IsNull(RecPhysic!可用数量), 0, RecPhysic!可用数量)
            !实际数量 = IIf(IsNull(RecPhysic!库存数量), 0, RecPhysic!库存数量)
            !实际金额 = IIf(IsNull(RecPhysic!库存金额), 0, RecPhysic!库存金额)
            !实际差价 = IIf(IsNull(RecPhysic!库存差价), 0, RecPhysic!库存差价)
            !库存数量 = IIf(IsNull(RecPhysic!实际数量), 0, RecPhysic!实际数量)
            
            '提取不分批药品的批号与效期信息
            gstrSQL = "Select 上次批号,效期,上次供应商id,上次生产日期 AS 生产日期,批准文号 From 药品库存 " & _
                     " Where 库房ID=[1] And 药品ID=[2] And 性质=1 "
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[提取不分批药品的批号与效期信息]", Lng源库房ID, CLng(RecPhysic!药品ID))
            
            If rsTemp.RecordCount <> 0 Then
                !批号 = Nvl(rsTemp!上次批号)
                If Not IsNull(rsTemp!生产日期) Then
                    !生产日期 = Nvl(rsTemp!生产日期)
                End If
                !上次供应商ID = Nvl(rsTemp!上次供应商ID, 0)
                If Not IsNull(rsTemp!效期) Then
                    !效期 = rsTemp!效期
                    If gtype_UserSysParms.P149_效期显示方式 = 1 And Nvl(!效期) <> "" Then
                        '换算为有效期
                        !效期 = Format(DateAdd("D", -1, !效期), "yyyy-mm-dd")
                    End If
                End If
                !批准文号 = IIf(IsNull(rsTemp!批准文号), "", rsTemp!批准文号)
            End If
            
            If Not blnStock Then Call Get可用库存(!药品ID, 0)
        End If
        
        '如果不显示对方库房的库存，需重新提取并更新
        If Not blnStock Then
            !可用数量 = mdbl可用数量
            !实际数量 = mdbl实际数量
            !实际金额 = mdbl实际金额
            !实际差价 = mdbl实际差价
            !库存数量 = mdbl库存数量
        End If
        
        !指导批发价 = RecPhysic!指导批发价
        !指导差价率 = RecPhysic!指导差价率
        !成本价 = IIf(Val(RecPhysic!初始成本价) = 0, Val(RecPhysic!指导批发价), RecPhysic!初始成本价)
        
        .Update
    End With
    
    CombinateRec = True
End Function

Private Function CheckData() As Boolean
    Dim DblCurStock As Double       '当前库存数
    '检测是否允许选择
    CheckData = False
    
    If Cmd确定.Enabled = False Then Exit Function
    
    'lng供应商ID不为零，表示退货，无库存时不准继续
    If lng供应商ID <> 0 Then
        If Msf批次.Visible Then
            If Val(Msf批次.TextMatrix(Msf批次.Row, mconIntCol上次供应商ID)) <> 0 And lng供应商ID <> Val(Msf批次.TextMatrix(Msf批次.Row, mconIntCol上次供应商ID)) Then
                MsgBox "你选择的退货商不是该药品的供应商，不能继续操作！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    If Msf批次.Visible Then
        If blnStock Then
            DblCurStock = Val(Msf批次.TextMatrix(Msf批次.Row, mconIntCol可用数量))
        Else
            DblCurStock = Get可用库存(Val(Msf药品规格.TextMatrix(Msf药品规格.Row, mconIntColSpec药品ID)), Val(Msf批次.TextMatrix(Msf批次.Row, mconIntCol批次)))
        End If
    Else
        If Not RecPhysic.EOF Then
            If blnStock Then
                DblCurStock = Val(Msf药品规格.TextMatrix(Msf药品规格.Row, mconIntColSpec可用数量))
            Else
                DblCurStock = Get可用库存(Val(Msf药品规格.TextMatrix(Msf药品规格.Row, mconIntColSpec药品ID)))
            End If
        End If
    End If
    
    If DblCurStock > 0 Then
        CheckData = True
        Exit Function
    End If
    
    '如果源库房与目库房为空，则表明是药品目录自己在进行常规设置，不判断
    If (Lng源库房ID = 0 And Lng目库房ID = 0) Then
        CheckData = True
        Exit Function
    End If
    
    '如果是盘点单调用药品选择器，则不需判断，直接退出
    If bln盘点单 Then
        CheckData = True
        Exit Function
    End If
    
    '如果是药品库存差价调整，则不需判断，直接退出
    If glngModul = 1303 Then
        CheckData = True
        Exit Function
    End If
    
    If Msf批次.Visible Or bln时价 Then
        If (DblCurStock > 0) Or Not blnPrice Or Msf批次.TextMatrix(Msf批次.Row, mconIntCol批号) = "新增批次药品" Then CheckData = True: Exit Function
        MsgBox "该" & IIf(bln时价, "时价", "批次") & "药品已经没有库存，不能继续操作！", vbInformation, gstrSysName
        Exit Function
    Else
        If blnCheck = False Then
           CheckData = True
           Exit Function
        End If
    End If
    
    Select Case IntStockCheck
    Case 1
        If MsgBox("该药品已经没有库存，是否继续！", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
    Case 2
        MsgBox "该药品已经没有库存，不能继续操作！", vbInformation, gstrSysName
        Exit Function
    End Select
    CheckData = True
End Function

Public Function ShowME(ByVal FrmMain As Form, ByVal 编辑模式 As Integer, Optional ByVal 源库房 As Long, _
                    Optional ByVal 目库房 As Long = 0, Optional ByVal 使用部门 As Long = 0, Optional ByVal Bln检测库存 As Boolean = True, _
                    Optional ByVal bln检查批次或时价 As Boolean = True, Optional ByVal bln盘点单据 As Boolean = False, Optional ByVal bln增加空批次 As Boolean = False, _
                    Optional ByVal bln显示库存 As Boolean = True, Optional ByVal lng供应商 As Long = 0, Optional ByVal bln盘无存储库房药品 As Boolean = False, _
                    Optional ByVal 领用方式 As Integer = 0, Optional ByVal bln包含停用药品 As Boolean = False) As ADODB.Recordset
    'bln检查库存:遵守批次药品及时价药品零库存不准出库原则，可强制允许not (批次 or 时价) 药品出库
    'bln检查批次或时价:允许零库存的批次药品及时价药品出库
    'lng供应商ID:不为零表示退货
    
    With Me
        .In_编辑状态 = 编辑模式
        .In_源库房 = 源库房
        .In_目库房 = 目库房
        .In_部门 = 使用部门
        .In_MainFrm = FrmMain
        bln盘点单 = bln盘点单据
        bln空批次 = bln增加空批次
        blnCheck = Bln检测库存
        blnPrice = bln检查批次或时价
        blnStock = bln显示库存
        lng供应商ID = lng供应商
        mblnNoStock = bln盘无存储库房药品
        int领用方式 = 领用方式
        mbln包含停用药品 = bln包含停用药品
        .Show 1, FrmMain
    End With
    Set ShowME = RecReturn.Clone
End Function

Public Function Get可用库存(ByVal lng药品ID As Long, Optional ByVal lng批次 As Long = 0) As Single
    Dim rsStock As New ADODB.Recordset
    
    gstrSQL = " Select Sum(A.可用数量" & StrUnitString & ") 可用数量,Sum(A.实际数量" & StrUnitString & ") 实际数量,sum(A.实际金额) 实际金额,sum(A.实际差价) 实际差价,Sum(A.实际数量) 库存数量 " & _
              " From 药品库存 A,药品规格 B " & _
              " Where A.药品ID=B.药品ID And A.性质=1 And A.药品ID=[1] " & IIf(lng批次 = 0, "", " And Nvl(A.批次,0)=[2] ")
    If Lng源库房ID <> 0 Or Lng目库房ID <> 0 Then
        gstrSQL = gstrSQL & " And A.库房ID=[3]"
    End If
    gstrSQL = gstrSQL & " Group By A.药品id"
    
    Set rsStock = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[获取可用库存]", lng药品ID, lng批次, IIf(Lng源库房ID = 0, Lng目库房ID, Lng源库房ID))
    
    mdbl可用数量 = 0
    mdbl实际差价 = 0
    mdbl实际金额 = 0
    mdbl实际数量 = 0
    mdbl库存数量 = 0
    If Not rsStock.EOF Then
        mdbl可用数量 = IIf(IsNull(rsStock!可用数量), 0, rsStock!可用数量)
        mdbl实际差价 = IIf(IsNull(rsStock!实际差价), 0, rsStock!实际差价)
        mdbl实际金额 = IIf(IsNull(rsStock!实际金额), 0, rsStock!实际金额)
        mdbl实际数量 = IIf(IsNull(rsStock!实际数量), 0, rsStock!实际数量)
        mdbl库存数量 = IIf(IsNull(rsStock!库存数量), 0, rsStock!库存数量)
    End If
    Get可用库存 = mdbl可用数量
End Function
