VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
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
   Begin VSFlex8Ctl.VSFlexGrid vsf批次 
      Height          =   1635
      Left            =   2640
      TabIndex        =   5
      Top             =   4140
      Width           =   6795
      _cx             =   11986
      _cy             =   2884
      Appearance      =   0
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
      BackColorSel    =   16053482
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483636
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
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Frm药品选择器.frx":0E42
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
   Begin VSFlex8Ctl.VSFlexGrid vsf药品规格 
      Height          =   2805
      Left            =   2640
      TabIndex        =   4
      Top             =   1290
      Width           =   6795
      _cx             =   11986
      _cy             =   4948
      Appearance      =   0
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
      BackColorBkg    =   -2147483636
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
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"Frm药品选择器.frx":0EB7
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
   Begin VB.CommandButton Cmd取消 
      Cancel          =   -1  'True
      Caption         =   "取消(&C)"
      Height          =   350
      Left            =   8250
      TabIndex        =   3
      Top             =   5850
      Width           =   1100
   End
   Begin VB.CommandButton Cmd确定 
      Caption         =   "确定(&O)"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   350
      Left            =   7020
      TabIndex        =   2
      Top             =   5850
      Width           =   1100
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
            Picture         =   "Frm药品选择器.frx":0F2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Frm药品选择器.frx":2586
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
            Picture         =   "Frm药品选择器.frx":4290
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
Private mstrPreType As String                   '上次选择的药品类别（西药、成药、草药）

'默认的列顺序
Const mstrDefaultSpec As String = "剂型,药名编码,来源,基本药物,药典ID,用途分类ID,剂量单位,药品编码,通用名称,药品名称,商品名,规格,产地,药名ID,药品ID,上次采购价,售价,售价单位,售价包装,门诊单位,门诊包装,住院单位,住院包装,药库单位,药库包装,可用数量,库存数量,库存金额,库存差价,有效期,药库分批,药房分批,时价,指导批发价,指导差价率,库房货位,批准文号,实际数量,留存数量,合同单位,药价级别"
Const mstrDefaultBatch1 As String = "RID,库房,批次,入库日期,批号,生产日期,失效期,产地,成本价,售价,可用数量,库存数量,库存金额,库存差价,上次供应商ID,实际数量,批准文号"
Private mstrDefaultBatch As String              '批次列顺序（因为效期列名可能会变化，所以实际用这个变量，在SetUserDefine中赋值）

'用户自定义的列顺序
Private mstrColumnSpecSequence As String            '规格
Private mstrColumnBatchSequence As String           '批次
    
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
Private LngLastSelect药品ID As Long             '上次选择的药品ID（用于是否刷新）
Private mbln中药库房 As Boolean
Private mblnNoStock As Boolean                  '本地参数：是否允许盘点没有设置存储库房的药品
Private int领用方式 As Integer                  '0-向库房领药;1-向科室留存领药
Private mbln包含停用药品 As Boolean
Private mbln忽略服务对象 As Boolean

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
Private mconIntColRID As Integer
Private mconIntCol库房 As Integer
Private mconIntCol批次 As Integer
Private mconIntCol入库日期 As Integer
Private mconIntCol批号 As Integer
Private mconIntCol生产日期 As Integer
Private mconIntCol失效期 As Integer
Private mconIntCol产地 As Integer
Private mconintCol成本价 As Integer
Private mconIntCol售价 As Integer
Private mconIntCol可用数量 As Integer
Private mconintCol库存数量 As Integer
Private mconIntCol库存金额 As Integer
Private mconIntCol库存差价 As Integer
Private mconIntCol上次供应商ID As Integer
Private mconIntCol实际数量 As Integer
Private mconIntCol批准文号 As Integer

'规格列表
Private Const mconIntColSpec列数 As Integer = 41
Private mconIntColSpec剂型 As Integer
Private mconIntColSpec药名编码 As Integer
Private mconIntColSpec通用名称 As Integer
Private mconIntColSpec药品来源 As Integer
Private mconIntColSpec基本药物 As Integer
Private mconIntColSpec药典ID As Integer
Private mconIntColSpec用途分类ID As Integer
Private mconIntColSpec剂量单位 As Integer
Private mconIntColSpec药品编码 As Integer
Private mconIntColSpec药品名称 As Integer
Private mconIntColSpec商品名 As Integer
Private mconIntColSpec规格 As Integer
Private mconIntColSpec产地 As Integer
Private mconIntColSpec药名ID As Integer
Private mconIntColSpec药品ID As Integer
Private mconIntColSpec上次采购价 As Integer
Private mconIntColSpec售价 As Integer
Private mconIntColSpec售价单位 As Integer
Private mconIntColSpec售价包装 As Integer
Private mconIntColSpec门诊单位 As Integer
Private mconIntColSpec门诊包装 As Integer
Private mconIntColSpec住院单位 As Integer
Private mconIntColSpec住院包装 As Integer
Private mconIntColSpec药库单位 As Integer
Private mconIntColSpec药库包装 As Integer
Private mconIntColSpec可用数量 As Integer
Private mconIntColSpec库存数量 As Integer
Private mconIntColSpec库存金额 As Integer
Private mconIntColSpec库存差价 As Integer
Private mconIntColSpec有效期 As Integer
Private mconIntColSpec药库分批 As Integer
Private mconIntColSpec药房分批 As Integer
Private mconIntColSpec时价 As Integer
Private mconIntColSpec指导批发价 As Integer
Private mconIntColSpec指导差价率 As Integer
Private mconIntColSpec库房货位 As Integer
Private mconIntColSpec批准文号 As Integer
Private mconIntColSpec实际数量 As Integer
Private mconIntColSpec留存数量 As Integer
Private mconIntColSpec合同单位 As Integer
Private mconIntColSpec药价级别 As Integer
Private Sub SetColumnValue(ByVal intType As Integer, ByVal str列名 As String, ByVal intValue As Integer)
    Select Case intType
    Case 0  '规格
        Select Case str列名
        Case "剂型"
            mconIntColSpec剂型 = intValue
        Case "药名编码"
            mconIntColSpec药名编码 = intValue
        Case "通用名称"
            mconIntColSpec通用名称 = intValue
        Case "药品来源", "来源"
            mconIntColSpec药品来源 = intValue
        Case "基本药物"
            mconIntColSpec基本药物 = intValue
        Case "药典ID"
            mconIntColSpec药典ID = intValue
        Case "用途分类ID"
            mconIntColSpec用途分类ID = intValue
        Case "剂量单位"
            mconIntColSpec剂量单位 = intValue
        Case "药品编码"
            mconIntColSpec药品编码 = intValue
        Case "药品名称"
            mconIntColSpec药品名称 = intValue
        Case "商品名"
            mconIntColSpec商品名 = intValue
        Case "规格"
            mconIntColSpec规格 = intValue
            
        Case "产地"
            mconIntColSpec产地 = intValue
        Case "药名ID"
            mconIntColSpec药名ID = intValue
        Case "药品ID"
            mconIntColSpec药品ID = intValue
        Case "上次采购价"
            mconIntColSpec上次采购价 = intValue
        Case "售价"
            mconIntColSpec售价 = intValue
        Case "售价单位"
            mconIntColSpec售价单位 = intValue
        Case "售价包装"
            mconIntColSpec售价包装 = intValue
        Case "门诊单位"
            mconIntColSpec门诊单位 = intValue
        Case "门诊包装"
            mconIntColSpec门诊包装 = intValue
        Case "住院单位"
            mconIntColSpec住院单位 = intValue
            
        Case "住院包装"
            mconIntColSpec住院包装 = intValue
        Case "药库单位"
            mconIntColSpec药库单位 = intValue
        Case "药库包装"
            mconIntColSpec药库包装 = intValue
        Case "可用数量"
            mconIntColSpec可用数量 = intValue
        Case "库存数量"
            mconIntColSpec库存数量 = intValue
        Case "库存金额"
            mconIntColSpec库存金额 = intValue
        Case "库存差价"
            mconIntColSpec库存差价 = intValue
        Case "有效期"
            mconIntColSpec有效期 = intValue
        Case "药库分批"
            mconIntColSpec药库分批 = intValue
        Case "药房分批"
            mconIntColSpec药房分批 = intValue
            
        Case "时价"
            mconIntColSpec时价 = intValue
        Case "指导批发价"
            mconIntColSpec指导批发价 = intValue
        Case "指导差价率"
            mconIntColSpec指导差价率 = intValue
        Case "库房货位"
            mconIntColSpec库房货位 = intValue
        Case "批准文号"
            mconIntColSpec批准文号 = intValue
        Case "实际数量"
            mconIntColSpec实际数量 = intValue
        Case "留存数量"
            mconIntColSpec留存数量 = intValue
        Case "合同单位"
            mconIntColSpec合同单位 = intValue
        Case "药价级别"
            mconIntColSpec药价级别 = intValue
        End Select
    Case 1 '批次
        Select Case str列名
        Case "RID"
            mconIntColRID = intValue
        Case "库房"
            mconIntCol库房 = intValue
        Case "批次"
            mconIntCol批次 = intValue
        Case "入库日期"
            mconIntCol入库日期 = intValue
        Case "批号"
            mconIntCol批号 = intValue
        Case "生产日期"
            mconIntCol生产日期 = intValue
        Case "失效期", "有效期至"
            mconIntCol失效期 = intValue
        Case "产地"
            mconIntCol产地 = intValue
        Case "成本价"
            mconintCol成本价 = intValue
        Case "售价"
            mconIntCol售价 = intValue
        Case "可用数量"
            mconIntCol可用数量 = intValue
        Case "库存数量"
            mconintCol库存数量 = intValue
        Case "库存金额"
            mconIntCol库存金额 = intValue
        Case "库存差价"
            mconIntCol库存差价 = intValue
        Case "上次供应商ID"
            mconIntCol上次供应商ID = intValue
        Case "实际数量"
            mconIntCol实际数量 = intValue
        Case "批准文号"
            mconIntCol批准文号 = intValue
        End Select
    End Select
End Sub

Private Sub SetUserDefine()
    Dim arrText As Variant
    Dim i
    
    '规格列设置
    mstrColumnSpecSequence = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name & mstrCaption & "\VSFlexGrid", "vsf药品规格" & vsf药品规格.Tag & "名称", "")
    
    mstrColumnSpecSequence = Replace(mstrColumnSpecSequence, "药品来源", "来源")
    
    If mstrColumnSpecSequence = "" Then
        mstrColumnSpecSequence = mstrDefaultSpec
    Else
        arrText = Split(mstrColumnSpecSequence, ",")
        
        '列数变了,使用缺省列顺序
        If mconIntColSpec列数 <> UBound(arrText) + 1 Then
            mstrColumnSpecSequence = mstrDefaultSpec
        End If
            
        '列标题变了,使用缺省列顺序
        For i = 0 To UBound(arrText)
            If InStr(1, "," & mstrDefaultSpec & ",", "," & arrText(i) & ",") = 0 Then
                mstrColumnSpecSequence = mstrDefaultSpec
                Exit For
            End If
        Next
    End If
    
    If mstrColumnSpecSequence = mstrDefaultSpec Then
        '默认列设置
        mconIntColSpec剂型 = 0
        mconIntColSpec药名编码 = 1
        mconIntColSpec药品来源 = 2
        mconIntColSpec基本药物 = 3
        mconIntColSpec药典ID = 4
        mconIntColSpec用途分类ID = 5
        mconIntColSpec剂量单位 = 6
        mconIntColSpec药品编码 = 7
        mconIntColSpec通用名称 = 8
        mconIntColSpec药品名称 = 9
        mconIntColSpec商品名 = 10
        mconIntColSpec规格 = 11
        mconIntColSpec产地 = 12
        mconIntColSpec药名ID = 13
        mconIntColSpec药品ID = 14
        mconIntColSpec上次采购价 = 15
        mconIntColSpec售价 = 16
        mconIntColSpec售价单位 = 17
        mconIntColSpec售价包装 = 18
        mconIntColSpec门诊单位 = 19
        mconIntColSpec门诊包装 = 20
        mconIntColSpec住院单位 = 21
        mconIntColSpec住院包装 = 22
        mconIntColSpec药库单位 = 23
        mconIntColSpec药库包装 = 24
        mconIntColSpec可用数量 = 25
        mconIntColSpec库存数量 = 26
        mconIntColSpec库存金额 = 27
        mconIntColSpec库存差价 = 28
        mconIntColSpec有效期 = 29
        mconIntColSpec药库分批 = 30
        mconIntColSpec药房分批 = 31
        mconIntColSpec时价 = 32
        mconIntColSpec指导批发价 = 33
        mconIntColSpec指导差价率 = 34
        mconIntColSpec库房货位 = 35
        mconIntColSpec批准文号 = 36
        mconIntColSpec实际数量 = 37
        mconIntColSpec留存数量 = 38
        mconIntColSpec合同单位 = 39
        mconIntColSpec药价级别 = 40
    Else
        '用户列设置
        arrText = Split(mstrColumnSpecSequence, ",")
        
        For i = 0 To mconIntColSpec列数 - 1
            Call SetColumnValue(0, arrText(i), i)
        Next
    End If
    
    
    '批次列设置
    If gtype_UserSysParms.P149_效期显示方式 = 1 Then
        mstrDefaultBatch = Replace(mstrDefaultBatch1, "失效期", "有效期至")
    Else
        mstrDefaultBatch = mstrDefaultBatch1
    End If
    
    mstrColumnBatchSequence = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name & mstrCaption & "\VSFlexGrid", "vsf批次" & vsf批次.Tag & "名称", "")
    
    If mstrColumnBatchSequence = "" Then
        mstrColumnBatchSequence = mstrDefaultBatch
    Else
        If gtype_UserSysParms.P149_效期显示方式 = 1 Then
            mstrColumnBatchSequence = Replace(mstrColumnBatchSequence, "失效期", "有效期至")
        Else
            mstrColumnBatchSequence = Replace(mstrColumnBatchSequence, "有效期至", "失效期")
        End If
        
        arrText = Split(mstrColumnBatchSequence, ",")
        
        '列数变了,使用缺省列顺序
        If mconIntCol列数 <> UBound(arrText) + 1 Then
            mstrColumnBatchSequence = mstrDefaultBatch
        End If
            
        '列标题变了,使用缺省列顺序
        For i = 0 To UBound(arrText)
            If InStr(1, "," & mstrColumnBatchSequence & ",", "," & arrText(i) & ",") = 0 Then
                mstrColumnBatchSequence = mstrDefaultBatch
                Exit For
            End If
        Next
    End If
    
    If mstrColumnBatchSequence = mstrDefaultBatch Then
        '默认列设置
         mconIntColRID = 0
         mconIntCol库房 = 1
         mconIntCol批次 = 2
         mconIntCol入库日期 = 3
         mconIntCol批号 = 4
         mconIntCol生产日期 = 5
         mconIntCol失效期 = 6
         mconIntCol产地 = 7
         mconintCol成本价 = 8
         mconIntCol售价 = 9
         mconIntCol可用数量 = 10
         mconintCol库存数量 = 11
         mconIntCol库存金额 = 12
         mconIntCol库存差价 = 13
         mconIntCol上次供应商ID = 14
         mconIntCol实际数量 = 15
         mconIntCol批准文号 = 16
    Else
        '用户列设置
        arrText = Split(mstrColumnBatchSequence, ",")
        
        For i = 0 To mconIntCol列数 - 1
            Call SetColumnValue(1, arrText(i), i)
        Next
    End If
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
        With vsf药品规格
            
            If BlnSetHeader Then
                .ExplorerBar = flexExSortShowAndMove
                
                .Cols = mconIntColSpec列数
                '卡片
                .TextMatrix(0, mconIntColSpec剂型) = "剂型"
                .TextMatrix(0, mconIntColSpec药名编码) = "药名编码"
                .TextMatrix(0, mconIntColSpec通用名称) = "通用名称"
                .TextMatrix(0, mconIntColSpec药品来源) = "来源"
                .TextMatrix(0, mconIntColSpec基本药物) = "基本药物"
                .TextMatrix(0, mconIntColSpec药典ID) = "药典ID"
                .TextMatrix(0, mconIntColSpec用途分类ID) = "用途分类ID"
                .TextMatrix(0, mconIntColSpec剂量单位) = "剂量单位"
                
                '规格
                .TextMatrix(0, mconIntColSpec药品编码) = "药品编码"
                .TextMatrix(0, mconIntColSpec药品名称) = "药品名称"
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
                .TextMatrix(0, mconIntColSpec留存数量) = "留存数量"
                .TextMatrix(0, mconIntColSpec合同单位) = "合同单位"
                .TextMatrix(0, mconIntColSpec药价级别) = "药价级别"
            End If
            
            For intCol = 0 To .Cols - 1
                .FixedAlignment(intCol) = flexAlignCenterCenter
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
            .ColAlignment(mconIntColSpec留存数量) = 7
            .ColAlignment(mconIntColSpec合同单位) = flexAlignLeftCenter
            .ColAlignment(mconIntColSpec药价级别) = flexAlignLeftCenter
            
'            If BlnStartUp = False Then
                .ColWidth(mconIntColSpec剂型) = 500

                '规格
                .ColWidth(mconIntColSpec药品编码) = 1000
                .ColWidth(mconIntColSpec药品名称) = 1800
                .ColWidth(mconIntColSpec商品名) = 1800
                .ColWidth(mconIntColSpec规格) = 1000
                .ColWidth(mconIntColSpec产地) = 1200
                .ColWidth(mconIntColSpec售价) = 1200
                .ColWidth(mconIntColSpec可用数量) = 1200
                .ColWidth(mconIntColSpec有效期) = 900
                .ColWidth(mconIntColSpec药库分批) = 900
                .ColWidth(mconIntColSpec药房分批) = 900
                .ColWidth(mconIntColSpec时价) = 900
                .ColWidth(mconIntColSpec库房货位) = 1500
                .ColWidth(mconIntColSpec批准文号) = 1000
                .ColWidth(mconIntColSpec留存数量) = 1000
                .ColWidth(mconIntColSpec合同单位) = 1500
                .ColWidth(mconIntColSpec药价级别) = 1000
                
                .Row = 1
                
                '恢复用户列设置
                Call RestoreFlexState(vsf药品规格, App.ProductName & "\" & Me.Name & mstrCaption)
                
                '特殊列的处理
                .ColWidth(mconIntColSpec药名编码) = 0
                .ColWidth(mconIntColSpec药品名称) = 0
                
                If gint药品名称显示 = 2 Then
                    .ColWidth(mconIntColSpec通用名称) = IIf(.ColWidth(mconIntColSpec通用名称) = 0, 1800, .ColWidth(mconIntColSpec通用名称))
                    .ColWidth(mconIntColSpec商品名) = IIf(.ColWidth(mconIntColSpec商品名) = 0, 1800, .ColWidth(mconIntColSpec商品名))
                ElseIf gint药品名称显示 = 0 Then
                    .ColWidth(mconIntColSpec通用名称) = IIf(.ColWidth(mconIntColSpec通用名称) = 0, 1800, .ColWidth(mconIntColSpec通用名称))
                    .ColWidth(mconIntColSpec商品名) = 0
                ElseIf gint药品名称显示 = 1 Then
                    .ColWidth(mconIntColSpec通用名称) = 0
                    .ColWidth(mconIntColSpec商品名) = IIf(.ColWidth(mconIntColSpec商品名) = 0, 1800, .ColWidth(mconIntColSpec商品名))
                End If
                
                .ColWidth(mconIntColSpec药典ID) = 0
                .ColWidth(mconIntColSpec用途分类ID) = 0
                .ColWidth(mconIntColSpec剂量单位) = 0
                .ColWidth(mconIntColSpec药名ID) = 0
                .ColWidth(mconIntColSpec药品ID) = 0
                .ColWidth(mconIntColSpec库存数量) = 0
                .ColWidth(mconIntColSpec库存金额) = 0
                .ColWidth(mconIntColSpec库存差价) = 0
                .ColWidth(mconIntColSpec指导批发价) = 0
                .ColWidth(mconIntColSpec指导差价率) = 0
                .ColWidth(mconIntColSpec实际数量) = 0
                If int领用方式 = 0 Then
                    .ColWidth(mconIntColSpec留存数量) = 0
                End If
                
                .ColWidth(mconIntColSpec售价单位) = IIf(mintUnit = mconint售价单位, 900, 0)
                .ColWidth(mconIntColSpec售价包装) = IIf(mintUnit = mconint售价单位, 900, 0)
                .ColWidth(mconIntColSpec门诊单位) = IIf(mintUnit = mconint门诊单位, 900, 0)
                .ColWidth(mconIntColSpec门诊包装) = IIf(mintUnit = mconint门诊单位, 900, 0)
                .ColWidth(mconIntColSpec住院单位) = IIf(mintUnit = mconint住院单位, 900, 0)
                .ColWidth(mconIntColSpec住院包装) = IIf(mintUnit = mconint住院单位, 900, 0)
                .ColWidth(mconIntColSpec药库单位) = IIf(mintUnit = mconint药库单位, 900, 0)
                .ColWidth(mconIntColSpec药库包装) = IIf(mintUnit = mconint药库单位, 900, 0)
                If mstrCaption = "药品外购入库管理" Then
                    If .ColWidth(mconIntColSpec上次采购价) = 0 Then .ColWidth(mconIntColSpec上次采购价) = 1200
                Else
                    .ColWidth(mconIntColSpec上次采购价) = 0
                End If
                If mstrCaption = "药品外购入库管理" Or mstrCaption = "药品计划管理" Then
                    If .ColWidth(mconIntColSpec合同单位) = 0 Then .ColWidth(mconIntColSpec合同单位) = 1500
                Else
                    .ColWidth(mconIntColSpec合同单位) = 0
                End If
'            End If
        End With
    Case 0
        With vsf批次
            If BlnSetHeader Then
                .ExplorerBar = flexExSortShowAndMove
            
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
                .FixedAlignment(intCol) = flexAlignCenterCenter
            Next
            .ColAlignment(mconIntCol产地) = 7
            .ColAlignment(mconintCol成本价) = 7
            .ColAlignment(mconIntCol售价) = 7
            .ColAlignment(mconIntCol可用数量) = 7
            .ColAlignment(mconintCol库存数量) = 7
            .ColAlignment(mconIntCol库存金额) = 7
            .ColAlignment(mconIntCol库存差价) = 7
            
            '特殊列的处理
            .ColWidth(mconIntColRID) = 0
            .ColWidth(mconIntCol批次) = 0
            .ColWidth(mconIntCol上次供应商ID) = 0
            .ColWidth(mconIntCol实际数量) = 0
            .ColWidth(mconIntCol入库日期) = IIf(mstrCaption = "药品移库管理" Or mstrCaption = "药品申领管理", 1000, 0)
            If BlnStartUp = False Then
                .ColWidth(mconIntCol库房) = 1200
                .ColWidth(mconIntCol批号) = 1000
                .ColWidth(mconIntCol生产日期) = 1000
                .ColWidth(mconIntCol失效期) = 1000
                .ColWidth(mconIntCol产地) = 1200
                .ColWidth(mconintCol成本价) = 1200
                .ColWidth(mconIntCol售价) = 1200
                .ColWidth(mconIntCol可用数量) = 1200
                .ColWidth(mconintCol库存数量) = 1200
                .ColWidth(mconIntCol库存金额) = 1200
                .ColWidth(mconIntCol库存差价) = 1200
                .ColWidth(mconIntCol批准文号) = 1000
                .Row = 1
                
                Call RestoreFlexState(vsf批次, App.ProductName & "\" & Me.Name & mstrCaption)
                
                '特殊列的处理
                .ColWidth(mconIntColRID) = 0
                .ColWidth(mconIntCol批次) = 0
                .ColWidth(mconIntCol上次供应商ID) = 0
                .ColWidth(mconIntCol实际数量) = 0
                .ColWidth(mconIntCol入库日期) = IIf(mstrCaption = "药品移库管理" Or mstrCaption = "药品申领管理", 1000, 0)
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
    vsf药品规格.SetFocus
End Sub

Private Sub Form_Load()
    mstrCaption = GetText(GetParentWindow(OutObj.hWnd))
    
    Call RestoreWinState(Me, App.ProductName, mstrCaption)
    On Error GoTo errHandle
    BlnStartUp = False
    BlnFirstStart = False
    mstrPreNode = ""
    
    '取售价单位
    StrUnitString = ""
    IntStockCheck = 0
    LngLastSelect药品ID = 0
    vsf批次.Visible = (In_编辑状态 = 2)
    
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
    Set RecUnit = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, Lng源库房ID)
    
    With RecUnit
        If Not .EOF Then
            IntStockCheck = !库存检查
        End If
    End With
    
    '检查源库房是否为药库
    If Lng源库房ID <> 0 Then
        int库房 = 3
        
        gstrSQL = "select 部门ID from 部门性质说明 where (工作性质 like '%药房' Or 工作性质 like '%制剂室') And 部门id=[1]"
        Set RecUnit = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, Lng源库房ID)
        
        If RecUnit.EOF Then
            RecUnit.Close
            
            gstrSQL = "select 部门ID from 部门性质说明 where 工作性质 like '%药库' And 部门id=[1]"
            Set RecUnit = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, Lng源库房ID)
            
            If Not RecUnit.EOF Then int库房 = 1
        Else
            int库房 = 2
        End If

    End If
    
    If mstrCaption = "药品申领管理" Then
        Call GetDrugDigit(Lng使用部门ID, mstrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
        Lng使用部门ID = 0
    ElseIf mstrCaption = "药品移库管理" Then
        Call GetDrugDigit(Lng使用部门ID, mstrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
        Lng使用部门ID = 0
    Else
        Call GetDrugDigit(IIf(Lng源库房ID = 0, Lng目库房ID, Lng源库房ID), mstrCaption, mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    End If
    
    mstrCostFormat = "'999999999990." & String(gtype_UserDrugDigits.Digit_成本价, "0") & "'"
    mstrPriceFormat = "'999999999990." & String(gtype_UserDrugDigits.Digit_零售价, "0") & "'"
    mstrNumberFormat = "'999999999990." & String(gtype_UserDrugDigits.Digit_数量, "0") & "'"
    mstrMoneyFormat = "'999999999990." & String(gtype_UserDrugDigits.Digit_金额, "0") & "'"

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
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function ReadAndSendDataToTvw() As Boolean
    Dim NodeThis As Node, ItemThis As ListItem
    Dim Int末级 As Integer
    Dim lng库房ID As Long
    Dim rs材质分类 As New ADODB.Recordset
    
    '药品用途分类是否有数据
    ReadAndSendDataToTvw = False
    On Error GoTo errHandle
    gstrSQL = " Select 编码,名称 From 诊疗项目类别 " & _
              " Where Instr([1],编码,1) > 0 " & _
              " Order by 编码"
    Set rs材质分类 = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, "567")
    
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
        Set RecData = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[检查部门性质]", lng库房ID)
        
        If Not RecData.EOF Then mbln中药库房 = True
        gstrSQL = "Select Distinct J.编码,J.名称 " & _
                 " From 诊疗执行科室 A,药品特性 B,药品剂型 J " & _
                 " Where A.诊疗项目ID=B.药名ID And B.药品剂型=J.名称" & _
                 " And A.执行科室ID=[1]"
        Set RecData = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[提取该库房现在剂型]", lng库房ID)
    Else
        gstrSQL = "Select 编码,名称 From 药品剂型"
        Call zldatabase.OpenRecordset(RecData, gstrSQL, "提取所有药品剂型")
    End If
    
    With RecData
        Lvw.ListItems.Clear
        Do While Not RecData.EOF
            Lvw.ListItems.Add , "K" & RecData!编码, RecData!名称, 1, 1
            RecData.MoveNext
        Loop
        
        If .State = 1 Then .Close
        gstrSQL = "Select ID,上级ID,名称,1 as 末级,decode(类型,1,'西成药',2,'中成药','中草药') as 材质,类型  " & _
                 " From 诊疗分类目录  " & _
                 " where 类型 in (1,2,3)  " & _
                 " Start With 上级ID IS NULL Connect By Prior ID=上级ID Order by level,ID"
        
        Call SQLTest(App.Title, Me.Caption, gstrSQL)
        Set RecData = zldatabase.OpenSQLRecord(gstrSQL, "ReadAndSendDataToTvw")
        Call SQLTest
        
        If RecData.EOF Then
            MsgBox "请初始化药品用途分类（药品用途分类）！", vbInformation, gstrSysName
            Exit Function
        End If
        
        '将药品用途分类数据装入
        Do While Not RecData.EOF
            Int末级 = IIf(RecData!末级 = 1, 2, 1)
            If IsNull(RecData!上级ID) Then
                Set NodeThis = Tvw药品用途分类.Nodes.Add("Root" & RecData!材质, 4, "K_" & RecData!Id, RecData!名称, Int末级, Int末级)
            Else
                Set NodeThis = Tvw药品用途分类.Nodes.Add("K_" & RecData!上级ID, 4, "K_" & RecData!Id, RecData!名称, Int末级, Int末级)
            End If
            NodeThis.Tag = RecData!类型   '存放分类类型:1-西成药,2-中成药,3-中草药
            RecData.MoveNext
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
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
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
    With vsf药品规格
        .Left = ImgUpDown_S.Left
        .Top = ImgLeftRight_S.Top
        .Width = ImgUpDown_S.Width
    End With
    With vsf批次
        If .Visible Then
            .Top = ImgUpDown_S.Top + ImgUpDown_S.Height
            .Height = ImgLeftRight_S.Top + ImgLeftRight_S.Height - .Top
            .Left = vsf药品规格.Left
            .Width = vsf药品规格.Width
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
    
    With vsf药品规格
        .Height = IIf(vsf批次.Visible = False, Tvw药品用途分类.Top + Tvw药品用途分类.Height - .Top, vsf批次.Top - 45 - .Top)
        If .RowIsVisible(.Row) = False Then .TopRow = .TopRow + IIf(.Row > .TopRow, 1, -1)
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If BlnStartUp = False Then Exit Sub
    
'    Call SaveColWidth("")
    Call SaveWinState(Me, App.ProductName, mstrCaption)
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

Private Sub vsf批次_AfterMoveColumn(ByVal Col As Long, Position As Long)
    Dim i As Integer
    
    For i = 0 To mconIntCol列数 - 1
        Call SetColumnValue(1, vsf批次.TextMatrix(0, i), i)
    Next
End Sub

Private Sub vsf批次_DblClick()
    On Error Resume Next
    If Cmd确定.Enabled = False Then Exit Sub
    
    With RecStock
        If .RecordCount <> 0 Then .MoveFirst
        If .EOF Then Exit Sub
        If .RecordCount = 0 Then Exit Sub
    End With
    Call Cmd确定_Click
End Sub

Private Sub vsf批次_GotFocus()
    Call SetGridFocus(vsf批次, True)
End Sub

Private Sub vsf批次_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then vsf批次_DblClick
End Sub

Private Sub vsf批次_LostFocus()
    Call SetGridFocus(vsf批次, False)
End Sub

Private Sub vsf药品规格_AfterMoveColumn(ByVal Col As Long, Position As Long)
    Dim i As Integer
    
    For i = 0 To mconIntColSpec列数 - 1
        Call SetColumnValue(0, vsf药品规格.TextMatrix(0, i), i)
    Next
End Sub

Private Sub vsf药品规格_DblClick()
    If RecPhysic.EOF Then Exit Sub
    If RecPhysic.RecordCount = 0 Then Exit Sub
    
    If Cmd确定.Enabled Then
        Cmd确定_Click
    Else
        MsgBox "该药品没有库存，不能继续操作！", vbInformation, gstrSysName
    End If
End Sub

Private Sub vsf药品规格_EnterCell()
    Dim Lng收费细目ID As Long, intCol As Integer, LngSelectRow As Long
    Dim strTmp As String, RecGetPrice As New ADODB.Recordset
    Dim strSql效期 As String
    Dim str售价 As String
    
    On Error GoTo errHandle
    With vsf药品规格
        '如果该规格药品的价格到执行时间还未执行,则触发
        Lng收费细目ID = Val(.TextMatrix(.Row, mconIntColSpec药品ID))
        If Lng收费细目ID = 0 Then
            If vsf批次.Visible Then
                vsf批次.Clear
                vsf批次.rows = 2
                Call SetFormat(0, True)
            Else
                Call SetFormat(0, True)
            End If
            Exit Sub
        End If
        
        If LngLastSelect药品ID = Lng收费细目ID Then Exit Sub
        LngLastSelect药品ID = Lng收费细目ID
        
        '检查并执行调价
        Call AutoAdjustPrice_ByID(Lng收费细目ID)
    End With
    
    If In_编辑状态 = 2 Then
        vsf批次.Visible = False
        '读出该药品规格下所有的药品批次库存信息
        bln时价 = (vsf药品规格.TextMatrix(vsf药品规格.Row, mconIntColSpec时价) = "是")
        int分批 = 0
        str售价 = vsf药品规格.TextMatrix(vsf药品规格.Row, mconIntColSpec售价)
        If vsf药品规格.TextMatrix(vsf药品规格.Row, mconIntColSpec药库分批) = "是" Or vsf药品规格.TextMatrix(vsf药品规格.Row, mconIntColSpec药房分批) = "是" Then
            If vsf药品规格.TextMatrix(vsf药品规格.Row, mconIntColSpec药库分批) = "是" And vsf药品规格.TextMatrix(vsf药品规格.Row, mconIntColSpec药房分批) = "是" Then
                int分批 = 3
            ElseIf vsf药品规格.TextMatrix(vsf药品规格.Row, mconIntColSpec药库分批) = "是" Then
                int分批 = 1
            Else
                int分批 = 2
            End If
        End If
        If Not ((int分批 = 3 And int库房 <> 3) Or (int分批 = 1 And int库房 = 1) Or (int分批 = 2 And int库房 = 2)) Then '如果该药品不分批
            vsf批次.Visible = False
            Form_Resize
        Else
            If vsf批次.Visible = False Then vsf批次.Visible = True
        End If
        Form_Resize
        
        With RecStock
            If .State = 1 Then .Close
            gstrSQL = "Select " & mstrColumnBatchSequence & " From ("
            If bln空批次 Then
                strSql效期 = IIf(gtype_UserSysParms.P149_效期显示方式 = 1, "有效期至", "失效期")
                gstrSQL = gstrSQL & "Select 1 RID,名称 库房,0 批次,'' 入库日期,'新增批次药品' 批号,NULL 生产日期,sysdate " & strSql效期 & ",'' 产地,'' 成本价,''售价," & _
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
                    strTmp = " To_Char(K.上次采购价," & mstrNumberFormat & ") 成本价, " & _
                             IIf(bln时价 = True, " Decode(Nvl(K.批次,0),0,Decode(Sign(K.实际数量),1,To_Char(K.实际金额/K.实际数量," & mstrNumberFormat & "),'" & str售价 & "'),To_char(Nvl(K.零售价, K.实际金额 / K.实际数量)," & mstrNumberFormat & ")) 售价, ", " '" & str售价 & "' 售价, ") & _
                             " To_Char(K.可用数量," & mstrNumberFormat & ") 可用数量," & _
                             " To_Char(K.实际数量," & mstrNumberFormat & ") 库存数量,"
                Case mconint门诊单位
                    strTmp = " To_Char(K.上次采购价*nvl(D.门诊包装,1)," & mstrNumberFormat & ") 成本价, " & _
                             IIf(bln时价 = True, " Decode(Nvl(K.批次,0),0,Decode(Sign(K.实际数量),1,To_Char(K.实际金额/K.实际数量*nvl(D.门诊包装,1)," & mstrNumberFormat & "),'" & str售价 & "'),To_char(Nvl(K.零售价, K.实际金额 / K.实际数量)*nvl(D.门诊包装,1)," & mstrNumberFormat & ")) 售价, ", " '" & str售价 & "' 售价, ") & _
                             " To_Char(K.可用数量" & StrUnitString & "," & mstrNumberFormat & ") 可用数量," & _
                             " To_Char(K.实际数量" & StrUnitString & "," & mstrNumberFormat & ") 库存数量,"
                Case mconint住院单位
                    strTmp = " To_Char(K.上次采购价*nvl(D.住院包装,1)," & mstrNumberFormat & ") 成本价, " & _
                             IIf(bln时价 = True, " Decode(Nvl(K.批次,0),0,Decode(Sign(K.实际数量),1,To_Char(K.实际金额/K.实际数量*nvl(D.住院包装,1)," & mstrNumberFormat & "),'" & str售价 & "'),To_char(Nvl(K.零售价, K.实际金额 / K.实际数量)*nvl(D.住院包装,1)," & mstrNumberFormat & ")) 售价, ", " '" & str售价 & "' 售价, ") & _
                             " To_Char(K.可用数量" & StrUnitString & "," & mstrNumberFormat & ") 可用数量," & _
                             " To_Char(K.实际数量" & StrUnitString & "," & mstrNumberFormat & ") 库存数量,"
                Case mconint药库单位
                    strTmp = " To_Char(K.上次采购价*nvl(D.药库包装,1)," & mstrNumberFormat & ") 成本价, " & _
                             IIf(bln时价 = True, " Decode(Nvl(K.批次,0),0,Decode(Sign(K.实际数量),1,To_Char(K.实际金额/K.实际数量*nvl(D.药库包装,1)," & mstrNumberFormat & "),'" & str售价 & "'),To_char(Nvl(K.零售价, K.实际金额 / K.实际数量)*nvl(D.药库包装,1)," & mstrNumberFormat & ")) 售价, ", " '" & str售价 & "' 售价, ") & _
                             " To_Char(K.可用数量" & StrUnitString & "," & mstrNumberFormat & ") 可用数量," & _
                             " To_Char(K.实际数量" & StrUnitString & "," & mstrNumberFormat & ") 库存数量,"
                End Select
            Else
                strTmp = "'' 成本价, '' 售价, '' 可用数量,'' 库存数量,"
            End If
            
            gstrSQL = gstrSQL & strTmp & IIf(blnStock, " To_Char(K.实际金额," & mstrMoneyFormat & ") 库存金额,", "'' 库存金额,") & _
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
        
        gstrSQL = gstrSQL & ")"
       
        Set RecStock = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, Lng目库房ID, IIf(Lng源库房ID = 0, Lng目库房ID, Lng源库房ID), LngLastSelect药品ID)
        
        Dim BlnState As Boolean
        With vsf批次
            If Not RecStock.EOF Then
                Set .DataSource = RecStock
                .ColWidth(mconIntColRID) = 0
            Else
                .Clear
                .rows = 2
            End If
            DoEvents
            Call SetFormat(0, RecStock.EOF)
            DoEvents
            If bln空批次 And RecStock.RecordCount <> 0 Then .Row = IIf(.rows > 2, 2, 1)
            BlnState = ((int分批 = 3 And int库房 <> 3) Or (int分批 = 1 And int库房 = 1) Or (int分批 = 2 And int库房 = 2)) And Not RecStock.EOF
            .Visible = BlnState
        End With
        Form_Resize
    End If
    
    '设置按钮状态
    With RecPhysic
        If .RecordCount <> 0 Then .MoveFirst
        .Find "药品ID=" & Val(vsf药品规格.TextMatrix(vsf药品规格.Row, mconIntColSpec药品ID))
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
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsf药品规格_GotFocus()
    Call SetGridFocus(vsf药品规格, True)
End Sub

Private Sub vsf药品规格_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then vsf药品规格_DblClick
End Sub

Private Sub vsf药品规格_LostFocus()
    Call SetGridFocus(vsf药品规格, False)
End Sub

Private Sub Tvw药品用途分类_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim strTmp As String, StrGroupBy As String
    Dim str单位转换串 As String
    Dim str显示留存 As String
    Dim strNodeType As String
    Dim str剂型 As String
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
    
            
    If Node.Key = mstrPreNode Then Exit Sub
    
    Call SaveWinState(Me, App.ProductName, mstrCaption)
    
    mstrPreNode = Node.Key
    
    strNodeType = Tvw药品用途分类.SelectedItem.Tag
    If strNodeType = "5" Then
        strNodeType = "1"
    ElseIf strNodeType = "6" Then
        strNodeType = "2"
    ElseIf strNodeType = "7" Then
        strNodeType = "3"
    End If
    
    DoEvents
    If mstrPreType = "" Then
        '表示为初始时，设置初始的表格状态
        mstrPreType = strNodeType

        vsf药品规格.Tag = mstrPreType
        vsf批次.Tag = mstrPreType
    ElseIf mstrPreType <> strNodeType Then
        BlnStartUp = False
        
        '保存上个类别的表格状态
        Call SaveFlexState(vsf药品规格, App.ProductName & "\" & Me.Name & mstrCaption)
        Call SaveFlexState(vsf批次, App.ProductName & "\" & Me.Name & mstrCaption)
        
        '设置当前的表格类型
        mstrPreType = strNodeType
        vsf药品规格.Tag = mstrPreType
        vsf批次.Tag = mstrPreType
    End If
    
    Call SetUserDefine
    DoEvents
    
    str显示留存 = IIf(int领用方式 = 1, ",To_Char(S.留存数量 ," & mstrNumberFormat & ") 留存数量", ",'' 留存数量 ")

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
    
    '取剂型
    str剂型 = Get剂型SQL
    
    With RecPhysic
        If .State = 1 Then .Close
        
        '对列头排顺序
        gstrSQL = "Select /*+ Rule*/ " & mstrColumnSpecSequence & " From ("
        gstrSQL = gstrSQL & " Select D.剂型,D.药名编码,D.通用名称,D.药品来源 as 来源,D.基本药物,D.药典ID,D.用途分类ID,D.剂量单位,D.药品编码,D.药品名称,D.商品名,D.规格," & IIf(IntEditState = 1, "D.产地", "Nvl(D.产地,S.产地)") & " AS 产地," & _
                " D.药名ID,D.药品ID,trim(to_char(D.初始成本价" & str单位转换串 & "," & mstrCostFormat & ")) As 上次采购价,trim(to_char(P.售价" & str单位转换串 & ", " & mstrPriceFormat & ")) As 售价," & _
                " D.售价单位,D.剂量系数 As 售价包装,D.门诊单位,D.门诊包装,D.住院单位,D.住院包装,D.药库单位,D.药库包装," & _
                IIf(blnStock, " To_Char(S.可用数量 " & StrUnitString & " ," & mstrNumberFormat & ") 可用数量,To_Char(S.库存数量 " & StrUnitString & "," & mstrNumberFormat & ") 库存数量,S.库存金额,S.库存差价,", "'' 可用数量,'' 库存数量,'' 库存金额,'' 库存差价,") & _
                " D.最大效期 有效期,D.药库分批,D.药房分批,D.时价,D.指导批发价,D.指导差价率,E.库房货位,D.批准文号,To_Char(S.库存数量 ," & mstrNumberFormat & ") 实际数量" & str显示留存 & ",D.合同单位,D.药价级别" & _
                " From"
        '药品信息，药品目录
        gstrSQL = gstrSQL & " (SELECT DISTINCT J.名称 剂型,C.编码 药名编码,C.名称 AS 通用名称,0 AS 药典ID,M.分类ID AS 用途分类ID,M.计算单位 AS 剂量单位,C.编码 AS 药品编码,C.名称 As 药品名称," & _
                IIf(gint药品名称显示 = 1, "Nvl(A.名称,C.名称) As 商品名", " A.名称 As 商品名") & ",C.规格,C.产地,D.药品来源,D.基本药物,D.批准文号, D.药名ID, D.药品ID, C.计算单位 AS 售价单位," & _
                " To_Char(D.剂量系数," & StrFormat & " ) 剂量系数,nvl(To_Char(D.最大效期,'9999990'),0) 最大效期," & _
                " DECODE(D.药库分批,1,'是','否') 药库分批,DECODE(D.药房分批,1,'是','否') 药房分批,DECODE(C.是否变价,1,'是','否') 时价," & _
                " D.门诊单位,To_Char(D.门诊包装," & StrFormat & " ) 门诊包装,D.住院单位," & _
                " To_Char(D.住院包装," & StrFormat & " ) 住院包装,D.药库单位,To_Char(D.药库包装," & StrFormat & " ) 药库包装," & _
                " To_Char(D.指导批发价," & mstrCostFormat & ") 指导批发价,nvl(D.成本价,0) 初始成本价,To_Char(D.指导差价率," & StrFormat & " ) 指导差价率,Q.名称 As 合同单位,D.药价级别" & _
                " FROM 收费项目目录 C,药品规格 D,收费项目别名 A,药品剂型 J,药品特性 T,诊疗项目目录 M,供应商 Q," & _
                "             (Select 执行科室ID,收费细目ID From 收费执行科室 Where 执行科室ID" & IIf(Lng源库房ID <> 0, "=[1]", IIf(Lng目库房ID <> 0, "=[2]", " Is Not NULL")) & " Group By 执行科室ID,收费细目ID) K," & _
                "             (Select 执行科室ID,收费细目ID From 收费执行科室 Where 执行科室ID" & IIf(Lng目库房ID <> 0, "=[2]", IIf(Lng源库房ID <> 0, "=[1]", " Is Not NULL")) & " Group By 执行科室ID,收费细目ID) I "
        
        If str剂型 <> "" Then
            gstrSQL = gstrSQL & ",Table(Cast(f_Str2list([5]) As zlTools.t_Strlist)) B"
        End If
        
        gstrSQL = gstrSQL & " WHERE C.ID=D.药品ID AND D.药名ID=T.药名ID AND T.药名ID=M.ID AND M.类别 IN ('5','6','7')" & _
                " AND D.药品ID=K.收费细目ID" & IIf(mblnNoStock, "(+)", "") & " " & _
                " And D.药品ID=I.收费细目ID" & IIf(mblnNoStock, "(+)", "") & " " & _
                " AND D.药品ID=A.收费细目ID(+) AND A.性质(+)=3 And (C.站点 = '" & gstrNodeNo & "' Or C.站点 is Null) " & _
                " AND T.药品剂型=J.名称(+) And D.合同单位ID=Q.ID(+) "
                'IIf(Lng使用部门ID <> 0, " And K.开单科室ID=I.开单科室ID And K.开单科室ID=" & Lng使用部门ID, "")
        gstrSQL = gstrSQL & "" & _
            " and ([2] is null" & _
                " or exists(select 1 from 部门性质说明 where 工作性质='制剂室' and 部门id=[2])" & _
                " or C.类别=(select distinct '5' from 部门性质说明 where 工作性质 like '西药%' and 部门id=[2])" & _
                " or C.类别=(select distinct '6' from 部门性质说明 where 工作性质 like '成药%' and 部门id=[2])" & _
                " or C.类别=(select distinct '7' from 部门性质说明 where 工作性质 like '中药%' and 部门id=[2]) Or [2]=0)" & _
            " and ([2] is null" & _
                " or exists(select 1 from 部门性质说明 where 工作性质 like '%药库' and 部门id=[2])" & _
                " or exists(select 1 from 部门性质说明 where 工作性质='制剂室' and 部门id=[2])"
        
        If mbln忽略服务对象 = True Then
            gstrSQL = gstrSQL & " Or 1=1)"
        Else
            gstrSQL = gstrSQL & " or decode(C.服务对象,1,1,3,1,0)=(select distinct '1' from 部门性质说明 where 工作性质 like '%药房' and 部门id=[2] and 服务对象 in(1,3))" & _
                " or decode(C.服务对象,2,1,3,1,0)=(select distinct '1' from 部门性质说明 where 工作性质 like '%药房' and 部门id=[2] and 服务对象 in(2,3)) Or [2]=0)"
        End If
        
        '查找指定药品用途分类的规格药品
        If Not (Node.Key Like "Root*") Then
            gstrSQL = gstrSQL & _
                    " And M.分类ID IN " & _
                    "     (Select ID from 诊疗分类目录 " & _
                    "     Where 类型 In (1,2,3)" & _
                    "     Start With ID=[6]" & _
                    "     Connect By Prior ID=上级ID)"
        Else
            gstrSQL = gstrSQL & " And M.类别=[7] "
        End If
        '只查找未停用的规格药品（需要根据传入参数决定，暂时只有盘点时该参数才可能为True）
        If mbln包含停用药品 = False Then
            gstrSQL = gstrSQL & " And (C.撤档时间 Is Null Or To_char(C.撤档时间,'yyyy-MM-dd')='3000-01-01')"
        End If
        
         '只查找指定剂型的规格药品
        If str剂型 <> "" Then
            gstrSQL = gstrSQL & " And T.药品剂型=B.Column_Value "
        End If
        
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
                        " To_Char(Sum(a.实际差价),'99999999999990.99999') 库存差价,'' 留存数量 " & _
                        " From 药品库存 a Where a.性质=1 "
            End If
        Else
            gstrSQL = gstrSQL & " (Select 药品id,' ' 产地, '' 可用数量," & _
                    " '' 库存数量,'' 库存金额,'' 库存差价,'' 留存数量" & _
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
        
        gstrSQL = gstrSQL & ")"
    End With
    
    Set RecPhysic = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption, Lng源库房ID, Lng目库房ID, Lng使用部门ID, Format(zldatabase.Currentdate(), "yyyy"), str剂型, Val(Mid(Node.Key, 3)), Node.Tag)
    
    With vsf药品规格
        If Not RecPhysic.EOF Then
            Set .DataSource = RecPhysic
        Else
            .Clear
            .rows = 2
            LngLastSelect药品ID = 0
        End If
        DoEvents
        Call SetFormat(1, RecPhysic.EOF)
    End With
    Cmd确定.Enabled = (RecPhysic.EOF <> True)
    
    Call vsf药品规格_EnterCell
         
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
            strReturn = strReturn & "," & ItemThis.Text & ""
        End If
    Next
    If mbln中药库房 And strReturn <> "" Then strReturn = strReturn & ",方剂"
    
    If strReturn = "" Then Exit Function
    strReturn = Mid(strReturn, 2)
    Get剂型SQL = strReturn
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
            .Fields.Append "基本药物", adVarChar, 30, adFldIsNullable
            .Fields.Append "通用名", adLongVarChar, 40, adFldIsNullable
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
    
    On Error GoTo errHandle
    CombinateRec = False
    With RecPhysic
        If .RecordCount <> 0 Then .MoveFirst
        .Find "药品ID=" & Val(vsf药品规格.TextMatrix(vsf药品规格.Row, mconIntColSpec药品ID))
        If .EOF Then
            MsgBox "发生内部错误！", vbInformation, gstrSysName
            Exit Function
        End If
        
        If ((int分批 = 3 And int库房 <> 3) Or (int分批 = 1 And int库房 = 1) Or (int分批 = 2 And int库房 = 2)) And In_编辑状态 = 2 Then
            With RecStock
                If .RecordCount <> 0 Then .MoveFirst
                .Find "批次=" & Val(vsf批次.TextMatrix(vsf批次.Row, mconIntCol批次))
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
    Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[提取该药品的零售单位价格]", CLng(RecPhysic!药品ID))
    
    dblPrice = 0
    If Not rsTemp.EOF Then
        dblPrice = NVL(rsTemp!现价, 0)
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
        !药品来源 = RecPhysic!来源
        !基本药物 = RecPhysic!基本药物
        !通用名 = RecPhysic!通用名称
        !药典ID = RecPhysic!药典ID
        !用途分类ID = RecPhysic!用途分类ID
        !剂量单位 = RecPhysic!剂量单位
        !药品编码 = RecPhysic!药品编码
        !商品名 = IIf(IsNull(RecPhysic!商品名), "", RecPhysic!商品名)
        !规格 = RecPhysic!规格
        !产地 = RecPhysic!产地
        !药名ID = RecPhysic!药名ID
        !药品ID = RecPhysic!药品ID
        !售价 = dblPrice
        !售价单位 = RecPhysic!售价单位
        !剂量系数 = RecPhysic!售价包装
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
            If vsf批次.TextMatrix(vsf批次.Row, mconIntCol批号) = "新增批次药品" Then
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
                    !产地 = NVL(RecStock!产地)
                    !上次供应商ID = NVL(RecStock!上次供应商ID, 0)
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
            Set rsTemp = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[提取不分批药品的批号与效期信息]", Lng源库房ID, CLng(RecPhysic!药品ID))
            
            If rsTemp.RecordCount <> 0 Then
                !批号 = NVL(rsTemp!上次批号)
                If Not IsNull(rsTemp!生产日期) Then
                    !生产日期 = NVL(rsTemp!生产日期)
                End If
                !上次供应商ID = NVL(rsTemp!上次供应商ID, 0)
                If Not IsNull(rsTemp!效期) Then
                    !效期 = rsTemp!效期
                    If gtype_UserSysParms.P149_效期显示方式 = 1 And NVL(!效期) <> "" Then
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
        !成本价 = IIf(Val(RecPhysic!上次采购价) = 0, Val(RecPhysic!指导批发价), Val(RecPhysic!上次采购价))
        Select Case mintUnit
        Case mconint门诊单位
            !成本价 = Val(!成本价) / Val(RecPhysic!门诊包装)
        Case mconint住院单位
            !成本价 = Val(!成本价) / Val(RecPhysic!住院包装)
        Case mconint药库单位
            !成本价 = Val(!成本价) / Val(RecPhysic!药库包装)
        End Select
        
        .Update
    End With
    
    CombinateRec = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckData() As Boolean
    Dim DblCurStock As Double       '当前库存数
    '检测是否允许选择
    CheckData = False
    
    If Cmd确定.Enabled = False Then Exit Function
    
    'lng供应商ID不为零，表示退货，无库存时不准继续
    If lng供应商ID <> 0 Then
        If vsf批次.Visible Then
            If Val(vsf批次.TextMatrix(vsf批次.Row, mconIntCol上次供应商ID)) <> 0 And lng供应商ID <> Val(vsf批次.TextMatrix(vsf批次.Row, mconIntCol上次供应商ID)) Then
                MsgBox "你选择的退货商不是该药品的供应商，不能继续操作！", vbInformation, gstrSysName
                Exit Function
            End If
        End If
    End If
    
    If vsf批次.Visible Then
        If blnStock Then
            DblCurStock = Val(vsf批次.TextMatrix(vsf批次.Row, mconIntCol可用数量))
        Else
            DblCurStock = Get可用库存(Val(vsf药品规格.TextMatrix(vsf药品规格.Row, mconIntColSpec药品ID)), Val(vsf批次.TextMatrix(vsf批次.Row, mconIntCol批次)))
        End If
    Else
        If Not RecPhysic.EOF Then
            If blnStock Then
                DblCurStock = Val(vsf药品规格.TextMatrix(vsf药品规格.Row, mconIntColSpec可用数量))
            Else
                DblCurStock = Get可用库存(Val(vsf药品规格.TextMatrix(vsf药品规格.Row, mconIntColSpec药品ID)))
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
    
    If vsf批次.Visible Or bln时价 Then
        If (DblCurStock > 0) Or Not blnPrice Or vsf批次.TextMatrix(vsf批次.Row, mconIntCol批号) = "新增批次药品" Then CheckData = True: Exit Function
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
                    Optional ByVal 领用方式 As Integer = 0, Optional ByVal bln包含停用药品 As Boolean = False, Optional ByVal bln忽略服务对象 As Boolean = False) As ADODB.Recordset
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
        mbln忽略服务对象 = bln忽略服务对象
        .Show 1, FrmMain
    End With
    Set ShowME = RecReturn.Clone
End Function

Public Function Get可用库存(ByVal lng药品ID As Long, Optional ByVal lng批次 As Long = 0) As Single
    Dim rsStock As New ADODB.Recordset
    
    On Error GoTo errHandle
    gstrSQL = " Select Sum(A.可用数量" & StrUnitString & ") 可用数量,Sum(A.实际数量" & StrUnitString & ") 实际数量,sum(A.实际金额) 实际金额,sum(A.实际差价) 实际差价,Sum(A.实际数量) 库存数量 " & _
              " From 药品库存 A,药品规格 B " & _
              " Where A.药品ID=B.药品ID And A.性质=1 And A.药品ID=[1] " & IIf(lng批次 = 0, "", " And Nvl(A.批次,0)=[2] ")
    If Lng源库房ID <> 0 Or Lng目库房ID <> 0 Then
        gstrSQL = gstrSQL & " And A.库房ID=[3]"
    End If
    gstrSQL = gstrSQL & " Group By A.药品id"
    
    Set rsStock = zldatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[获取可用库存]", lng药品ID, lng批次, IIf(Lng源库房ID = 0, Lng目库房ID, Lng源库房ID))
    
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
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
