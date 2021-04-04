VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frm处方发药列表 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3060
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5970
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame fraColSel 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4200
      TabIndex        =   0
      Top             =   300
      Width           =   195
      Begin VB.Image imgColSel 
         Height          =   195
         Left            =   0
         Picture         =   "frm处方发药列表.frx":0000
         ToolTipText     =   "选择需要显示的列(ALT+C)"
         Top             =   0
         Width           =   195
      End
   End
   Begin MSComctlLib.ImageList imgCheck 
      Left            =   4680
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16777215
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":054E
            Key             =   ""
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":0AE8
            Key             =   ""
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":1082
            Key             =   ""
            Object.Tag             =   "3"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   5280
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   42
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":11DC
            Key             =   "打印11"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":1576
            Key             =   "当前"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":7DD8
            Key             =   "指示器"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":E63A
            Key             =   "附件"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":EBD4
            Key             =   "报告"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":EF6E
            Key             =   "标志"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":F308
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":F6A2
            Key             =   "图标"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":FA3C
            Key             =   "选择"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":1044E
            Key             =   "Person"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":16CB0
            Key             =   "未检"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":1D512
            Key             =   "在检"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":23D74
            Key             =   "已检"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":2A5D6
            Key             =   "Family"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":30E38
            Key             =   "分类"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":311D2
            Key             =   "分类_选中"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":3156C
            Key             =   "套餐"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":37DCE
            Key             =   "类型"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":3E630
            Key             =   "照片"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":44E92
            Key             =   "参数"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":4B6F4
            Key             =   "指标"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":51F56
            Key             =   "体检"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":587B8
            Key             =   "病历样式"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":5F01A
            Key             =   "病历文件"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":6587C
            Key             =   "规则"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":6C0DE
            Key             =   "收费"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":6CAF0
            Key             =   "诊断"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":73352
            Key             =   "创建"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":79BB4
            Key             =   "确认"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":80416
            Key             =   "开始"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":86C78
            Key             =   "结束"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":8D4DA
            Key             =   "部份"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":8D874
            Key             =   "全部"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":8DC0E
            Key             =   "部份总检"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":8DFA8
            Key             =   "全部总检"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":8E342
            Key             =   "总检"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":8E6DC
            Key             =   "打印"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":8F0EE
            Key             =   "已经打印"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":8FB00
            Key             =   "呼叫"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":9009A
            Key             =   "未取药"
         EndProperty
         BeginProperty ListImage41 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":90634
            Key             =   ""
         EndProperty
         BeginProperty ListImage42 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm处方发药列表.frx":96E96
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfColSel 
      Height          =   1095
      Left            =   1920
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   1470
      _cx             =   2593
      _cy             =   1931
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
      BackColorFixed  =   8421504
      ForeColorFixed  =   16777215
      BackColorSel    =   14737632
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   -1  'True
      AllowUserResizing=   0
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   0
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   2
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   250
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   $"frm处方发药列表.frx":9D6F8
      ScrollTrack     =   -1  'True
      ScrollBars      =   2
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
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   960
      Left            =   0
      TabIndex        =   2
      Top             =   240
      Width           =   1800
      _cx             =   3175
      _cy             =   1693
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
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483643
      GridColor       =   10329501
      GridColorFixed  =   10329501
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   16777215
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   5
      FixedRows       =   1
      FixedCols       =   1
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frm处方发药列表.frx":9D746
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
Attribute VB_Name = "frm处方发药列表"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mblnOutPut As Boolean
'列表显示条件
Private Type Type_ShowListCondition
    int列表类型 As Integer                          '0-配药确认,1-待配药;2-已配药;3-待发药;4-退药
    bln过滤模式 As Boolean
    bln是否呼叫 As Boolean
    bln是否签到确认 As Boolean
    bln取药确认 As Boolean                          '取药确认权限
    bln配药 As Boolean
    bln处方审查 As Boolean
End Type
Private mcondition As Type_ShowListCondition

Private mstrUnallowSetColHide As String             '不允许设置隐藏的列
Private mstrUnallowShow As String                   '不允许显示的列

Private mrsList As ADODB.Recordset
Private mIntOldRow As Integer

Private mintLocate As Integer
Private mstrFindType As String
Private mstrFind As String
Private mblnSortByName As Boolean                   '判断是否按姓名排序
Public mstrLastName As String                       '上次发药的病人姓名
Private mstrLastNo As String                        '上次选择的NO
Private mblnFreshList As Boolean
Private mblnNoRefreshDetail As Boolean
Private mblnFindOver As Boolean

Private Type FindProcess
    FindType As String
    FindContent As String
    StartRow As Integer
End Type
Private mFindProcess As FindProcess

'处方类型：普通、儿科、急诊、精二、精一、麻醉
Private Enum 处方类型
    普通 = 0
    儿科 = 1
    急诊 = 2
    精二 = 3
    精一 = 4
    麻醉 = 5
End Enum

'用户定义的处方颜色，从注册表取的字符串，用;分隔
Private mstrUserRecipeColor As String

Private mint金额显示 As Integer     '金额显示方式：0-显示应收金额,1-显示实收金额,2-显示应收和实收金额
Private mbln取药确认 As Boolean       '是否启用病人实际取药确认模式：0-不启用，1-启用
Private mintShowBill配药 As Integer '0-显示所有配药单,1-只显示未打印的待配药单据,2-只显示已打印的待配药单据

Private mintMoneyDigit As Integer           '金额小数位数

'列表类型
Private Enum mListType
    配药确认 = 0
    待配药 = 1
    已配药 = 2
    待发药 = 3
    超时未发 = 4
    退药 = 5
End Enum

Private Enum mFindType
    单据号 = 1
    门诊号 = 2
    姓名 = 3
    身份证 = 4
    IC卡 = 5
    医保号 = 6
    住院号 = 7
End Enum

Private Const mconIntCol列数 = 35
Private mIntCol当前行 As Integer
Private mintcol选择 As Integer
Private mIntCol审核 As Integer
Private mIntCol呼叫 As Integer
Private mIntCol颜色 As Integer
Private mIntCol处方类型 As Integer
Private mIntCol标志 As Integer
Private mIntCol类型 As Integer
Private mIntCol单据 As Integer
Private mIntCol收费 As Integer
Private mIntCol配药人 As Integer
Private mIntColNO As Integer
Private mIntCol姓名 As Integer
Private mIntCol金额 As Integer
Private mIntCol日期 As Integer
Private mIntCol签到日期 As Integer
Private mIntCol可操作 As Integer
Private mIntCol说明 As Integer
Private mIntCol就诊卡号 As Integer
Private mIntCol门诊号 As Integer
Private mIntCol身份证 As Integer
Private mIntColIC卡 As Integer
Private mIntCol病人ID As Integer
Private mIntCol医保号 As Integer
Private mIntCol住院号 As Integer
Private mIntCol实收金额  As Integer
Private mIntCol门诊标志  As Integer
Private mIntCol记录性质  As Integer
Private mIntCol收费类别  As Integer
Private mIntCol拼音简码 As Integer
Private mIntCol五笔简码 As Integer
Private mIntCol排队状态 As Integer
Private mIntCol发药窗口 As Integer
Private mIntCol未取药 As Integer
Private mIntCol审查结果 As Integer


Public Function GetPrintObject(ByVal blnOutPut As Boolean) As Object
    mblnOutPut = blnOutPut
    If vsfList.rows = 1 Then
        Set GetPrintObject = Nothing
    Else
        Set GetPrintObject = vsfList
    End If
End Function

Public Sub SetFontSize(ByVal intFont As Integer)
    With vsfList
        .Font.Size = intFont
        Me.Font.Size = .Font.Size
        .Cell(flexcpFontSize, 0, 0, .rows - 1, .Cols - 1) = .Font.Size
        
        .RowHeightMin = TextHeight("刘") + 100
        .RowHeightMax = TextHeight("刘") + 100
        .Refresh
    End With
End Sub
Public Function FindSpecialRow(ByVal strFindType As String, ByVal strFindContent As String, Optional strNos As String, Optional ByRef objSquareCard As Object, Optional ByVal str姓名 As String) As Boolean
    '如果是银行卡，则strFindContent格式为：卡ID|卡号
    'strNOs返回当前找到的行的病人在列表中的所有处方号：单据,NO|单据,NO
    Dim intCol As Integer
    Dim intFindRow As Integer
    Dim strNo As String
    Dim lng病人ID As Long
    Dim strName As String
    Dim intCount As Integer
    
    mblnFindOver = True
    
    With mFindProcess
        .FindType = strFindType
        .FindContent = UCase(strFindContent)
        .StartRow = 1
    End With
    
    With vsfList
        Select Case strFindType
            Case "姓名"
                intCol = mIntCol姓名
                
                If zlCommFun.IsCharAlpha(mFindProcess.FindContent) Then
                    '全字母时匹配简码
                    If zldatabase.GetPara("简码方式") = 0 Then
                        intCol = mIntCol拼音简码
                    Else
                        intCol = mIntCol五笔简码
                    End If
                End If
            Case "单据号"
                intCol = mIntColNO
            Case "门诊号"
                intCol = mIntCol门诊号
            Case "身份证"
                intCol = mIntCol身份证
            Case "IC卡"
                intCol = mIntCol病人ID
            Case "医保号"
                intCol = mIntCol医保号
            Case "住院号"
                intCol = mIntCol住院号
            Case Else
                '其余为消费卡，按病人ID查找
                intCol = mIntCol病人ID
                
                If InStr(1, strFindContent, "|") >= 1 Then
                    mFindProcess.FindContent = zlfuncCard_GetPatiID(objSquareCard, Val(Split(strFindContent, "|")(0)), Split(strFindContent, "|")(1))
                End If
        End Select
        
        If str姓名 <> "" Then
            Do Until mFindProcess.StartRow + 1 >= .rows
                mFindProcess.StartRow = .FindRow(mFindProcess.FindContent, mFindProcess.StartRow, intCol)
                
                If mFindProcess.StartRow = -1 Then Exit Do
                If .TextMatrix(mFindProcess.StartRow, mIntCol姓名) = str姓名 Then Exit Do
                
                mFindProcess.StartRow = mFindProcess.StartRow + 1
            Loop
        Else
            mFindProcess.StartRow = .FindRow(mFindProcess.FindContent, mFindProcess.StartRow, intCol)
        End If
        
        If mFindProcess.StartRow > 0 Then
            .Row = mFindProcess.StartRow
            .TopRow = .Row
            FindSpecialRow = True
            strNo = .TextMatrix(.Row, mIntColNO)
            lng病人ID = Val(.TextMatrix(.Row, mIntCol病人ID))
            strName = .TextMatrix(.Row, mIntCol姓名)
            
            If mFindProcess.StartRow + 1 >= .rows Then
                mFindProcess.StartRow = 1
            Else
                mFindProcess.StartRow = mFindProcess.StartRow + 1
            End If
        Else
            mFindProcess.StartRow = 1
        End If
                
        If strNo <> "" Then
            For intCount = 1 To .rows - 1
                If lng病人ID > 0 Then
                    If Val(.TextMatrix(intCount, mIntCol病人ID)) = lng病人ID Then
                        strNos = IIf(strNos = "", "", strNos & "|") & .TextMatrix(intCount, mIntCol单据) & "," & .TextMatrix(intCount, mIntColNO)
                    End If
                Else
                    If .TextMatrix(intCount, mIntCol姓名) = strName Then
                        strNos = IIf(strNos = "", "", strNos & "|") & .TextMatrix(intCount, mIntCol单据) & "," & .TextMatrix(intCount, mIntColNO)
                    End If
                End If
            Next
        End If
    End With
    
    mblnFindOver = False
End Function

Private Sub FindNextPati(ByVal blnFirst As Boolean)
'    Static intStar As Integer
'    Dim n As Integer
'    Dim strFind As String
'    Dim blnDo As Boolean
'
'    If BlnFirst Then intStar = 1
'
'    If Trim(txtFind.Text) = "" Then Exit Sub
'
'    strFind = Trim(txtFind.Text)
'
'    With Msf列表
'        If .Rows < 2 Then Exit Sub
'
'        For n = intStar To .Rows - 1
'            Select Case lblFind.Tag
'                Case FindType.就诊卡
'                    If Trim(.TextMatrix(n, 处方列名.就诊卡号)) = strFind Then blnDo = True
'                Case FindType.门诊号
'                    If Trim(.TextMatrix(n, 处方列名.门诊号)) = strFind Then blnDo = True
'                Case FindType.单据号
'                    If Trim(.TextMatrix(n, 处方列名.NO)) = strFind Then blnDo = True
'                Case FindType.姓名
'                    If mblnCard = True Then
'                        If Trim(.TextMatrix(n, 处方列名.就诊卡号)) = strFind Then blnDo = True
'                    Else
'                        If gbytCode = 1 Then
'                            If Trim(.TextMatrix(n, 处方列名.姓名)) Like "*" & strFind & "*" Or mWBX(Trim(.TextMatrix(n, 处方列名.姓名)), 1) Like "*" & UCase(strFind) & "*" Then blnDo = True
'                        Else
'                            If Trim(.TextMatrix(n, 处方列名.姓名)) Like "*" & strFind & "*" Or mPinYin(Trim(.TextMatrix(n, 处方列名.姓名))) Like "*" & UCase(strFind) & "*" Then blnDo = True
'                        End If
'                    End If
'                Case FindType.身份证
'                    If Trim(.TextMatrix(n, 处方列名.身份证)) = strFind Then blnDo = True
'                Case FindType.IC卡
'                    If Trim(.TextMatrix(n, 处方列名.IC卡)) = strFind Then blnDo = True
'            End Select
'
'            If blnDo Then
'                txtFind.Tag = txtFind.Text
'                .Row = n
'                Call Msf列表_EnterCell
'                .TopRow = n
'                intStar = n + 1
'                If intStar > .Rows - 1 Then intStar = .Rows - 1
'                Exit Sub
'            End If
'        Next
'    End With
'    intStar = 1
'    txtFind.SetFocus
    
End Sub
Public Function GetCurrentRecipe() As String
    '取当前处方
    '返回：0单据|1NO|2日期|3病人ID|4记录性质|5门诊标志|6处方类型|7收费类型|8病人姓名|9发药窗口|10呼叫|11未取药|12行号
    
    With vsfList
        If .Row = 0 Then Exit Function
        If Val(.TextMatrix(.Row, mIntCol单据)) = 0 Then Exit Function

        GetCurrentRecipe = .TextMatrix(.Row, mIntCol单据) & "|" & _
            .TextMatrix(.Row, mIntColNO) & "|" & _
            .TextMatrix(.Row, mIntCol日期) & "|" & _
            .TextMatrix(.Row, mIntCol病人ID) & "|" & _
            .TextMatrix(.Row, mIntCol记录性质) & "|" & _
            .TextMatrix(.Row, mIntCol门诊标志) & "|" & _
            .TextMatrix(.Row, mIntCol处方类型) & "|" & _
            .TextMatrix(.Row, mIntCol收费类别) & "|" & _
            .TextMatrix(.Row, mIntCol姓名) & "|" & _
            .TextMatrix(.Row, mIntCol发药窗口) & "|" & _
            .TextMatrix(.Row, mIntCol呼叫) & "|" & _
            .TextMatrix(.Row, mIntCol未取药) & "|" & _
            .Row
    End With
End Function

Public Function GetCurrentBatchRecipe() As String
    '发药时提取当前所选处方
    '返回：单据,NO,病人ID,金额,未审核,记录性质,门诊标志|单据,NO,病人ID,实收金额,未审核,记录性质,门诊标志
    Dim i As Integer
    Dim strRecipe As String
    
    If mblnFreshList = True Then Exit Function
    
    With vsfList
        If mcondition.bln过滤模式 = False Then
            If .TextMatrix(.Row, mIntColNO) <> "" Then
                strRecipe = .TextMatrix(.Row, mIntCol单据) & "," & _
                            .TextMatrix(.Row, mIntColNO) & "," & _
                            .TextMatrix(.Row, mIntCol病人ID) & "," & _
                            .TextMatrix(.Row, mIntCol实收金额) & "," & _
                            .TextMatrix(.Row, mIntCol收费) & "," & _
                            .TextMatrix(.Row, mIntCol记录性质) & "," & _
                            .TextMatrix(.Row, mIntCol门诊标志) & "," & _
                            .TextMatrix(.Row, mIntCol收费类别)
            End If
        Else
            For i = 1 To .rows - 1
                If .TextMatrix(i, mIntColNO) <> "" And Val(.TextMatrix(i, mIntCol标志)) = 1 Then
                    strRecipe = IIf(strRecipe = "", "", strRecipe & "|") & _
                        .TextMatrix(i, mIntCol单据) & "," & _
                        .TextMatrix(i, mIntColNO) & "," & _
                        .TextMatrix(i, mIntCol病人ID) & "," & _
                        .TextMatrix(i, mIntCol实收金额) & "," & _
                        .TextMatrix(i, mIntCol收费) & "," & _
                        .TextMatrix(i, mIntCol记录性质) & "," & _
                        .TextMatrix(i, mIntCol门诊标志) & "," & _
                        .TextMatrix(.Row, mIntCol收费类别)
                End If
            Next
        End If
        GetCurrentBatchRecipe = strRecipe
    End With
End Function
Sub InitList(ByVal intType As Integer)
    Dim i As Integer
    Dim n As Integer
    Dim str列设置 As String
    Dim arr列设置
    Dim bln列设置无效 As Boolean
    
    '''初始化列顺序
    '默认列顺序
    mIntCol当前行 = 0
    mintcol选择 = 1
    mIntCol审核 = 2
    mIntCol呼叫 = 3
    mIntCol颜色 = 4
    mIntCol处方类型 = 5
    mIntCol类型 = 6
    mIntColNO = 7
    mIntCol姓名 = 8
    mIntCol金额 = 9
    mIntCol实收金额 = 10
    mIntCol日期 = 11
    mIntCol签到日期 = 12
    mIntCol可操作 = 13
    mIntCol说明 = 14
    mIntCol就诊卡号 = 15
    mIntCol门诊号 = 16
    mIntCol身份证 = 17
    mIntColIC卡 = 18
    mIntCol病人ID = 19
    mIntCol医保号 = 20
    mIntCol住院号 = 21
    mIntCol标志 = 22
    mIntCol单据 = 23
    mIntCol收费 = 24
    mIntCol配药人 = 25
    mIntCol门诊标志 = 26
    mIntCol记录性质 = 27
    mIntCol收费类别 = 28
    mIntCol拼音简码 = 29
    mIntCol五笔简码 = 30
    mIntCol排队状态 = 31
    mIntCol发药窗口 = 32
    mIntCol未取药 = 33
    mIntCol审查结果 = 34
    
    '恢复用户自定义列顺序
    str列设置 = LoadListColState
    If str列设置 <> "" Then
        arr列设置 = Split(str列设置, "|")
        If UBound(arr列设置) + 1 <> mconIntCol列数 Then
            str列设置 = ""
        Else
            For n = 0 To UBound(arr列设置)
                If Split(arr列设置(n), ",")(0) = "" Then
                    bln列设置无效 = True
                    Exit For
                End If
            Next
            
            If bln列设置无效 = True Then
                str列设置 = ""
            Else
                For n = 0 To UBound(arr列设置)
                    SetColumnValue Split(arr列设置(n), ",")(0), n
                Next
            End If
        End If
    End If
     
    '初始化未发药清单
    With vsfList
        .Redraw = flexRDNone
        
        .rows = 1
        .rows = 2
        .Cols = mconIntCol列数
        .ExplorerBar = IIf(intType = mListType.待发药 And mcondition.bln过滤模式 = True, flexExNone, flexExSortShowAndMove)
        
        .Cell(flexcpPicture, 1, mIntCol当前行, 1, mIntCol当前行) = Me.imgList.ListImages(2).Picture
        .Cell(flexcpPictureAlignment, 1, mIntCol当前行, .rows - 1, mIntCol当前行) = flexPicAlignRightCenter
        
        VsfGridColFormat vsfList, mIntCol当前行, "", 250, flexAlignCenterCenter, "当前行"
        
        VsfGridColFormat vsfList, mintcol选择, "", IIf(intType = mListType.待发药 And mcondition.bln过滤模式 = True, 300, 0), flexAlignCenterCenter, "选择"
        VsfGridColFormat vsfList, mIntCol审核, "审", IIf((intType = mListType.待发药 Or intType = mListType.待配药) And mcondition.bln处方审查 = True, 300, 0), flexAlignCenterCenter, "审核"
        VsfGridColFormat vsfList, mIntCol呼叫, "呼叫", 500, flexAlignCenterCenter, "呼叫"
        VsfGridColFormat vsfList, mIntCol颜色, "类型", 500, flexAlignCenterCenter, "类型"
        VsfGridColFormat vsfList, mIntCol处方类型, "处方类型", 0, flexAlignCenterCenter, "处方类型"
        VsfGridColFormat vsfList, mIntCol标志, "1", 0, flexAlignCenterCenter, "标志"
        VsfGridColFormat vsfList, mIntCol类型, "类别", 1000, flexAlignLeftCenter, "类别"
        VsfGridColFormat vsfList, mIntCol单据, "单据", 0, flexAlignCenterCenter, "单据"
        VsfGridColFormat vsfList, mIntCol收费, "收费", 0, flexAlignCenterCenter, "收费"
        VsfGridColFormat vsfList, mIntCol配药人, "配药人", 0, flexAlignCenterCenter, "配药人"
        
        If mbln取药确认 = True Or intType = mListType.待配药 Then
            VsfGridColFormat vsfList, mIntColNO, "NO", 1100, flexAlignRightCenter, "NO"
        Else
            VsfGridColFormat vsfList, mIntColNO, "NO", 800, flexAlignLeftCenter, "NO"
        End If
        
        VsfGridColFormat vsfList, mIntCol姓名, "姓名", 800, flexAlignLeftCenter, "姓名"
        
        VsfGridColFormat vsfList, mIntCol金额, "应收金额", IIf(mint金额显示 = 1, 0, 1000), flexAlignRightCenter, "应收金额"
        VsfGridColFormat vsfList, mIntCol实收金额, "实收金额", IIf(mint金额显示 = 0, 0, 1000), flexAlignRightCenter, "实收金额"
        VsfGridColFormat vsfList, mIntCol日期, "日期", 1500, flexAlignLeftCenter, "日期"
        VsfGridColFormat vsfList, mIntCol签到日期, "签到日期", 1500, flexAlignLeftCenter, "签到日期"
        VsfGridColFormat vsfList, mIntCol可操作, "可操作", 0, flexAlignCenterCenter, "可操作"
        VsfGridColFormat vsfList, mIntCol说明, "说明", 1500, flexAlignLeftCenter, "说明"
        VsfGridColFormat vsfList, mIntCol就诊卡号, "就诊卡号", 1000, flexAlignLeftCenter, "就诊卡号"
        VsfGridColFormat vsfList, mIntCol门诊号, "门诊号", 1000, flexAlignLeftCenter, "门诊号"
        VsfGridColFormat vsfList, mIntCol身份证, "身份证", 1600, flexAlignLeftCenter, "身份证"
        VsfGridColFormat vsfList, mIntColIC卡, "IC卡", 1600, flexAlignLeftCenter, "IC卡"
        VsfGridColFormat vsfList, mIntCol病人ID, "病人ID", 0, flexAlignCenterCenter, "病人ID"
        VsfGridColFormat vsfList, mIntCol医保号, "医保号", 1500, flexAlignLeftCenter, "医保号"
        VsfGridColFormat vsfList, mIntCol住院号, "住院号", 1000, flexAlignLeftCenter, "住院号"
        
        VsfGridColFormat vsfList, mIntCol门诊标志, "门诊标志", 0, flexAlignCenterCenter, "门诊标志"
        VsfGridColFormat vsfList, mIntCol记录性质, "记录性质", 0, flexAlignCenterCenter, "记录性质"
        VsfGridColFormat vsfList, mIntCol收费类别, "收费类型", 0, flexAlignCenterCenter, "收费类型"
        VsfGridColFormat vsfList, mIntCol拼音简码, "拼音简码", 0, flexAlignCenterCenter, "拼音简码"
        VsfGridColFormat vsfList, mIntCol五笔简码, "五笔简码", 0, flexAlignCenterCenter, "五笔简码"
        VsfGridColFormat vsfList, mIntCol排队状态, "排队状态", 0, flexAlignCenterCenter, "排队状态"
        VsfGridColFormat vsfList, mIntCol发药窗口, "发药窗口", 0, flexAlignCenterCenter, "发药窗口"
        VsfGridColFormat vsfList, mIntCol未取药, "未取药", 0, flexAlignCenterCenter, "未取药"
        VsfGridColFormat vsfList, mIntCol审查结果, "审查结果", 0, flexAlignCenterCenter, "审查结果"
        
        mstrUnallowSetColHide = "NO"
        mstrUnallowShow = "当前行;处方类型;标志;单据;收费;配药人;可操作;病人ID;未审核;门诊标志;记录性质;收费类型;拼音简码;五笔简码;排队状态;发药窗口;未取药;审查结果"
        If mint金额显示 = 0 Then mstrUnallowShow = mstrUnallowShow & ";实收金额"
        If mint金额显示 = 1 Then mstrUnallowShow = mstrUnallowShow & ";应收金额"
        If mcondition.bln过滤模式 = False Then mstrUnallowShow = mstrUnallowShow & ";" & "选择"
        If mcondition.int列表类型 = mListType.退药 Or Not mcondition.bln是否签到确认 Then mstrUnallowShow = mstrUnallowShow & ";" & "签到日期"
        If mcondition.int列表类型 <> mListType.待发药 Or Not mcondition.bln是否呼叫 Then mstrUnallowShow = mstrUnallowShow & ";" & "呼叫"
        
        '恢复自定义列宽（不包括不允许显示的列）
        If str列设置 <> "" Then
            arr列设置 = Split(str列设置, "|")
            For n = 0 To UBound(arr列设置)
                If IsInString(mstrUnallowShow, Split(arr列设置(n), ",")(0), ";") = False Then
                    For i = 0 To vsfList.Cols - 1
                        If Split(arr列设置(n), ",")(0) = vsfList.ColKey(i) Then
                            vsfList.ColWidth(i) = Val(Split(arr列设置(n), ",")(1))
                        End If
                    Next
                End If
            Next
        End If
        
        If .ColWidth(mIntCol颜色) = 0 Then .ColWidth(mIntCol颜色) = 500
        
        If mcondition.int列表类型 = mListType.待发药 And mcondition.bln是否呼叫 Then
            .ColHidden(mIntCol呼叫) = False
        Else
            .ColHidden(mIntCol呼叫) = True
        End If
        
        If mcondition.int列表类型 <> mListType.退药 And mcondition.bln是否签到确认 Then
            .ColHidden(mIntCol签到日期) = False
        Else
            .ColHidden(mIntCol签到日期) = True
        End If
                
        .RowSel = 1
        
        .Redraw = flexRDDirect
    End With
End Sub

Public Sub SetPrintFlag(ByVal lngRow As Long)
    '在主界面调用打印配药单后，设置待配药列表中的打印图标
    If mcondition.int列表类型 <> mListType.待配药 Then Exit Sub
    If lngRow <= 0 Or lngRow > vsfList.rows - 1 Then Exit Sub
    
    vsfList.Redraw = flexRDNone
    
    If mintShowBill配药 = 1 Then
        vsfList.RemoveItem lngRow
        
        If lngRow <= vsfList.rows - 1 Then
            vsfList.Row = lngRow
        Else
            vsfList.Row = vsfList.rows - 1
        End If
        
        Call vsfList_EnterCell
    Else
        vsfList.Cell(flexcpPicture, lngRow, mIntColNO) = Me.imgList.ListImages("打印").Picture
        vsfList.Cell(flexcpPictureAlignment, lngRow, mIntColNO) = flexPicAlignLeftCenter
    End If
    
    vsfList.Redraw = flexRDDirect
End Sub
Private Sub SaveListColState(ByVal int类型 As Integer)
    Dim str列设置 As String
    Dim i As Integer
    Dim strType As String
    
    If Val(zldatabase.GetPara("使用个性化风格")) = 0 Then Exit Sub
    
    Select Case int类型
        Case mListType.配药确认
            strType = "配药确认"
        Case mListType.待配药
            strType = "待配药"
        Case mListType.已配药
            strType = "已配药"
        Case mListType.待发药
            strType = "待发药"
        Case mListType.超时未发
            strType = "超时未发"
        Case mListType.退药
            strType = "退药"
    End Select
    
    With vsfList
        For i = 0 To .Cols - 1
            str列设置 = IIf(str列设置 = "", "", str列设置 & "|") & vsfList.ColKey(i) & "," & .ColWidth(i)
        Next
    End With
    
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(vsfList), strType, str列设置)
End Sub

Private Function LoadListColState() As String
    Dim str列设置 As String
    Dim i As Integer
    Dim strType As String
    
    If Val(zldatabase.GetPara("使用个性化风格")) = 0 Then Exit Function
    
    Select Case mcondition.int列表类型
        Case mListType.配药确认
            strType = "配药确认"
        Case mListType.待配药
            strType = "待配药"
        Case mListType.已配药
            strType = "已配药"
        Case mListType.待发药
            strType = "待发药"
        Case mListType.超时未发
            strType = "超时未发"
        Case mListType.退药
            strType = "退药"
    End Select
    
    LoadListColState = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(vsfList), strType, "")
End Function

Private Sub SetMainComandBars(ByVal intListType As Integer, ByVal lngRow As Long)
    '根据当前记录清单类型及当前记录，设置主窗体的菜单状态
    Dim cbrControl As CommandBarControl
    Dim cbrMenu As CommandBarControl
    Dim bln允许取消 As Boolean
    Dim int门诊标志 As Integer
    Dim int记录性质 As Integer
    Dim Int单据 As Integer
    Dim strNo As String
    Dim blnAddSign As Boolean
    Dim blnVeirfySign As Boolean
    Dim int未取药 As Integer
    Dim int审核结果 As Integer
    Dim dateNow As Date
    Dim dblime As Double
    
    If lngRow = 0 Then Exit Sub
    
    int门诊标志 = Val(vsfList.TextMatrix(lngRow, mIntCol门诊标志))
    int记录性质 = Val(vsfList.TextMatrix(lngRow, mIntCol记录性质))
    Int单据 = Val(vsfList.TextMatrix(lngRow, mIntCol单据))
    strNo = vsfList.TextMatrix(lngRow, mIntColNO)
    int未取药 = Val(vsfList.TextMatrix(lngRow, mIntCol未取药))
    int审核结果 = Val(vsfList.TextMatrix(lngRow, mIntCol审查结果))
    
    '退药时和取消配药可以验证电子签名
    If intListType = mListType.退药 Or (intListType = mListType.待发药 And mcondition.bln配药) Then
        If gblnESign处方发药 = True Then
            blnAddSign = RecipeSendWork_JudgeSign(Int单据, strNo, Val(vsfList.TextMatrix(vsfList.Row, mIntCol可操作)), 0, CDate(vsfList.TextMatrix(vsfList.Row, mIntCol日期)))
            
            Set cbrMenu = frm药品处方发药New.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_VerifySign, , True)
            Set cbrControl = frm药品处方发药New.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_VerifySign, , True)

            If Not cbrMenu Is Nothing Then cbrMenu.Enabled = blnAddSign
            If Not cbrControl Is Nothing Then cbrControl.Enabled = blnAddSign
        End If
    End If
    
    '病人实际取药确认
    If intListType = mListType.退药 Then
        Set cbrMenu = frm药品处方发药New.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_TakeDrug, , True)
        Set cbrControl = frm药品处方发药New.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_TakeDrug, , True)

        If Not cbrMenu Is Nothing Then
            If mbln取药确认 = True And mcondition.bln取药确认 = True Then
                cbrMenu.Enabled = (int未取药 = 1)
            Else
                cbrMenu.Visible = False
            End If
        End If
        If Not cbrControl Is Nothing Then
            If mbln取药确认 = True And mcondition.bln取药确认 = True Then
                cbrControl.Enabled = (int未取药 = 1)
            Else
                cbrControl.Visible = False
            End If
        End If
    End If
    
    '未审核处方不能叫号
    If mcondition.bln是否呼叫 And intListType = mListType.待发药 Then
        Set cbrMenu = frm药品处方发药New.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Call, , True)
        Set cbrControl = frm药品处方发药New.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Call, , True)

        If Not cbrMenu Is Nothing Then
            If mcondition.bln处方审查 Then
                cbrMenu.Enabled = (int审核结果 = 1)
            Else
                cbrMenu.Enabled = True
            End If
        End If
        If Not cbrControl Is Nothing Then
            If mcondition.bln处方审查 Then
                cbrControl.Enabled = (int审核结果 = 1)
            Else
                cbrControl.Enabled = True
            End If
        End If
    End If
    
    '发药状态时，当单据已经超过3天，则不可以进行呼叫操作
    If intListType = mListType.待发药 Then
        dateNow = zldatabase.Currentdate
        dblime = dateNow - CDate(vsfList.TextMatrix(lngRow, mIntCol日期))
        Set cbrMenu = frm药品处方发药New.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Call, , True)
        Set cbrControl = frm药品处方发药New.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Recipe_Call, , True)
        If dblime > 3 Then
            If Not cbrMenu Is Nothing Then cbrMenu.Enabled = False
            If Not cbrControl Is Nothing Then cbrControl.Enabled = False
        Else
            If Not cbrMenu Is Nothing Then cbrMenu.Enabled = True
            If Not cbrControl Is Nothing Then cbrControl.Enabled = True
        End If
    End If
End Sub

Public Sub SetParams()
    mstrUserRecipeColor = zldatabase.GetPara("处方颜色", glngSys, 1341)
    If mstrUserRecipeColor = "" Then mstrUserRecipeColor = GetDefaultRecipeColor
    
    mcondition.bln取药确认 = IsInString(gstrprivs, "取药确认", ";")

    mint金额显示 = Val(zldatabase.GetPara("金额显示方式", glngSys, 1341, 0))
    mbln取药确认 = (Val(zldatabase.GetPara("启用病人实际取药确认模式", glngSys, 1341, 0)) = 1)
    mintShowBill配药 = Val(zldatabase.GetPara("待配药单据打印显示方式", glngSys, 1341, 0))
End Sub

Public Sub ShowList(ByVal intType As Integer, ByVal bln过滤模式 As Boolean, ByVal bln是否呼叫 As Boolean, ByVal bln是否签到确认 As Boolean, ByVal bln配药 As Boolean, ByVal bln处方审查 As Boolean, Optional ByVal strFindType As String = "", Optional ByVal strFind As String = "")
    vsfColSel.Visible = False
    
    If mcondition.int列表类型 <> intType Then
        mintLocate = 1
        mcondition.int列表类型 = intType
        mcondition.bln过滤模式 = bln过滤模式
    End If
    
    mcondition.bln是否呼叫 = bln是否呼叫
    mcondition.bln是否签到确认 = bln是否签到确认
    mcondition.bln配药 = bln配药
    mcondition.bln处方审查 = bln处方审查
    
    mstrFindType = strFindType
    mstrFind = strFind
    
    Call InitList(mcondition.int列表类型)
    
    Call InitColSelList(mcondition.int列表类型)
End Sub
Private Sub Form_Load()
    '取金额位数
    mintMoneyDigit = Val(zldatabase.GetPara("费用金额保留位数", glngSys, 0))
    
    Call SetParams
End Sub

Private Sub Form_Resize()
    vsfList.Move 0, 0, Me.Width, Me.Height
    
    fraColSel.Left = vsfList.ColWidth(0) - fraColSel.Width - 50
    fraColSel.Top = (vsfList.RowHeight(0) - fraColSel.Height) / 2 + 30
    fraColSel.ZOrder
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrLastNo = ""
    
    Call SaveListColState(mcondition.int列表类型)
    
    '没有启用个性化设置时删除用户排序设置
    On Error Resume Next
    If Val(zldatabase.GetPara("使用个性化风格")) = 0 Then
        DeleteSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品处方发药", "处方清单排序" & mListType.配药确认
        DeleteSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品处方发药", "处方清单排序" & mListType.待配药
        DeleteSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品处方发药", "处方清单排序" & mListType.已配药
        DeleteSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品处方发药", "处方清单排序" & mListType.待发药
        DeleteSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品处方发药", "处方清单排序" & mListType.超时未发
        DeleteSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品处方发药", "处方清单排序" & mListType.退药
    End If
End Sub
Private Sub imgColSel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim i As Long
    
    If Button = 1 Then '列选择器
        '根据当前状态直接确定勾选状态
        With vsfColSel
            If .Visible Then
                .Visible = False
                vsfList.SetFocus
            Else
                For i = .FixedRows To .rows - 1
                    If vsfList.ColHidden(.RowData(i)) Or vsfList.ColWidth(.RowData(i)) = 0 Then
                        .TextMatrix(i, 0) = 0
                    Else
                        .TextMatrix(i, 0) = 1
                    End If
                Next
                
                .Height = .RowHeightMin * .rows + 150
                .Top = fraColSel.Top + fraColSel.Height
                If .Top + .Height > Me.ScaleHeight - vsfList.Top Then
                    .Height = Me.ScaleHeight - .Top - vsfList.Top
                    .Width = 1750
                Else
                    .Width = 1470
                End If
                
                .Left = fraColSel.Left
                .ZOrder
                .Visible = True
                .SetFocus
            End If
        End With
    End If
End Sub


Private Sub vsfColSel_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngCol As Long
    
    If Col = 0 Then
        lngCol = vsfColSel.RowData(Row)
        If Val(vsfColSel.TextMatrix(Row, 0)) <> 0 Then
            vsfList.ColWidth(lngCol) = vsfList.ColData(lngCol)
            vsfList.ColHidden(lngCol) = False
        Else
            vsfList.ColWidth(lngCol) = 0
            vsfList.ColHidden(lngCol) = True
        End If
    End If
    
    Call SaveListColState(mcondition.int列表类型)
End Sub

Private Sub vsfColSel_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsfColSel
        If NewRow >= .FixedRows - 1 And NewCol >= .FixedCols - 1 Then
            .ForeColorSel = .Cell(flexcpForeColor, NewRow, 1)
            .Col = 0
        End If
    End With
End Sub


Private Sub vsfColSel_LostFocus()
    vsfColSel.Visible = False
End Sub

Private Sub vsfColSel_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> 0 Or vsfColSel.Cell(flexcpForeColor, Row, 1) = vsfColSel.BackColorFixed Then Cancel = True
End Sub


Private Sub vsfList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    Dim i As Integer
    
    '重设列选择列表
    Call InitColSelList(mcondition.int列表类型)
    
    '重设列顺序号
    For i = 0 To vsfList.Cols - 1
        Call SetColumnValue(vsfList.TextMatrix(0, i), i)
    Next
    
    '保存列表的状态
    Call SaveListColState(mcondition.int列表类型)
End Sub

Private Sub vsfList_AfterSort(ByVal Col As Long, Order As Integer)
    If Col = mIntCol姓名 Then
        mblnSortByName = True
    Else
        mblnSortByName = False
    End If
    
    Call vsfList_EnterCell

    '保存处方清单的用户排序规则
    '保存规则：
    '子项＝列表类型
    '值=列号|升/降序
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品处方发药", "处方清单排序" & mcondition.int列表类型, Col & "|" & Order)
End Sub


Private Sub vsfList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    '保存列表的状态
    Call SaveListColState(mcondition.int列表类型)
End Sub
Private Sub vsfList_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    '设置不能移动的列
    Select Case mcondition.int列表类型
        Case mListType.待配药, mListType.已配药, mListType.超时未发, mListType.配药确认
            If Col = mIntCol颜色 Then
                Position = mIntCol颜色
            End If

            If Col <> mIntCol颜色 And Position = mIntCol颜色 Then
                Position = Col
            End If
        Case mListType.待发药
            If Col = mIntCol颜色 Then
                Position = mIntCol颜色
            End If
            
            If Col = mintcol选择 Then
                Position = mintcol选择
            End If
            
            If Col = mIntCol呼叫 Then
                Position = mIntCol呼叫
            End If
            
            If (Col <> mIntCol颜色 And Position = mIntCol颜色) Or (Col <> mintcol选择 And Position = mintcol选择) Or (Col <> mIntCol呼叫 And Position = mIntCol呼叫) Then
                Position = Col
            End If
    End Select
End Sub

Private Sub vsfList_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '设置不能调整列宽的列
    Select Case mcondition.int列表类型
        Case mListType.待配药, mListType.已配药, mListType.超时未发
            If Col = mIntCol当前行 Or Col = mIntCol颜色 Then Cancel = True
        Case mListType.待发药
            If Col = mIntCol当前行 Or Col = mIntCol颜色 Or Col = mintcol选择 Or Col = mIntCol呼叫 Then Cancel = True
        Case Else
            If Col = 0 Then Cancel = True
    End Select
End Sub

Private Sub InitColSelList(ByVal intListType As Integer)
    Dim i As Integer
    
    With vsfColSel
        .Tag = intListType
        
        .rows = .FixedRows
        For i = 1 To vsfList.Cols - 1
            '不在不允许显示列表的列才能加入列选择列表
            If IsInString(mstrUnallowShow, vsfList.ColKey(i), ";") = False Then
                .rows = .rows + 1
                .TextMatrix(.rows - 1, 1) = vsfList.TextMatrix(0, i)
                .RowData(.rows - 1) = i
                
                '列宽为空或者隐藏的列设置为不勾选
                If Not (vsfList.ColWidth(i) = 0 Or vsfList.ColHidden(i)) Then
                    .TextMatrix(.rows - 1, 0) = 0
                End If
                
                '指定的列设置为不能设置隐藏
                If IsInString(mstrUnallowSetColHide, vsfList.ColKey(i), ";") = True Then
                    .Cell(flexcpForeColor, .rows - 1, 1) = .BackColorFixed
                End If
            End If
        Next
    End With
End Sub
Public Sub RefreshList(ByVal intType As Integer, ByVal rsData As ADODB.Recordset, Optional ByVal strNo As String, Optional ByVal blnNoRefreshDetail As Boolean)
    '创建处方列表数据集的副本
    Set mrsList = rsData
    Dim intRow As Integer
    Dim lngColor As Long
    Dim strSort As String
    Dim lngFindRow As Long
    Dim strFind As String
    Dim intFindCol As Integer
    Dim lngTime As Long
    Dim dateNow As Date
    
    
    mblnFreshList = True
    
    mblnNoRefreshDetail = blnNoRefreshDetail
    
    mcondition.bln过滤模式 = (Val(GetSetting("ZLSOFT", "公共模块\操作\" & App.ProductName & "\" & "药品处方发药", "界面定位", 0)) = 1)
    
    Call InitList(intType)
    
    With vsfList
        .Redraw = flexRDNone
        .MergeCells = flexMergeNever
        .rows = 1
        
        If mrsList.EOF Then
            .rows = 2
            .Cell(flexcpText, 1, mIntCol呼叫, 1, .Cols - 1) = "没有找到满足条件的记录......"
            .MergeCells = flexMergeRestrictRows
            .MergeRow(1) = True
            frm药品处方发药New.ClearForm_Detail
            frm药品处方发药New.ClearForm_Recipe
        Else
            Do While Not mrsList.EOF
                If intRow <> 0 And (.TextMatrix(intRow, mIntColNO) = mrsList!NO And .TextMatrix(intRow, mIntCol单据) = mrsList!单据) And mcondition.bln处方审查 And mcondition.int列表类型 <> mListType.退药 Then
                    If (mcondition.int列表类型 = mListType.待配药 Or mcondition.int列表类型 = mListType.待发药) Then
                        If Val(Nvl(mrsList!审查id)) <> 0 And .TextMatrix(intRow, mIntCol审查结果) <> mrsList!审查结果 Then
                            .TextMatrix(intRow, mIntCol审查结果) = 2
                        End If
                    End If
                Else
                intRow = intRow + 1
                .rows = intRow + 1
        
                .TextMatrix(intRow, mintcol选择) = ""

                If mcondition.int列表类型 = mListType.待发药 Then
                    If zlStr.Nvl(mrsList!呼叫时间) <> "" Then
'                        .Cell(flexcpPicture, intRow, mIntCol呼叫, intRow, mIntCol呼叫) = Me.imgList.ListImages(39).Picture
'                        .Cell(flexcpPictureAlignment, intRow, mIntCol呼叫, intRow, mIntCol呼叫) = flexPicAlignLeftCenter
                        .Cell(flexcpFontBold, intRow, mIntCol呼叫, intRow, mIntCol呼叫) = True
                        
                        dateNow = Sys.Currentdate
                        lngTime = DateDiff("n", mrsList!呼叫时间, dateNow)
                        If lngTime > 60 Then
                            .TextMatrix(intRow, mIntCol呼叫) = ">60"
                        Else
                            .TextMatrix(intRow, mIntCol呼叫) = IIf(lngTime < 0, 0, lngTime)
                        End If
                    End If
                End If
                
                If (mcondition.int列表类型 = mListType.待配药 Or mcondition.int列表类型 = mListType.待发药) And mcondition.bln处方审查 Then
                    If Val(Nvl(mrsList!审查id)) <> 0 Then
                        If mrsList!审查结果 = 1 Then
                            .Cell(flexcpPicture, intRow, mIntCol审核, intRow, mIntCol审核) = Me.imgList.ListImages(41).Picture
                             .TextMatrix(intRow, mIntCol审查结果) = 1
                        Else
                            .Cell(flexcpPicture, intRow, mIntCol审核, intRow, mIntCol审核) = Me.imgList.ListImages(42).Picture
                            .TextMatrix(intRow, mIntCol审查结果) = 2
                        End If
                    Else
                        .Cell(flexcpPicture, intRow, mIntCol审核, intRow, mIntCol审核) = Me.imgList.ListImages(41).Picture
                        .TextMatrix(intRow, mIntCol审查结果) = 1
                    End If
                End If
                
                If mcondition.int列表类型 <> mListType.退药 Then
                    .TextMatrix(intRow, mIntCol签到日期) = zlStr.Nvl(mrsList!签到时间)
                End If
                
                If mrsList!处方类型 = 1 Then
                    .TextMatrix(intRow, mIntCol颜色) = "儿科"
                ElseIf mrsList!处方类型 = 2 Then
                    .TextMatrix(intRow, mIntCol颜色) = "急诊"
                ElseIf mrsList!处方类型 = 3 Then
                    .TextMatrix(intRow, mIntCol颜色) = "精二"
                ElseIf mrsList!处方类型 = 4 Then
                    .TextMatrix(intRow, mIntCol颜色) = "精一"
                ElseIf mrsList!处方类型 = 5 Then
                    .TextMatrix(intRow, mIntCol颜色) = "麻醉"
                Else
                    .TextMatrix(intRow, mIntCol颜色) = "普通"
                End If

                .TextMatrix(intRow, mIntCol处方类型) = IIf(IsNull(mrsList!处方类型), "", mrsList!处方类型)
                .TextMatrix(intRow, mIntCol标志) = mrsList!标志
                .TextMatrix(intRow, mIntCol类型) = IIf(IsNull(mrsList!类型), "", mrsList!类型)
                
                .TextMatrix(intRow, mIntCol单据) = mrsList!单据
                .TextMatrix(intRow, mIntCol收费) = mrsList!已收费
                .TextMatrix(intRow, mIntCol配药人) = IIf(IsNull(mrsList!配药人), "", mrsList!配药人)
                
                .TextMatrix(intRow, mIntColNO) = mrsList!NO
                
                Select Case intType
                    Case mListType.待配药
                        If mrsList!打印状态 = 1 Then    '只有待配药环节记录集才有“打印状态”
                            .Cell(flexcpPicture, intRow, mIntColNO) = Me.imgList.ListImages("打印").Picture
                            .Cell(flexcpPictureAlignment, intRow, mIntColNO) = flexPicAlignLeftCenter
                        End If
                    
                    Case mListType.退药
                        If mbln取药确认 = True Then
                            .TextMatrix(intRow, mIntCol未取药) = zlStr.Nvl(mrsList!未取药, 0)
                            If Val(.TextMatrix(intRow, mIntCol未取药)) = 1 Then
                                .Cell(flexcpPicture, intRow, mIntColNO) = Me.imgList.ListImages("未取药").Picture
                                .Cell(flexcpPictureAlignment, intRow, mIntColNO) = flexPicAlignRightCenter
                            End If
                        End If
                End Select
               
                .TextMatrix(intRow, mIntCol姓名) = IIf(IsNull(mrsList!姓名), "", mrsList!姓名)
                
                .TextMatrix(intRow, mIntCol金额) = zlStr.FormatEx(Val(mrsList!金额), mintMoneyDigit, , True)
                .TextMatrix(intRow, mIntCol实收金额) = zlStr.FormatEx(Val(mrsList!实收金额), mintMoneyDigit, , True)
                .TextMatrix(intRow, mIntCol日期) = mrsList!日期
                .TextMatrix(intRow, mIntCol可操作) = mrsList!可操作
                .TextMatrix(intRow, mIntCol说明) = IIf(IsNull(mrsList!说明), "", mrsList!说明)
                .TextMatrix(intRow, mIntCol就诊卡号) = IIf(IsNull(mrsList!就诊卡号), "", mrsList!就诊卡号)
                
                .TextMatrix(intRow, mIntCol门诊号) = IIf(IsNull(mrsList!门诊号), "", mrsList!门诊号)
                .TextMatrix(intRow, mIntCol身份证) = IIf(IsNull(mrsList!身份证号), "", mrsList!身份证号)
                .TextMatrix(intRow, mIntColIC卡) = IIf(IsNull(mrsList!IC卡号), "", mrsList!IC卡号)
                .TextMatrix(intRow, mIntCol病人ID) = IIf(IsNull(mrsList!病人ID), "", mrsList!病人ID)
                .TextMatrix(intRow, mIntCol医保号) = IIf(IsNull(mrsList!医保号), "", mrsList!医保号)
                .TextMatrix(intRow, mIntCol住院号) = IIf(IsNull(mrsList!住院号), "", mrsList!住院号)
                
                .TextMatrix(intRow, mIntCol门诊标志) = mrsList!门诊标志
                .TextMatrix(intRow, mIntCol记录性质) = mrsList!记录性质
                .TextMatrix(intRow, mIntCol收费类别) = mrsList!收费类别
                
                .TextMatrix(intRow, mIntCol拼音简码) = mPinYin(IIf(IsNull(mrsList!姓名), "", mrsList!姓名))
                .TextMatrix(intRow, mIntCol五笔简码) = mWBX(IIf(IsNull(mrsList!姓名), "", mrsList!姓名), 1)
                
                If intType = mListType.配药确认 Then
                    .TextMatrix(intRow, mIntCol排队状态) = zlStr.Nvl(mrsList!排队状态)
                End If
                
                If intType <> mListType.退药 Then
                    .TextMatrix(intRow, mIntCol发药窗口) = zlStr.Nvl(mrsList!发药窗口)
                End If
                
                .Cell(flexcpBackColor, intRow, mIntCol颜色, intRow, mIntCol颜色) = Val(Split(mstrUserRecipeColor, ";")(Val(mrsList!处方类型)))
                
                '设置颜色
                lngColor = IIf(mcondition.int列表类型 <> mListType.退药 Or mrsList!可操作 = 0, &H80000008, IIf(mrsList!可操作 = 1, glng正常, IIf(mrsList!可操作 = 2, glng发药, glng退药)))
                .Cell(flexcpForeColor, intRow, 1, intRow, .Cols - 1) = lngColor
                .Cell(flexcpForeColor, intRow, mIntCol呼叫, intRow, mIntCol呼叫) = vbRed
                
                '病人类型用不同前景色，字体加粗区别
                .Cell(flexcpForeColor, intRow, mIntCol姓名, intRow, mIntCol姓名) = zldatabase.GetPatiColor(IIf(IsNull(mrsList!病人类型), "", mrsList!病人类型))
                End If
                mrsList.MoveNext
            Loop
            
            If mcondition.bln过滤模式 = True Then
                .Cell(flexcpPicture, 0, mintcol选择, .rows - 1, mintcol选择) = LoadResPicture("checked", vbResBitmap)
                .Cell(flexcpPictureAlignment, 0, mintcol选择, .rows - 1, mintcol选择) = flexAlignCenterCenter
                .Cell(flexcpText, 0, mIntCol标志, .rows - 1, mIntCol标志) = 1
            End If
            
            If mintLocate = 0 Or mintLocate > intRow Then mintLocate = 1
            If mcondition.bln过滤模式 Then
                .Row = 1
                .TopRow = .Row
            End If
        End If
        
        '查找状态下定位
        If mcondition.bln过滤模式 = False And mstrFind <> "" Then
            mFindProcess.StartRow = 1
            FindSpecialRow mstrFindType, mstrFind
        End If
        
        mblnSortByName = False
        
        '恢复用户排序规则
        '子项＝列表类型
        '值=列号|升/降序
        strSort = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品处方发药", "处方清单排序" & mcondition.int列表类型, "")
        If strSort <> "" And InStr(1, strSort, "|") > 0 Then    '值不能为空，并且要有分隔符
            If Val(Split(strSort, "|")(0)) > 0 And Val(Split(strSort, "|")(0)) < .Cols - 1 Then     '返回的列序号要在清单列范围内
                If .ColHidden(Val(Split(strSort, "|")(0))) = False Then      '返回的列必须不是隐藏
                    .ColSort(Val(Split(strSort, "|")(0))) = IIf(Val(Split(strSort, "|")(1)) = 2, 2, 1)
                    .Col = Val(Split(strSort, "|")(0))
                    .Sort = flexSortUseColSort
                    
                    If Val(Split(strSort, "|")(0)) = mIntCol姓名 Then
                        mblnSortByName = True
                    End If
                End If
            End If
        End If
        
        If mcondition.bln过滤模式 = False Then
            If mintLocate = 0 Or mintLocate > intRow Then mintLocate = 1
            
            If mblnSortByName = True And mstrLastName <> "" Then
                '按姓名排序时查找上次发药病人的下张单据
                strFind = mstrLastName
                intFindCol = mIntCol姓名
            ElseIf strNo <> "" Then
                strFind = strNo
                intFindCol = mIntColNO
                mintLocate = 1
            Else
                '按上次选择的NO查找
                strFind = mstrLastNo
                intFindCol = mIntColNO
                mintLocate = 1
            End If
            
            If strFind <> "" Then
                lngFindRow = .FindRow(strFind, mintLocate, intFindCol)
                If lngFindRow > 0 Then
'                    .Row = 0
                    .Row = lngFindRow
                Else
                    lngFindRow = .FindRow(strFind, 1, intFindCol)
                    If lngFindRow > 0 Then
'                        .Row = 0
                        .Row = lngFindRow
                    Else
'                        .Row = 0
                        .Row = mintLocate
                    End If
                End If
            Else
                If .rows > 1 Then .Row = 1
            End If
            .TopRow = .Row
        End If
        
        .Redraw = flexRDDirect
    End With
    
    mblnFreshList = False
End Sub

Private Sub SetColumnValue(ByVal str列名 As String, ByVal intValue As Integer)
    Select Case str列名
        Case "呼叫"
            mIntCol呼叫 = intValue
        Case "类型"
            mIntCol颜色 = intValue
        Case "选择"
            mintcol选择 = intValue
        Case "类别"
            mIntCol类型 = intValue
        Case "NO"
            mIntColNO = intValue
        Case "姓名"
            mIntCol姓名 = intValue
        Case "金额", "应收金额"
            mIntCol金额 = intValue
        Case "实收金额"
            mIntCol实收金额 = intValue
        Case "日期"
            mIntCol日期 = intValue
        Case "签到日期"
            mIntCol签到日期 = intValue
        Case "说明"
            mIntCol说明 = intValue
        Case "就诊卡号"
            mIntCol就诊卡号 = intValue
        Case "门诊号"
            mIntCol门诊号 = intValue
    
        Case "身份证"
            mIntCol身份证 = intValue
        Case "IC卡"
            mIntColIC卡 = intValue
        Case "病人ID"
            mIntCol病人ID = intValue
        Case "可操作"
            mIntCol可操作 = intValue
        Case "医保号"
            mIntCol医保号 = intValue
        Case "住院号"
            mIntCol住院号 = intValue
        Case "处方类型"
            mIntCol处方类型 = intValue
    End Select
End Sub

Private Sub vsfList_Click()
    Dim intCheck As Integer
    Dim strCheck As String
    
    With vsfList
        If mcondition.bln过滤模式 = False Then Exit Sub
        If .MouseRow < 0 Then Exit Sub
        If .MouseCol <> mintcol选择 Then Exit Sub
        
        If IsNumeric(.TextMatrix(.rows - 1, mIntCol单据)) Then
            intCheck = Abs(.Cell(flexcpText, .MouseRow, mIntCol标志, .MouseRow, mIntCol标志) - 1)
        Else
            intCheck = Abs(.Cell(flexcpText, 0, mIntCol标志, 0, mIntCol标志) - 1)
            .TextMatrix(0, mIntCol标志) = intCheck
        End If
        strCheck = IIf(intCheck = 1, "checked", "unchecked")

        If .MouseRow = 0 Then
            If IsNumeric(.TextMatrix(.rows - 1, mIntCol单据)) Then .Cell(flexcpText, 0, mIntCol标志, .rows - 1, mIntCol标志) = intCheck
            .Cell(flexcpPicture, 0, mintcol选择, .rows - 1, mintcol选择) = LoadResPicture(strCheck, vbResBitmap)
            .Cell(flexcpPictureAlignment, 0, mintcol选择, .rows - 1, mintcol选择) = flexAlignCenterCenter
        Else
            If IsNumeric(.TextMatrix(.rows - 1, mIntCol单据)) Then
                .Cell(flexcpText, .MouseRow, mIntCol标志, .MouseRow, mIntCol标志) = intCheck
            Else
                .Cell(flexcpPicture, 0, mintcol选择, 0, mintcol选择) = LoadResPicture(strCheck, vbResBitmap)
                .Cell(flexcpPictureAlignment, 0, mintcol选择, 0, mintcol选择) = flexAlignCenterCenter
            End If
            .Cell(flexcpPicture, .MouseRow, mintcol选择, .MouseRow, mintcol选择) = LoadResPicture(strCheck, vbResBitmap)
            .Cell(flexcpPictureAlignment, .MouseRow, mintcol选择, .MouseRow, mintcol选择) = flexAlignCenterCenter
        End If
    End With
End Sub

Private Sub vsfList_EnterCell()
    Dim lngColor As Long
    Dim lngNameColor As Long
    
    If mblnOutPut = True Then Exit Sub
    If mblnNoRefreshDetail = True Then Exit Sub
'    If mblnFindOver = True Then Exit Sub
    
    With vsfList
        If .Row = 0 Then Exit Sub
        If Not gobjPass Is Nothing Then Call gobjPass.zlPassClearLight_YF
        .Cell(flexcpPicture, 1, 0, .rows - 1, 0) = Nothing
        .Cell(flexcpPicture, .Row, 0, .Row, 0) = Me.imgList.ListImages(2).Picture
        
        lngColor = IIf(mcondition.int列表类型 <> mListType.退药 Or Val(.TextMatrix(.Row, mIntCol可操作)) <= 1, glng正常, IIf(Val(.TextMatrix(.Row, mIntCol可操作)) = 2, glng发药, glng退药))
        lngNameColor = .Cell(flexcpForeColor, .Row, mIntCol姓名, .Row, mIntCol姓名)
        
        '选中行时的前景色用病人类型颜色标识
        .ForeColorSel = lngNameColor
        
        If Val(.TextMatrix(.Row, mIntCol单据)) = 0 Then Exit Sub
        
        If mblnFreshList = False Then mstrLastNo = .TextMatrix(.Row, mIntColNO)
        
        mintLocate = .Row
        
        SetMainComandBars mcondition.int列表类型, .Row
        
        If mcondition.int列表类型 = mListType.退药 Then
            If Trim(.TextMatrix(.Row, mIntCol说明)) = "" Then
                frm药品处方发药New.RefreshDetail_Return Val(.TextMatrix(.Row, mIntCol单据)), .TextMatrix(.Row, mIntColNO), .TextMatrix(.Row, mIntCol日期), Val(.TextMatrix(.Row, mIntCol可操作)), Val(.TextMatrix(.Row, mIntCol门诊标志)), Val(.TextMatrix(.Row, mIntCol记录性质))
            Else
                frm药品处方发药New.RefreshDetail_Return Val(.TextMatrix(.Row, mIntCol单据)), .TextMatrix(.Row, mIntColNO), .TextMatrix(.Row, mIntCol日期), Val(.TextMatrix(.Row, mIntCol可操作)), Val(.TextMatrix(.Row, mIntCol门诊标志)), Val(.TextMatrix(.Row, mIntCol记录性质)), False, Val(Mid(.TextMatrix(.Row, mIntCol说明), InStr(1, .TextMatrix(.Row, mIntCol说明), "第") + 1, InStr(1, .TextMatrix(.Row, mIntCol说明), "次") - InStr(1, .TextMatrix(.Row, mIntCol说明), "第") - 1))
            End If
        Else
            frm药品处方发药New.RefreshDetail_Send 0, Val(.TextMatrix(.Row, mIntCol单据)), .TextMatrix(.Row, mIntColNO), Val(.TextMatrix(.Row, mIntCol门诊标志)), Val(.TextMatrix(.Row, mIntCol记录性质)), Val(.TextMatrix(.Row, mIntCol排队状态)), Val(.TextMatrix(.Row, mIntCol审查结果))
        End If
    End With
End Sub

Private Sub vsfList_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> mintcol选择 Then Cancel = True
End Sub

Public Sub SetCalling()
    With Me.vsfList
        If mIntOldRow <> 0 And mIntOldRow <= .rows - 1 Then
            If .Cell(flexcpText, mIntOldRow, mIntCol呼叫, mIntOldRow, mIntCol呼叫) = "" Then
                .Cell(flexcpText, mIntOldRow, mIntCol呼叫, mIntOldRow, mIntCol呼叫) = "0"
                .Cell(flexcpPicture, mIntOldRow, mIntCol呼叫, mIntOldRow, mIntCol呼叫) = Nothing
            End If
        End If
        mIntOldRow = .Row
        
        .Cell(flexcpText, .Row, mIntCol呼叫, .Row, mIntCol呼叫) = ""
        .Cell(flexcpPicture, .Row, mIntCol呼叫, .Row, mIntCol呼叫) = Me.imgList.ListImages(39).Picture
        .Cell(flexcpPictureAlignment, .Row, mIntCol呼叫, .Row, mIntCol呼叫) = flexPicAlignCenterCenter
    End With
End Sub

Public Sub SetSign(ByRef strNo As String)
    Dim i As Integer
    
    With Me.vsfList
        For i = 1 To .rows - 1
            If InStr(1, strNo, .TextMatrix(i, mIntColNO)) <> 0 Then
                strNo = Replace(strNo, .TextMatrix(.Row, mIntColNO), "")
                .Row = i
                Exit Sub
            End If
        Next
    End With
End Sub
