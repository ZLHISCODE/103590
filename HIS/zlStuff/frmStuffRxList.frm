VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmStuffRxList 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5760
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3600
   ScaleWidth      =   5760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.ImageList imgList 
      Left            =   4800
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   39
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":0000
            Key             =   "打印11"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":039A
            Key             =   "当前"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":6BFC
            Key             =   "指示器"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":D45E
            Key             =   "附件"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":D9F8
            Key             =   "报告"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":DD92
            Key             =   "标志"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":E12C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":E4C6
            Key             =   "图标"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":E860
            Key             =   "选择"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":F272
            Key             =   "Person"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":15AD4
            Key             =   "未检"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":1C336
            Key             =   "在检"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":22B98
            Key             =   "已检"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":293FA
            Key             =   "Family"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":2FC5C
            Key             =   "分类"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":2FFF6
            Key             =   "分类_选中"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":30390
            Key             =   "套餐"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":36BF2
            Key             =   "类型"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":3D454
            Key             =   "照片"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":43CB6
            Key             =   "参数"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":4A518
            Key             =   "指标"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":50D7A
            Key             =   "体检"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":575DC
            Key             =   "病历样式"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":5DE3E
            Key             =   "病历文件"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":646A0
            Key             =   "规则"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":6AF02
            Key             =   "收费"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":6B914
            Key             =   "诊断"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":72176
            Key             =   "创建"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":789D8
            Key             =   "确认"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":7F23A
            Key             =   "开始"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":85A9C
            Key             =   "结束"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":8C2FE
            Key             =   "部份"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":8C698
            Key             =   "全部"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":8CA32
            Key             =   "部份总检"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":8CDCC
            Key             =   "全部总检"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":8D166
            Key             =   "总检"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":8D500
            Key             =   "打印"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":8DF12
            Key             =   "已经打印"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffRxList.frx":8E924
            Key             =   "呼叫"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfColSel 
      Height          =   1455
      Left            =   4080
      TabIndex        =   0
      Top             =   1440
      Visible         =   0   'False
      Width           =   1470
      _cx             =   2593
      _cy             =   2566
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
      FormatString    =   $"frmStuffRxList.frx":8EEBE
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
      Height          =   2160
      Left            =   360
      TabIndex        =   1
      Top             =   600
      Width           =   2400
      _cx             =   4233
      _cy             =   3810
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
      FormatString    =   $"frmStuffRxList.frx":8EF0C
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
   Begin VB.Image imgSel 
      Height          =   195
      Left            =   4320
      Picture         =   "frmStuffRxList.frx":8EF81
      ToolTipText     =   "选择需要显示的列(ALT+C)"
      Top             =   840
      Width           =   195
   End
End
Attribute VB_Name = "frmStuffRxList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'定义表格列
Private Const mconIntCol列数 = 25
Private mIntCol当前行 As Integer
Private mIntCol处方类型 As Integer
Private mIntCol标志 As Integer
Private mIntCol类型 As Integer
Private mIntCol单据 As Integer
Private mIntCol收费 As Integer
Private mIntColNO As Integer
Private mIntCol姓名 As Integer
Private mIntCol金额 As Integer
Private mIntCol日期 As Integer
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
Private mIntCol库房ID  As Integer
Private mIntCol记录状态  As Integer

Private Const glng退药 As Long = &HC0&
Private Const glng发药 As Long = &HC00000
Private Const glng正常 As Long = &H80000008

'设置列的隐藏的变量
Private mstrUnallowSetColHide As String
Private mstrUnallowShow As String

'页面常量
Private Enum mListType
    待发料 = 0
    退料 = 1
End Enum

Private Enum mFindType
    单据号 = 0
    门诊号 = 1
    姓名 = 2
    身份证 = 3
    IC卡 = 4
    医保号 = 5
    住院号 = 6
End Enum

Private Type FindProcess
    FindType As Integer
    FindContent As String
    StartRow As Integer
End Type
Private mFindProcess As FindProcess

Private mintType As Integer               '当前页面类型
Private mlng库房id As Long
Private mFMT As g_FmtString

Private Sub Form_Load()
    Call InitVsfList
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    Me.vsfList.Move 0, 0, Me.Width, Me.Height

    err.Clear
End Sub

Private Sub InitVsfList()
    Dim i As Integer
    Dim n As Integer
    Dim str列设置 As String
    Dim arr列设置
    
    '''初始化列顺序
    '默认列顺序
    mIntCol当前行 = 0
    mIntCol处方类型 = 1
    mIntCol标志 = 2
    mIntCol类型 = 3
    mIntCol单据 = 4
    mIntCol收费 = 5
    mIntColNO = 6
    mIntCol姓名 = 7
    mIntCol金额 = 8
    mIntCol日期 = 9
    mIntCol可操作 = 10
    mIntCol说明 = 11
    mIntCol就诊卡号 = 12
    mIntCol门诊号 = 13
    mIntCol身份证 = 14
    mIntColIC卡 = 15
    mIntCol病人ID = 16
    mIntCol医保号 = 17
    mIntCol住院号 = 18
    mIntCol实收金额 = 19
    mIntCol门诊标志 = 20
    mIntCol记录性质 = 21
    mIntCol收费类别 = 22
    mIntCol库房ID = 23
    mIntCol记录状态 = 24
    
    '恢复用户自定义列顺序
    str列设置 = LoadListColState
    If str列设置 <> "" Then
        arr列设置 = Split(str列设置, "|")
        If UBound(arr列设置) + 1 <> mconIntCol列数 Then
            str列设置 = ""
        Else
            For n = 0 To UBound(arr列设置)
                SetColumnValue Split(arr列设置(n), ",")(0), n
            Next
        End If
    End If
     
    '初始化未发药清单
    With vsfList
        .Redraw = flexRDNone
        
        .Rows = 1
        .Rows = 2
        .Cols = mconIntCol列数
        
        .Cell(flexcpPicture, 1, mIntCol当前行, 1, mIntCol当前行) = Me.imgList.ListImages(2).Picture
        .Cell(flexcpPictureAlignment, 1, mIntCol当前行, .Rows - 1, mIntCol当前行) = flexPicAlignRightCenter
        
        VsfGridColFormat vsfList, mIntCol当前行, "", 250, flexAlignCenterCenter, "当前行"

        VsfGridColFormat vsfList, mIntCol处方类型, "处方类型", 0, flexAlignCenterCenter, "处方类型"
        VsfGridColFormat vsfList, mIntCol标志, "1", 0, flexAlignCenterCenter, "标志"
        VsfGridColFormat vsfList, mIntCol类型, "类别", 1000, flexAlignLeftCenter, "类别"
        VsfGridColFormat vsfList, mIntCol单据, "单据", 0, flexAlignCenterCenter, "单据"
        VsfGridColFormat vsfList, mIntCol收费, "收费", 0, flexAlignCenterCenter, "收费"
        VsfGridColFormat vsfList, mIntColNO, "NO", 800, flexAlignLeftCenter, "NO"
        VsfGridColFormat vsfList, mIntCol姓名, "姓名", 800, flexAlignLeftCenter, "姓名"
        
        VsfGridColFormat vsfList, mIntCol金额, "金额", 1200, flexAlignRightCenter, "金额"
        VsfGridColFormat vsfList, mIntCol日期, "日期", 1500, flexAlignLeftCenter, "日期"
        VsfGridColFormat vsfList, mIntCol可操作, "可操作", 0, flexAlignCenterCenter, "可操作"
        VsfGridColFormat vsfList, mIntCol说明, "说明", 1500, flexAlignLeftCenter, "说明"
        VsfGridColFormat vsfList, mIntCol就诊卡号, "就诊卡号", 1000, flexAlignLeftCenter, "就诊卡号"
        VsfGridColFormat vsfList, mIntCol门诊号, "门诊号", 1000, flexAlignLeftCenter, "门诊号"
        VsfGridColFormat vsfList, mIntCol身份证, "身份证", 1600, flexAlignLeftCenter, "身份证"
        VsfGridColFormat vsfList, mIntColIC卡, "IC卡", 1600, flexAlignLeftCenter, "IC卡"
        VsfGridColFormat vsfList, mIntCol病人ID, "病人ID", 0, flexAlignCenterCenter, "病人ID"
        VsfGridColFormat vsfList, mIntCol医保号, "医保号", 1500, flexAlignLeftCenter, "医保号"
        VsfGridColFormat vsfList, mIntCol住院号, "住院号", 1000, flexAlignLeftCenter, "住院号"
        
        VsfGridColFormat vsfList, mIntCol实收金额, "实收金额", 0, flexAlignCenterCenter, "实收金额"
        VsfGridColFormat vsfList, mIntCol门诊标志, "门诊标志", 0, flexAlignCenterCenter, "门诊标志"
        VsfGridColFormat vsfList, mIntCol记录性质, "记录性质", 0, flexAlignCenterCenter, "记录性质"
        VsfGridColFormat vsfList, mIntCol收费类别, "收费类型", 0, flexAlignCenterCenter, "收费类型"
        VsfGridColFormat vsfList, mIntCol库房ID, "库房id", 0, flexAlignCenterCenter, "库房id"
        VsfGridColFormat vsfList, mIntCol记录状态, "记录状态", 0, flexAlignCenterCenter, "记录状态"
        
        
        mstrUnallowSetColHide = "NO"
        mstrUnallowShow = "当前行;处方类型;标志;单据;收费;可操作;病人ID;未审核;实收金额;门诊标志;记录性质;收费类型;库房id"

        
        '恢复自定义列宽（不包括不允许显示的列）
        If str列设置 <> "" Then
            arr列设置 = Split(str列设置, "|")
            For n = 0 To UBound(arr列设置)
                If InStr(mstrUnallowShow, Split(arr列设置(n), ",")(0)) > 0 Then
                    For i = 0 To vsfList.Cols - 1
                        If Split(arr列设置(n), ",")(0) = vsfList.ColKey(i) Then
                            vsfList.ColWidth(i) = Val(Split(arr列设置(n), ",")(1))
                        End If
                    Next
                End If
            Next
        End If
        
        .RowSel = 1
        .Redraw = flexRDDirect
    End With
End Sub


Private Function LoadListColState() As String
    Dim strType As String
    
    If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then Exit Function
    
    Select Case mintType
        Case mListType.待发料
            strType = "待发料"
        Case mListType.退料
            strType = "退料"
    End Select
    
    LoadListColState = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & Me.Name & "\" & TypeName(vsfList), strType, "")
End Function


Private Sub SetColumnValue(ByVal str列名 As String, ByVal intValue As Integer)
    Select Case str列名
        Case "类别"
            mIntCol类型 = intValue
        Case "NO"
            mIntColNO = intValue
        Case "姓名"
            mIntCol姓名 = intValue
            
        Case "金额"
            mIntCol金额 = intValue
        Case "日期"
            mIntCol日期 = intValue

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
    End Select
                
End Sub


Public Sub RefreshList(ByVal rsTemp As Recordset, ByVal intType As Integer)
'功能：实现将数据表现到vsf表格
'参数1：rsTemp是刷新列表的数据集
'参数2：intType是代表当前的业务类型，1-待发药，2-已发药
    Dim intCount As Integer
    Dim strType As String
    Dim lngColor As Long
    Dim int可操作 As Integer
    Dim dblMoney As Double
    Dim str合计 As String
    
    On Error GoTo ErrHandle
    
    mintType = intType
    
    With Me.vsfList
        .Rows = 1
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        intCount = 1
        
        Do While Not rsTemp.EOF
            .TextMatrix(intCount, .ColIndex("当前行")) = intCount
            If rsTemp!单据 = 24 Then
                strType = "收费"
            Else
                strType = "记账"
            End If
            
            If rsTemp!已收费 = 0 Then strType = "（未）" & strType
            
            .TextMatrix(intCount, .ColIndex("类别")) = strType
            .TextMatrix(intCount, .ColIndex("标志")) = rsTemp!门诊标志
            .TextMatrix(intCount, .ColIndex("单据")) = rsTemp!单据
            .TextMatrix(intCount, .ColIndex("收费")) = rsTemp!已收费
            
            .TextMatrix(intCount, .ColIndex("NO")) = rsTemp!NO
            .TextMatrix(intCount, .ColIndex("姓名")) = rsTemp!姓名
            .TextMatrix(intCount, .ColIndex("金额")) = rsTemp!金额
            .TextMatrix(intCount, .ColIndex("日期")) = rsTemp!日期
            .TextMatrix(intCount, .ColIndex("就诊卡号")) = NVL(rsTemp!就诊卡号)
            .TextMatrix(intCount, .ColIndex("门诊号")) = NVL(rsTemp!门诊号)
            .TextMatrix(intCount, .ColIndex("身份证")) = NVL(rsTemp!身份证号)
            .TextMatrix(intCount, .ColIndex("IC卡")) = NVL(rsTemp!IC卡号)
            .TextMatrix(intCount, .ColIndex("病人id")) = NVL(rsTemp!病人id)
            .TextMatrix(intCount, .ColIndex("医保号")) = NVL(rsTemp!医保号)
            .TextMatrix(intCount, .ColIndex("住院号")) = NVL(rsTemp!住院号)
            .TextMatrix(intCount, .ColIndex("门诊标志")) = rsTemp!门诊标志
            .TextMatrix(intCount, .ColIndex("实收金额")) = Format(rsTemp!金额, mFMT.FM_金额)
            .TextMatrix(intCount, .ColIndex("记录性质")) = rsTemp!记录性质
            .TextMatrix(intCount, .ColIndex("库房id")) = rsTemp!库房ID
            .TextMatrix(intCount, .ColIndex("记录状态")) = rsTemp!记录状态
            .TextMatrix(intCount, .ColIndex("说明")) = rsTemp!说明
            
            dblMoney = dblMoney + Val(rsTemp!金额)
            
            '判断当前的退药次数
            If rsTemp!记录状态 = 1 Or rsTemp!记录状态 Mod 3 = 0 Then
                int可操作 = 1
            Else
                int可操作 = rsTemp!记录状态 Mod 3 + 1
            End If
            .TextMatrix(intCount, .ColIndex("可操作")) = int可操作
            
            
            '设置数据行颜色
            '设置颜色
            lngColor = IIf(intType = 0 Or int可操作 = 0, &H80000008, IIf(int可操作 = 1, glng正常, IIf(int可操作 = 2, glng发药, glng退药)))
            .Cell(flexcpForeColor, intCount, 1, intCount, .Cols - 1) = lngColor
              
            rsTemp.MoveNext
            intCount = intCount + 1
        Loop
        .Row = 1
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsfList_EnterCell()
'刷新明细信息
    Dim rsTemp As Recordset
    
    With Me.vsfList
        If .Row = 0 Then Exit Sub
        Call frmStuffRxSend.RefreshSendData(Val(.TextMatrix(.Row, .ColIndex("单据"))), .TextMatrix(.Row, .ColIndex("NO")), Val(.TextMatrix(.Row, .ColIndex("库房id"))), Val(.TextMatrix(.Row, .ColIndex("记录状态"))), Val(.TextMatrix(.Row, .ColIndex("可操作"))))
    End With
End Sub


Public Sub SetFontSize(ByVal intFont As Integer)
    Me.vsfList.FontSize = intFont
End Sub

Public Function FindSpecialRow(ByVal intFindType As Integer, ByVal strFindContent As String) As Boolean
    '如果是银行卡，则strFindContent格式为：卡ID|卡号
    Dim intCol As Integer
    Dim intFindRow As Integer
    
    With mFindProcess
        .FindType = intFindType
        .FindContent = UCase(strFindContent)
        .StartRow = 1
    End With
    
    With vsfList
        Select Case mFindProcess.FindType
            Case mFindType.姓名
                intCol = mIntCol姓名
                
                If zlCommFun.IsCharAlpha(mFindProcess.FindContent) Then
'                    '全字母时匹配简码
'                    If zlDatabase.GetPara("简码方式") = 0 Then
'                        intCol = mIntCol拼音简码
'                    Else
'                        intCol = mIntCol五笔简码
'                    End If
                End If
            Case mFindType.单据号
                intCol = mIntColNO
            Case mFindType.门诊号
                intCol = mIntCol门诊号
            Case mFindType.身份证
                intCol = mIntCol身份证
            Case mFindType.IC卡
                intCol = mIntCol病人ID
            Case mFindType.医保号
                intCol = mIntCol医保号
            Case mFindType.住院号
                intCol = mIntCol住院号
            Case Else
                '其余为消费卡，按病人ID查找
                intCol = mIntCol病人ID
                mFindProcess.FindContent = zlfuncCard_GetPatiID(Val(Split(strFindContent, "|")(0)), Split(strFindContent, "|")(1))
        End Select
        
        mFindProcess.StartRow = .FindRow(mFindProcess.FindContent, mFindProcess.StartRow, intCol)
        
        If mFindProcess.StartRow > 0 Then
            .Row = mFindProcess.StartRow
            .TopRow = .Row
            FindSpecialRow = True
            If mFindProcess.StartRow + 1 >= .Rows Then
                mFindProcess.StartRow = 1
            Else
                mFindProcess.StartRow = mFindProcess.StartRow + 1
            End If
        Else
            mFindProcess.StartRow = 1
        End If
        
    End With
End Function
