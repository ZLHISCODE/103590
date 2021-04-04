VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmPriceQuery 
   BackColor       =   &H8000000A&
   Caption         =   "收费项目与价目"
   ClientHeight    =   5160
   ClientLeft      =   165
   ClientTop       =   3750
   ClientWidth     =   8160
   Icon            =   "frmPriceQuery.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5160
   ScaleWidth      =   8160
   ShowInTaskbar   =   0   'False
   WindowState     =   2  'Maximized
   Begin VSFlex8Ctl.VSFlexGrid msh从属项目 
      Height          =   1530
      Left            =   2715
      TabIndex        =   7
      ToolTipText     =   "从属项目"
      Top             =   2850
      Width           =   5400
      _cx             =   9525
      _cy             =   2699
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
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483634
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
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
      FormatString    =   $"frmPriceQuery.frx":030A
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
      Editable        =   1
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
   Begin VB.PictureBox picHBar_S 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   825
      MousePointer    =   7  'Size N S
      ScaleHeight     =   60
      ScaleWidth      =   6075
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   2250
      Width           =   6075
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshSum 
      Height          =   945
      Left            =   3075
      TabIndex        =   5
      Top             =   960
      Width           =   3195
      _ExtentX        =   5636
      _ExtentY        =   1667
      _Version        =   393216
      FixedCols       =   0
      BackColorBkg    =   -2147483643
      GridColor       =   8421504
      GridColorFixed  =   8421504
      GridColorUnpopulated=   8421504
      AllowBigSelection=   0   'False
      FocusRect       =   0
      GridLinesFixed  =   1
      SelectionMode   =   1
      AllowUserResizing=   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox picV 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3225
      Left            =   2055
      MousePointer    =   9  'Size W E
      ScaleHeight     =   3225
      ScaleWidth      =   45
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   900
      Width           =   45
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   2490
      Top             =   1020
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceQuery.frx":03DB
            Key             =   "R"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceQuery.frx":06F5
            Key             =   "C"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceQuery.frx":084F
            Key             =   "P"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwMain_S 
      Height          =   3525
      Left            =   75
      TabIndex        =   3
      Top             =   960
      Width           =   2325
      _ExtentX        =   4101
      _ExtentY        =   6218
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList ilsColor 
      Left            =   5640
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceQuery.frx":0CA1
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceQuery.frx":0EBD
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceQuery.frx":10D9
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceQuery.frx":12F3
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceQuery.frx":150F
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMono 
      Left            =   4920
      Top             =   60
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceQuery.frx":172B
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceQuery.frx":1947
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceQuery.frx":1B63
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceQuery.frx":1D7D
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmPriceQuery.frx":1F99
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   8160
      _CBHeight       =   780
      _Version        =   "6.7.8988"
      Child1          =   "tbrThis"
      MinHeight1      =   720
      Width1          =   5370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   720
         Left            =   30
         TabIndex        =   1
         Top             =   30
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsMono"
         HotImageList    =   "ilsColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "Preview"
               Object.ToolTipText     =   "预览"
               Object.Tag             =   "预览"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "Print"
               Description     =   "打印"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查找"
               Key             =   "Find"
               Object.ToolTipText     =   "查找"
               Object.Tag             =   "查找"
               ImageKey        =   "Find"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Quit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   2
      Top             =   4800
      Width           =   8160
      _ExtentX        =   14393
      _ExtentY        =   635
      SimpleText      =   $"frmPriceQuery.frx":21B5
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmPriceQuery.frx":21FC
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   9313
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFilePreview 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuFileLine1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "标准按钮(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolSplit 
            Caption         =   "-"
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "文本标签(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewSplit1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewShowDynamic 
         Caption         =   "显示变价项目(&D)"
      End
      Begin VB.Menu mnuViewSplit3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "查找(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewSplit2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "刷新(&R) "
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpTitle 
         Caption         =   "帮助主题(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "Web上的中联"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "中联主页(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "中联论坛(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "发送反馈(&K)..."
         End
      End
      Begin VB.Menu mnuHelpWebL 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)…"
      End
   End
End
Attribute VB_Name = "frmPriceQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum Column价目
    col编码 = 0
    col名称 = 1
    col规格 = 2
    col产地 = 3
    col售价单位 = 4
    col收入项目 = 5
    col价格 = 6
    col附术 = 7
    col加班 = 8
End Enum

Dim strPrivs As String   '模块权限
Dim mblnTradeName As Boolean        '是否以商品名显示成药
Dim rsTemp As New ADODB.Recordset

Dim mblnLoad As Boolean
Dim msngStartX As Single    '移动前鼠标的位置
Dim mstrKey As String       '前一个树节点的关键值
Dim mrs价目 As New ADODB.Recordset
Dim msngOldY As Single
Private Type 售价精度
    收费项目小数 As Integer
    药品项目小数 As Integer
    卫材项目小数 As Integer
End Type
Private m售价精度 As 售价精度

Private mstrVbFormat As String
Private mstrOraFormat As String
Private mlngPreRow As Long

Private Sub Form_Activate()
    If Me.tvwMain_S.Nodes.Count = 0 Then
        MsgBox "没有建立项目，或者权限不具备", vbExclamation, gstrSysName
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub Form_Load()
    strPrivs = gstrPrivs
    mblnLoad = True
    RestoreWinState Me, App.ProductName
    mnuViewShowDynamic.Checked = (GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name & "\Menu", "mnuViewShowDynamic状态", "False") = "True")
    
    Call init小数位数
    
    mstrVbFormat = GetFmtString(1, False)
    mstrOraFormat = GetFmtString(1, True)
    
    mrs价目.CursorLocation = adUseClient
    '得到查询的时间范围
    Call InitSum
    
    mblnTradeName = False
'    gstrSQL = "Select nvl(参数值,'0') From 系统参数表 Where 参数名='西成药以商品名显示'"
    mblnTradeName = IIf(zlDatabase.GetPara(74, 100, , -1) = 1, True, False)
    
    '装入分类
    With tvwMain_S.Nodes
        .Clear
        If InStr(1, strPrivs, "收费项目") <> 0 Then
            .Add , , "K0", "[0]收费项目", "R", "R"
            tvwMain_S.Nodes("K0").Sorted = True
            tvwMain_S.Nodes("K0").Tag = "0"
            Call FillTree(0)
        End If
        If InStr(1, strPrivs, "西成药") <> 0 Then
            .Add , , "K1", "[1]西成药", "R", "R"
            tvwMain_S.Nodes("K1").Sorted = True
            tvwMain_S.Nodes("K1").Tag = 1
            Call FillTree(1)
        End If
        If InStr(1, strPrivs, "中成药") <> 0 Then
            .Add , , "K2", "[2]中成药", "R", "R"
            tvwMain_S.Nodes("K2").Sorted = True
            tvwMain_S.Nodes("K2").Tag = 2
            Call FillTree(2)
        End If
        If InStr(1, strPrivs, "中草药") <> 0 Then
            .Add , , "K3", "[3]中草药", "R", "R"
            tvwMain_S.Nodes("K3").Sorted = True
            tvwMain_S.Nodes("K3").Tag = 3
            
            Call FillTree(3)
        End If
        If InStr(1, strPrivs, "卫生材料") <> 0 Then
            .Add , , "K7", "[4]卫生材料", "R", "R"
            tvwMain_S.Nodes("K7").Sorted = True
            tvwMain_S.Nodes("K7").Tag = 7
            Call FillTree(7)
        End If
    End With
    If Me.tvwMain_S.Nodes.Count <> 0 Then
        Me.tvwMain_S.Nodes(1).Expanded = True
        Me.tvwMain_S.Nodes(1).Selected = True
        Call Me.tvwMain_S.Nodes(1).EnsureVisible
    End If
    
End Sub

Private Sub InitSum()
    '初始化汇总表的样式
    With mshSum
        ClearGrid mshSum, 9

'        .MergeCells = flexMergeRestrictRows
'        .MergeCol(col编码) = True
        .TextMatrix(0, col编码) = "编码"
        .TextMatrix(0, col名称) = "收费细目"
        .TextMatrix(0, col规格) = "规格"
        .TextMatrix(0, col产地) = "产地"
        .TextMatrix(0, col售价单位) = "单位"
        .TextMatrix(0, col收入项目) = "收入项目"
        .TextMatrix(0, col价格) = "价格"
        .TextMatrix(0, col附术) = "附术收费率"
        .TextMatrix(0, col加班) = "加班加价率"
        
        .ColWidth(col编码) = 1000
        .ColWidth(col名称) = 2500
        .ColWidth(col规格) = 1600
        .ColWidth(col产地) = 1500
        .ColWidth(col售价单位) = 600
        .ColWidth(col收入项目) = 900
        .ColWidth(col价格) = 1100
        .ColWidth(col附术) = 800
        .ColWidth(col加班) = 800
        
        .ColAlignment(col编码) = 1
        .ColAlignment(col名称) = 1
        .ColAlignment(col规格) = 1
        .ColAlignment(col产地) = 1
        .ColAlignment(col售价单位) = 1
        .ColAlignment(col收入项目) = 1
        .ColAlignment(col价格) = 7
        .ColAlignment(col附术) = 7
        .ColAlignment(col加班) = 7
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mrs价目.Sort = ""
    If mrs价目.State = 1 Then mrs价目.Close
    
    mstrKey = ""
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name & "\Menu", "mnuViewShowDynamic状态", mnuViewShowDynamic.Checked
    SaveWinState Me, App.ProductName
End Sub

Private Sub Form_Resize()
    Dim sngTop As Single, sngBottom As Single
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    sngTop = IIf(cbrThis.Visible, cbrThis.Top + cbrThis.Height, 0)
    sngBottom = ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0)
    '右边
    'tvwMain_S的位置
    tvwMain_S.Top = sngTop
    tvwMain_S.Height = IIf(sngBottom - tvwMain_S.Top > 0, sngBottom - tvwMain_S.Top, 0)
    tvwMain_S.Left = 0
    'picV的位置
    picV.Top = sngTop
    picV.Height = tvwMain_S.Height
    picV.Left = tvwMain_S.Left + tvwMain_S.Width
        
        
    mshSum.Left = picV.Left + picV.Width
    mshSum.Width = ScaleWidth - mshSum.Left
    mshSum.Top = sngTop
    
    If picHBar_S.Top > Me.ScaleHeight - 2000 Then picHBar_S.Top = Me.ScaleHeight - 2000
    picHBar_S.Left = mshSum.Left
    picHBar_S.Width = mshSum.Width
    If msh从属项目.Visible = False Then
        mshSum.Height = IIf(sngBottom - mshSum.Top > 0, sngBottom - mshSum.Top, 0)
    Else
        mshSum.Height = picHBar_S.Top - mshSum.Top '  IIf(sngBottom - mshSum.Top > 0, sngBottom - mshSum.Top, 0)
    End If
    With msh从属项目
        If .Visible = True Then
            .Left = mshSum.Left
            .Top = picHBar_S.Top + picHBar_S.Height
            .Height = IIf(sngBottom - .Top > 0, sngBottom - .Top, 0)
            .Width = Me.ScaleWidth - .Left
        End If
    End With
    Refresh
End Sub

Private Sub mnuViewFind_Click()
    With frmPriceFind
        .Left = Me.Left + Me.Width - .Width
        .Top = Me.Top + Me.Height - .Height
        .Show vbModal, Me
    End With
End Sub

Private Sub mnuViewRefresh_Click()
    FillSum
End Sub

Private Sub mnuViewShowDynamic_Click()
    mnuViewShowDynamic.Checked = Not mnuViewShowDynamic.Checked
    
    mstrKey = ""
    Call FillSum
End Sub

 

Private Sub mshSum_GotFocus()
    Call MenuSet
    mshSum.BackColorSel = &H8000000D
End Sub

Private Sub mshSum_LostFocus()
    Call MenuSet
    mshSum.BackColorSel = &H8000000F
End Sub

Private Sub mshSum_RowColChange()
    With mshSum
        If .Row <> mlngPreRow Then
            mlngPreRow = .Row
            
            Load重属项目 .RowData(.Row)
        End If
    End With
End Sub

Private Sub msh从属项目_GotFocus()
    msh从属项目.BackColorSel = &H8000000D
End Sub

Private Sub msh从属项目_LostFocus()
    msh从属项目.BackColorSel = &H8000000F
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnuViewTool, 2
    End If
End Sub

Public Sub tvwMain_S_NodeClick(ByVal Node As MSComctlLib.Node)
    Dim intType As Integer
    mlngPreRow = -1
    Select Case Val(Node.Tag)
    Case 1, 2, 3        '西成药,中成药,中草药
        intType = 2     'intType-1- 收费项目,2-药品项目,3-卫材项目
    Case 7              '卫生材料
        intType = 3     'intType-1- 收费项目,2-药品项目,3-卫材项目
    Case Else           '收费项目
        intType = 1
    End Select
    mstrVbFormat = GetFmtString(intType, False)
    mstrOraFormat = GetFmtString(intType, True)
    FillSum
    Call Set从属项目
    Call mshSum_RowColChange
End Sub

Private Sub picV_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        msngStartX = x
    End If
End Sub

Private Sub picV_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim sngTemp As Single
    If Button = 1 Then
        sngTemp = picV.Left + x - msngStartX
        If sngTemp > 1500 And ScaleWidth - (sngTemp + picV.Width) > 1600 Then
            picV.Left = sngTemp
            tvwMain_S.Width = picV.Left - tvwMain_S.Left
            Form_Resize
        End If
    End If
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFilePrintSet_Click()
    zlPrintSet
End Sub

Private Sub mnuFileExcel_Click()
    subPrint 3
End Sub

Private Sub mnuFilePreview_Click()
    subPrint 2
End Sub

Private Sub mnuFilePrint_Click()
    subPrint 1
End Sub

Private Sub tabMain_Click()
    mstrKey = ""
    Call FillSum
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Select Case Button.Key
        Case "Find"
            mnuViewFind_Click
        Case "Quit"
            mnuFileExit_Click
        Case "Print"
            mnuFilePrint_Click
        Case "Preview"
            mnuFilePreview_Click
        Case "Help"
            mnuHelpTitle_Click
    End Select
    
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
    cbrThis.Bands("only").MinHeight = tbrThis.Height
    Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim buttTemp As Button
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For Each buttTemp In tbrThis.Buttons
        If mnuViewToolText.Checked Then
            buttTemp.Caption = buttTemp.Tag
        Else
            buttTemp.Caption = ""
        End If
    Next
    cbrThis.Bands("only").MinHeight = tbrThis.Height
    Form_Resize
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Form_Resize
End Sub

Private Sub mnuHelpAbout_Click()
    ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
   Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(hWnd)
End Sub

Private Sub subPrint(bytMode As Byte)
'功能:进行打印,预览和输出到EXCEL
'参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim nod As Node
    
    Set nod = tvwMain_S.SelectedItem
    Do Until nod.Parent Is Nothing
        Set nod = nod.Parent
    Loop
    
    Set objPrint.Body = mshSum
    objPrint.Title.Text = nod.Text & "类项目价目表"
    objRow.Add "医院名称：" & gstr单位名称
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "打印人：" & gstrUserName
    objRow.Add "打印时间：" & Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    objPrint.BelowAppRows.Add objRow
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Function FillTree(lngKind As Long) As Boolean
    '功能:装入收费类别和收费细目的所有分类到tvwMain_S
    '本程序中树节点比其它程序的KEY值多一个字符，即第二位的类别编码
    Dim objNode As Node
    
    Select Case lngKind
    Case 0
        gstrSQL = "Select id, 上级id, 编码, 名称 " & _
                "  From 收费分类目录" & _
                " Start With 上级ID Is Null" & _
                " Connect By Prior id = 上级ID"
    Case 1, 2, 3, 7
        gstrSQL = "Select id, 上级id, 编码, 名称 " & _
                "  From 诊疗分类目录" & _
                " Where 类型 = " & lngKind & _
                " Start With 上级ID Is Null" & _
                " Connect By Prior id = 上级ID"
    End Select
    Call OpenRecordset(rsTemp, Me.Caption)
    With tvwMain_S.Nodes
        Do Until rsTemp.EOF
            If IsNull(rsTemp("上级id")) Then
                Set objNode = .Add("K" & lngKind, tvwChild, "K" & lngKind & rsTemp("id"), "【" & rsTemp("编码") & "】" & rsTemp("名称"), "P", "P")
            Else
                Set objNode = .Add("K" & lngKind & rsTemp("上级id"), tvwChild, "K" & lngKind & rsTemp("id"), "【" & rsTemp("编码") & "】" & rsTemp("名称"), "P", "P")
            End If
            objNode.Tag = lngKind
            objNode.Sorted = True
            rsTemp.MoveNext
        Loop
    End With
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub Set从属项目()
    '--------------------------------------------------------------------------------------------------
    '功能:设置从属项目的相关属性
    '编制:刘兴宏
    '日期:2007/09/27
    '--------------------------------------------------------------------------------------------------
    Dim blnVisible As Boolean
    Dim blnOldHide As Boolean       '上次是否影藏的
    
    If Me.tvwMain_S.SelectedItem Is Nothing Then
        blnVisible = False
    Else
        blnVisible = Val(Me.tvwMain_S.SelectedItem.Tag) = 0
    End If
    blnOldHide = msh从属项目.Visible
    
    With msh从属项目
        .Visible = blnVisible
        picHBar_S.Visible = blnVisible
    End With
    If blnOldHide = False And blnVisible Then
        If picHBar_S.Top < Me.ScaleHeight - picHBar_S.Top - IIf(stbThis.Visible = False, 0, stbThis.Height) Then
            picHBar_S.Top = Me.ScaleHeight - IIf(stbThis.Visible = False, 0, stbThis.Height) - 2000
        End If
    End If
    If picHBar_S.Top > Me.ScaleHeight - IIf(stbThis.Visible = False, 0, stbThis.Height) - 2000 Then
         picHBar_S.Top = Me.ScaleHeight - IIf(stbThis.Visible = False, 0, stbThis.Height) - 2000
    End If
    Call Form_Resize
End Sub
Public Sub FillSum()
    '功能:装入各种统计数据
    Dim nod As Node
    Dim str材质分类 As String

    If tvwMain_S.SelectedItem Is Nothing Then
        ClearGrid mshSum
        Call MenuSet
        Exit Sub
    End If
    If mstrKey = tvwMain_S.SelectedItem.Key Then Exit Sub
    mstrKey = tvwMain_S.SelectedItem.Key
    Set nod = tvwMain_S.SelectedItem
    
    '根据不同的节点，做出不同的显示
    Select Case Mid(nod.Key, 2, 1)
    Case "0"
        mshSum.TextMatrix(0, col产地) = "产地"
        mshSum.ColWidth(col产地) = 0
        mshSum.ColWidth(col附术) = 1100
        mshSum.ColWidth(col加班) = 1000
        mshSum.MergeCol(col编码) = True
        mshSum.MergeCol(col名称) = True
        
        If nod.Image = "R" Then
            gstrSQL = "Select id,编码,名称,加班加价,是否变价,规格,产地,计算单位" & _
                    " From 收费项目目录" & _
                    " Where 类别 not in ('4','5','6','7') And " & IIf(mnuViewShowDynamic.Checked, "", " 是否变价=0 And ") & _
                    "       (撤档时间 is null or 撤档时间=to_date('3000-01-01','yyyy-mm-dd'))"
        Else
            gstrSQL = "Select id,编码,名称,加班加价,是否变价,规格,产地,计算单位" & _
                    " From 收费项目目录" & _
                    " Where 类别 not in ('4','5','6','7') And " & IIf(mnuViewShowDynamic.Checked, "", " 是否变价=0 And ") & _
                    "       (撤档时间 is null or 撤档时间=to_date('3000-01-01','yyyy-mm-dd')) And" & _
                    "       分类ID in (" & _
                    "           Select Id From 收费分类目录 start with ID=" & Mid(nod.Key, 3) & " connect by prior id=上级ID)"
        End If
        gstrSQL = "select A.收费细目ID,C.编码,C.名称 as 收费细目,C.规格,C.产地,C.计算单位 as 售价单位,B.名称 as 收入项目,A.原价,A.现价,A.附术收费率,decode(C.加班加价,1,A.加班加价率,0) as 加班加价率,C.是否变价 " & _
                   " from 收费价目 A,收入项目 B, " & _
                   "   (" & gstrSQL & ") C " & _
                   " Where A.收费细目ID = C.ID And A.收入项目ID = B.ID " & _
                   "       and A.执行日期<=sysdate and (A.终止日期>=sysdate or a.终止日期 is null) " & _
                   " order by C.编码"
    
    Case "1", "2", "3", "7"
        Select Case Mid(nod.Key, 2, 1)
        Case "1", "2"
            mshSum.TextMatrix(0, col产地) = "厂牌"
            mshSum.ColWidth(col产地) = 1500
            mshSum.ColWidth(col规格) = 1600
        Case "3"
            mshSum.TextMatrix(0, col产地) = "产地"
            mshSum.ColWidth(col产地) = 1000
            mshSum.ColWidth(col规格) = 0
        Case "7"
            mshSum.TextMatrix(0, col产地) = "产地"
            mshSum.ColWidth(col产地) = 1000
            mshSum.ColWidth(col规格) = 1600
        End Select
        mshSum.ColWidth(col加班) = 0
        mshSum.ColWidth(col附术) = 0
        mshSum.MergeCol(col编码) = False
        mshSum.MergeCol(col名称) = False
        
        If nod.Image = "R" Then
            If mblnTradeName = False Then
                gstrSQL = "Select id,编码,名称,加班加价,是否变价,规格,产地,计算单位" & _
                        " From 收费项目目录" & _
                        " Where 类别 ='" & Switch(nod.Key = "K1", 5, nod.Key = "K2", 6, nod.Key = "K3", 7, nod.Key = "K7", 4) & "'" & _
                        "       And " & IIf(mnuViewShowDynamic.Checked, "", " 是否变价=0 And ") & _
                        "       (撤档时间 is null or 撤档时间=to_date('3000-01-01','yyyy-mm-dd'))"
            Else
                gstrSQL = "Select id, 编码, nvl(N.名称,I.名称) As 名称, 加班加价, 是否变价, 规格, 产地, 计算单位" & _
                        " From 收费项目目录 I,收费项目别名 N" & _
                        " Where 类别 ='" & Switch(nod.Key = "K1", 5, nod.Key = "K2", 6, nod.Key = "K3", 7, nod.Key = "K7", 4) & "'" & _
                        "       And " & IIf(mnuViewShowDynamic.Checked, "", " 是否变价=0 And ") & _
                        "       (撤档时间 is null or 撤档时间=to_date('3000-01-01','yyyy-mm-dd'))" & _
                        "       And I.Id=N.收费细目id(+) And N.性质(+)=3 And N.码类(+)=1"
            End If
        Else
            If mblnTradeName = False Then
                If Mid(nod.Key, 2, 1) = 7 Then
                    gstrSQL = "Select I.id, I.编码, I.名称, I.加班加价, I.是否变价, I.规格, I.产地, I.计算单位" & _
                            "  From 收费项目目录 I,材料特性 T,诊疗项目目录 Z" & _
                            " Where I.类别 ='" & Switch(Mid(nod.Key, 1, 2) = "K1", 5, Mid(nod.Key, 1, 2) = "K2", 6, Mid(nod.Key, 1, 2) = "K3", 7, Mid(nod.Key, 1, 2) = "K7", 4) & "'" & _
                            "       And " & IIf(mnuViewShowDynamic.Checked, "", " I.是否变价=0 And ") & _
                            "       (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','yyyy-mm-dd'))" & _
                            "       And I.Id=T.材料id And T.诊疗id=Z.Id" & _
                            "       And Z.分类ID in (" & _
                            "           Select Id From 诊疗分类目录 start with ID=" & Mid(nod.Key, 3) & " connect by prior id=上级ID)"
                Else
                    gstrSQL = "Select I.id, I.编码, I.名称, I.加班加价, I.是否变价, I.规格, I.产地, I.计算单位" & _
                            "  From 收费项目目录 I,药品规格 T,诊疗项目目录 Z" & _
                            " Where I.类别 ='" & Switch(Mid(nod.Key, 1, 2) = "K1", 5, Mid(nod.Key, 1, 2) = "K2", 6, Mid(nod.Key, 1, 2) = "K3", 7, Mid(nod.Key, 1, 2) = "K7", 4) & "'" & _
                            "       And " & IIf(mnuViewShowDynamic.Checked, "", " I.是否变价=0 And ") & _
                            "       (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','yyyy-mm-dd'))" & _
                            "       And I.Id=T.药品id And T.药名id=Z.Id" & _
                            "       And Z.分类ID in (" & _
                            "           Select Id From 诊疗分类目录 start with ID=" & Mid(nod.Key, 3) & " connect by prior id=上级ID)"
                End If
            Else
                If Mid(nod.Key, 2, 1) = 7 Then
                    gstrSQL = "Select I.id, I.编码, nvl(N.名称,I.名称) As 名称, I.加班加价, I.是否变价, I.规格, I.产地, I.计算单位" & _
                            "  From 收费项目目录 I,收费项目别名 N,材料特性 T,诊疗项目目录 Z" & _
                            " Where I.类别 ='" & Switch(Mid(nod.Key, 1, 2) = "K1", 5, Mid(nod.Key, 1, 2) = "K2", 6, Mid(nod.Key, 1, 2) = "K3", 7, Mid(nod.Key, 1, 2) = "K7", 4) & "'" & _
                            "       And " & IIf(mnuViewShowDynamic.Checked, "", " I.是否变价=0 And ") & _
                            "       (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','yyyy-mm-dd'))" & _
                            "       And I.Id=N.收费细目id(+) And N.性质(+)=3 And N.码类(+)=1" & _
                            "       And I.Id=T.材料id And T.诊疗id=Z.Id" & _
                            "       And Z.分类ID in (" & _
                            "           Select Id From 诊疗分类目录 start with ID=" & Mid(nod.Key, 3) & " connect by prior id=上级ID)"
                Else
                    gstrSQL = "Select I.id, I.编码, nvl(N.名称,I.名称) As 名称, I.加班加价, I.是否变价, I.规格, I.产地, I.计算单位" & _
                            "  From 收费项目目录 I,收费项目别名 N,药品规格 T,诊疗项目目录 Z" & _
                            " Where I.类别 ='" & Switch(Mid(nod.Key, 1, 2) = "K1", 5, Mid(nod.Key, 1, 2) = "K2", 6, Mid(nod.Key, 1, 2) = "K3", 7, Mid(nod.Key, 1, 2) = "K7", 4) & "'" & _
                            "       And " & IIf(mnuViewShowDynamic.Checked, "", " I.是否变价=0 And ") & _
                            "       (I.撤档时间 is null or I.撤档时间=to_date('3000-01-01','yyyy-mm-dd'))" & _
                            "       And I.Id=N.收费细目id(+) And N.性质(+)=3 And N.码类(+)=1" & _
                            "       And I.Id=T.药品id And T.药名id=Z.Id" & _
                            "       And Z.分类ID in (" & _
                            "           Select Id From 诊疗分类目录 start with ID=" & Mid(nod.Key, 3) & " connect by prior id=上级ID)"
                End If
            End If
        End If
        gstrSQL = "select A.收费细目ID,C.编码,C.名称 as 收费细目,C.规格,C.产地,C.计算单位 as 售价单位,B.名称 as 收入项目,A.原价,A.现价,A.附术收费率,decode(C.加班加价,1,A.加班加价率,0) as 加班加价率,C.是否变价 " & _
                   " from 收费价目 A,收入项目 B, " & _
                   "   (" & gstrSQL & ") C " & _
                   " Where A.收费细目ID = C.ID And A.收入项目ID = B.ID " & _
                   "       and A.执行日期<=sysdate and (A.终止日期>=sysdate or a.终止日期 is null) " & _
                   " order by C.编码"
    End Select
    
    MousePointer = 11
    If mrs价目.State = 1 Then mrs价目.Close
    Call OpenRecordset(mrs价目, Me.Caption)
    
    Call ReList
    MousePointer = 0
    Call MenuSet
End Sub

Private Sub ReList()
    Dim lngRow As Long
    Dim lngID  As Long
    Dim lngCount As Long
    
    
    MousePointer = 11
    mshSum.Redraw = False
    ClearGrid mshSum
    If mrs价目.RecordCount <> 0 Then
        mshSum.Rows = mrs价目.RecordCount + 1
    End If
    lngRow = 1
    With mshSum
        Do Until mrs价目.EOF
            If mrs价目("收费细目ID") <> lngID Then
                lngID = mrs价目("收费细目ID")
                lngCount = lngCount + 1
            End If
            .RowData(lngRow) = lngID
            .TextMatrix(lngRow, col编码) = mrs价目("编码")
            .TextMatrix(lngRow, col名称) = mrs价目("收费细目")
            .TextMatrix(lngRow, col规格) = IIf(IsNull(mrs价目("规格")), "", mrs价目("规格"))
            .TextMatrix(lngRow, col产地) = IIf(IsNull(mrs价目("产地")), "", mrs价目("产地"))
            .TextMatrix(lngRow, col售价单位) = IIf(IsNull(mrs价目("售价单位")), "", mrs价目("售价单位"))
            .TextMatrix(lngRow, col收入项目) = mrs价目("收入项目")
            If mrs价目("是否变价") = 1 Then
                .TextMatrix(lngRow, col价格) = Format(mrs价目("原价"), mstrVbFormat) & "～" & Format(mrs价目("现价"), mstrVbFormat)
            Else
                .TextMatrix(lngRow, col价格) = Format(mrs价目("现价"), mstrVbFormat)
            End If
            .TextMatrix(lngRow, col附术) = Format(mrs价目("附术收费率"), "0.00;-0.00; ; ")
            .TextMatrix(lngRow, col加班) = Format(mrs价目("加班加价率"), "0.00;-0.00; ; ")
            lngRow = lngRow + 1
            mrs价目.MoveNext
        Loop
    End With
    mshSum.Redraw = True
    stbThis.Panels(2).Text = "共有收费项目" & lngCount & "条"
    MousePointer = 0

End Sub

Private Sub ClearGrid(objGrid As MSHFlexGrid, Optional lngCols As Long = 0)
'功能：清除表格,并完成部分初始化
    Dim i As Long
    
    With objGrid
        If lngCols > 0 Then
            '如果有列数传进来，那就初始化它
            .Cols = lngCols
            .AllowBigSelection = True
            .FillStyle = flexFillRepeat
            .Col = 0
            .Row = 0
            .ColSel = .Cols - 1
            .RowSel = 0
            .CellAlignment = 4
            .FillStyle = flexFillSingle
            .AllowBigSelection = False
            .Row = 1
        End If
        
        .Rows = 2
        .RowData(1) = 0
        For i = 0 To objGrid.Cols - 1
            objGrid.TextMatrix(1, i) = ""
        Next
    
    End With
End Sub

Private Sub MenuSet()
'功能:显示菜单和工具栏的状态(打印)
    Dim blnPrint As Boolean
    
    blnPrint = Not (mshSum.Rows = 2 And mshSum.TextMatrix(1, col编码) = "")

    mnuFilePreview.Enabled = blnPrint
    mnuFilePrint.Enabled = blnPrint
    mnuFileExcel.Enabled = blnPrint
    tbrThis.Buttons("Preview").Enabled = blnPrint
    tbrThis.Buttons("Print").Enabled = blnPrint
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub






Private Sub picHBar_S_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button <> 1 Then Exit Sub
    msngOldY = y
End Sub

Private Sub picHBar_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    '分割条设置
    
    If Button <> 1 Then Exit Sub
    
    With picHBar_S
        If .Top + y < 2000 Then Exit Sub
        If .Top + y > ScaleHeight - 2000 Then Exit Sub
        .Move .Left, .Top + y - msngOldY
    End With
    
    Call Form_Resize
    
    
End Sub

Private Sub picHBar_S_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
        msngOldY = 0
End Sub
Private Sub Load重属项目(ByVal lng收费细目ID As Long)
    '---------------------------------------------------------------------------------------
    '功能:加载重属项目
    '参数:
    '编制:刘兴宏
    '日期:2007/09/28
    '---------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim j As Integer, i As Integer, intCol As Integer, lngRow As Long
    
    If msh从属项目.Visible = False Then Exit Sub
    
    
    gstrSQL = "" & _
        "   Select a.主项ID,a.从项ID,a.固有从属,a.从项数次,b.名称,b.编码 项目编码,c.编码 ,c.名称 类别, " & vbCrLf & _
        "           Nvl(B.撤档时间,to_Date('3000-01-01','YYYY-MM-DD')) As 撤档时间," & vbCrLf & _
        "           decode(nvl(b.是否变价,0),1,ltrim(rtrim(to_char(sum(d.原价),'" & mstrOraFormat & "')))||'～'||ltrim(rtrim(to_char(sum(d.现价),'" & mstrOraFormat & "'))),ltrim(rtrim(to_char(sum(d.现价),'" & mstrOraFormat & "'))))  AS  价格 " & vbCrLf & _
        "   From 收费从属项目 a,收费项目目录 b,收费项目类别 c ,收费价目 d " & vbCrLf & _
        "   Where c.编码=b.类别 and  a.从项ID=b.id  and b.id=d.收费细目id  and 主项ID=[1] " & vbCrLf & _
        "           AND NVL (D.终止日期, TO_DATE ('3000-01-01', 'YYYY-MM-DD')) = TO_DATE ('3000-01-01', 'YYYY-MM-DD') " & _
        "   GROUP BY a.ROWID,a.主项ID,b.是否变价,a.从项ID,a.固有从属,a.从项数次,b.名称,b.编码,b.撤档时间 ,c.编码 ,c.名称 " & _
        " ORDER BY a.ROWID "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng收费细目ID)
    With msh从属项目
        .Redraw = flexRDNone
        If rsTemp.RecordCount = 0 Then
            .Rows = 2
            For i = 0 To .Cols - 1
                .TextMatrix(1, i) = ""
            Next
            .Redraw = flexRDBuffered
            Exit Sub
        End If
        .Rows = rsTemp.RecordCount + 1
        Dim dbl单价 As Double, intTemp As Integer
        
        i = 1
        dbl单价 = 0
        Do While Not rsTemp.EOF
            .TextMatrix(i, .ColIndex("收费类别")) = "(" & NVL(rsTemp!编码) & ")" & NVL(rsTemp!类别)
            .TextMatrix(i, .ColIndex("收费项目")) = "[" & NVL(rsTemp!项目编码) & "]" & NVL(rsTemp!名称)
            .TextMatrix(i, .ColIndex("次数")) = NVL(rsTemp!从项数次)
            intTemp = Val(NVL(rsTemp!固有从属))
            If intTemp = 0 Then
                .TextMatrix(i, .ColIndex("固定")) = "0-不固定"
            ElseIf intTemp = 2 Then
                .TextMatrix(i, .ColIndex("固定")) = "2-按比例计算"
            Else
                .TextMatrix(i, .ColIndex("固定")) = "1-固定"
            End If
            
            If Format(rsTemp!撤档时间, "YYYY-MM-DD") <> "3000-01-01" Then
                lngRow = .Row: intCol = .Col
                .Row = i
                For j = 0 To .Cols - 1
                    .Col = j
                    .CellForeColor = &HFF&
                Next
                .Row = lngRow: .Col = intCol
                .TextMatrix(i, .ColIndex("状态")) = "停用"
            Else
                .TextMatrix(i, .ColIndex("状态")) = ""
            End If
            .TextMatrix(i, .ColIndex("单价")) = NVL(rsTemp("价格"))
            If IsNumeric(.TextMatrix(i, .ColIndex("单价"))) Then
                dbl单价 = dbl单价 + Val(.TextMatrix(i, .ColIndex("单价")))
            End If
            i = i + 1
            rsTemp.MoveNext
        Loop
        '加入合计行
        .Rows = .Rows + 1
        .TextMatrix(i, .ColIndex("收费类别")) = ""
        .TextMatrix(i, .ColIndex("收费项目")) = "合计"
        .TextMatrix(i, .ColIndex("次数")) = ""
        .TextMatrix(i, .ColIndex("固定")) = ""
        .TextMatrix(i, .ColIndex("状态")) = ""
        .TextMatrix(i, .ColIndex("单价")) = Format(dbl单价, mstrVbFormat)
        .Redraw = flexRDBuffered
    End With
End Sub
Private Function NVL(rsObj As Field, Optional ByVal varValue As Variant = "") As Variant
    '-----------------------------------------------------------------------------------
    '功能:取某字段的值
    '参数:rsObj          被检查的字段
    '     varValue       当rsObj为NULL值时的取新值
    '返回:如果不为空值,返回原来的值,如果为空值,则返回指定的varValue值
    '-----------------------------------------------------------------------------------
    If IsNull(rsObj) Then
        NVL = varValue
    Else
        NVL = rsObj
    End If
End Function

Private Sub init小数位数()
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    '    类别    Number(1)   1-药品,2-卫材
    '    内容    Number(1)   1-成本价，2-零售价,3-数量
    '    单位    Number(1)   1,2,3,4；药品分别为售价、门诊、住院、药库单位，卫材分别为散装、包装单位。
    '    精度    Number(1)   取值为2-7。
    
    
    m售价精度.药品项目小数 = 7
    m售价精度.卫材项目小数 = 7
    m售价精度.收费项目小数 = 3
    strSQL = "Select * from 药品卫材精度 where 类别 in (1,2) and 内容=2 and  单位=1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取卫生材料的小数位数精度", 2)
    Do While Not rsTemp.EOF
        Select Case Val(NVL(rsTemp!类别))
        Case 1
            m售价精度.药品项目小数 = Val(NVL(rsTemp!精度, "7"))
        Case 2
            m售价精度.卫材项目小数 = Val(NVL(rsTemp!精度, "7"))
        End Select
        rsTemp.MoveNext
    Loop
End Sub
Private Function GetFmtString(ByVal intType As Integer, Optional blnOracle As Boolean = False) As String
    '------------------------------------------------------------------------------------------------------
    '功能:返回指定的小数格式串
    '入参:intType-1- 收费项目,2-药品项目,3-卫材项目
    '     blnOracle-返回是oracle的格式串还是Vb的格式串
    '出参:
    '返回:返回指定的格式串
    '修改人:刘兴宏
    '修改时间:2007/3/6
    '------------------------------------------------------------------------------------------------------
    Dim int位数 As Integer
    Select Case intType
    Case 2  '-药品项目
         int位数 = m售价精度.药品项目小数
    Case 3  '卫材项目
         int位数 = m售价精度.卫材项目小数
    Case Else       '收费项目
         int位数 = m售价精度.收费项目小数
    End Select
    If blnOracle Then
       GetFmtString = "99999999999990." & String(int位数, "9") & ""
    Else
       GetFmtString = "#0." & String(int位数, "0") & ";-#0." & String(int位数, "0") & ";0; "
    End If
End Function
 
