VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmMedicalItems 
   Caption         =   "体检项目管理"
   ClientHeight    =   6720
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10455
   Icon            =   "frmMedicalItems.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6360
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMedicalItems.frx":1CFA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   13361
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
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
   Begin MSComctlLib.TreeView tvw 
      Height          =   1770
      Left            =   120
      TabIndex        =   3
      Top             =   825
      Width           =   2190
      _ExtentX        =   3863
      _ExtentY        =   3122
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ils16"
      Appearance      =   1
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   1244
      BandCount       =   1
      _CBWidth        =   10455
      _CBHeight       =   705
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinWidth1       =   4995
      MinHeight1      =   645
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   645
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   1138
         ButtonWidth     =   820
         ButtonHeight    =   1138
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ilsMenu"
         HotImageList    =   "ilsMenuHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   7
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "预览"
               Object.ToolTipText     =   "预览"
               Object.Tag             =   "预览"
               ImageKey        =   "Preview"
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "打印"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageKey        =   "Print"
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "组合"
               Key             =   "组合"
               Object.ToolTipText     =   "组合"
               Object.Tag             =   "组合"
               ImageKey        =   "Item"
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_4"
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "帮助"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageKey        =   "Quit"
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   8325
      Top             =   3210
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalItems.frx":258E
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalItems.frx":27AE
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalItems.frx":29CE
            Key             =   "Class"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalItems.frx":2BEA
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalItems.frx":2E06
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalItems.frx":3020
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalItems.frx":3240
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalItems.frx":3460
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalItems.frx":3680
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalItems.frx":38A0
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMenuHot 
      Left            =   7515
      Top             =   3210
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   10
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalItems.frx":3AC0
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalItems.frx":3CE0
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalItems.frx":3F00
            Key             =   "Class"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalItems.frx":411C
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalItems.frx":4338
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalItems.frx":468A
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalItems.frx":48AA
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalItems.frx":4ACA
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalItems.frx":4CEA
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalItems.frx":4F0A
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid vsfPrint 
      Height          =   780
      Left            =   7980
      TabIndex        =   4
      Top             =   5505
      Visible         =   0   'False
      Width           =   2250
      _ExtentX        =   3969
      _ExtentY        =   1376
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   270
      MergeCells      =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   1740
      Left            =   3405
      TabIndex        =   5
      Top             =   3885
      Width           =   3030
      _cx             =   5345
      _cy             =   3069
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
      BackColorSel    =   16772055
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
   Begin MSComctlLib.ImageList ils16 
      Left            =   1260
      Top             =   4665
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
            Picture         =   "frmMedicalItems.frx":512A
            Key             =   "Class"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalItems.frx":56C4
            Key             =   "Root"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid lvw 
      Height          =   1740
      Left            =   3495
      TabIndex        =   6
      Top             =   855
      Width           =   3030
      _cx             =   5345
      _cy             =   3069
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
      BackColorSel    =   16772055
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483632
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   2
      HighLight       =   1
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   6
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   ""
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
   Begin VB.Image imgX_S 
      Height          =   45
      Left            =   4860
      MousePointer    =   7  'Size N S
      Top             =   2385
      Width           =   5115
   End
   Begin VB.Image imgY_S 
      Height          =   4395
      Left            =   2670
      MousePointer    =   9  'Size W E
      Top             =   840
      Width           =   45
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrintView 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileOutExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditSelect 
         Caption         =   "检查组合(&S)"
      End
      Begin VB.Menu mnuEditSerial 
         Caption         =   "打印顺序(&N)"
      End
      Begin VB.Menu mnuEditPrintNumber 
         Caption         =   "打印份数(&P)"
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
         Begin VB.Menu mnuViewToolText 
            Caption         =   "文本标签(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuView_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "刷新(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpTopic 
         Caption         =   "帮助主题(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&Web上的中联"
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
      Begin VB.Menu h1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
End
Attribute VB_Name = "frmMedicalItems"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'（１）窗体级变量定义**************************************************************************************************
Private mblnStartUp As Boolean                          '窗体启动标志
Private mstrVsf As String                               '表格列标题
Private mstrKey As String                               '保存以前的选择

Private mlngLoop As Long
Private WithEvents mobjPopMenu As clsPopMenu                '自定义弹出菜单对象
Attribute mobjPopMenu.VB_VarHelpID = -1
Private mbytPopMenu As Byte

Private mblnShowAll As Boolean

Private Enum mCol
    名称 = 0
    类别 = 6
End Enum


'（２）自定义过程或函数************************************************************************************************
Private Function InitLoad() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:初始化数据，发生在窗体的Load事件
    '------------------------------------------------------------------------------------------------------------------
    Dim strVsf As String
    
    On Error GoTo errHand
    
    mstrKey = ""
            
    If Val(GetSetting("ZLSOFT", "私有全局\" & gstrDBUser, "使用个性化风格", "0")) = 1 Then
        '使用个性化设置
        mstrVsf = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "表格标题", mstrVsf)
                        
    End If
    
    mstrVsf = "名称,3000,1,1,1,;编码,900,1,1,1,;简码,1200,1,1,1,;单位,1200,1,1,1,;打印顺序,810,4,1,1,;打印份数,810,4,1,1,;类别,810,4,1,1,;所属分类,900,1,1,1,"
    Call CreateVsf(lvw, mstrVsf)

        
    mstrVsf = "序号,600,1,1,1,;中文名,3000,1,1,1,;编码,900,1,1,1,;英文名,900,1,1,1,;类型,1200,1,1,1,;长度,600,7,1,1,;小数,600,7,1,1,;单位,1080,1,1,1,"
    Call CreateVsf(vsf, mstrVsf)
        
    InitLoad = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub ApplyPrivilege(ByVal strPrivilege As String)
    '------------------------------------------------------------------------------------------------------------------
    '功能： 应用权限处理
    '参数： strPrivilege                    权限
    '------------------------------------------------------------------------------------------------------------------
    
    '调试语句
        
    If InStr(strPrivilege, "项目组合") = 0 And InStr(strPrivilege, "打印排序") = 0 Then
        mnuEdit.Visible = False
    ElseIf InStr(strPrivilege, "项目组合") = 0 Then
        mnuEditSelect.Visible = False
    ElseIf InStr(strPrivilege, "打印排序") = 0 Then
        mnuEditSerial.Visible = False
    End If
    
    tbrThis.Buttons("组合").Visible = mnuEdit.Visible And mnuEditSelect.Visible
    tbrThis.Buttons("Split_4").Visible = tbrThis.Buttons("组合").Visible
End Sub

Private Sub AdjustEnableState()
    '------------------------------------------------------------------------------------------------------------------
    '功能： 调整各功能菜单的可用状态
    '------------------------------------------------------------------------------------------------------------------
    
    mnuFilePrint.Enabled = True
    mnuFilePrintView.Enabled = True
    mnuFileOutExcel.Enabled = True
    
    mnuEditSelect.Enabled = True
    
    If Val(lvw.RowData(lvw.Row)) = 0 Then
        mnuFilePrint.Enabled = False
        mnuFilePrintView.Enabled = False
        mnuFileOutExcel.Enabled = False
        
        mnuEditSelect.Enabled = False
    End If
    
    tbrThis.Buttons("预览").Enabled = mnuFilePrintView.Enabled
    tbrThis.Buttons("打印").Enabled = mnuFilePrint.Enabled

    tbrThis.Buttons("组合").Enabled = mnuEditSelect.Enabled
    
End Sub

Private Sub RefreshStateInfo()
    '------------------------------------------------------------------------------------------------------------------
    '功能：刷新状态栏显示信息
    '------------------------------------------------------------------------------------------------------------------
'    If lvw.SelectedItem Is Nothing Then
'        stbThis.Panels(2).Text = "共有 " & lvw.ListItems.Count & " 个体检类型！"
'    Else
'        If vsf.Rows = 2 And vsf.RowData(1) = 0 Then
'            stbThis.Panels(2).Text = "共有 " & lvw.ListItems.Count & " 个体检类型！"
'        Else
'            stbThis.Panels(2).Text = "共有 " & lvw.ListItems.Count & " 个体检类型，“" & lvw.SelectedItem.Text & "”下有 " & vsf.Rows - 1 & " 个体检项目！"
'        End If
'    End If
    
End Sub

Public Function GetItem(ByRef lngKey As Long, ByVal intFoot As Integer) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：供编辑数据窗体调用，接口函数
    '------------------------------------------------------------------------------------------------------------------
    Dim lngIndex As Long
    Dim objItem As ListItem
    
    On Error GoTo errHand
    
    Set objItem = lvw.ListItems("K" & lngKey)
    If Not (objItem Is Nothing) Then
        
        lngIndex = objItem.Index
        lngIndex = lngIndex + intFoot
        
        Set objItem = Nothing
        Set objItem = lvw.ListItems(lngIndex)
        
        If Not (objItem Is Nothing) Then lngKey = Val(Mid(objItem.Key, 2))
            
        GetItem = True
    Else
        GetItem = False
    End If
    
    Exit Function
    
errHand:
    
End Function

Private Function ClearData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '------------------------------------------------------------------------------------------------------------------
    
    strMenuItem = ";" & strMenuItem & ";"
    
    If InStr(strMenuItem, ";体检类型分类;") > 0 Then
        tvw.Nodes.Clear
    End If
        
    If InStr(strMenuItem, ";体检类型;") > 0 Then
        Call ResetVsf(lvw)
    End If
    If InStr(strMenuItem, ";体检项目;") > 0 Then
        Call ResetVsf(vsf)
    End If
        
End Function

Private Function RefreshData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：刷新/装载数据
    '------------------------------------------------------------------------------------------------------------------
    Dim lngKey As Long
    Dim rs As New ADODB.Recordset
    Dim objNode As Node
    Dim rsPrice As New ADODB.Recordset
    
    On Error GoTo errHand
    
    Select Case strMenuItem
    Case "体检类型分类"
        
        tvw.Nodes.Clear
        
        tvw.Nodes.Add , , "Root", "所有项目", "Root", "Root"
        
        gstrSQL = "select * " & _
             "from (Select DISTINCT ID,上级ID,编码,名称 " & _
                     "From 诊疗分类目录 " & _
                    "Where 类型 = 5 " & _
                    "Start With ID IN (SELECT DISTINCT 分类id FROM 诊疗项目目录 WHERE (撤档时间 = To_Date('30000101', 'YYYYMMDD') Or 撤档时间 is NULL) AND 类别 In ('C','D')) " & _
                   "Connect by Prior 上级ID = ID " & _
                   ") A " & _
            "ORDER BY A.编码"
        
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        Do Until rs.EOF
            If IsNull(rs("上级id")) Then
                tvw.Nodes.Add "Root", tvwChild, "_" & rs("id"), "【" & rs("编码") & "】" & rs("名称"), "Class", "Class"
            Else
                tvw.Nodes.Add "_" & rs("上级id"), tvwChild, "_" & rs("id"), "【" & rs("编码") & "】" & rs("名称"), "Class", "Class"
            End If
            rs.MoveNext
        Loop
        
    Case "体检类型"
        
        lvw.Rows = 2
        lvw.Cell(flexcpText, 1, 0, 1, lvw.Cols - 1) = ""
        lvw.RowData(1) = 0
    
        If tvw.SelectedItem Is Nothing Then Exit Function
                
        gstrSQL = "Select A.ID,A.名称,A.编码,D.简码,A.计算单位 As 单位,Decode(A.类别,'C','检验','检查') As 类别,c.排列顺序 As 打印顺序,e.排列顺序 As 打印份数,B.名称 as 所属分类 " & _
                    "From "
    
        If Val(Mid(tvw.SelectedItem.Key, 2)) > 0 Then
            
            If mblnShowAll Then
                gstrSQL = gstrSQL & "(Select ID,名称 From 诊疗分类目录 Connect by Prior ID=上级id Start With ID = " & Val(Mid(tvw.SelectedItem.Key, 2)) & ") B,"
            Else
                gstrSQL = gstrSQL & "(Select ID,名称 From 诊疗分类目录 where ID = " & Val(Mid(tvw.SelectedItem.Key, 2)) & ") B,"
            End If
            
        Else
            gstrSQL = gstrSQL & "(Select ID,名称 From 诊疗分类目录) B,"
        End If
    
        gstrSQL = gstrSQL & _
                        "诊疗项目目录 A,体检项目排列 c,体检项目排列 e, " & _
                        "(Select * From 诊疗项目别名 Where 性质=1 And 码类=1) D " & _
                    "Where (A.撤档时间 Is Null Or A.撤档时间 = To_Date('3000-01-01','YYYY-MM-DD')) " & _
                            "and A.ID=D.诊疗项目id(+) " & _
                            "and A.类别 In ('C','D') " & _
                            "and B.ID=A.分类ID and a.id=c.诊疗项目id(+) and c.排列性质(+)=1  and a.id=e.诊疗项目id(+) and e.排列性质(+)=2"
                
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        If rs.BOF = False Then
            
            Call FillGrid(lvw, rs)
            
        End If
                
    Case "体检项目"
        
        vsf.Rows = 2
        vsf.Cell(flexcpText, 1, 0, 1, vsf.Cols - 1) = ""
        vsf.RowData(1) = 0
    
        If tvw.SelectedItem Is Nothing Then Exit Function
        
        gstrSQL = "Select c.控件号 As 序号,a.ID,a.编码,a.中文名,a.英文名,Decode(A.类型,0,'数字',1,'文本',2,'日期') As 类型,A.长度,A.小数,A.单位 " & _
                    "From 诊治所见项目 a,病历元素目录 b,病历所见单 c where b.类型=-1 and  b.id=c.元素id and c.所见项id=a.id and c.行=[1] Order By c.控件号 "
        
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(lvw.RowData(lvw.Row)))
        If rs.BOF = False Then
            Call FillGrid(vsf, rs)
        End If
    End Select
    
    RefreshData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function MenuClick(ByVal strMenuItem As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：数据编辑/处理
    '------------------------------------------------------------------------------------------------------------------
    Dim lngKey As Long
    Dim blnTran As Boolean
    Dim lngLoop As Long
    Dim strSQL() As String
                
    On Error GoTo errHand
    
    ReDim Preserve strSQL(1 To 1)
    
    Select Case strMenuItem
    Case "检查组合"
        
        If Val(lvw.RowData(lvw.Row)) = 0 Then Exit Function
        If lvw.TextMatrix(lvw.Row, mCol.类别) = "检验" Then Exit Function
        
        If frmMedicalItemsEdit.ShowEdit(Me, Val(lvw.RowData(lvw.Row))) Then
            Call lvw_AfterRowColChange(0, 0, lvw.Row, lvw.Col)
        End If
        
    End Select
    
    blnTran = True
    gcnOracle.BeginTrans
    
    For lngLoop = 1 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(lngLoop), Me.Caption)
    Next
    
    gcnOracle.CommitTrans
    blnTran = False
    
    
    Call AdjustEnableState
    Call RefreshStateInfo
    
    MenuClick = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
    If blnTran Then gcnOracle.RollbackTrans
End Function

Private Sub PrintData(ByVal bytMode As Byte)
    '------------------------------------------------------------------------------------------------------------------
    '功能： 打印数据
    '参数： bytMode                         打印方式（1-打印；2-预览；3-输出到Excel）
    '------------------------------------------------------------------------------------------------------------------
    Dim objPrint As New zlPrint1Grd
    Dim objRow As zlTabAppRow
    
'    If lvw.SelectedItem Is Nothing Then Exit Sub
'
'    If UserInfo.姓名 = "" Then Call GetUserInfo
'
'    objPrint.Title.Text = "体检项目清单"
'    Call CopyGrid(vsf, vsfPrint)
'
'    Set objRow = New zlTabAppRow
'    objRow.Add "类型：" & lvw.SelectedItem.Text
'    objRow.Add ""
'
'    objPrint.UnderAppRows.Add objRow
'
'    Set objPrint.Body = vsfPrint
'
'    If bytMode = 1 Then bytMode = zlPrintAsk(objPrint)
'
'    If bytMode >= 1 And bytMode <= 3 Then Call zlPrintOrView1Grd(objPrint, bytMode)
        
End Sub

'（３）窗体及其控件的事件处理******************************************************************************************
Private Sub Form_Activate()
    
    If mblnStartUp = False Then Exit Sub
    DoEvents
    
'    Call mnuViewIcon_Click(lvw.View)
    Call mnuViewRefresh_Click
    
    mblnStartUp = False
End Sub

Private Sub Form_Load()
    
    mblnStartUp = True
    
    Call RestoreWinState(Me, App.ProductName)
    Call InitLoad
    Call ApplyPrivilege(gstrPrivs)
    
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    If Me.WindowState = 1 Then Exit Sub
    
    '处理特殊情况
    
    If imgX_S.Top > Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - 1000 Then
        imgX_S.Top = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - 1000
    End If
    
    If imgY_S.Left > Me.ScaleWidth - 1000 Then
        imgY_S.Left = Me.ScaleWidth - 1000
    End If
    
    With tvw
        .Left = 0
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0)
        .Width = imgY_S.Left
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
    
    With imgY_S
        .Top = tvw.Top
        .Height = tvw.Height
    End With
    
    With lvw
        .Left = imgY_S.Left + imgY_S.Width
        .Top = tvw.Top
        .Width = Me.ScaleWidth - .Left
        .Height = imgX_S.Top - .Top
    End With
    
    With imgX_S
        .Left = lvw.Left
        .Width = lvw.Width
    End With
    
    With vsf
        .Left = lvw.Left
        .Top = imgX_S.Top + imgX_S.Height
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - .Top
    End With

    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Cancel = mblnStartUp
    If Cancel Then Exit Sub
    
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "表格标题", mstrVsf)
                
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub imgX_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    On Error Resume Next
    
    imgX_S.Top = imgX_S.Top + Y
    
    If imgX_S.Top < 1500 Then imgX_S.Top = 1500
    If Me.Height - imgX_S.Top - imgX_S.Height < 1000 Then imgX_S.Top = Me.Height - imgX_S.Height - 1000
    
            
    Form_Resize
End Sub

Private Sub imgY_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    imgY_S.Left = imgY_S.Left + X
    
    If imgY_S.Left < 1500 Then imgY_S.Left = 1500
    If Me.Width - imgY_S.Left - imgY_S.Width < 1000 Then imgY_S.Left = Me.Width - imgY_S.Width - 1000

    Form_Resize
End Sub

Private Sub lvw_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    If OldRow <> NewRow Then
        
        Call RefreshData("体检项目")
        
    End If
    
    Call AdjustEnableState
    
End Sub

Private Sub lvw_DblClick()
    '
End Sub

Private Sub mnuEditPrintNumber_Click()
     If frmMedicalItemsPrint.ShowEdit(Me) And Not (tvw.SelectedItem Is Nothing) Then
        Call tvw_NodeClick(tvw.SelectedItem)
     End If
End Sub

Private Sub mnuEditSelect_Click()
    
    Call MenuClick("检查组合")
End Sub

Private Sub mnuEditSerial_Click()
    If frmMedicalItemsArrange.ShowEdit(Me) And Not (tvw.SelectedItem Is Nothing) Then
        Call tvw_NodeClick(tvw.SelectedItem)
    End If
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileOutExcel_Click()
    Call PrintData(3)
End Sub

Private Sub mnuFilePrint_Click()
    Call PrintData(1)
End Sub

Private Sub mnuFilePrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub mnuFilePrintView_Click()
    Call PrintData(2)
End Sub

Private Sub mnuHelpAbout_Click()
    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
End Sub

Private Sub mnuHelpTopic_Click()
   Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub mnuViewRefresh_Click()
    Dim strKey As String
    Dim strKeyClass As String
        
    '保存体检类型分类、体检类型
    If Not (tvw.SelectedItem Is Nothing) Then strKeyClass = tvw.SelectedItem.Key
            
    Call ClearData("体检项目;体检类型;体检类型分类")
    
    Call RefreshData("体检类型分类")
    
    '恢复刷新前选择的体检类型分类
    
    If tvw.Nodes.Count > 0 Then
        tvw.Nodes(1).Selected = True
        tvw.Nodes(1).Expanded = True
    End If
    
    On Error Resume Next
    tvw.Nodes(strKeyClass).Selected = True
    tvw.Nodes(strKeyClass).EnsureVisible
    On Error GoTo 0
    
    If Not (tvw.SelectedItem Is Nothing) Then
    
        Call RefreshData("体检类型")
        Call RefreshData("体检项目")
        
    End If
    
    Call AdjustEnableState
    Call RefreshStateInfo
    
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim intLoop As Integer
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For intLoop = 1 To tbrThis.Buttons.Count
        tbrThis.Buttons(intLoop).Caption = IIf(mnuViewToolText.Checked, tbrThis.Buttons(intLoop).Tag, "")
    Next
    cbrThis.Bands(1).MinHeight = tbrThis.Height
    Call Form_Resize
    
End Sub


Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(tbrThis.hWnd, objPoint)
    
    Select Case Button.Key
    Case "预览"
        Call mnuFilePrintView_Click
    Case "打印"
        
        Call mnuFilePrint_Click
                
    Case "组合"
        Call mnuEditSelect_Click
    Case "帮助"
        Call mnuHelpTopic_Click
    Case "退出"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool
End Sub


Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
    
    Call ClearData("体检类型;体检项目")
    
    Call RefreshData("体检类型")
    
    Call RefreshData("体检项目")
    
    Call AdjustEnableState
    
End Sub

Private Sub vsf_GotFocus()
    vsf.BackColorSel = COLOR.焦点
End Sub

Private Sub vsf_LostFocus()
    vsf.BackColorSel = COLOR.非焦点
End Sub


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

