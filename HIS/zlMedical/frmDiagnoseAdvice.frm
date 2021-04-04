VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmDiagnoseAdvice 
   Caption         =   "体检诊断建议"
   ClientHeight    =   6720
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   10455
   Icon            =   "frmDiagnoseAdvice.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   6720
   ScaleWidth      =   10455
   ShowInTaskbar   =   0   'False
   Begin RichTextLib.RichTextBox rtb 
      Height          =   1935
      Left            =   3105
      TabIndex        =   6
      Top             =   3240
      Width           =   3105
      _ExtentX        =   5477
      _ExtentY        =   3413
      _Version        =   393217
      BorderStyle     =   0
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      TextRTF         =   $"frmDiagnoseAdvice.frx":1CFA
   End
   Begin MSComctlLib.TabStrip tbs 
      Height          =   315
      Left            =   3150
      TabIndex        =   5
      Top             =   2460
      Width           =   3795
      _ExtentX        =   6694
      _ExtentY        =   556
      Style           =   2
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   3
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&1.参考建议"
            Key             =   "参考建议"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&2.评估规则"
            Key             =   "评估规则"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&3.依据项目"
            Key             =   "依据项目"
            Object.ToolTipText     =   "依据的体检项目"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
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
            Picture         =   "frmDiagnoseAdvice.frx":1D97
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
      Left            =   330
      TabIndex        =   4
      Top             =   945
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
   Begin MSComctlLib.ImageList ils16 
      Left            =   1035
      Top             =   4695
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
            Picture         =   "frmDiagnoseAdvice.frx":262B
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":2A7D
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":2D97
            Key             =   "class"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvw 
      Height          =   1485
      Left            =   3135
      TabIndex        =   3
      Top             =   870
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   2619
      View            =   3
      Arrange         =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ils32"
      SmallIcons      =   "ils16"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   405
      Top             =   4695
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":3331
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":3783
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   10455
      _ExtentX        =   18441
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   10455
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinWidth1       =   4995
      MinHeight1      =   720
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   720
         Left            =   30
         TabIndex        =   2
         Top             =   30
         Width           =   10335
         _ExtentX        =   18230
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ilsMenu"
         HotImageList    =   "ilsMenuHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   13
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
               Caption         =   "分类"
               Key             =   "分类"
               Object.ToolTipText     =   "分类"
               Object.Tag             =   "分类"
               ImageKey        =   "Class"
               Style           =   5
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_2"
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "增加"
               Key             =   "增加"
               Object.ToolTipText     =   "增加"
               Object.Tag             =   "增加"
               ImageKey        =   "Add"
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "修改"
               Object.ToolTipText     =   "修改"
               Object.Tag             =   "修改"
               ImageKey        =   "Modify"
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "删除"
               Object.ToolTipText     =   "删除"
               Object.Tag             =   "删除"
               ImageKey        =   "Delete"
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_3"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "列表"
               Key             =   "列表"
               Object.ToolTipText     =   "列表"
               Object.Tag             =   "列表"
               ImageKey        =   "View"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Large"
                     Text            =   "大图标(&G)"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Small"
                     Text            =   "小图标(&M)"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "List"
                     Text            =   "列表(&L)"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "Detail"
                     Text            =   "详细资料(&D)"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_5"
               Style           =   3
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "帮助"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageKey        =   "Help"
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
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
            Picture         =   "frmDiagnoseAdvice.frx":3A9D
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":3CBD
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":3EDD
            Key             =   "Class"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":40F9
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":4315
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":452F
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":474F
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":496F
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":4B8F
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":4DAF
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
            Picture         =   "frmDiagnoseAdvice.frx":4FCF
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":51EF
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":540F
            Key             =   "Class"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":562B
            Key             =   "Item"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":5847
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":5B99
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":5DB9
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":5FD9
            Key             =   "View"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":61F9
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDiagnoseAdvice.frx":6419
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   1995
      Left            =   4035
      TabIndex        =   7
      Top             =   3630
      Width           =   4275
      _cx             =   7541
      _cy             =   3519
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
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483633
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
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   270
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
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
      Begin VB.Line lnY 
         Index           =   0
         Visible         =   0   'False
         X1              =   1965
         X2              =   1965
         Y1              =   1590
         Y2              =   2805
      End
      Begin VB.Line lnX 
         Index           =   0
         Visible         =   0   'False
         X1              =   1140
         X2              =   2925
         Y1              =   1725
         Y2              =   1725
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfItem 
      Height          =   1995
      Left            =   6435
      TabIndex        =   8
      Top             =   3060
      Width           =   4275
      _cx             =   7541
      _cy             =   3519
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
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483633
      GridColorFixed  =   -2147483633
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
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   3
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   270
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   -1  'True
      FormatString    =   ""
      ScrollTrack     =   0   'False
      ScrollBars      =   3
      ScrollTips      =   0   'False
      MergeCells      =   1
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
      Begin VB.Line lnX1 
         Index           =   0
         Visible         =   0   'False
         X1              =   1140
         X2              =   2925
         Y1              =   1725
         Y2              =   1725
      End
      Begin VB.Line lnY1 
         Index           =   0
         Visible         =   0   'False
         X1              =   1965
         X2              =   1965
         Y1              =   1590
         Y2              =   2805
      End
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
      Begin VB.Menu mnuEditClass 
         Caption         =   "诊断分类(&C)"
         Begin VB.Menu mnuEditClassAdd 
            Caption         =   "增加分类(&A)"
         End
         Begin VB.Menu mnuEditClassModify 
            Caption         =   "修改分类(&M)"
         End
         Begin VB.Menu mnuEditClassDelete 
            Caption         =   "删除分类(&D)"
         End
      End
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditAdd 
         Caption         =   "增加诊断(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "修改诊断(&M)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "删除诊断(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEdit_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditAdvice 
         Caption         =   "参考建议(S)"
      End
      Begin VB.Menu mnuEditCondition 
         Caption         =   "评估规则(&P)"
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
      Begin VB.Menu mnuViewIcon 
         Caption         =   "大图标(&G)"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "小图标(&M)"
         Index           =   1
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "列表(&L)"
         Index           =   2
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "详细资料(&D)"
         Index           =   3
      End
      Begin VB.Menu mnuView_2 
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
Attribute VB_Name = "frmDiagnoseAdvice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'（１）窗体级变量定义**************************************************************************************************
Private mblnStartUp As Boolean                          '窗体启动标志
Private mstrVsf As String                               '表格列标题
Private mstrKey As String                               '保存以前的选择
Private Const mstrLvw As String = "名称,2100,0,1;编码,900,0,0;简码,900,0,0;是否疾病,900,0,0;诊断建议,2100,0,0;复查间隔,900,0,0;随访期限,900,0,0;所属分类,1500,0,0"
Private mlngLoop As Long
Private WithEvents mobjPopMenu As clsPopMenu                '自定义弹出菜单对象
Attribute mobjPopMenu.VB_VarHelpID = -1
Private mbytPopMenu As Byte

'（２）自定义过程或函数************************************************************************************************
Private Function InitLoad() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:初始化数据，发生在窗体的Load事件
    '------------------------------------------------------------------------------------------------------------------
    Dim strVsf As String
    
    On Error GoTo errHand
    
    mstrKey = ""
    lvw.Tag = "可变化的"
    
    If lvw.ListItems.Count = 0 Then zlControl.LvwSelectColumns lvw, mstrLvw, True

    strVsf = "组名,900,1,1,1,;性别,600,1,1,1,;开始年龄,900,1,1,1,;结束年龄,900,1,1,1,;项目,2100,1,1,1,;条件,900,1,1,1,;项目值,900,1,1,1,"
    Call CreateVsf(vsf, strVsf)
    
    vsf.MergeCol(0) = True
    
    
    strVsf = "名称,2400,1,1,1,;编码,1200,1,1,1,;类别,900,1,1,1,"
    Call CreateVsf(vsfItem, strVsf)
    vsfItem.Cols = vsfItem.Cols + 1
    vsfItem.ColWidth(vsfItem.Cols - 1) = 15
    
    
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
    'strPrivilege = "基本;增删改"
    
    '不具有“增删改”权限时
    If InStr(strPrivilege, "增删改") = 0 Then
        mnuEdit.Visible = False
    End If
    
    tbrThis.Buttons("分类").Visible = mnuEdit.Visible
    
    tbrThis.Buttons("增加").Visible = mnuEdit.Visible
    tbrThis.Buttons("修改").Visible = mnuEdit.Visible
    tbrThis.Buttons("删除").Visible = mnuEdit.Visible
    
    tbrThis.Buttons("Split_2").Visible = mnuEdit.Visible
    tbrThis.Buttons("Split_3").Visible = mnuEdit.Visible

End Sub

Private Sub AdjustEnableState()
    '------------------------------------------------------------------------------------------------------------------
    '功能： 调整各功能菜单的可用状态
    '------------------------------------------------------------------------------------------------------------------
    
    mnuFilePrint.Enabled = True
    mnuFilePrintView.Enabled = True
    mnuFileOutExcel.Enabled = True
    
    mnuEditClassAdd.Enabled = True
    mnuEditClassModify.Enabled = True
    mnuEditClassDelete.Enabled = True
    
    mnuEditClass.Enabled = True
    
    mnuEditAdd.Enabled = True
    mnuEditModify.Enabled = True
    mnuEditDelete.Enabled = True
    mnuEditAdvice.Enabled = True
    mnuEditCondition.Enabled = True
    
    If lvw.ListItems.Count = 0 Then
                
        mnuEditModify.Enabled = False
        mnuEditDelete.Enabled = False
        mnuEditAdvice.Enabled = False
        mnuEditCondition.Enabled = False
    End If
    
    If Not (tvw.SelectedItem Is Nothing) Then
        If tvw.SelectedItem.Key = "K0" Then
            mnuEditClassModify.Enabled = False
            mnuEditClassDelete.Enabled = False
        End If
    End If
    
    tbrThis.Buttons("预览").Enabled = mnuFilePrintView.Enabled
    tbrThis.Buttons("打印").Enabled = mnuFilePrint.Enabled
    
    tbrThis.Buttons("分类").Enabled = mnuEditClassModify.Enabled Or mnuEditClassAdd.Enabled
    tbrThis.Buttons("增加").Enabled = mnuEditAdd.Enabled
    tbrThis.Buttons("修改").Enabled = mnuEditModify.Enabled
    tbrThis.Buttons("删除").Enabled = mnuEditDelete.Enabled
        
End Sub

Private Sub RefreshStateInfo()
    '------------------------------------------------------------------------------------------------------------------
    '功能：刷新状态栏显示信息
    '------------------------------------------------------------------------------------------------------------------
    
    stbThis.Panels(2).Text = "共有 " & lvw.ListItems.Count & " 个体检诊断！"
    
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

Public Function EditRefresh(ByVal strMenuItem As String, ByVal lngKey As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：供编辑数据窗体调用，接口函数
    '------------------------------------------------------------------------------------------------------------------
    Dim lngSvrKey As Long
    
    On Error GoTo errHand

    Select Case strMenuItem
    Case "体检诊断分类"
        
        Call ClearData("体检诊断分类;体检诊断目录;体检诊断附项")
        
        Call RefreshData("体检诊断分类")
        
        On Error Resume Next
        tvw.Nodes("K" & lngKey).Selected = True
        tvw.Nodes("K" & lngKey).EnsureVisible
        On Error GoTo 0
        
        If tvw.Nodes.Count > 0 Then
            If tvw.SelectedItem Is Nothing Then
                tvw.Nodes(1).Selected = True
                tvw.Nodes(1).EnsureVisible
            End If
        End If
        
        Call RefreshData("体检诊断目录")
        
        Call tbs_Click
        
        
    Case "体检诊断目录"
    
        Call ClearData("体检诊断目录;体检诊断附项")
        
        Call RefreshData("体检诊断目录")
        
        '恢复体检类型
        Call zlControl.LvwRestoreItem(lvw, "K" & lngKey)
            
        Call tbs_Click
   
    End Select
    
    Call AdjustEnableState
    Call RefreshStateInfo
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function ClearData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:
    '------------------------------------------------------------------------------------------------------------------
    
    strMenuItem = ";" & strMenuItem & ";"
    
    If InStr(strMenuItem, ";体检诊断分类;") > 0 Then
        tvw.Nodes.Clear
    End If
    
    If InStr(strMenuItem, ";体检诊断目录;") > 0 Then
        lvw.ListItems.Clear
    End If
    
    If InStr(strMenuItem, ";体检诊断附项;") > 0 Then
        rtb.Text = ""
        Call ResetVsf(vsf)
        Call ResetVsf(vsfItem)
        Call AppendRows(vsf, lnX, lnY)
        Call AppendRows(vsfItem, lnX1, lnY1)
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
    Case "体检诊断分类"
        
        gstrSQL = GetPublicSQL(SQL.体检诊断分类)
        
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        
        If rs.BOF = False Then Call FillTreeData(tvw, rs)
        
    Case "体检诊断目录"
        
        If Val(Mid(tvw.SelectedItem.Key, 2)) = 0 Then
            
            gstrSQL = " SELECT E.序号 AS ID," & _
                              "E.编码," & _
                              "E.名称," & _
                              "E.简码," & _
                              "E.诊断建议,Decode(E.复查间隔,Null,'',Trim(To_Char(E.复查间隔,'99999'))||'月') As 复查间隔,Decode(E.随访期限,Null,'',Trim(To_Char(E.随访期限,'99999'))||'月') as 随访期限," & _
                              "Decode(E.是否疾病,1,'√','') As 是否疾病," & _
                              "1 as 图标, "
                                         
            gstrSQL = gstrSQL & _
                              "F.名称 AS 所属分类 " & _
                         "FROM 体检诊断建议 E, 体检诊断建议 F " & _
                        "WHERE E.末级 = 1 AND E.上级序号 = F.序号(+)"
        Else
            gstrSQL = " SELECT E.序号 AS ID," & _
                      "E.编码," & _
                      "E.名称," & _
                      "E.简码," & _
                        "E.诊断建议,Decode(E.复查间隔,Null,'',Trim(To_Char(E.复查间隔,'99999'))||'月') as 复查间隔,Decode(E.随访期限,Null,'',Trim(To_Char(E.随访期限,'99999'))||'月') as 随访期限," & _
                      "Decode(E.是否疾病,1,'√','') As 是否疾病," & _
                      "1 as 图标,"
                                         
                gstrSQL = gstrSQL & _
                      "F.名称 AS 所属分类 " & _
                 "FROM 体检诊断建议 E, 体检诊断建议 F " & _
                "WHERE E.末级 = 1 AND E.上级序号 = F.序号 AND E.上级序号=[1]"
            
        End If
        
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Mid(tvw.SelectedItem.Key, 2)))
        If rs.BOF = False Then Call FillLvw(lvw, rs)
        
    Case "依据项目"
        
        If lvw.SelectedItem Is Nothing Then Exit Function
        
        lngKey = Val(Mid(lvw.SelectedItem.Key, 2))
        If lngKey = 0 Then Exit Function
        
        gstrSQL = "SELECT B.ID,B.名称,b.编码,c.名称 As 类别 FROM 体检诊断依据 A,诊疗项目目录 B,诊疗项目类别 c WHERE A.诊疗项目id=B.ID And A.诊断序号=[1] And c.编码=b.类别"
        
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
        If rs.BOF = False Then
            Call FillGrid(vsfItem, rs)
        
            Call AppendRows(vsfItem, lnX1, lnY1)
        End If
    Case "体检诊断建议"
    
        If lvw.SelectedItem Is Nothing Then Exit Function
        
        lngKey = Val(Mid(lvw.SelectedItem.Key, 2))
        If lngKey = 0 Then Exit Function
        
        gstrSQL = "select 参考建议 from 体检诊断建议 WHERE 序号=[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
        
        If rs.BOF = False Then rtb.Text = zlCommFun.NVL(rs("参考建议").Value)
        
    Case "体检诊断条件"
        
        If lvw.SelectedItem Is Nothing Then Exit Function
        
        Call ResetVsf(vsf)
        Call AppendRows(vsf, lnX, lnY)
        
        lngKey = Val(Mid(lvw.SelectedItem.Key, 2))
        If lngKey = 0 Then Exit Function
        
        gstrSQL = "SELECT b.ID,a.分组名 As 组名,b.中文名 As 项目,a.关系式 As 条件,a.条件值 As 项目值,a.性别,a.开始年龄,a.结束年龄 from 体检诊断评估 A,诊治所见项目 B WHERE A.项目id=B.ID AND A.诊断序号=[1] ORDER BY A.分组名"
                   
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngKey)
        If rs.BOF = False Then
            Call LoadGrid(vsf, rs)
            Call AppendRows(vsf, lnX, lnY)
        End If
        
    End Select
    
    RefreshData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
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
    Case "增加分类"
        
        If tvw.SelectedItem Is Nothing Then Exit Function
        
        If Not frmDiagnoseAdviceClass.ShowEdit(Me, 0, Val(Mid(tvw.SelectedItem.Key, 2))) Then Exit Function
        
        
    Case "修改分类"
        
        If tvw.SelectedItem Is Nothing Then Exit Function
        If tvw.SelectedItem.Key = "K0" Then Exit Function
                        
        If Not frmDiagnoseAdviceClass.ShowEdit(Me, Val(Mid(tvw.SelectedItem.Key, 2)), Val(Mid(tvw.SelectedItem.Parent.Key, 2))) Then Exit Function
        
        
    Case "删除分类"
        
        If tvw.SelectedItem Is Nothing Then Exit Function
        If tvw.SelectedItem.Key = "K0" Then Exit Function
        
        If MsgBox("你真的要删除“" & tvw.SelectedItem.Text & "”分类？" & vbCrLf & "删除分类同时也删除对应的体检诊断建议。", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        
        lngKey = Val(Mid(tvw.SelectedItem.Key, 2))
        If lngKey = 0 Then Exit Function
        
        strSQL(ReDimArray(strSQL)) = "ZL_体检诊断建议_DELETE(" & lngKey & ")"
        
    Case "增加诊断"
        If tvw.SelectedItem Is Nothing Then Exit Function
        
        If Not frmDiagnoseAdviceEdit.ShowEdit(Me, 0, Val(Mid(tvw.SelectedItem.Key, 2))) Then Exit Function
        
    Case "修改诊断"
        If tvw.SelectedItem Is Nothing Then Exit Function
        If lvw.SelectedItem Is Nothing Then Exit Function
        lngKey = Val(Mid(lvw.SelectedItem.Key, 2))
        If lngKey = 0 Then Exit Function
        
        If Not frmDiagnoseAdviceEdit.ShowEdit(Me, lngKey, Val(Mid(tvw.SelectedItem.Key, 2))) Then Exit Function
        
    Case "删除诊断"
        If tvw.SelectedItem Is Nothing Then Exit Function
        If lvw.SelectedItem Is Nothing Then Exit Function
        
        If MsgBox("你真的要删除“" & lvw.SelectedItem.Text & "”？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
        
        lngKey = Val(Mid(lvw.SelectedItem.Key, 2))
        If lngKey = 0 Then Exit Function
        
        strSQL(ReDimArray(strSQL)) = "ZL_体检诊断建议_DELETE(" & lngKey & ")"
        
    Case "参考建议"
        If lvw.SelectedItem Is Nothing Then Exit Function
        
        lngKey = Val(Mid(lvw.SelectedItem.Key, 2))
        If lngKey = 0 Then Exit Function
        
        If Not frmDiagnoseAdviceContent.ShowEdit(Me, lngKey) Then Exit Function
    
    Case "评估规则"
        
        If lvw.SelectedItem Is Nothing Then Exit Function
        
        lngKey = Val(Mid(lvw.SelectedItem.Key, 2))
        If lngKey = 0 Then Exit Function
        
        If Not frmDiagnoseAdviceEvaluate.ShowEdit(Me, lngKey) Then Exit Function
        
    End Select
    
    blnTran = True
    gcnOracle.BeginTrans
    
    For lngLoop = 1 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(lngLoop), Me.Caption)
    Next
    
    gcnOracle.CommitTrans
    blnTran = False
        
    Select Case strMenuItem
    Case "删除分类"
        
        If Not (tvw.SelectedItem Is Nothing) Then tvw.Nodes.Remove tvw.SelectedItem.Index
        
        Call ClearData("体检诊断建议")
        Call RefreshData("体检诊断建议")
        If Not (lvw.SelectedItem Is Nothing) Then Call RefreshData("体检诊断附项")
                
        
    Case "删除诊断"
    
        '删除行
        lngLoop = lvw.SelectedItem.Index
        lvw.ListItems.Remove lngLoop
        Call NextLvwPos(lvw, lngLoop)
        
    Case "参考建议", "评估规则"
        Call tbs_Click
    End Select
    
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
    Dim objPrint As New zlPrintLvw
        
    If tvw.SelectedItem Is Nothing Then Exit Sub
    If lvw.SelectedItem Is Nothing Then Exit Sub
    
    If UserInfo.姓名 = "" Then Call GetUserInfo

    objPrint.Title.Text = "体检诊断建议清单"
    
    objPrint.UnderAppItems.Add "分类：" & tvw.SelectedItem.Text
    objPrint.UnderAppItems.Add ""
        
    Set objPrint.Body.objData = lvw

    If bytMode = 1 Then bytMode = zlPrintAsk(objPrint)

    If bytMode >= 1 And bytMode <= 3 Then Call zlPrintOrViewLvw(objPrint, bytMode)
        
End Sub

'（３）窗体及其控件的事件处理******************************************************************************************
Private Sub Form_Activate()
    
    If mblnStartUp = False Then Exit Sub
    DoEvents
    
    Call mnuViewIcon_Click(lvw.View)
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
    
    With tbs
        .Left = imgX_S.Left
        .Top = imgX_S.Top + imgX_S.Height
        .Width = imgX_S.Width
    End With
    
    With rtb
        .Left = lvw.Left
        .Top = tbs.Top + tbs.Height + 30
        .Width = Me.ScaleWidth - .Left
        .Height = Me.ScaleHeight - IIf(stbThis.Visible, stbThis.Height, 0) - .Top
    End With
    
    With vsf
        .Left = rtb.Left
        .Top = rtb.Top
        .Width = rtb.Width
        .Height = rtb.Height
    End With

    With vsfItem
        .Left = rtb.Left
        .Top = rtb.Top
        .Width = rtb.Width
        .Height = rtb.Height
    End With
    
    Call AppendRows(vsf, lnX, lnY)
    Call AppendRows(vsfItem, lnX1, lnY1)
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

Private Sub lvw_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
        
    Call zlControl.LvwSortColumn(lvw, ColumnHeader.Index)
    
End Sub

Private Sub lvw_DblClick()
    If mnuEdit.Visible And mnuEditModify.Visible And mnuEditModify.Enabled Then Call mnuEditModify_Click
End Sub

Private Sub lvw_ItemClick(ByVal Item As MSComctlLib.ListItem)

    If mstrKey = Item.Key Then Exit Sub
    mstrKey = Item.Key
    
    Call ClearData("体检诊断附项")
    Call tbs_Click
    
    '恢复
    Call AdjustEnableState
    Call RefreshStateInfo
    
End Sub

Private Sub lvw_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call lvw_DblClick
End Sub

Private Sub lvw_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    
    mbytPopMenu = 2
    Set mobjPopMenu = New clsPopMenu
    mobjPopMenu.ShowPopupMenuByCursor

End Sub

Private Sub mnuEditAdvice_Click()
    Call MenuClick("参考建议")
End Sub

Private Sub mnuEditClassAdd_Click()
    Call MenuClick("增加分类")
End Sub

Private Sub mnuEditClassDelete_Click()
    Call MenuClick("删除分类")
End Sub

Private Sub mnuEditClassModify_Click()
    Call MenuClick("修改分类")
End Sub

Private Sub mnuEditCondition_Click()
    Call MenuClick("评估规则")
End Sub

Private Sub mnuEditDelete_Click()
    Call MenuClick("删除诊断")
End Sub

Private Sub mnuEditModify_Click()
    Call MenuClick("修改诊断")
End Sub

Private Sub mnuEditAdd_Click()
    Call MenuClick("增加诊断")
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

Private Sub mnuViewIcon_Click(Index As Integer)
    mnuViewIcon(0).Checked = False
    mnuViewIcon(1).Checked = False
    mnuViewIcon(2).Checked = False
    mnuViewIcon(3).Checked = False
    
    mnuViewIcon(Index).Checked = True
    
    lvw.View = Index
End Sub

Private Sub mnuViewRefresh_Click()
    Dim strKey As String
    Dim strKeyClass As String
        
    '保存体检类型分类、体检类型
    If Not (tvw.SelectedItem Is Nothing) Then strKeyClass = tvw.SelectedItem.Key
    strKey = zlControl.LvwSaveItem(lvw)
            
    Call ClearData("体检诊断分类;体检诊断目录;体检诊断附项")
    
    Call RefreshData("体检诊断分类")
    
    '恢复刷新前选择的体检类型分类
    
    tvw.Nodes(1).Selected = True
    tvw.Nodes(1).Expanded = True
    
    On Error Resume Next
    tvw.Nodes(strKeyClass).Selected = True
    tvw.Nodes(strKeyClass).EnsureVisible
    On Error GoTo 0
    
    If Not (tvw.SelectedItem Is Nothing) Then
        Call RefreshData("体检诊断目录")
        
        '恢复刷新前选择的体检类型
        Call zlControl.LvwRestoreItem(lvw, strKey)
        
        Call tbs_Click
        
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

Private Sub mobjPopMenu_MenuBeforeShow(Cancel As Boolean)
    
    Select Case mbytPopMenu
    Case 1
        
        If mnuEdit.Visible = False Then Exit Sub
        
        If mnuEditClassAdd.Visible Then mobjPopMenu.Add 1, mnuEditClassAdd.Caption, , , mnuEditClassAdd.Enabled
        If mnuEditClassModify.Visible Then mobjPopMenu.Add 2, mnuEditClassModify.Caption, , , mnuEditClassModify.Enabled
        If mnuEditClassDelete.Visible Then mobjPopMenu.Add 3, mnuEditClassDelete.Caption, , , mnuEditClassDelete.Enabled
        
    Case 2
        
        If mnuEdit.Visible = False Then Exit Sub
        
        If mnuEditAdd.Visible Then mobjPopMenu.Add 1, mnuEditAdd.Caption, , , mnuEditAdd.Enabled
        If mnuEditModify.Visible Then mobjPopMenu.Add 2, mnuEditModify.Caption, , , mnuEditModify.Enabled
        If mnuEditDelete.Visible Then mobjPopMenu.Add 3, mnuEditDelete.Caption, , , mnuEditDelete.Enabled
            
        mobjPopMenu.Add 4, "-", , 2, True
        
        If mnuEditAdvice.Visible Then mobjPopMenu.Add 5, mnuEditAdvice.Caption, , , mnuEditAdvice.Enabled
        If mnuEditCondition.Visible Then mobjPopMenu.Add 6, mnuEditCondition.Caption, , , mnuEditCondition.Enabled
        
    End Select
    
End Sub

Private Sub mobjPopMenu_MenuClick(ByVal Key As Long, ByVal Caption As String)
    Select Case mbytPopMenu
    Case 1
        Select Case Key
        Case 1
            Call mnuEditClassAdd_Click
        Case 2
            Call mnuEditClassModify_Click
        Case 3
            Call mnuEditClassDelete_Click
        End Select
    Case 2
        Select Case Key
        Case 1
            Call mnuEditAdd_Click
        Case 2
            Call mnuEditModify_Click
        Case 3
            Call mnuEditDelete_Click
        Case 5
            Call mnuEditAdvice_Click
        Case 6
            Call mnuEditCondition_Click
        End Select
    End Select
End Sub


Private Sub rtb_DblClick()
    If mnuEdit.Visible And mnuEditAdvice.Visible And mnuEditAdvice.Enabled Then
        Call mnuEditAdvice_Click
    End If
End Sub

Private Sub rtb_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call rtb_DblClick
    End If
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(tbrThis.hWnd, objPoint)
    
    Select Case Button.Key
    Case "预览"
        Call mnuFilePrintView_Click
    Case "打印"
        
        Call mnuFilePrint_Click
        
    Case "分类"
                
        mbytPopMenu = 1
        Set mobjPopMenu = New clsPopMenu
        Call mobjPopMenu.ShowPopupMenu(objPoint.X * 15 + Button.Left - 15, objPoint.Y * 15 + Button.Top + Button.Height + 15)
        
    Case "增加"
        Call mnuEditAdd_Click
    Case "修改"
        Call mnuEditModify_Click
    Case "删除"
        Call mnuEditDelete_Click
    Case "列表"
        Call mnuViewIcon_Click(IIf(lvw.View = 3, 0, lvw.View + 1))
    Case "帮助"
        Call mnuHelpTopic_Click
    Case "退出"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tbrThis_ButtonDropDown(ByVal Button As MSComctlLib.Button)
    Call tbrThis_ButtonClick(Button)
End Sub

Private Sub tbrThis_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Select Case ButtonMenu.Key
    Case "Large"
        Call mnuViewIcon_Click(0)
    Case "Small"
        Call mnuViewIcon_Click(1)
    Case "List"
        Call mnuViewIcon_Click(2)
    Case "Detail"
        Call mnuViewIcon_Click(3)
    End Select
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool
End Sub

Private Sub tbs_Click()

    If lvw.SelectedItem Is Nothing Then Exit Sub
    
'    rtb.Visible = True
'    vsf.Visible = True
'    vsf.Visible = True
    Select Case tbs.SelectedItem.Key
    Case "参考建议"
        rtb.Visible = True
        vsf.Visible = False
        vsfItem.Visible = False
        
        Call RefreshData("体检诊断建议")

        
    Case "评估规则"
        
        rtb.Visible = False
        vsf.Visible = True
        vsfItem.Visible = False
        Call RefreshData("体检诊断条件")

    
        
    Case "依据项目"
        
        rtb.Visible = False
        vsf.Visible = False
        vsfItem.Visible = True
        
        Call RefreshData("依据项目")

        
    End Select
End Sub

Private Sub tvw_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If Button = 2 Then
    
        mbytPopMenu = 1
        Set mobjPopMenu = New clsPopMenu
        mobjPopMenu.ShowPopupMenuByCursor
        
    End If
End Sub

Private Sub tvw_NodeClick(ByVal Node As MSComctlLib.Node)
    
    Call ClearData("体检诊断目录;体检诊断附项")
    
    Call RefreshData("体检诊断目录")
    Call tbs_Click
    
    Call AdjustEnableState
    Call RefreshStateInfo
    
End Sub

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call AppendRows(vsf, lnX, lnY)
End Sub

Private Sub vsf_DblClick()
    If mnuEdit.Visible And mnuEditAdvice.Visible And mnuEditCondition.Enabled Then
        Call mnuEditCondition_Click
    End If
End Sub

Private Sub vsf_GotFocus()
    vsf.BackColorSel = COLOR.焦点
End Sub

Private Sub vsf_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Call vsf_DblClick
    End If
End Sub

Private Sub vsf_LostFocus()
    vsf.BackColorSel = COLOR.非焦点
End Sub

Private Sub vsfItem_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call AppendRows(vsfItem, lnX1, lnY1)
End Sub

Private Sub vsfItem_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Call AppendRows(vsfItem, lnX1, lnY1)
End Sub

Private Sub vsfItem_GotFocus()
    vsfItem.BackColorSel = COLOR.焦点
End Sub

Private Sub vsfItem_LostFocus()
    vsfItem.BackColorSel = COLOR.非焦点
End Sub


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

