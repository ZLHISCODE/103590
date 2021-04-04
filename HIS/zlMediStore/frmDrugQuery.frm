VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Begin VB.Form frmDrugQuery 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   Caption         =   "药品库存查询"
   ClientHeight    =   7110
   ClientLeft      =   165
   ClientTop       =   825
   ClientWidth     =   11325
   Icon            =   "frmDrugQuery.frx":0000
   LockControls    =   -1  'True
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   7110
   ScaleWidth      =   11325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '屏幕中心
   WindowState     =   2  'Maximized
   Begin VB.PictureBox picColor 
      BorderStyle     =   0  'None
      Height          =   255
      Left            =   4920
      ScaleHeight     =   255
      ScaleWidth      =   4935
      TabIndex        =   16
      Top             =   5760
      Width           =   4935
      Begin VB.PictureBox picColor2 
         BackColor       =   &H00FF0000&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   2040
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   19
         Top             =   0
         Width           =   260
      End
      Begin VB.PictureBox picColor1 
         BackColor       =   &H000000C0&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   1
         Left            =   600
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   18
         Top             =   0
         Width           =   260
      End
      Begin VB.PictureBox picColor1 
         BackColor       =   &H000000FF&
         BorderStyle     =   0  'None
         Height          =   255
         Index           =   0
         Left            =   3000
         ScaleHeight     =   255
         ScaleWidth      =   255
         TabIndex        =   17
         Top             =   0
         Width           =   260
      End
      Begin VB.Label lblColor2 
         AutoSize        =   -1  'True
         Caption         =   "近效期"
         Height          =   180
         Left            =   2325
         TabIndex        =   24
         Top             =   30
         Width           =   540
      End
      Begin VB.Label lblColor1 
         AutoSize        =   -1  'True
         Caption         =   "规格："
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   1
         Left            =   0
         TabIndex        =   23
         Top             =   30
         Width           =   540
      End
      Begin VB.Label lblColor1 
         AutoSize        =   -1  'True
         Caption         =   "停用"
         Height          =   180
         Index           =   2
         Left            =   900
         TabIndex        =   22
         Top             =   30
         Width           =   360
      End
      Begin VB.Label lblColor1 
         AutoSize        =   -1  'True
         Caption         =   "批次："
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   3
         Left            =   1440
         TabIndex        =   21
         Top             =   30
         Width           =   540
      End
      Begin VB.Label lblColor1 
         AutoSize        =   -1  'True
         Caption         =   "可用数量不足或预减"
         Height          =   180
         Index           =   0
         Left            =   3285
         TabIndex        =   20
         Top             =   30
         Width           =   1620
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfColSel 
      Height          =   1095
      Left            =   9360
      TabIndex        =   15
      Top             =   1440
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
      FormatString    =   $"frmDrugQuery.frx":0982
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
   Begin VB.CheckBox Chk剂型 
      Appearance      =   0  'Flat
      Caption         =   "全选"
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   70
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   5655
      Value           =   1  'Checked
      Width           =   675
   End
   Begin VB.PictureBox picFind 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000004&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   1470
      ScaleHeight     =   255
      ScaleWidth      =   2415
      TabIndex        =   9
      TabStop         =   0   'False
      Top             =   6810
      Width           =   2415
      Begin VB.TextBox txt药品信息 
         Appearance      =   0  'Flat
         Height          =   255
         Left            =   780
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   0
         Width           =   1245
      End
      Begin VB.Label lbl药品信息 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "查找药品"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   30
         TabIndex        =   11
         Top             =   30
         Width           =   720
      End
   End
   Begin MSComctlLib.ListView lst剂型_S 
      Height          =   720
      Left            =   0
      TabIndex        =   8
      Top             =   5910
      Width           =   2865
      _ExtentX        =   5054
      _ExtentY        =   1270
      View            =   2
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      Checkboxes      =   -1  'True
      HoverSelection  =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6744
      Width           =   11328
      _ExtentX        =   19976
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmDrugQuery.frx":09D0
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14182
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   1429
            MinWidth        =   1411
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
            Object.ToolTipText     =   "当前数字键状态"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1411
            MinWidth        =   1411
            Text            =   "大写"
            TextSave        =   "大写"
            Key             =   "STACAPS"
            Object.ToolTipText     =   "当前大写键状态"
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
   Begin MSComctlLib.TreeView tvwSection_S 
      Height          =   4350
      Left            =   60
      TabIndex        =   1
      Top             =   1275
      Width           =   2790
      _ExtentX        =   4921
      _ExtentY        =   7673
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   494
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "imglvw"
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList imglvw 
      Left            =   2985
      Top             =   2205
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
            Picture         =   "frmDrugQuery.frx":1264
            Key             =   "root"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":2F6E
            Key             =   "child"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":4C78
            Key             =   "clock"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picVLine_S 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0C0C0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5460
      Left            =   2940
      MousePointer    =   9  'Size W E
      ScaleHeight     =   5460
      ScaleWidth      =   45
      TabIndex        =   2
      Top             =   1305
      Width           =   45
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   1125
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11325
      _ExtentX        =   19976
      _ExtentY        =   1984
      BandCount       =   2
      _CBWidth        =   11325
      _CBHeight       =   1125
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinHeight1      =   720
      Width1          =   2730
      NewRow1         =   0   'False
      Caption2        =   "库房"
      Child2          =   "cboStock"
      MinHeight2      =   300
      Width2          =   6780
      NewRow2         =   -1  'True
      Begin VB.ComboBox cboStock 
         Height          =   300
         Left            =   585
         TabIndex        =   6
         Text            =   "cboStock"
         Top             =   780
         Width           =   10650
      End
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   720
         Left            =   165
         TabIndex        =   4
         Top             =   30
         Width           =   11070
         _ExtentX        =   19526
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgTbrStard"
         HotImageList    =   "imgTbrHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   14
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "预览"
               Object.ToolTipText     =   "预览"
               Object.Tag             =   "预览"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "打印"
               Description     =   "打印"
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               Caption         =   "重置"
               Key             =   "重置"
               Object.ToolTipText     =   "重置条件"
               Object.Tag             =   "重置"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Object.Visible         =   0   'False
               ImageIndex      =   3
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查找"
               Key             =   "查找"
               Object.ToolTipText     =   "查找"
               Object.Tag             =   "查找"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "明细"
               Key             =   "明细"
               Object.ToolTipText     =   "药品明细帐"
               Object.Tag             =   "明细"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "总帐"
               Key             =   "总帐"
               Object.ToolTipText     =   "药品总帐"
               Object.Tag             =   "总帐"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "刷新"
               Key             =   "刷新"
               Object.ToolTipText     =   "刷新"
               Object.Tag             =   "刷新"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "字体"
               Key             =   "字体"
               Object.ToolTipText     =   "字体"
               Object.Tag             =   "字体"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "帮助"
               Description     =   "帮助"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "退出"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList imgTbrHot 
      Left            =   1425
      Top             =   780
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":B4DA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":B6F6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":B912
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":BB2C
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":BD48
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":BF62
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":C17C
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":C398
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":C5B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":C7D0
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":C9EA
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgTbrStard 
      Left            =   690
      Top             =   810
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   11
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":D2C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":D4E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":D6FC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":D916
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":DB32
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":DD4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":DF66
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":E182
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":E39E
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":E5BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmDrugQuery.frx":E7D4
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   1245
      Left            =   4080
      TabIndex        =   12
      Top             =   1440
      Width           =   5055
      _cx             =   8916
      _cy             =   2196
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
      BackColorSel    =   16769992
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   15724527
      GridColor       =   0
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   29
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDrugQuery.frx":F0AE
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
      VirtualData     =   0   'False
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
   Begin VSFlex8Ctl.VSFlexGrid vsfBatch 
      Height          =   1245
      Left            =   3960
      TabIndex        =   13
      Top             =   4080
      Width           =   5055
      _cx             =   8916
      _cy             =   2196
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
      MouseIcon       =   "frmDrugQuery.frx":F4AD
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      BackColorFixed  =   -2147483633
      ForeColorFixed  =   -2147483630
      BackColorSel    =   16053482
      ForeColorSel    =   -2147483640
      BackColorBkg    =   -2147483633
      BackColorAlternate=   15724527
      GridColor       =   0
      GridColorFixed  =   0
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   1
      HighLight       =   1
      AllowSelection  =   -1  'True
      AllowBigSelection=   -1  'True
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   17
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   255
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmDrugQuery.frx":F4C9
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
      VirtualData     =   0   'False
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
   Begin VB.Label lbl剂型_S 
      Alignment       =   2  'Center
      BorderStyle     =   1  'Fixed Single
      Caption         =   "药品剂型"
      Height          =   285
      Left            =   15
      MousePointer    =   7  'Size N S
      TabIndex        =   7
      Top             =   5625
      Width           =   2865
   End
   Begin VB.Label lbl分批_S 
      Alignment       =   2  'Center
      BackColor       =   &H8000000D&
      Caption         =   "分批库存"
      Height          =   180
      Left            =   3600
      MousePointer    =   7  'Size N S
      TabIndex        =   5
      Top             =   3240
      Width           =   6585
   End
   Begin VB.Menu mnuFile 
      Caption         =   "文件(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "打印设置(&S)"
      End
      Begin VB.Menu mnuFilePrintView 
         Caption         =   "打印预览(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "打印(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuExcel 
         Caption         =   "输出到&Excel"
      End
      Begin VB.Menu mnuFileBatch 
         Caption         =   "批量打印明细帐(&B)"
      End
      Begin VB.Menu mnuViewLine 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileOpen 
         Caption         =   "参数设置(&R)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFileBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "报表(&R)"
      Visible         =   0   'False
      Begin VB.Menu mnuReportItem 
         Caption         =   "-"
         Index           =   0
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "查看(&V)"
      Begin VB.Menu mnuViewToolbar 
         Caption         =   "工具栏(&T)"
         Begin VB.Menu mnuViewToolbarStand 
            Caption         =   "标准按钮(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolbarText 
            Caption         =   "文本标签(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFont 
         Caption         =   "字体(&F)"
         Begin VB.Menu mnuViewFontSize 
            Caption         =   "小字体"
            Checked         =   -1  'True
            Index           =   0
         End
         Begin VB.Menu mnuViewFontSize 
            Caption         =   "中字体"
            Index           =   1
         End
         Begin VB.Menu mnuViewFontSize 
            Caption         =   "大字体"
            Index           =   2
         End
      End
      Begin VB.Menu mnuViewForeColor 
         Caption         =   "前景色(&C)"
      End
      Begin VB.Menu mnuViewBackColor 
         Caption         =   "背景色(&B)"
      End
      Begin VB.Menu mnuviewLine2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "刷新(&R)"
         Shortcut        =   {F5}
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "查找(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuviewLineNoVerify 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewNoVerify 
         Caption         =   "未审单据查询(&N)"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "帮助(&H)"
      Begin VB.Menu mnuHelpTitle 
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
      Begin VB.Menu mnuHelpBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
   Begin VB.Menu mnuPopu 
      Caption         =   "弹出"
      Visible         =   0   'False
      Begin VB.Menu mnuPopuOpen 
         Caption         =   "打开(&O)"
      End
      Begin VB.Menu mnuPopuBar0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopuFontSize 
         Caption         =   "小字体"
         Checked         =   -1  'True
         Index           =   0
      End
      Begin VB.Menu mnuPopuFontSize 
         Caption         =   "中字体"
         Index           =   1
      End
      Begin VB.Menu mnuPopuFontSize 
         Caption         =   "大字体"
         Index           =   2
      End
   End
   Begin VB.Menu mnuReportBill 
      Caption         =   "报表菜单"
      Visible         =   0   'False
      Begin VB.Menu mnuBill 
         Caption         =   "单据(&D)"
      End
   End
End
Attribute VB_Name = "frmDrugQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'--------------------------------------------------------
Dim intFont As Integer
'Dim WithEvents DataRecordSet  As adodb.Recordset
Dim appData As Collection
Public intChoose级数 As Byte         '0-售价单位;1-门诊单位;2-药库单位;3-住院单位
Dim bln库存数 As Boolean
Dim bln包含停用药品 As Boolean
Dim intMonths As Integer
Dim BlnHourse As Boolean '为真表示库房
Public BlnDO As Boolean
Dim Bln西成药 As Boolean '表示是否具有查询西成药的权限
Dim Bln中成药 As Boolean '表示是否具有查询中成药的权限
Dim Bln中草药 As Boolean '表示是否具有查询中草药的权限
Dim Str材质 As String
Dim StrSort As String    '表示药品类别
Dim mstrPrivs As String
Private mlngMode As Long
Private mblnViewCost As Boolean       '查看成本价 true-允许查看 false-不允许查看

Private mblnRefresh As Boolean                  '是否正在刷新
Private mstrUnShow_List As String               '不允许显示的列：药品列表
Private mstrUnShow_Batch As String              '不允许显示的列：批次列表

Private LngCardRow As Long
Private LngPhysicRow As Long
Private StrCardSortBy As String                 '排序列
'Modified By 朱玉宝 2003-12-10 地区：泸州 选中颜色改为蓝色，灰色度有调整，替换深灰为蓝色
Private Const glng白色 As Long = &H80000005
Private Const glng黑色 As Long = &H80000008
Private Const glng蓝色 As Long = &HFFCECE
Private Const glng本色 As Long = &H8000000F
Private Const glng灰色 As Long = &HC0C0C0
Private Const glng红色 As Long = &HC0           '停用

Private mStr成本价 As String
Private mStr单价 As String
Private mStr数量 As String
Private mStr金额 As String
Private mStrMax金额 As String

Private mblnExportState As Boolean          '数据输出状态

'从参数表中取药品价格、数量、金额小数位数
Private mintCostDigit As Integer            '成本价小数位数
Private mintPriceDigit As Integer           '售价小数位数
Private mintNumberDigit As Integer          '数量小数位数
Private mintMoneyDigit As Integer           '金额小数位数

'----------------------
'三张报表的变量设置
Public WithEvents ObjReport As zl9Report.clsReport
Attribute ObjReport.VB_VarHelpID = -1
Private lngCurReport As Long
Private CurSheet As Object
Dim strNoS As String
'-----------------------

Private Type Type_SQLCondition
    str通用名 As String
    str编码 As String
    str简码 As String
    str别名 As String
    str规格 As String
    str产地 As String
    str药品信息 As String
    lng药品分类 As Long
    lng库房ID As Long
End Type

Private SQLCondition As Type_SQLCondition

Private Enum IniListType
    AllList = 0
    MainList = 1
    BatchList = 2
End Enum

Private Sub SetSortCode()
    '根据药品编码返回格式化的排序编码
    '编码中可能含有"-"符号，查找所有编码中"-"前最多几位，"-"后最多几位，所有编码都按最大位数进行格式化处理
    Dim lngRow As Long
    Dim int前缀 As Integer
    Dim int后缀 As Integer
    Dim str编码前缀 As String
    Dim str编码后缀 As String
    Dim blnLine As Boolean
    
    With vsfList
        For lngRow = 1 To vsfList.rows - 1
            If InStr(1, .TextMatrix(lngRow, .ColIndex("编码")), "-") > 0 Then
                blnLine = True
                If Len(Mid(.TextMatrix(lngRow, .ColIndex("编码")), 1, InStr(.TextMatrix(lngRow, .ColIndex("编码")), "-") - 1)) > int前缀 Then
                    int前缀 = Len(Mid(.TextMatrix(lngRow, .ColIndex("编码")), 1, InStr(.TextMatrix(lngRow, .ColIndex("编码")), "-") - 1))
                End If
                
                If Len(Mid(.TextMatrix(lngRow, .ColIndex("编码")), InStr(.TextMatrix(lngRow, .ColIndex("编码")), "-") + 1)) > int后缀 Then
                    int后缀 = Len(Mid(.TextMatrix(lngRow, .ColIndex("编码")), InStr(.TextMatrix(lngRow, .ColIndex("编码")), "-") + 1))
                End If
            Else
                If Len(.TextMatrix(lngRow, .ColIndex("编码"))) > int前缀 Then
                    int前缀 = Len(.TextMatrix(lngRow, .ColIndex("编码")))
                End If
            End If
        Next
        
        For lngRow = 1 To .rows - 1
            If blnLine = False Then
                .TextMatrix(lngRow, .ColIndex("排序编码")) = Format(.TextMatrix(lngRow, .ColIndex("编码")), String(int前缀, "0"))
            Else
                If InStr(.TextMatrix(lngRow, .ColIndex("编码")), "-") > 0 Then
                    str编码前缀 = Mid(.TextMatrix(lngRow, .ColIndex("编码")), 1, InStr(.TextMatrix(lngRow, .ColIndex("编码")), "-") - 1)
                    str编码后缀 = Mid(.TextMatrix(lngRow, .ColIndex("编码")), InStr(.TextMatrix(lngRow, .ColIndex("编码")), "-") + 1)
                    
                    str编码前缀 = Format(str编码前缀, String(int前缀, "0"))
                    str编码后缀 = Format(str编码后缀, String(int后缀, "0"))
                Else
                    str编码前缀 = Format(.TextMatrix(lngRow, .ColIndex("编码")), String(int前缀, "0"))
                    str编码后缀 = String(int后缀, "0")
                End If
                
                .TextMatrix(lngRow, .ColIndex("排序编码")) = str编码前缀 & "-" & str编码后缀
            End If
        Next
    End With
End Sub

Private Sub SetDrugDigit(ByVal intUnit As Integer)
    Dim strUnit As String
    Dim intDrugUnit As Integer
    
    Const conInt计算精度 As Integer = 0
    
    Const conInt药品 As Integer = 1
    
    '药品库存查询参数设置的单位，顺序可能和其他模块设置不一致
    Const conint售价单位 As Integer = 1
    Const conint门诊单位 As Integer = 2
    Const conint住院单位 As Integer = 4
    Const conint药库单位 As Integer = 3
        
    Const conInt成本价 As Integer = 1
    Const conInt售价 As Integer = 2
    Const conInt数量 As Integer = 3
    Const conInt金额 As Integer = 4
    
    intDrugUnit = intUnit
    
    Select Case intDrugUnit
        Case conint售价单位            '售价单位：主要是制剂室
            intDrugUnit = 1
        Case conint门诊单位
            intDrugUnit = 2
        Case conint住院单位
            intDrugUnit = 3
        Case conint药库单位
            intDrugUnit = 4
    End Select

    '分别取药品成本价、售价、数量、金额的小数位数
    mintCostDigit = GetDigit(conInt计算精度, conInt药品, conInt成本价, intDrugUnit)
    mintPriceDigit = GetDigit(conInt计算精度, conInt药品, conInt售价, intDrugUnit)
    mintNumberDigit = GetDigit(conInt计算精度, conInt药品, conInt数量, intDrugUnit)
    mintMoneyDigit = GetDigit(conInt计算精度, conInt药品, conInt金额)
    
    mStr成本价 = "####0." & String(mintCostDigit, "0") & ";-####0." & String(mintCostDigit, "0") & "; ;"
    mStr单价 = "####0." & String(mintPriceDigit, "0") & ";-####0." & String(mintPriceDigit, "0") & "; ;"
    mStr数量 = "####0." & String(mintNumberDigit, "0") & ";-####0." & String(mintNumberDigit, "0") & "; ;"
    mStr金额 = "####0." & String(mintMoneyDigit, "0") & ";-####0." & String(mintMoneyDigit, "0") & "; ;"
    
    mStrMax金额 = "####0." & String(5, "0") & ";-####0." & String(5, "0") & "; ;"
End Sub
Private Sub open药品总帐()
    Call ObjReport.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1309_1", Me, "库房=" & Me.cboStock.Text & "|" & IIf(Me.cboStock.ItemData(Me.cboStock.ListIndex) = 0, " is not null ", "=" & Me.cboStock.ItemData(Me.cboStock.ListIndex)))
End Sub

Private Sub open药品明细帐()
'    If DataRecordSet Is Nothing Then Exit Sub
'    If Not (DataRecordSet.State = 1) Then Exit Sub
'    If DataRecordSet.RecordCount = 0 Then Exit Sub
    
    If vsfList.Row = 0 Then Exit Sub
    If vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("名称")) = "" Then Exit Sub
    
    If cboStock.ItemData(cboStock.ListIndex) = 0 Then
        Call ObjReport.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1309_2", Me, "药品=" & vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("名称")) & "|" & Val(vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("药品ID"))), "库房=所有库房|is not null", "单位=" & Choose(intChoose级数, "售价单位", "门诊单位", "药库单位", "住院单位") & "|" & Choose(intChoose级数, 1, 3, 2, 4))    ' , "开始日期=" & Format(DateAdd("m", -1, Sys.Currentdate), "yyyy-MM-DD"), "结束日期=" & Format(Sys.Currentdate, "yyyy-MM-DD"))
    Else
        Call ObjReport.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1309_2", Me, "药品=" & vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("名称")) & "|" & Val(vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("药品ID"))), "库房=" & cboStock.Text & "|=  " & cboStock.ItemData(cboStock.ListIndex), "单位=" & Choose(intChoose级数, "售价单位", "门诊单位", "药库单位", "住院单位") & "|" & Choose(intChoose级数, 1, 3, 2, 4)) ' , "开始日期=" & Format(DateAdd("m", -1, Sys.Currentdate), "yyyy-MM-DD"), "结束日期=" & Format(Sys.Currentdate, "yyyy-MM-DD"))
    End If
End Sub
Private Sub open药品明细表()
    Call ObjReport.ReportOpen(gcnOracle, glngSys, "ZL1_INSIDE_1309_3", Me, "库房=" & Me.cboStock.Text & "|" & IIf(Me.cboStock.ItemData(Me.cboStock.ListIndex) = 0, " is not null ", "=" & Me.cboStock.ItemData(Me.cboStock.ListIndex)), "单位=" & Choose(intChoose级数, "售价单位", "门诊单位", "药库单位", "住院单位") & "|" & Choose(intChoose级数, 1, 3, 2, 4))
End Sub

Private Sub cboStock_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim str工作性质 As String
    
    str工作性质 = "H,I,J,K,L,M,N"

    If KeyCode <> vbKeyReturn Then Exit Sub
    If cboStock.ListCount = 0 Then Call zlControl.ControlSetFocus(vsfList): Exit Sub
    
    If cboStock.ListIndex >= 0 Then
        If Val(cboStock.Tag) = cboStock.ItemData(cboStock.ListIndex) Then
            Call zlControl.ControlSetFocus(vsfList, True)
            Exit Sub
        End If
    End If
    
    If Select部门选择器(Me, cboStock, Trim(cboStock.Text), str工作性质, IIf(zlStr.IsHavePrivs(mstrPrivs, "所有库房"), False, True)) = False Then
        Exit Sub
    End If
    If cboStock.ListIndex >= 0 Then
        cboStock.Tag = cboStock.ItemData(cboStock.ListIndex)
    End If
End Sub

Private Sub cboStock_KeyPress(KeyAscii As Integer)
    '屏蔽输入单引号
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub


Private Sub cboStock_Validate(Cancel As Boolean)
    If cboStock.ListCount > 0 Then
        If cboStock.ListIndex = -1 Then
            MsgBox "请选择一个药库或者药房！", vbInformation, gstrSysName
            Cancel = True
        End If
    End If
End Sub


Private Sub cbrThis_Resize()
    Form_Resize
End Sub

Private Sub cboStock_Click()
    If Me.tvwSection_S.Nodes.count = 0 Then Exit Sub
    Me.tvwSection_S.Tag = ""
    
'    If cboStock.Text = "所有库房" Then
'        vsfList.ColHidden(vsfList.ColIndex("储备情况")) = True
'        vsfBatch.ColHidden(vsfBatch.ColIndex("储备情况")) = False
'    Else
'        vsfList.ColHidden(vsfList.ColIndex("储备情况")) = False
'        vsfBatch.ColHidden(vsfBatch.ColIndex("储备情况")) = True
'    End If
'
'    If bln库存数 = False Then
'        vsfList.ColHidden(vsfList.ColIndex("储备情况")) = True
'        vsfBatch.ColHidden(vsfBatch.ColIndex("储备情况")) = True
'    End If
    
    If Val(cboStock.Tag) <> cboStock.ItemData(cboStock.ListIndex) Then
        If IIf(Val(cboStock.Tag) = 0, 0, 1) <> IIf(cboStock.ItemData(cboStock.ListIndex) = 0, 0, 1) Then
            SaveVsFlexState vsfList, App, Me, IIf(Val(cboStock.Tag) = 0, "所有库房", "")
            SaveVsFlexState vsfBatch, App, Me, IIf(Val(cboStock.Tag) = 0, "所有库房", "")
            
            RestoreVsFlexState vsfList, App, Me, IIf(cboStock.ItemData(cboStock.ListIndex) = 0, "所有库房", "")
            RestoreVsFlexState vsfBatch, App, Me, IIf(cboStock.ItemData(cboStock.ListIndex) = 0, "所有库房", "")
        End If
        
        cboStock.Tag = cboStock.ItemData(cboStock.ListIndex)
        
        ReFreshDrugData Me.cboStock.ItemData(Me.cboStock.ListIndex), IIf(Left(Me.tvwSection_S.SelectedItem.Key, 1) = "R", 0, Mid(Me.tvwSection_S.SelectedItem.Key, 2))
    End If
End Sub

Private Sub Chk剂型_Click()
    Dim lstItem As ListItem
    
    For Each lstItem In Me.lst剂型_S.ListItems
        lstItem.Checked = (Chk剂型.Value = 1)
    Next
    
    DoEvents
    
    ReFreshDrugData Me.cboStock.ItemData(Me.cboStock.ListIndex), IIf(Left(Me.tvwSection_S.SelectedItem.Key, 1) = "R", 0, Mid(Me.tvwSection_S.SelectedItem.Key, 2))
End Sub

Private Sub Form_Activate()
    ReFreshDrugData Me.cboStock.ItemData(Me.cboStock.ListIndex), IIf(Left(Me.tvwSection_S.SelectedItem.Key, 1) = "R", 0, Mid(Me.tvwSection_S.SelectedItem.Key, 2))
End Sub

Private Sub Form_Load()
    intChoose级数 = Val(zlDatabase.GetPara("单位", glngSys, 1309, 3))
    Call SetDrugDigit(intChoose级数)
    
    bln库存数 = (zlDatabase.GetPara("是否显示无库存药品", glngSys, 1309) = 1)
    intMonths = Val(zlDatabase.GetPara("效期报警月数", glngSys, 1309, 3))
    bln包含停用药品 = (zlDatabase.GetPara("是否显示停用药品", glngSys, 1309) = 1)
    intFont = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品库存查询", "字体", 0)

    mlngMode = glngModul
    mstrPrivs = gstrprivs
    gstrStockSearchPrivs = gstrprivs '专门针对库存查询的权限
    mblnViewCost = zlStr.IsHavePrivs(mstrPrivs, "查看成本价")

    Call mnuViewFontSize_Click(intFont)
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngMode, mstrPrivs)
    
    '加载报表
    Set ObjReport = New zl9Report.clsReport
    Call 设置权限
    
    Call SetParent(picFind.hWnd, stbThis.hWnd)
    picFind.Top = 80
    picFind.Left = stbThis.Panels(1).Width + 160
    
    If Not ReFreshTreeView() Then Unload Me: Exit Sub
    
    RestoreWinState Me, App.ProductName, Me.Caption
    RestoreVsFlexState vsfList, App, Me, IIf(cboStock.ItemData(cboStock.ListIndex) = 0, "所有库房", "")
    RestoreVsFlexState vsfBatch, App, Me, IIf(cboStock.ItemData(cboStock.ListIndex) = 0, "所有库房", "")

    Call SetFormat(IniListType.AllList)
    
    stbThis.Panels(2).Picture = picColor
    
'    Set vsfBatch.Icons = imgList.Icons '设置关联的图标控件
End Sub

Private Sub Form_Resize()
    Dim intTop As Integer, intButton As Integer
    If Me.WindowState = 1 Then Exit Sub
    intTop = IIf(Me.cbrThis.Visible, Me.cbrThis.Height, 0)
    intButton = IIf(Me.stbThis.Visible, Me.stbThis.Height, 0)
    
    On Error Resume Next
    Me.picVLine_S.Top = intTop + Me.ScaleTop
    Me.picVLine_S.Height = Me.ScaleHeight - Me.tvwSection_S.Top - intButton
    If Me.picVLine_S.Left < 500 Then Me.picVLine_S.Left = 500
    If Me.picVLine_S.Left > Me.ScaleWidth - 500 Then Me.picVLine_S.Left = Me.ScaleWidth - 500
    
    Me.tvwSection_S.Left = Me.ScaleLeft
    Me.tvwSection_S.Width = Me.picVLine_S.Left - Me.tvwSection_S.Left
    Me.tvwSection_S.Top = Me.ScaleTop + intTop
    Me.tvwSection_S.Height = Me.lbl剂型_S.Top - Me.tvwSection_S.Top
    
    If Me.ScaleWidth - Me.picVLine_S.Left - Me.picVLine_S.Width < 500 Then
        Me.Width = Me.picVLine_S.Left + Me.picVLine_S.Width + 500
    End If
    If Me.ScaleHeight - Me.lbl分批_S.Top - Me.lbl分批_S.Height < 500 Then
        Me.Height = Me.lbl分批_S.Top + Me.lbl分批_S.Height + 2000
    End If
    If Me.ScaleHeight - Me.lbl剂型_S.Top - Me.lbl剂型_S.Height < 500 Then
        Me.Height = Me.lbl剂型_S.Top + Me.lbl剂型_S.Height + 2000
    End If
    Me.lbl剂型_S.Left = Me.tvwSection_S.Left
    Me.lbl剂型_S.Width = Me.tvwSection_S.Width
    Me.Chk剂型.Left = Me.lbl剂型_S.Left + 55
    Me.Chk剂型.Top = Me.lbl剂型_S.Top + 30
    With lst剂型_S
        .Top = Me.lbl剂型_S.Top + Me.lbl剂型_S.Height
        .Height = Me.ScaleHeight - .Top - intButton
        .Width = Me.lbl剂型_S.Width
        .Left = Me.lbl剂型_S.Left
    End With

    Me.lbl分批_S.Left = Me.picVLine_S.Left + Me.picVLine_S.Width - 20
    Me.lbl分批_S.Width = Me.ScaleWidth - Me.lbl分批_S.Left
    With Me.vsfBatch
        .Left = Me.lbl分批_S.Left
        .Width = Me.lbl分批_S.Width
    End With
    
    Me.vsfList.Left = Me.lbl分批_S.Left
    Me.vsfList.Width = Me.lbl分批_S.Width
        
    If Me.vsfBatch.Visible Then
        With Me.vsfBatch
            .Top = Me.lbl分批_S.Top + Me.lbl分批_S.Height
            .Height = Me.ScaleHeight - .Top - intButton
        End With
        Me.vsfList.Top = intTop + 50
        Me.vsfList.Height = Me.lbl分批_S.Top - Me.vsfList.Top
    Else
        Me.vsfList.Top = intTop + 50
        Me.vsfList.Height = Me.ScaleHeight - Me.vsfList.Top - intButton
    End If

'    If cboStock.Text = "所有库房" Then
'        vsfList.ColHidden(vsfList.ColIndex("储备情况")) = True
'        vsfBatch.ColHidden(vsfBatch.ColIndex("储备情况")) = False
'    Else
'        vsfList.ColHidden(vsfList.ColIndex("储备情况")) = False
'        vsfBatch.ColHidden(vsfBatch.ColIndex("储备情况")) = True
'    End If
'    If bln库存数 = False Then
'        vsfList.ColHidden(vsfList.ColIndex("储备情况")) = True
'        vsfBatch.ColHidden(vsfBatch.ColIndex("储备情况")) = True
'    End If
    
    With picColor
        .Top = Me.ScaleHeight - .Height - 30
        .Left = Me.ScaleWidth - stbThis.Panels(3).Width - stbThis.Panels(4).Width - .Width - 300
    End With
End Sub
Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next

    SaveWinState Me, App.ProductName, Me.Caption
    SaveVsFlexState vsfList, App, Me, IIf(Val(cboStock.Tag) = 0, "所有库房", "")
    SaveVsFlexState vsfBatch, App, Me, IIf(Val(cboStock.Tag) = 0, "所有库房", "")
End Sub

Private Sub lbl分批_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Me.lbl分批_S.Top = Me.lbl分批_S.Top + y
        If Me.lbl分批_S.Top < 5000 Then Me.lbl分批_S.Top = 5000
        If Me.Height - Me.lbl分批_S.Top < 2000 Then Me.lbl分批_S.Top = Me.Height - 2000
        Form_Resize
    End If
End Sub

Private Sub lbl剂型_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Me.lbl剂型_S.Top = Me.lbl剂型_S.Top + y
        If Me.lbl剂型_S.Top < 2000 Then Me.lbl剂型_S.Top = 2000
        If Me.Height - Me.lbl剂型_S.Top < 2000 Then Me.lbl剂型_S.Top = Me.Height - 2000
        Form_Resize
    End If
End Sub

Private Sub lst剂型_S_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    ReFreshDrugData Me.cboStock.ItemData(Me.cboStock.ListIndex), IIf(Left(Me.tvwSection_S.SelectedItem.Key, 1) = "R", 0, Mid(Me.tvwSection_S.SelectedItem.Key, 2))
End Sub

Private Sub SaveVsFlexState(objGrid As VSFlexGrid, objApp As App, objForm As Form, Optional strType As String)
    '保存VSFlexGrid控件的列状态：列名，列键值，列状态（0-隐藏;1-显示），列宽，列对齐方式
    '格式：列名1,列键值1,列状态1,列宽1,列对其方式1|列名2,列键值2,列状态2,列宽2,列对其方式2。。。
    'objApp：工程对象
    'objForm：主窗口对象
    'strType：根据业务情况，同一表格可能有多个显示方式，自定义显示方式名称
    Dim strText As String
    Dim i As Integer
    
    If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then
        Exit Sub
    End If
    
    With objGrid
        For i = 0 To .Cols - 1
            strText = IIf(strText = "", "", strText & "|") & .TextMatrix(0, i) & "," & .ColKey(i) & "," & IIf(.colHidden(i) = True, 0, 1) & "," & .ColWidth(i) & "," & .ColAlignment(i)
        Next
    End With
    
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & objApp.ProductName & "\" & objForm.Name & objForm.Caption & "\" & TypeName(objGrid), objGrid.Name & strType, strText
End Sub

Private Sub RestoreVsFlexState(objGrid As VSFlexGrid, objApp As App, objForm As Form, Optional strType As String)
    '恢复VSFlexGrid控件的列状态（同时恢复列顺序）：列名，列键值，列状态（0-隐藏;1-显示），列宽，列对齐方式
    '格式：列名1,列键值1,列状态1,列宽1,列对其方式1|列名2,列键值2,列状态2,列宽2,列对其方式2。。。
    'objApp：工程对象
    'objForm：主窗口对象
    'strType：根据业务情况，同一表格可能有多个显示方式，自定义显示方式名称
    Dim strText As String
    Dim i As Integer
    Dim intCols As Integer
    Dim arrText
    
    Dim strName As String
    Dim strkey As String
    Dim blnHidden As Boolean
    Dim dblWidth As Double
    Dim intAlignment As Integer
    
    Dim blnFindKey As Boolean
    
    If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then
        Exit Sub
    End If
    
    strText = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & objApp.ProductName & "\" & objForm.Name & objForm.Caption & "\" & TypeName(objGrid), objGrid.Name & strType, "")
    
    '注册表值为空，不处理
    If strText = "" Then Exit Sub
    
    arrText = Array()
    arrText = Split(strText, "|")
    
    '列数不相等，退出
    If UBound(arrText) + 1 <> objGrid.Cols Then Exit Sub
    
    '没找到列键值，不处理
    For i = 0 To UBound(arrText)
        strkey = Split(arrText(i), ",")(1)
        blnFindKey = False
        For intCols = 0 To objGrid.Cols - 1
            If strkey = objGrid.ColKey(intCols) Then
                blnFindKey = True
                Exit For
            End If
        Next
        If blnFindKey = False Then
            Exit Sub
        End If
    Next
    
    '恢复列状态
    With objGrid
        .Clear
        For i = 0 To UBound(arrText)
            .TextMatrix(0, i) = Split(arrText(i), ",")(0)
            .ColKey(i) = Split(arrText(i), ",")(1)
            .colHidden(i) = IIf(Val(Split(arrText(i), ",")(2)) = 0, True, False)
            .ColWidth(i) = Val(Split(arrText(i), ",")(3))
            .ColAlignment(i) = Val(Split(arrText(i), ",")(4))

            .FixedAlignment(i) = flexAlignCenterCenter
        Next
    End With
End Sub

Private Sub mnuEXCEL_Click()
    grdPrint 1
End Sub

Private Sub mnuFileBatch_Click()
    With Frm批量打印明细帐
        .Show 1, Me
    End With
End Sub

Private Sub mnufileexit_Click()
    Unload Me
End Sub

Private Sub mnuFileOpen_Click()
    BlnDO = False
    frmDrugQueryParaSet.In_权限 = mstrPrivs
    frmDrugQueryParaSet.Show 1, Me
    
    If Not BlnDO Then Exit Sub
    
    intChoose级数 = Val(zlDatabase.GetPara("单位", glngSys, 1309, 3))
    Call SetDrugDigit(intChoose级数)
    
    bln库存数 = (zlDatabase.GetPara("是否显示无库存药品", glngSys, 1309) = 1)
    intMonths = Val(zlDatabase.GetPara("效期报警月数", glngSys, 1309))
    bln包含停用药品 = (zlDatabase.GetPara("是否显示停用药品", glngSys, 1309) = 1)
    
    If Me.tvwSection_S.Nodes.count = 0 Then Exit Sub
    Me.tvwSection_S.Tag = ""
    ReFreshDrugData Me.cboStock.ItemData(Me.cboStock.ListIndex), IIf(Left(Me.tvwSection_S.SelectedItem.Key, 1) = "R", 0, Mid(Me.tvwSection_S.SelectedItem.Key, 2))
End Sub

Private Sub mnuFilePrint_Click()
    grdPrint 3
End Sub

Private Sub mnuFilePrintSet_Click()
     zlPrintSet
End Sub

Private Sub mnuFilePrintView_Click()
  grdPrint 0
End Sub
Private Sub grdPrint(blnIsPreview As Byte)
    '---------------------------------------------------
    '功能：    根据屏幕组织表上附加项目，打印预览
    '参数：
    '     blnIsPreview: 0表示预览 1表示输出到EXCEL 其它表示打印
    '返回：
    '---------------------------------------------------
    Dim objPrint As New zlPrint1Grd
    Dim objRow As zlTabAppRow
    
    objPrint.Title.Text = "药品库存查询"
    Set objRow = New zlTabAppRow
    objRow.Add "库房：" & Me.cboStock.Text
    objRow.Add "药品用途：" & Me.tvwSection_S.SelectedItem.Text
    objRow.Add "截止日期：" & Format(Sys.Currentdate, "yyyy年MM月DD日")
    objPrint.UnderAppRows.Add objRow
    Set objRow = New zlTabAppRow
    objRow.Add "打印人:" & UserInfo.用户姓名
    objRow.Add "打印时间:" & Format(Sys.Currentdate, "yyyy年MM月DD日 HH:MM")
    objPrint.BelowAppRows.Add objRow
    Set objPrint.Body = vsfList
    
    mblnExportState = True
    
    If blnIsPreview = 0 Then
         zlPrintOrView1Grd objPrint, 2
    Else
      If blnIsPreview = 1 Then
            zlPrintOrView1Grd objPrint, 3
      Else
        Select Case zlPrintAsk(objPrint)
            Case 1
                 zlPrintOrView1Grd objPrint, 1
            Case 2
                zlPrintOrView1Grd objPrint, 2
            Case 3
                zlPrintOrView1Grd objPrint, 3
        End Select
      End If
    End If
    Set objPrint = Nothing
    mblnExportState = False
End Sub

Private Sub mnuHelpAbout_Click()
   ShowAbout Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision
End Sub

Private Sub mnuHelpTitle_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub mnuReportItem_Click(Index As Integer)
    '默认参数：库房=库房id，分类=分类id，药品=药品id
    Dim strReportName As String
    
    strReportName = Split(mnuReportItem(Index).Tag, ",")(1)
    
    Select Case strReportName
        Case "ZL1_INSIDE_1309_2"
            Call open药品明细帐
        Case "ZL1_INSIDE_1309_3"
            Call open药品明细表
        Case "ZL1_INSIDE_1309_1"
            Call open药品总帐
        Case Else
            Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, _
                "药品=", _
                "库房=" & IIf(Val(cboStock.ItemData(cboStock.ListIndex)) = 0, "", Val(cboStock.ItemData(cboStock.ListIndex))), _
                "分类=" & IIf(SQLCondition.lng药品分类 = 0, "", SQLCondition.lng药品分类))
    End Select
End Sub

Private Sub mnuViewFind_Click()
    Dim strFind As String
    Me.tvwSection_S.Tag = ""
    strFind = Frm库存查找.GetSearch(Me, _
         SQLCondition.str通用名, _
         SQLCondition.str编码, _
         SQLCondition.str简码, _
         SQLCondition.str别名, _
         SQLCondition.str规格, _
         SQLCondition.str产地)
    
    If strFind = "" Then Exit Sub
    If Not ReFreshDrugData(cboStock.ItemData(cboStock.ListIndex), IIf(Left(Me.tvwSection_S.SelectedItem.Key, 1) = "R", 0, Mid(Me.tvwSection_S.SelectedItem.Key, 2)), strFind, False) Then Exit Sub
    Me.tvwSection_S.Tag = "T"
End Sub

Private Sub mnuViewFontSize_Click(Index As Integer)
    Dim i As Integer
    For i = 0 To 2
        Me.mnuViewFontSize(i).Checked = False
    Next
    Me.mnuViewFontSize(Index).Checked = True

    Select Case Index
    Case 0
        Me.vsfList.Font.Size = 9
        Me.tvwSection_S.Font.Size = 9
        vsfBatch.Font.Size = 9
     Case 1
        Me.vsfList.Font.Size = 11
        Me.tvwSection_S.Font.Size = 11
        vsfBatch.Font.Size = 11
    Case 2
        Me.vsfList.Font.Size = 15
        Me.tvwSection_S.Font.Size = 15
        vsfBatch.Font.Size = 15
    End Select
    intFont = Index
    
    SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品库存查询", "字体", intFont
    
    Form_Resize
    Me.Refresh
End Sub

Private Sub mnuViewForeColor_Click()
    Dim lngForeColor As Long
    lngForeColor = zlGetColor(Me.vsfList.ForeColor)
    Me.vsfList.ForeColor = lngForeColor
    
End Sub

Private Sub mnuViewBackColor_Click()
    Dim lngBackColor As Long
    lngBackColor = zlGetColor(Me.vsfList.BackColor)
    Me.vsfList.BackColor = lngBackColor
End Sub


Private Sub mnuViewNoVerify_Click()
    frmQueryUnVerify.ShowCard Me, cboStock, intChoose级数, mintNumberDigit
End Sub

Private Sub mnuViewRefresh_Click()
    ReFreshDrugData Me.cboStock.ItemData(Me.cboStock.ListIndex), IIf(Left(Me.tvwSection_S.SelectedItem.Key, 1) = "R", 0, Mid(Me.tvwSection_S.SelectedItem.Key, 2))
End Sub



Private Sub mnuViewStatus_Click()
    Me.mnuViewStatus.Checked = Not Me.mnuViewStatus.Checked
    Me.stbThis.Visible = Me.mnuViewStatus.Checked
    Form_Resize
End Sub

Private Sub mnuViewToolbarStAnd_Click()
    Dim intCount As Integer
    Me.mnuViewToolbarStand.Checked = Not Me.mnuViewToolbarStand.Checked
    Me.mnuViewToolbarText.Enabled = Me.mnuViewToolbarStand.Checked
    Me.cbrThis.Visible = Me.mnuViewToolbarStand.Checked
    If Me.mnuViewToolbarText.Checked Then
        For intCount = 1 To Me.tbrThis.Buttons.count
            Me.tbrThis.Buttons(intCount).Caption = Me.tbrThis.Buttons(intCount).Tag
        Next
    Else
        For intCount = 1 To Me.tbrThis.Buttons.count
            Me.tbrThis.Buttons(intCount).Caption = ""
        Next
    End If
    Me.cbrThis.Bands(1).MinHeight = Me.tbrThis.Height
    Me.cbrThis.Refresh
    Form_Resize

End Sub
Private Sub mnuViewToolbarText_Click()
    Dim intCount As Integer
    Me.mnuViewToolbarText.Checked = Not Me.mnuViewToolbarText.Checked
    If Me.mnuViewToolbarText.Checked Then
        For intCount = 1 To Me.tbrThis.Buttons.count
            Me.tbrThis.Buttons(intCount).Caption = Me.tbrThis.Buttons(intCount).Tag
        Next
    Else
        For intCount = 1 To Me.tbrThis.Buttons.count
            Me.tbrThis.Buttons(intCount).Caption = ""
        Next
    End If
    Me.cbrThis.Bands(1).MinHeight = Me.tbrThis.Height
    Me.cbrThis.Refresh
    Form_Resize
End Sub

Private Sub vsfBatch_BeforeMoveColumn(ByVal Col As Long, Position As Long)
    '重设列选择列表
    Call InitColSelList(IniListType.BatchList, vsfBatch)
End Sub

Private Sub vsfBatch_Click()
    vsfBatch.BackColorSel = glngRowByFocus
'    vsfBatch.GridColorFixed = &H80000008
'    vsfBatch.GridColor = &H80000008
    
    vsfList.BackColorSel = glngRowByNotFocus
'    vsfList.GridColorFixed = &H80000010
'    vsfList.GridColor = &H80000010
End Sub

Private Sub vsfBatch_EnterCell()
    If mblnRefresh = True Then Exit Sub
    If vsfBatch.Row = 0 Then Exit Sub

    With vsfBatch
        '当前行设置
        .Redraw = flexRDNone

        .ForeColorSel = .Cell(flexcpForeColor, .Row, 1)

        .Redraw = flexRDDirect
    End With
End Sub

Private Sub vsfBatch_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Long
    
    If Button = 2 Then '列选择器
        If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then Exit Sub
        
        If vsfBatch.MouseRow <> 0 Then Exit Sub
        
        InitColSelList IniListType.BatchList, vsfBatch
        
        '根据当前状态直接确定勾选状态
        With vsfColSel
            If .Visible Then
                .Visible = False
                vsfBatch.SetFocus
            Else
                For i = .FixedRows To .rows - 1
                    If vsfBatch.colHidden(.RowData(i)) Or vsfBatch.ColWidth(.RowData(i)) = 0 Then
                        .TextMatrix(i, 0) = 0
                    Else
                        .TextMatrix(i, 0) = 1
                    End If
                Next
                
                .Height = .RowHeightMin * .rows + 150
                .Top = vsfBatch.Top + vsfBatch.RowHeight(0)
                .Width = 1750
'                If .Top + .Height > Me.ScaleHeight - vsfBatch.Top Then
'                    .Height = Me.ScaleHeight - .Top - vsfBatch.Top
'                    .Width = 1750
'                Else
'                    .Width = 1470
'                End If
                
                .Left = vsfBatch.Left + x
                .ZOrder
                .Visible = True
                .SetFocus
            End If
        End With
    End If
End Sub

Private Sub vsfColSel_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngCol As Long
    
    lngCol = vsfColSel.RowData(Row)
    If Val(vsfColSel.TextMatrix(Row, 0)) <> 0 Then
        If Val(vsfColSel.Tag) = IniListType.MainList Then
            vsfList.colHidden(lngCol) = False
        Else
            vsfBatch.colHidden(lngCol) = False
        End If
    Else
        If Val(vsfColSel.Tag) = IniListType.MainList Then
            vsfList.colHidden(lngCol) = True
        Else
            vsfBatch.colHidden(lngCol) = True
        End If
    End If
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


Private Sub InitColSelList(ByVal intListType As Integer, ByVal objGrid As VSFlexGrid)
    Dim i As Integer
    Dim strUnShow As String
    
    If intListType = IniListType.MainList Then
        strUnShow = mstrUnShow_List
    ElseIf intListType = IniListType.BatchList Then
        strUnShow = mstrUnShow_Batch
    End If
        
    With vsfColSel
        .Tag = intListType
        
        .rows = .FixedRows
        For i = 0 To objGrid.Cols - 1
            '不在不允许显示列表的列才能加入列选择列表
            If InStr(1, ";" & strUnShow & ";", ";" & objGrid.ColKey(i) & ";") = 0 Then
                .rows = .rows + 1
                .TextMatrix(.rows - 1, 1) = objGrid.TextMatrix(0, i)
                .RowData(.rows - 1) = i
                
                '列宽为空或者隐藏的列设置为不勾选
                If Not (objGrid.ColWidth(i) = 0 Or objGrid.colHidden(i)) Then
                    .TextMatrix(.rows - 1, 0) = 0
                End If
                
'                '指定的列设置为不能设置隐藏
'                If IsInString(mstrUnallowSetColHide, objGrid.ColKey(i), ";") = True Then
'                    .Cell(flexcpForeColor, .Rows - 1, 1) = .BackColorFixed
'                End If
            End If
        Next
    End With
End Sub

Private Sub vsfList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    '重设列选择列表
    Call InitColSelList(IniListType.MainList, vsfList)
End Sub

Private Sub vsfList_AfterSort(ByVal Col As Long, Order As Integer)
    With vsfList
        If Col = .ColIndex("编码") Then
            .Col = .ColIndex("排序编码")
            .Sort = Order
        End If
    End With
End Sub


Private Sub vsfList_Click()
    vsfList.BackColorSel = glngRowByFocus
'    vsfList.GridColorFixed = &H80000008
'    vsfList.GridColor = &H80000008
    
    vsfBatch.BackColorSel = glngRowByNotFocus
    vsfBatch.ForeColorSel = IIf(Val(vsfBatch.RowData(vsfBatch.Row)) = 0, glng黑色, glng报警)
'    vsfBatch.GridColorFixed = &H80000010
'    vsfBatch.GridColor = &H80000010
End Sub

Private Sub vsfList_DblClick()
'    If DataRecordSet Is Nothing Then Exit Sub
'    If Not (DataRecordSet.State = 1) Then Exit Sub
'    If DataRecordSet.RecordCount = 0 Then Exit Sub
    
    If vsfList.MouseRow = 0 Or vsfList.MouseRow = -1 Then Exit Sub
    
    If vsfList.TextMatrix(vsfList.MouseRow, vsfList.ColIndex("名称")) = "" Then Exit Sub
    
    If cboStock.ItemData(cboStock.ListIndex) = 0 Then
        Call ObjReport.ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_INSIDE_1309_2", "ZL8_INSIDE_1309_2"), Me, "药品=" & vsfList.TextMatrix(vsfList.MouseRow, vsfList.ColIndex("名称")) & "|" & Val(vsfList.TextMatrix(vsfList.MouseRow, vsfList.ColIndex("药品ID"))), "库房=所有库房|is not null", "单位=" & Choose(intChoose级数, "售价单位", "门诊单位", "药库单位", "住院单位") & "|" & Choose(intChoose级数, 1, 3, 2, 4))    ' , "开始日期=" & Format(DateAdd("m", -1, Sys.Currentdate), "yyyy-MM-DD"), "结束日期=" & Format(Sys.Currentdate, "yyyy-MM-DD"))
    Else
        Call ObjReport.ReportOpen(gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "ZL1_INSIDE_1309_2", "ZL8_INSIDE_1309_2"), Me, "药品=" & vsfList.TextMatrix(vsfList.MouseRow, vsfList.ColIndex("名称")) & "|" & Val(vsfList.TextMatrix(vsfList.MouseRow, vsfList.ColIndex("药品ID"))), "库房=" & cboStock.Text & "|=  " & cboStock.ItemData(cboStock.ListIndex), "单位=" & Choose(intChoose级数, "售价单位", "门诊单位", "药库单位", "住院单位") & "|" & Choose(intChoose级数, 1, 3, 2, 4))  ' , "开始日期=" & Format(DateAdd("m", -1, Sys.Currentdate), "yyyy-MM-DD"), "结束日期=" & Format(Sys.Currentdate, "yyyy-MM-DD"))
    End If
End Sub

Private Sub vsfList_EnterCell()
    On Error Resume Next
    
    If mblnExportState = True Then Exit Sub
    If mblnRefresh = True Then Exit Sub
    If vsfList.Row = 0 Then Exit Sub
    If vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("名称")) = "" Then
        RefreshBatch Me.cboStock.ItemData(Me.cboStock.ListIndex), 0
        Me.vsfBatch.Visible = False
        Me.lbl分批_S.Visible = False
        Me.vsfBatch.rows = 2
        Me.vsfBatch.Redraw = flexRDDirect
        Call Form_Resize
        Exit Sub
    End If
    
    With vsfList
        '当前行设置
        .Redraw = flexRDNone
        
        '正常药品和停用药品的前景色
        .ForeColorSel = IIf(Trim(.TextMatrix(.Row, vsfList.ColIndex("撤档时间"))) = "", glng黑色, glng红色)
        
        .Redraw = flexRDDirect
    
        '提取批次信息
        RefreshBatch Me.cboStock.ItemData(Me.cboStock.ListIndex), Val(.TextMatrix(.Row, .ColIndex("药品ID")))
        
        If Me.tvwSection_S.Tag <> "T" Then Exit Sub
        
        Err = 0
       
        Me.tvwSection_S.Nodes("_" & vsfList.TextMatrix(.Row, .ColIndex("用途分类ID"))).Selected = True
        Me.tvwSection_S.Nodes("_" & vsfList.TextMatrix(.Row, .ColIndex("用途分类ID"))).Expanded = True
    End With
End Sub

Private Sub picVLine_S_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Me.picVLine_S.Left = Me.picVLine_S.Left + x
        Form_Resize
    End If
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    With Button
        Select Case .Key
        Case "预览"
            mnuFilePrintView_Click
        Case "打印"
            grdPrint 3
        Case "总帐"
            Call open药品总帐
        Case "明细"
            Call open药品明细帐
        Case "查找"
            mnuViewFind_Click
        Case "刷新"
            mnuViewRefresh_Click
        Case "字体"
             PopupMenu mnuViewFont
        Case "前景色"
            mnuViewForeColor_Click
        Case "背景色" '
            mnuViewBackColor_Click
        Case "帮助"
            mnuHelpTitle_Click
        Case "退出"
           mnufileexit_Click
        End Select
    End With
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mnuViewToolbar
    End If
End Sub

Private Sub tvwSection_S_GotFocus()
    If Me.tvwSection_S.Tag = "T" Then Me.tvwSection_S.Tag = "F"
End Sub

Private Sub tvwSection_S_NodeClick(ByVal Node As MSComctlLib.Node)
    If Me.tvwSection_S.Tag = "T" Then Exit Sub
    ReFreshDrugData Me.cboStock.ItemData(Me.cboStock.ListIndex), IIf(Left(Me.tvwSection_S.SelectedItem.Key, 1) = "R", 0, Mid(Me.tvwSection_S.SelectedItem.Key, 2))
End Sub

Private Function ReFreshTreeView() As Boolean
    '-------------------------------------------------------------------------
    '--功能:重新获取的树型结构数据
    '--参数:
    '--返回:如果数据库打开成功,则返True,否则返回False
    '-------------------------------------------------------------------------
    Dim objNode As Node
    Dim RecDept As New ADODB.Recordset
    Dim RecDrug As New ADODB.Recordset
    Dim Str材质 As String
    Dim i As Integer
    Dim RsTreeRecordset As ADODB.Recordset
    
    ReFreshTreeView = False
    
    On Error GoTo ErrHand
    
    gstrSQL = "Select distinct a.ID,(a.编码 || '-' || a.名称) As 名称 From 部门表 a,部门性质说明 b,部门性质分类 C " & _
              "Where (a.站点 = [2] Or a.站点 is Null) And a.id=b.部门id And b.工作性质=c.名称 And (c.编码 in ('H','I','J','K','L','M','N'))" & _
              IIf(zlStr.IsHavePrivs(mstrPrivs, "所有库房"), "", " And A.id In (Select 部门ID From 部门人员 Where 人员ID=[1])") & _
              "  and (to_char(a.撤档时间,'yyyy-mm-dd')='3000-01-01' or a.撤档时间 is null) " & _
              "Order By A.编码 || '-' || A.名称 "
    Set RecDept = zlDatabase.OpenSQLRecord(gstrSQL, "取所有库房", UserInfo.用户ID, gstrNodeNo)
    
    With RecDept
        If .RecordCount = 0 Then
            MsgBox "药库体系未建立或权限不足，不能执行本程序!", vbInformation, gstrSysName
            Exit Function
        End If
        If zlStr.IsHavePrivs(mstrPrivs, "所有库房") Then
            Me.cboStock.Clear
            Me.cboStock.AddItem "所有库房"
            Me.cboStock.ItemData(Me.cboStock.NewIndex) = 0
            Me.cboStock.ListIndex = Me.cboStock.NewIndex
        End If
        Do While Not .EOF
            Me.cboStock.AddItem .Fields("名称").Value
            Me.cboStock.ItemData(Me.cboStock.NewIndex) = .Fields("ID").Value
            .MoveNext
        Loop
        
        gstrSQL = "Select Distinct 部门id From 部门人员 Where 缺省 = 1 And 人员id = [1]"
        Set RecDept = zlDatabase.OpenSQLRecord(gstrSQL, "取缺省部门", UserInfo.用户ID)
        
        Me.cboStock.ListIndex = 0
        
        '定位到缺省部门
        If Not RecDept.EOF Then
            For i = 0 To Me.cboStock.ListCount - 1
                If Me.cboStock.ItemData(i) = RecDept!部门ID Then
                    Me.cboStock.ListIndex = i
                    Exit For
                End If
            Next
        End If
        
        Me.cboStock.Tag = Me.cboStock.ItemData(Me.cboStock.ListIndex)
    End With
    
    Str材质 = ""
    StrSort = ""
    If zlStr.IsHavePrivs(mstrPrivs, "西成药") Then
        Bln西成药 = True
        Str材质 = "1"
        StrSort = ",'5'"
    Else
        Bln西成药 = False
    End If
    
    If zlStr.IsHavePrivs(mstrPrivs, "中成药") Then
        Bln中成药 = True
        If Str材质 = "" Then
            Str材质 = "2"
        Else
            Str材质 = Str材质 & ",2"
        End If
        StrSort = StrSort & ",'6'"
    Else
        Bln中成药 = False
    End If
    
    If zlStr.IsHavePrivs(mstrPrivs, "中草药") Then
        Bln中草药 = True
        If Str材质 = "" Then
            Str材质 = "3"
        Else
            Str材质 = Str材质 & ",3"
        End If
        StrSort = StrSort & ",'7'"
    Else
        Bln中草药 = False
    End If
    
    If Str材质 = "" Then
        MsgBox "对不起，必须得有一个管理药品材质的权限，请与系统管理员联系！", vbOKOnly, gstrSysName
        Exit Function
    End If
    
    gstrSQL = "Select a.id,a.上级id,a.名称,Decode(a.类型,1,'5',2,'6','7') As 材质 " & _
              "From 诊疗分类目录 a " & _
              "Where a.类型 in (select * from Table(Cast(f_Num2List([1]) As zlTools.t_NumList))) " & _
              "Start with a.上级id is null " & _
              "Connect by prior a.id=a.上级id " & _
              "Order by level,a.id"
    Set RsTreeRecordset = zlDatabase.OpenSQLRecord(gstrSQL, "药品用途分类", Str材质)
    
    With RsTreeRecordset
        If .RecordCount = 0 Then
            MsgBox "药品用途体系未建立，不能执行本程序!", vbInformation, gstrSysName
            Exit Function
        End If
        Me.tvwSection_S.Nodes.Clear
        
        If Bln西成药 = True Then
            Me.tvwSection_S.Nodes.Add , , "R" & "5", "西成药", "child"
        End If
        
        If Bln中草药 = True Then
            Me.tvwSection_S.Nodes.Add , , "R" & "7", "中草药", "child"
        End If
        
        If Bln中成药 = True Then
            Me.tvwSection_S.Nodes.Add , , "R" & "6", "中成药", "child"
        End If
        
        Do While Not .EOF
            If IsNull(.Fields("上级id").Value) Then
                Set objNode = Me.tvwSection_S.Nodes.Add("R" & !材质, 4, "_" & .Fields("id").Value, .Fields("名称").Value, "child")
            Else
                Set objNode = Me.tvwSection_S.Nodes.Add("_" & .Fields("上级id").Value, 4, "_" & .Fields("id").Value, .Fields("名称").Value, "child")
            End If
            .MoveNext
         Loop
         
         If tvwSection_S.Nodes(1).Children <> 0 Then
            tvwSection_S.Nodes(1).Child.Selected = True
         Else
            tvwSection_S.Nodes(1).Selected = True
         End If
    End With
    With RecDrug
        gstrSQL = " Select 编码,名称 From 药品剂型 "
        Call zlDatabase.OpenRecordset(RecDrug, gstrSQL, "药品剂型")
        
        If .RecordCount = 0 Then
            MsgBox "药品剂型未建立,不能执行程序!", vbInformation, gstrSysName
            Exit Function
        End If
        Me.lst剂型_S.ListItems.Clear
        Do While Not .EOF
            Me.lst剂型_S.ListItems.Add , "K" & !编码, !名称
            Me.lst剂型_S.ListItems("K" & !编码).Checked = True
            .MoveNext
        Loop
        Me.lst剂型_S.ListItems(1).Selected = True
    End With
    ReFreshTreeView = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then Resume
    Unload Me
End Function

Private Function ReFreshDrugData(ByVal lngDeptId As Long, _
    Optional ByVal lngUseId As Long = 0, Optional ByVal strFind As String = "", Optional ByVal Click As Boolean = True) As Boolean
    '-------------------------------------------------------------------------
    '--功能:重新获取的药品库存数
    '--参数:
    '       lngDeptId:药品房id
    '       lngUseId:用途id值
    '       strFind:用于快速查找（输入编码、名称或简码）
    '       Click为真表示点击选择,为假表示查找
    '--返回:
    '-------------------------------------------------------------------------
    Dim strOrder As String, strSql As String, str收费项目目录 As String
    Dim str剂型 As String
    Dim lstItem As ListItem
    Dim blnAllCheck As Boolean
    Dim rsData As ADODB.Recordset
    Dim strSql药名 As String
    
    blnAllCheck = True
    str剂型 = ""
    strOrder = ""
    If strFind = "" Then
        str收费项目目录 = " 收费项目目录 "
    Else
        str收费项目目录 = "(Select distinct A.ID, A.编码, A.名称, A.规格, A.产地, A.是否变价, A.撤档时间, A.类别, A.计算单位 " & _
                 " From 收费项目目录 A,收费项目别名 B " & _
                 " Where A.ID=B.收费细目ID " & strFind & ")"
    End If
    
    Call FS.ShowFlash("正在查找数据,请稍候 ...", Me)
    DoEvents
    For Each lstItem In Me.lst剂型_S.ListItems
        If lstItem.Checked Then
            str剂型 = str剂型 & "," & lstItem
        Else
            blnAllCheck = False
        End If
    Next
    If str剂型 <> "" Then
        If blnAllCheck = True Then
            str剂型 = ""
        Else
            str剂型 = Mid(str剂型, 2) & ",方剂"
        End If
    Else
        str剂型 = "小宝"
    End If
    
    If lngDeptId = 0 Then
        Select Case intChoose级数
        Case 1
            gstrSQL = ",A.计算单位 as 单位,'' as 上次采购价,Decode(Nvl(a.是否变价, 0), 0, Nvl(p.现价, 0), Decode(Nvl(b.零售价, 0), 0, Nvl(p.现价, 0), b.零售价)) as 当前售价,nvl(M.剂量系数,0) as 系数1,Sum(B.可用数量) As 可用数量,Sum(B.实际数量) As 实际数量,Sum(B.实际金额) As 实际金额,Sum(B.实际差价) As 实际差价,sum(b.平均成本价) 平均成本价,Decode(To_Char(A.撤档时间,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(A.撤档时间,'yyyy-MM-dd')) 撤档时间,1 as 除数, Nvl(A.是否变价, 0) 变价, G.名称 As 上次供应商, '' 库房货位 "
            strOrder = " Group by M.药品ID,X.分类id,A.编码,M.基本药物,M.标识码,A.名称,L.名称,A.规格,decode(b.上次产地,null,decode(m.上次产地,null,a.产地,m.上次产地),b.上次产地),m.原产地,M.药库分批,A.计算单位,Decode(Nvl(a.是否变价, 0), 0, Nvl(p.现价, 0), Decode(Nvl(b.零售价, 0), 0, Nvl(p.现价, 0), b.零售价)) ,nvl(M.剂量系数,0),Decode(To_Char(A.撤档时间,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(A.撤档时间,'yyyy-MM-dd')),Nvl(A.是否变价, 0),G.名称"
        Case 2
            gstrSQL = ",M.门诊单位 as 单位,'' as 上次采购价,Decode(Nvl(a.是否变价, 0), 0, Nvl(p.现价, 0), Decode(Nvl(b.零售价, 0), 0, Nvl(p.现价, 0), b.零售价)) * Nvl(m.门诊包装, 0) as 当前售价,nvl(M.门诊包装,0) as 系数1,Sum(B.可用数量/Decode(M.门诊包装,0,1,null,1,M.门诊包装)) as 可用数量,Sum(B.实际数量/Decode(M.门诊包装,0,1,null,1,M.门诊包装)) as 实际数量,Sum(B.实际金额) As 实际金额,Sum(B.实际差价) As 实际差价,sum(b.平均成本价)*Decode(M.门诊包装,0,1,null,1,M.门诊包装) 平均成本价,Decode(To_Char(A.撤档时间,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(A.撤档时间,'yyyy-MM-dd')) 撤档时间,Decode(M.门诊包装,0,1,null,1,M.门诊包装) as 除数, Nvl(A.是否变价, 0) 变价, G.名称 As 上次供应商, '' 库房货位 "
            strOrder = " Group by M.药品ID,X.分类id,A.编码,M.基本药物,M.标识码,A.名称,L.名称,A.规格,decode(b.上次产地,null,decode(m.上次产地,null,a.产地,m.上次产地),b.上次产地),m.原产地,M.药库分批,M.门诊单位,Decode(Nvl(a.是否变价, 0), 0, Nvl(p.现价, 0), Decode(Nvl(b.零售价, 0), 0, Nvl(p.现价, 0), b.零售价)) * Nvl(m.门诊包装, 0),nvl(M.门诊包装,0),Decode(M.门诊包装,0,1,null,1,M.门诊包装),Decode(To_Char(A.撤档时间,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(A.撤档时间,'yyyy-MM-dd')),Nvl(A.是否变价, 0),G.名称"
        Case 3
            gstrSQL = ",M.药库单位 as 单位,'' as 上次采购价,Decode(Nvl(a.是否变价, 0), 0, Nvl(p.现价, 0), Decode(Nvl(b.零售价, 0), 0, Nvl(p.现价, 0), b.零售价)) * Nvl(m.药库包装, 0) as 当前售价,nvl(M.药库包装,0) as 系数1,Sum(B.可用数量/Decode(M.药库包装,0,1,null,1,M.药库包装)) as 可用数量, Sum(B.实际数量/Decode(M.药库包装,0,1,null,1,M.药库包装)) as 实际数量,Sum(B.实际金额) As 实际金额,Sum(B.实际差价) As 实际差价,sum(b.平均成本价)*Decode(M.药库包装,0,1,null,1,M.药库包装) 平均成本价,Decode(To_Char(A.撤档时间,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(A.撤档时间,'yyyy-MM-dd')) 撤档时间,Decode(M.药库包装,0,1,null,1,M.药库包装) as 除数, Nvl(A.是否变价, 0) 变价, G.名称 As 上次供应商,'' 库房货位 "
            strOrder = " Group by M.药品ID,X.分类id,A.编码,M.基本药物,M.标识码,A.名称,L.名称,A.规格,decode(b.上次产地,null,decode(m.上次产地,null,a.产地,m.上次产地),b.上次产地),m.原产地,M.药库分批,M.药库单位,Decode(Nvl(a.是否变价, 0), 0, Nvl(p.现价, 0), Decode(Nvl(b.零售价, 0), 0, Nvl(p.现价, 0), b.零售价)) * Nvl(m.药库包装, 0),nvl(M.药库包装,0),Decode(M.药库包装,0,1,null,1,M.药库包装),Decode(To_Char(A.撤档时间,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(A.撤档时间,'yyyy-MM-dd')),Nvl(A.是否变价, 0),G.名称"
        Case 4
            gstrSQL = ",M.住院单位 as 单位,'' as 上次采购价,Decode(Nvl(a.是否变价, 0), 0, Nvl(p.现价, 0), Decode(Nvl(b.零售价, 0), 0, Nvl(p.现价, 0), b.零售价)) * Nvl(m.住院包装, 0) as 当前售价,nvl(M.住院包装,0) as 系数1,Sum(B.可用数量/Decode(M.住院包装,0,1,null,1,M.住院包装)) as 可用数量, Sum(B.实际数量/Decode(M.住院包装,0,1,null,1,M.住院包装)) as 实际数量,Sum(B.实际金额) As 实际金额,Sum(B.实际差价) As 实际差价,sum(b.平均成本价)*Decode(M.住院包装,0,1,null,1,M.住院包装) 平均成本价,Decode(To_Char(A.撤档时间,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(A.撤档时间,'yyyy-MM-dd')) 撤档时间,Decode(M.住院包装,0,1,null,1,M.住院包装) as 除数, Nvl(A.是否变价, 0) 变价, G.名称 As 上次供应商,'' 库房货位 "
            strOrder = " Group by M.药品ID,X.分类id,A.编码,M.基本药物,M.标识码,A.名称,L.名称,A.规格,decode(b.上次产地,null,decode(m.上次产地,null,a.产地,m.上次产地),b.上次产地),m.原产地,M.药库分批,M.住院单位,Decode(Nvl(a.是否变价, 0), 0, Nvl(p.现价, 0), Decode(Nvl(b.零售价, 0), 0, Nvl(p.现价, 0), b.零售价)) * Nvl(m.住院包装, 0),nvl(M.住院包装,0),Decode(M.住院包装,0,1,null,1,M.住院包装),Decode(To_Char(A.撤档时间,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(A.撤档时间,'yyyy-MM-dd')),Nvl(A.是否变价, 0),G.名称"
        End Select
    Else
        Select Case intChoose级数
        Case 1
            gstrSQL = ",A.计算单位 as 单位,Nvl(Avg(S.上次采购价), m.成本价) as 上次采购价,Decode(Nvl(a.是否变价, 0), 0, Nvl(p.现价, 0), Decode(Nvl(s.零售价, 0), 0, Nvl(p.现价, 0), s.零售价)) As 当前售价,nvl(M.剂量系数,0) as 系数1,Sum(S.可用数量) as 可用数量, Sum(S.实际数量) as 实际数量,sum(i.下限) as 下限,Sum(S.实际金额) as 实际金额,Sum(S.实际差价) as 实际差价,sum(s.平均成本价) 平均成本价,Decode(To_Char(A.撤档时间,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(A.撤档时间,'yyyy-MM-dd')) 撤档时间,1 as 除数, Nvl(A.是否变价, 0) 变价, G.名称 As 上次供应商, C.库房货位 "
            strOrder = " Group by M.药品ID,A.编码,M.基本药物,M.标识码,X.分类id,A.名称,L.名称,A.规格,decode(s.上次产地,null,decode(m.上次产地,null,a.产地,m.上次产地),s.上次产地),Nvl(m.原产地, s.原产地),nvl(M.最大效期,0),s.库存效期,s.报警,M.药库分批,A.计算单位,Decode(Nvl(a.是否变价, 0), 0, Nvl(p.现价, 0), Decode(Nvl(s.零售价, 0), 0, Nvl(p.现价, 0), s.零售价)),nvl(M.剂量系数,0),Decode(To_Char(A.撤档时间,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(A.撤档时间,'yyyy-MM-dd')),Nvl(A.是否变价, 0),G.名称, C.库房货位, m.成本价"
        Case 2
            gstrSQL = ",M.门诊单位 as 单位,Nvl(Avg(S.上次采购价*nvl(M.门诊包装,0)), m.成本价*nvl(M.门诊包装,0)) as 上次采购价,Decode(Nvl(a.是否变价, 0), 0, Nvl(p.现价, 0), Decode(Nvl(s.零售价, 0), 0, Nvl(p.现价, 0), s.零售价)) * Nvl(m.门诊包装, 0) As 当前售价,nvl(M.门诊包装,0) as 系数1,Sum(S.可用数量 /Decode(M.门诊包装,0,1,null,1,M.门诊包装)) as 可用数量, Sum(S.实际数量 /Decode(M.门诊包装,0,1,null,1,M.门诊包装)) as 实际数量,sum(i.下限/Decode(M.门诊包装,0,1,null,1,M.门诊包装)) as 下限,Sum(S.实际金额) as 实际金额,Sum(S.实际差价) as 实际差价,sum(s.平均成本价)*Decode(M.门诊包装,0,1,null,1,M.门诊包装) 平均成本价,Decode(To_Char(A.撤档时间,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(A.撤档时间,'yyyy-MM-dd')) 撤档时间,Decode(M.门诊包装,0,1,null,1,M.门诊包装) as 除数, Nvl(A.是否变价, 0) 变价, G.名称 As 上次供应商, C.库房货位 "
            strOrder = " Group by M.药品ID,A.编码,M.基本药物,M.标识码,X.分类id,A.名称,L.名称,A.规格,decode(s.上次产地,null,decode(m.上次产地,null,a.产地,m.上次产地),s.上次产地),Nvl(m.原产地, s.原产地),nvl(M.最大效期,0),s.库存效期,s.报警,M.药库分批,M.门诊单位,Decode(Nvl(a.是否变价, 0), 0, Nvl(p.现价, 0), Decode(Nvl(s.零售价, 0), 0, Nvl(p.现价, 0), s.零售价)) * Nvl(m.门诊包装, 0),nvl(M.门诊包装,0),Decode(M.门诊包装,0,1,null,1,M.门诊包装),Decode(To_Char(A.撤档时间,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(A.撤档时间,'yyyy-MM-dd')),Nvl(A.是否变价, 0),G.名称, C.库房货位, m.成本价"
        Case 3
            gstrSQL = ",M.药库单位 as 单位,Nvl(Avg(S.上次采购价*nvl(M.药库包装,0)), m.成本价*nvl(M.药库包装,0)) as 上次采购价,Decode(Nvl(a.是否变价, 0), 0, Nvl(p.现价, 0), Decode(Nvl(s.零售价, 0), 0, Nvl(p.现价, 0), s.零售价)) * Nvl(m.药库包装, 0) As 当前售价,nvl(M.药库包装,0) as 系数1,Sum(S.可用数量 /Decode(M.药库包装,0,1,null,1,M.药库包装)) as 可用数量,Sum(S.实际数量 /Decode(M.药库包装,0,1,null,1,M.药库包装)) as 实际数量,sum(i.下限/Decode(M.药库包装,0,1,null,1,M.药库包装)) as 下限,Sum(S.实际金额) as 实际金额,Sum(S.实际差价) as 实际差价,sum(s.平均成本价)*Decode(M.药库包装,0,1,null,1,M.药库包装) 平均成本价,Decode(To_Char(A.撤档时间,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(A.撤档时间,'yyyy-MM-dd')) 撤档时间,Decode(M.药库包装,0,1,null,1,M.药库包装) as 除数, Nvl(A.是否变价, 0) 变价, G.名称 As 上次供应商, C.库房货位 "
            strOrder = " Group by M.药品ID,A.编码,M.基本药物,M.标识码,X.分类id,A.名称,L.名称,A.规格,decode(s.上次产地,null,decode(m.上次产地,null,a.产地,m.上次产地),s.上次产地),Nvl(m.原产地, s.原产地),nvl(M.最大效期,0),s.库存效期,s.报警,M.药库分批,M.药库单位,Decode(Nvl(a.是否变价, 0), 0, Nvl(p.现价, 0), Decode(Nvl(s.零售价, 0), 0, Nvl(p.现价, 0), s.零售价)) * Nvl(m.药库包装, 0),nvl(M.药库包装,0),Decode(M.药库包装,0,1,null,1,M.药库包装),Decode(To_Char(A.撤档时间,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(A.撤档时间,'yyyy-MM-dd')),Nvl(A.是否变价, 0),G.名称, C.库房货位, m.成本价"
        Case 4
            gstrSQL = ",M.住院单位 as 单位,Nvl(Avg(S.上次采购价*nvl(M.住院包装,0)), m.成本价*nvl(M.住院包装,0)) as 上次采购价,Decode(Nvl(a.是否变价, 0), 0, Nvl(p.现价, 0), Decode(Nvl(s.零售价, 0), 0, Nvl(p.现价, 0), s.零售价)) * Nvl(m.住院包装, 0) As 当前售价,nvl(M.住院包装,0) as 系数1,Sum(S.可用数量 /Decode(M.住院包装,0,1,null,1,M.住院包装)) as 可用数量, Sum(S.实际数量 /Decode(M.住院包装,0,1,null,1,M.住院包装)) as 实际数量,sum(i.下限/Decode(M.住院包装,0,1,null,1,M.住院包装)) as 下限,Sum(S.实际金额) as 实际金额,Sum(S.实际差价) as 实际差价,sum(s.平均成本价)*Decode(M.住院包装,0,1,null,1,M.住院包装) 平均成本价,Decode(To_Char(A.撤档时间,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(A.撤档时间,'yyyy-MM-dd')) 撤档时间,Decode(M.住院包装,0,1,null,1,M.住院包装) as 除数, Nvl(A.是否变价, 0) 变价, G.名称 As 上次供应商, C.库房货位 "
            strOrder = " Group by M.药品ID,A.编码,M.基本药物,M.标识码,X.分类id,A.名称,L.名称,A.规格,decode(s.上次产地,null,decode(m.上次产地,null,a.产地,m.上次产地),s.上次产地),Nvl(m.原产地, s.原产地),nvl(M.最大效期,0),s.库存效期,s.报警,M.药库分批,M.住院单位,Decode(Nvl(a.是否变价, 0), 0, Nvl(p.现价, 0), Decode(Nvl(s.零售价, 0), 0, Nvl(p.现价, 0), s.零售价)) * Nvl(m.住院包装, 0),nvl(M.住院包装,0),Decode(M.住院包装,0,1,null,1,M.住院包装),Decode(To_Char(A.撤档时间,'yyyy-MM-dd'),'3000-01-01',' ',To_Char(A.撤档时间,'yyyy-MM-dd')),Nvl(A.是否变价, 0),G.名称, C.库房货位, m.成本价"
        End Select
    End If
    
    On Error GoTo ErrHand:

    If lngDeptId = 0 Then
        strSql = "SELECT Distinct M.药品ID,X.分类ID As 用途分类ID,A.编码,M.基本药物,M.标识码 As 药卡号,A.名称 As 通用名,L.名称 As 商品名,A.规格,decode(b.上次产地,null,decode(m.上次产地,null,a.产地,m.上次产地),b.上次产地) AS 产地,m.原产地 as 原产地, NULL AS 效期,DECODE(M.药库分批,1,'是','否') AS 药库分批 " & gstrSQL & _
                " FROM 药品规格 M,收费价目 P," & str收费项目目录 & " A "
                
        strSql = strSql & " ,(Select a.药品id, Avg(a.上次采购价) As 上次成本价,Sum(a.实际数量 * a.平均成本价) / Decode(Sum(Nvl(a.实际数量, 0)), 0, 1, Sum(Nvl(a.实际数量, 0))) as 平均成本价," & _
                " Sum(a.实际数量 * a.零售价) / Decode(Sum(Nvl(a.实际数量, 0)), 0, 1, Sum(Nvl(a.实际数量, 0))) as 零售价, '' 上次产地, Max(nvl(a.批次,0)) As 批次, Sum(a.可用数量) As 可用数量," & _
                " Sum(a.实际数量) As 实际数量, Sum(a.实际金额) As 实际金额, Sum(a.实际差价) As 实际差价 " & _
                " From 药品库存 A, 药品规格 B, 诊疗项目目录 C, 收费项目目录 D " & _
                " Where a.药品id = b.药品id And b.药名id = c.Id And b.药品id = d.Id And 性质 = 1 and (Nvl(a.可用数量,0)<>0 or Nvl(a.实际数量,0)<>0 or Nvl(a.实际金额,0)<>0 or Nvl(a.实际差价,0)<>0) "
        If Click Then
            strSql = strSql & (IIf(lngUseId = 0, " AND D.类别=[10]", _
                 " AND C.分类ID IN ( SELECT ID FROM 诊疗分类目录 Q Where Q.类型 In (1,2,3) START WITH Q.ID= [8] CONNECT BY PRIOR ID=上级ID)"))
        End If
        strSql = strSql & " Group By a.药品id) B "
        
        strSql = strSql & " ,诊疗项目目录 X,药品特性 T,收费项目别名 L, 供应商 G " & _
                " WHERE M.药名ID=X.ID And X.ID=T.药名ID And Nvl(M.上次供应商id, 0) = G.ID(+) " & _
                " AND M.药品ID=P.收费细目ID AND SYSDATE BETWEEN P.执行日期 AND NVL(P.终止日期,SYSDATE) " & _
                GetPriceClassString("P") & _
                IIf(bln包含停用药品, "", " AND (TO_CHAR(A.撤档时间, 'YYYY-MM-DD') = '3000-01-01' OR A.撤档时间 IS NULL) ") & _
                " AND A.ID=M.药品ID AND M.药品ID=B.药品ID(+)  " & _
                " And M.药品ID=L.收费细目ID(+) And L.性质(+)=3 And L.码类(+)=1"
        If Not Click Then
            strSql = strSql & " And A.类别 in (" & Mid(StrSort, 2) & ")"
        Else
            strSql = strSql & (IIf(lngUseId = 0, " AND A.类别=[10]", _
                 " AND X.分类ID IN ( SELECT ID FROM 诊疗分类目录 Q Where Q.类型 In (1,2,3) START WITH Q.ID= [8] CONNECT BY PRIOR ID=上级ID)"))
        End If
        
        If str剂型 <> "" Then
'            StrSql = StrSql & " And T.药品剂型=E.Column_Value "
            strSql = strSql & " And Instr(',' || [11] || ',' , T.药品剂型)>0 "
        End If

        strSql = strSql + strOrder
    Else
        strSql = "SELECT M.药品ID,X.分类ID As 用途分类ID,A.编码,M.基本药物,M.标识码 As 药卡号,A.名称 As 通用名,L.名称 As 商品名,A.规格,decode(s.上次产地,null,decode(m.上次产地,null,a.产地,m.上次产地),s.上次产地) AS 产地,Nvl(m.原产地, s.原产地) as 原产地,NVL(M.最大效期,0) AS 效期,s.报警,s.库存效期,DECODE(M.药库分批,1,'是','否') AS 药库分批 " & gstrSQL & _
                " FROM 药品规格 M,收费价目 P,(select nvl(下限,0) as 下限,药品id from 药品储备限额 where 库房id=[9]) I ," & str收费项目目录 & _
                " A,诊疗项目目录 X,药品特性 T "
        strSql = strSql & " ,(Select b.药品id, b.上次采购价, b.平均成本价,b.零售价,a.上次产地,a.原产地, b.可用数量,b.实际数量,b.实际金额, b.实际差价, A.上次供应商id,Decode(Sign(Add_Months(Sysdate, " & intMonths & ") - 效期), -1, 0, 1) 报警,效期 as 库存效期 " & _
                "  From 药品库存 a,(SELECT a.药品ID,avg(a.上次采购价) AS 上次采购价,Sum(a.实际数量 * a.平均成本价) / Decode(Sum(Nvl(a.实际数量, 0)), 0, 1, Sum(Nvl(a.实际数量, 0))) 平均成本价," & _
                " Sum(a.实际数量 * a.零售价) / Decode(Sum(Nvl(a.实际数量, 0)), 0, 1, Sum(Nvl(a.实际数量, 0))) 零售价, " & _
                " Max(nvl(a.批次,0)) AS 批次,SUM(a.可用数量) AS 可用数量,SUM(a.实际数量) AS 实际数量,SUM(a.实际金额) AS 实际金额,SUM(a.实际差价) AS 实际差价 " & _
                "  FROM 药品库存 A, 药品规格 B, 诊疗项目目录 C, 收费项目目录 D " & _
                " WHERE a.药品id = b.药品id And b.药名id = c.Id And b.药品id = d.Id And a.库房ID=[9] AND a.性质=1 " & _
                " and (Nvl(a.可用数量,0)<>0 or Nvl(a.实际数量,0)<>0 or Nvl(a.实际金额,0)<>0 or Nvl(a.实际差价,0)<>0)  "
        If Click Then
            strSql = strSql & (IIf(lngUseId = 0, " AND D.类别=[10]", _
                 " AND C.分类ID IN ( SELECT ID FROM 诊疗分类目录 Q Where Q.类型 In (1,2,3) START WITH Q.ID= [8] CONNECT BY PRIOR ID=上级ID)"))
        End If
        strSql = strSql & "  GROUP BY a.药品ID) b Where a.库房ID=[9] and a.药品id=b.药品id And a.性质 = 1 And nvl(a.批次,0) = b.批次) S, "
        
        strSql = strSql & " 收费项目别名 L, 供应商 G, (Select Distinct 收费细目id, 执行科室id From 收费执行科室 Where 执行科室id=[9]) K, " & _
                " (Select 库房id, 药品id, 库房货位 From 药品储备限额 Where 库房id = [9] And 库房货位 Is Not Null) C "
        
        strSql = strSql & " WHERE  i.药品id(+)=m.药品id and M.药品ID =A.ID And M.药名ID =X.ID And X.ID=T.药名ID And Nvl(S.上次供应商id, 0) = G.ID(+) " & _
                " And M.药品ID=L.收费细目ID(+) And L.性质(+)=3 AND L.码类(+)=1 And M.药品id = C.药品id(+) " & _
                " AND M.药品ID=S.药品ID(+) And M.药品ID = K.收费细目id " & _
                IIf(bln包含停用药品, "", " AND (TO_CHAR(A.撤档时间, 'YYYY-MM-DD') = '3000-01-01' OR A.撤档时间 IS NULL)  ") & _
                "       And M.药品ID+0=P.收费细目ID AND SYSDATE BETWEEN P.执行日期 AND NVL(P.终止日期,SYSDATE) " & _
                GetPriceClassString("P")
        
        If Not Click Then
            strSql = strSql & " And A.类别 in (" & Mid(StrSort, 2) & ")"
        Else
            strSql = strSql & (IIf(lngUseId = 0, " AND A.类别=[10]", _
                 " AND X.分类ID IN ( SELECT ID FROM 诊疗分类目录 Q Where Q.类型 In (1,2,3) START WITH Q.ID= [8] CONNECT BY PRIOR ID=上级ID)"))
        End If
        
        If str剂型 <> "" Then
            strSql = strSql & " And T.药品剂型 in (select * from Table(Cast(f_Str2list([11]) As zlTools.t_Strlist))) "
        End If
        
        strSql = strSql + strOrder
    End If
    gstrSQL = "Select * From (" & strSql & ")" & IIf(bln库存数, " Where NVL(实际数量,0)<>0 ", "")
    gstrSQL = gstrSQL & " Order By 编码"
    
    SQLCondition.lng药品分类 = lngUseId
    SQLCondition.lng库房ID = lngDeptId
    
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, _
            SQLCondition.str通用名, _
            SQLCondition.str编码, _
            SQLCondition.str简码, _
            SQLCondition.str别名, _
            SQLCondition.str规格, _
            SQLCondition.str产地, _
            SQLCondition.str药品信息, _
            SQLCondition.lng药品分类, _
            SQLCondition.lng库房ID, _
            Mid(tvwSection_S.SelectedItem.Key, 2), _
            str剂型)
    
    With rsData
        If .RecordCount = 0 Then
            mnuExcel.Enabled = False
            mnuFilePrint.Enabled = False
            mnuFilePrintView.Enabled = False
            mnuViewFind.Enabled = False
'            mnuViewList.Enabled = False
            tbrThis.Buttons.Item(1).Enabled = False
            tbrThis.Buttons.Item(2).Enabled = False
            tbrThis.Buttons.Item(6).Enabled = False
            tbrThis.Buttons.Item(7).Enabled = False
        Else
            mnuExcel.Enabled = True
            mnuFilePrint.Enabled = True
            mnuFilePrintView.Enabled = True
            mnuViewFind.Enabled = True
'            mnuViewList.Enabled = True
            tbrThis.Buttons.Item(1).Enabled = True
            tbrThis.Buttons.Item(2).Enabled = True
            tbrThis.Buttons.Item(6).Enabled = True
            tbrThis.Buttons.Item(7).Enabled = True
        End If
    End With
        
    Call FS.StopFlash
    Call SetFormat(IniListType.MainList)
    If Not rsData.EOF Then DataBound rsData
    
    With vsfList
        .Row = 1
        Call vsfList_EnterCell
    End With
    
    ReFreshDrugData = (rsData.RecordCount <> 0)
    If ReFreshDrugData Then Me.vsfList.SetFocus
    Exit Function
ErrHand:
    Call FS.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Exit Function
End Function

Private Sub RefreshBatch(lng库房ID As Long, lng药品id As Long)
    '-------------------------------------------------------------------------
    '--功能:重新获取的药品分批库存数
    '--参数:
    '       lng库房Id:药品房id
    '       lng药品Id:用途id值
    '--返回:
    '-------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim intRow As Long
    Dim intCol As Long
    Dim lngColor As Long
    
    Dim int分批 As Integer
    Dim int药房 As Integer
    Dim lng下限 As Long
    Dim strTemp As String
    Dim dbl包装系数 As Double
    Dim Dbl数量 As Double
    
    On Error GoTo ErrHand
    
    mblnRefresh = True
            
    Me.vsfBatch.Redraw = flexRDNone
    Me.vsfBatch.rows = 1

    gstrSQL = "Select 1 From 部门性质说明 Where 部门id=[1] And 工作性质 IN ('中药房','西药房','成药房')"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[药房判断]", lng库房ID)
    
    If rsTemp.EOF Then
        int药房 = 0
    Else
        int药房 = 1
    End If
    
    If lng药品id = 0 Then Exit Sub
    
    dbl包装系数 = vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("包装"))
    
    gstrSQL = " Select Decode(nvl(药库分批,0),1,Decode(Nvl(药房分批,0),1,2,1),0) As 分批 " & _
              " From 药品规格 Where 药品id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取分批性质", lng药品id)
        
    '如果药库分批且药房分批（int分批=2）；仅药库分批（int分批=1）；不分批（int分批=0）
    If Not rsTemp.EOF And Not rsTemp.BOF Then
        int分批 = rsTemp!分批
        If lng库房ID = 0 Or (int药房 = 1 And int分批 = 2) Or (int药房 = 0 And int分批 <> 0) Then
            '是所有库房 或者 是药库且库存分批，则显示分库房分批库存
            If lng库房ID = 0 Then
                gstrSQL = "Select 库房, 批号, 平均成本价, 失效期, 报警, 产地, 原产地,上次采购价, Sum(可用数量) As 可用数量, Sum(实际数量) As 实际数量, Sum(实际金额) As 实际金额, Sum(实际差价) As 实际差价," & _
                          "     填制日期 , NO, 供药单位, 供应商, 库房货位, 下限, 售价, 是否变价, 现价 " & _
                        " From ( " & _
                        " SELECT (D.编码 || '-' || D.名称) AS 库房,s.上次批号 AS 批号,avg(s.平均成本价) 平均成本价,s.效期 AS 失效期,DECODE(Nvl(SIGN(ADD_MONTHS(SYSDATE," & intMonths & ")-S.效期),-1),-1,0,1) 报警,NULL AS 产地,Null As 原产地,NULL AS 上次采购价," & _
                        "        SUM(S.可用数量)/" & dbl包装系数 & " AS 可用数量,SUM(S.实际数量)/" & dbl包装系数 & " AS 实际数量,SUM(S.实际金额) AS 实际金额,SUM(S.实际差价) AS 实际差价," & _
                        "        NULL AS 填制日期,NULL AS NO,NULL AS 供药单位,NULL As 供应商, C.库房货位,下限/" & dbl包装系数 & " as 下限, decode(nvl(s.零售价,0),0,decode(sum(s.实际数量),0,0,Sum(s.实际金额) / Sum(s.实际数量)),s.零售价)*" & dbl包装系数 & " As 售价,nvl(a.是否变价,0) as 是否变价,b.现价*" & dbl包装系数 & "as 现价 " & _
                        " FROM 药品库存 S,部门表 D, (Select Distinct 收费细目id, 执行科室id From 收费执行科室) K, 药品储备限额 C,收费项目目录 A,收费价目 B " & _
                        " WHERE S.库房ID=D.ID AND S.性质=1 AND S.药品ID=[1] And S.库房id = C.库房id(+) And S.药品id = C.药品id(+) " & _
                        "       And K.执行科室id(+) = S.库房ID And K.收费细目id(+) = S.药品ID AND s.药品id=a.id and a.Id = b.收费细目id And Sysdate Between 执行日期 And 终止日期 " & _
                        GetPriceClassString("B") & " AND (Nvl(S.实际数量,0)<>0 OR Nvl(S.实际金额,0)<>0 OR Nvl(S.实际差价,0)<>0) " & _
                        " GROUP BY D.编码 || '-' || D.名称,s.上次批号, C.库房货位,下限,是否变价,现价,批次,零售价,s.效期)" & _
                        " Group By 批号,库房, 平均成本价, 失效期, 报警, 产地, 上次采购价, 填制日期, NO, 供药单位, 供应商, 库房货位, 下限, 售价, 是否变价, 现价" & _
                        " order by 库房"
            Else
               gstrSQL = "SELECT (D.编码 || '-' || D.名称) AS 库房,S.上次批号 As 批号,s.平均成本价, S.效期 失效期, S.上次产地 As 产地,Decode(s.原产地, Null, a.原产地, s.原产地) As 原产地,DECODE(Nvl(SIGN(ADD_MONTHS(SYSDATE," & intMonths & ")-S.效期),-1),-1,0,1) 报警," & _
                        "        S.可用数量/" & dbl包装系数 & " AS 可用数量,S.实际数量/" & dbl包装系数 & " AS 实际数量,S.实际金额,S.实际差价," & _
                        "        S.上次采购价*" & dbl包装系数 & " AS 上次采购价, G.名称 As 供应商,'' 库房货位, decode(nvl(s.零售价,0),0,decode(s.实际数量,0,0,s.实际金额 / s.实际数量),s.零售价)*" & dbl包装系数 & " As 售价,nvl(a.是否变价,0) as 是否变价, b.现价*" & dbl包装系数 & "as 现价" & _
                        " FROM 药品库存 S,部门表 D,药品规格 A, 供应商 G, (Select Distinct 收费细目id, 执行科室id From 收费执行科室) K,收费项目目录 A, 收费价目 B " & _
                        " WHERE S.库房ID=D.ID AND A.药品ID=S.药品ID" & _
                        "       AND S.药品ID=[1] AND S.性质=1 AND S.库房ID=[2] And Nvl(S.上次供应商id, 0) = G.ID(+) " & _
                        "       And K.执行科室id(+) = S.库房ID And K.收费细目id(+) = S.药品ID and s.药品id=a.id and a.Id = b.收费细目id And Sysdate Between 执行日期 And 终止日期 " & _
                        GetPriceClassString("B") & " AND (Nvl(S.实际数量,0)<>0 OR Nvl(S.实际金额,0)<>0 OR Nvl(S.实际差价,0)<>0) " & _
                        " ORDER BY D.编码 || '-' || D.名称,S.上次批号"
            End If
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng药品id, lng库房ID)
            
            Call SetFormat(IniListType.BatchList)
            
            Me.vsfBatch.rows = 2
            With rsTemp
                Do While Not .EOF
                    If lng库房ID = 0 Then
                        Dbl数量 = 0
                        
                        If !库房 <> vsfBatch.TextMatrix(vsfBatch.rows - 2, vsfBatch.ColIndex("库房")) And vsfBatch.rows - 2 <> 0 Then
                            For intRow = 1 To vsfBatch.rows - 1
                                If vsfBatch.TextMatrix(intRow, vsfBatch.ColIndex("库房")) = vsfBatch.TextMatrix(vsfBatch.rows - 2, vsfBatch.ColIndex("库房")) Then
                                    Dbl数量 = Dbl数量 + Val(vsfBatch.TextMatrix(intRow, vsfBatch.ColIndex("实际数量")))
                                End If
                            Next
                            
                            vsfBatch.MergeCells = flexMergeRestrictRows
                            vsfBatch.MergeRow(vsfBatch.rows - 1) = True
                            
                            For intCol = 0 To vsfBatch.Cols - 1
                                vsfBatch.TextMatrix(vsfBatch.rows - 1, intCol) = vsfBatch.TextMatrix(vsfBatch.rows - 2, vsfBatch.ColIndex("库房")) & "实际数量为：" & Dbl数量 & vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("单位"))
                            Next
                            
                            Dbl数量 = 0
                            vsfBatch.rows = vsfBatch.rows + 1
                        End If
                    End If
                    
                    Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("库房")) = !库房
                    Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("批号")) = IIf(IsNull(!批号), "", !批号)
                    Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("效期")) = Format(!失效期, "yyyy年MM月dd日")
                    If gtype_UserSysParms.P149_效期显示方式 = 1 And Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("效期")) <> "" And cboStock.Text <> "所有库房" Then
                        '换算为有效期
                        Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("效期")) = Format(DateAdd("D", -1, Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("效期"))), "yyyy-mm-dd")
                    End If
                    
                    Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("产地")) = IIf(IsNull(!产地), "", !产地)
                    Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("原产地")) = IIf(IsNull(!原产地), "", !原产地)
                    Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("可用数量")) = Format(!可用数量, mStr数量)
                    Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("实际数量")) = Format(!实际数量, mStr数量)
                    Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("实际金额")) = Format(!实际金额, mStr金额)
                    Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("实际差价")) = Format(!实际差价, mStr金额)
                    If !是否变价 = 0 Then
                        Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("售价")) = Format(!现价, mStr单价)
                    Else
                        Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("售价")) = Format(!售价, mStr单价)
                    End If
                    Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("上次采购价")) = Format(!上次采购价, mStr成本价)
                    If !实际数量 <> 0 Then
'                        Me.vsfBatch.TextMatrix(vsfbatch.rows-1, vsfBatch.ColIndex("成本价")) = Format((!实际金额 - !实际差价) / !实际数量, mStr数量)
'                        Me.vsfBatch.TextMatrix(vsfbatch.rows-1, vsfBatch.ColIndex("成本金额")) = Format(!实际金额 - !实际差价, mStr金额)
                        Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("成本价")) = Format(!平均成本价 * dbl包装系数, mStr成本价)
                        Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("成本金额")) = Format(!平均成本价 * dbl包装系数 * !实际数量, mStr金额)
                    End If
          
                    Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("供应商")) = IIf(IsNull(!供应商), "", !供应商)
                    Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("库房货位")) = IIf(IsNull(!库房货位), "", !库房货位)
                    
                    With vsfBatch
                        If Me.vsfBatch.TextMatrix(vsfBatch.rows - 1, vsfBatch.ColIndex("实际数量")) <> "" And lng库房ID = 0 Then
                            lng下限 = IIf(IsNull(rsTemp!下限), 0, rsTemp!下限)
                            If lng下限 = 0 Then
                                strTemp = "无下限"
                            ElseIf Val(.TextMatrix(vsfBatch.rows - 1, .ColIndex("实际数量"))) > lng下限 Then
                                strTemp = "充足"
                            ElseIf Val(.TextMatrix(vsfBatch.rows - 1, .ColIndex("实际数量"))) = lng下限 Then
                                strTemp = "持平"
                            ElseIf Val(.TextMatrix(vsfBatch.rows - 1, .ColIndex("实际数量"))) < lng下限 Then
                                strTemp = "不足"
                            End If
                            .TextMatrix(vsfBatch.rows - 1, .ColIndex("储备情况")) = strTemp
                        End If
                    End With
                    '根据记录状态的不同，进行着色
                    lngColor = IIf(!报警 = 0, glng正常, glng报警)
                    Me.vsfBatch.Cell(flexcpForeColor, vsfBatch.rows - 1, 0, vsfBatch.rows - 1, vsfBatch.Cols - 1) = lngColor
                    '失效期药品在列前加上时钟图标
                    If !报警 = 1 And IsNull(!失效期) = False Then
                        Me.vsfBatch.Cell(flexcpPicture, vsfBatch.rows - 1, 0, vsfBatch.rows - 1, 0) = imglvw.ListImages(3).Picture
                    End If
                    
                    Me.vsfBatch.RowData(vsfBatch.rows - 1) = 0 ' CStr(!报警)
                    
                    '实际数量，金额，差价为0，可用数量不为0的表示是预减可用数量数据，红色字体显示
                    If zlCommFun.NVL(!实际数量, 0) = 0 And zlCommFun.NVL(!实际金额, 0) = 0 And zlCommFun.NVL(!实际差价, 0) = 0 Then
                        Me.vsfBatch.Cell(flexcpForeColor, vsfBatch.rows - 1, 0, vsfBatch.rows - 1, vsfBatch.Cols - 1) = vbRed
                    End If
                    
                    '特殊数据用红色字体显示
                    '1.可用数量<0的表示是可用数量预减或被占用(如先填出库单，再盘亏审核)
                    '2.实际数量<=0，这种可能是没有进行库存检查，或金额存在误差，或错误数据
                    If zlCommFun.NVL(!可用数量, 0) < 0 Or zlCommFun.NVL(!实际数量, 0) <= 0 Then
                        Me.vsfBatch.Cell(flexcpForeColor, vsfBatch.rows - 1, 0, vsfBatch.rows - 1, vsfBatch.Cols - 1) = vbRed
                    End If
                    
                    .MoveNext
                    vsfBatch.rows = vsfBatch.rows + 1
                Loop
                
                If lng库房ID = 0 Then
                    Dbl数量 = 0
                    vsfBatch.MergeCells = flexMergeRestrictRows
                    vsfBatch.MergeRow(vsfBatch.rows - 1) = True
                    
                    For intRow = 1 To vsfBatch.rows - 1
                        If vsfBatch.TextMatrix(intRow, vsfBatch.ColIndex("库房")) = vsfBatch.TextMatrix(vsfBatch.rows - 2, vsfBatch.ColIndex("库房")) Then
                            Dbl数量 = Dbl数量 + Val(vsfBatch.TextMatrix(intRow, vsfBatch.ColIndex("实际数量")))
                        End If
                    Next
                    For intCol = 0 To vsfBatch.Cols - 1
                        vsfBatch.TextMatrix(vsfBatch.rows - 1, intCol) = vsfBatch.TextMatrix(vsfBatch.rows - 2, vsfBatch.ColIndex("库房")) & "实际数量为：" & Dbl数量 & vsfList.TextMatrix(vsfList.Row, vsfList.ColIndex("单位"))
                    Next
                Else
                    vsfBatch.RemoveItem vsfBatch.rows - 1
                End If
            End With
        End If
    End If
    
    If Me.vsfBatch.rows = 1 Then
        Me.vsfBatch.Visible = False
        Me.lbl分批_S.Visible = False
        Me.vsfBatch.rows = 2
    Else
        Me.vsfBatch.Visible = True
        Me.lbl分批_S.Visible = True
    End If
'    Me.vsfBatch.FixedRows = 1
'    Me.vsfBatch.Row = 1
    Me.vsfBatch.Redraw = flexRDDirect
    Call Form_Resize
    
    mblnRefresh = False
    
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    mblnRefresh = False
    Exit Sub
End Sub
Private Sub mnuBill_Click()
    Dim strNo As String
    Dim byt单据 As Integer
    Dim byt记录状态 As Integer
    
    Select Case Mid(strNoS, 4)
        Case "_INSIDE_1309_1"  '总帐
            strNo = Mid(Trim(CurSheet.TextMatrix(CurSheet.Row, 3)), 3)
            byt单据 = Val(CurSheet.TextMatrix(CurSheet.Row, 1))
            byt记录状态 = Val(CurSheet.TextMatrix(CurSheet.Row, 11))
        Case "_INSIDE_1309_2"  '明细帐
            strNo = Trim(CurSheet.TextMatrix(CurSheet.Row, 3))
            byt单据 = Val(CurSheet.TextMatrix(CurSheet.Row, 2))
            byt记录状态 = Val(CurSheet.TextMatrix(CurSheet.Row, 1))
        Case "_INSIDE_1309_3"  '明细表
        
    End Select
    
    If strNo = "" Or byt单据 = 0 Or byt记录状态 = 99 Then Exit Sub
    If byt单据 = 0 Then Exit Sub
    ShowBill Me, strNo, byt记录状态, byt单据
End Sub

Private Sub ObjReport_ReportActive(ByVal strNo As String, Form As Object)
    lngCurReport = Form.hWnd
    strNoS = strNo
End Sub

Private Sub ObjReport_SheetDblClick(ByVal strNo As String, Sheet As Object, frmParent As Object)
    lngCurReport = frmParent.hWnd
    strNoS = strNo
    Set CurSheet = Sheet
    If Mid(UCase(strNo), 4) = "_INSIDE_1309_3" Then Exit Sub
    mnuBill_Click
End Sub

Private Sub ObjReport_SheetMouseDown(ByVal strNo As String, Button As Integer, Shift As Integer, x As Single, y As Single, Sheet As Object, frmParent As Object)
    lngCurReport = frmParent.hWnd
    strNoS = strNo
    Set CurSheet = Sheet
    If Mid(UCase(strNo), 4) <> "_INSIDE_1309_3" Then
        If Button = 2 Then PopupMenu mnuReportBill, 2
    End If
End Sub

Private Sub SetMenu(ByVal intState As Integer)
    If intState = 0 Then mnuReportBill.Visible = False: Exit Sub
End Sub
Private Sub ShowBill(frmObject As Object, strNo As String, int记录状态 As Integer, int单据 As Integer, Optional bln在用 As Boolean = False)
    '--------------------------------------------------------------------------------------
    '功能:显示指定单据
    '参数:
    '       frmObject:窗体
    '           strNo:单据号
    '     int记录状态:单据状态(mod(记录状态,3)=1-正常记录;mod(记录状态,3)=2-冲销记录;mod(记录状态,3)=0-已经冲销的记录)
    '         int单据:单据类别( 库房:1-外购入库单;2-其它入库;3-移库单;4-领用;5-其它出库;6-盘存;7-更换单;
    '                           在用:1-领用;2-销售;3-报废单;4-权属变更)
    '--------------------------------------------------------------------------------------
    On Error GoTo errHandle
    Select Case int单据
        Case 1
            frmPurchaseCard.ShowCard frmObject, strNo, 4, int记录状态
        Case 2
            frmSelfMakeCard.ShowCard frmObject, strNo, 4, int记录状态
        Case 3
            frmAccordDrugCard.ShowCard frmObject, strNo, 4, int记录状态
        Case 4
            frmOtherInputCard.ShowCard frmObject, strNo, 4, int记录状态
        Case 5
            frmDiffPriceAdjustCard.ShowCard frmObject, strNo, 4, int记录状态
        Case 6
            frmTransferCard.ShowCard frmObject, strNo, 4, int记录状态
        Case 7
            frmDrawCard.ShowCard frmObject, strNo, 4, False, int记录状态
        Case 11
            frmOtherOutputCard.ShowCard frmObject, strNo, 4, int记录状态
        Case 12
            frmCheckCard.ShowCard frmObject, strNo, 4, int记录状态
        Case 13
            Dim rsTemp As New ADODB.Recordset
            gstrSQL = "Select id,单据,NO,nvl(价格id,0) as 价格id " & _
                      "From 药品收发记录 " & _
                      "Where No=[1] and 费用ID is null And 单据=[2] "
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption & "[取价格记录ID]", strNo, int单据)
                  
            If rsTemp.EOF Or rsTemp.BOF Then Exit Sub
              
            gstrUserName = UserInfo.用户姓名
            With frmAdjust
                .lngBillId = rsTemp!价格id
                .lngMediId = 1
                .intUnit = intChoose级数
                .Show 1, frmObject
            End With
        Case Else
            
            With Frm单据See
                .int记录状态 = int记录状态
                .byt单据 = int单据
                .strNo = strNo
                .mstrPrivs = mstrPrivs
                .Show 1, frmObject
            End With
    End Select
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub 设置权限()
    If Not zlStr.IsHavePrivs(mstrPrivs, "药品明细帐") Then
        tbrThis.Buttons("明细").Visible = False
        mnuFileBatch.Visible = False
    End If
    If Not zlStr.IsHavePrivs(mstrPrivs, "药品总帐") Then
        tbrThis.Buttons("总帐").Visible = False
    End If
End Sub
Private Sub SetFormat(ByVal intType As Integer)
    Dim intCol As Integer
    Dim rsDetail As New Recordset
    Dim bln中药分类 As Boolean
    Dim int类型 As Integer
    
    On Error GoTo errHandle
    
    If Val(cboStock.ItemData(cboStock.ListIndex)) < 0 Then Exit Sub
    
    gstrSQL = "select a.类型 from 诊疗分类目录 a where a.id=[1]"
    If Left(Me.tvwSection_S.SelectedItem.Key, 1) <> "R" Then
        Set rsDetail = zlDatabase.OpenSQLRecord(gstrSQL, "判断选择药品的类别", Mid(Me.tvwSection_S.SelectedItem.Key, 2))
        int类型 = rsDetail!类型
    End If
    
    If int类型 = 3 Or (Left(Me.tvwSection_S.SelectedItem.Key, 1) = "R" And Right(Me.tvwSection_S.SelectedItem.Key, 1) = "7") Then bln中药分类 = True
    If bln中药分类 Then
        vsfList.ColWidth(vsfList.ColIndex("原产地")) = 1500
        vsfBatch.ColWidth(vsfBatch.ColIndex("原产地")) = 1500
    Else
        vsfList.ColWidth(vsfList.ColIndex("原产地")) = 0
        vsfBatch.ColWidth(vsfBatch.ColIndex("原产地")) = 0
    End If
    
    If intType = IniListType.AllList Or intType = IniListType.MainList Then
        With vsfList
            .rows = 1
            .rows = 2

            If gint药品名称显示 = 2 Then
                '显示商品名列
                .ColWidth(.ColIndex("商品名")) = IIf(.ColWidth(.ColIndex("商品名")) = 0, 2000, .ColWidth(.ColIndex("商品名")))
            Else
                '不单独显示商品名列
                .ColWidth(.ColIndex("商品名")) = 0
            End If
            
            .ColWidth(.ColIndex("上次采购价")) = IIf(mblnViewCost, IIf(Me.cboStock.ItemData(Me.cboStock.ListIndex) = 0, 0, IIf(.ColWidth(.ColIndex("上次采购价")) = 0, 1000, .ColWidth(.ColIndex("上次采购价")))), 0)
            .ColWidth(.ColIndex("平均成本价")) = IIf(mblnViewCost, IIf(Me.cboStock.ItemData(Me.cboStock.ListIndex) = 0, 0, IIf(.ColWidth(.ColIndex("平均成本价")) = 0, 1000, .ColWidth(.ColIndex("平均成本价")))), 0)
            .ColWidth(.ColIndex("库存差价")) = IIf(mblnViewCost, IIf(.ColWidth(.ColIndex("库存差价")) = 0, 1000, .ColWidth(.ColIndex("库存差价"))), 0)
            .ColWidth(.ColIndex("成本金额")) = IIf(mblnViewCost, IIf(.ColWidth(.ColIndex("成本金额")) = 0, 1000, .ColWidth(.ColIndex("成本金额"))), 0)
            .ColWidth(.ColIndex("上次供应商")) = IIf(zlStr.IsHavePrivs(mstrPrivs, "供应商查询"), IIf(.ColWidth(.ColIndex("上次供应商")) = 0, 2500, .ColWidth(.ColIndex("上次供应商"))), 0)
            .ColWidth(.ColIndex("库房货位")) = IIf(Me.cboStock.ItemData(Me.cboStock.ListIndex) = 0, 0, IIf(.ColWidth(.ColIndex("库房货位")) = 0, 1500, .ColWidth(.ColIndex("库房货位"))))
            .ColWidth(.ColIndex("储备情况")) = IIf(Me.cboStock.ItemData(Me.cboStock.ListIndex) = 0, 0, IIf(.ColWidth(.ColIndex("储备情况")) = 0, 1500, .ColWidth(.ColIndex("储备情况"))))
            .Row = 1
            
            mstrUnShow_List = "药品ID;用途分类ID;效期;剂量系数;撤档时间;包装"
            If .ColWidth(.ColIndex("商品名")) = 0 Then mstrUnShow_List = mstrUnShow_List & ";商品名"
            If .ColWidth(.ColIndex("上次采购价")) = 0 Then mstrUnShow_List = mstrUnShow_List & ";上次采购价"
            If .ColWidth(.ColIndex("平均成本价")) = 0 Then mstrUnShow_List = mstrUnShow_List & ";平均成本价"
            If .ColWidth(.ColIndex("库存差价")) = 0 Then mstrUnShow_List = mstrUnShow_List & ";库存差价"
            If .ColWidth(.ColIndex("成本金额")) = 0 Then mstrUnShow_List = mstrUnShow_List & ";成本金额"
            If .ColWidth(.ColIndex("上次供应商")) = 0 Then mstrUnShow_List = mstrUnShow_List & ";上次供应商"
            If .ColWidth(.ColIndex("库房货位")) = 0 Then mstrUnShow_List = mstrUnShow_List & ";库房货位"
            If .ColWidth(.ColIndex("储备情况")) = 0 Then mstrUnShow_List = mstrUnShow_List & ";储备情况"
            If .ColWidth(.ColIndex("原产地")) = 0 Then mstrUnShow_List = mstrUnShow_List & ";原产地"
        End With
    End If
    
    If intType = IniListType.AllList Or intType = IniListType.BatchList Then
        With vsfBatch
            .rows = 1
            .rows = 2
            
            .TextMatrix(0, .ColIndex("效期")) = IIf(gtype_UserSysParms.P149_效期显示方式 = 1, "有效期至", "失效期")
            
            If Val(cboStock.ItemData(cboStock.ListIndex)) = 0 Then
                .ColWidth(.ColIndex("库房")) = IIf(.ColWidth(.ColIndex("库房")) = 0, 1500, .ColWidth(.ColIndex("库房")))
                .ColWidth(.ColIndex("批号")) = IIf(.ColWidth(.ColIndex("批号")) = 0, 1500, .ColWidth(.ColIndex("批号")))
                .ColWidth(.ColIndex("效期")) = IIf(.ColWidth(.ColIndex("效期")) = 0, 1500, .ColWidth(.ColIndex("效期")))
                .ColWidth(.ColIndex("产地")) = 0
                .ColWidth(.ColIndex("上次采购价")) = 0
                .ColWidth(.ColIndex("供应商")) = 0
                .ColWidth(.ColIndex("库房货位")) = IIf(.ColWidth(.ColIndex("库房货位")) = 0, 1500, .ColWidth(.ColIndex("库房货位")))
                .ColWidth(.ColIndex("储备情况")) = IIf(.ColWidth(.ColIndex("储备情况")) = 0, 1500, .ColWidth(.ColIndex("储备情况")))
                .ColWidth(.ColIndex("原产地")) = 0
            Else
                .ColWidth(.ColIndex("库房")) = 0
                .ColWidth(.ColIndex("批号")) = IIf(.ColWidth(.ColIndex("批号")) = 0, 1500, .ColWidth(.ColIndex("批号")))
                .ColWidth(.ColIndex("效期")) = IIf(.ColWidth(.ColIndex("效期")) = 0, 1500, .ColWidth(.ColIndex("效期")))
                .ColWidth(.ColIndex("产地")) = IIf(.ColWidth(.ColIndex("产地")) = 0, 1500, .ColWidth(.ColIndex("产地")))
                .ColWidth(.ColIndex("上次采购价")) = IIf(mblnViewCost, IIf(.ColWidth(.ColIndex("上次采购价")) = 0, 1500, .ColWidth(.ColIndex("上次采购价"))), 0)
                .ColWidth(.ColIndex("供应商")) = IIf(zlStr.IsHavePrivs(mstrPrivs, "供应商查询"), IIf(.ColWidth(.ColIndex("供应商")) = 0, 2500, .ColWidth(.ColIndex("供应商"))), 0)
                .ColWidth(.ColIndex("库房货位")) = 0
                .ColWidth(.ColIndex("储备情况")) = 0
                If bln中药分类 Then
                    .ColWidth(.ColIndex("原产地")) = IIf(.ColWidth(.ColIndex("原产地")) = 0, 1500, .ColWidth(.ColIndex("原产地")))
                End If
            End If

            .ColWidth(.ColIndex("实际差价")) = IIf(mblnViewCost, IIf(.ColWidth(.ColIndex("实际差价")) = 0, 1500, .ColWidth(.ColIndex("实际差价"))), 0)
            .ColWidth(.ColIndex("成本价")) = IIf(mblnViewCost, IIf(.ColWidth(.ColIndex("成本价")) = 0, 1500, .ColWidth(.ColIndex("成本价"))), 0)
            .ColWidth(.ColIndex("成本金额")) = IIf(mblnViewCost, IIf(.ColWidth(.ColIndex("成本金额")) = 0, 1500, .ColWidth(.ColIndex("成本金额"))), 0)
            
            mstrUnShow_Batch = "不可能有这行"
            If .ColWidth(.ColIndex("库房")) = 0 Then mstrUnShow_Batch = mstrUnShow_Batch & ";库房"
            If .ColWidth(.ColIndex("批号")) = 0 Then mstrUnShow_Batch = mstrUnShow_Batch & ";批号"
            If .ColWidth(.ColIndex("效期")) = 0 Then mstrUnShow_Batch = mstrUnShow_Batch & ";效期"
            If .ColWidth(.ColIndex("产地")) = 0 Then mstrUnShow_Batch = mstrUnShow_Batch & ";产地"
            If .ColWidth(.ColIndex("上次采购价")) = 0 Then mstrUnShow_Batch = mstrUnShow_Batch & ";上次采购价"
            If .ColWidth(.ColIndex("供应商")) = 0 Then mstrUnShow_Batch = mstrUnShow_Batch & ";供应商"
            If .ColWidth(.ColIndex("库房货位")) = 0 Then mstrUnShow_Batch = mstrUnShow_Batch & ";库房货位"
            If .ColWidth(.ColIndex("实际差价")) = 0 Then mstrUnShow_Batch = mstrUnShow_Batch & ";实际差价"
            If .ColWidth(.ColIndex("成本价")) = 0 Then mstrUnShow_Batch = mstrUnShow_Batch & ";成本价"
            If .ColWidth(.ColIndex("成本金额")) = 0 Then mstrUnShow_Batch = mstrUnShow_Batch & ";成本金额"
            If .ColWidth(.ColIndex("储备情况")) = 0 Then mstrUnShow_Batch = mstrUnShow_Batch & ";储备情况"
            If .ColWidth(.ColIndex("原产地")) = 0 Then mstrUnShow_Batch = mstrUnShow_Batch & ";原产地"
        End With
    End If
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub DataBound(ByVal rsData As ADODB.Recordset)
    Dim lngColor As Long
    Dim lngRow As Long, lngCol As Long
    Dim lngPrice As Single
    Dim lng下限 As Long
    Dim strTemp As String
    Dim rsTemp As ADODB.Recordset
    
    If rsData.EOF Then Exit Sub
    
    On Error GoTo ErrHand
    
    mblnRefresh = True
    
    With vsfList
        .Redraw = flexRDNone
        Do While Not rsData.EOF
            If rsData.AbsolutePosition > .rows - 1 Then .rows = .rows + 1
            .Row = rsData.AbsolutePosition
            '填充数据
            .TextMatrix(.Row, .ColIndex("药品ID")) = rsData!药品id
            .TextMatrix(.Row, .ColIndex("用途分类ID")) = rsData!用途分类id
            .TextMatrix(.Row, .ColIndex("编码")) = rsData!编码
            .TextMatrix(.Row, .ColIndex("标识码")) = zlCommFun.NVL(rsData!药卡号, "")
            
            If gint药品名称显示 = 0 Or gint药品名称显示 = 2 Then
                .TextMatrix(.Row, .ColIndex("名称")) = rsData!通用名
            Else
                .TextMatrix(.Row, .ColIndex("名称")) = IIf(IsNull(rsData!商品名), rsData!通用名, rsData!商品名)
            End If
            
            .TextMatrix(.Row, .ColIndex("商品名")) = IIf(IsNull(rsData!商品名), "", rsData!商品名)
            .TextMatrix(.Row, .ColIndex("基本药物")) = IIf(IsNull(rsData!基本药物), "", rsData!基本药物)
            .TextMatrix(.Row, .ColIndex("规格")) = zlCommFun.NVL(rsData!规格, "")
            .TextMatrix(.Row, .ColIndex("产地")) = zlCommFun.NVL(rsData!产地, "")
            .TextMatrix(.Row, .ColIndex("原产地")) = zlCommFun.NVL(rsData!原产地, "")
            .TextMatrix(.Row, .ColIndex("效期")) = zlCommFun.NVL(rsData!效期, "")
            .TextMatrix(.Row, .ColIndex("药库分批")) = zlCommFun.NVL(rsData!药库分批, "否")
            .TextMatrix(.Row, .ColIndex("单位")) = zlCommFun.NVL(rsData!单位, "")
            .TextMatrix(.Row, .ColIndex("平均成本价")) = Format(rsData!平均成本价, mStr成本价)
            .TextMatrix(.Row, .ColIndex("上次采购价")) = Format(rsData!上次采购价, mStr成本价)
            
            lngPrice = rsData!当前售价
            
            .TextMatrix(.Row, .ColIndex("当前售价")) = Format(lngPrice, mStr单价)
            .TextMatrix(.Row, .ColIndex("剂量系数")) = Format(rsData!系数1, mStr单价)
            .TextMatrix(.Row, .ColIndex("可用数量")) = Format(rsData!可用数量, mStr数量)
            .TextMatrix(.Row, .ColIndex("库存数量")) = Format(rsData!实际数量, mStr数量)
            .TextMatrix(.Row, .ColIndex("库存金额")) = Format(rsData!实际金额, mStr金额)
            .TextMatrix(.Row, .ColIndex("库存差价")) = Format(rsData!实际差价, mStr金额)
            .TextMatrix(.Row, .ColIndex("成本金额")) = Format(rsData!平均成本价 * rsData!实际数量, mStr金额)
            .TextMatrix(.Row, .ColIndex("撤档时间")) = zlCommFun.NVL(rsData!撤档时间, "")
            .TextMatrix(.Row, .ColIndex("包装")) = zlCommFun.NVL(rsData!除数, 1)
            .TextMatrix(.Row, .ColIndex("上次供应商")) = zlCommFun.NVL(rsData!上次供应商, "")
            .TextMatrix(.Row, .ColIndex("库房货位")) = zlCommFun.NVL(rsData!库房货位, "")
            If .TextMatrix(.Row, .ColIndex("库存数量")) <> "" And cboStock.Tag <> 0 Then
                lng下限 = IIf(IsNull(rsData!下限), 0, rsData!下限)
                If lng下限 = 0 Then
                    strTemp = "无下限"
                ElseIf Val(.TextMatrix(.Row, .ColIndex("库存数量"))) > lng下限 Then
                    strTemp = "充足"
                ElseIf Val(.TextMatrix(.Row, .ColIndex("库存数量"))) = lng下限 Then
                    strTemp = "持平"
                ElseIf Val(.TextMatrix(.Row, .ColIndex("库存数量"))) < lng下限 Then
                    strTemp = "不足"
                End If
                .TextMatrix(.Row, .ColIndex("储备情况")) = strTemp
            End If
            '上色
            If bln包含停用药品 Then
                lngColor = IIf(Trim(.TextMatrix(.Row, .ColIndex("撤档时间"))) = "", glng黑色, glng红色)
                .Cell(flexcpForeColor, .Row, 0, .Row, .Cols - 1) = lngColor
            End If
            If cboStock.ItemData(Me.cboStock.ListIndex) > 0 Then
                '检查是否有批次效期过了，只有选择具体某个库房时才处理
                If rsData!报警 = 1 And IsNull(rsData!库存效期) = False Then
                    .Cell(flexcpPicture, .Row, 0, .Row, 0) = imglvw.ListImages(3).Picture
                End If
            End If
            rsData.MoveNext
        Loop
        
        '填写排序编码
        Call SetSortCode
                
        .Redraw = flexRDDirect
    End With
    
    mblnRefresh = False
    Exit Sub
ErrHand:
    If ErrCenter() Then Resume
    Call SaveErrLog
    mblnRefresh = False
End Sub

'Modified By 朱玉宝 2003-12-10 地区：泸州
Private Sub txt药品信息_GotFocus()
    Call zlControl.TxtSelAll(txt药品信息)
End Sub

Private Sub txt药品信息_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strFind As String, StrBit As String
    Dim strInput As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    txt药品信息.Text = Replace(txt药品信息.Text, "'", "")
    strInput = Trim(UCase(txt药品信息.Text))
    
    If strInput = "" Then Exit Sub
    
    StrBit = GetSetting(appName:="ZLSOFT", Section:="公共模块\操作", Key:="输入匹配", Default:="0")
    StrBit = IIf(StrBit = "0", "%", "")
    
    strFind = " And (A.编码 Like [7] OR B.名称 Like [7] OR B.简码 LIKE [7])"
    
    If IsNumeric(strInput) Then                         '10,11.输入全是数字时只匹配编码'对于药品,则要匹配简码(码类为3的数字码)
        If Mid(gtype_UserSysParms.P44_输入匹配, 1, 1) = "1" Then strFind = " And (A.编码 Like [7] Or B.简码 Like [7] And B.码类=3)"
    ElseIf zlStr.IsCharAlpha(strInput) Then         '01,11.输入全是字母时只匹配简码
        If Mid(gtype_UserSysParms.P44_输入匹配, 2, 1) = "1" Then strFind = " And B.简码 Like [7] "
    ElseIf zlStr.IsCharChinese(strInput) Then
        strFind = " And B.名称 Like [7] "
    End If
    
    SQLCondition.str药品信息 = StrBit & strInput & "%"
        
    If Not ReFreshDrugData(cboStock.ItemData(cboStock.ListIndex), IIf(Left(Me.tvwSection_S.SelectedItem.Key, 1) = "R", 0, Mid(Me.tvwSection_S.SelectedItem.Key, 2)), strFind, False) Then Exit Sub
    Me.tvwSection_S.Tag = "T"
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

Private Sub vsfList_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Long
    
    If Button = 2 Then '列选择器
        If Val(zlDatabase.GetPara("使用个性化风格")) = 0 Then Exit Sub
        
        If vsfList.MouseRow <> 0 Then Exit Sub
        
        InitColSelList IniListType.MainList, vsfList
        
        '根据当前状态直接确定勾选状态
        With vsfColSel
            If .Visible Then
                .Visible = False
                vsfList.SetFocus
            Else
                For i = .FixedRows To .rows - 1
                    If vsfList.colHidden(.RowData(i)) Or vsfList.ColWidth(.RowData(i)) = 0 Then
                        .TextMatrix(i, 0) = 0
                    Else
                        .TextMatrix(i, 0) = 1
                    End If
                Next
                
                .Height = .RowHeightMin * .rows + 150
                .Top = vsfList.Top + vsfList.RowHeight(0)
                If .Top + .Height > Me.ScaleHeight - vsfList.Top Then
                    .Height = Me.ScaleHeight - .Top - vsfList.Top
                    .Width = 1750
                Else
                    .Width = 1470
                End If
                
                .Left = vsfList.Left + x
                .ZOrder
                .Visible = True
                .SetFocus
            End If
        End With
    End If
End Sub



