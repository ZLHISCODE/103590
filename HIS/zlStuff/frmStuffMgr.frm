VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmStuffMgr 
   BackColor       =   &H8000000A&
   Caption         =   "卫材目录管理"
   ClientHeight    =   8280
   ClientLeft      =   60
   ClientTop       =   600
   ClientWidth     =   11415
   Icon            =   "frmStuffMgr.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8280
   ScaleWidth      =   11415
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picClass 
      Height          =   4455
      Left            =   0
      ScaleHeight     =   4395
      ScaleWidth      =   2355
      TabIndex        =   16
      Top             =   840
      Width           =   2415
      Begin VB.CommandButton cmdKind 
         Caption         =   "过滤结果(&1)"
         Height          =   300
         Index           =   1
         Left            =   0
         TabIndex        =   18
         TabStop         =   0   'False
         Top             =   300
         UseMaskColor    =   -1  'True
         Width           =   2295
      End
      Begin VB.CommandButton cmdKind 
         Caption         =   "分    类(&0)"
         Height          =   300
         Index           =   0
         Left            =   0
         TabIndex        =   17
         TabStop         =   0   'False
         Top             =   15
         Width           =   2295
      End
      Begin MSComctlLib.TreeView tvwClass 
         Height          =   3030
         Left            =   0
         TabIndex        =   19
         Top             =   840
         Width           =   2295
         _ExtentX        =   4048
         _ExtentY        =   5345
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   494
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         ImageList       =   "imgList"
         BorderStyle     =   1
         Appearance      =   1
      End
   End
   Begin VB.PictureBox picCost 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2640
      Left            =   -810
      ScaleHeight     =   2640
      ScaleWidth      =   5730
      TabIndex        =   14
      Top             =   5310
      Width           =   5730
      Begin VSFlex8Ctl.VSFlexGrid vsCost 
         Height          =   2070
         Left            =   225
         TabIndex        =   15
         Top             =   30
         Width           =   7080
         _cx             =   12488
         _cy             =   3651
         Appearance      =   3
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
         ForeColorSel    =   0
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
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
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   8
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmStuffMgr.frx":030A
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
         ExplorerBar     =   3
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
         BackColorFrozen =   16777215
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.PictureBox picVBar_S 
      Appearance      =   0  'Flat
      BackColor       =   &H8000000A&
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
      Height          =   6660
      Left            =   2520
      MousePointer    =   9  'Size W E
      ScaleHeight     =   6660
      ScaleWidth      =   45
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   840
      Width           =   45
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
      Height          =   45
      Left            =   3360
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   6075
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   5520
      Width           =   6075
   End
   Begin VB.PictureBox picTabPrice 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2550
      Left            =   2640
      ScaleHeight     =   2550
      ScaleWidth      =   3735
      TabIndex        =   9
      Top             =   4590
      Width           =   3735
      Begin VB.Frame fraComment 
         BorderStyle     =   0  'None
         Height          =   480
         Left            =   60
         TabIndex        =   10
         Top             =   1905
         Width           =   7410
         Begin VB.Label lblComment 
            AutoSize        =   -1  'True
            Caption         =   "1、时价材料，指导批发价6元/支。。。"
            Height          =   180
            Index           =   3
            Left            =   0
            TabIndex        =   12
            Top             =   30
            Width           =   3150
         End
         Begin VB.Label lblComment 
            AutoSize        =   -1  'True
            Caption         =   "2、最高售价198.25元/支；根据病人身份费别进行优惠或加价。"
            Height          =   180
            Index           =   4
            Left            =   0
            TabIndex        =   11
            Top             =   270
            Width           =   5040
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsPrice 
         Height          =   1725
         Left            =   90
         TabIndex        =   13
         Top             =   105
         Width           =   7365
         _cx             =   12991
         _cy             =   3043
         Appearance      =   3
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
         ForeColorSel    =   0
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
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
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   15
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmStuffMgr.frx":0416
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
         ExplorerBar     =   3
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
         BackColorFrozen =   16777215
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.PictureBox picStuffLst 
      Appearance      =   0  'Flat
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2640
      Left            =   6885
      ScaleHeight     =   2640
      ScaleWidth      =   2670
      TabIndex        =   7
      Top             =   4680
      Width           =   2670
      Begin VSFlex8Ctl.VSFlexGrid vsStuff 
         Height          =   2070
         Left            =   0
         TabIndex        =   8
         Top             =   0
         Width           =   7080
         _cx             =   12488
         _cy             =   3651
         Appearance      =   3
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
         ForeColorSel    =   0
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
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
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   37
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmStuffMgr.frx":061F
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
         ExplorerBar     =   3
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
         BackColorFrozen =   16777215
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin XtremeSuiteControls.TabControl TbList 
      Height          =   3630
      Left            =   2445
      TabIndex        =   6
      Top             =   4170
      Width           =   8475
      _Version        =   589884
      _ExtentX        =   14949
      _ExtentY        =   6403
      _StockProps     =   64
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   10080
      Top             =   1200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   4
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":0B4B
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":10E5
            Key             =   "expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":167F
            Key             =   "start"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":1B91
            Key             =   "stop"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwItems 
      Height          =   2895
      Left            =   2520
      TabIndex        =   1
      Top             =   960
      Width           =   7200
      _ExtentX        =   12700
      _ExtentY        =   5106
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      AllowReorder    =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   2
      Top             =   7905
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmStuffMgr.frx":1DAB
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15055
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
   Begin ComCtl3.CoolBar clbThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   11415
      _ExtentX        =   20135
      _ExtentY        =   1376
      BandCount       =   1
      _CBWidth        =   11415
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tlbThis"
      MinWidth1       =   24000
      MinHeight1      =   720
      Width1          =   8730
      FixedBackground1=   0   'False
      Key1            =   "Comm"
      NewRow1         =   0   'False
      Begin MSComctlLib.Toolbar tlbThis 
         Height          =   720
         Left            =   30
         TabIndex        =   20
         Top             =   30
         Width           =   24000
         _ExtentX        =   42333
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "imgGray"
         HotImageList    =   "imgColor"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   20
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "预览"
               Key             =   "Preview"
               Description     =   "预览"
               Object.ToolTipText     =   "预览当前表"
               Object.Tag             =   "预览"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "打印"
               Key             =   "Print"
               Description     =   "打印"
               Object.ToolTipText     =   "打印当前表"
               Object.Tag             =   "打印"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "分类"
               Key             =   "Class"
               Description     =   "分类"
               Object.ToolTipText     =   "调整材料分类"
               Object.Tag             =   "分类"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split4"
               Style           =   3
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "增加"
               Key             =   "增加"
               Description     =   "增加"
               Object.ToolTipText     =   "增加规格或品种"
               Object.Tag             =   "增加"
               ImageKey        =   "Add"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   2
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "增加品种"
                     Object.Tag             =   "增加品种"
                     Text            =   "增加品种"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Key             =   "增加规格"
                     Object.Tag             =   "增加规格"
                     Text            =   "增加规格"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "修改"
               Description     =   "修改"
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
               Key             =   "Split1"
               Description     =   "删除"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "启用"
               Key             =   "Start"
               Description     =   "启用"
               Object.ToolTipText     =   "启用指定的停用材料"
               Object.Tag             =   "启用"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "停用"
               Key             =   "Stop"
               Description     =   "停用"
               Object.ToolTipText     =   "停用指定的在用材料"
               Object.Tag             =   "停用"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split3"
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split12"
               Style           =   3
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查看"
               Key             =   "查看"
               Object.ToolTipText     =   "查看"
               Object.Tag             =   "查看"
               ImageKey        =   "View"
               Style           =   5
               BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
                  NumButtonMenus  =   4
                  BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "大图标"
                  EndProperty
                  BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "小图标"
                  EndProperty
                  BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "列表"
                  EndProperty
                  BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                     Text            =   "详细资料"
                  EndProperty
               EndProperty
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "查找"
               Key             =   "Find"
               Description     =   "查找"
               Object.ToolTipText     =   "查找卫材目录"
               Object.Tag             =   "查找"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "过滤"
               Key             =   "Filter"
               Description     =   "过滤"
               Object.ToolTipText     =   "过滤卫材目录"
               Object.Tag             =   "过滤"
               ImageIndex      =   16
            EndProperty
            BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "Help"
               Description     =   "帮助"
               Object.ToolTipText     =   "当前帮助主题"
               Object.Tag             =   "帮助"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "Exit"
               Description     =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   11
            EndProperty
         EndProperty
         Begin VB.PictureBox picFind 
            BorderStyle     =   0  'None
            Height          =   300
            Left            =   9000
            ScaleHeight     =   285.714
            ScaleMode       =   0  'User
            ScaleWidth      =   495
            TabIndex        =   22
            Top             =   240
            Width           =   495
            Begin VB.Label lbl查找 
               Caption         =   "查找"
               Height          =   255
               Left            =   120
               TabIndex        =   23
               Top             =   75
               Width           =   495
            End
         End
         Begin VB.TextBox txtFind 
            Height          =   300
            Left            =   9600
            MaxLength       =   10
            TabIndex        =   21
            Tag             =   "简码"
            Top             =   240
            Width           =   1425
         End
      End
   End
   Begin MSComctlLib.ImageList imgGray 
      Left            =   7680
      Top             =   525
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":263D
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":2857
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":2A71
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":2C8B
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":2EA5
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":359F
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":37B9
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":39D3
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":40CD
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":42E7
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":4507
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":4727
            Key             =   "Add"
            Object.Tag             =   "Add"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":4941
            Key             =   "Modify"
            Object.Tag             =   "Modify"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":4B5B
            Key             =   "Delete"
            Object.Tag             =   "Delete"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":4D75
            Key             =   "View"
            Object.Tag             =   "View"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":4F8F
            Key             =   "Filter"
            Object.Tag             =   "Filter"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgColor 
      Left            =   6915
      Top             =   525
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   16
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":52A9
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":54C9
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":56E9
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":5903
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":5B1D
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":6217
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":6431
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":664B
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":6D45
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":6F5F
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":717F
            Key             =   "Quit"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":739F
            Key             =   "Add"
            Object.Tag             =   "Add"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":75B9
            Key             =   "Modify"
            Object.Tag             =   "Modify"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":77D3
            Key             =   "Delete"
            Object.Tag             =   "Delete"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":79ED
            Key             =   "View"
            Object.Tag             =   "View"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffMgr.frx":7C07
            Key             =   "Filter"
            Object.Tag             =   "Filter"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picUp_S 
      Appearance      =   0  'Flat
      BackColor       =   &H80000011&
      BorderStyle     =   0  'None
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   2790
      ScaleHeight     =   255
      ScaleWidth      =   7170
      TabIndex        =   5
      Top             =   3900
      Width           =   7170
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
      Begin VB.Menu mnuFileMlPrint 
         Caption         =   "目录打印(&M)"
      End
      Begin VB.Menu mnuFileMlPreview 
         Caption         =   "目录预览(&Y)"
      End
      Begin VB.Menu mnuFileSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePara 
         Caption         =   "参数设置(&A)"
         Shortcut        =   {F12}
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileSpt2 
         Caption         =   "-"
         Visible         =   0   'False
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuClass 
      Caption         =   "分类(&K)"
      Begin VB.Menu mnuClassAdd 
         Caption         =   "新增(&I)"
         Shortcut        =   +{INSERT}
      End
      Begin VB.Menu mnuClassMod 
         Caption         =   "修改(&U)"
      End
      Begin VB.Menu mnuClassDel 
         Caption         =   "删除(&E)"
         Shortcut        =   +{DEL}
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "材料(&E)"
      Begin VB.Menu mnuEditAddName 
         Caption         =   "增加品种(&P)"
      End
      Begin VB.Menu mnuEditModifyName 
         Caption         =   "修改品种(&E)"
      End
      Begin VB.Menu mnuEditDeleName 
         Caption         =   "删除品种(&D)"
      End
      Begin VB.Menu mnuEditSplit 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditItemAdd 
         Caption         =   "新增规格(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditItemMod 
         Caption         =   "修改规格(&M)"
         Shortcut        =   ^M
      End
      Begin VB.Menu mnuEditItemDel 
         Caption         =   "删除规格(&D)"
         Shortcut        =   {DEL}
      End
      Begin VB.Menu mnuEditSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditItemPart 
         Caption         =   "存储库房(&F)..."
      End
      Begin VB.Menu mnuEditSpecLimit 
         Caption         =   "储备限量(&L)..."
      End
      Begin VB.Menu mnuEditSpecSelf 
         Caption         =   "自制材料(&H)..."
      End
      Begin VB.Menu mnuEditUnit 
         Caption         =   "中标单位(&Z)"
      End
      Begin VB.Menu mnuEditSpt2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditSubsection 
         Caption         =   "分段加成率(&J)"
      End
      Begin VB.Menu mnuEditExcel 
         Caption         =   "导入项目"
      End
      Begin VB.Menu mnuEditSpt3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditStart 
         Caption         =   "启用(&R)"
      End
      Begin VB.Menu mnuEditStop 
         Caption         =   "停用(&S)"
      End
      Begin VB.Menu mnuEditSpt4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditPrice 
         Caption         =   "调价(&T)"
      End
   End
   Begin VB.Menu mnuPrice 
      Caption         =   "价格(&P)"
      Begin VB.Menu mnuPriceTable 
         Caption         =   "调价记录表(&S)"
      End
      Begin VB.Menu mnuPriceLists 
         Caption         =   "材料价目表(&L)..."
         Shortcut        =   ^L
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
      Begin VB.Menu mnuViewStates 
         Caption         =   "状态栏(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewSpt1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewIcon 
         Caption         =   "大图标(&G)"
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
      Begin VB.Menu mnuViewSp 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewStoped 
         Caption         =   "显示停用材料(&C)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewPrices 
         Caption         =   "显示历史价格(&H)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuViewDownLevel 
         Caption         =   "显示所有下级(&X)"
      End
      Begin VB.Menu mnuViewSpt2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "查找(&F)..."
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuFilter 
         Caption         =   "过滤(&T)"
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuViewSpt3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "刷新(&R)"
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
         Caption         =   "Web上的中联(&W)"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "中联主页(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "中联论坛(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "发送反馈(&E)..."
         End
      End
      Begin VB.Menu mnuHelp1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "关于(&A)..."
      End
   End
End
Attribute VB_Name = "frmStuffMgr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mstrPrivs As String       '用户具有本程序的具体权限

Dim rsTemp As New ADODB.Recordset
Dim objNode As Node
Dim objItem As ListItem
Dim intCount As Integer, intRow As Integer, intCol As Integer
Dim strTemp As String
Dim mintUnit  As Integer
Private mlngModule As Long
Dim mstrMaterialID As String
Private Declare Function SetParent Lib "user32 " (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private mstrFindValue As String
Private mrsFind As ADODB.Recordset
Private mrsRefClasses As ADODB.Recordset
Private mrsRefRecords As ADODB.Recordset

Private Enum colPrice
    NO = 0
    序号
    材料类型
    指导批价
    扣率
    指导售价
    指导差率
    执行情况
    屏蔽费别
    材料
    单位
    售价
    收入项目
    说明
    执行日期
End Enum
Private mbln显示下级 As Boolean         '是否显示下级所有材
'----------------------------------------------------------------------------------------------------------
'刘兴宏:增加小数位数的格式串
'修改:2007/03/06
Private mFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------


Private Sub SetMenuEnable()
    Dim blnStop As Boolean  '品种停用
    Dim blnListStop As Boolean  '规格停用
    Dim blnSel As Boolean   '选中品种的
    Dim blnListSel As Boolean '选中的材料
    
    Dim blnData As Boolean '存在数据否
    Dim blnSel分类 As Boolean '选中分类
        
    blnSel分类 = Not tvwClass.SelectedItem Is Nothing
    
    blnData = Me.lvwItems.ListItems.count <> 0
    blnSel = Not Me.lvwItems.SelectedItem Is Nothing
    If blnSel Then
        blnStop = Me.lvwItems.SelectedItem.Icon = "stop"
    End If
    
    With vsStuff
        blnListSel = .RowData(.Row) <> 0
        If blnSel Then
            blnListStop = .TextMatrix(.Row, .ColIndex("撤档时间")) <> ""
        End If
    End With
    
    
    '对品种的相关操作进行设置
    mnuEditModifyName.Enabled = blnSel And Not blnStop
    mnuEditDeleName.Enabled = blnSel And Not blnStop
    
    '确定卫生材料的增删改
    mnuEditItemAdd.Enabled = blnSel And Not blnStop
    tlbThis.Buttons("增加").ButtonMenus("增加规格").Enabled = mnuEditItemAdd.Enabled
    
    mnuEditItemMod.Enabled = blnListSel And Not blnListStop
    mnuEditItemDel.Enabled = blnListSel And Not blnListStop
    mnuEditPrice.Enabled = blnListSel And Not blnListStop
    
    If Me.ActiveControl Is lvwItems Then
        tlbThis.Buttons("修改").Enabled = blnSel And Not blnStop
        tlbThis.Buttons("删除").Enabled = blnSel And Not blnStop
        mnuEditStop.Enabled = Not blnStop And blnSel
        mnuEditStart.Enabled = blnStop And blnSel
    Else
        tlbThis.Buttons("修改").Enabled = blnListSel And Not blnListStop
        tlbThis.Buttons("删除").Enabled = blnListSel And Not blnListStop
        mnuEditStop.Enabled = Not blnListStop And blnListSel
        mnuEditStart.Enabled = blnListStop And blnListSel
    End If
    
    tlbThis.Buttons("Start").Enabled = mnuEditStart.Enabled
    tlbThis.Buttons("Stop").Enabled = mnuEditStop.Enabled
    
    '确定分类属性
    mnuClassDel.Enabled = blnSel分类
    mnuClassMod.Enabled = blnSel分类
    
    '需要检查权限
    Dim blnVisible As Boolean
    
    If Me.ActiveControl Is Me.vsStuff Then
        blnVisible = True
    ElseIf Me.ActiveControl Is Me.vsPrice Then
        blnVisible = zlStr.IsHavePrivs(mstrPrivs, "售价管理")
    ElseIf Me.ActiveControl Is Me.vsCost Then
        blnVisible = zlStr.IsHavePrivs(mstrPrivs, "成本价管理")
    Else
        blnVisible = True
    End If
    
    '确定打印及预览
    mnuFileExcel.Enabled = blnData
    mnuFilePreview.Enabled = blnData
    mnuFilePrint.Enabled = blnData
    tlbThis.Buttons("Preview").Enabled = blnData
    tlbThis.Buttons("Print").Enabled = blnData
End Sub
Private Sub clbThis_Resize()
    Me.clbThis.Bands(1).MinHeight = Me.tlbThis.Height
    Me.clbThis.Refresh
    Call Form_Resize
End Sub

Private Sub cmdKind_Click(Index As Integer)
    Dim intCount As Integer
    Call SaveListViewState(Me.lvwItems, Me.Name & Val(Me.tvwClass.Tag), Me.lvwItems.View)
    For intCount = Me.cmdKind.LBound To Me.cmdKind.UBound
        If intCount <= Index Then
            Me.cmdKind(intCount).Tag = 0
        Else
            Me.cmdKind(intCount).Tag = 1
        End If
    Next
    
    '无数据，不能打印或预览 2010-5-10
    mnuFileMlPrint.Enabled = Not tvwClass.SelectedItem Is Nothing
    mnuFileMlPreview.Enabled = Not tvwClass.SelectedItem Is Nothing
    
    '装数据并调整界面
    If Me.lvwItems.Visible Then
        Call picClass_Resize
        Me.tvwClass.SetFocus
    End If
    If Index = 0 Then
        If Val(tvwClass.Tag) <> Index Then
            Me.tvwClass.Tag = Index
            Call zlRefClasses
        End If
        Me.mnuViewFind.Enabled = True
        Me.tlbThis.Buttons("Find").Enabled = True
    Else
        Me.tvwClass.Tag = Index
        Call zlRefClasses
        Me.mnuViewFind.Enabled = False
        Me.tlbThis.Buttons("Find").Enabled = False
        frmStuffFind.Hide
    End If
End Sub

Private Sub Form_Activate()
    Me.lvwItems.Visible = True
End Sub

Private Sub Form_Load()
    Dim rsTemp As New ADODB.Recordset
    Dim rs收入项目 As ADODB.Recordset
    Dim bln收入项目 As Boolean
    
    '界面恢复
    Call RestoreWinState(Me, App.ProductName)
    mlngModule = glngModul
    mbln显示下级 = Val(zlDatabase.GetPara("包含下级卫材", glngSys, mlngModule, "0")) = 1
    
    Me.mnuViewDownLevel.Checked = mbln显示下级
    Me.mnuViewStoped.Checked = False
    Me.mnuViewPrices.Checked = False
    
    mnuViewIcon(0).Checked = False
    mnuViewIcon(1).Checked = False
    mnuViewIcon(2).Checked = False
    mnuViewIcon(3).Checked = False
    mnuViewIcon(lvwItems.View).Checked = True
    
    '检查是否以库房单位显示价格
    mintUnit = Get定价单位
  
    '刘兴宏:增加小数格式化串
    With mFMT
        .FM_成本价 = GetFmtString(mintUnit, g_成本价)
        .FM_金额 = GetFmtString(mintUnit, g_金额)
        .FM_零售价 = GetFmtString(mintUnit, g_售价)
        .FM_数量 = GetFmtString(mintUnit, g_数量)
    End With
    
    
    'mstrFormat = IIf(mintUnit, "####0.0000;-####0.0000; ;", "####0.0000000;-####0.0000000; ;")
    '可直接通过菜单进行的权限控制
    mstrPrivs = ";" & gstrPrivs & ";"

    '控制权限
    Call SetPopedom
    '初始价格列头
     
    Call InitHgdPrivceHeadcol
       
'    Me.picHBar_S.Top = Me.ScaleHeight - IIf(Me.stbThis.Visible, Me.stbThis.Height, 0) - 2500
    '初始例头
    Call InitLvwStuffColHead
    
    '获取相关数据
    zlRefClasses
    '2006-04-25:刘兴宏,统一增加报表发布到模块的功能
    Call zlDatabase.ShowReportMenu(Me, glngSys, glngModul, gstrPrivs)
     
    Call SetMenuEnable
    Call InitTabCtl
    Call vsCost_LostFocus
    Call vsPrice_LostFocus
    Call vsStuff_LostFocus
    
    If Val(zlDatabase.GetPara("使用个性化风格", 0, 0)) = 1 Then
        strTemp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\分割", "横向", "0")
        If strTemp <> "0" Then
            Me.picVBar_S.Left = CLng(strTemp)
        End If
        strTemp = GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\分割", "纵向", "0")
        If strTemp <> "0" Then
            Me.picHBar_S.Top = CLng(strTemp)
        End If
    End If
    
    zl_vsGrid_Para_Restore mlngModule, vsCost, Me.Caption, "成本价调价记录"
    zl_vsGrid_Para_Restore mlngModule, vsPrice, Me.Caption, "售价调价记录"
    zl_vsGrid_Para_Restore mlngModule, vsStuff, Me.Caption, "材料规格记录"
    
    If vsStuff.ColWidth(vsStuff.ColIndex("编码")) = 0 Then vsStuff.ColWidth(vsStuff.ColIndex("编码")) = 1600
    vsStuff.ColHidden(vsStuff.ColIndex("编码")) = False
    
    Call cmdKind_Click(0)
    
    gstrSQL = "select 参数值 from zlParameters where 模块=1711 and 参数名='收入项目对应'"
    Set rs收入项目 = zlDatabase.OpenSQLRecord(gstrSQL, "收入项目", "收入项目对应")
    If rs收入项目.RecordCount > 0 Then
        If IsNull(rs收入项目!参数值) Then
            bln收入项目 = True
        End If
    End If
    
    If bln收入项目 = True Then
        MsgBox "请设置各材质对应的收入项目！", vbInformation, gstrSysName
        frmStuffPara.ShowMe mstrPrivs, Me
        If gblnIncomeItem = False Then
            Unload Me
        End If
    End If
End Sub
Private Sub InitTabCtl()
    '-----------------------------------------------------------------------------------------------------------
    '功能:初始化Tab控件
    '编制:刘兴宏
    '日期:2007/05/24
    '-----------------------------------------------------------------------------------------------------------
    TbList.SetImageList Me.imgList
    TbList.InsertItem 0, "材料规格(&1)", picStuffLst.hwnd, 0
    TbList.InsertItem 1, "售价调价记录(&2)", picTabPrice.hwnd, 0
    TbList.InsertItem 2, "成本价调价记录(&3)", picCost.hwnd, 0
    
    TbList.Item(2).Visible = InStr(1, mstrPrivs, ";成本价管理;") > 0
    TbList.PaintManager.Appearance = xtpTabAppearancePropertyPage2003
End Sub

Private Sub SetPopedom()
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:权限控制
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim blnModify As Boolean    '允许修改
    Dim bln调价 As Boolean
    
    blnModify = InStr(1, mstrPrivs, ";售价管理;") <> 0      '有售价,可以更改
    blnModify = blnModify Or InStr(1, mstrPrivs, ";成本价管理;") <> 0      '有成本价,也可以更改
    blnModify = blnModify Or InStr(1, mstrPrivs, ";管理扣率;") <> 0      '有扣率,也可以更改
    blnModify = blnModify Or InStr(1, mstrPrivs, ";指导价格管理;") <> 0      '有指导价格,也可以更改
    blnModify = blnModify Or InStr(1, mstrPrivs, ";医保用料目录;") <> 0      '有;医保用料目录,也可以更改
    
'    mnuFilePara.Visible = InStr(1, mstrPrivs, ";参数设置;") <> 0
'    mnuFileSpt2.Visible = mnuFilePara.Visible
    
    bln调价 = zlStr.IsHavePrivs(mstrPrivs, "售价管理") Or zlStr.IsHavePrivs(mstrPrivs, "成本价管理")

    mnuFileMlPreview.Visible = InStr(1, mstrPrivs, ";目录打印;") <> 0
    mnuFileMlPrint.Visible = InStr(1, mstrPrivs, ";目录打印;") <> 0
    
    '价格子菜单
    If InStr(1, mstrPrivs, ";查询卫材价目表;") = 0 And InStr(1, mstrPrivs, ";调价记录查询;") = 0 Then '子菜单都不显示，直接控制父菜单不显示
        mnuPrice.Visible = False
    Else
        mnuPriceLists.Visible = InStr(1, mstrPrivs, ";查询卫材价目表;") <> 0
        mnuPriceTable.Visible = InStr(1, mstrPrivs, ";调价记录查询;") <> 0
    End If
    
    mnuEditSubsection.Visible = InStr(1, mstrPrivs, ";分段加成率;") <> 0
    mnuEditSpt3.Visible = mnuEditSubsection.Visible
    
    mnuEditSpecLimit.Visible = InStr(1, mstrPrivs, ";上下限控制;") <> 0 Or InStr(1, mstrPrivs, ";盘点属性设置;") <> 0
    
    mnuEditItemAdd.Visible = InStr(1, mstrPrivs, ";卫材品名管理;") <> 0
    mnuEditItemMod.Visible = InStr(1, mstrPrivs, ";卫材品名管理;") <> 0 Or _
                InStr(1, mstrPrivs, ";管理扣率;") <> 0 Or _
                InStr(1, mstrPrivs, ";指导价格管理;") <> 0 Or _
                InStr(1, mstrPrivs, ";售价管理;") <> 0 Or _
                InStr(1, mstrPrivs, ";成本价管理;") <> 0 Or _
                InStr(1, mstrPrivs, ";医保用料目录;") <> 0 Or _
                InStr(1, mstrPrivs, ";服务对象;") <> 0
    mnuEditPrice.Visible = InStr(1, mstrPrivs, ";售价管理;") <> 0 And InStr(1, mstrPrivs, ";成本价管理;") <> 0
    mnuEditItemDel.Visible = mnuEditItemAdd.Visible
    
    mnuEditAddName.Visible = mnuEditItemAdd.Visible
    mnuEditDeleName.Visible = mnuEditItemAdd.Visible
    mnuEditModifyName.Visible = mnuEditItemAdd.Visible
    
    mnuEditSpt1.Visible = mnuEditItemAdd.Visible Or mnuEditItemMod.Visible
    mnuEditStop.Visible = mnuEditItemAdd.Visible
    mnuEditStart.Visible = mnuEditItemAdd.Visible
    mnuEditSpt2.Visible = mnuEditItemAdd.Visible
    
    tlbThis.Buttons("增加").Visible = mnuEditItemAdd.Visible
    tlbThis.Buttons("修改").Visible = blnModify    ' mnuEditItemMod.Visible
    tlbThis.Buttons("删除").Visible = mnuEditItemAdd.Visible
    
    tlbThis.Buttons("Split1").Visible = mnuEditItemAdd.Visible
    tlbThis.Buttons("Start").Visible = mnuEditItemAdd.Visible
    tlbThis.Buttons("Stop").Visible = mnuEditItemAdd.Visible
    
    
    mnuClass.Visible = InStr(1, mstrPrivs, ";卫材分类管理;") <> 0
    tlbThis.Buttons("Class").Visible = mnuClass.Visible
    tlbThis.Buttons("Split4").Visible = mnuClass.Visible
    
    tlbThis.Buttons("Split").Visible = mnuEditItemAdd.Visible Or mnuClass.Visible Or mnuEditItemMod.Visible
    tlbThis.Buttons("Split4").Visible = mnuEditItemAdd.Visible And mnuClass.Visible
    
    
End Sub
Private Sub InitHgdPrivceHeadcol()
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:初始价格列头
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    '价格表格设置
    With Me.vsPrice
        .Redraw = flexRDNone
        .Rows = 2
        For intCol = 0 To .Cols - 1
            .TextMatrix(1, intCol) = ""
            .Cell(flexcpData, 1, intCol) = ""
        Next
        .RowData(1) = 0
        .Redraw = flexRDBuffered
        .ExtendLastCol = True
    End With
    With Me.vsStuff
        .Redraw = flexRDNone
        .Rows = 2
        For intCol = 0 To .Cols - 1
            .TextMatrix(1, intCol) = ""
            .Cell(flexcpData, 1, intCol) = ""
        Next
                 
        .RowData(1) = 0
        .Redraw = flexRDBuffered
        .ExtendLastCol = True
    End With
    vsCost.ExtendLastCol = True
End Sub
Private Sub Form_Resize()
    Dim lngTools As Single, lngStatus As Single
    
    SetParent txtFind.hwnd, tlbThis.hwnd
    SetParent picFind.hwnd, tlbThis.hwnd
    txtFind.Left = Me.ScaleWidth - txtFind.Width - 200
    picFind.Left = txtFind.Left - 100 - picFind.Width
    
    If WindowState = 1 Then Exit Sub
    lngTools = IIf(Me.clbThis.Visible, Me.clbThis.Height, 0)
    lngStatus = IIf(Me.stbThis.Visible, Me.stbThis.Height, 0)
    
    err = 0: On Error Resume Next
    
    With Me.picVBar_S
        .Top = lngTools
        .Height = Me.ScaleHeight - lngTools - lngStatus
        If .Left < 2000 Then .Left = 2000
        If .Left > Me.ScaleWidth - 4000 Then .Left = Me.ScaleWidth - 4000
    End With
    With Me.picHBar_S
        .Left = Me.picVBar_S.Left + Me.picVBar_S.Width
        .Width = Me.ScaleWidth - .Left
        If .Top < 2000 Then .Top = 2000
        If .Top > Me.ScaleHeight - lngStatus - 2500 Then .Top = Me.ScaleHeight - lngStatus - 2500
    End With
    
    With Me.picClass
        .Left = Me.ScaleLeft
        .Top = lngTools
        .Height = Me.ScaleHeight - picClass.Top - lngStatus
        .Width = Me.picVBar_S.Left - Me.picClass.Left
    End With
    
    With Me.lvwItems
        .Left = Me.picVBar_S.Left + Me.picVBar_S.Width
        .Top = lngTools
        .Height = Me.picHBar_S.Top - .Top
        .Width = Me.ScaleWidth - .Left
    End With
    
    With Me.picUp_S
        .Left = Me.picVBar_S.Left + Me.picVBar_S.Width
        .Top = Me.picHBar_S.Top + Me.picHBar_S.Height
        '.Height = Me.ScaleHeight - lngStatus - .Top + 15
        .Width = Me.ScaleWidth - .Left + 15
    End With
    With Me.TbList
        .Left = picUp_S.Left
        .Top = picUp_S.Height + picUp_S.Top + 10
        .Height = ScaleHeight - .Top - lngStatus
        .Width = ScaleWidth - .Left
    End With
    
    
'    With Me.fraComment
'        .Left = picUp_S.Left + 20
'        .Width = ScaleWidth - .Left
'        .Top = ScaleHeight - .Height - lngStatus
'    End With
'
'    With Me.vsPrice
'        .Left = picUp_S.Left
'        .Top = picUp_S.Height + picUp_S.Top + 10
'        .Width = picUp_S.Width - 30
'        .Height = fraComment.Top - .Top - 30
'    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
    Call SaveListViewState(Me.lvwItems, Me.Name & Val(Me.tvwClass.Tag), Me.lvwItems.View)
    
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\分割", "横向", Me.picVBar_S.Left)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\分割", "纵向", Me.picHBar_S.Top)
    
    Call zlDatabase.SetPara("包含下级卫材", IIf(mnuViewDownLevel.Checked = True, 1, 0), glngSys, mlngModule)
    zl_vsGrid_Para_Save mlngModule, vsCost, Me.Caption, "成本价调价记录"
    zl_vsGrid_Para_Save mlngModule, vsPrice, Me.Caption, "售价调价记录"
    zl_vsGrid_Para_Save mlngModule, vsStuff, Me.Caption, "材料规格记录"
    mstrFindValue = ""
    Set mrsFind = Nothing
End Sub



Private Sub mnuEditAddName_Click()
    '增加品种
    Dim lng分类id As Long
    If Me.tvwClass.SelectedItem Is Nothing Then
        ShowMsgBox "尚未设置分类,不能增删卫生材料！"
        Exit Sub
    End If
    If tvwClass.SelectedItem Is Nothing Then
        lng分类id = 0
    Else
        lng分类id = Val(Mid(tvwClass.SelectedItem.Key, 2))
    End If
    If frmStuffBreed.ShowEditCard(Me, g新增, "", lng分类id, mstrPrivs) = False Then
        Exit Sub
    End If
    If lvwItems.SelectedItem Is Nothing Then
        Call zlRefRecords
    Else
        Call zlRefRecords(Val(Mid(lvwItems.SelectedItem.Key, 2)))
    End If
    Call SetMenuEnable
End Sub

Private Sub mnuEditDeleName_Click()
    Dim lng诊疗ID As Long
    Dim rsSpec As New ADODB.Recordset
    On Error GoTo ErrHand
    
    With Me.lvwItems
        If .SelectedItem Is Nothing Then Exit Sub
        If MsgBox("真的删除品种为“" & .SelectedItem.Text & "”吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
        lng诊疗ID = Mid(.SelectedItem.Key, 2)
        
        '参数为:材料id
        gstrSQL = "Zl_材料品种_Delete(" & lng诊疗ID & ")"
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
        
        Call .ListItems.Remove(.SelectedItem.Key)
        If .SelectedItem Is Nothing Then
            Call InitHgdPrivceHeadcol
        Else
            Call lvwItems_ItemClick(.SelectedItem)
        End If
    End With
    Call SetMenuEnable
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditExcel_Click()
    frmItemImport.ShowMe 3, Me
    
    Call zlRefClasses
    Call zlRefRecords
End Sub

Private Sub mnuEditModifyName_Click()

    Dim lng分类id As Long, lng诊疗ID As Long
    If lvwItems.SelectedItem Is Nothing Then Exit Sub
    lng诊疗ID = Val(Mid(lvwItems.SelectedItem.Key, 2))
    If lng诊疗ID = 0 Then Exit Sub
    lng分类id = Val(lvwItems.SelectedItem.Tag)
    
    If Me.lvwItems.SelectedItem.Icon = "stop" Then
        ShowMsgBox "不能对停用卫生材料品种进行修改！"
        Exit Sub
    End If
    
    If frmStuffBreed.ShowEditCard(Me, g修改, lng诊疗ID, lng分类id, mstrPrivs) = False Then
        Exit Sub
    End If
    Call zlRefRecords(lng诊疗ID)
    Call SetMenuEnable

End Sub

Private Sub mnuFilter_Click()
    With FrmStuffFilter
        Call .ShowMe(Me, mnuViewStoped.Checked)
    End With
End Sub

Public Sub zlGetFilter(ByVal strMaterialId As String)
    mstrMaterialID = strMaterialId
    Call cmdKind_Click(1)
End Sub
Private Sub picClass_Resize()
    Dim intCount As Integer
    err = 0: On Error Resume Next
    For intCount = Me.cmdKind.LBound To Me.cmdKind.UBound
        Me.cmdKind(intCount).Left = Me.picClass.ScaleLeft + 15
        Me.cmdKind(intCount).Width = Me.picClass.ScaleWidth
        Me.cmdKind(intCount).Height = 300
        If Val(Me.cmdKind(intCount).Tag) = 0 Then
            Me.cmdKind(intCount).Top = Me.picClass.ScaleTop + 285 * intCount
            Me.tvwClass.Top = Me.picClass.ScaleTop + 285 * (intCount + 1)
        Else
            Me.cmdKind(intCount).Top = Me.picClass.ScaleHeight - 285 * (Me.cmdKind.UBound - intCount + 1)
        End If
    Next
    Me.tvwClass.Left = Me.picClass.ScaleLeft + 15
    Me.tvwClass.Width = Me.picClass.ScaleWidth
    Me.tvwClass.Height = Me.picClass.ScaleHeight - 285 * (Me.cmdKind.UBound + 1) - 15
End Sub


Private Sub picCost_Resize()
    err = 0: On Error Resume Next
    With picCost
        vsCost.Move .ScaleLeft, .ScaleTop, .ScaleWidth, .ScaleHeight
    End With
End Sub

Private Sub TbList_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Select Case Item.Index
    Case 0
        zlControl.ControlSetFocus vsStuff, True
    Case 1
        zlControl.ControlSetFocus vsPrice, True
    Case Else
        zlControl.ControlSetFocus vsCost, True
    End Select
End Sub

Private Sub txtFind_GotFocus()
    zlControl.TxtSelAll txtFind
    OS.OpenIme True
End Sub

Private Sub txtfind_KeyPress(KeyAscii As Integer)
    Dim strTag As String
    Dim strSearch As String
        
    On Error GoTo ErrHandle
    If KeyAscii = vbKeyReturn Then
        zlControl.TxtSelAll txtFind
        If txtFind.Text = "" Then Exit Sub
        If mstrFindValue <> txtFind.Text And txtFind.Text <> "" Then
            mstrFindValue = txtFind.Text
            Set mrsFind = Nothing
            
            strSearch = " And (C.编码 Like [1] OR N.名称 Like [1] OR N.简码 LIKE upper([1]))"
            gstrSQL = "" & _
                    "   Select distinct I.分类ID,B.材料ID,B.诊疗ID" & _
                    "   From 诊疗项目目录 I,收费项目别名 N,材料特性 B,收费项目目录 C" & _
                    "   Where   I.类别='4' And I.id=b.诊疗id and b.材料ID=N.收费细目id and b.材料id=C.id " & strSearch
            If mnuViewStoped.Checked = False Then gstrSQL = gstrSQL & " And (C.撤档时间 Is NULL Or C.撤档时间 >=to_date('3000-01-01 00:00:00','yyyy-mm-dd hh24:mi:ss'))"
            Set mrsFind = zlDatabase.OpenSQLRecord(gstrSQL, "材料查询", UCase(Trim(txtFind.Text)) & "%")
                    
            If mrsFind.RecordCount > 0 Then
                Call zlLocateItem(mrsFind!分类id, mrsFind!诊疗id, mrsFind!材料ID)
            End If
        Else
            If Not mrsFind.EOF Then
                mrsFind.MoveNext
                If Not mrsFind.EOF Then
                    Call zlLocateItem(mrsFind!分类id, mrsFind!诊疗id, mrsFind!材料ID)
                Else
                    MsgBox "已查询到最后一条记录！", vbInformation, gstrSysName
                    mrsFind.MoveFirst
                    Call zlLocateItem(mrsFind!分类id, mrsFind!诊疗id, mrsFind!材料ID)
                End If
            ElseIf mrsFind.RecordCount <> 0 And mrsFind.EOF Then
                mrsFind.MoveFirst
                Call zlLocateItem(mrsFind!分类id, mrsFind!诊疗id, mrsFind!材料ID)
            End If
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsCost_DblClick()
    Call ShowCostBill
    
End Sub

Private Sub vsCost_GotFocus()
    With vsCost
         .SelectionMode = flexSelectionByRow
         .BackColorSel = &H8000000D
    End With
End Sub

Private Sub vsCost_LostFocus()
    With vsCost
         .BackColorSel = GRD_LOSTFOCUS_COLORSEL
    End With
End Sub

Private Sub ShowCostBill()
    '-----------------------------------------------------------------------------------------------------------
    '功能:显示成本价调成金额
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-11-11 16:57:43
    '-----------------------------------------------------------------------------------------------------------
    Dim strNo As String
    
    With vsCost
        strNo = Trim(.TextMatrix(.Row, .ColIndex("NO")))
        If strNo = "" Then Exit Sub
        frmDiffPriceAdjustCard.ShowCard Me, strNo, 4, 1
    End With
    
    
End Sub


Private Sub vsPrice_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    If mnuPrice.Visible = False Then Exit Sub
    Call PopupMenu(Me.mnuPrice, 2)
End Sub

Private Sub vsPrice_RowColChange()
    
    Dim rsCheck As New ADODB.Recordset
    With Me.vsPrice
        Me.lblComment(3).Caption = "1、" & _
            .TextMatrix(.Row, .ColIndex("材料类型")) & "卫生材料，" & _
             "指导批价" & .TextMatrix(.Row, .ColIndex("指导批价")) & "元/" & .TextMatrix(.Row, colPrice.单位) & "，" & _
            "采购扣率" & .TextMatrix(.Row, .ColIndex("扣率")) & "%；"

            Me.lblComment(4).Caption = "2、" & _
            "指导售价" & .TextMatrix(.Row, .ColIndex("指导售价")) & "元/" & .TextMatrix(.Row, colPrice.单位) & "，" & _
            "指导差率" & .TextMatrix(.Row, .ColIndex("指导差价率")) & "%；" & _
            IIf(Val(.TextMatrix(.Row, .ColIndex("屏蔽费别"))) = 0, "根据病人身份费别进行优惠或加价。", "不受病人身份费别影响。")
    End With
End Sub
Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwItems.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwItems.SortOrder = IIf(Me.lvwItems.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwItems.SortKey = ColumnHeader.Index - 1
        Me.lvwItems.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwItems_DblClick()
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call frmStuffBreed.ShowEditCard(Me, g查看, Mid(Me.lvwItems.SelectedItem.Key, 2), , mstrPrivs)
End Sub

Private Function Load价格信息(ByVal lng诊疗ID As Long) As Boolean
    '--------------------------------------------------------------------------------
    '功能: 加载卫生材料的价格信息
    '参数:lng诊疗ID-诊疗ID
    '返回:成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/05/24
    '--------------------------------------------------------------------------------
    err = 0: On Error GoTo ErrHand
    
    For intCount = Me.lblComment.LBound To Me.lblComment.UBound
        Me.lblComment(intCount).Caption = ""
    Next
    '----------填写价格-----------------
    gstrSQL = " Select P.ID,p.NO ,p.序号,decode(I.是否变价,1,'时价','定价') as 类型,  " & _
             "      nvl(S.指导批发价,0) as 指导批价,nvl(S.扣率,0) as 扣率, " & _
             "      nvl(S.指导零售价,0) as 指导售价,nvl(S.指导差价率,0) as 指导差率,nvl(I.屏蔽费别,0)  as 屏蔽费别, " & _
             "      decode(sign(P.执行日期-sysdate),1,1,decode(sign(P.终止日期-sysdate),-1,-1,0)) as 执行情况, " & _
             "      '['||I.编码||']'||I.名称||' '||I.规格||' '||I.产地 as 材料,I.计算单位 as 单位,S.包装单位,S.换算系数,S.材料ID, " & _
             "      P.现价 as 售价,U.名称 as 收入项目,P.调价说明, " & _
             "      to_char(P.执行日期,'YYYY-MM-DD HH24:MI:SS') as 执行日期 " & _
             "   From 收费价目 P,收入项目 U,收费项目目录 I,材料特性 S" & _
             "   Where P.收费细目ID=I.ID and P.收入项目ID=U.ID and I.ID=S.材料ID" & _
             "       and S.诊疗ID=[1]" & GetPriceClassString("P")
    
    If Me.mnuViewPrices.Checked = False Then
        gstrSQL = gstrSQL & "       and (P.终止日期 is null or P.终止日期>=sysdate)"
    End If
    gstrSQL = gstrSQL & " order by I.编码,P.执行日期 desc"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng诊疗ID)

    With rsTemp
        Me.vsPrice.Redraw = flexRDNone
        If .BOF Or .EOF Then
            With Me.vsPrice
                .Rows = .FixedRows + 1: .RowData(.FixedRows) = 0
                For intCol = 0 To .Cols - 1
                    .TextMatrix(.FixedRows, intCol) = ""
                Next
            End With
        Else
            Me.vsPrice.Rows = Me.vsPrice.FixedRows + .RecordCount
        End If

        Do While Not .EOF
            Me.vsPrice.RowData(.AbsolutePosition) = Val(zlStr.Nvl(!Id))
            Me.vsPrice.TextMatrix(.AbsolutePosition, vsPrice.ColIndex("NO")) = zlStr.Nvl(!NO)
            Me.vsPrice.TextMatrix(.AbsolutePosition, vsPrice.ColIndex("序号")) = zlStr.Nvl(!序号)
            Me.vsPrice.TextMatrix(.AbsolutePosition, vsPrice.ColIndex("材料类型")) = zlStr.Nvl(!类型)
            Me.vsPrice.TextMatrix(.AbsolutePosition, vsPrice.ColIndex("指导批价")) = Format(!指导批价 * IIf(mintUnit = 0, 1, zlStr.Nvl(!换算系数, 1)), mFMT.FM_成本价)
            Me.vsPrice.TextMatrix(.AbsolutePosition, vsPrice.ColIndex("扣率")) = Format(!扣率, GFM_VBKL)
            Me.vsPrice.TextMatrix(.AbsolutePosition, vsPrice.ColIndex("指导售价")) = Format(!指导售价 * IIf(mintUnit = 0, 1, zlStr.Nvl(!换算系数, 1)), mFMT.FM_零售价)
            Me.vsPrice.TextMatrix(.AbsolutePosition, vsPrice.ColIndex("指导差价率")) = Format(!指导差率, GFM_VBCJL)

            Me.vsPrice.TextMatrix(.AbsolutePosition, vsPrice.ColIndex("执行情况")) = zlStr.Nvl(!执行情况)
            Me.vsPrice.TextMatrix(.AbsolutePosition, vsPrice.ColIndex("材料")) = zlStr.Nvl(!材料)
            Me.vsPrice.Cell(flexcpData, .AbsolutePosition, vsPrice.ColIndex("材料")) = zlStr.Nvl(!材料ID)
            Me.vsPrice.TextMatrix(.AbsolutePosition, vsPrice.ColIndex("单位")) = IIf(mintUnit = 0, zlStr.Nvl(!单位), zlStr.Nvl(!包装单位))
            Me.vsPrice.TextMatrix(.AbsolutePosition, vsPrice.ColIndex("售价")) = Format(!售价 * IIf(mintUnit = 0, 1, zlStr.Nvl(!换算系数, 1)), mFMT.FM_零售价)

            Me.vsPrice.TextMatrix(.AbsolutePosition, vsPrice.ColIndex("收入项目")) = zlStr.Nvl(!收入项目)
            Me.vsPrice.TextMatrix(.AbsolutePosition, vsPrice.ColIndex("说明")) = zlStr.Nvl(!调价说明)
            Me.vsPrice.TextMatrix(.AbsolutePosition, vsPrice.ColIndex("执行日期")) = zlStr.Nvl(!执行日期)
            Me.vsPrice.TextMatrix(.AbsolutePosition, vsPrice.ColIndex("屏蔽费别")) = zlStr.Nvl(!屏蔽费别)

            Me.vsPrice.Row = .AbsolutePosition
            For intCol = 0 To Me.vsPrice.Cols - 1
                Me.vsPrice.Col = intCol
                Select Case !执行情况
                Case -1     '已执行
                    Me.vsPrice.CellBackColor = RGB(240, 240, 240)
                Case 0      '正在执行
                    Me.vsPrice.CellBackColor = RGB(255, 255, 255)
                Case 1      '未执行
                    Me.vsPrice.CellBackColor = RGB(225, 255, 255)
                End Select
            Next
            .MoveNext
        Loop

        Me.vsPrice.Row = Me.vsPrice.FixedRows
        If Me.vsPrice.ColWidth(vsPrice.ColIndex("材料")) = 0 Or Me.vsPrice.ColWidth(vsPrice.ColIndex("单位")) = 0 Then
            Me.vsPrice.ColWidth(vsPrice.ColIndex("材料")) = 3500
            Me.vsPrice.ColWidth(vsPrice.ColIndex("单位")) = 550
        End If
        Me.vsPrice.Redraw = flexRDBuffered
    End With
    Call vsPrice_RowColChange
    Load价格信息 = True
    Exit Function
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub SetVsGridColor(ByVal objVsGrid As Object, ByVal lngRow As Long, ByVal OleColor As OLE_COLOR, ByVal blnBackColor As Boolean)
    '--------------------------------------------------------------------------------
    '功能: 设置指定行的背景或前景色
    '参数:objVsGrid-指定的网格控件
    '     lngRow -指定的行
    '     olecolor-指定的颜色值
    '     blnBackColor-设置背景色
    '编制:刘兴宏
    '日期:2007/05/24
    '--------------------------------------------------------------------------------
    Dim lngCol As Long
    Dim lngCurRow As Long, lngCurCol As Long
    With objVsGrid
        lngCurRow = .Row: lngCurCol = .Col
        For intCol = 0 To .Cols - 1
            .Col = intCol
            If blnBackColor = False Then
                .CellForeColor = OleColor
            Else
                .CellBackColor = OleColor
            End If
        Next
        .Row = lngCurRow: .Col = lngCurCol
    End With
End Sub
Private Function Load材料规格信息(ByVal lng诊疗ID As Long) As Boolean
    '--------------------------------------------------------------------------------
    '功能: 加载卫生材料的规格信息
    '参数:lng诊疗ID-诊疗ID
    '返回:成功,返回true,否则返回False
    '编制:刘兴宏
    '日期:2007/05/24
    '--------------------------------------------------------------------------------
    Dim lngRow As Long
    Dim lngPre材料ID As Long, lngLocaleRow As Long
    err = 0: On Error GoTo ErrHand:
    gstrSQL = " " & _
         " SELECT  Distinct b.id,a.诊疗id, b.编码,b.名称,b.规格,b.产地,b.计算单位 AS 散装单位,a.包装单位, a.换算系数,b.服务对象,b.费用类型,a.成本价," & _
         "      a.材质分类,a.存储条件,a.许可证号,a.许可证有效期,b.标识主码,b.标识子码,decode(b.是否变价,1,'实价','定价') as 是否变价," & _
         "      a.最大效期,a.灭菌效期,decode(a.招标材料,1,'是','否') as 招标材料,decode(a.无菌性材料,1,'是','否')  无菌性材料, " & _
         "      decode(a.一次性材料,1,'是','否') 一次性材料  ,  decode(a.自制材料,1,'是','否') 自制材料 ,a.货源情况,a.材料来源, " & _
         "      a.指导批发价,a.指导零售价,a.指导差价率,a.扣率, decode(a.库房分批,1,'是','否') 库房分批,decode(a.在用分批,1,'是','否') 在用分批," & _
         "      decode(A.原材料,1,'是','否') 原材料,  decode(a.跟踪在用,1,'是','否') 跟踪在用,decode(a.核算材料,1,'是','否') 核算材料," & _
         "      to_Char(b.建档时间,'yyyy-mm-dd') as  建档时间," & _
         "      nvl(b.撤档时间,to_date('3000-01-01','YYYY-MM-DD')) as 撤档时间,C.名称 As 商品名, A.高值材料 " & _
         " FROM 材料特性 a,收费项目目录 b, 收费项目别名 C " & _
         " WHERE a.材料id=b.id  And B.ID = C.收费细目id(+) And C.性质(+) = 3 and a.诊疗id=[1] "
 
    If Me.mnuViewStoped.Checked = False Then
        gstrSQL = gstrSQL & " and (B.撤档时间 is null or B.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
    End If
    gstrSQL = gstrSQL & " order by B.编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng诊疗ID)
    
    lngLocaleRow = -1
    With vsStuff
        .Redraw = flexRDNone
        lngPre材料ID = .RowData(.Row)
        If rsTemp.EOF Then
            .Rows = .FixedRows + 1: .RowData(.FixedRows) = 0
            For intCol = 0 To .Cols - 1
                .TextMatrix(.FixedRows, intCol) = ""
            Next
        Else
            .Rows = .FixedRows + rsTemp.RecordCount
        End If
        
        lngRow = 1
        Do While Not rsTemp.EOF
            .RowData(lngRow) = Val(zlStr.Nvl(rsTemp!Id))
            
            .Cell(flexcpData, lngRow, .ColIndex("编码")) = zlStr.Nvl(rsTemp!诊疗id)
            .TextMatrix(lngRow, .ColIndex("ID")) = zlStr.Nvl(rsTemp!Id)
            .TextMatrix(lngRow, .ColIndex("编码")) = zlStr.Nvl(rsTemp!编码)
            .TextMatrix(lngRow, .ColIndex("规格")) = zlStr.Nvl(rsTemp!规格)
            .TextMatrix(lngRow, .ColIndex("产地")) = zlStr.Nvl(rsTemp!产地)
            .TextMatrix(lngRow, .ColIndex("商品名")) = zlStr.Nvl(rsTemp!商品名)
            .TextMatrix(lngRow, .ColIndex("散装单位")) = zlStr.Nvl(rsTemp!散装单位)
            .TextMatrix(lngRow, .ColIndex("包装单位")) = zlStr.Nvl(rsTemp!包装单位)
            .TextMatrix(lngRow, .ColIndex("换算系数")) = zlStr.Nvl(rsTemp!换算系数)
            If zlStr.Nvl(rsTemp!服务对象) = 1 Then
                .TextMatrix(lngRow, .ColIndex("服务对象")) = "门诊"
            ElseIf zlStr.Nvl(rsTemp!服务对象) = 2 Then
                .TextMatrix(lngRow, .ColIndex("服务对象")) = "住院"
            ElseIf zlStr.Nvl(rsTemp!服务对象) = 3 Then
                .TextMatrix(lngRow, .ColIndex("服务对象")) = "门诊和住院"
            Else
                .TextMatrix(lngRow, .ColIndex("服务对象")) = "不应用于病人"
            End If
            .TextMatrix(lngRow, .ColIndex("医保类型")) = zlStr.Nvl(rsTemp!费用类型)
            .TextMatrix(lngRow, .ColIndex("材质分类")) = zlStr.Nvl(rsTemp!材质分类)
            .TextMatrix(lngRow, .ColIndex("存储条件")) = zlStr.Nvl(rsTemp!存储条件)
            .TextMatrix(lngRow, .ColIndex("许可证号")) = zlStr.Nvl(rsTemp!许可证号)
            .TextMatrix(lngRow, .ColIndex("许可证效期")) = zlStr.Nvl(rsTemp!许可证有效期)
            .TextMatrix(lngRow, .ColIndex("标识主码")) = zlStr.Nvl(rsTemp!标识主码)
            .TextMatrix(lngRow, .ColIndex("标识子码")) = zlStr.Nvl(rsTemp!标识子码)
            .TextMatrix(lngRow, .ColIndex("最大效期")) = zlStr.Nvl(rsTemp!最大效期)
            .TextMatrix(lngRow, .ColIndex("灭菌效期")) = zlStr.Nvl(rsTemp!灭菌效期)
            .TextMatrix(lngRow, .ColIndex("招标材料")) = zlStr.Nvl(rsTemp!招标材料)

            .TextMatrix(lngRow, .ColIndex("无菌材料")) = zlStr.Nvl(rsTemp!无菌性材料)
            .TextMatrix(lngRow, .ColIndex("一次性材料")) = zlStr.Nvl(rsTemp!一次性材料)
            .TextMatrix(lngRow, .ColIndex("自制材料")) = zlStr.Nvl(rsTemp!自制材料)
            .TextMatrix(lngRow, .ColIndex("货源情况")) = zlStr.Nvl(rsTemp!货源情况)
            .TextMatrix(lngRow, .ColIndex("材料来源")) = zlStr.Nvl(rsTemp!材料来源)
            .TextMatrix(lngRow, .ColIndex("指导批价")) = Format(rsTemp!指导批发价 * IIf(mintUnit = 0, 1, zlStr.Nvl(rsTemp!换算系数, 1)), mFMT.FM_成本价)
            .TextMatrix(lngRow, .ColIndex("成本价")) = Format(rsTemp!成本价 * IIf(mintUnit = 0, 1, zlStr.Nvl(rsTemp!换算系数, 1)), mFMT.FM_成本价)
            
            .TextMatrix(lngRow, .ColIndex("指导售价")) = Format(rsTemp!指导零售价 * IIf(mintUnit = 0, 1, zlStr.Nvl(rsTemp!换算系数, 1)), mFMT.FM_零售价)
            .TextMatrix(lngRow, .ColIndex("指导差价率")) = Format(rsTemp!指导差价率, GFM_VBCJL)
            .TextMatrix(lngRow, .ColIndex("扣率")) = Format(rsTemp!扣率, GFM_VBKL)
            .TextMatrix(lngRow, .ColIndex("库房分批")) = zlStr.Nvl(rsTemp!库房分批)
            .TextMatrix(lngRow, .ColIndex("在用分批")) = zlStr.Nvl(rsTemp!在用分批)
            .TextMatrix(lngRow, .ColIndex("原料材料")) = zlStr.Nvl(rsTemp!原材料)
            .TextMatrix(lngRow, .ColIndex("跟踪在用")) = zlStr.Nvl(rsTemp!跟踪在用)
            .TextMatrix(lngRow, .ColIndex("核算材料")) = zlStr.Nvl(rsTemp!核算材料)
            .TextMatrix(lngRow, .ColIndex("高值材料")) = IIf(IsNull(rsTemp!高值材料) Or rsTemp!高值材料 = "0", "否", "是")
            .TextMatrix(lngRow, .ColIndex("建档时间")) = zlStr.Nvl(rsTemp!建档时间)
            If .RowData(lngRow) = lngPre材料ID Then
                lngLocaleRow = lngRow
            End If
            .Row = lngRow
            If Format(rsTemp!撤档时间, "YYYY-MM-DD") <> "3000-01-01" Then
                .TextMatrix(lngRow, .ColIndex("撤档时间")) = zlStr.Nvl(rsTemp!撤档时间)
                Call SetVsGridColor(vsStuff, lngRow, &HFF&, False)
            Else
                .TextMatrix(lngRow, .ColIndex("撤档时间")) = ""
                '如果是招标材料，用颜色区分是否是时价还是定价材料
                If Trim(zlStr.Nvl(rsTemp!招标材料)) = "是" Then
                    Call SetVsGridColor(vsStuff, lngRow, IIf(rsTemp!是否变价 = "定价", &H800000, &H800080), False)
                Else
                    Call SetVsGridColor(vsStuff, lngRow, IIf(rsTemp!是否变价 = "定价", &H0, &H40&), False)
                End If
            End If
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If lngLocaleRow > 0 Then
            .Row = lngLocaleRow
        Else
            .Row = 1
        End If
        Call .ShowCell(.Row, .ColIndex("编码"))
        .Redraw = flexRDBuffered
    End With
   Load材料规格信息 = True
   Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub lvwItems_ItemClick(ByVal Item As MSComctlLib.ListItem)
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    Dim lng诊疗ID As Long
    
    err = 0: On Error GoTo ErrHand
    lng诊疗ID = Val(Mid(Item.Key, 2))
    Call Load价格信息(lng诊疗ID)
    Call LoadCostData(lng诊疗ID)
    '加载规格信息:
    Call Load材料规格信息(lng诊疗ID)
    Call SetMenuEnable
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub lvwItems_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call lvwItems_DblClick
End Sub

Private Sub lvwItems_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
   
   PopupMenu mnuEdit
End Sub
Private Sub mnuClassAdd_Click()
    With frmClinicClass
        .lblKind.Tag = 7
        If Me.tvwClass.SelectedItem Is Nothing Then
            .txtParent.Tag = 0
        Else
            .txtParent.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
        End If
        .Tag = "增加"
        .Show 1, Me
    End With
    
    Call zlRefClasses
End Sub

Private Sub mnuClassDel_Click()
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    If MsgBox("真的删除该分类“" & Me.tvwClass.SelectedItem.Text & "”吗", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    err = 0: On Error GoTo ErrHand
    gstrSQL = "zl_诊疗分类目录_delete(" & Mid(Me.tvwClass.SelectedItem.Key, 2) & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
                    
    
    Dim strParentKey As String
    If Me.tvwClass.SelectedItem.Next Is Nothing Then
        If Me.tvwClass.SelectedItem.Parent Is Nothing Then
            Call zlRefClasses
        Else
            strParentKey = Me.tvwClass.SelectedItem.Parent.Key
            Call Me.tvwClass.Nodes.Remove(Me.tvwClass.SelectedItem.Key)
            If Me.tvwClass.SelectedItem Is Nothing Then
                    Me.lvwItems.ListItems.Clear
                    Call InitHgdPrivceHeadcol
            Else
                Call zlRefRecords
            End If
        End If
    Else
        Call zlRefClasses
    End If
    SetMenuEnable
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuClassMod_Click()

    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    With frmClinicClass
        .lblKind.Tag = 7
        If Me.tvwClass.SelectedItem.Parent Is Nothing Then
            .txtParent.Tag = 0
            .txtParent.Text = "(无)"
            .txtUpCode.Text = ""
            .txtCode.Text = Mid(Split(Me.tvwClass.SelectedItem.Text, "]")(0), 2)
            .txtCode.MaxLength = Len(.txtCode.Text)
            .txtCode.Tag = .txtCode.MaxLength
        Else
            .txtParent.Tag = Mid(Me.tvwClass.SelectedItem.Parent.Key, 2)
            .txtParent.Text = Me.tvwClass.SelectedItem.Parent.Text
            .txtUpCode.Text = Mid(Split(Me.tvwClass.SelectedItem.Parent.Text, "]")(0), 2)
            .txtCode.Text = Mid(Split(Me.tvwClass.SelectedItem.Text, "]")(0), Len(.txtUpCode.Text) + 2)
            .txtCode.MaxLength = Len(.txtCode.Text)
            .txtCode.Tag = .txtCode.MaxLength
        End If
        .txtName = Split(Me.tvwClass.SelectedItem.Text, "]")(1)
        .txtSymbol = Me.tvwClass.SelectedItem.Tag
        .Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
        .Show 1, Me
    End With
    Call zlRefClasses
End Sub

Private Sub mnuEditItemAdd_Click()
    Dim lng诊疗ID As Long
    Dim blnStop As Boolean
    Dim lng分类id As Long
    
    If Me.lvwItems.ListItems.count = 0 Then
        ShowMsgBox "尚未设置品种,不能增删卫生材料的规格！"
        Exit Sub
    End If
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    
    blnStop = Me.lvwItems.SelectedItem.Icon = "stop"
    If blnStop Then
        ShowMsgBox "该品种已经被停用,不能增加规格!"
        Exit Sub
    End If
    
    lng诊疗ID = Val(Mid(Me.lvwItems.SelectedItem.Key, 2))
    
    If tvwClass.SelectedItem Is Nothing Then
        lng分类id = 0
    Else
        lng分类id = Val(Mid(tvwClass.SelectedItem.Key, 2))
    End If
    
    If frmStuffSpec.ShowEditCard(Me, g新增, lng诊疗ID, lng分类id, "", mstrPrivs) = False Then
        Exit Sub
    End If
    Call Load材料规格信息(Val(Mid(lvwItems.SelectedItem.Key, 2)))
    Call SetMenuEnable
    
End Sub

Private Sub mnuEditItemDel_Click()
    Dim lng诊疗ID As Long
    Dim lng材料ID As Long
    Dim blnStop As Boolean
    Dim str规格 As String
    With vsStuff
          lng材料ID = .RowData(.Row)
          lng诊疗ID = Val(.Cell(flexcpData, .Row, .ColIndex("编码")))
          blnStop = Trim(.TextMatrix(.Row, .ColIndex("撤档时间"))) <> ""
          str规格 = Trim(.TextMatrix(.Row, .ColIndex("规格")))
    End With
    If lng材料ID = 0 Then Exit Sub
    If lng诊疗ID = 0 Then Exit Sub
    
    On Error GoTo ErrHand
    If MsgBox("真的删除规格为“" & str规格 & "”吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
        
    '参数为:材料id
    gstrSQL = "zl_卫生材料_DELETE(" & lng材料ID & ")"
    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
    With vsStuff
        If .Rows = 2 Then
            Call InitHgdPrivceHeadcol
        Else
            .RemoveItem .Row
        End If
    End With
    Call SetMenuEnable
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditItemMod_Click()
    Dim lng诊疗ID As Long
    Dim lng材料ID As Long
    Dim blnStop As Boolean
    Dim lng分类id As Long
    
    With vsStuff
          lng材料ID = .RowData(.Row)
          lng诊疗ID = Val(.Cell(flexcpData, .Row, .ColIndex("编码")))
          blnStop = Trim(.TextMatrix(.Row, .ColIndex("撤档时间"))) <> ""
    End With
    If lng材料ID = 0 Then Exit Sub
    If lng诊疗ID = 0 Then Exit Sub
    If blnStop Then
        ShowMsgBox "不能对停用卫生材料进行修改！"
        Exit Sub
    End If
    If tvwClass.SelectedItem Is Nothing Then
        lng分类id = 0
    Else
        lng分类id = Val(Mid(tvwClass.SelectedItem.Key, 2))
    End If
    If frmStuffSpec.ShowEditCard(Me, g修改, lng诊疗ID, lng分类id, CStr(lng材料ID), mstrPrivs) = False Then
        Exit Sub
    End If
    If Me.lvwItems.SelectedItem Is Nothing Then
        Call InitHgdPrivceHeadcol
    Else
        Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
    End If
    Call SetMenuEnable
    
End Sub
Private Function Get分类ID() As Long

    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    gstrSQL = "Select 分类id From 诊疗项目目录 where id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(Me.lvwItems.SelectedItem.Tag))
    If rsTemp.EOF Then
        Get分类ID = 0
    Else
        Get分类ID = Val(zlStr.Nvl(rsTemp!分类id))
    End If
    rsTemp.Close
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub mnuEditItemPart_Click()
    Dim bln编辑 As Boolean, lng材料ID As Long
    With vsStuff
        lng材料ID = .RowData(.Row)
    End With
    With frm存储库房
        bln编辑 = (InStr(1, mstrPrivs, ";存储库房;") <> 0)
        Call .ShowMe(Me, lng材料ID, bln编辑)
    End With
End Sub



Private Sub mnuEditSpecLimit_Click()
    With frmStuffLimit
        If InStr(1, mstrPrivs, ";上下限控制;") = 0 And InStr(1, mstrPrivs, ";盘点属性设置;") = 0 Then
            .cmdClose.Tag = "查阅"
        Else
            .cmdClose.Tag = "修改"
        End If
        .strPrivs = mstrPrivs
        .Show 1, Me
    End With
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
End Sub

Private Sub mnuEditSpecSelf_Click()
    Dim lng材料ID As Long
    
    With vsStuff
        lng材料ID = .RowData(.Row)
    End With
'
'    If Me.lvwItems.SelectedItem Is Nothing Then
'        lng材料id = 0
'    Else
'        lng材料id = Val(Mid(Me.lvwItems.SelectedItem.Key, 2))
'    End If

    With frmStuffMember
        If InStr(1, mstrPrivs, ";自制材料构成;") = 0 Then
            .cmdClose.Tag = "查阅"
        Else
            .cmdClose.Tag = "修改"
        End If
          
        .lblMedi.Tag = lng材料ID
        .msfMember.Tag = "自制"
        .Show 1, Me
    End With

    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
End Sub
Private Sub mnuEditStart_Click()
    Dim lng诊疗ID As Long
    Dim lng材料ID As Long
    Dim blnStop As Boolean
    
    If Me.ActiveControl Is lvwItems Then
            With Me.lvwItems
                If .SelectedItem Is Nothing Then Exit Sub
                If .SelectedItem.Icon = "start" Then Exit Sub
                
                If MsgBox("真的重新启用品种为“" & .SelectedItem.Text & "”吗？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
                
                '收费项目目录ID(材料ID)
                gstrSQL = "Zl_材料品种_Reuse(" & Val(Mid(.SelectedItem.Key, 2)) & ")"
                
                err = 0: On Error GoTo ErrHand
                zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
                .SelectedItem.SubItems(Me.lvwItems.ColumnHeaders("_撤档时间").Index - 1) = ""
                                
                .SelectedItem.Icon = "start": .SelectedItem.SmallIcon = "start"
                '恢复启用项目显示颜色
                .SelectedItem.ForeColor = .ForeColor
                For intCount = 1 To .ColumnHeaders.count - 1
                    .SelectedItem.ListSubItems(intCount).ForeColor = .ForeColor
                Next
                
            End With
    Else
            With Me.vsStuff
                lng材料ID = .RowData(.Row)
                lng诊疗ID = Val(.Cell(flexcpData, .Row, .ColIndex("编码")))
                blnStop = Trim(.TextMatrix(.Row, .ColIndex("撤档时间"))) <> ""
                
                If lng材料ID = 0 Then Exit Sub
                If blnStop = False Then Exit Sub
                
                If MsgBox("真的重新启用规格为:“" & .TextMatrix(.Row, .ColIndex("规格")) & "”吗？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
                
                '收费项目目录ID(材料ID)
                gstrSQL = "zl_卫生材料_REUSE(" & lng材料ID & ")"
                
                err = 0: On Error GoTo ErrHand
                zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
                .Cell(flexcpText, .Row, .ColIndex("撤档时间")) = ""
                Call SetVsGridColor(vsStuff, .Row, .ForeColor, False)
                
                With lvwItems
                    lvwItems.SelectedItem.SubItems(Me.lvwItems.ColumnHeaders("_撤档时间").Index - 1) = ""
                    lvwItems.SelectedItem.Icon = "start": lvwItems.SelectedItem.SmallIcon = "start"
                    '恢复启用项目显示颜色
                    .SelectedItem.ForeColor = .ForeColor
                    For intCount = 1 To .ColumnHeaders.count - 1
                        .SelectedItem.ListSubItems(intCount).ForeColor = .ForeColor
                    Next
                End With
            End With
    End If
    Call SetMenuEnable
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditStop_Click()
    Dim lng诊疗ID As Long
    Dim lng材料ID As Long
    Dim blnStop As Boolean
    
    If Me.ActiveControl Is lvwItems Then
        With Me.lvwItems
            If .SelectedItem Is Nothing Then Exit Sub
            If .SelectedItem.Icon = "stop" Then Exit Sub
            If MsgBox("真的要停用品种为“" & .SelectedItem.Text & "”吗？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
            
            gstrSQL = "Zl_材料品种_Stop(" & Val(Mid(.SelectedItem.Key, 2)) & ")"
        
            err = 0: On Error GoTo ErrHand
            zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
            
            If Me.mnuViewStoped.Checked = True Then
                .SelectedItem.Icon = "stop": .SelectedItem.SmallIcon = "stop"
                '将停用项目显示为红色
                .SelectedItem.SubItems(Me.lvwItems.ColumnHeaders("_撤档时间").Index - 1) = Format(sys.Currentdate, "yyyy-mm-dd")
                .SelectedItem.ForeColor = &HFF&
                For intCount = 1 To .ColumnHeaders.count - 1
                    .SelectedItem.ListSubItems(intCount).ForeColor = &HFF&
                Next
                
            Else
                Call .ListItems.Remove(.SelectedItem.Key)
            End If
            
        End With
    Else
            With Me.vsStuff
                lng材料ID = .RowData(.Row)
                lng诊疗ID = Val(.Cell(flexcpData, .Row, .ColIndex("编码")))
                blnStop = Trim(.TextMatrix(.Row, .ColIndex("撤档时间"))) <> ""
                
                If lng材料ID = 0 Then Exit Sub
                If blnStop Then Exit Sub
                
                If MsgBox("真的重新停用规格为“" & .TextMatrix(.Row, .ColIndex("规格")) & "”吗？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then Exit Sub
                
                '收费项目目录ID(材料ID)
                gstrSQL = "zl_卫生材料_STOP(" & lng材料ID & ")"
                
                err = 0: On Error GoTo ErrHand
                zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
                
                If Me.mnuViewStoped.Checked = True Then
                    .Cell(flexcpText, .Row, .ColIndex("撤档时间")) = Format(sys.Currentdate, "yyyy-mm-dd")
                    Call SetVsGridColor(vsStuff, .Row, &HFF&, False)
                Else
                    If .Rows = 2 Then
                        Call InitHgdPrivceHeadcol
                    Else
                        .RemoveItem .Row
                    End If
                End If
            End With
    End If
    Call SetMenuEnable
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mnuEditSubsection_Click()
    Dim blnreturn As Boolean
    frm加成率设置.EditCard Me, mstrPrivs, blnreturn
End Sub

Private Sub mnuEditUnit_Click()
    Dim lng材料ID  As Long, lng诊疗ID As Long
    '中标单位
    '招标材料中标单位设置
    With vsStuff
          lng材料ID = .RowData(.Row)
          lng诊疗ID = Val(.Cell(flexcpData, .Row, .ColIndex("编码")))
    End With
    
    With frmStuffUnitMgr
        Dim strType As String

        If lng材料ID = 0 Then   '如果没有记录就退出，因为无法判断药品材质
            .lblTag = 0
        Else
            .lblTag = lng材料ID
        End If
        
        .frmTag = 4
        .strPrivs = mstrPrivs
        If InStr(1, mstrPrivs, ";中标单位;") <> 0 Then
            .cmdClose.Tag = ""
        Else
            .cmdClose.Tag = "查阅"
        End If
        .Show 1, Me
    End With
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
End Sub

Private Sub mnuFileExcel_Click()
    Call zlRptPrint(3)
End Sub

Private Sub mnuFileMlPreview_Click()
    Dim strPara As String
    If Val(Mid(Me.tvwClass.SelectedItem.Key, 2)) > 0 Then
        strPara = " ID=" & Val(Mid(Me.tvwClass.SelectedItem.Key, 2))
    Else
        strPara = " 上级ID is null"
    End If
    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1711", Me, "分类=" & strPara & " ", 1)
End Sub

Private Sub mnuFileMlPrint_Click()
    Dim strPara As String
    If Val(Mid(Me.tvwClass.SelectedItem.Key, 2)) > 0 Then
        strPara = " ID=" & Val(Mid(Me.tvwClass.SelectedItem.Key, 2))
    Else
        strPara = " 上级ID is null"
    End If
    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1711", Me, "分类=" & strPara & " ", 2)

End Sub

Private Sub mnuFilePara_Click()
    '模块公共参数已经调整到药品参数设置模块，目前没有私有或本机参数，暂时屏蔽参数设置界面
'   frmStuffPara.ShowMe mstrPrivs, Me
   'mbln显示下级 = Val(GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\材料显示模式\", "显示下级", 0)) = 1
   
End Sub

Private Sub mnuFilePreView_Click()
    Call zlRptPrint(0)
End Sub

Private Sub mnuFilePrint_Click()
    Call zlRptPrint(1)
End Sub

Private Sub mnuFilePrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub mnuHelpAbout_Click()
    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
End Sub

Private Sub mnuHelpTitle_Click()
    Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub mnufileexit_Click()
    Unload Me
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hwnd)
End Sub


Private Sub mnuPriceLists_Click()
    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1711_2", Me)
End Sub

Private Sub mnuPriceTable_Click()
    Call ReportOpen(gcnOracle, glngSys, "ZL1_BILL_1711_1", Me)
End Sub

Private Sub mnuViewDownLevel_Click()
    mnuViewDownLevel.Checked = Not mnuViewDownLevel.Checked
    mbln显示下级 = mnuViewDownLevel.Checked
    Call mnuViewRefresh_Click
End Sub

Private Sub mnuViewFind_Click()
    With frmStuffFind
        Call .ShowMe(Me, mnuViewStoped.Checked)
    End With
End Sub

Private Sub mnuViewIcon_Click(Index As Integer)
   Dim intTemp As Integer
    For intTemp = 0 To 3
        mnuViewIcon(intTemp).Checked = False
    Next
    mnuViewIcon(Index).Checked = True
    lvwItems.View = Index
    lvwItems.Refresh
End Sub

Private Sub mnuViewPrices_Click()
    Me.mnuViewPrices.Checked = Not Me.mnuViewPrices.Checked
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
End Sub

Private Sub mnuViewRefresh_Click()
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    Call zlRefRecords
End Sub

Private Sub mnuViewStates_Click()
    Me.mnuViewStates.Checked = Not Me.mnuViewStates.Checked
    Me.stbThis.Visible = Me.mnuViewStates.Checked
    Form_Resize
End Sub

Private Sub mnuViewStoped_Click()
    Me.mnuViewStoped.Checked = Not Me.mnuViewStoped.Checked
    Call zlRefRecords
End Sub
Private Sub mnuReportItem_Click(Index As Integer)
    Dim lng分类id As Long
    Dim lng材料ID As Long
    
    If tvwClass.SelectedItem Is Nothing Then
        lng分类id = 0
    Else
        lng分类id = Val(Mid(tvwClass.SelectedItem.Key, 2))
    End If
    
    lng材料ID = 0
    If Not Me.lvwItems.SelectedItem Is Nothing Then
        lng材料ID = Val(Mid(Me.lvwItems.SelectedItem.Key, 2))
    End If
    
    '2006-04-25:刘兴宏:增加自定义报表发布到模块的功能
    Call ReportOpen(gcnOracle, Split(mnuReportItem(Index).Tag, ",")(0), Split(mnuReportItem(Index).Tag, ",")(1), Me, "分类=" & lng分类id, "材料=" & lng材料ID)
End Sub
Private Sub mnuViewToolbarStAnd_Click()
    Me.mnuViewToolbarStand.Checked = Not Me.mnuViewToolbarStand.Checked
    Me.clbThis.Visible = Me.mnuViewToolbarStand.Checked
    Form_Resize
End Sub

Private Sub mnuViewToolbarText_Click()
    Dim i As Integer
    Me.mnuViewToolbarText.Checked = Not Me.mnuViewToolbarText.Checked
    If Me.mnuViewToolbarText.Checked Then
        For i = 1 To Me.tlbThis.Buttons.count
            Me.tlbThis.Buttons(i).Caption = Me.tlbThis.Buttons(i).Tag
        Next
    Else
        For i = 1 To Me.tlbThis.Buttons.count
            Me.tlbThis.Buttons(i).Caption = ""
        Next
    End If
    Me.clbThis.Bands(1).MinHeight = Me.tlbThis.Height
    Me.clbThis.Refresh
    Form_Resize
End Sub

Private Sub picHBar_S_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        picHBar_S.BackColor = &H8000000C
    End If
End Sub

Private Sub picHBar_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Me.picHBar_S.Top = Me.picHBar_S.Top + Y
    End If
End Sub

Private Sub picHBar_S_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Call Form_Resize
        picHBar_S.BackColor = &H8000000A
    End If
End Sub


Private Sub picUp_S_Paint()
    '功能:打印时间范围
    picUp_S.Cls
    picUp_S.CurrentX = 90
    picUp_S.CurrentY = 60
    picUp_S.Print "材料规格及价格信息"
End Sub

Private Sub picVBar_S_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        picVBar_S.BackColor = &H8000000C
    End If
End Sub

Private Sub picVBar_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Me.picVBar_S.Left = Me.picVBar_S.Left + X
    End If
End Sub

Private Sub picVBar_S_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        Call Form_Resize
        picVBar_S.BackColor = &H8000000A
    End If
End Sub

Private Sub tlbThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim lvwTemp As ListView
    
    Select Case Button.Key
    Case "Preview"
        Call mnuFilePreView_Click
    Case "Print"
        Call mnuFilePrint_Click
    Case "Class"
        Call PopupMenu(Me.mnuClass, 2)
    Case "增加"
        If Me.ActiveControl Is lvwItems Then
            mnuEditAddName_Click
        Else
            mnuEditItemAdd_Click
        End If
    Case "修改"
        If Me.ActiveControl Is lvwItems Then
            If mnuEditModifyName.Visible Then
                mnuEditModifyName_Click
            End If
        Else
            mnuEditItemMod_Click
        End If
    Case "删除"
        If Me.ActiveControl Is lvwItems Then
            mnuEditDeleName_Click
        Else
            mnuEditItemDel_Click
        End If
    Case "Start"
        Call mnuEditStart_Click
    Case "Stop"
        Call mnuEditStop_Click
    Case "Find"
        Call mnuViewFind_Click
    Case "Filter"
        Call mnuFilter_Click
    Case "View"
        Set lvwTemp = lvwItems
        mnuViewIcon(lvwTemp.View).Checked = False
        If lvwTemp.View = 3 Then
            mnuViewIcon(0).Checked = True
            lvwTemp.View = 0
        Else
            mnuViewIcon(lvwTemp.View + 1).Checked = True
            lvwTemp.View = lvwTemp.View + 1
        End If
        
    Case "Help"
        Call mnuHelpTitle_Click
    Case "Exit"
        Call mnufileexit_Click
    End Select
End Sub

Private Sub tlbThis_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)
    Dim i As Integer
    Select Case ButtonMenu.Key
    Case "增加品种"
        Call mnuEditAddName_Click
    Case "增加规格"
        Call mnuEditItemAdd_Click
    Case Else
        For i = 0 To 3
            mnuViewIcon(i).Checked = False
        Next
        mnuViewIcon(ButtonMenu.Index - 1).Checked = True
        lvwItems.View = ButtonMenu.Index - 1
    End Select
End Sub

Private Sub tlbThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    PopupMenu Me.mnuViewToolbar, 2
End Sub

Private Sub tvwClass_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    If InStr(1, mstrPrivs, ";卫材分类管理;") = 0 Then Exit Sub
    Call PopupMenu(Me.mnuClass, 2)
End Sub

Private Sub tvwClass_NodeClick(ByVal Node As MSComctlLib.Node)
    If Me.lvwItems.Tag = Node.Key Then Exit Sub
    Me.lvwItems.Tag = Node.Key
    Call zlRefRecords
    
    Call SetMenuEnable

End Sub
Private Sub InitLvwStuffColHead()
    '-----------------------------------------------------------------------------------------------------------
    '--功  能:初始列头
    '--入参数:
    '--出参数:
    '--返  回:
    '-----------------------------------------------------------------------------------------------------------
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "_名称", "名称", 2500
        .Add , "_编码", "编码", 1000
        .Add , "_英文名称", "英文名称", 1000
        .Add , "_散装单位", "散装单位", 1000
        .Add , "_适用性别", "适用性别", 1200
        .Add , "_建档时间", "建档时间", 1200
        .Add , "_撤档时间", "撤档时间", 1200
    End With
    
    With Me.lvwItems
        .ListItems.Clear
        .ColumnHeaders("_编码").Position = 1
        .SortKey = .ColumnHeaders("_编码").Index - 1
        .SortOrder = lvwAscending
    End With
    Call RestoreListViewState(Me.lvwItems, Me.Name, Me.lvwItems.View)
    

End Sub

Private Sub RefClassesAppend()
'创建动态纪录集
    On Error GoTo ErrHand
    Set mrsRefClasses = New ADODB.Recordset
    If mrsRefClasses.State <> 1 Then
        With mrsRefClasses
            Call .Fields.Append("ID", adDouble, 20, adFldIsNullable)
            Call .Fields.Append("上级id", adDouble, 20, adFldIsNullable)
            Call .Fields.Append("编码", adLongVarChar, 20, adFldIsNullable)
            Call .Fields.Append("名称", adLongVarChar, 100, adFldIsNullable)
            Call .Fields.Append("简码", adLongVarChar, 100, adFldIsNullable)
            
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open '打开纪录集
        End With
    End If
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub RefRecordsAppend()
'创建动态纪录集
    On Error GoTo ErrHand
    Set mrsRefRecords = New ADODB.Recordset
    If mrsRefRecords.State <> 1 Then
        With mrsRefRecords
            Call .Fields.Append("诊疗id", adDouble, 20, adFldIsNullable)
            Call .Fields.Append("分类id", adDouble, 20, adFldIsNullable)
            Call .Fields.Append("编码", adLongVarChar, 50, adFldIsNullable)
            Call .Fields.Append("名称", adLongVarChar, 100, adFldIsNullable)
            Call .Fields.Append("英文名称", adLongVarChar, 100, adFldIsNullable)
            Call .Fields.Append("散装单位", adLongVarChar, 50, adFldIsNullable)
            Call .Fields.Append("建档时间", adDate, , adFldIsNullable)
            Call .Fields.Append("撤档时间", adDate, , adFldIsNullable)
            Call .Fields.Append("适用性别", adDouble, 10, adFldIsNullable)

            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .LockType = adLockOptimistic
            .Open '打开纪录集
        End With
    End If
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub zlRefClasses()
    '---------------------------------------------
    '填写诊疗分类项目
    '---------------------------------------------
    Dim lngNode As Long
    Dim arrExecute As Variant
    Dim strID As String
    Dim i As Integer
    Dim str纪录ID串 As String
    
    err = 0
    On Error GoTo ErrHand:
    
    If Val(tvwClass.Tag) = 0 Then
        gstrSQL = "" & _
            "   select ID,上级ID,编码,名称,简码" & _
            "   From 诊疗分类目录" & _
            "   Where 类型 = 7 " & _
            "   start with 上级ID is null Connect by prior ID=上级ID"
            
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    Else
        If mstrMaterialID = "" Then
            tvwClass.Nodes.Clear
            Me.lvwItems.ListItems.Clear
            Exit Sub
        End If
        
        gstrSQL = "" & _
            "   select /*+ Rule*/ Distinct A.ID,A.上级ID,A.编码,A.名称,A.简码" & _
            "   From 诊疗分类目录 A,诊疗项目目录 B, Table(Cast(f_Num2List([1]) As zlTools.t_NumList)) C " & _
            "   Where A.类型 = 7 " & _
            "  And A.id=B.分类id And B.id=C.Column_Value "
            
        'mstrMaterialID串可能超过4K，分解后分别执行SQL再汇总数据集
        Call RefClassesAppend
        arrExecute = GetArrayByStr(mstrMaterialID, 4000, ",")
        For i = 0 To UBound(arrExecute)
            strID = arrExecute(i)
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strID)
            
            Do While Not rsTemp.EOF
                If InStr(str纪录ID串 & ",", "," & rsTemp!Id & ",") = 0 Then
                    With mrsRefClasses
                        .AddNew
                        .Fields!Id = rsTemp!Id
                        .Fields!上级id = rsTemp!上级id
                        .Fields!编码 = rsTemp!编码
                        .Fields!名称 = rsTemp!名称
                        .Fields!简码 = rsTemp!简码
                        .Update
                    End With
                    str纪录ID串 = str纪录ID串 & "," & rsTemp!Id
                End If
                
                rsTemp.MoveNext
            Loop
        Next
        
        Set rsTemp = Nothing
        Set rsTemp = mrsRefClasses.Clone
    End If
    
    If Not tvwClass.SelectedItem Is Nothing Then
        lngNode = Val(Mid(tvwClass.SelectedItem.Key, 2))
    End If
    
    With rsTemp
        tvwClass.Nodes.Clear
        If Val(tvwClass.Tag) = 1 Then Set objNode = Me.tvwClass.Nodes.Add(, , "_ALL", "所有过滤结果", "close")
        Do While Not .EOF
            If Val(tvwClass.Tag) = 0 Then
                If IsNull(!上级id) Then
                    Set objNode = Me.tvwClass.Nodes.Add(, , "_" & !Id, "[" & !编码 & "]" & !名称, "close", "expend")
                Else
                    Set objNode = Me.tvwClass.Nodes.Add("_" & !上级id, tvwChild, "_" & !Id, "[" & !编码 & "]" & !名称, "close", "expend")
                End If
            Else
                Set objNode = Me.tvwClass.Nodes.Add("_ALL", tvwChild, "_" & !Id, "[" & !编码 & "]" & !名称, "close", "expend")
            End If
            
            objNode.Sorted = True
            objNode.Tag = IIf(IsNull(!简码), "", !简码)
         '   objNode.ExpandedImage = "expend"
            .MoveNext
        Loop
    End With
    
    err = 0
    On Error Resume Next
    If Me.tvwClass.Nodes.count > 0 Then
        If lngNode <> 0 Then
            Me.tvwClass.Nodes("_" & lngNode).Selected = True
            Me.tvwClass.Nodes("_" & lngNode).Expanded = True
            If err <> 0 Then
                Me.tvwClass.Nodes(1).Selected = True
                Me.tvwClass.Nodes(1).Expanded = True
            End If
        Else
            Me.tvwClass.Nodes(1).Selected = True
             Me.tvwClass.Nodes(1).Expanded = True
        End If
        Call zlRefRecords
    Else
        Me.lvwItems.ListItems.Clear
    End If
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Public Sub zlRefRecords(Optional lngLocale诊疗ID As Long)
    '---------------------------------------------
    '填写卫生材料列表
    '---------------------------------------------
    Dim lngId As Long
    Dim lngPreItemID As Long
    Dim arrExecute As Variant
    Dim strID As String
    Dim i As Integer
    
    err = 0
    On Error GoTo ErrHand
    If Me.lvwItems.SelectedItem Is Nothing Then
        lngPreItemID = 0
    Else
        lngPreItemID = Val(Mid(Me.lvwItems.SelectedItem.Key, 2))
    End If
    If tvwClass.SelectedItem Is Nothing Then
        lngId = 0
    Else
        lngId = Val(Mid(tvwClass.SelectedItem.Key, 2))
    End If
    
    If Val(tvwClass.Tag) = 0 Then
        gstrSQL = " " & _
         " SELECT distinct a.id as 诊疗ID,A.分类ID, a.编码,a.名称,c.名称 as 英文名称,a.计算单位 as 散装单位," & _
         "      to_char(a.建档时间,'yyyy-mm-dd') as 建档时间,nvl(a.撤档时间,to_date('3000-01-01','YYYY-MM-DD')) as 撤档时间,Nvl(A.适用性别,0) As 适用性别 " & _
         " FROM 诊疗项目目录 a ,诊疗项目别名 c" & _
         " WHERE a.id=c.诊疗项目id(+) and c.性质(+)=2 " & _
                IIf(mbln显示下级, " And a.分类id in (Select ID From 诊疗分类目录 where 类型=7 start with id=[1] connect by prior id=上级id  )", "  and  A.分类id=[1]")
    Else
        gstrSQL = " " & _
         " SELECT /*+ Rule*/ distinct a.id as 诊疗ID,A.分类ID, a.编码,a.名称,c.名称 as 英文名称,a.计算单位 as 散装单位," & _
         "      to_char(a.建档时间,'yyyy-mm-dd') as 建档时间,nvl(a.撤档时间,to_date('3000-01-01','YYYY-MM-DD')) as 撤档时间,Nvl(A.适用性别,0) As 适用性别 " & _
         " FROM 诊疗项目目录 a ,诊疗项目别名 c, Table(cast(f_Num2List([2]) as zlTools.t_NumList)) B " & _
         " WHERE a.id=c.诊疗项目id(+) and c.性质(+)=2 And a.ID=b.Column_Value "
         
         If lngId = 0 Then
            gstrSQL = gstrSQL & IIf(mbln显示下级, "", " And a.分类id =[1]")
         Else
            gstrSQL = gstrSQL & " And a.分类id =[1]"
         End If
    End If
    
    If Me.mnuViewStoped.Checked = False Then
        gstrSQL = gstrSQL & " and (a.撤档时间 is null or a.撤档时间=to_date('3000-01-01','YYYY-MM-DD'))"
    End If
    gstrSQL = gstrSQL & " order by a.编码"

    If Val(tvwClass.Tag) = 0 Then
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngId)
    Else
        'mstrMaterialID串可能超过4K，分解后分别执行SQL再汇总数据集
        Call RefRecordsAppend
        arrExecute = GetArrayByStr(mstrMaterialID, 4000, ",")
        For i = 0 To UBound(arrExecute)
            strID = arrExecute(i)
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngId, strID)
               
            Do While Not rsTemp.EOF
                    With mrsRefRecords
                        .AddNew
                        .Fields!诊疗id = rsTemp!诊疗id
                        .Fields!分类id = rsTemp!分类id
                        .Fields!编码 = rsTemp!编码
                        .Fields!名称 = rsTemp!名称
                        .Fields!英文名称 = rsTemp!英文名称
                        .Fields!散装单位 = rsTemp!散装单位
                        .Fields!建档时间 = rsTemp!建档时间
                        .Fields!撤档时间 = rsTemp!撤档时间
                        .Fields!适用性别 = rsTemp!适用性别
                        .Update
                    End With
                rsTemp.MoveNext
            Loop
        Next
    
        Set rsTemp = Nothing
        Set rsTemp = mrsRefRecords.Clone
    End If
    
    With rsTemp
            Me.lvwItems.ListItems.Clear
            Do While Not .EOF
                Set objItem = Me.lvwItems.ListItems.Add(, "_" & !诊疗id, zlStr.Nvl(!名称))
                If Format(!撤档时间, "YYYY-MM-DD") = "3000-01-01" Then
                    objItem.Icon = "start": objItem.SmallIcon = "start"
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_撤档时间").Index - 1) = ""
                Else
                    objItem.Icon = "stop": objItem.SmallIcon = "stop"
                    objItem.SubItems(Me.lvwItems.ColumnHeaders("_撤档时间").Index - 1) = Format(!撤档时间, "yyyy-mm-dd")
                End If
                
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_编码").Index - 1) = !编码
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_英文名称").Index - 1) = zlStr.Nvl(!英文名称)
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_散装单位").Index - 1) = zlStr.Nvl(!散装单位)
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_适用性别").Index - 1) = IIf(!适用性别 = 1, "男性", IIf(!适用性别 = 2, "女性", "无性别区分"))
                objItem.SubItems(Me.lvwItems.ColumnHeaders("_建档时间").Index - 1) = zlStr.Nvl(!建档时间)
                objItem.Tag = zlStr.Nvl(!分类id)
                If !诊疗id = lngLocale诊疗ID Then
                    objItem.Selected = True
                End If
                
                If Format(!撤档时间, "YYYY-MM-DD") <> "3000-01-01" Then
                    objItem.ForeColor = &HFF&
                    For intCount = 1 To Me.lvwItems.ColumnHeaders.count - 1
                        objItem.ListSubItems(intCount).ForeColor = &HFF&
                    Next
                End If
                If Val(zlStr.Nvl(!诊疗id)) = lngPreItemID Then
                    objItem.Selected = True
                    objItem.EnsureVisible
                    
                End If
                 .MoveNext
            Loop
    End With
    If Me.lvwItems.ListItems.count > 0 Then
    
        If Me.lvwItems.SelectedItem Is Nothing Then
            Me.lvwItems.ListItems(1).Selected = True
        End If
        Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
        
        err = 0: On Error Resume Next
        DoEvents: Me.lvwItems.SelectedItem.EnsureVisible
        Me.stbThis.Panels(2).Text = "该分类共有" & Me.lvwItems.ListItems.count & "种卫材品种"
    Else
        
        For intCount = Me.lblComment.LBound To Me.lblComment.UBound
            Me.lblComment(intCount).Caption = ""
        Next
        Me.stbThis.Panels(2).Text = ""
        Call InitHgdPrivceHeadcol
    End If
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub zlRptPrint(ByVal bytMode As Byte)
    '-------------------------------------------------
    '功能:记录表打印
    '参数:bytMode=1 打印;2 预览;3 输出到EXCEL
    '-------------------------------------------------
    Dim objPrint As New zlPrintLvw
    err = 0: On Error Resume Next
    Set objPrint.Body.objData = Me.lvwItems
    objPrint.Title.Text = "卫生材料品种清单"
    
    objPrint.UnderAppItems.Add "分类：" & Me.tvwClass.SelectedItem.Text
    objPrint.BelowAppItems.Add "打印时间：" & Now
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrViewLvw objPrint, bytMode
    Else
        zlPrintOrViewLvw objPrint, bytMode
    End If
End Sub

Public Sub zlLocateItem(lng分类id As Long, ByVal ln诊疗ID As Long, lng材料ID As Long)
    '---------------------------------------------
    '功能:定位到指定的材料上去，在查找时使用
    '编制:刘兴宏
    '日期:2007-05-28
    '---------------------------------------------
    Dim lngRow As Long
    Dim lstItem As ListItem, tvwNode As Node
    On Error GoTo ErrHand
    Set tvwNode = tvwClass.SelectedItem
    Set lstItem = lvwItems.SelectedItem
    
    Set Me.tvwClass.SelectedItem = Me.tvwClass.Nodes("_" & lng分类id)
    Me.tvwClass.Nodes("_" & lng分类id).Selected = True
    Me.tvwClass.SelectedItem.EnsureVisible
    Call zlRefRecords
    Set Me.lvwItems.SelectedItem = Me.lvwItems.ListItems("_" & ln诊疗ID)
    Me.lvwItems.SelectedItem.EnsureVisible
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
    '定位到材料ID上去
    With vsStuff
        For lngRow = 1 To .Rows - 1
            If .RowData(lngRow) = lng材料ID Then
                .Row = lngRow
                Call .ShowCell(.Row, 1)
            End If
        Next
    End With
    Exit Sub
ErrHand:
    Set tvwClass.SelectedItem = tvwNode
    Call zlRefRecords
    Set Me.lvwItems.SelectedItem = Me.lvwItems.ListItems(lstItem.Key)
    Me.lvwItems.SelectedItem.EnsureVisible
    Call lvwItems_ItemClick(Me.lvwItems.SelectedItem)
End Sub

Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hwnd)
End Sub

Private Sub picStuffLst_Resize()
    err = 0: On Error Resume Next
    With picStuffLst
        vsStuff.Top = 20
        vsStuff.Left = 20
        vsStuff.Width = .ScaleWidth - 40
        vsStuff.Height = .ScaleHeight - 40
    End With
End Sub
Private Sub picTabPrice_Resize()
    err = 0: On Error Resume Next
    With picTabPrice
        vsPrice.Top = 20
        vsPrice.Left = 20
        vsPrice.Width = .ScaleWidth - 40
        fraComment.Top = .ScaleHeight - fraComment.Height - 40
        fraComment.Left = vsPrice.Left
        vsPrice.Height = fraComment.Top - vsPrice.Top - 20
    End With
End Sub

Private Sub vsStuff_DblClick()
    Dim lng材料ID As Long
    Dim lng分类id As Long
    
    With vsStuff
        If .RowData(.Row) = 0 Then Exit Sub
        lng材料ID = .RowData(.Row)
    End With
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    
    If tvwClass.SelectedItem Is Nothing Then
        lng分类id = 0
    Else
        lng分类id = Val(Mid(tvwClass.SelectedItem.Key, 2))
    End If
    Call frmStuffSpec.ShowEditCard(Me, g查看, Mid(Me.lvwItems.SelectedItem.Key, 2), lng分类id, CStr(lng材料ID), mstrPrivs)
    
End Sub

 

Private Sub vsStuff_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   If Button <> 2 Then Exit Sub
   PopupMenu mnuEdit
End Sub

Private Sub vsStuff_RowColChange()
    Call SetMenuEnable
End Sub
Private Function LoadCostData(ByVal lng诊疗ID As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '功能:填充成本价调价信息
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2008-11-10 15:17:09
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    
    
    '原成本价 NUMBER(16,7),效期 DATE,批号 VARCHAR2(20),产地 VARCHAR2(40)
    
    gstrSQL = "" & _
    " Select '' as NO, I.ID As 材料ID, '[' || I.编码 || ']' || I.名称 || ' ' || I.规格 || ' ' || A.产地 As 材料, " & _
    "        P.编码,P.名称 As 库房, A.批号 As 批号, " & _
    "        I.计算单位 As 单位, S.包装单位,S.换算系数, " & _
    "        A.新成本价 As 成本价,A.原成本价,'' as 执行日期, '未执行' 摘要 " & _
    " From 成本价调价信息 A, 材料特性 S, 收费项目目录 I,部门表 P" & _
    " Where A.药品ID=S.材料ID And A.药品ID=I.ID and A.库房id=P.id" & _
    "       And S.诊疗id = [1] And A.执行日期 is NULL " & _
            IIf(Me.mnuViewStoped.Checked = False, "and (I.撤档时间 is null or I.撤档时间>=to_date('3000-01-01','YYYY-MM-DD'))", "")
    
  gstrSQL = gstrSQL & " Union all " & vbCrLf & _
    " Select  B.NO as NO, I.ID As 材料ID, '[' || I.编码 || ']' || I.名称 || ' ' || I.规格 || ' ' || A.产地 As 材料, " & _
    "        P.编码,P.名称 As 库房, A.批号 , " & _
    "        I.计算单位 As 单位, S.包装单位,S.换算系数, " & _
    "        A.新成本价 As 成本价,A.原成本价,to_char(A.执行日期,'yyyy-mm-dd') 执行日期, B.摘要 " & _
    " From 成本价调价信息 A,药品收发记录 B, 材料特性 S, 收费项目目录 I,部门表 P" & _
    " Where A.收发ID=B.ID and A.药品ID=S.材料ID And A.药品ID=I.ID and A.库房id=P.id" & _
    "       And S.诊疗id = [1] And A.执行日期 is NOt NULL " & _
            IIf(Me.mnuViewStoped.Checked = False, "and (I.撤档时间 is null or I.撤档时间>=to_date('3000-01-01','YYYY-MM-DD'))", "")
    
    gstrSQL = gstrSQL & " Order By 编码,材料,执行日期 Desc, NO Desc "
    
    err = 0: On Error GoTo ErrHand:
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng诊疗ID)
    
    With vsCost
        .Rows = 2
        .Clear 1
        .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
        If rsTemp.RecordCount = 0 Then
            LoadCostData = True: Exit Function
        End If
        
        .Rows = .FixedRows + rsTemp.RecordCount
        .Redraw = flexRDNone
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("NO")) = zlStr.Nvl(rsTemp!NO)
            .TextMatrix(lngRow, .ColIndex("卫材信息")) = zlStr.Nvl(rsTemp!材料)
            .Cell(flexcpData, lngRow, .ColIndex("卫材信息")) = zlStr.Nvl(rsTemp!材料ID)
            .TextMatrix(lngRow, .ColIndex("库房")) = zlStr.Nvl(rsTemp!库房)
            .TextMatrix(lngRow, .ColIndex("批号")) = zlStr.Nvl(rsTemp!批号)
            .TextMatrix(lngRow, .ColIndex("单位")) = IIf(mintUnit = 0, zlStr.Nvl(rsTemp!单位), zlStr.Nvl(rsTemp!包装单位))
            .Cell(flexcpData, lngRow, .ColIndex("单位")) = IIf(mintUnit = 0, 1, Val(zlStr.Nvl(rsTemp!换算系数)))
            .TextMatrix(lngRow, .ColIndex("成本价")) = Format(Val(zlStr.Nvl(rsTemp!成本价)) * Val(.Cell(flexcpData, lngRow, .ColIndex("单位"))), mFMT.FM_成本价)
            .TextMatrix(lngRow, .ColIndex("执行日期")) = zlStr.Nvl(rsTemp!执行日期)
            .TextMatrix(lngRow, .ColIndex("说明")) = zlStr.Nvl(rsTemp!摘要)
            
            .Cell(flexcpBackColor, lngRow, 0, lngRow, .Cols - 1) = IIf(zlStr.Nvl(rsTemp!执行日期) = "", RGB(225, 255, 255), RGB(240, 240, 240))
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        .Row = 1
        .Redraw = flexRDBuffered
    End With
    LoadCostData = True
    Exit Function
ErrHand:
    vsCost.Redraw = flexRDBuffered
    If ErrCenter = 1 Then
        Resume
    End If
End Function

 

Private Sub vsStuff_GotFocus()
    With vsStuff
        .SelectionMode = flexSelectionByRow
        .BackColorSel = &H8000000D
    End With
End Sub

Private Sub vsStuff_LostFocus()
    With vsStuff
        .BackColorSel = GRD_LOSTFOCUS_COLORSEL
    End With
End Sub

Private Sub vsPrice_GotFocus()
    With vsPrice
        .SelectionMode = flexSelectionByRow
        .BackColorSel = &H8000000D
    End With
End Sub

Private Sub vsPrice_LostFocus()
    With vsPrice
        .BackColorSel = GRD_LOSTFOCUS_COLORSEL
    End With
End Sub

Private Sub mnuEditPrice_Click()
    If checkNotExecutePrice(vsStuff.TextMatrix(vsStuff.Row, vsStuff.ColIndex("ID"))) Then Exit Sub
    Call frmStuffPriceCard.ShowMe(Me, 0, "", 0, 1, vsStuff.TextMatrix(vsStuff.Row, vsStuff.ColIndex("ID")))
    frmStuffMgr.SetFocus
End Sub

Private Function checkNotExecutePrice(Optional ByVal lngDrugID As Long = 0) As Boolean
    '功能 ：检查是否存在未执行的价格
    Dim RecCheck As New ADODB.Recordset
    On Error GoTo ErrHandle
    
    checkNotExecutePrice = False
    '判断是否有未执行的历史价格
    gstrSQL = " Select Count(*) Records From 收费价目 Where 变动原因=0 " & GetPriceClassString("") & _
        " And 执行日期 > Sysdate And 收费细目ID=[1]"
    Set RecCheck = zlDatabase.OpenSQLRecord(gstrSQL, "", lngDrugID)

    With RecCheck
        If Not .EOF Then
            If Not IsNull(!Records) Then
                If !Records <> 0 Then
                    MsgBox "该规格还存在未执行的售价调价记录，未执行卫材不能调价！", vbInformation, gstrSysName
                    checkNotExecutePrice = True
                    Exit Function
                End If
            End If
        End If
    End With

    '检查是否还有未执行的成本价调价计划
    gstrSQL = "Select 1 From 成本价调价信息 Where 药品id = [1] And 执行日期 Is Null And Rownum = 1 "
    Set RecCheck = zlDatabase.OpenSQLRecord(gstrSQL, "", lngDrugID)

    If RecCheck.RecordCount > 0 Then
        MsgBox "该规格还存在未执行的成本价调价记录，未执行卫材不能调价！", vbInformation, gstrSysName
        checkNotExecutePrice = True
        Exit Function
    End If
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
