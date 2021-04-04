VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Begin VB.Form frmLaterVisit 
   Caption         =   "体检随访管理"
   ClientHeight    =   7185
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11490
   Icon            =   "frmLaterVisit.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7185
   ScaleWidth      =   11490
   ShowInTaskbar   =   0   'False
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   6825
      Width           =   11490
      _ExtentX        =   20267
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmLaterVisit.frx":1CFA
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联信息产业公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15187
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
   Begin MSComctlLib.ImageList ils16 
      Left            =   3945
      Top             =   5700
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
            Picture         =   "frmLaterVisit.frx":258E
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisit.frx":29E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisit.frx":2CFA
            Key             =   "class"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils32 
      Left            =   3315
      Top             =   5700
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
            Picture         =   "frmLaterVisit.frx":3294
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisit.frx":36E6
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   705
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   11490
      _ExtentX        =   20267
      _ExtentY        =   1244
      BandCount       =   1
      _CBWidth        =   11490
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
         Width           =   11370
         _ExtentX        =   20055
         _ExtentY        =   1138
         ButtonWidth     =   820
         ButtonHeight    =   1138
         AllowCustomize  =   0   'False
         Style           =   1
         ImageList       =   "ilsMenu"
         HotImageList    =   "ilsMenuHot"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   12
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
               Object.ToolTipText     =   "打印"
               Object.Tag             =   "打印"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "增加"
               Key             =   "增加"
               Object.ToolTipText     =   "增加"
               Object.Tag             =   "增加"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "修改"
               Key             =   "修改"
               Object.ToolTipText     =   "修改"
               Object.Tag             =   "修改"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "删除"
               Key             =   "删除"
               Object.ToolTipText     =   "删除"
               Object.Tag             =   "删除"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_2"
               Style           =   3
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "结束"
               Key             =   "结束"
               Object.ToolTipText     =   "结束"
               Object.Tag             =   "结束"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "恢复"
               Key             =   "恢复"
               Object.ToolTipText     =   "恢复"
               Object.Tag             =   "恢复"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_3"
               Style           =   3
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "帮助"
               Key             =   "帮助"
               Object.ToolTipText     =   "帮助"
               Object.Tag             =   "帮助"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "退出"
               Key             =   "退出"
               Object.ToolTipText     =   "退出"
               Object.Tag             =   "退出"
               ImageIndex      =   9
            EndProperty
         EndProperty
      End
   End
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   9765
      Top             =   465
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisit.frx":3A00
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisit.frx":3C20
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisit.frx":3E40
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisit.frx":405A
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisit.frx":427A
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisit.frx":449A
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisit.frx":46B4
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisit.frx":48CE
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisit.frx":4AEE
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMenuHot 
      Left            =   8955
      Top             =   465
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   9
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisit.frx":4D0E
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisit.frx":4F2E
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisit.frx":514E
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisit.frx":54A0
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisit.frx":56C0
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisit.frx":58E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisit.frx":5AFA
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisit.frx":5D14
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisit.frx":5F34
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   1575
      Index           =   2
      Left            =   3600
      TabIndex        =   6
      Top             =   825
      Width           =   5805
      _cx             =   10239
      _cy             =   2778
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
      HighLight       =   0
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
      Begin VB.Line lnY2 
         Index           =   0
         Visible         =   0   'False
         X1              =   270
         X2              =   270
         Y1              =   420
         Y2              =   1635
      End
      Begin VB.Line lnX2 
         Index           =   0
         Visible         =   0   'False
         X1              =   -555
         X2              =   1230
         Y1              =   555
         Y2              =   555
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfContent 
      Height          =   2685
      Left            =   3630
      TabIndex        =   8
      Top             =   2685
      Width           =   5805
      _cx             =   10239
      _cy             =   4736
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
      GridColor       =   12698049
      GridColorFixed  =   -2147483632
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483643
      FocusRect       =   0
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   0
      GridLinesFixed  =   2
      GridLineWidth   =   1
      Rows            =   1
      Cols            =   6
      FixedRows       =   0
      FixedCols       =   0
      RowHeightMin    =   255
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
      AutoSizeMode    =   1
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
      WordWrap        =   -1  'True
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
   Begin MSComctlLib.ImageList ils13 
      Left            =   10095
      Top             =   3795
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisit.frx":6154
            Key             =   "状态"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisit.frx":64EE
            Key             =   "完成"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisit.frx":6A88
            Key             =   "up"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisit.frx":6C4A
            Key             =   "down"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisit.frx":6E0C
            Key             =   "people"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisit.frx":73A6
            Key             =   "people1"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmLaterVisit.frx":DC08
            Key             =   "bill"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsf 
      Height          =   1020
      Index           =   0
      Left            =   540
      TabIndex        =   9
      Top             =   1335
      Width           =   1620
      _cx             =   2857
      _cy             =   1799
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
      HighLight       =   0
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
      Begin VB.Line lnX0 
         Index           =   0
         Visible         =   0   'False
         X1              =   -555
         X2              =   1230
         Y1              =   555
         Y2              =   555
      End
      Begin VB.Line lnY0 
         Index           =   0
         Visible         =   0   'False
         X1              =   270
         X2              =   270
         Y1              =   420
         Y2              =   1635
      End
   End
   Begin VB.Frame fraInfo 
      Height          =   765
      Left            =   225
      TabIndex        =   3
      Top             =   5100
      Width           =   2925
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1380
         TabIndex        =   7
         Top             =   210
         Width           =   1140
      End
      Begin VB.CommandButton cmdMenu 
         Height          =   270
         Left            =   75
         Picture         =   "frmLaterVisit.frx":E1A2
         Style           =   1  'Graphical
         TabIndex        =   4
         TabStop         =   0   'False
         Top             =   225
         Width           =   285
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&6.姓名"
         Height          =   180
         Index           =   1
         Left            =   420
         TabIndex        =   5
         Tag             =   "姓名"
         Top             =   270
         Width           =   540
      End
   End
   Begin VB.Image imgX_S 
      Height          =   45
      Left            =   3630
      MousePointer    =   7  'Size N S
      Top             =   2460
      Width           =   5115
   End
   Begin VB.Image imgY_S 
      Height          =   4395
      Left            =   3090
      MousePointer    =   9  'Size W E
      Top             =   765
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
      Begin VB.Menu mnuFilePara 
         Caption         =   "参数设置(&R)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFile_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu mnuEdit 
      Caption         =   "编辑(&E)"
      Begin VB.Menu mnuEditAdd 
         Caption         =   "增加随访(&A)"
         Shortcut        =   ^A
      End
      Begin VB.Menu mnuEditModify 
         Caption         =   "修改随访(&M)"
      End
      Begin VB.Menu mnuEditDelete 
         Caption         =   "删除随访(&D)"
      End
      Begin VB.Menu mnuEdit_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuEditStop 
         Caption         =   "结束随访(S)"
      End
      Begin VB.Menu mnuEditRestore 
         Caption         =   "恢复随访(&P)"
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
      Begin VB.Menu mnuViewFilter 
         Caption         =   "过滤(&F)"
         Shortcut        =   ^F
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
Attribute VB_Name = "frmLaterVisit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'（１）窗体级变量定义**************************************************************************************************
Private mblnStartUp As Boolean                          '窗体启动标志
Private WithEvents PopMenu As clsPopMenu                '自定义弹出菜单对象
Attribute PopMenu.VB_VarHelpID = -1
Private mbytPopMenu As Byte
Private mlngSvrKey(0 To 2)  As Long                     '用于保存各个区域选中的行关键字
Private mintIndex As Integer
Private mlngLoop As Long
Private mlng结论id As Long
Private mblnDataMoved As Boolean

Private Type CONDITION
    结论id As Long
    随访期内 As Boolean             '只显示随访期内的人员
    开始时间 As String              '历史随访人员的体检开始时间
    结束时间 As String              '历史随访人员的体检结束时间
    随访人员 As Boolean             '只显示当前要随访的人员,前提是随访期内为真的
End Type

Private mConditon As CONDITION

Private Enum mCol
    状态 = 0
    姓名 = 1
    门诊号
    性别
    单位
    体检单号
    随访开始
    随访期限
    上次随访
    随访终止
End Enum

'（２）自定义过程或函数************************************************************************************************

Private Function SaveRow(ByVal objVsf As Object) As String
    SaveRow = objVsf.RowData(objVsf.Row)
End Function

Private Sub InheritRestoreRow(ByVal objVsf As Object, ByVal strKey As String)
    '--------------------------------------------------------------------------------------------------------
    '功能:继承RestoreRow过程
    '参数:
    '返回:
    '--------------------------------------------------------------------------------------------------------
    On Error Resume Next
        
    Call RestoreRow(objVsf, Val(strKey))
    
End Sub


Private Function InitLoad() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能:初始化数据，发生在窗体的Load事件
    '------------------------------------------------------------------------------------------------------------------
    Dim strVsf As String
    
    On Error GoTo errHand
    
    mintIndex = 0
    
    mConditon.结论id = 0
    mConditon.随访期内 = True
    mConditon.随访人员 = True
    mConditon.开始时间 = Format(zlDatabase.Currentdate - 30, "yyyy-MM-dd")
    mConditon.结束时间 = Format(zlDatabase.Currentdate, "yyyy-MM-dd")
    
    mlngSvrKey(0) = 0
    mlngSvrKey(1) = 0
    mlngSvrKey(2) = 0
        
    strVsf = ",255,4,1,1,[状态];姓名,900,1,1,1,;门诊号,900,1,1,1,;性别,450,1,1,1,;单位,1500,1,1,1,;体检单号,900,1,1,1,;随访开始,990,1,1,1,;随访期限,900,1,1,1,;上次随访,990,1,1,1,;随访终止,990,1,1,1,"
    Call CreateVsf(vsf(0), strVsf)
    vsf(0).Cols = vsf(0).Cols + 1
    vsf(0).ColWidth(vsf(0).Cols - 1) = 15
    Set vsf(0).Cell(flexcpPicture, 0, 0) = ils13.ListImages("状态").Picture
    
    vsf(0).ColFormat(mCol.随访开始) = "yyyy-MM-dd"
    vsf(0).ColFormat(mCol.上次随访) = "yyyy-MM-dd"
    vsf(0).ColFormat(mCol.随访终止) = "yyyy-MM-dd"
        
    strVsf = ",255,4,1,1,[状态];No,1500,1,1,1,;随访日期,1200,1,1,1,;随访人,900,1,1,1,;登记时间,1800,1,1,1,"
    Call CreateVsf(vsf(2), strVsf)
    vsf(2).Cols = vsf(2).Cols + 1
    vsf(2).ColWidth(vsf(2).Cols - 1) = 15
    Set vsf(2).Cell(flexcpPicture, 0, 0) = ils13.ListImages("状态").Picture
        
    vsfContent.FixedRows = 0
    vsfContent.FixedCols = 0
    vsfContent.Rows = 1
    vsfContent.Cols = 3
    vsfContent.ColWidth(0) = 450
    vsfContent.ColWidth(1) = 600
    vsfContent.ColWidth(2) = 900
    vsfContent.ColAlignment(1) = flexAlignLeftTop
    vsfContent.ColAlignment(2) = flexAlignLeftTop
    
    vsfContent.ExtendLastCol = True
    
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
    'strPrivilege = "基本;随访;结束随访;恢复随访"
    
    '不具有“随访人员”权限时
    
    strPrivilege = ";" & strPrivilege & ";"
    
    '不具有“随访”权限时
    If InStr(strPrivilege, ";随访;") = 0 Then
        mnuEditAdd.Visible = False
        mnuEditModify.Visible = False
        mnuEditDelete.Visible = False
    End If
    
    '不具有“结束随访”权限时
    If InStr(strPrivilege, "结束随访") = 0 Then
        mnuEditStop.Visible = False
    End If
    
    '不具有“恢复随访”权限时
    If InStr(strPrivilege, "恢复随访") = 0 Then
        If InStr(strPrivilege, "结束随访") = 0 And InStr(strPrivilege, "随访") = 0 Then
            mnuEdit.Visible = False
        Else
            mnuEditRestore.Visible = False
        End If
    End If
        
    mnuEdit_1.Visible = mnuEditAdd.Visible And (mnuEditStop.Visible Or mnuEditRestore.Visible)
    
    tbrThis.Buttons("增加").Visible = mnuEditAdd.Visible
    tbrThis.Buttons("修改").Visible = mnuEditModify.Visible
    tbrThis.Buttons("删除").Visible = mnuEditDelete.Visible
    
    tbrThis.Buttons("结束").Visible = mnuEditStop.Visible
    tbrThis.Buttons("恢复").Visible = mnuEditRestore.Visible
    
    tbrThis.Buttons("Split_2").Visible = mnuEditAdd.Visible
    tbrThis.Buttons("Split_3").Visible = mnuEditStop.Visible Or mnuEditRestore.Visible

End Sub

Private Sub AdjustEnableState()
    '------------------------------------------------------------------------------------------------------------------
    '功能： 调整各功能菜单的可用状态
    '------------------------------------------------------------------------------------------------------------------
    
    mnuFilePrint.Enabled = True
    mnuFilePrintView.Enabled = True
    mnuFileOutExcel.Enabled = True
    
    mnuEditAdd.Enabled = True
    mnuEditModify.Enabled = True
    mnuEditDelete.Enabled = True
    mnuEditStop.Enabled = True
    mnuEditRestore.Enabled = True
    
    If Val(vsf(mintIndex).RowData(1)) = 0 Then
        mnuEditAdd.Enabled = False
        mnuEditModify.Enabled = False
        mnuEditDelete.Enabled = False
        mnuEditStop.Enabled = False
        mnuEditRestore.Enabled = False
    Else
        If mintIndex = 1 Then
            mnuEditAdd.Enabled = False
            mnuEditModify.Enabled = False
            mnuEditDelete.Enabled = False
'            mnuEditStop.Enabled = False
        Else
'            mnuEditRestore.Enabled = False
        End If
        
    End If
    
    If vsf(0).TextMatrix(vsf(0).Row, mCol.状态) = "people" Then
        mnuEditRestore.Enabled = False
    Else
        mnuEditStop.Enabled = False
        
    End If
    
    If Val(vsf(2).RowData(1)) = 0 Then
    
        mnuFilePrint.Enabled = False
        mnuFilePrintView.Enabled = False
        mnuFileOutExcel.Enabled = False
        
        mnuEditModify.Enabled = False
        mnuEditDelete.Enabled = False
    End If
    
    
    tbrThis.Buttons("预览").Enabled = mnuFilePrintView.Enabled
    tbrThis.Buttons("打印").Enabled = mnuFilePrint.Enabled
    
    tbrThis.Buttons("增加").Enabled = mnuEditAdd.Enabled
    tbrThis.Buttons("修改").Enabled = mnuEditModify.Enabled
    tbrThis.Buttons("删除").Enabled = mnuEditDelete.Enabled
    
    tbrThis.Buttons("结束").Enabled = mnuEditStop.Enabled
    tbrThis.Buttons("恢复").Enabled = mnuEditRestore.Enabled
        
End Sub

Private Sub RefreshStateInfo()
    '------------------------------------------------------------------------------------------------------------------
    '功能：刷新状态栏显示信息
    '------------------------------------------------------------------------------------------------------------------
    Dim strInfo As String
    Dim lngIndex As Long
    Dim lngLoop As Long
        
    On Error Resume Next

    
    Select Case mintIndex
    Case 0
        
        If Val(vsf(mintIndex).RowData(1)) = 0 Then
            strInfo = strInfo & "没有随访期中的人员。"
        Else
            strInfo = strInfo & "有" & vsf(mintIndex).Rows - 1 & "个随访期中的人员。"
        End If
        
    Case 1
        If Val(vsf(mintIndex).RowData(1)) = 0 Then
            strInfo = strInfo & "没有随访结束的人员。"
        Else
            strInfo = strInfo & "有" & vsf(mintIndex).Rows - 1 & "个随访结束的人员。"
        End If
    End Select
 
    stbThis.Panels(2).Text = "在" & vsf(mintIndex).Tag & "期间" & strInfo
End Sub

Public Function EditRefresh(ByVal strMenuItem As String, ByVal strNO As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：供编辑数据窗体调用，接口函数
    '------------------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    
    On Error GoTo errHand

    Select Case strMenuItem
    Case "随访记录"
        
        Call ClearData("随访记录")
        Call ClearData("随访情况")
      
        Call RefreshData("随访记录")
        
        lngRow = vsf(2).FindRow(strNO, , 1, , True)
        If lngRow > 0 Then
            vsf(2).Row = lngRow
            vsf(2).ShowCell vsf(2).Row, vsf(2).Col
        End If
        
        Call RefreshData("随访情况")
    
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
    
    Select Case strMenuItem
    Case "随访记录"
        Call ResetVsf(vsf(2))
    Case "随访情况"
        vsfContent.Rows = 1
        vsfContent.RowData(0) = 0
        vsfContent.Cell(flexcpText, 0, 0, 0, vsfContent.Cols - 1) = ""
    Case "正在随访"
        Call ResetVsf(vsf(0))
        
    End Select
        
End Function

Private Function RefreshData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '功能：刷新/装载数据
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    Dim strStart As String
    Dim strEnd As String
    Dim lngTime As Long
    Dim strSQL As String
    Dim blnDataMoved As Boolean
    Dim strTmp As String
    
    On Error GoTo errHand
    
    lngTime = Val(GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "随访间隔", 1)) * 30
        
    If mConditon.结论id > 0 Then
        strSQL = " And B.体检病历id In (Select 病历记录id From 病人病历内容 a,病历元素目录 b,体检人员结论 c where c.记录性质=0 And c.病历id=a.id and a.元素编码=b.编码 and Upper(b.部件)='ZL9CISCORE.USRMEDICALSUM' and 结论id=" & mConditon.结论id & ")"
    End If
    
    If mConditon.随访期内 = False Then
        strSQL = strSQL & " And C.体检时间>=TO_DATE('" & mConditon.开始时间 & "','yyyy-mm-dd hh24:mi:ss') AND C.体检时间<=TO_DATE('" & mConditon.结束时间 & "','yyyy-mm-dd hh24:mi:ss')"
        vsf(0).Tag = Format(mConditon.开始时间, "yyyy-MM-dd") & "～" & Format(mConditon.结束时间, "yyyy-MM-dd")
    Else
        strSQL = strSQL & " And B.随访开始+B.随访期限*30>=SYSDATE And B.随访终止 IS NULL "
    End If
    
    If mConditon.随访人员 Then
        strSQL = strSQL & " AND B.随访时间+" & lngTime & "<=SYSDATE "
    End If
    
    Select Case strMenuItem
    Case "正在随访"
                
        gstrSQL = "SELECT   A.病人ID AS ID," & _
                            "A.姓名," & _
                            "A.性别,A.门诊号," & _
                            "A.出生日期," & _
                            "B.随访时间 AS 上次随访," & _
                            "Decode(B.随访期限,null,'',0,'',trim(to_char(B.随访期限))||'月') as 随访期限," & _
                            "B.随访终止," & _
                            "B.随访开始," & _
                            "C.体检号 AS 体检单号," & _
                            "Decode(SIGN(SYSDATE - (B.随访开始+B.随访期限*30)),1,'people1',Decode(B.随访终止,NULL,'people','people1')) AS 状态 " & _
                    "FROM 病人信息 A," & _
                         "体检人员档案 B, " & _
                         "体检登记记录 C " & _
                    "WHERE A.病人ID = B.病人ID " & _
                          "AND B.登记id=C.ID " & _
                          "AND B.随访开始 IS NOT NULL " & strSQL
        
        '数据转储处理
        '--------------------------------------------------------------------------------------------------------------
        blnDataMoved = False
        If mConditon.随访期内 = False Then
            blnDataMoved = zlDatabase.DateMoved(Format(mConditon.开始时间, "yyyy-MM-dd HH:mm:ss"), , glngSys, Me.Caption)
        Else
            blnDataMoved = True
        End If
        If blnDataMoved Then
            strTmp = gstrSQL
            strTmp = Replace(strTmp, "体检人员档案", "H体检人员档案")
            strTmp = Replace(strTmp, "体检登记记录", "H体检登记记录")
            strTmp = Replace(strTmp, "病人病历内容", "H病人病历内容")
            strTmp = Replace(strTmp, "体检人员结论", "H体检人员结论")
            gstrSQL = gstrSQL & " Union All " & strTmp
        End If
        
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
         If rs.BOF = False Then
            Call LoadGrid(vsf(0), rs, Array(), , ils13)
        End If

        
    Case "随访记录"
        
        gstrSQL = "select DISTINCT 1 AS ID," & _
                         "NO," & _
                         "TO_CHAR(随访时间, 'yyyy-mm-dd') AS 随访日期," & _
                         "随访人," & _
                         "TO_CHAR(登记时间, 'yyyy-mm-dd hh24:mi') AS 登记时间," & _
                         "'bill' AS 状态 " & _
                    "from 体检随访记录 " & _
                    "WHERE 病人id=[1] " & _
                            "AND 体检单号=[2]"
                            
        '数据转储处理
        '--------------------------------------------------------------------------------------------------------------
        mblnDataMoved = DataMove(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "体检单号")), 3)
        If mblnDataMoved Then
            gstrSQL = Replace(gstrSQL, "体检随访记录", "H体检随访记录")
        End If
        
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsf(mintIndex).RowData(vsf(mintIndex).Row)), vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "体检单号")))
        If rs.BOF = False Then
            Call LoadGrid(vsf(2), rs, , , ils13)
        End If
        
        Call vsf_AfterRowColChange(2, 0, 0, vsf(2).Row, vsf(2).Col)
        
    Case "随访情况"
        
        vsfContent.Rows = 1
        vsfContent.RowData(0) = 0
        vsfContent.Cell(flexcpText, 0, 0, 0, vsfContent.Cols - 1) = ""
        
        gstrSQL = "select * from (" & _
                    "select 诊断描述," & _
                           "诊断描述 AS 随访结果," & _
                           "诊断描述 AS 随访情况," & _
                           "序号 AS 排序1," & _
                           "1 AS 排序2 " & _
                      "From 体检随访记录 where no=[1] " & _
                    "Union All " & _
                    "select '' AS 诊断描述," & _
                           "'结果:' AS 随访结果," & _
                           "DECODE(随访结果,1,'正常',2,'观察',3,'复查',4,'治疗','') AS 随访情况," & _
                           "序号 AS 排序1," & _
                           "2 AS 排序2 " & _
                      "From 体检随访记录 where no=[1] " & _
                    "Union All " & _
                    "select '' AS 诊断描述," & _
                           "'描述:' AS 随访结果," & _
                           "随访情况," & _
                           "序号 AS 排序1," & _
                           "3 AS 排序2 " & _
                      "From 体检随访记录 where no=[1] " & _
                    ") " & _
                    "order by 排序1,排序2 "
                    
        '数据转储处理
        '--------------------------------------------------------------------------------------------------------------
        If mblnDataMoved Then
        
            gstrSQL = Replace(gstrSQL, "体检随访记录", "H体检随访记录")
                        
        End If
        
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, vsf(2).TextMatrix(vsf(2).Row, 1))
        If rs.BOF = False Then
            Do While Not rs.EOF
                
                If Val(vsfContent.RowData(vsfContent.Rows - 1)) = 1 Then
                    vsfContent.Rows = vsfContent.Rows + 1
                End If
                
                vsfContent.RowData(vsfContent.Rows - 1) = 1
                
                vsfContent.TextMatrix(vsfContent.Rows - 1, 1) = zlCommFun.NVL(rs("随访结果").Value, "")
                vsfContent.TextMatrix(vsfContent.Rows - 1, 2) = zlCommFun.NVL(rs("随访情况").Value, "")
                
                Select Case zlCommFun.NVL(rs("排序2").Value, 0)
                Case 1
                    
                    vsfContent.MergeRow(vsfContent.Rows - 1) = True
                    vsfContent.TextMatrix(vsfContent.Rows - 1, 0) = zlCommFun.NVL(rs("排序1").Value) & "、" & zlCommFun.NVL(rs("诊断描述").Value, "")
                    vsfContent.Cell(flexcpFontBold, vsfContent.Rows - 1, 0, vsfContent.Rows - 1, vsfContent.Cols - 1) = True
                    
                Case 3
                    
                    '画线
                    vsfContent.Select vsfContent.Rows - 1, 1, vsfContent.Rows - 1, vsfContent.Cols - 1
                    vsfContent.CellBorder -2147483633, 0, 0, 0, 1, 0, 0
                
                End Select
                
                rs.MoveNext
            Loop
        End If
        
        Call vsfContent.AutoSize(vsfContent.Cols - 1, vsfContent.Cols - 1)
        
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
    Dim rs As New ADODB.Recordset
    Dim strSQL() As String
                    
    Dim lng病人id As Long
    Dim str体检单号 As String
                    
    On Error GoTo errHand
    
    ReDim Preserve strSQL(1 To 1)
    
    lng病人id = Val(vsf(mintIndex).RowData(vsf(mintIndex).Row))
    str体检单号 = vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "体检单号"))
    
    Select Case strMenuItem
    Case "参数设置"
            
        If Not frmLaterVisitPara.ShowPara(Me) Then GoTo Over
        
    Case "条件过滤"
        
        If Not frmLaterVisitFilter.ShowPara(Me, mConditon.结论id, mConditon.随访期内, mConditon.随访人员, mConditon.开始时间, mConditon.结束时间) Then GoTo Over
        
    Case "增加随访"
        
        '数据转储处理
        '--------------------------------------------------------------------------------------------------------------
        If mblnDataMoved Then
            ShowSimpleMsg "此随访记录已经转出，不能再操作。"
            GoTo Over
        End If
        
        If Not frmLaterVisitEdit.ShowEdit(Me, lng病人id & "'" & str体检单号 & "'") Then GoTo Over
        
    Case "修改随访"
        
        If vsf(2).TextMatrix(vsf(2).Row, 1) = "" Then GoTo Over
        
        '数据转储处理
        '--------------------------------------------------------------------------------------------------------------
        If mblnDataMoved Then
            ShowSimpleMsg "此随访记录已经转出，不能再操作。"
            GoTo Over
        End If
        
        If Not frmLaterVisitEdit.ShowEdit(Me, lng病人id & "'" & str体检单号 & "'" & vsf(2).TextMatrix(vsf(2).Row, 1)) Then GoTo Over
        
    Case "删除随访"
        
        If vsf(2).TextMatrix(vsf(2).Row, 1) = "" Then GoTo Over
        
        '数据转储处理
        '--------------------------------------------------------------------------------------------------------------
        If mblnDataMoved Then
            ShowSimpleMsg "此随访记录已经转出，不能再删除。"
            GoTo Over
        End If
        
        If MsgBox("真的要删除“" & vsf(2).TextMatrix(vsf(2).Row, 1) & "”的随访记录吗？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo Then GoTo Over
        
        strSQL(ReDimArray(strSQL)) = "ZL_体检随访记录_DELETE('" & Trim(vsf(2).TextMatrix(vsf(2).Row, 1)) & "')"
        
    Case "结束随访"
    
        If vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, 1) = "" Then GoTo Over
        
        '数据转储处理
        '--------------------------------------------------------------------------------------------------------------
        If mblnDataMoved Then
            ShowSimpleMsg "此随访记录已经转出，不能再操作。"
            GoTo Over
        End If
        
        If MsgBox("真的要结束“" & vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, 1) & "”的随访吗？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo Then GoTo Over
           
        strSQL(ReDimArray(strSQL)) = "ZL_体检人员档案_STOP(" & lng病人id & ",'" & str体检单号 & "')"
        
    Case "恢复随访"
        
        If vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, 1) = "" Then GoTo Over
        
        '数据转储处理
        '--------------------------------------------------------------------------------------------------------------
        If mblnDataMoved Then
            ShowSimpleMsg "此随访记录已经转出，不能再操作。"
            GoTo Over
        End If
        
        If MsgBox("真的要恢复“" & vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, 1) & "”的随访吗？", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo Then GoTo Over
           
        strSQL(ReDimArray(strSQL)) = "ZL_体检人员档案_RESTORE(" & lng病人id & ",'" & str体检单号 & "')"
                          
    End Select
    
    blnTran = True
    gcnOracle.BeginTrans
    
    For lngLoop = 1 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(lngLoop), Me.Caption)
    Next
    
    gcnOracle.CommitTrans
    blnTran = False
        
    Select Case strMenuItem
    Case "增加随访", "修改随访", "删除随访"
        
        '修改人员列表的上次随访时间
        gstrSQL = "SELECT A.随访时间 FROM 体检人员档案 A,体检登记记录 B WHERE A.登记id=B.ID AND A.病人id=[1] AND B.体检号=[2]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng病人id, str体检单号)
        If rs.BOF = False Then
            vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "上次随访")) = Format(zlCommFun.NVL(rs("随访时间")), "yyyy-MM-dd")
        End If
        
        If strMenuItem = "删除随访" Then
            If vsf(2).Rows = 2 Then
                vsf(2).Cell(flexcpText, 1, 0, 1, vsf(2).Cols - 1) = ""
                vsf(2).RowData(1) = 0
                Set vsf(2).Cell(flexcpPicture, 1, 0) = Nothing
            Else
                vsf(2).RemoveItem vsf(2).Row
            End If
        End If
        
        Call ClearData("随访情况")
        Call RefreshData("随访情况")
        
    Case "结束随访", "恢复随访", "参数设置", "条件过滤"
        
        Call mnuViewRefresh_Click
        
    End Select
    
    Call AdjustEnableState
    Call RefreshStateInfo
    
    MenuClick = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
    If blnTran Then gcnOracle.RollbackTrans
    
Over:

End Function

Private Sub PrintData(ByVal bytMode As Byte)
    '------------------------------------------------------------------------------------------------------------------
    '功能： 打印数据
    '参数： bytMode                         打印方式（1-打印；2-预览；3-输出到Excel）
    '------------------------------------------------------------------------------------------------------------------
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim lngCX As Long

    If UserInfo.姓名 = "" Then Call GetUserInfo

    objPrint.Title.Text = "体检随访记录"
    
    Set objRow = New zlTabAppRow
    objRow.Add ""
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "病人:" & vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "姓名")) & " 体检单号:" & vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "体检单号"))
    objRow.Add ""
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "随访日期:" & vsf(2).TextMatrix(vsf(2).Row, GetCol(vsf(2), "随访日期")) & " 随访人:" & vsf(2).TextMatrix(vsf(2).Row, GetCol(vsf(2), "随访人"))
    objRow.Add ""
    objPrint.UnderAppRows.Add objRow
    
    lngCX = vsfContent.ColWidth(2)
    
    vsfContent.ColWidth(2) = 8100
    vsfContent.AutoSize vsfContent.Cols - 1, vsfContent.Cols - 1
    
    Set objPrint.Body = vsfContent

    If bytMode = 1 Then bytMode = zlPrintAsk(objPrint)

    If bytMode >= 1 And bytMode <= 3 Then Call zlPrintOrView1Grd(objPrint, bytMode)
    
    vsfContent.ColWidth(2) = lngCX
    vsfContent.AutoSize vsfContent.Cols - 1, vsfContent.Cols - 1
End Sub


Private Sub cmdMenu_Click()
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(cmdMenu.hWnd, objPoint)
    
    mbytPopMenu = 3
    Set PopMenu = New clsPopMenu
    Call PopMenu.ShowPopupMenu(objPoint.X * Screen.TwipsPerPixelX, objPoint.Y * Screen.TwipsPerPixelY - 1845)
    
    txt(1).Text = ""
    LocationObj txt(1)
    
End Sub

'（３）窗体及其控件的事件处理******************************************************************************************
Private Sub Form_Activate()
    
    If mblnStartUp = False Then Exit Sub
    DoEvents
    
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
    
    If imgY_S.Left > Me.ScaleWidth - 1000 Then imgY_S.Left = Me.ScaleWidth - 1000
        
    With vsf(0)
        .Left = 0
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0)
        .Width = imgY_S.Left
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0) - fraInfo.Height + 120
    End With
    
    With fraInfo
        .Left = 0
        .Top = vsf(0).Top + vsf(0).Height - 120
        .Width = vsf(0).Width
    End With
    
    With txt(1)
        .Width = fraInfo.Width - .Left - 75
    End With
    
    With imgY_S
        .Top = vsf(0).Top
        .Height = vsf(0).Height
    End With
    
    With vsf(2)
        .Left = imgY_S.Left + imgY_S.Width
        .Top = vsf(0).Top
        .Width = Me.ScaleWidth - .Left
        .Height = imgX_S.Top - .Top
    End With
    
    With imgX_S
        .Left = vsf(2).Left
        .Width = vsf(2).Width
    End With
    
    With vsfContent
        .Left = vsf(2).Left
        .Top = imgX_S.Top + imgX_S.Height
        .Width = vsf(2).Width
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
    
    Call vsfContent.AutoSize(vsfContent.Cols - 1, vsfContent.Cols - 1)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    
    Cancel = mblnStartUp
    If Cancel Then Exit Sub
    
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


Private Sub mnuEditStop_Click()
    Call MenuClick("结束随访")
End Sub

Private Sub mnuEditRestore_Click()
    Call MenuClick("恢复随访")
End Sub

Private Sub mnuEditDelete_Click()
    Call MenuClick("删除随访")
End Sub

Private Sub mnuEditModify_Click()
    Call MenuClick("修改随访")
End Sub

Private Sub mnuEditAdd_Click()
    Call MenuClick("增加随访")
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileOutExcel_Click()
    Call PrintData(3)
End Sub

Private Sub mnuFilePara_Click()
    Call MenuClick("参数设置")
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

Private Sub mnuViewFilter_Click()
    Call MenuClick("条件过滤")
End Sub

Private Sub mnuViewRefresh_Click()
    Dim strSvrKey As String
        
    zlCommFun.ShowFlash "请稍候，正在刷新数据...", Me
    DoEvents

    strSvrKey = SaveRow(vsf(mintIndex))
    
    LockWindowUpdate vsf(mintIndex).hWnd
    
    Call ClearData("正在随访")
    Call ClearData("随访记录")
    Call ClearData("随访情况")
    
    Call RefreshData("正在随访")
    
    Call InheritRestoreRow(vsf(mintIndex), Val(strSvrKey))
    
    LockWindowUpdate 0
    
    zlCommFun.StopFlash
    
    mlngSvrKey(mintIndex) = -1
    Call vsf_AfterRowColChange(mintIndex, 0, 0, vsf(mintIndex).Row, vsf(mintIndex).Col)
    
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

Private Sub PopMenu_MenuBeforeShow(Cancel As Boolean)
    
    Select Case mbytPopMenu
    Case 1
        If mnuEditAdd.Visible Then PopMenu.Add 1, mnuEditAdd.Caption, , , mnuEditAdd.Enabled
        If mnuEditModify.Visible Then PopMenu.Add 2, mnuEditModify.Caption, , , mnuEditModify.Enabled
        If mnuEditDelete.Visible Then PopMenu.Add 3, mnuEditDelete.Caption, , , mnuEditDelete.Enabled
    Case 3
        
        PopMenu.Add 1, "&1.姓名", , , True, , (lbl(1).Tag = "姓名")
        PopMenu.Add 2, "&2.性别", , , True, , (lbl(1).Tag = "性别")
        PopMenu.Add 3, "&3.单位", , , True, , (lbl(1).Tag = "单位")
        PopMenu.Add 4, "-", , 2, True
        PopMenu.Add 5, "&4.体检单号", , , True, , (lbl(1).Tag = "体检单号")
        PopMenu.Add 6, "&5.门诊号", , , True, , (lbl(1).Tag = "门诊号")
        
    End Select
    
End Sub

Private Sub PopMenu_MenuClick(ByVal Key As Long, ByVal Caption As String)
    Select Case mbytPopMenu
    Case 1
        Select Case Key
        Case 1
            Call mnuEditAdd_Click
        Case 2
            Call mnuEditModify_Click
        Case 3
            Call mnuEditDelete_Click
        End Select
    Case 3
    
        Caption = Mid(Caption, 4)
        
        lbl(1).Caption = "&6." & Left(Trim(Caption), Len(Trim(Caption)) - 1)
        lbl(1).Tag = Left(Trim(Caption), Len(Trim(Caption)) - 1)
        
    End Select
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(tbrThis.hWnd, objPoint)
    
    Select Case Button.Key
    Case "预览"
        Call mnuFilePrintView_Click
    Case "打印"
        Call mnuFilePrint_Click
    Case "增加"
        Call mnuEditAdd_Click
    Case "修改"
        Call mnuEditModify_Click
    Case "删除"
        Call mnuEditDelete_Click
    Case "结束"
        Call mnuEditStop_Click
    Case "恢复"
        Call mnuEditRestore_Click
    Case "帮助"
        Call mnuHelpTopic_Click
    Case "退出"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tbrThis_ButtonDropDown(ByVal Button As MSComctlLib.Button)
    Call tbrThis_ButtonClick(Button)
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool
End Sub

Private Sub txt_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt(Index)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim lngLoop As Long
    Dim strCol As String
    Dim lngCol As Long
    
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
        
    Dim lngRow As Long
    
    If KeyAscii = vbKeyReturn Then
        
        If Index = 1 And Trim(txt(Index).Text) <> "" Then
            
            Call txt_LostFocus(Index)
            
            strCol = Mid(lbl(1).Caption, 4)
            lngCol = GetCol(vsf(mintIndex), strCol)
            
            lngRow = 0
            If vsf(mintIndex).Row + 1 <= vsf(mintIndex).Rows - 1 Then
                For lngLoop = vsf(mintIndex).Row + 1 To vsf(mintIndex).Rows - 1
                    If InStr(vsf(mintIndex).TextMatrix(lngLoop, lngCol), txt(Index).Text) > 0 Then
                        lngRow = lngLoop
                        Exit For
                    End If
                Next
            End If
            
            If lngRow = 0 Then
                For lngLoop = 1 To vsf(mintIndex).Row
                    If InStr(vsf(mintIndex).TextMatrix(lngLoop, lngCol), txt(Index).Text) > 0 Then
                        lngRow = lngLoop
                        Exit For
                    End If
                Next
            End If
            
            If lngRow <= 0 Then
                ShowSimpleMsg "没有找到符合要求的信息！"
                txt(Index).Text = ""
            Else
                vsf(mintIndex).ShowCell lngRow, vsf(mintIndex).Col
                vsf(mintIndex).Row = lngRow
            End If
        End If
        
        txt(Index).SetFocus
        zlControl.TxtSelAll txt(Index)
    Else
        If Index = 1 Then
            Select Case lbl(1).Tag
            Case "体检单号"
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
            End Select
        End If
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    
    If lbl(1).Tag = "体检单号" Then
        Dim intYear As Integer
        Dim strYear As String
        '自动补齐单据号
        If (UCase(Left(txt(Index).Text, 1)) < "A" Or UCase(Left(txt(Index).Text, 1)) > "Z") And Trim(txt(Index).Text) <> "" Then
            intYear = Format(zlDatabase.Currentdate, "YYYY") - 1990
            strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
            txt(Index).Text = strYear & Right("0000000" & txt(Index).Text, 7)
        End If
    End If
End Sub


Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
        
    If OldRow = NewRow Then Exit Sub
    
    Call SelectRow(vsf(Index), OldRow, NewRow)
    
    mlngSvrKey(Index) = Val(vsf(Index).RowData(NewRow))
    
    Select Case Index
    Case 0, 1
        Call ClearData("随访记录")
        Call ClearData("随访情况")
            
        Call RefreshData("随访记录")
        Call RefreshData("随访情况")
    Case 2
        Call ClearData("随访情况")
        
        Call RefreshData("随访情况")
    End Select
    
    Call AdjustEnableState
    
End Sub

Private Sub vsf_BeforeUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = (Col = 0)
End Sub

Private Sub vsf_DblClick(Index As Integer)
    If mnuEditModify.Visible And mnuEdit.Visible And mnuEditModify.Enabled Then
        Call mnuEditModify_Click
    End If
End Sub

Private Sub vsf_GotFocus(Index As Integer)
    vsf(Index).BackColorSel = COLOR.焦点
    Call SelectRow(vsf(Index), 1, vsf(Index).Row)
End Sub

Private Sub vsf_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Call vsf_DblClick(2)
    End If
End Sub

Private Sub vsf_LostFocus(Index As Integer)
    vsf(Index).BackColorSel = COLOR.非焦点
    Call SelectRow(vsf(Index), 1, vsf(Index).Row)
End Sub

Private Sub vsf_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Index = 2 Then
        If Button <> 2 Then Exit Sub

        Call SendLMouseButton(vsf(Index).hWnd, X, Y)
        
        mbytPopMenu = 1
        Set PopMenu = New clsPopMenu
        Call PopMenu.ShowPopupMenuByCursor
        
    End If
End Sub


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '功能:链接到中联论坛
    '修改人:刘兴宏
    '修改日期:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub


