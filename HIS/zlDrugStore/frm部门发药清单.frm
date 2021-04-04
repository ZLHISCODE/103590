VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frm部门发药清单 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   7725
   ClientLeft      =   105
   ClientTop       =   105
   ClientWidth     =   12825
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7725
   ScaleWidth      =   12825
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox picInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   7215
      Left            =   8880
      ScaleHeight     =   7185
      ScaleWidth      =   3705
      TabIndex        =   17
      Top             =   120
      Width           =   3735
      Begin VB.PictureBox picHscSend 
         BackColor       =   &H00FFEDDD&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   3735
         TabIndex        =   25
         Tag             =   "0"
         Top             =   0
         Width           =   3735
         Begin VB.CheckBox chk所有诊断 
            BackColor       =   &H00FFEDDD&
            Caption         =   "所有"
            Height          =   180
            Left            =   2520
            TabIndex        =   26
            Top             =   30
            Width           =   735
         End
         Begin VB.Label lblDiag 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEDDD&
            Caption         =   "临床诊断"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   120
            TabIndex        =   27
            Top             =   30
            Width           =   2280
         End
      End
      Begin VB.PictureBox Pic用药理由 
         BackColor       =   &H00FFEDDD&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   3735
         TabIndex        =   23
         Tag             =   "0"
         Top             =   1800
         Width           =   3735
         Begin VB.Label lbl用药理由 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEDDD&
            Caption         =   "抗菌药物相关信息"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   600
            TabIndex        =   24
            Top             =   0
            Width           =   1680
         End
      End
      Begin VB.TextBox txt用药理由 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   1215
         Left            =   0
         MultiLine       =   -1  'True
         TabIndex        =   22
         Top             =   2160
         Width           =   3735
      End
      Begin VB.PictureBox picDoctor 
         BackColor       =   &H00FFEDDD&
         BorderStyle     =   0  'None
         Height          =   255
         Left            =   0
         ScaleHeight     =   255
         ScaleWidth      =   3735
         TabIndex        =   20
         Tag             =   "0"
         Top             =   3480
         Width           =   3735
         Begin VB.Label lblDoctor 
            Alignment       =   2  'Center
            AutoSize        =   -1  'True
            BackColor       =   &H00FFEDDD&
            Caption         =   "开单医生签名图片"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   780
            TabIndex        =   21
            Top             =   0
            Width           =   1440
         End
      End
      Begin VB.PictureBox pic签名图片 
         AutoRedraw      =   -1  'True
         BackColor       =   &H8000000B&
         Height          =   1050
         Left            =   960
         ScaleHeight     =   66
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   90
         TabIndex        =   19
         Top             =   3960
         Visible         =   0   'False
         Width           =   1410
      End
      Begin VB.PictureBox picSign 
         AutoRedraw      =   -1  'True
         Height          =   210
         Left            =   2640
         ScaleHeight     =   10
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   18
         TabIndex        =   18
         Top             =   4200
         Visible         =   0   'False
         Width           =   330
      End
      Begin VSFlex8Ctl.VSFlexGrid vsf诊断 
         Height          =   1335
         Left            =   0
         TabIndex        =   28
         Top             =   360
         Width           =   3720
         _cx             =   6562
         _cy             =   2355
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
         BackColorSel    =   16761024
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
         Rows            =   5
         Cols            =   2
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frm部门发药清单.frx":0000
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
   Begin VB.Frame fraH 
      BackColor       =   &H80000007&
      Height          =   5895
      Left            =   6600
      MousePointer    =   9  'Size W E
      TabIndex        =   16
      Top             =   1680
      Width           =   15
   End
   Begin VB.PictureBox picHsc 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000010&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   0
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   5535
      TabIndex        =   13
      TabStop         =   0   'False
      Top             =   2880
      Visible         =   0   'False
      Width           =   5535
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   960
      Index           =   0
      Left            =   0
      TabIndex        =   11
      Top             =   0
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
      BackColorSel    =   13487565
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
      SelectionMode   =   3
      GridLines       =   1
      GridLinesFixed  =   1
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
      FormatString    =   $"frm部门发药清单.frx":003D
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
      OwnerDraw       =   1
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
   Begin VSFlex8Ctl.VSFlexGrid vsfList 
      Height          =   960
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   0
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
      BackColorSel    =   13487565
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
      FormatString    =   $"frm部门发药清单.frx":00B2
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
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSComctlLib.ImageList imgGroup 
      Left            =   5040
      Top             =   1200
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
            Picture         =   "frm部门发药清单.frx":0127
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":0281
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":03DB
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picAssist 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   9375
      TabIndex        =   6
      Top             =   2160
      Visible         =   0   'False
      Width           =   9375
      Begin VB.ComboBox cbo核查人 
         Height          =   300
         Left            =   3480
         TabIndex        =   15
         Text            =   "cbo核查人"
         Top             =   60
         Width           =   1900
      End
      Begin VB.ComboBox cbo配药人 
         Height          =   300
         Left            =   600
         TabIndex        =   8
         TabStop         =   0   'False
         Text            =   "cbo配药人"
         Top             =   60
         Width           =   1900
      End
      Begin VB.ComboBox cbo发药单格式 
         Height          =   300
         Left            =   6960
         Style           =   2  'Dropdown List
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   60
         Width           =   1905
      End
      Begin VB.Label lbl核查人 
         Caption         =   "核查人"
         Height          =   180
         Left            =   2880
         TabIndex        =   14
         Top             =   120
         Width           =   615
      End
      Begin VB.Label lbl配药人 
         AutoSize        =   -1  'True
         Caption         =   "配药人"
         Height          =   180
         Left            =   0
         TabIndex        =   10
         Top             =   120
         Width           =   540
      End
      Begin VB.Label lbl发药单格式 
         AutoSize        =   -1  'True
         Caption         =   "发药单格式"
         Height          =   180
         Left            =   6000
         TabIndex        =   9
         Top             =   120
         Width           =   900
      End
   End
   Begin VB.Frame fraColSel 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   4500
      TabIndex        =   0
      Top             =   300
      Width           =   195
      Begin VB.Image imgColSel 
         Height          =   195
         Left            =   0
         Picture         =   "frm部门发药清单.frx":0535
         ToolTipText     =   "选择需要显示的列(ALT+C)"
         Top             =   0
         Width           =   195
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfColSel 
      Height          =   855
      Left            =   2040
      TabIndex        =   1
      Top             =   0
      Visible         =   0   'False
      Width           =   1470
      _cx             =   2593
      _cy             =   1508
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
      FormatString    =   $"frm部门发药清单.frx":0A83
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
      Index           =   2
      Left            =   0
      TabIndex        =   3
      Top             =   0
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
      BackColorSel    =   13487565
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
      FormatString    =   $"frm部门发药清单.frx":0AD1
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
      Index           =   3
      Left            =   0
      TabIndex        =   4
      Top             =   0
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
      BackColorSel    =   13487565
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
      FormatString    =   $"frm部门发药清单.frx":0B46
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
      Index           =   4
      Left            =   0
      TabIndex        =   5
      Top             =   0
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
      BackColorSel    =   13487565
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
      FormatString    =   $"frm部门发药清单.frx":0BBB
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
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin MSComctlLib.ImageList imgCheck 
      Left            =   5040
      Top             =   0
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
            Picture         =   "frm部门发药清单.frx":0C30
            Key             =   ""
            Object.Tag             =   "1"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":11CA
            Key             =   ""
            Object.Tag             =   "2"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":1764
            Key             =   ""
            Object.Tag             =   "3"
         EndProperty
      EndProperty
   End
   Begin VSFlex8Ctl.VSFlexGrid vsfChargeOff 
      Height          =   960
      Left            =   120
      TabIndex        =   12
      Top             =   1080
      Visible         =   0   'False
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
      BackColorSel    =   13487565
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
      FormatString    =   $"frm部门发药清单.frx":18BE
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
   Begin MSComctlLib.ImageList imgList 
      Left            =   5040
      Top             =   600
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   40
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":1933
            Key             =   "打印11"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":1CCD
            Key             =   "当前"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":852F
            Key             =   "指示器"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":ED91
            Key             =   "附件"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":F32B
            Key             =   "报告"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":F6C5
            Key             =   "标志"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":FA5F
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":FDF9
            Key             =   "图标"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":10193
            Key             =   "选择"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":10BA5
            Key             =   "Person"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":17407
            Key             =   "未检"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":1DC69
            Key             =   "在检"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":244CB
            Key             =   "已检"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":2AD2D
            Key             =   "Family"
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":3158F
            Key             =   "分类"
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":31929
            Key             =   "分类_选中"
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":31CC3
            Key             =   "套餐"
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":38525
            Key             =   "类型"
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":3ED87
            Key             =   "照片"
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":455E9
            Key             =   "参数"
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":4BE4B
            Key             =   "指标"
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":526AD
            Key             =   "体检"
         EndProperty
         BeginProperty ListImage23 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":58F0F
            Key             =   "病历样式"
         EndProperty
         BeginProperty ListImage24 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":5F771
            Key             =   "病历文件"
         EndProperty
         BeginProperty ListImage25 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":65FD3
            Key             =   "规则"
         EndProperty
         BeginProperty ListImage26 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":6C835
            Key             =   "收费"
         EndProperty
         BeginProperty ListImage27 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":6D247
            Key             =   "诊断"
         EndProperty
         BeginProperty ListImage28 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":73AA9
            Key             =   "创建"
         EndProperty
         BeginProperty ListImage29 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":7A30B
            Key             =   "确认"
         EndProperty
         BeginProperty ListImage30 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":80B6D
            Key             =   "开始"
         EndProperty
         BeginProperty ListImage31 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":873CF
            Key             =   "结束"
         EndProperty
         BeginProperty ListImage32 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":8DC31
            Key             =   "部份"
         EndProperty
         BeginProperty ListImage33 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":8DFCB
            Key             =   "全部"
         EndProperty
         BeginProperty ListImage34 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":8E365
            Key             =   "部份总检"
         EndProperty
         BeginProperty ListImage35 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":8E6FF
            Key             =   "全部总检"
         EndProperty
         BeginProperty ListImage36 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":8EA99
            Key             =   "总检"
         EndProperty
         BeginProperty ListImage37 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":8EE33
            Key             =   "打印"
         EndProperty
         BeginProperty ListImage38 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":8F845
            Key             =   "已经打印"
         EndProperty
         BeginProperty ListImage39 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":90257
            Key             =   "药品"
         EndProperty
         BeginProperty ListImage40 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm部门发药清单.frx":96AB9
            Key             =   "高危"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   3840
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frm部门发药清单"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'从参数表中取药品价格、数量、金额小数位数
Private mintUnit As Integer                 '单位系数：1-售价;2-门诊;3-住院;4-药库
Private mlngMode As Long

Private mblnOutPut As Boolean

Private mintCostDigit As Integer            '成本价小数位数
Private mintPriceDigit As Integer           '售价小数位数
Private mintNumberDigit As Integer          '数量小数位数
Private mintMoneyDigit As Integer           '金额小数位数

'工具栏菜单
Private Const conMenu_Tool_ShowShortage = 101       '显示缺药
'Private Const conMenu_Tool_ShowRefuse = 102         '显示拒发
Private Const conMenu_Tool_ShowReturnSend = 103     '显示退药待发
Private Const conMenu_Tool_SumByBatch = 104         '按批次汇总
Private Const conMenu_Tool_ShowAllProcess = 105     '显示所有过程单据
Private Const conMenu_Tool_ShowPlug = 106           '调用插件：合理用药
Private Const conMenu_Tool_ShowInfo = 107           '显示扩展信息

'弹出菜单
'发药时
Private Const conMenu_StatusPopup = 3                 '单据状态
Private Const conMenu_Status_Verify = 301             '发药
Private Const conMenu_Status_Reject = 303             '拒发确认
Private Const conMenu_Status_Return = 304             '退药
Private Const conMenu_Status_RefuseRestore = 308      '拒发恢复
Private Const conMenu_Status_Shortage = 309           '缺药
Private Const conMenu_Status_NoProcess = 310          '不处理
Private Const conMenu_Status_AllSend = 311            '全部发药
Private Const conMenu_Status_AllReject = 312          '全部拒发
Private Const conMenu_Status_AllNoProcess = 313       '全部不处理
'退药时
Private Const conMenu_Status_AllReturn = 321          '全部退药
Private Const conMenu_Status_AllCancel = 322          '全部取消退药
'药品名称
Private Const conMenu_MediPopup = 4                   '药品名称显示
Private Const conMenu_Medi_CodeAddName = 401          '显示编码和名称
Private Const conMenu_Medi_Code = 402                 '显示编码
Private Const conMenu_Medi_Name = 403                 '显示名称

'其它定义
Private mdblSumListHeight As Double                 '记录汇总发药列表原始的高度
Private mdblSendListHeight As Double                 '记录发药列表原始的高度
Private mblnResize As Boolean

Private mstrFindCondition As String                 '查找条件

Private mblnShowReject As Boolean

'数据集
Private mrsSendList As ADODB.Recordset              '发药数据集
Private mrsChargeOff As ADODB.Recordset             '销账数据集
Private mrsReturnList As ADODB.Recordset            '退药数据集

'列表显示条件
Private Type Type_ShowListCondition
    intListType As Integer                          '0-未发;1-汇总;2-缺药;3-拒发;4-已发
    bln按批次汇总 As Boolean
    bln按科室汇总 As Boolean
    intShowPass As Integer                           '是否显示合理用药（PASS）
    bln医生查询 As Boolean
    bln显示退药待发单据 As Boolean
    bln显示扩展信息 As Boolean
    bln显示缺药 As Boolean
    bln显示过程单据 As Boolean
    bln显示所有诊断 As Boolean
    lng药房id As Long
    bln修改留存数量 As Boolean
    bln启用退药销账 As Boolean
    bln药品储备 As Boolean
    bln允许未审核处方发药 As Boolean
    int药品名称编码显示 As Integer
    int发药单格式 As Integer
    bln配制中心 As Boolean
    str高危发放 As String
    str高危分类 As String
    int退药待发单据默认为发药状态 As Integer
    bln显示原产地 As Boolean
End Type
Private mcondition As Type_ShowListCondition

'更新标志
Private mblnRefresh As Boolean                      '刷新标志
Private mblnSendChange As Boolean                   '待发药清单中的状态发生变化
Private mblnDrop As Boolean                     '在KeyDown中判断下拉列表是否弹出
Private Const CB_GETDROPPEDSTATE = &H157
Private Const CB_SHOWDROPDOWN = &H14F

'列表类型
Private Enum mListType
    发药 = 0
    汇总 = 1
    缺药 = 2
    拒发 = 3
    退药 = 4
End Enum

'执行状态
Private Enum mState
    缺药 = 0
    发药 = 1
    拒发 = 2
    不处理 = 3
    拒发_恢复 = 4
    拒发_不处理 = 5
    退药 = 6
    退药_原始记录 = 7
    退药_发药记录 = 8
    退药_退药记录 = 9
    转出记录 = 10
End Enum

'通过菜单改变状态的处理
Private Enum mChangeState
    发药 = 0
    拒发 = 1
    缺药 = 2
    不处理 = 3
    全部发药 = 4
    全部拒发 = 5
    全部不处理 = 6
End Enum

'单据汇总类型
Private Enum mSubTotalType
    SubSum = 0                  '合计
    SubByDept = 1               '按领药部门小计
    SubByPeople = 2             '按病人小计
    SubByNo = 3                 '按单据小计
    SubByDrug = 4               '按药品小计
    SubByHosNumber = 5          '按住院号小计
    SubByBedNumber = 6          '按床号小计
    SubByPeopleDept = 7         '按病人科室
End Enum

Private mstrUnallowSetColHide(0 To 4) As String         '不允许设置隐藏的列
Private mstrUnallowShow(0 To 4) As String                   '不允许显示的列

'未发药列表
Private Const mconIntCol发药_列数 As Integer = 61
Private mIntCol发药_当前行 As Integer
Private mIntCol发药_审查结果 As Integer
Private mIntCol发药_分组符 As Integer
Private mIntCol发药_科室 As Integer
Private mIntCol发药_开单医生 As Integer
Private mIntCol发药_状态 As Integer
Private mIntCol发药_类型 As Integer
Private mIntCol发药_发药类型 As Integer
Private mIntCol发药_NO As Integer
Private mIntCol发药_记帐员 As Integer
Private mIntCol发药_床号 As Integer
Private mIntCol发药_病人类型 As Integer
Private mIntCol发药_姓名 As Integer
Private mIntCol发药_性别 As Integer
Private mIntCol发药_年龄 As Integer
Private mIntCol发药_住院号 As Integer
Private mIntCol发药_品名 As Integer
Private mIntCol发药_皮试结果 As Integer
Private mIntCol发药_其它名 As Integer
Private mIntCol发药_英文名 As Integer
Private mIntCol发药_配方名称 As Integer
Private mIntCol发药_规格 As Integer
Private mIntCol发药_生产商 As Integer
Private mIntCol发药_原产地 As Integer
Private mIntCol发药_批号 As Integer
Private mIntCol发药_效期 As Integer
Private mIntCol发药_付 As Integer
Private mIntCol发药_数量 As Integer
Private mIntCol发药_单价 As Integer
Private mIntCol发药_金额 As Integer
Private mIntCol发药_单量 As Integer
Private mIntCol发药_频次 As Integer
Private mIntCol发药_用法 As Integer
Private mIntCol发药_用药次数 As Integer
Private mIntCol发药_用药目的 As Integer
Private mIntCol发药_记帐时间 As Integer
Private mIntCol发药_说明 As Integer
Private mIntCol发药_单据 As Integer
Private mIntCol发药_医嘱id As Integer
Private mIntCol发药_退药人 As Integer
Private mIntCol发药_库房货位 As Integer
Private mIntCol发药_相关ID As Integer
Private mIntCol发药_药品ID As Integer
Private mIntCol发药_单量单位 As Integer
Private mIntCol发药_领药部门 As Integer
Private mIntCol发药_领药部门id As Integer
Private mIntCol发药_药品编码和名称 As Integer
Private mIntCol发药_药品编码 As Integer
Private mIntCol发药_药品名称 As Integer
Private mIntCol发药_收发ID As Integer
Private mIntCol发药_执行状态 As Integer
Private mIntCol发药_领药号 As Integer
Private mIntCol发药_已收费 As Integer
Private mIntCol发药_病人ID As Integer
Private mIntCol发药_主页ID As Integer
Private mIntCol发药_用药理由 As Integer
Private mIntCol发药_高危药品 As Integer
Private mIntCol发药_结果 As Integer
Private mIntCol发药_禁忌药品说明 As Integer
Private mIntCol发药_开单部门id As Integer
Private mIntCol发药_脚注 As Integer

'汇总列表
Private Const mconIntCol汇总_列数 As Integer = 14
Private mIntCol汇总_当前行 As Integer
Private mIntCol汇总_品名 As Integer
Private mIntCol汇总_规格 As Integer
Private mIntCol汇总_生产商 As Integer
Private mIntCol汇总_原产地 As Integer
Private mIntCol汇总_批号 As Integer
Private mIntCol汇总_效期 As Integer
Private mIntCol汇总_数量 As Integer
Private mIntCol汇总_单位 As Integer
Private mIntCol汇总_单价 As Integer
Private mIntCol汇总_金额 As Integer
Private mIntCol汇总_药品编码和名称 As Integer
Private mIntCol汇总_药品编码 As Integer
Private mIntCol汇总_药品名称 As Integer

Private Const mconIntCol科室汇总_列数 As Integer = 25
Private mIntCol科室汇总_当前行 As Integer
Private mIntCol科室汇总_科室 As Integer
Private mIntCol科室汇总_品名 As Integer
Private mIntCol科室汇总_规格 As Integer
Private mIntCol科室汇总_生产商 As Integer
Private mIntCol科室汇总_原产地 As Integer
Private mIntCol科室汇总_批号 As Integer
Private mIntCol科室汇总_效期 As Integer
Private mIntCol科室汇总_应发数量 As Integer
Private mIntCol科室汇总_留存数量 As Integer
Private mIntCol科室汇总_销帐数量 As Integer
Private mIntCol科室汇总_实发数量 As Integer
Private mIntCol科室汇总_单位 As Integer
Private mIntCol科室汇总_单价 As Integer
Private mIntCol科室汇总_应发金额 As Integer
Private mIntCol科室汇总_实发金额 As Integer
Private mIntCol科室汇总_批次 As Integer
Private mIntCol科室汇总_科室ID As Integer
Private mIntCol科室汇总_药品ID As Integer
Private mIntCol科室汇总_领药部门 As Integer
Private mIntCol科室汇总_领药部门id As Integer
Private mIntCol科室汇总_药品编码和名称 As Integer
Private mIntCol科室汇总_药品编码 As Integer
Private mIntCol科室汇总_药品名称 As Integer
Private mIntCol科室汇总_包装 As Integer

'销账列表
Private Const mconIntCol销账_列数 As Integer = 14
Private mIntCol销账_当前行 As Integer
Private mIntCol销账_申请科室 As Integer
Private mIntCol销账_单据 As Integer
Private mIntCol销账_NO As Integer
Private mIntCol销账_药品ID As Integer
Private mIntCol销账_申请时间 As Integer
Private mIntCol销账_收发序号 As Integer
Private mIntCol销账_生产商 As Integer
Private mIntCol销账_批号 As Integer
Private mIntCol销账_效期 As Integer
Private mIntCol销账_准退数量 As Integer
Private mIntCol销账_销帐数量 As Integer
Private mIntCol销账_包装 As Integer
Private mIntCol销账_单位 As Integer

'缺药列表
Private Const mconIntCol缺药_列数 As Integer = 21
Private mIntCol缺药_当前行 As Integer
Private mIntCol缺药_科室 As Integer
Private mIntCol缺药_NO As Integer
Private mIntCol缺药_类型 As Integer
Private mIntCol缺药_发药类型 As Integer
Private mIntCol缺药_床号 As Integer
Private mIntCol缺药_姓名 As Integer
Private mIntCol缺药_性别 As Integer
Private mIntCol缺药_品名 As Integer
Private mIntCol缺药_规格 As Integer
Private mIntCol缺药_生产商 As Integer
Private mIntCol缺药_原产地 As Integer
Private mIntCol缺药_批号 As Integer
Private mIntCol缺药_效期 As Integer
Private mIntCol缺药_数量 As Integer
Private mIntCol缺药_单价 As Integer
Private mIntCol缺药_金额 As Integer
Private mIntCol缺药_药品编码和名称 As Integer
Private mIntCol缺药_药品编码 As Integer
Private mIntCol缺药_药品名称 As Integer
Private mIntCol缺药_脚注 As Integer

'拒发列表
Private Const mconIntCol拒发_列数 As Integer = 24
Private mIntCol拒发_当前行 As Integer
Private mIntCol拒发_科室 As Integer
Private mIntCol拒发_状态 As Integer
Private mIntCol拒发_NO As Integer
Private mIntCol拒发_类型 As Integer
Private mIntCol拒发_发药类型 As Integer
Private mIntCol拒发_床号 As Integer
Private mIntCol拒发_姓名 As Integer
Private mIntCol拒发_性别 As Integer
Private mIntCol拒发_品名 As Integer
Private mIntCol拒发_规格 As Integer
Private mIntCol拒发_生产商 As Integer
Private mIntCol拒发_原产地 As Integer
Private mIntCol拒发_批号 As Integer
Private mIntCol拒发_效期 As Integer
Private mIntCol拒发_数量 As Integer
Private mIntCol拒发_单价 As Integer
Private mIntCol拒发_金额 As Integer
Private mIntCol拒发_药品编码和名称 As Integer
Private mIntCol拒发_药品编码 As Integer
Private mIntCol拒发_药品名称 As Integer
Private mIntCol拒发_执行状态 As Integer
Private mIntCol拒发_收发ID As Integer
Private mIntCol拒发_脚注 As Integer

'退药列表
Private Const mconIntCol退药_列数 As Integer = 50
Private mIntCol退药_当前行 As Integer
Private mIntCol退药_审查结果 As Integer
Private mIntCol退药_分组符 As Integer
Private mIntCol退药_科室 As Integer
Private mIntCol退药_状态 As Integer
Private mIntCol退药_类型 As Integer
Private mIntCol退药_发药类型 As Integer
Private mIntCol退药_NO As Integer
Private mIntCol退药_床号 As Integer
Private mIntCol退药_姓名 As Integer
Private mIntCol退药_性别 As Integer
Private mIntCol退药_住院号 As Integer
Private mIntCol退药_品名 As Integer
Private mIntCol退药_其它名 As Integer
Private mIntCol退药_英文名 As Integer
Private mIntCol退药_规格 As Integer
Private mIntCol退药_生产商 As Integer
Private mIntCol退药_原产地 As Integer
Private mIntCol退药_批号 As Integer
Private mIntCol退药_效期 As Integer
Private mIntCol退药_付 As Integer
Private mIntCol退药_数量 As Integer
Private mIntCol退药_已退数 As Integer
Private mIntCol退药_准退数 As Integer
Private mIntCol退药_退药数 As Integer
Private mIntCol退药_单价 As Integer
Private mIntCol退药_金额 As Integer
Private mIntCol退药_单量 As Integer
Private mIntCol退药_频次 As Integer
Private mIntCol退药_用法 As Integer
Private mIntCol退药_操作员 As Integer
Private mIntCol退药_发药时间 As Integer
Private mIntCol退药_单据 As Integer
Private mIntCol退药_医嘱id As Integer
Private mIntCol退药_领药人 As Integer
Private mIntCol退药_库房货位 As Integer
Private mIntCol退药_相关ID As Integer
Private mIntCol退药_药品ID As Integer
Private mIntCol退药_单量单位 As Integer
Private mIntCol退药_药品编码和名称 As Integer
Private mIntCol退药_药品编码 As Integer
Private mIntCol退药_药品名称 As Integer
Private mIntCol退药_收发ID As Integer
Private mIntCol退药_执行状态 As Integer
Private mIntCol退药_发药号 As Integer
Private mIntCol退药_领药部门id As Integer
Private mIntCol退药_发送时间 As Integer
Private mIntCol退药_脚注 As Integer
Private mIntCol退药_病人ID As Integer
Private mIntCol退药_主页ID As Integer

Public Sub SetSendBillStateByCustom(ByVal str收发ids As String)
    '自定义审核功能，根据返回的数据更新界面发药状态（取消发药）
    Dim intState As Integer
    Dim strState As String
    Dim lngColor As Long
    Dim i As Long
    Dim lng相关ID As Long
    Dim strNo As String
    
    If mcondition.intListType <> mListType.发药 Then Exit Sub
    If str收发ids = "" Then Exit Sub
    
    With vsfList(mListType.发药)
        intState = mState.不处理
        strState = "不处理"
        lngColor = mListColor.State_UnProcess
        
        .Redraw = flexRDNone
        For i = 1 To .rows - 1
            If .IsSubtotal(i) = False And Val(.TextMatrix(i, mIntCol发药_执行状态)) = mState.发药 _
                And InStr("," & str收发ids & ",", "," & Val(.TextMatrix(i, mIntCol发药_收发ID)) & ",") > 0 Then
                .TextMatrix(i, mIntCol发药_执行状态) = intState
                .TextMatrix(i, mIntCol发药_状态) = strState
                
                .Cell(flexcpBackColor, i, 1, i, .Cols - 1) = lngColor
                
                mrsSendList.Filter = "收发ID=" & Val(.TextMatrix(i, mIntCol发药_收发ID))
                
                mrsSendList!执行状态 = intState
                mrsSendList!状态 = strState
                
                mrsSendList.Update
                
                mblnSendChange = True
            End If
        Next
        
        .Redraw = flexRDDirect
    End With
End Sub
Private Sub GetDiagnosis(ByVal lngRow As Long)
    '根据所在行的病人ID得到病人的诊断记录
    Dim strTmp As String
    Dim i As Integer
    
    
    With vsf诊断
        
        If vsfList(mListType.发药).IsSubtotal(lngRow) = True Then 'Or vsfList(mListType.发药).TextMatrix(lngRow, mIntCol发药_领药部门) = "" Then
            .rows = 1
            .TextMatrix(0, 0) = ""
            .TextMatrix(0, 1) = ""
            lblDiag.Caption = "临床诊断"
            .Tag = ""
            Exit Sub
        End If
        
        If Val(.Tag) = Val(vsfList(mListType.发药).TextMatrix(lngRow, mIntCol发药_病人ID)) Then Exit Sub
        
        strTmp = RecipeSendWork_GetDiagnosis(IIf(chk所有诊断.Value = 1, 3, 2), Val(vsfList(mListType.发药).TextMatrix(lngRow, mIntCol发药_病人ID)), Val(vsfList(mListType.发药).TextMatrix(lngRow, mIntCol发药_主页ID)))
    
        .Tag = vsfList(mListType.发药).TextMatrix(lngRow, mIntCol发药_病人ID)
        
        lblDiag.Caption = "[" & vsfList(mListType.发药).TextMatrix(lngRow, mIntCol发药_姓名) & "]的临床诊断"
        
        .Redraw = flexRDNone
        .rows = 1
        .TextMatrix(0, 0) = ""
        .TextMatrix(0, 1) = ""
            
        If strTmp <> "" Then
            strTmp = strTmp & "|"
            For i = 0 To UBound(Split(strTmp, "|"))
                If Split(strTmp, "|")(i) <> "" Then
                    If i > 0 Then .rows = .rows + 1
                    .TextMatrix(i, 0) = Split(Split(strTmp, "|")(i), ",")(0)
                    .TextMatrix(i, 1) = Split(Split(strTmp, "|")(i), ",")(1)
                End If
            Next
        End If
        .Redraw = flexRDDirect
        
        If .TextMatrix(0, 0) = "" Then
            lblDiag.Caption = lblDiag.Caption & "(0)"
        Else
            lblDiag.Caption = lblDiag.Caption & "(" & .rows & ")"
        End If
    End With
End Sub

Public Function GetRecordInfo() As String
    '返回当前记录的信息
    '返回：单据|NO|病人ID
    
    If mcondition.intListType = mListType.发药 Or mcondition.intListType = mListType.退药 Then
        With vsfList(mcondition.intListType)
            If .Row = 0 Then Exit Function
            If .IsSubtotal(.Row) = True Then Exit Function
            If .TextMatrix(.Row, .ColIndex("收发ID")) = "" Then Exit Function
            
            GetRecordInfo = .TextMatrix(.Row, .ColIndex("单据")) & "|" & .TextMatrix(.Row, .ColIndex("NO")) & "|" & .TextMatrix(.Row, .ColIndex("药品ID"))
        End With
    End If
End Function
Public Sub ClearList(ByVal intType As Integer)
    '清空列表
    
    Select Case intType
        Case mListType.发药
            Set mrsSendList = Nothing
            Set mrsChargeOff = Nothing
            
            vsfList(mListType.发药).rows = 1
            vsfList(mListType.发药).rows = 2
            vsfList(mListType.汇总).rows = 1
            vsfList(mListType.汇总).rows = 2
            vsfList(mListType.拒发).rows = 1
            vsfList(mListType.拒发).rows = 2
            vsfList(mListType.缺药).rows = 1
            vsfList(mListType.缺药).rows = 2
            
            vsfChargeOff.rows = 1
            Me.pic签名图片.Visible = False
            Me.txt用药理由.Text = ""
            vsf诊断.rows = 0
            vsf诊断.rows = 4
        Case mListType.退药
            Set mrsReturnList = Nothing
            
            vsfList(mListType.退药).rows = 1
            vsfList(mListType.退药).rows = 2
    End Select
End Sub
Public Function GetPrintObject(ByVal blnOutPut As Boolean) As Object
    mblnOutPut = blnOutPut
    If vsfList(mcondition.intListType).rows = 1 Then
        Set GetPrintObject = Nothing
    Else
        Set GetPrintObject = vsfList(mcondition.intListType)
    End If
End Function

Private Sub Modify退药待发(ByVal blnShow As Boolean)
    '切换显示退药待发状态时：如果显示则更新记录为发药状态；如果不显示则更新为不处理状态
    
    If mrsSendList Is Nothing Then Exit Sub

    mrsSendList.Filter = "记录状态>1"
    If blnShow = True Then
        Do While Not mrsSendList.EOF
            If mrsSendList!执行状态 = mState.不处理 And mcondition.int退药待发单据默认为发药状态 = 1 Then
                If mrsSendList!高危药品 = 0 Or (mrsSendList!高危药品 > 0 And InStr(1, mcondition.str高危发放, mrsSendList!高危药品) = 0) And Not (InStr("," & mcondition.str高危分类 & ",", "," & mrsSendList!高危药品 & ",") > 0 And mrsSendList!高危药品 > 0) Then
                    mrsSendList!状态 = "发药"
                    mrsSendList!执行状态 = mState.发药
                    mrsSendList.Update
                End If
            End If
            mrsSendList.MoveNext
        Loop
    Else
        Do While Not mrsSendList.EOF
            If mrsSendList!执行状态 = mState.发药 Then
                mrsSendList!状态 = "不处理"
                mrsSendList!执行状态 = mState.不处理
                mrsSendList.Update
            End If
            mrsSendList.MoveNext
        Loop
    End If
End Sub

Public Sub SetAllReturn()
    '退药状态时，设置全退
    Dim n As Long
    
    If mcondition.intListType <> mListType.退药 Then Exit Sub
    
    With vsfList(mListType.退药)
        For n = 1 To .rows - 1
            If .IsSubtotal(n) = False Then
                If Val(.TextMatrix(n, mIntCol退药_执行状态)) = mState.退药_原始记录 And Val(.TextMatrix(n, mIntCol退药_准退数)) > 0 Then
                    .TextMatrix(n, mIntCol退药_退药数) = Val(.TextMatrix(n, mIntCol退药_准退数))
                    .TextMatrix(n, mIntCol退药_状态) = "退药"
                    .TextMatrix(n, mIntCol退药_执行状态) = mState.退药
                    
                    mrsReturnList.Filter = "收发ID=" & Val(.TextMatrix(n, mIntCol退药_收发ID))
                    
                    mrsReturnList!状态 = .TextMatrix(n, mIntCol退药_状态)
                    mrsReturnList!执行状态 = Val(.TextMatrix(n, mIntCol退药_执行状态))
                    mrsReturnList!退药数 = Val(.TextMatrix(n, mIntCol退药_退药数))
                    mrsReturnList.Update
                End If
            End If
        Next
    End With
End Sub


Public Sub SetAllNotReturn()
    '退药状态时，设置全不退
    Dim n As Long
    
    If mcondition.intListType <> mListType.退药 Then Exit Sub
    
    With vsfList(mListType.退药)
        For n = 1 To .rows - 1
            If .IsSubtotal(n) = False Then
                If Val(.TextMatrix(n, mIntCol退药_执行状态)) = mState.退药 Then
                    .TextMatrix(n, mIntCol退药_退药数) = ""
                    .TextMatrix(n, mIntCol退药_状态) = "不处理"
                    .TextMatrix(n, mIntCol退药_执行状态) = mState.退药_原始记录
                    
                    mrsReturnList.Filter = "收发ID=" & Val(.TextMatrix(n, mIntCol退药_收发ID))
                    
                    mrsReturnList!状态 = .TextMatrix(n, mIntCol退药_状态)
                    mrsReturnList!执行状态 = Val(.TextMatrix(n, mIntCol退药_执行状态))
                    mrsReturnList!退药数 = Val(.TextMatrix(n, mIntCol退药_退药数))
                    mrsReturnList.Update
                End If
            End If
        Next
    End With
End Sub

Public Sub SetFontSize(ByVal intFont As Integer)
    Dim objVSF As VSFlexGrid
    
    For Each objVSF In vsfList
        objVSF.Font.Size = intFont
        Me.Font.Size = objVSF.Font.Size
        objVSF.Cell(flexcpFontSize, 0, 0, objVSF.rows - 1, objVSF.Cols - 1) = objVSF.Font.Size
        
        objVSF.RowHeightMin = TextHeight("刘") + 100
        objVSF.RowHeightMax = TextHeight("刘") + 100
        objVSF.Refresh
    Next
End Sub
Public Sub AfterSendRefresh()
    '发药后更新发药数据集
    
    '删除已发药的记录
    If Not mrsSendList Is Nothing Then
        With mrsSendList
            .Filter = "执行状态=1"
            Do While Not .EOF
                .Delete
                .MoveNext
            Loop
            .UpdateBatch
            
            .Filter = ""
        End With
    End If
    
    If Not mrsChargeOff Is Nothing Then
        With mrsChargeOff
            .Filter = "执行标志=1"
            Do While Not .EOF
                .Delete
                .MoveNext
            Loop
            .UpdateBatch
            
            .Filter = ""
        End With
    End If
    
    '更新明细界面
    RefreshList mListType.发药, mrsSendList, mrsChargeOff
End Sub

Public Sub AfterReturnRefresh()
    '发药后更新发药数据集
    
    '删除已发药的记录
    With mrsReturnList
        .Filter = "执行状态=" & mState.退药
        Do While Not .EOF
            .Delete
            .MoveNext
        Loop
        .UpdateBatch
        
        .Filter = ""
    End With
    
    '更新明细界面
    RefreshList mListType.退药, mrsReturnList
End Sub
Public Sub AfterRejectRefresh()
    '拒发后更新发药数据集
    
    '修改拒发药的记录
    With mrsSendList
        .Filter = "执行状态=" & mState.拒发
        Do While Not .EOF
            !执行状态 = mState.拒发_不处理
            !状态 = "不处理"
            
            .MoveNext
        Loop
        .UpdateBatch
        
        .Filter = ""
    End With
    
    '更新明细界面
    RefreshList mListType.发药, mrsSendList
End Sub

Public Sub AfterRejectRestoreRefresh()
    '拒发恢复后更新发药数据集
    
    '修改拒发恢复的记录
    With mrsSendList
        .Filter = "执行状态=" & mState.拒发_恢复
        Do While Not .EOF
            !执行状态 = mState.发药
            !状态 = "发药"
            
            .MoveNext
        Loop
        .UpdateBatch
        
        .Filter = ""
    End With
    
    '更新明细界面
    RefreshList mListType.发药, mrsSendList
End Sub
Public Sub FindRecord(ByVal intType As Integer, Optional ByVal strFind As String = "")
    Dim lng收发ID As Long
    Dim int收发ID列 As Integer
    Dim strFilter As String
    
    '如果查找条件为空，并且上次查找条件也为空，则退出
    If strFind = "" And mstrFindCondition = "" Then Exit Sub
     
    If strFind <> "" And strFind <> mstrFindCondition Then
        '查找条件不为空，并且不等于上次查找条件，则重新过滤数据集
        mstrFindCondition = strFind
        If intType = mListType.退药 Then
            If mrsReturnList Is Nothing Then Exit Sub
            
            mrsReturnList.Filter = mstrFindCondition
            If mrsReturnList.RecordCount = 0 Then Exit Sub
            
            lng收发ID = mrsReturnList!收发ID
            int收发ID列 = mIntCol退药_收发ID
        Else
            If mrsSendList Is Nothing Then Exit Sub
            
            mrsSendList.Filter = mstrFindCondition
            If mrsSendList.RecordCount = 0 Then Exit Sub
            
            lng收发ID = mrsSendList!收发ID
            int收发ID列 = mIntCol发药_收发ID
        End If
    Else
        '查找条件为空，或者等于上次查找条件，则在已过滤范围中查找下条记录
        If intType = mListType.退药 Then
            If mrsReturnList Is Nothing Then Exit Sub
            
            If mrsReturnList.RecordCount = 0 Then Exit Sub
                
            int收发ID列 = mIntCol退药_收发ID
            
            mrsReturnList.MoveNext
            
            If Not mrsReturnList.EOF Then
                lng收发ID = mrsReturnList!收发ID
            Else
                mrsReturnList.MoveFirst
                lng收发ID = mrsReturnList!收发ID
            End If
                    
        Else
            If mrsSendList Is Nothing Then Exit Sub
            
            If mrsSendList.RecordCount = 0 Then Exit Sub
            
            int收发ID列 = mIntCol发药_收发ID
            
            mrsSendList.MoveNext
            
            If Not mrsSendList.EOF Then
                lng收发ID = mrsSendList!收发ID
            Else
                mrsSendList.MoveFirst
                lng收发ID = mrsSendList!收发ID
            End If
                
        End If
    End If
    
    '根据查找到的收发ID，在表格中定位
    With vsfList(intType)
        .Row = .FindRow(lng收发ID, 1, int收发ID列)
    End With
End Sub

Public Function GetReturnDate() As String
    '取待退药单据日期
    '返回：日期
    
    With vsfList(mListType.退药)
        If .Row = 0 Then Exit Function
        If .IsSubtotal(.Row) = True Then Exit Function
        If .TextMatrix(.Row, mIntCol退药_发药时间) = "" Then Exit Function

        GetReturnDate = .TextMatrix(.Row, mIntCol退药_发药时间)
    End With
End Function

Public Function GetSendedInfo() As String
    '返回已发药单据信息：领药部门|领药部门ID|汇总发药号
    
    With vsfList(mListType.退药)
        If .Row = 0 Then Exit Function
        If .IsSubtotal(.Row) = True Then Exit Function
        If .TextMatrix(.Row, mIntCol退药_发药号) = "" Then Exit Function

        GetSendedInfo = .TextMatrix(.Row, mIntCol退药_科室) & "|" & .TextMatrix(.Row, mIntCol退药_领药部门id) & "|" & .TextMatrix(.Row, mIntCol退药_发药号)
    End With
End Function
Public Function GetSendRecord() As ADODB.Recordset
    '用于向主界面返回发药记录集
    
    If mrsSendList Is Nothing Then
        Set GetSendRecord = Nothing
        Exit Function
    Else
        mrsSendList.Filter = ""
        Set GetSendRecord = mrsSendList
    End If
End Function

Public Function GetReturnRecord() As ADODB.Recordset
    '用于向主界面返回退药记录集
    
    If mrsReturnList Is Nothing Then
        Set GetReturnRecord = Nothing
    Else
        mrsReturnList.Filter = ""
        Set GetReturnRecord = mrsReturnList
    End If
End Function

Public Function GetChargeOffRecord() As ADODB.Recordset
    '用于向主界面返回销帐记录集
    Dim i As Integer
    Dim lng领药部门ID As Long
    Dim lng药品id As Long
    
    With vsfList(mListType.汇总)
        If mcondition.bln按科室汇总 = True Then
            For i = 1 To .rows - 1
                lng领药部门ID = Val(.TextMatrix(i, mIntCol科室汇总_领药部门id))
                lng药品id = Val(.TextMatrix(i, mIntCol科室汇总_药品ID))
                 
                If Not mrsChargeOff Is Nothing Then
                    mrsChargeOff.Filter = "审核标志>0 And 领药部门ID=" & lng领药部门ID & " And 药品ID=" & lng药品id
                    Do While Not mrsChargeOff.EOF
                        mrsChargeOff!执行标志 = 1
                        mrsChargeOff.Update
                        mrsChargeOff.MoveNext
                    Loop
                End If
            Next
        End If
    End With
    
    Set GetChargeOffRecord = mrsChargeOff
End Function

Public Function GetStayRecord() As ADODB.Recordset
    '用于向主界面返回留存记录集
    Dim i As Integer
    Dim rsStay As ADODB.Recordset                  '留存数据集
    
    Set rsStay = New ADODB.Recordset
    With rsStay
        If .State = 1 Then .Close
        .Fields.Append "领药部门ID", adDouble, 18, adFldIsNullable
        .Fields.Append "药品ID", adDouble, 18, adFldIsNullable
        .Fields.Append "批次", adDouble, 18, adFldIsNullable
        .Fields.Append "留存数量", adDouble, 18, adFldIsNullable
        .Fields.Append "单价", adLongVarChar, 50, adFldIsNullable
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    If mcondition.bln按科室汇总 = True Then
        With vsfList(mListType.汇总)
            For i = 1 To .rows - 1
                If Val(.TextMatrix(i, mIntCol科室汇总_领药部门id)) > 0 And Val(.TextMatrix(i, mIntCol科室汇总_留存数量)) > 0 Then
                    rsStay.AddNew
                    rsStay!领药部门ID = Val(.TextMatrix(i, mIntCol科室汇总_领药部门id))
                    rsStay!药品ID = Val(.TextMatrix(i, mIntCol科室汇总_药品ID))
                    rsStay!批次 = Val(.TextMatrix(i, mIntCol科室汇总_批次))
                    rsStay!留存数量 = Val(.TextMatrix(i, mIntCol科室汇总_留存数量)) * Val(.TextMatrix(i, mIntCol科室汇总_包装))
                    rsStay!单价 = Val(.TextMatrix(i, mIntCol科室汇总_单价)) / Val(.TextMatrix(i, mIntCol科室汇总_包装))
                    rsStay.Update
                End If
            Next
        End With
    End If
    
    Set GetStayRecord = rsStay
End Function

Public Function Get当前发药单格式() As Integer
    Get当前发药单格式 = IIf(cbo发药单格式.ListIndex = -1, 1, cbo发药单格式.ListIndex + 1)
End Function

Public Function Get当前配药人() As String
    '用于向主窗口返回配药人
    
    If InStr(cbo配药人.Text, "-") > 0 Then
        Get当前配药人 = Mid(cbo配药人.Text, InStr(cbo配药人.Text, "-") + 1)
    Else
        Get当前配药人 = cbo配药人.Text
    End If
End Function

Public Function Get当前核查人() As String
    '用于向主窗口返回配药人
    
    If InStr(cbo核查人.Text, "-") > 0 Then
        Get当前核查人 = Mid(cbo核查人.Text, InStr(cbo核查人.Text, "-") + 1)
    Else
        Get当前核查人 = cbo核查人.Text
    End If
End Function


Private Sub InitColSelList(ByVal intListType As Integer)
    Dim i As Integer
    
    With vsfColSel
        .Tag = intListType
        
        .rows = .FixedRows
        For i = 1 To vsfList(intListType).Cols - 1
            '不在不允许显示列表的列才能加入列选择列表
            If IsInString(mstrUnallowShow(intListType), vsfList(intListType).ColKey(i), ";") = False Then
                If (mcondition.bln显示原产地 And vsfList(intListType).ColKey(i) = "原产地") Or vsfList(intListType).ColKey(i) <> "原产地" Then
                    .rows = .rows + 1
                    .TextMatrix(.rows - 1, 1) = vsfList(intListType).ColKey(i)
                    .RowData(.rows - 1) = i
                End If
                
                '列宽为空或者隐藏的列设置为不勾选
                If Not (vsfList(intListType).ColWidth(i) = 0 Or vsfList(intListType).ColHidden(i)) Then
                    .TextMatrix(.rows - 1, 0) = 0
                End If
                
                '指定的列设置为不能设置隐藏
                If IsInString(mstrUnallowSetColHide(intListType), vsfList(intListType).ColKey(i), ";") = True Then
                    .Cell(flexcpForeColor, .rows - 1, 1) = .BackColorFixed
                End If
            End If
        Next
    End With
End Sub

Private Sub InitList(ByVal intType As Integer)
    '根据参数初始化列表
    
    Select Case intType
        Case mListType.发药
            Call InitList_Send
        Case mListType.汇总
            Call InitList_Sum
            Call InitList_ChargeOff
        Case mListType.缺药
            Call InitList_Shortage
        Case mListType.拒发
            Call InitList_Reject
        Case mListType.退药
            Call InitList_Return
        Case Else
            Call InitList_Send
            Call InitList_Sum
            Call InitList_ChargeOff
            Call InitList_Shortage
            Call InitList_Reject
            Call InitList_Return
    End Select
End Sub

Private Sub SaveListColState(Optional intType As Integer = -1)
    Dim str列设置 As String
    Dim i As Integer
    Dim strType As String
    Dim n As Integer
    Dim intStart As Integer
    Dim intEnd As Integer
    
    If Val(zlDataBase.GetPara("使用个性化风格")) = 0 Then Exit Sub
    
    If intType = -1 Then
        intStart = 0
        intEnd = vsfList.count - 1
    Else
        intStart = intType
        intEnd = intType
    End If
    
    For n = intStart To intEnd
        Select Case n
            Case mListType.发药
                strType = "发药"
            Case mListType.汇总
                If mcondition.bln按科室汇总 = True Then
                    strType = "科室汇总"
                Else
                    strType = "汇总"
                End If
            Case mListType.缺药
                strType = "缺药"
            Case mListType.拒发
                strType = "拒发"
            Case mListType.退药
                strType = "退药"
        End Select
        
        str列设置 = ""
        With vsfList(n)
            For i = 0 To .Cols - 1
                str列设置 = IIf(str列设置 = "", "", str列设置 & "|") & .ColKey(i) & "," & IIf(.ColHidden(i) = True, 0, .ColWidth(i))
            Next
        End With
        
        Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(vsfList(n)), strType, str列设置)
    Next
End Sub

Private Function LoadListColState(ByVal intType As Integer) As String
    Dim str列设置 As String
    Dim i As Integer
    Dim strType As String
    
    If Val(zlDataBase.GetPara("使用个性化风格")) = 0 Then Exit Function
    
    Select Case intType
        Case mListType.发药
            strType = "发药"
        Case mListType.汇总
            If mcondition.bln按科室汇总 = True Then
                strType = "科室汇总"
            Else
                strType = "汇总"
            End If
        Case mListType.缺药
            strType = "缺药"
        Case mListType.拒发
            strType = "拒发"
        Case mListType.退药
            strType = "退药"
    End Select
    
    LoadListColState = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(vsfList(intType)), strType, "")
End Function
Private Sub InitList_Send()
    Dim i As Integer
    Dim n As Integer
    Dim str列设置 As String
    Dim arr列设置
    
    '''初始化列顺序
    '默认列顺序
    mIntCol发药_当前行 = 0
    mIntCol发药_分组符 = 1
    mIntCol发药_审查结果 = 2
    mIntCol发药_领药部门 = 3
    mIntCol发药_科室 = 4
    mIntCol发药_开单医生 = 5
    mIntCol发药_状态 = 6
    mIntCol发药_类型 = 7
    mIntCol发药_发药类型 = 8
    mIntCol发药_NO = 9
    mIntCol发药_记帐员 = 10
    mIntCol发药_床号 = 11
    mIntCol发药_病人类型 = 12
    mIntCol发药_姓名 = 13
    mIntCol发药_性别 = 14
    mIntCol发药_年龄 = 15
    mIntCol发药_住院号 = 16
    mIntCol发药_品名 = 17
    mIntCol发药_皮试结果 = 18
    mIntCol发药_其它名 = 19
    mIntCol发药_英文名 = 20
    mIntCol发药_配方名称 = 21
    mIntCol发药_规格 = 22
    mIntCol发药_生产商 = 23
    mIntCol发药_原产地 = 24
    mIntCol发药_批号 = 25
    mIntCol发药_效期 = 26
    mIntCol发药_付 = 27
    mIntCol发药_数量 = 28
    mIntCol发药_单价 = 29
    mIntCol发药_金额 = 30
    mIntCol发药_单量 = 31
    mIntCol发药_频次 = 32
    mIntCol发药_用法 = 33
    mIntCol发药_用药次数 = 34
    mIntCol发药_用药目的 = 35
    mIntCol发药_禁忌药品说明 = 36
    mIntCol发药_记帐时间 = 37
    mIntCol发药_说明 = 38
    mIntCol发药_单据 = 39
    mIntCol发药_医嘱id = 40
    mIntCol发药_退药人 = 41
    mIntCol发药_库房货位 = 42
    mIntCol发药_相关ID = 43
    mIntCol发药_药品ID = 44
    mIntCol发药_单量单位 = 45
    mIntCol发药_领药部门id = 46
    mIntCol发药_药品编码和名称 = 47
    mIntCol发药_药品编码 = 48
    mIntCol发药_药品名称 = 49
    mIntCol发药_收发ID = 50
    mIntCol发药_执行状态 = 51
    mIntCol发药_领药号 = 52
    mIntCol发药_已收费 = 53
    mIntCol发药_病人ID = 54
    mIntCol发药_主页ID = 55
    mIntCol发药_用药理由 = 56
    mIntCol发药_高危药品 = 57
    mIntCol发药_结果 = 58
    mIntCol发药_开单部门id = 59
    mIntCol发药_脚注 = 60
    
    '恢复用户自定义顺序
    str列设置 = LoadListColState(mListType.发药)
    If str列设置 <> "" Then
        arr列设置 = Split(str列设置, "|")
        If UBound(arr列设置) + 1 <> mconIntCol发药_列数 Then
            str列设置 = ""
        Else
            For n = 0 To UBound(arr列设置)
                SetColumnValue mListType.发药, Split(arr列设置(n), ",")(0), n
            Next
        End If
    End If
    
    '初始化未发药清单
    With vsfList(mListType.发药)
        .Redraw = flexRDNone
        
        .rows = 2
        
'        .RowHeightMax = 255
        .Cols = mconIntCol发药_列数
        
'        .Cell(flexcpPicture, 1, mIntCol发药_当前行, 1, mIntCol发药_当前行) = Me.imgList.ListImages(2).Picture
'        .Cell(flexcpPictureAlignment, 1, mIntCol发药_当前行, .Rows - 1, mIntCol发药_当前行) = flexPicAlignRightCenter
        
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_当前行, "", 250, flexAlignCenterCenter, "当前行"
        
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_审查结果, "警", IIf(Not (gobjPass Is Nothing) And IsInString(gstrprivs, "合理用药监测", ";"), 300, 0), flexAlignCenterCenter, "审查结果"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_分组符, "组", 300, flexAlignRightCenter, "分组"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_科室, "病人科室", 1000, flexAlignLeftCenter, "病人科室"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_开单医生, "开单医生", IIf(mcondition.bln医生查询 = True, 1100, 0), flexAlignLeftCenter, "开单医生"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_状态, "状态", 1000, flexAlignLeftCenter, "状态"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_类型, "类型", 1000, flexAlignLeftCenter, "类型"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_发药类型, "发药类型", 1000, flexAlignLeftCenter, "发药类型"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_NO, "NO", 900, flexAlignLeftCenter, "NO"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_记帐员, "记帐员", 800, flexAlignLeftCenter, "记帐员"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_床号, "床号", 1000, flexAlignLeftCenter, "床号"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_病人类型, "病人类型", 1000, flexAlignLeftCenter, "病人类型"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_姓名, "姓名", 700, flexAlignLeftCenter, "姓名"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_性别, "性别", 700, flexAlignLeftCenter, "性别"
        
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_年龄, "年龄", 700, flexAlignLeftCenter, "年龄"
        
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_住院号, "住院号", 1200, flexAlignLeftCenter, "住院号"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_品名, "药品名称", 2500, flexAlignLeftCenter, "药品名称"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_皮试结果, "", 800, flexAlignLeftCenter, "皮试结果"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_其它名, "其它名", 2000, flexAlignLeftCenter, "其它名"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_英文名, "英文名", 2000, flexAlignLeftCenter, "英文名"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_配方名称, "配方名称", 2000, flexAlignLeftCenter, "配方名称"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_规格, "规格", 1500, flexAlignLeftCenter, "规格"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_生产商, "生产商", 1500, flexAlignLeftCenter, "生产商"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_原产地, "原产地", 1500, flexAlignLeftCenter, "原产地"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_批号, "批号", 1500, flexAlignLeftCenter, "批号"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_效期, "效期", 1500, flexAlignLeftCenter, "效期"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_付, "付", 300, flexAlignRightCenter, "付"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_数量, "数量", 1200, flexAlignRightCenter, "数量"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_单价, "单价", 1200, flexAlignRightCenter, "单价"
        
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_金额, "金额", 1200, flexAlignRightCenter, "金额"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_单量, "单量", 1200, flexAlignRightCenter, "单量"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_频次, "频次", 500, flexAlignLeftCenter, "频次"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_用法, "用法", 800, flexAlignLeftCenter, "用法"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_用药次数, "用药次数", 900, flexAlignRightCenter, "用药次数"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_用药目的, "用药目的", 0, flexAlignLeftCenter, "用药目的"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_禁忌药品说明, "禁忌药品说明", 1500, flexAlignLeftCenter, "禁忌药品说明"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_记帐时间, "记帐时间", 1800, flexAlignLeftCenter, "记帐时间"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_说明, "说明", 1200, flexAlignLeftCenter, "说明"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_单据, "单据", 0, flexAlignCenterCenter, "单据"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_医嘱id, "医嘱id", 0, flexAlignCenterCenter, "医嘱id"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_退药人, "退药人", 1000, flexAlignLeftCenter, "退药人"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_库房货位, "库房货位", IIf(mcondition.bln药品储备 = True, 1200, 0), flexAlignLeftCenter, "库房货位"
        
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_相关ID, "相关ID", 0, flexAlignCenterCenter, "相关ID"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_药品ID, "药品ID", 0, flexAlignCenterCenter, "药品ID"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_单量单位, "单量单位", 0, flexAlignLeftCenter, "单量单位"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_领药部门, "领药部门", 1000, flexAlignLeftCenter, "领药部门"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_领药部门id, "领药部门id", 0, flexAlignCenterCenter, "领药部门id"
        
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_药品编码和名称, "药品编码和名称", 0, flexAlignCenterCenter, "药品编码和名称"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_药品编码, "药品编码", 0, flexAlignCenterCenter, "药品编码"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_药品名称, "药名", 0, flexAlignCenterCenter, "药名"
        
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_收发ID, "收发ID", 0, flexAlignCenterCenter, "收发ID"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_执行状态, "状态标志", 0, flexAlignCenterCenter, "状态标志"
        
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_领药号, "领药号", 1000, flexAlignLeftCenter, "领药号"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_已收费, "已收费", 0, flexAlignLeftCenter, "已收费"
        
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_病人ID, "病人ID", 0, flexAlignCenterCenter, "病人ID"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_主页ID, "主页ID", 0, flexAlignCenterCenter, "主页ID"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_用药理由, "用药理由", 0, flexAlignCenterCenter, "用药理由"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_高危药品, "高危药品", 0, flexAlignCenterCenter, "高危药品"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_结果, "结果", 0, flexAlignCenterCenter, "结果"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_开单部门id, "开单部门id", 0, flexAlignCenterCenter, "开单部门id"
        VsfGridColFormat vsfList(mListType.发药), mIntCol发药_脚注, "脚注", 1000, flexAlignLeftCenter, "脚注"
        
        mstrUnallowSetColHide(mListType.发药) = "状态;药品名称;数量"
        mstrUnallowShow(mListType.发药) = "当前行;审查结果;分组;单据;医嘱id;用药目的;相关ID;药品ID;皮试结果;单量单位;领药部门id;药品编码和名称;药品编码;药名;收发ID;状态标志;已收费;病人ID;主页ID;用药理由;高危药品;结果;开单部门id"
        If mcondition.bln药品储备 = False Then mstrUnallowShow(mListType.发药) = mstrUnallowShow(mListType.发药) & ";库房货位"
        If mcondition.bln医生查询 = False Then mstrUnallowShow(mListType.发药) = mstrUnallowShow(mListType.发药) & ";开单医生"
        
        '恢复自定义列宽（不包括不允许显示的列）
        If str列设置 <> "" Then
            arr列设置 = Split(str列设置, "|")
            For n = 0 To UBound(arr列设置)
                If IsInString(mstrUnallowShow(mListType.发药), Split(arr列设置(n), ",")(0), ";") = False Then
                    For i = 0 To .Cols - 1
                        If Split(arr列设置(n), ",")(0) = .ColKey(i) Then
                            .ColWidth(i) = Val(Split(arr列设置(n), ",")(1))
                        End If
                    Next
                End If
            Next
        End If
        
        If mcondition.bln显示原产地 = False Then VsfGridColFormat vsfList(mListType.发药), mIntCol发药_原产地, "原产地", 0, flexAlignLeftCenter, "原产地"
        
        .Redraw = flexRDDirect
    End With
    
End Sub

Private Sub InitList_Reject()
    Dim i As Integer
    Dim n As Integer
    Dim str列设置 As String
    Dim arr列设置
    
    '''初始化列顺序
    '默认列顺序
    mIntCol拒发_当前行 = 0
    mIntCol拒发_科室 = 1
    mIntCol拒发_状态 = 2
    mIntCol拒发_NO = 3
    mIntCol拒发_类型 = 4
    mIntCol拒发_发药类型 = 5
    mIntCol拒发_床号 = 6
    mIntCol拒发_姓名 = 7
    mIntCol拒发_性别 = 8
    mIntCol拒发_品名 = 9
    mIntCol拒发_规格 = 10
    mIntCol拒发_生产商 = 11
    mIntCol拒发_原产地 = 12
    mIntCol拒发_批号 = 13
    mIntCol拒发_效期 = 14
    mIntCol拒发_数量 = 15
    mIntCol拒发_单价 = 16
    mIntCol拒发_金额 = 17
    mIntCol拒发_药品编码和名称 = 18
    mIntCol拒发_药品编码 = 19
    mIntCol拒发_药品名称 = 20
    mIntCol拒发_执行状态 = 21
    mIntCol拒发_收发ID = 22
    mIntCol拒发_脚注 = 23
    
    '恢复用户自定义顺序
    str列设置 = LoadListColState(mListType.拒发)
    If str列设置 <> "" Then
        arr列设置 = Split(str列设置, "|")
        If UBound(arr列设置) + 1 <> mconIntCol拒发_列数 Then
            str列设置 = ""
        Else
            For n = 0 To UBound(arr列设置)
               SetColumnValue mListType.拒发, Split(arr列设置(n), ",")(0), n
            Next
        End If
    End If
    
    '初始化拒发药清单
    With vsfList(mListType.拒发)
        .Redraw = flexRDNone
        .rows = 2
        .Cols = mconIntCol拒发_列数
        
        .Cell(flexcpPicture, 1, mIntCol拒发_当前行, 1, mIntCol拒发_当前行) = Me.ImgList.ListImages(2).Picture
        .Cell(flexcpPictureAlignment, 1, mIntCol拒发_当前行, .rows - 1, mIntCol拒发_当前行) = flexPicAlignRightCenter
        
        VsfGridColFormat vsfList(mListType.拒发), mIntCol拒发_当前行, "", 250, flexAlignCenterCenter, "当前行"
        VsfGridColFormat vsfList(mListType.拒发), mIntCol拒发_科室, "病人科室", 1200, flexAlignLeftCenter, "病人科室"
        VsfGridColFormat vsfList(mListType.拒发), mIntCol拒发_状态, "状态", 1000, flexAlignLeftCenter, "状态"
        VsfGridColFormat vsfList(mListType.拒发), mIntCol拒发_NO, "NO", 800, flexAlignLeftCenter, "NO"
        VsfGridColFormat vsfList(mListType.拒发), mIntCol拒发_类型, "类型", 1000, flexAlignLeftCenter, "类型"
        VsfGridColFormat vsfList(mListType.拒发), mIntCol拒发_发药类型, "发药类型", 1000, flexAlignLeftCenter, "发药类型"
        VsfGridColFormat vsfList(mListType.拒发), mIntCol拒发_床号, "床号", 800, flexAlignLeftCenter, "床号"
        VsfGridColFormat vsfList(mListType.拒发), mIntCol拒发_姓名, "姓名", 1000, flexAlignLeftCenter, "姓名"
        VsfGridColFormat vsfList(mListType.拒发), mIntCol拒发_性别, "性别", 1000, flexAlignLeftCenter, "性别"
        VsfGridColFormat vsfList(mListType.拒发), mIntCol拒发_品名, "药品名称", 2500, flexAlignLeftCenter, "药品名称"
        VsfGridColFormat vsfList(mListType.拒发), mIntCol拒发_规格, "规格", 1500, flexAlignLeftCenter, "规格"
        VsfGridColFormat vsfList(mListType.拒发), mIntCol拒发_生产商, "生产商", 1500, flexAlignLeftCenter, "生产商"
        VsfGridColFormat vsfList(mListType.拒发), mIntCol拒发_原产地, "原产地", 1500, flexAlignLeftCenter, "原产地"
        VsfGridColFormat vsfList(mListType.拒发), mIntCol拒发_批号, "批号", 1500, flexAlignLeftCenter, "批号"
        VsfGridColFormat vsfList(mListType.拒发), mIntCol拒发_效期, "效期", 1500, flexAlignLeftCenter, "效期"
        VsfGridColFormat vsfList(mListType.拒发), mIntCol拒发_数量, "数量", 1200, flexAlignRightCenter, "数量"
        VsfGridColFormat vsfList(mListType.拒发), mIntCol拒发_单价, "单价", 1200, flexAlignRightCenter, "单价"
        VsfGridColFormat vsfList(mListType.拒发), mIntCol拒发_金额, "金额", 1200, flexAlignRightCenter, "金额"
        
        VsfGridColFormat vsfList(mListType.拒发), mIntCol拒发_药品编码和名称, "药品编码和名称", 0, flexAlignCenterCenter, "药品编码和名称"
        VsfGridColFormat vsfList(mListType.拒发), mIntCol拒发_药品编码, "药品编码", 0, flexAlignCenterCenter, "药品编码"
        VsfGridColFormat vsfList(mListType.拒发), mIntCol拒发_药品名称, "药品名称", 0, flexAlignCenterCenter, "药名"
        
        VsfGridColFormat vsfList(mListType.拒发), mIntCol拒发_执行状态, "执行状态", 0, flexAlignCenterCenter, "执行状态"
        VsfGridColFormat vsfList(mListType.拒发), mIntCol拒发_收发ID, "收发ID", 0, flexAlignCenterCenter, "收发ID"
        
        VsfGridColFormat vsfList(mListType.拒发), mIntCol拒发_脚注, "脚注", 1200, flexAlignLeftCenter, "脚注"
        
        mstrUnallowSetColHide(mListType.拒发) = "状态;药品名称;数量"
        mstrUnallowShow(mListType.拒发) = "当前行;药品编码和名称;药品编码;药名;执行状态;收发ID"
        
        '恢复自定义列宽（不包括不允许显示的列）
        If str列设置 <> "" Then
            arr列设置 = Split(str列设置, "|")
            For n = 0 To UBound(arr列设置)
                If IsInString(mstrUnallowShow(mListType.拒发), Split(arr列设置(n), ",")(0), ";") = False Then
                    For i = 0 To .Cols - 1
                        If Split(arr列设置(n), ",")(0) = .ColKey(i) Then
                            .ColWidth(i) = Val(Split(arr列设置(n), ",")(1))
                        End If
                    Next
                End If
            Next
        End If
        
        If mcondition.bln显示原产地 = False Then VsfGridColFormat vsfList(mListType.拒发), mIntCol拒发_原产地, "原产地", 0, flexAlignLeftCenter, "原产地"
        
        .Redraw = flexRDDirect
    End With
    
End Sub

Private Sub InitList_Shortage()
    Dim i As Integer
    Dim n As Integer
    Dim str列设置 As String
    Dim arr列设置
    
    '''初始化列顺序
    '默认列顺序
    mIntCol缺药_当前行 = 0
    mIntCol缺药_科室 = 1
    mIntCol缺药_NO = 2
    mIntCol缺药_类型 = 3
    mIntCol缺药_发药类型 = 4
    mIntCol缺药_床号 = 5
    mIntCol缺药_姓名 = 6
    mIntCol缺药_性别 = 7
    mIntCol缺药_品名 = 8
    mIntCol缺药_规格 = 9
    mIntCol缺药_生产商 = 10
    mIntCol缺药_原产地 = 11
    mIntCol缺药_批号 = 12
    mIntCol缺药_效期 = 13
    mIntCol缺药_数量 = 14
    mIntCol缺药_单价 = 15
    mIntCol缺药_金额 = 16
    mIntCol缺药_药品编码和名称 = 17
    mIntCol缺药_药品编码 = 18
    mIntCol缺药_药品名称 = 19
    mIntCol缺药_脚注 = 20
    
    '恢复用户自定义顺序
    str列设置 = LoadListColState(mListType.缺药)
    If str列设置 <> "" Then
        arr列设置 = Split(str列设置, "|")
        If UBound(arr列设置) + 1 <> mconIntCol缺药_列数 Then
            str列设置 = ""
        Else
            For n = 0 To UBound(arr列设置)
                SetColumnValue mListType.缺药, Split(arr列设置(n), ",")(0), n
            Next
        End If
    End If
    
    '初始化缺药清单
    With vsfList(mListType.缺药)
        .Redraw = flexRDNone
        .rows = 2
        .Cols = mconIntCol缺药_列数
        
        .Cell(flexcpPicture, 1, mIntCol缺药_当前行, 1, mIntCol缺药_当前行) = Me.ImgList.ListImages(2).Picture
        .Cell(flexcpPictureAlignment, 1, mIntCol缺药_当前行, .rows - 1, mIntCol缺药_当前行) = flexPicAlignRightCenter
        
        VsfGridColFormat vsfList(mListType.缺药), mIntCol缺药_当前行, "", 250, flexAlignCenterCenter, "当前行"
        VsfGridColFormat vsfList(mListType.缺药), mIntCol缺药_科室, "病人科室", 1200, flexAlignLeftCenter, "病人科室"
        VsfGridColFormat vsfList(mListType.缺药), mIntCol缺药_NO, "NO", 800, flexAlignLeftCenter, "NO"
        VsfGridColFormat vsfList(mListType.缺药), mIntCol缺药_类型, "类型", 1000, flexAlignCenterCenter, "类型"
        VsfGridColFormat vsfList(mListType.缺药), mIntCol缺药_发药类型, "发药类型", 1000, flexAlignCenterCenter, "发药类型"
        VsfGridColFormat vsfList(mListType.缺药), mIntCol缺药_床号, "床号", 800, flexAlignLeftCenter, "床号"
        VsfGridColFormat vsfList(mListType.缺药), mIntCol缺药_姓名, "姓名", 1000, flexAlignLeftCenter, "姓名"
        VsfGridColFormat vsfList(mListType.缺药), mIntCol缺药_性别, "性别", 1000, flexAlignLeftCenter, "性别"
        VsfGridColFormat vsfList(mListType.缺药), mIntCol缺药_品名, "药品名称", 2500, flexAlignLeftCenter, "药品名称"
        VsfGridColFormat vsfList(mListType.缺药), mIntCol缺药_规格, "规格", 1500, flexAlignLeftCenter, "规格"
        VsfGridColFormat vsfList(mListType.缺药), mIntCol缺药_生产商, "生产商", 1500, flexAlignLeftCenter, "生产商"
        VsfGridColFormat vsfList(mListType.缺药), mIntCol缺药_原产地, "原产地", 1500, flexAlignLeftCenter, "原产地"
        VsfGridColFormat vsfList(mListType.缺药), mIntCol缺药_批号, "批号", 1500, flexAlignLeftCenter, "批号"
        VsfGridColFormat vsfList(mListType.缺药), mIntCol缺药_效期, "效期", 1500, flexAlignLeftCenter, "效期"
        VsfGridColFormat vsfList(mListType.缺药), mIntCol缺药_数量, "数量", 1200, flexAlignRightCenter, "数量"
        VsfGridColFormat vsfList(mListType.缺药), mIntCol缺药_单价, "单价", 1200, flexAlignRightCenter, "单价"
        VsfGridColFormat vsfList(mListType.缺药), mIntCol缺药_金额, "金额", 1200, flexAlignRightCenter, "金额"
        VsfGridColFormat vsfList(mListType.缺药), mIntCol缺药_药品编码和名称, "药品编码和名称", 0, flexAlignRightCenter, "药品编码和名称"
        VsfGridColFormat vsfList(mListType.缺药), mIntCol缺药_药品编码, "药品编码", 0, flexAlignRightCenter, "药品编码"
        VsfGridColFormat vsfList(mListType.缺药), mIntCol缺药_药品名称, "药名", 0, flexAlignRightCenter, "药名"
        VsfGridColFormat vsfList(mListType.缺药), mIntCol缺药_脚注, "脚注", 1200, flexAlignLeftCenter, "脚注"
        
        mstrUnallowSetColHide(mListType.缺药) = "药品名称;数量"
        mstrUnallowShow(mListType.缺药) = "当前行;药品编码和名称;药品编码;药名"
        
        '恢复自定义列宽（不包括不允许显示的列）
        If str列设置 <> "" Then
            arr列设置 = Split(str列设置, "|")
            For n = 0 To UBound(arr列设置)
                If IsInString(mstrUnallowShow(mListType.缺药), Split(arr列设置(n), ",")(0), ";") = False Then
                    For i = 0 To .Cols - 1
                        If Split(arr列设置(n), ",")(0) = .ColKey(i) Then
                            .ColWidth(i) = Val(Split(arr列设置(n), ",")(1))
                        End If
                    Next
                End If
            Next
        End If
        
        If mcondition.bln显示原产地 = False Then VsfGridColFormat vsfList(mListType.缺药), mIntCol缺药_原产地, "原产地", 0, flexAlignLeftCenter, "原产地"
        
        .Redraw = flexRDDirect
    End With
    
End Sub

Private Sub InitList_ChargeOff()
    '''初始化列顺序
    '默认列顺序
    mIntCol销账_当前行 = 0
    mIntCol销账_申请科室 = 1
    mIntCol销账_单据 = 2
    mIntCol销账_NO = 3
    mIntCol销账_药品ID = 4
    mIntCol销账_申请时间 = 5
    mIntCol销账_收发序号 = 6
    mIntCol销账_生产商 = 7
    mIntCol销账_批号 = 8
    mIntCol销账_效期 = 9
    mIntCol销账_准退数量 = 10
    mIntCol销账_销帐数量 = 11
    mIntCol销账_包装 = 12
    mIntCol销账_单位 = 13
    
    '恢复用户自定义顺序
    
    '初始化缺药清单
    With vsfChargeOff
        .Redraw = flexRDNone
        .rows = 2
        .Cols = mconIntCol销账_列数
        
        .Cell(flexcpPicture, 1, mIntCol销账_当前行, 1, mIntCol销账_当前行) = Me.ImgList.ListImages(2).Picture
        .Cell(flexcpPictureAlignment, 1, mIntCol销账_当前行, .rows - 1, mIntCol销账_当前行) = flexPicAlignRightCenter
        
        VsfGridColFormat vsfChargeOff, mIntCol销账_当前行, "", 250, flexAlignCenterCenter
        
        VsfGridColFormat vsfChargeOff, mIntCol销账_申请科室, "申请科室", 1500, flexAlignLeftCenter, "申请科室"
        VsfGridColFormat vsfChargeOff, mIntCol销账_单据, "单据", 0, flexAlignCenterCenter, "单据"
        VsfGridColFormat vsfChargeOff, mIntCol销账_NO, "NO", 1200, flexAlignLeftCenter, "NO"
        VsfGridColFormat vsfChargeOff, mIntCol销账_药品ID, "药品ID", 0, flexAlignLeftCenter, "药品ID"
        VsfGridColFormat vsfChargeOff, mIntCol销账_申请时间, "申请时间", 2000, flexAlignRightCenter, "申请时间"
        VsfGridColFormat vsfChargeOff, mIntCol销账_收发序号, "收发序号", 0, flexAlignLeftCenter, "收发序号"
        
        VsfGridColFormat vsfChargeOff, mIntCol销账_生产商, "生产商", 2000, flexAlignLeftCenter, "生产商"
        VsfGridColFormat vsfChargeOff, mIntCol销账_批号, "批号", 1000, flexAlignLeftCenter, "批号"
        VsfGridColFormat vsfChargeOff, mIntCol销账_效期, "效期", 1500, flexAlignLeftCenter, "效期"
        VsfGridColFormat vsfChargeOff, mIntCol销账_准退数量, "准退数量", 1000, flexAlignRightCenter, "准退数量"
        VsfGridColFormat vsfChargeOff, mIntCol销账_销帐数量, "销帐数量", 1000, flexAlignRightCenter, "销帐数量"
        
        VsfGridColFormat vsfChargeOff, mIntCol销账_包装, "包装", 0, flexAlignLeftCenter, "包装"
        VsfGridColFormat vsfChargeOff, mIntCol销账_单位, "单位", 1000, flexAlignLeftCenter, "单位"
        
        .Redraw = flexRDDirect
    End With
    
End Sub

Private Function GetChargeOffCount(ByVal lng科室ID As Long, ByVal lng药品id As Long, ByVal lng批次 As Long) As Double
    Dim dblSum As Double
    
    With mrsChargeOff
        If mrsChargeOff Is Nothing Then Exit Function

        
        If mcondition.bln按批次汇总 Then
            .Filter = "领药部门id=" & lng科室ID & " And 药品ID=" & lng药品id & " And 批次=" & lng批次 & " And 审核标志 = 1"
        Else
            .Filter = "领药部门id=" & lng科室ID & " And 药品ID=" & lng药品id & " And 审核标志 = 1"
        End If
        If .EOF Then Exit Function
        
        Do While Not .EOF
            Debug.Print mrsChargeOff!批次
            dblSum = dblSum + !销帐数量 / !包装
            .MoveNext
        Loop
        
    End With
    
    GetChargeOffCount = dblSum
End Function
Private Sub InitList_Return()
    Dim i As Integer
    Dim n As Integer
    Dim str列设置 As String
    Dim arr列设置
    
    '''初始化列顺序
    '默认列顺序
    mIntCol退药_当前行 = 0
    mIntCol退药_审查结果 = 1
    mIntCol退药_分组符 = 2
    mIntCol退药_科室 = 3
    mIntCol退药_状态 = 4
    mIntCol退药_类型 = 5
    mIntCol退药_发药类型 = 6
    mIntCol退药_NO = 7
    mIntCol退药_床号 = 8
    mIntCol退药_姓名 = 9
    mIntCol退药_性别 = 10
    mIntCol退药_住院号 = 11
    mIntCol退药_品名 = 12
    mIntCol退药_其它名 = 13
    mIntCol退药_英文名 = 14
    mIntCol退药_规格 = 15
    mIntCol退药_生产商 = 16
    mIntCol退药_原产地 = 17
    mIntCol退药_批号 = 18
    mIntCol退药_效期 = 19
    mIntCol退药_付 = 20
    mIntCol退药_数量 = 21
    mIntCol退药_已退数 = 22
    mIntCol退药_准退数 = 23
    mIntCol退药_退药数 = 24
    mIntCol退药_单价 = 25
    mIntCol退药_金额 = 26
    mIntCol退药_单量 = 27
    mIntCol退药_频次 = 28
    mIntCol退药_用法 = 29
    mIntCol退药_操作员 = 30
    mIntCol退药_发药时间 = 31
    mIntCol退药_领药人 = 32
    mIntCol退药_发药号 = 33
    mIntCol退药_发送时间 = 34
    mIntCol退药_单据 = 35
    mIntCol退药_医嘱id = 36
    
    mIntCol退药_库房货位 = 37
    mIntCol退药_相关ID = 38
    mIntCol退药_药品ID = 39
    mIntCol退药_单量单位 = 40
    mIntCol退药_药品编码和名称 = 41
    mIntCol退药_药品编码 = 42
    mIntCol退药_药品名称 = 43
    mIntCol退药_收发ID = 44
    mIntCol退药_执行状态 = 45
    mIntCol退药_领药部门id = 46
    mIntCol退药_脚注 = 47
    mIntCol退药_病人ID = 48
    mIntCol退药_主页ID = 49
    
    '恢复用户自定义顺序
    str列设置 = LoadListColState(mListType.退药)
    If str列设置 <> "" Then
        arr列设置 = Split(str列设置, "|")
        If UBound(arr列设置) + 1 <> mconIntCol退药_列数 Then
            str列设置 = ""
        Else
            For n = 0 To UBound(arr列设置)
                SetColumnValue mListType.退药, Split(arr列设置(n), ",")(0), n
            Next
        End If
    End If

    '初始化退药清单
    With vsfList(mListType.退药)
        .Redraw = flexRDNone
        .rows = 2
        .Cols = mconIntCol退药_列数
        
        .Cell(flexcpPicture, 1, mIntCol退药_当前行, 1, mIntCol退药_当前行) = Me.ImgList.ListImages(2).Picture
        .Cell(flexcpPictureAlignment, 1, mIntCol退药_当前行, .rows - 1, mIntCol退药_当前行) = flexPicAlignRightCenter
        
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_当前行, "", 250, flexAlignCenterCenter, "当前行"
        
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_审查结果, "警", IIf(mcondition.intShowPass <> 0 And IsInString(gstrprivs, "合理用药监测", ";"), 300, 0), flexAlignCenterCenter, "审查结果"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_分组符, "组", 300, flexAlignCenterCenter, "分组"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_科室, "病人科室", 1200, flexAlignLeftCenter, "病人科室"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_状态, "状态", 1000, flexAlignLeftCenter, "状态"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_类型, "类型", 1000, flexAlignLeftCenter, "类型"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_发药类型, "发药类型", 1000, flexAlignLeftCenter, "发药类型"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_NO, "NO", 900, flexAlignLeftCenter, "NO"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_床号, "床号", 600, flexAlignLeftCenter, "床号"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_姓名, "姓名", 700, flexAlignLeftCenter, "姓名"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_性别, "性别", 700, flexAlignLeftCenter, "性别"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_住院号, "住院号", 1200, flexAlignLeftCenter, "住院号"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_品名, "药品名称", 2500, flexAlignLeftCenter, "药品名称"
        
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_其它名, "其它名", 2000, flexAlignLeftCenter, "其它名"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_英文名, "英文名", 2000, flexAlignLeftCenter, "英文名"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_规格, "规格", 1500, flexAlignLeftCenter, "规格"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_生产商, "生产商", 1500, flexAlignCenterCenter, "生产商"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_原产地, "原产地", 1500, flexAlignCenterCenter, "原产地"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_批号, "批号", 1500, flexAlignLeftCenter, "批号"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_效期, "效期", 1500, flexAlignLeftCenter, "效期"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_付, "付", 300, flexAlignRightCenter, "付"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_数量, "数量", 1000, flexAlignRightCenter, "数量"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_已退数, "已退数", 1000, flexAlignRightCenter, "已退数"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_准退数, "准退数", 1000, flexAlignRightCenter, "准退数"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_退药数, "退药数", 1000, flexAlignRightCenter, "退药数"
        
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_单价, "单价", 1000, flexAlignRightCenter, "单价"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_金额, "金额", 1000, flexAlignRightCenter, "金额"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_单量, "单量", 1000, flexAlignRightCenter, "单量"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_频次, "频次", 500, flexAlignLeftCenter, "频次"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_用法, "用法", 800, flexAlignLeftCenter, "用法"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_操作员, "操作员", 800, flexAlignLeftCenter, "操作员"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_发药时间, "发药时间", 1500, flexAlignLeftCenter, "发药时间"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_单据, "单据", 0, flexAlignCenterCenter, "单据"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_医嘱id, "医嘱id", 0, flexAlignCenterCenter, "医嘱id"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_领药人, "领/退药人", 1000, flexAlignLeftCenter, "领/退药人"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_发药号, "发药号", 1200, flexAlignLeftCenter, "发药号"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_发送时间, "发送时间", 1200, flexAlignLeftCenter, "发送时间"
        
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_单据, "单据", 0, flexAlignCenterCenter, "单据"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_医嘱id, "医嘱id", 0, flexAlignCenterCenter, "医嘱id"
        
        
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_库房货位, "库房货位", IIf(mcondition.bln药品储备 = True, 1200, 0), flexAlignLeftCenter, "库房货位"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_相关ID, "相关ID", 0, flexAlignCenterCenter, "相关ID"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_药品ID, "药品ID", 0, flexAlignCenterCenter, "药品ID"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_单量单位, "单量单位", 0, flexAlignCenterCenter, "单量单位"
        
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_药品编码和名称, "药品编码和名称", 0, flexAlignCenterCenter, "药品编码和名称"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_药品编码, "药品编码", 0, flexAlignCenterCenter, "药品编码"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_药品名称, "药名", 0, flexAlignCenterCenter, "药名"
        
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_收发ID, "收发ID", 0, flexAlignCenterCenter, "收发ID"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_执行状态, "执行状态", 0, flexAlignCenterCenter, "执行状态"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_领药部门id, "领药部门id", 0, flexAlignLeftCenter, "领药部门id"
        
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_脚注, "脚注", 1200, flexAlignLeftCenter, "脚注"
            
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_病人ID, "病人ID", 0, flexAlignLeftCenter, "病人ID"
        VsfGridColFormat vsfList(mListType.退药), mIntCol退药_主页ID, "主页ID", 0, flexAlignLeftCenter, "主页ID"
            
        mstrUnallowSetColHide(mListType.退药) = "药品名称;数量;退药数"
        mstrUnallowShow(mListType.退药) = "当前行;审查结果;分组;单据;医嘱id;相关ID;药品ID;单量单位;药品编码和名称;药品编码;药名;执行状态;收发ID;领药部门id;病人ID;主页ID"
        If mcondition.bln药品储备 = False Then mstrUnallowShow(mListType.退药) = mstrUnallowShow(mListType.退药) & ";库房货位"
        
        '恢复自定义列宽（不包括不允许显示的列）
        If str列设置 <> "" Then
            arr列设置 = Split(str列设置, "|")
            For n = 0 To UBound(arr列设置)
                If IsInString(mstrUnallowShow(mListType.退药), Split(arr列设置(n), ",")(0), ";") = False Then
                    For i = 0 To .Cols - 1
                        If Split(arr列设置(n), ",")(0) = .ColKey(i) Then
                            .ColWidth(i) = Val(Split(arr列设置(n), ",")(1))
                        End If
                    Next
                End If
            Next
        End If
        
        If mcondition.bln显示原产地 = False Then VsfGridColFormat vsfList(mListType.退药), mIntCol退药_原产地, "原产地", 0, flexAlignCenterCenter, "原产地"
        
        .Redraw = flexRDDirect
    End With
    
End Sub
Private Sub InitList_Sum()
    Dim int当前行 As Integer
    Dim i As Integer
    Dim n As Integer
    Dim str列设置 As String
    Dim arr列设置
    
    '''初始化列顺序
    '默认列顺序
    mIntCol汇总_当前行 = 0
    mIntCol汇总_品名 = 1
    mIntCol汇总_规格 = 2
    mIntCol汇总_生产商 = 3
    mIntCol汇总_原产地 = 4
    mIntCol汇总_批号 = 5
    mIntCol汇总_效期 = 6
    mIntCol汇总_数量 = 7
    mIntCol汇总_单位 = 8
    mIntCol汇总_单价 = 9
    mIntCol汇总_金额 = 10
    mIntCol汇总_药品编码和名称 = 11
    mIntCol汇总_药品编码 = 12
    mIntCol汇总_药品名称 = 13
    
    mIntCol科室汇总_当前行 = 0
    mIntCol科室汇总_领药部门 = 1
    mIntCol科室汇总_科室 = 2
    mIntCol科室汇总_品名 = 3
    mIntCol科室汇总_规格 = 4
    mIntCol科室汇总_生产商 = 5
    mIntCol科室汇总_原产地 = 6
    mIntCol科室汇总_批号 = 7
    mIntCol科室汇总_效期 = 8
    mIntCol科室汇总_应发数量 = 9
    mIntCol科室汇总_留存数量 = 10
    mIntCol科室汇总_销帐数量 = 11
    mIntCol科室汇总_实发数量 = 12
    mIntCol科室汇总_单位 = 13
    mIntCol科室汇总_单价 = 14
    mIntCol科室汇总_应发金额 = 15
    mIntCol科室汇总_实发金额 = 16
    mIntCol科室汇总_批次 = 17
    mIntCol科室汇总_科室ID = 18
    mIntCol科室汇总_药品ID = 19
    mIntCol科室汇总_领药部门id = 20
    mIntCol科室汇总_药品编码和名称 = 21
    mIntCol科室汇总_药品编码 = 22
    mIntCol科室汇总_药品名称 = 23
    mIntCol科室汇总_包装 = 24
    
    '恢复用户自定义顺序
    str列设置 = LoadListColState(mListType.汇总)
    If str列设置 <> "" Then
        arr列设置 = Split(str列设置, "|")
        If UBound(arr列设置) + 1 <> IIf(mcondition.bln按科室汇总 = True, mconIntCol科室汇总_列数, mconIntCol汇总_列数) Then
            str列设置 = ""
        Else
            For n = 0 To UBound(arr列设置)
                SetColumnValue mListType.汇总, Split(arr列设置(n), ",")(0), n
            Next
        End If
    End If
    
    '''初始化汇总发药清单
    With vsfList(mListType.汇总)
        .Redraw = flexRDNone
        .rows = 2
        .Cols = IIf(mcondition.bln按科室汇总, mconIntCol科室汇总_列数, mconIntCol汇总_列数)
        
        If mcondition.bln按科室汇总 Then
            int当前行 = mIntCol科室汇总_当前行
        Else
            int当前行 = mIntCol汇总_当前行
        End If
        
        .Cell(flexcpPicture, 1, int当前行, 1, int当前行) = Me.ImgList.ListImages(2).Picture
        .Cell(flexcpPictureAlignment, 1, int当前行, .rows - 1, int当前行) = flexPicAlignRightCenter
        
        VsfGridColFormat vsfList(mListType.汇总), int当前行, "", 250, flexAlignCenterCenter, "当前行"
        
        If mcondition.bln按科室汇总 = False Then
            VsfGridColFormat vsfList(mListType.汇总), mIntCol汇总_品名, "药品名称", 2500, flexAlignLeftCenter, "药品名称"
            VsfGridColFormat vsfList(mListType.汇总), mIntCol汇总_规格, "规格", 1500, flexAlignLeftCenter, "规格"
            VsfGridColFormat vsfList(mListType.汇总), mIntCol汇总_生产商, "生产商", 1500, flexAlignLeftCenter, "生产商"
            VsfGridColFormat vsfList(mListType.汇总), mIntCol汇总_原产地, "原产地", 1500, flexAlignLeftCenter, "原产地"
            VsfGridColFormat vsfList(mListType.汇总), mIntCol汇总_批号, "批号", 1200, flexAlignLeftCenter, "批号"
            VsfGridColFormat vsfList(mListType.汇总), mIntCol汇总_效期, "效期", 1200, flexAlignLeftCenter, "效期"
            VsfGridColFormat vsfList(mListType.汇总), mIntCol汇总_数量, "数量", 1200, flexAlignRightCenter, "数量"
            VsfGridColFormat vsfList(mListType.汇总), mIntCol汇总_单位, "单位", 500, flexAlignCenterCenter, "单位"
            VsfGridColFormat vsfList(mListType.汇总), mIntCol汇总_单价, "单价", 1200, flexAlignRightCenter, "单价"
            VsfGridColFormat vsfList(mListType.汇总), mIntCol汇总_金额, "金额", 1200, flexAlignRightCenter, "金额"
            
            VsfGridColFormat vsfList(mListType.汇总), mIntCol汇总_药品编码和名称, "药品编码和名称", 0, flexAlignLeftCenter, "药品编码和名称"
            VsfGridColFormat vsfList(mListType.汇总), mIntCol汇总_药品编码, "药品编码", 0, flexAlignLeftCenter, "药品编码"
            VsfGridColFormat vsfList(mListType.汇总), mIntCol汇总_药品名称, "药名", 0, flexAlignLeftCenter, "药名"
        Else
            VsfGridColFormat vsfList(mListType.汇总), mIntCol科室汇总_科室, "开单科室", 1200, flexAlignLeftCenter, "开单科室"
            VsfGridColFormat vsfList(mListType.汇总), mIntCol科室汇总_品名, "药品名称", 2500, flexAlignLeftCenter, "药品名称"
            VsfGridColFormat vsfList(mListType.汇总), mIntCol科室汇总_规格, "规格", 1500, flexAlignLeftCenter, "规格"
            VsfGridColFormat vsfList(mListType.汇总), mIntCol科室汇总_生产商, "生产商", 1500, flexAlignLeftCenter, "生产商"
            VsfGridColFormat vsfList(mListType.汇总), mIntCol科室汇总_原产地, "原产地", 1500, flexAlignLeftCenter, "原产地"
            VsfGridColFormat vsfList(mListType.汇总), mIntCol科室汇总_批号, "批号", 1200, flexAlignLeftCenter, "批号"
            VsfGridColFormat vsfList(mListType.汇总), mIntCol科室汇总_效期, "效期", 1200, flexAlignLeftCenter, "效期"
            
            VsfGridColFormat vsfList(mListType.汇总), mIntCol科室汇总_应发数量, "应发数量", 1200, flexAlignRightCenter, "应发数量"
            VsfGridColFormat vsfList(mListType.汇总), mIntCol科室汇总_留存数量, "留存数量", 1200, flexAlignRightCenter, "留存数量"
            VsfGridColFormat vsfList(mListType.汇总), mIntCol科室汇总_销帐数量, "销帐数量", IIf(mcondition.bln启用退药销账 = True, 1200, 0), flexAlignRightCenter, "销帐数量"
            VsfGridColFormat vsfList(mListType.汇总), mIntCol科室汇总_实发数量, "实发数量", 1200, flexAlignRightCenter, "实发数量"

            VsfGridColFormat vsfList(mListType.汇总), mIntCol科室汇总_单位, "单位", 500, flexAlignCenterCenter, "单位"
            VsfGridColFormat vsfList(mListType.汇总), mIntCol科室汇总_单价, "单价", 1200, flexAlignRightCenter, "单价"
            VsfGridColFormat vsfList(mListType.汇总), mIntCol科室汇总_应发金额, "应发金额", 1200, flexAlignRightCenter, "应发金额"
            VsfGridColFormat vsfList(mListType.汇总), mIntCol科室汇总_实发金额, "实发金额", 1200, flexAlignRightCenter, "实发金额"
            VsfGridColFormat vsfList(mListType.汇总), mIntCol科室汇总_批次, "批次", 0, flexAlignRightCenter, "批次"

            VsfGridColFormat vsfList(mListType.汇总), mIntCol科室汇总_科室ID, "科室ID", 0, flexAlignCenterCenter, "科室ID"
            VsfGridColFormat vsfList(mListType.汇总), mIntCol科室汇总_药品ID, "药品ID", 0, flexAlignLeftCenter, "药品ID"
            VsfGridColFormat vsfList(mListType.汇总), mIntCol科室汇总_领药部门, "领药部门", 1200, flexAlignLeftCenter, "领药部门"
            VsfGridColFormat vsfList(mListType.汇总), mIntCol科室汇总_领药部门id, "领药部门id", 0, flexAlignLeftCenter, "领药部门id"
            
            VsfGridColFormat vsfList(mListType.汇总), mIntCol科室汇总_药品编码和名称, "药品编码和名称", 0, flexAlignLeftCenter, "药品编码和名称"
            VsfGridColFormat vsfList(mListType.汇总), mIntCol科室汇总_药品编码, "药品编码", 0, flexAlignLeftCenter, "药品编码"
            VsfGridColFormat vsfList(mListType.汇总), mIntCol科室汇总_药品名称, "药名", 0, flexAlignLeftCenter, "药名"
            
            VsfGridColFormat vsfList(mListType.汇总), mIntCol科室汇总_包装, "包装", 0, flexAlignLeftCenter, "包装"
        End If
        
        mstrUnallowSetColHide(mListType.汇总) = "药品名称;数量;应发数量;实发数量;单位"
        If mcondition.bln启用退药销账 = True Then mstrUnallowSetColHide(mListType.汇总) = mstrUnallowSetColHide(mListType.汇总) & ";销帐数量"
        
        mstrUnallowShow(mListType.汇总) = "当前行;批次;科室ID;药品ID;领药部门id;药品编码和名称;药品编码;药名;包装"
        If mcondition.bln启用退药销账 = False Then mstrUnallowShow(mListType.汇总) = mstrUnallowShow(mListType.汇总) & ";销帐数量"
        
        '恢复自定义列宽（不包括不允许显示的列）
        If str列设置 <> "" Then
            arr列设置 = Split(str列设置, "|")
            For n = 0 To UBound(arr列设置)
                If IsInString(mstrUnallowShow(mListType.汇总), Split(arr列设置(n), ",")(0), ";") = False Then
                    For i = 0 To .Cols - 1
                        If Split(arr列设置(n), ",")(0) = .ColKey(i) Then
                            If IsInString(mstrUnallowSetColHide(mListType.汇总), Split(arr列设置(n), ",")(0), ";") = True Then
                                '如果是不允许隐藏的列，则列宽不能为0
                                If Val(Split(arr列设置(n), ",")(1)) <> 0 Then
                                    .ColWidth(i) = Val(Split(arr列设置(n), ",")(1))
                                End If
                            Else
                                .ColWidth(i) = Val(Split(arr列设置(n), ",")(1))
                            End If
                        End If
                    Next
                End If
            Next
        End If
        
        If mcondition.bln按科室汇总 = False Then
            If mcondition.bln显示原产地 = False Then VsfGridColFormat vsfList(mListType.汇总), mIntCol汇总_原产地, "原产地", 0, flexAlignLeftCenter, "原产地"
        Else
            If mcondition.bln显示原产地 = False Then VsfGridColFormat vsfList(mListType.汇总), mIntCol科室汇总_原产地, "原产地", 0, flexAlignLeftCenter, "原产地"
        End If
        
        .Redraw = flexRDDirect
    End With
    
End Sub

Private Sub SetGroup(ByVal Bill As VSFlexGrid, ByVal bln是否分组 As Boolean)
    Dim n As Integer
    Dim str上行相关 As String
    Dim str本行相关 As String
    Dim str下行相关 As String
    Dim int列名_相关ID As Integer
    Dim int列名_NO As Integer
    Dim int列名_分组符 As Integer
    Dim bln是否存在分组 As Boolean
    Dim bln汇总行分组符 As Boolean
    
    '总行数小于四行时没有必要分组（1行固定行＋2行汇总行）
    If Bill.rows < 4 Then Exit Sub
    
    str上行相关 = "-1"
        
    '按相关ID分组
    With Bill
        Select Case Bill.index
        Case mListType.发药
            int列名_相关ID = mIntCol发药_相关ID
            int列名_NO = mIntCol发药_NO
            int列名_分组符 = mIntCol发药_分组符
        Case mListType.退药
            int列名_相关ID = mIntCol退药_相关ID
            int列名_NO = mIntCol退药_NO
            int列名_分组符 = mIntCol退药_分组符
        End Select
        
        .Redraw = flexRDNone
        
        .Cell(flexcpPicture, 1, int列名_分组符, .rows - 1, int列名_分组符) = Nothing
                
        If Not bln是否分组 Then
            .ColWidth(int列名_分组符) = 0
            .Redraw = flexRDDirect
            Exit Sub
        Else
            .ColWidth(int列名_分组符) = 250
        End If
        
        For n = 1 To .rows - 1
            .Row = n
            .Col = int列名_分组符
            If .IsSubtotal(n) = False And .TextMatrix(n, int列名_相关ID) <> "" Then
                str本行相关 = IIf(.TextMatrix(n, int列名_相关ID) = 0, "0", .TextMatrix(n, int列名_NO) & .TextMatrix(n, int列名_相关ID))
                If n + 1 <= .rows - 1 Then
                    If .IsSubtotal(n + 1) = False And .TextMatrix(n + 1, int列名_相关ID) <> "" Then  '如果下行为记录行时
                        str下行相关 = IIf(.TextMatrix(n + 1, int列名_相关ID) = 0, "-1", .TextMatrix(n + 1, int列名_NO) & .TextMatrix(n + 1, int列名_相关ID))
                    ElseIf n + 2 <= .rows - 1 Then  '如果下行为汇总行行时
                        If .IsSubtotal(n + 2) = False And .TextMatrix(n + 2, int列名_相关ID) <> "" Then    '如果下下行为记录行时
                            str下行相关 = IIf(.TextMatrix(n + 2, int列名_相关ID) = 0, "-1", .TextMatrix(n + 2, int列名_NO) & .TextMatrix(n + 2, int列名_相关ID))
                        Else
                            str下行相关 = "-1"
                        End If
                    Else
                        str下行相关 = "-1"
                    End If
                Else
                    str下行相关 = "-1"
                End If
                
                If str本行相关 = str上行相关 Then
                    If str本行相关 = str下行相关 Then
                        .Cell(flexcpPicture, n, int列名_分组符) = imgGroup.ListImages(2).Picture
                    Else
                        .Cell(flexcpPicture, n, int列名_分组符) = imgGroup.ListImages(3).Picture
                    End If
                ElseIf str本行相关 = str下行相关 Then
                        .Cell(flexcpPicture, n, int列名_分组符) = imgGroup.ListImages(1).Picture
                    bln是否存在分组 = True
                End If
            
                str上行相关 = IIf(str本行相关 = "0", "-1", str本行相关)
            Else
                '如果该行是汇总行，则要根据下行的相关ID判断分组符号
                If n + 1 <= .rows - 1 Then
                    If .IsSubtotal(n + 1) = False And .TextMatrix(n + 1, int列名_相关ID) <> "" Then
                        If str上行相关 <> "-1" And str上行相关 = IIf(.TextMatrix(n + 1, int列名_相关ID) = 0, "-1", .TextMatrix(n + 1, int列名_NO) & .TextMatrix(n + 1, int列名_相关ID)) Then
                            .Cell(flexcpPicture, n, int列名_分组符) = imgGroup.ListImages(2).Picture
                        End If
                    End If
                End If
            End If
        Next
        
        .Cell(flexcpPictureAlignment, 1, int列名_分组符, .rows - 1, int列名_分组符) = flexAlignRightCenter
        
        If Not bln是否存在分组 Then .ColWidth(int列名_分组符) = 0
        
        .Redraw = flexRDDirect

    End With
    
End Sub

Private Sub Load发药单格式()
    Dim rsData As ADODB.Recordset
    
    Set rsData = DeptSendWork_Get发药单格式("ZL1_BILL_1342")
    
    With cbo发药单格式
        .Clear
        
        Do While Not rsData.EOF
            .AddItem rsData!格式
            rsData.MoveNext
        Loop
        
        If mcondition.int发药单格式 <= .ListCount - 1 And mcondition.int发药单格式 >= 0 Then
            .ListIndex = mcondition.int发药单格式
        Else
            .ListIndex = 0
        End If
        
        If rsData.RecordCount = 1 Then
            .Enabled = False
        End If
    End With
End Sub

Private Sub Load配药人(ByVal lng药房id As Long)
    Dim rsData As ADODB.Recordset
    Dim intIndex As Integer
    
    Set rsData = DeptSendWork_Get配药人(lng药房id)
    
    With rsData
        Me.cbo配药人.Clear
        If .EOF Then Exit Sub
        Do While Not .EOF
            cbo配药人.AddItem !姓名

            If gstrUserName = !名称 Then
                intIndex = .AbsolutePosition - 1
            End If

            .MoveNext
        Loop

        cbo配药人.Enabled = Not cbo配药人.ListCount = 0

        If intIndex <> -1 Then cbo配药人.ListIndex = intIndex
    End With
End Sub

Public Sub Load核查人(ByVal lng药房id As Long)
    Dim rsData As ADODB.Recordset
    
    Set rsData = DeptSendWork_Get核查人(lng药房id)
    
    With cbo核查人
        .Clear
        
        Do While Not rsData.EOF
            .AddItem rsData!姓名
            rsData.MoveNext
        Loop
    End With
    
    cbo核查人.Text = gstrUserAbbr & "-" & gstrUserName
End Sub
Public Sub RefreshList(ByVal intType As Integer, ByVal rsData As ADODB.Recordset, Optional ByVal rsChargeOff As ADODB.Recordset)
    '刷新列表
    
    Select Case intType
        Case mListType.发药
            '创建发药数据集的副本
            Set mrsSendList = rsData
            
            '创建销账数据集的副本
            Set mrsChargeOff = rsChargeOff
    
            '发药状态时，同时刷新待发药、汇总发药、缺药、拒发等列表
            mblnRefresh = True
            
            '根据界面选项来改变退药待发单据的发药状态
            Modify退药待发 mcondition.bln显示退药待发单据
            
            Call InitList_Send
            Call RefreshList_Send
            
            Call InitList_ChargeOff
            
            Call InitList_Sum
            Call RefreshList_Sum
            
            Call InitList_Shortage
            Call RefreshList_Shortage
            
            Call InitList_Reject
            Call RefreshList_Reject
            
            mblnRefresh = False
        Case mListType.退药
            '创建退药数据集的副本
            Set mrsReturnList = rsData
            
            Call InitList_Return
            Call RefreshList_Return
    End Select
    
    Call InitColSelList(mcondition.intListType)
    Form_Resize
End Sub


Private Sub RefreshList_Send()
    '刷新待发药列表
    Dim lngRow As Long
    Dim str科室 As String
    Dim lngStateColor As Double
    Dim strFilter As String
    Dim i As Long
    Dim dateCurrent As Date
    
    If mrsSendList Is Nothing Then Exit Sub
    
    dateCurrent = Sys.Currentdate
    
    '是否显示退药待发药品
    If mcondition.bln显示退药待发单据 = True Then
        '显示正常的发药药品
        strFilter = strFilter & "执行状态=" & mState.发药
        
        '是否显示缺药药品
        If mcondition.bln显示缺药 = True Then
           strFilter = strFilter & " Or 执行状态=" & mState.缺药
        End If
    
        '显示不处理、拒发的药品（上次操作的药品）
        strFilter = strFilter & " Or 执行状态=" & mState.不处理 & " Or 执行状态=" & mState.拒发
    Else
        '显示正常的发药药品
        strFilter = "(记录状态=1 And 执行状态=" & mState.发药 & ")"
        
        '是否显示缺药药品
        If mcondition.bln显示缺药 = True Then
           strFilter = strFilter & " Or (记录状态=1 And 执行状态=" & mState.缺药 & ")"
        End If
        
        '显示不处理、拒发的药品（上次操作的药品）
        strFilter = strFilter & " Or (记录状态=1 And 执行状态=" & mState.不处理 & ")" & " Or (记录状态=1 And 执行状态=" & mState.拒发 & ")"
    End If

    With vsfList(mListType.发药)
        mrsSendList.Filter = strFilter
'        mrsSendList.Filter = "(记录状态=1 And 执行状态=1) Or (记录状态=1 and 执行状态=0) Or (记录状态=1 and 执行状态=3) Or (记录状态=1 and 执行状态=2)"
'        mrsSendList.Sort = "领药部门,领药号,姓名,NO,序号"
        mrsSendList.Sort = "领药部门,NO,相关ID"
        
        .Redraw = flexRDNone
        .MergeCells = flexMergeNever
        .rows = 1
        
        .Subtotal flexSTClear
        
        If mrsSendList.EOF Then
            .rows = 2
            .Cell(flexcpText, 1, mIntCol发药_科室, 1, .Cols - 1) = "没有找到满足条件的记录......"
            .MergeCells = flexMergeRestrictRows
            .MergeRow(1) = True
        Else
            Do While Not mrsSendList.EOF
                lngRow = lngRow + 1
                .rows = lngRow + 1
                
                .TextMatrix(lngRow, mIntCol发药_分组符) = ""
                .TextMatrix(lngRow, mIntCol发药_科室) = mrsSendList!科室
                .TextMatrix(lngRow, mIntCol发药_开单医生) = IIf(IsNull(mrsSendList!开单医生), "", mrsSendList!开单医生)
                
                .TextMatrix(lngRow, mIntCol发药_类型) = mrsSendList!类型
                .TextMatrix(lngRow, mIntCol发药_发药类型) = IIf(InStr(1, mrsSendList!扣率, "3") > 1, "离院带药", IIf(InStr(1, mrsSendList!扣率, "4") > 1, "自取药", "院内用药"))
                .TextMatrix(lngRow, mIntCol发药_NO) = mrsSendList!NO
                .TextMatrix(lngRow, mIntCol发药_记帐员) = mrsSendList!记帐员
                .TextMatrix(lngRow, mIntCol发药_床号) = IIf(IsNull(mrsSendList!床号), "", mrsSendList!床号)
                .TextMatrix(lngRow, mIntCol发药_姓名) = mrsSendList!姓名
                .TextMatrix(lngRow, mIntCol发药_性别) = mrsSendList!性别
                .TextMatrix(lngRow, mIntCol发药_病人类型) = zlStr.NVL(mrsSendList!病人类型)
                .Cell(flexcpForeColor, lngRow, mIntCol发药_姓名, lngRow, mIntCol发药_姓名) = zlStr.NVL(mrsSendList!颜色, 0)
                
                .TextMatrix(lngRow, mIntCol发药_年龄) = IIf(IsNull(mrsSendList!年龄), "", mrsSendList!年龄)
                
                .TextMatrix(lngRow, mIntCol发药_住院号) = IIf(IsNull(mrsSendList!住院号), "", mrsSendList!住院号)
                
                If mrsSendList!抗生素 <> 0 Then
                    .Cell(flexcpPicture, lngRow, mIntCol发药_品名) = Me.ImgList.ListImages(39).Picture
                    .Cell(flexcpPictureAlignment, lngRow, mIntCol发药_品名) = flexPicAlignLeftCenter
                End If
                
                If mrsSendList!高危药品 > 0 Then
                    .Cell(flexcpPicture, lngRow, mIntCol发药_品名) = Me.ImgList.ListImages("高危").Picture
                    .Cell(flexcpPictureAlignment, lngRow, mIntCol发药_品名) = flexPicAlignLeftCenter
                End If
                
                If mcondition.int药品名称编码显示 = 0 Then
                    .TextMatrix(lngRow, mIntCol发药_品名) = mrsSendList!品名
                ElseIf mcondition.int药品名称编码显示 = 1 Then
                    .TextMatrix(lngRow, mIntCol发药_品名) = mrsSendList!药品编码
                Else
                    .TextMatrix(lngRow, mIntCol发药_品名) = mrsSendList!药品名称
                End If
                
                .TextMatrix(lngRow, mIntCol发药_其它名) = IIf(IsNull(mrsSendList!其它名), "", mrsSendList!其它名)
                .TextMatrix(lngRow, mIntCol发药_英文名) = IIf(IsNull(mrsSendList!英文名), "", mrsSendList!英文名)
                .TextMatrix(lngRow, mIntCol发药_配方名称) = IIf(IsNull(mrsSendList!配方名称), "", mrsSendList!配方名称)
                .TextMatrix(lngRow, mIntCol发药_规格) = IIf(IsNull(mrsSendList!规格), "", mrsSendList!规格)
                .TextMatrix(lngRow, mIntCol发药_生产商) = IIf(IsNull(mrsSendList!产地), "", mrsSendList!产地)
                .TextMatrix(lngRow, mIntCol发药_原产地) = IIf(IsNull(mrsSendList!原产地), "", mrsSendList!原产地)
                .TextMatrix(lngRow, mIntCol发药_批号) = IIf(IsNull(mrsSendList!批号), "", mrsSendList!批号)
                .TextMatrix(lngRow, mIntCol发药_效期) = IIf(IsNull(mrsSendList!效期), "", mrsSendList!效期)
                .TextMatrix(lngRow, mIntCol发药_付) = mrsSendList!付
                .TextMatrix(lngRow, mIntCol发药_数量) = mrsSendList!数量
                .TextMatrix(lngRow, mIntCol发药_单价) = Format(mrsSendList!单价, "#0." & String(mintPriceDigit, "0"))
                
                .TextMatrix(lngRow, mIntCol发药_金额) = zlStr.FormatEx(mrsSendList!金额, mintMoneyDigit, , True)
                .TextMatrix(lngRow, mIntCol发药_单量) = mrsSendList!单量
                .TextMatrix(lngRow, mIntCol发药_频次) = IIf(IsNull(mrsSendList!频次), "", mrsSendList!频次)
                .TextMatrix(lngRow, mIntCol发药_用法) = IIf(IsNull(mrsSendList!用法), "", mrsSendList!用法)
                .TextMatrix(lngRow, mIntCol发药_用药次数) = IIf(IsNull(mrsSendList!用药次数), "", mrsSendList!用药次数)
                .TextMatrix(lngRow, mIntCol发药_用药目的) = mrsSendList!用药目的
                .TextMatrix(lngRow, mIntCol发药_记帐时间) = mrsSendList!记帐时间
                .TextMatrix(lngRow, mIntCol发药_说明) = IIf(IsNull(mrsSendList!说明), "", mrsSendList!说明)
                .TextMatrix(lngRow, mIntCol发药_单据) = mrsSendList!单据
                .TextMatrix(lngRow, mIntCol发药_医嘱id) = mrsSendList!医嘱id
                .TextMatrix(lngRow, mIntCol发药_退药人) = ""
                .TextMatrix(lngRow, mIntCol发药_库房货位) = IIf(IsNull(mrsSendList!库房货位), "", mrsSendList!库房货位)
                
                .TextMatrix(lngRow, mIntCol发药_相关ID) = IIf(IsNull(mrsSendList!相关ID), 0, mrsSendList!相关ID)
                .TextMatrix(lngRow, mIntCol发药_药品ID) = mrsSendList!药品ID
                .TextMatrix(lngRow, mIntCol发药_单量单位) = mrsSendList!单量单位
                .TextMatrix(lngRow, mIntCol发药_领药部门) = mrsSendList!领药部门
                .TextMatrix(lngRow, mIntCol发药_领药部门id) = mrsSendList!领药部门ID
                
                .TextMatrix(lngRow, mIntCol发药_药品编码和名称) = mrsSendList!药品编码和名称
                .TextMatrix(lngRow, mIntCol发药_药品编码) = mrsSendList!药品编码
                .TextMatrix(lngRow, mIntCol发药_药品名称) = mrsSendList!药品名称
                .TextMatrix(lngRow, mIntCol发药_禁忌药品说明) = zlStr.NVL(mrsSendList!禁忌药品说明)
                
                .TextMatrix(lngRow, mIntCol发药_收发ID) = mrsSendList!收发ID
                
                .TextMatrix(lngRow, mIntCol发药_领药号) = IIf(IsNull(mrsSendList!领药号), "", mrsSendList!领药号)
                .TextMatrix(lngRow, mIntCol发药_已收费) = mrsSendList!已收费
                
                If mrsSendList!是否皮试 = 1 Then
                    .TextMatrix(lngRow, mIntCol发药_皮试结果) = Get皮试结果(mrsSendList!病人ID, mrsSendList!药名ID, dateCurrent, mrsSendList!开嘱时间, mrsSendList!主页ID)
                End If
                
                .TextMatrix(lngRow, mIntCol发药_病人ID) = mrsSendList!病人ID
                .TextMatrix(lngRow, mIntCol发药_主页ID) = mrsSendList!主页ID
                .TextMatrix(lngRow, mIntCol发药_用药理由) = zlStr.NVL(mrsSendList!用药理由)
                .TextMatrix(lngRow, mIntCol发药_高危药品) = zlStr.NVL(mrsSendList!高危药品, 0)
                .TextMatrix(lngRow, mIntCol发药_结果) = NVL(mrsSendList!审查结果, 0)
                .TextMatrix(lngRow, mIntCol发药_开单部门id) = NVL(mrsSendList!开单部门id, 0)
                                
                .TextMatrix(lngRow, mIntCol发药_状态) = mrsSendList!状态
                .TextMatrix(lngRow, mIntCol发药_执行状态) = mrsSendList!执行状态
                            
                .TextMatrix(lngRow, mIntCol发药_脚注) = NVL(mrsSendList!医生嘱托, "")
                            
                '设置状态的颜色
                If mrsSendList!执行状态 = mState.缺药 Then
                    lngStateColor = mListColor.State_Shortage
                ElseIf mrsSendList!执行状态 = mState.发药 Then
                    lngStateColor = mListColor.State_Send
                ElseIf mrsSendList!执行状态 = mState.拒发 Then
                    lngStateColor = mListColor.State_Reject
                ElseIf mrsSendList!执行状态 = mState.不处理 Then
                    lngStateColor = mListColor.State_UnProcess
                End If
                
                .Cell(flexcpBackColor, lngRow, 1, lngRow, .Cols - 1) = lngStateColor
                
'                设置合理用药标志 (PASS)
                If Not gobjPass Is Nothing Then
                    .Cell(flexcpPicture, lngRow, mIntCol发药_审查结果, lngRow, mIntCol发药_审查结果) = gobjPass.zlPassSetWarnLight_YF(Val(mrsSendList!审查结果))
                    .Cell(flexcpPictureAlignment, lngRow, mIntCol发药_审查结果, lngRow, mIntCol发药_审查结果) = flexPicAlignCenterCenter
                End If
                
                '以下用于测试PASS
'                .Cell(flexcpPicture, lngRow, mIntCol发药_审查结果, lngRow, mIntCol发药_审查结果) = frmPublic.imgPass.ListImages(Val(mrsSendList!审查结果) + 2).Picture
'                .Cell(flexcpPictureAlignment, lngRow, mIntCol发药_审查结果, lngRow, mIntCol发药_审查结果) = flexPicAlignCenterCenter
                
                '特殊药品粗体显示
                If IsInString("毒性药;麻醉药;精神I类;精神II类", zlStr.NVL(mrsSendList!毒理分类), ";") = True And zlStr.NVL(mrsSendList!毒理分类) <> "" Then
                    .Cell(flexcpFontBold, lngRow, mIntCol发药_品名, lngRow, mIntCol发药_品名) = True
                End If
                
                '检查库存下限
                If mcondition.bln药品储备 = True And mrsSendList!库存下限 = 0 Then
                    .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = mListColor.LowerLimit
                End If
                
                '该医嘱是否在部门发药中进行过发药操作
                If mrsSendList!药师审核标志 = 0 Then
                    .Cell(flexcpForeColor, lngRow, 0, lngRow, .Cols - 1) = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\界面设置\" & App.ProductName & "\1345", "未审核医嘱颜色", 33023)
                End If
                
                mrsSendList.MoveNext
            Loop
            
            '设置数量列为加粗显示
            .Cell(flexcpFontBold, 1, mIntCol发药_数量, .rows - 1, mIntCol发药_数量) = True
            
            '皮试结果标示
            For i = 1 To .rows - 1
                If .IsSubtotal(i) = False Then
                    If .TextMatrix(i, mIntCol发药_皮试结果) = "(+)" Then
                        .Cell(flexcpForeColor, i, mIntCol发药_皮试结果, i, mIntCol发药_皮试结果) = vbRed
                    ElseIf .TextMatrix(i, mIntCol发药_皮试结果) = "(-)" Then
                        .Cell(flexcpForeColor, i, mIntCol发药_皮试结果, i, mIntCol发药_皮试结果) = vbBlue
                    Else
                        .Cell(flexcpForeColor, i, mIntCol发药_皮试结果, i, mIntCol发药_皮试结果) = &H80000008
                    End If
                End If
            Next
                    
            SetSubTotal vsfList(mListType.发药), "领药部门"
        End If
        
        .Redraw = flexRDDirect
    End With
    
    SetGroup vsfList(mListType.发药), True
End Sub

Private Sub RefreshList_Return()
    '刷新退药列表
    
    Dim intRow As Integer
    Dim lngStateColor As Double
    Dim strFilter As String
    
    If mrsReturnList Is Nothing Then Exit Sub
    
    '显示正常的退药药品
    strFilter = "执行状态=" & mState.退药_原始记录 & " And 准退数>0 "
    
    '是否显示所有过程单据
    If mcondition.bln显示过程单据 = True Then
        strFilter = "执行状态=" & mState.退药_原始记录 & " Or 执行状态=" & mState.退药_发药记录 & " Or 执行状态=" & mState.退药_退药记录
    End If
    
    strFilter = strFilter & " Or 执行状态=" & mState.退药
    
    With vsfList(mListType.退药)
        mrsReturnList.Filter = strFilter
'        mrsReturnList.Sort = "科室,发药号,姓名,NO,序号"
        mrsReturnList.Sort = "科室,NO,相关ID"
        
        .Redraw = flexRDNone
        .MergeCells = flexMergeNever
        .rows = 1
        
        If mrsReturnList.EOF Then
            .rows = 2
            .Cell(flexcpText, 1, mIntCol退药_科室, 1, .Cols - 1) = "没有找到满足条件的记录......"
            .MergeCells = flexMergeRestrictRows
            .MergeRow(1) = True
        Else
            Do While Not mrsReturnList.EOF
                intRow = intRow + 1
                .rows = intRow + 1
                
                .TextMatrix(intRow, mIntCol退药_科室) = mrsReturnList!科室
                .TextMatrix(intRow, mIntCol退药_类型) = mrsReturnList!类型
                .TextMatrix(intRow, mIntCol退药_发药类型) = IIf(Right(mrsReturnList!扣率, 1) = 3, "离院带药", IIf(Right(mrsReturnList!扣率, 1) = 4, "自取药", "院内用药"))
                .TextMatrix(intRow, mIntCol退药_NO) = mrsReturnList!NO
                
                .TextMatrix(intRow, mIntCol退药_床号) = IIf(IsNull(mrsReturnList!床号), "", mrsReturnList!床号)
                .TextMatrix(intRow, mIntCol退药_姓名) = mrsReturnList!姓名
                .TextMatrix(intRow, mIntCol退药_性别) = mrsReturnList!性别
                .TextMatrix(intRow, mIntCol退药_住院号) = IIf(IsNull(mrsReturnList!住院号), "", mrsReturnList!住院号)
                
                If mcondition.int药品名称编码显示 = 0 Then
                    .TextMatrix(intRow, mIntCol退药_品名) = mrsReturnList!品名
                ElseIf mcondition.int药品名称编码显示 = 1 Then
                    .TextMatrix(intRow, mIntCol退药_品名) = mrsReturnList!药品编码
                Else
                    .TextMatrix(intRow, mIntCol退药_品名) = mrsReturnList!药品名称
                End If
                
                If mrsReturnList!高危药品 > 0 Then
                    .Cell(flexcpPicture, intRow, mIntCol退药_品名) = Me.ImgList.ListImages("高危").Picture
                    .Cell(flexcpPictureAlignment, intRow, mIntCol退药_品名) = flexPicAlignLeftCenter
                End If
                
                .TextMatrix(intRow, mIntCol退药_其它名) = IIf(IsNull(mrsReturnList!其它名), "", mrsReturnList!其它名)
                
                .TextMatrix(intRow, mIntCol退药_英文名) = IIf(IsNull(mrsReturnList!英文名), "", mrsReturnList!英文名)
                .TextMatrix(intRow, mIntCol退药_规格) = IIf(IsNull(mrsReturnList!规格), "", mrsReturnList!规格)
                .TextMatrix(intRow, mIntCol退药_生产商) = IIf(IsNull(mrsReturnList!产地), "", mrsReturnList!产地)
                .TextMatrix(intRow, mIntCol退药_原产地) = IIf(IsNull(mrsReturnList!原产地), "", mrsReturnList!原产地)
                .TextMatrix(intRow, mIntCol退药_批号) = IIf(IsNull(mrsReturnList!批号), "", mrsReturnList!批号)
                .TextMatrix(intRow, mIntCol退药_效期) = IIf(IsNull(mrsReturnList!效期), "", mrsReturnList!效期)
                .TextMatrix(intRow, mIntCol退药_付) = mrsReturnList!付
                
                .TextMatrix(intRow, mIntCol退药_数量) = mrsReturnList!数量
                .TextMatrix(intRow, mIntCol退药_已退数) = mrsReturnList!已退数
                .TextMatrix(intRow, mIntCol退药_准退数) = mrsReturnList!准退数
                
                If IIf(IsNull(mrsReturnList!退药数), 0, mrsReturnList!退药数) > 0 And mrsReturnList!执行状态 = mState.退药 Then
                    .TextMatrix(intRow, mIntCol退药_退药数) = mrsReturnList!退药数
                End If
                
                .TextMatrix(intRow, mIntCol退药_单价) = Format(mrsReturnList!单价, "#0." & String(mintPriceDigit, "0"))
        
                .TextMatrix(intRow, mIntCol退药_金额) = Format(mrsReturnList!金额, "#0." & String(mintPriceDigit, "0"))
                .TextMatrix(intRow, mIntCol退药_单量) = mrsReturnList!单量
                .TextMatrix(intRow, mIntCol退药_频次) = IIf(IsNull(mrsReturnList!频次), "", mrsReturnList!频次)
                .TextMatrix(intRow, mIntCol退药_用法) = IIf(IsNull(mrsReturnList!用法), "", mrsReturnList!用法)
                
                .TextMatrix(intRow, mIntCol退药_操作员) = mrsReturnList!操作员
                
                .TextMatrix(intRow, mIntCol退药_发药时间) = mrsReturnList!发药时间
                .TextMatrix(intRow, mIntCol退药_单据) = mrsReturnList!单据
                .TextMatrix(intRow, mIntCol退药_医嘱id) = mrsReturnList!医嘱id
                .TextMatrix(intRow, mIntCol退药_领药人) = IIf(IsNull(mrsReturnList!领药人), "", mrsReturnList!领药人)
                .TextMatrix(intRow, mIntCol退药_领药部门id) = mrsReturnList!领药部门ID
                .TextMatrix(intRow, mIntCol退药_库房货位) = IIf(IsNull(mrsReturnList!库房货位), "", mrsReturnList!库房货位)
               
                .TextMatrix(intRow, mIntCol退药_相关ID) = IIf(IsNull(mrsReturnList!相关ID), 0, mrsReturnList!相关ID)
                .TextMatrix(intRow, mIntCol退药_药品ID) = mrsReturnList!药品ID
                .TextMatrix(intRow, mIntCol退药_单量单位) = mrsReturnList!单量单位
                
                .TextMatrix(intRow, mIntCol退药_药品编码和名称) = mrsReturnList!药品编码和名称
                .TextMatrix(intRow, mIntCol退药_药品编码) = mrsReturnList!药品编码
                .TextMatrix(intRow, mIntCol退药_药品名称) = mrsReturnList!药品名称
                
                .TextMatrix(intRow, mIntCol退药_收发ID) = mrsReturnList!收发ID
                .TextMatrix(intRow, mIntCol退药_状态) = mrsReturnList!状态
                .TextMatrix(intRow, mIntCol退药_执行状态) = mrsReturnList!执行状态
                .TextMatrix(intRow, mIntCol退药_发药号) = mrsReturnList!发药号
                .TextMatrix(intRow, mIntCol退药_发送时间) = zlStr.NVL(mrsReturnList!发送时间)
                .TextMatrix(intRow, mIntCol退药_脚注) = IIf(IsNull(mrsReturnList!医生嘱托), "", mrsReturnList!医生嘱托)
                            
                .TextMatrix(intRow, mIntCol退药_病人ID) = mrsReturnList!病人ID
                .TextMatrix(intRow, mIntCol退药_主页ID) = mrsReturnList!主页ID
                            
                '设置状态的颜色
                If mrsReturnList!执行状态 = mState.退药_原始记录 Then
                    lngStateColor = mListColor.Return_Original
                ElseIf mrsReturnList!执行状态 = mState.退药_发药记录 Then
                    lngStateColor = mListColor.Return_Sended
                ElseIf mrsReturnList!执行状态 = mState.退药_退药记录 Then
                    lngStateColor = mListColor.Return_Returned
                End If
                
                '设置记录的前景色
                .Cell(flexcpForeColor, intRow, 1, intRow, .Cols - 1) = lngStateColor
                
                '设置合理用药标志（PASS）
                If mcondition.intShowPass = 1 Then
                    If mrsReturnList!审查结果 > -1 And mrsReturnList!审查结果 < 5 Then
                        .Cell(flexcpPicture, intRow, mIntCol退药_审查结果, intRow, mIntCol退药_审查结果) = frmPublic.imgPass.ListImages(Val(mrsReturnList!审查结果) + 1).Picture
                        .Cell(flexcpPictureAlignment, intRow, mIntCol退药_审查结果, intRow, mIntCol退药_审查结果) = flexPicAlignCenterCenter
                    End If
                ElseIf mcondition.intShowPass = 3 Then
                    If mrsReturnList!审查结果 = 1 Then
                        .Cell(flexcpPicture, intRow, mIntCol退药_审查结果, intRow, mIntCol退药_审查结果) = frmPublic.imgPass.ListImages(3).Picture
                    ElseIf mrsReturnList!审查结果 = 2 Then
                        .Cell(flexcpPicture, intRow, mIntCol退药_审查结果, intRow, mIntCol退药_审查结果) = frmPublic.imgPass.ListImages(2).Picture
                    ElseIf mrsReturnList!审查结果 = 3 Then
                        .Cell(flexcpPicture, intRow, mIntCol退药_审查结果, intRow, mIntCol退药_审查结果) = frmPublic.imgPass.ListImages(1).Picture
                    End If
                    .Cell(flexcpPictureAlignment, intRow, mIntCol退药_审查结果, intRow, mIntCol退药_审查结果) = flexPicAlignCenterCenter
                End If
                
                '以下用于测试PASS
'                .Cell(flexcpPicture, intRow, mIntCol退药_审查结果, intRow, mIntCol退药_审查结果) = frmPublic.imgPass.ListImages(Val(mrsReturnList!审查结果) + 2).Picture
'                .Cell(flexcpPictureAlignment, intRow, mIntCol退药_审查结果, intRow, mIntCol退药_审查结果) = flexPicAlignCenterCenter
                
                '特殊药品粗体显示
                If IsInString("毒性药;麻醉药;精神I类;精神II类", zlStr.NVL(mrsReturnList!毒理分类), ";") = True And zlStr.NVL(mrsReturnList!毒理分类) <> "" Then
                    .Cell(flexcpFontBold, intRow, mIntCol退药_品名, intRow, mIntCol退药_品名) = True
                End If
                
                mrsReturnList.MoveNext
            Loop
            
            '设置数量列为加粗显示
            .Cell(flexcpFontBold, 1, mIntCol退药_退药数, .rows - 1, mIntCol退药_退药数) = True
        End If
        
        .Redraw = flexRDDirect
    End With
    
    SetGroup vsfList(mListType.退药), mcondition.bln显示过程单据 = False
End Sub
Private Sub RefreshList_Reject()
    '刷新拒发列表
    Dim intRow As Integer
    
    If mrsSendList Is Nothing Then Exit Sub
    
    mrsSendList.Filter = "执行状态=" & mState.拒发 & " Or 执行状态=" & mState.拒发_不处理
    mrsSendList.Sort = "领药部门,姓名,NO,品名"
    
    With vsfList(mListType.拒发)
        .Redraw = flexRDNone
        .rows = 1

        Do While Not mrsSendList.EOF
            intRow = intRow + 1
            .rows = intRow + 1
            
            .TextMatrix(intRow, mIntCol拒发_科室) = mrsSendList!科室
            .TextMatrix(intRow, mIntCol拒发_NO) = mrsSendList!NO
            .TextMatrix(intRow, mIntCol拒发_类型) = mrsSendList!类型
            .TextMatrix(intRow, mIntCol拒发_发药类型) = IIf(mrsSendList!扣率 = 3, "离院带药", IIf(mrsSendList!扣率 = 4, "自取药", "院内用药"))
            .TextMatrix(intRow, mIntCol拒发_床号) = IIf(IsNull(mrsSendList!床号), "", mrsSendList!床号)
            .TextMatrix(intRow, mIntCol拒发_姓名) = mrsSendList!姓名
            .TextMatrix(intRow, mIntCol拒发_性别) = mrsSendList!性别
            
            If mcondition.int药品名称编码显示 = 0 Then
                .TextMatrix(intRow, mIntCol拒发_品名) = mrsSendList!品名
            ElseIf mcondition.int药品名称编码显示 = 1 Then
                .TextMatrix(intRow, mIntCol拒发_品名) = mrsSendList!药品编码
            Else
                .TextMatrix(intRow, mIntCol拒发_品名) = mrsSendList!药品名称
            End If
            
            If mrsSendList!高危药品 > 0 Then
                .Cell(flexcpPicture, intRow, mIntCol拒发_品名) = Me.ImgList.ListImages("高危").Picture
                .Cell(flexcpPictureAlignment, intRow, mIntCol拒发_品名) = flexPicAlignLeftCenter
            End If
            
            .TextMatrix(intRow, mIntCol拒发_规格) = IIf(IsNull(mrsSendList!规格), "", mrsSendList!规格)
            .TextMatrix(intRow, mIntCol拒发_生产商) = IIf(IsNull(mrsSendList!产地), "", mrsSendList!产地)
            .TextMatrix(intRow, mIntCol拒发_原产地) = IIf(IsNull(mrsSendList!原产地), "", mrsSendList!原产地)
            .TextMatrix(intRow, mIntCol拒发_批号) = IIf(IsNull(mrsSendList!批号), "", mrsSendList!批号)
            .TextMatrix(intRow, mIntCol拒发_效期) = IIf(IsNull(mrsSendList!效期), "", mrsSendList!效期)
            .TextMatrix(intRow, mIntCol拒发_数量) = mrsSendList!数量
            
            .TextMatrix(intRow, mIntCol拒发_单价) = Format(mrsSendList!单价, "#0." & String(mintPriceDigit, "0"))
            .TextMatrix(intRow, mIntCol拒发_金额) = Format(mrsSendList!金额, "#0." & String(mintPriceDigit, "0"))
            
            .TextMatrix(intRow, mIntCol拒发_药品编码和名称) = mrsSendList!药品编码和名称
            .TextMatrix(intRow, mIntCol拒发_药品编码) = mrsSendList!药品编码
            .TextMatrix(intRow, mIntCol拒发_药品名称) = mrsSendList!药品名称
            
            .TextMatrix(intRow, mIntCol拒发_执行状态) = mrsSendList!执行状态
            
            If mrsSendList!执行状态 = mState.拒发 Then
                .TextMatrix(intRow, mIntCol拒发_状态) = ""
            Else
                .TextMatrix(intRow, mIntCol拒发_状态) = mrsSendList!状态
            End If
            
            .TextMatrix(intRow, mIntCol拒发_收发ID) = mrsSendList!收发ID
            .TextMatrix(intRow, mIntCol拒发_脚注) = NVL(mrsSendList!医生嘱托, "")
            
            '特殊药品粗体显示
            If IsInString("毒性药;麻醉药;精神I类;精神II类", zlStr.NVL(mrsSendList!毒理分类), ";") = True And zlStr.NVL(mrsSendList!毒理分类) <> "" Then
                .Cell(flexcpFontBold, intRow, mIntCol拒发_品名, intRow, mIntCol拒发_品名) = True
            End If
            
            mrsSendList.MoveNext
        Loop
        
        .Redraw = flexRDDirect
   End With
End Sub

Private Sub RefreshList_Shortage()
    '刷新缺药列表
    
    Dim intRow As Integer
    
    If mrsSendList Is Nothing Then Exit Sub
    
    mrsSendList.Filter = "执行状态=" & mState.缺药
    mrsSendList.Sort = "领药部门,姓名,NO,品名"
    
    With vsfList(mListType.缺药)
        .Redraw = flexRDNone
        .rows = 1

        Do While Not mrsSendList.EOF
            intRow = intRow + 1
            .rows = intRow + 1
            
            .TextMatrix(intRow, mIntCol缺药_科室) = mrsSendList!科室
            .TextMatrix(intRow, mIntCol缺药_NO) = mrsSendList!NO
            .TextMatrix(intRow, mIntCol缺药_类型) = mrsSendList!类型
            .TextMatrix(intRow, mIntCol缺药_发药类型) = IIf(mrsSendList!扣率 = 3, "离院带药", IIf(mrsSendList!扣率 = 4, "自取药", "院内用药"))
            .TextMatrix(intRow, mIntCol缺药_床号) = IIf(IsNull(mrsSendList!床号), "", mrsSendList!床号)
            .TextMatrix(intRow, mIntCol缺药_姓名) = mrsSendList!姓名
            .TextMatrix(intRow, mIntCol缺药_性别) = mrsSendList!性别
            
            If mcondition.int药品名称编码显示 = 0 Then
                .TextMatrix(intRow, mIntCol缺药_品名) = mrsSendList!品名
            ElseIf mcondition.int药品名称编码显示 = 1 Then
                .TextMatrix(intRow, mIntCol缺药_品名) = mrsSendList!药品编码
            Else
                .TextMatrix(intRow, mIntCol缺药_品名) = mrsSendList!药品名称
            End If
            
            If mrsSendList!抗生素 <> 0 Then
                .Cell(flexcpPicture, intRow, mIntCol缺药_品名) = Me.ImgList.ListImages(39).Picture
                .Cell(flexcpPictureAlignment, intRow, mIntCol缺药_品名) = flexPicAlignLeftCenter
            End If
            
            If mrsSendList!高危药品 > 0 Then
                .Cell(flexcpPicture, intRow, mIntCol缺药_品名) = Me.ImgList.ListImages("高危").Picture
                .Cell(flexcpPictureAlignment, intRow, mIntCol缺药_品名) = flexPicAlignLeftCenter
            End If
                        
            .TextMatrix(intRow, mIntCol缺药_规格) = IIf(IsNull(mrsSendList!规格), "", mrsSendList!规格)
            .TextMatrix(intRow, mIntCol缺药_生产商) = IIf(IsNull(mrsSendList!产地), "", mrsSendList!产地)
            .TextMatrix(intRow, mIntCol缺药_原产地) = IIf(IsNull(mrsSendList!原产地), "", mrsSendList!原产地)
            .TextMatrix(intRow, mIntCol缺药_批号) = IIf(IsNull(mrsSendList!批号), "", mrsSendList!批号)
            .TextMatrix(intRow, mIntCol缺药_效期) = IIf(IsNull(mrsSendList!效期), "", mrsSendList!效期)
            .TextMatrix(intRow, mIntCol缺药_数量) = mrsSendList!数量
            
            .TextMatrix(intRow, mIntCol缺药_单价) = Format(mrsSendList!单价, "#0." & String(mintPriceDigit, "0"))
            .TextMatrix(intRow, mIntCol缺药_金额) = Format(mrsSendList!金额, "#0." & String(mintPriceDigit, "0"))
            
            .TextMatrix(intRow, mIntCol缺药_药品编码和名称) = mrsSendList!药品编码和名称
            .TextMatrix(intRow, mIntCol缺药_药品编码) = mrsSendList!药品编码
            .TextMatrix(intRow, mIntCol缺药_药品名称) = mrsSendList!药品名称
            .TextMatrix(intRow, mIntCol缺药_脚注) = NVL(mrsSendList!医生嘱托, "")
            
            '特殊药品粗体显示
            If IsInString("毒性药;麻醉药;精神I类;精神II类", zlStr.NVL(mrsSendList!毒理分类), ";") = True And zlStr.NVL(mrsSendList!毒理分类) <> "" Then
                .Cell(flexcpFontBold, intRow, mIntCol缺药_品名, intRow, mIntCol缺药_品名) = True
            End If
            
            mrsSendList.MoveNext
        Loop
        
        .Redraw = flexRDDirect
   End With
End Sub

Private Function RefreshList_ChargeOff(ByVal lng科室ID As Long, ByVal lng药品id As Long) As Boolean
    '刷新销账列表
    Dim intRow As Integer
    Dim dblSumNum As Double
    
    If mrsChargeOff Is Nothing Then Exit Function
    
    mrsChargeOff.Filter = "领药部门id=" & lng科室ID & " And 药品ID=" & lng药品id & " And 审核标志 = 1"
    mrsChargeOff.Sort = "NO,收发序号 Desc"
    
    If mrsChargeOff.EOF Then Exit Function
    
    With vsfChargeOff
        .Redraw = flexRDNone
        .rows = 1

        Do While Not mrsChargeOff.EOF
            intRow = intRow + 1
            .rows = intRow + 1
            
            .TextMatrix(intRow, mIntCol销账_申请科室) = mrsChargeOff!领药部门
            .TextMatrix(intRow, mIntCol销账_单据) = mrsChargeOff!单据
            .TextMatrix(intRow, mIntCol销账_NO) = mrsChargeOff!NO
            .TextMatrix(intRow, mIntCol销账_药品ID) = mrsChargeOff!药品ID
            .TextMatrix(intRow, mIntCol销账_申请时间) = Format(mrsChargeOff!申请时间, "yyyy-mm-dd hh:mm:ss")
            .TextMatrix(intRow, mIntCol销账_生产商) = IIf(IsNull(mrsChargeOff!产地), "", mrsChargeOff!产地)
            .TextMatrix(intRow, mIntCol销账_批号) = IIf(IsNull(mrsChargeOff!批号), "", mrsChargeOff!批号)
            .TextMatrix(intRow, mIntCol销账_效期) = Format(mrsChargeOff!效期, "yyyy-mm-dd")
            .TextMatrix(intRow, mIntCol销账_准退数量) = zlStr.FormatEx(mrsChargeOff!准退数量 / mrsChargeOff!包装, 5)
            .TextMatrix(intRow, mIntCol销账_销帐数量) = zlStr.FormatEx(mrsChargeOff!销帐数量 / mrsChargeOff!包装, 5)
            .TextMatrix(intRow, mIntCol销账_包装) = IIf(IsNull(mrsChargeOff!包装), "", mrsChargeOff!包装)
            .TextMatrix(intRow, mIntCol销账_单位) = IIf(IsNull(mrsChargeOff!单位), "", mrsChargeOff!单位)
            .TextMatrix(intRow, mIntCol销账_收发序号) = IIf(IsNull(mrsChargeOff!收发序号), "", mrsChargeOff!收发序号)
            
            dblSumNum = dblSumNum + mrsChargeOff!销帐数量 / mrsChargeOff!包装
            
           mrsChargeOff.MoveNext
        Loop
        
        intRow = intRow + 1
        .rows = intRow + 1
            
        .TextMatrix(intRow, mIntCol销账_NO) = "合计"
        .TextMatrix(intRow, mIntCol销账_销帐数量) = zlStr.FormatEx(dblSumNum, 5)
        
        .Redraw = flexRDDirect
   End With
   
   RefreshList_ChargeOff = True
End Function


Private Sub RefreshList_Sum()
    '刷新汇总列表
    
    Dim str科室汇总 As String
    Dim str药品汇总 As String
    Dim dblSumNumber As Double
    Dim dblSumMoney As Double
    Dim intRow As Integer
    Dim strSum As String
    Dim intSumType As Integer
    Dim strFilter As String
    Dim n As Integer
    
    If mrsSendList Is Nothing Then Exit Sub
    
    strFilter = "执行状态=" & mState.发药
    
    '是否显示退药待发药品
    If mcondition.bln显示退药待发单据 = False Then
        strFilter = strFilter & " And 记录状态=1 "
    End If

    mrsSendList.Filter = strFilter
    
    With vsfList(mListType.汇总)
        .Redraw = flexRDNone
        .rows = 1
        .MergeCells = flexMergeNever
        .Subtotal flexSTClear
        
        If mrsSendList.EOF Then
            .rows = 2
            .Cell(flexcpText, 1, IIf(mcondition.bln按科室汇总 = True, mIntCol科室汇总_科室, mIntCol汇总_品名), 1, .Cols - 1) = "没有找到满足条件的记录......"
            .MergeCells = flexMergeRestrictRows
            .MergeRow(1) = True
        Else
            '按科室、药品汇总
            If mcondition.bln按科室汇总 = True Then
                mrsSendList.Sort = "领药部门,品名,批次"
'                 If mcondition.bln按批次汇总 = True Then
'                    .ColWidth(mIntCol科室汇总_批号) = 1200
                Do While Not mrsSendList.EOF
                    Debug.Print mrsSendList!领药部门 & mrsSendList!品名 & IIf(mcondition.bln按批次汇总, IIf(IsNull(mrsSendList!批次), 0, mrsSendList!批次), "")
                    If str科室汇总 <> mrsSendList!领药部门 & mrsSendList!品名 & IIf(mcondition.bln按批次汇总, IIf(IsNull(mrsSendList!批次), 0, mrsSendList!批次), "") Then
                        intRow = intRow + 1
                        .rows = intRow + 1
                        
                        str科室汇总 = mrsSendList!领药部门 & mrsSendList!品名 & IIf(mcondition.bln按批次汇总, IIf(IsNull(mrsSendList!批次), 0, mrsSendList!批次), "")
                        dblSumNumber = mrsSendList!实际数量
                        dblSumMoney = Val(mrsSendList!金额)
                        
                        .TextMatrix(intRow, mIntCol科室汇总_当前行) = ""
                        
                        .TextMatrix(intRow, mIntCol科室汇总_科室) = mrsSendList!科室
                        If mcondition.int药品名称编码显示 = 0 Then
                            .TextMatrix(intRow, mIntCol科室汇总_品名) = mrsSendList!品名
                        ElseIf mcondition.int药品名称编码显示 = 1 Then
                            .TextMatrix(intRow, mIntCol科室汇总_品名) = mrsSendList!药品编码
                        Else
                            .TextMatrix(intRow, mIntCol科室汇总_品名) = mrsSendList!药品名称
                        End If
                        
                        If mrsSendList!抗生素 <> 0 Then
                            .Cell(flexcpPicture, intRow, mIntCol科室汇总_品名) = Me.ImgList.ListImages(39).Picture
                            .Cell(flexcpPictureAlignment, intRow, mIntCol科室汇总_品名) = flexPicAlignLeftCenter
                        End If
                        
                        If mrsSendList!高危药品 > 0 Then
                            .Cell(flexcpPicture, intRow, mIntCol科室汇总_品名) = Me.ImgList.ListImages("高危").Picture
                            .Cell(flexcpPictureAlignment, intRow, mIntCol科室汇总_品名) = flexPicAlignLeftCenter
                        End If
                
                        .TextMatrix(intRow, mIntCol科室汇总_规格) = IIf(IsNull(mrsSendList!规格), "", mrsSendList!规格)
                        .TextMatrix(intRow, mIntCol科室汇总_生产商) = IIf(IsNull(mrsSendList!产地), "", mrsSendList!产地)
                        .TextMatrix(intRow, mIntCol科室汇总_原产地) = IIf(IsNull(mrsSendList!原产地), "", mrsSendList!原产地)
                        .TextMatrix(intRow, mIntCol科室汇总_批号) = IIf(IsNull(mrsSendList!批号), "", mrsSendList!批号)
                        .TextMatrix(intRow, mIntCol科室汇总_效期) = IIf(IsNull(mrsSendList!效期), "", mrsSendList!效期)
                        
                        .TextMatrix(intRow, mIntCol科室汇总_应发数量) = mrsSendList!实际数量
                        .TextMatrix(intRow, mIntCol科室汇总_留存数量) = zlStr.FormatEx(mrsSendList!留存数量, 5)
                        If mcondition.bln启用退药销账 = True Then
                            .TextMatrix(intRow, mIntCol科室汇总_销帐数量) = FormatEx(GetChargeOffCount(mrsSendList!领药部门ID, mrsSendList!药品ID, mrsSendList!批次), 5)
                        Else
                            .TextMatrix(intRow, mIntCol科室汇总_销帐数量) = "0"
                        End If
                        
                        .TextMatrix(intRow, mIntCol科室汇总_实发数量) = mrsSendList!实际数量
                        .TextMatrix(intRow, mIntCol科室汇总_单位) = mrsSendList!单位
                        
                        .TextMatrix(intRow, mIntCol科室汇总_单价) = Format(mrsSendList!单价, "#0." & String(mintPriceDigit, "0"))
                        .TextMatrix(intRow, mIntCol科室汇总_应发金额) = Format(mrsSendList!金额, "#0." & String(mintPriceDigit, "0"))
                        
                        .TextMatrix(intRow, mIntCol科室汇总_批次) = IIf(IsNull(mrsSendList!批次), 0, mrsSendList!批次)
                        .TextMatrix(intRow, mIntCol科室汇总_科室ID) = mrsSendList!科室ID
                        .TextMatrix(intRow, mIntCol科室汇总_药品ID) = mrsSendList!药品ID
                        
                        .TextMatrix(intRow, mIntCol科室汇总_领药部门) = mrsSendList!领药部门
                        .TextMatrix(intRow, mIntCol科室汇总_领药部门id) = mrsSendList!领药部门ID
                        
                        .TextMatrix(intRow, mIntCol科室汇总_药品编码和名称) = mrsSendList!品名
                        .TextMatrix(intRow, mIntCol科室汇总_药品编码) = mrsSendList!药品编码
                        .TextMatrix(intRow, mIntCol科室汇总_药品名称) = mrsSendList!药品名称
                        
                        .TextMatrix(intRow, mIntCol科室汇总_包装) = mrsSendList!包装
                        
                        '如果存在上一行（上一行不是固定行时），格式化上一行的数量、金额
                        If intRow - 1 > 0 Then
                            .TextMatrix(intRow - 1, mIntCol科室汇总_应发数量) = zlStr.FormatEx(Val(.TextMatrix(intRow - 1, mIntCol科室汇总_应发数量)), 5)
                            .TextMatrix(intRow - 1, mIntCol科室汇总_留存数量) = zlStr.FormatEx(Val(.TextMatrix(intRow - 1, mIntCol科室汇总_留存数量)), 5)
                            .TextMatrix(intRow - 1, mIntCol科室汇总_销帐数量) = zlStr.FormatEx(Val(.TextMatrix(intRow - 1, mIntCol科室汇总_销帐数量)), 5)
                            .TextMatrix(intRow - 1, mIntCol科室汇总_实发数量) = zlStr.FormatEx(Val(.TextMatrix(intRow - 1, mIntCol科室汇总_实发数量)), 5)
                            .TextMatrix(intRow - 1, mIntCol科室汇总_应发金额) = zlStr.FormatEx(Val(.TextMatrix(intRow - 1, mIntCol科室汇总_应发金额)), mintMoneyDigit, , True)
                            .TextMatrix(intRow - 1, mIntCol科室汇总_实发金额) = zlStr.FormatEx(Val(.TextMatrix(intRow - 1, mIntCol科室汇总_实发金额)), mintMoneyDigit, , True)
                        End If
                        
                        '特殊药品粗体显示
                        If IsInString("毒性药;麻醉药;精神I类;精神II类", zlStr.NVL(mrsSendList!毒理分类), ";") = True And zlStr.NVL(mrsSendList!毒理分类) <> "" Then
                            .Cell(flexcpFontBold, intRow, mIntCol科室汇总_品名, intRow, mIntCol科室汇总_品名) = True
                        End If
                    Else
                        dblSumNumber = dblSumNumber + mrsSendList!实际数量
                        dblSumMoney = dblSumMoney + Val(mrsSendList!金额)
                        
                        .TextMatrix(intRow, mIntCol科室汇总_应发数量) = dblSumNumber
                        .TextMatrix(intRow, mIntCol科室汇总_实发数量) = dblSumNumber
                        .TextMatrix(intRow, mIntCol科室汇总_应发金额) = zlStr.FormatEx(dblSumMoney, mintMoneyDigit, , True)
                        
                        
                    End If
                    
                    mrsSendList.MoveNext
                Loop
                
                '设置数量列为加粗显示
                .Cell(flexcpFontBold, 1, mIntCol科室汇总_实发数量, .rows - 1, mIntCol科室汇总_实发数量) = True
                
                '统计实际发药数量
                For n = 1 To .rows - 1
                    If .TextMatrix(n, 0) <> "小计" Then
                        '应发数量小于了销帐数量，实发为负数（表示科室将实物退药），留存数为0
                        If Val(.TextMatrix(n, mIntCol科室汇总_应发数量)) - Val(.TextMatrix(n, mIntCol科室汇总_销帐数量)) < 0 Then
                            .TextMatrix(n, mIntCol科室汇总_实发数量) = zlStr.FormatEx(Val(.TextMatrix(n, mIntCol科室汇总_应发数量)) - Val(.TextMatrix(n, mIntCol科室汇总_销帐数量)), 5)
                            .TextMatrix(n, mIntCol科室汇总_留存数量) = 0
                        Else
                            If Val(.TextMatrix(n, mIntCol科室汇总_留存数量)) > 0 Then
                                '如果留存数量不为0（从药品留存计划取值），根据实际应发数量计算（实际应发＝应发数量－销帐数量）
                                If Val(.TextMatrix(n, mIntCol科室汇总_留存数量)) > Val(.TextMatrix(n, mIntCol科室汇总_应发数量)) - Val(.TextMatrix(n, mIntCol科室汇总_销帐数量)) Then
                                    '留存数量大于了实际应发数量，则留存数量＝实际应发数量
                                    .TextMatrix(n, mIntCol科室汇总_留存数量) = zlStr.FormatEx(Val(.TextMatrix(n, mIntCol科室汇总_应发数量)) - Val(.TextMatrix(n, mIntCol科室汇总_销帐数量)), 5)
                                End If
                                
                                '实发数量＝应发数量－留存数量－销帐数量
                                .TextMatrix(n, mIntCol科室汇总_实发数量) = zlStr.FormatEx(Val(.TextMatrix(n, mIntCol科室汇总_应发数量) - Val(.TextMatrix(n, mIntCol科室汇总_留存数量)) - Val(.TextMatrix(n, mIntCol科室汇总_销帐数量))), 5)
                            ElseIf Val(.TextMatrix(n, mIntCol科室汇总_留存数量)) = 0 Then
                                .TextMatrix(n, mIntCol科室汇总_实发数量) = FormatEx(Val(.TextMatrix(n, mIntCol科室汇总_应发数量) - Val(.TextMatrix(n, mIntCol科室汇总_销帐数量))), 5)
                            End If
                        End If
                        
                        .TextMatrix(n, mIntCol科室汇总_实发金额) = zlStr.FormatEx(Val(.TextMatrix(n, mIntCol科室汇总_应发金额)) / Val(.TextMatrix(n, mIntCol科室汇总_应发数量)) * Val(.TextMatrix(n, mIntCol科室汇总_实发数量)), mintMoneyDigit, , True)
                        
                        .Row = n
                        .Col = mIntCol科室汇总_实发数量
                        .CellFontBold = True
                        If Val(.TextMatrix(n, mIntCol科室汇总_实发数量)) < 0 Then
                            .CellForeColor = vbRed
                        ElseIf Val(.TextMatrix(n, mIntCol科室汇总_实发数量)) > 0 Then
                            .CellForeColor = vbBlue
                        End If
                    End If
                Next
               
                '按批次汇总（显示批号列）
                If mcondition.bln按批次汇总 = True Then
                    .ColWidth(mIntCol科室汇总_批号) = 1200
                Else
                    .ColWidth(mIntCol科室汇总_批号) = 0
                End If
                
                '设置小计，合计
                SetSubTotal vsfList(mListType.汇总), "领药部门"
            Else
            '按药品汇总
                mrsSendList.Sort = "品名"
                If mrsSendList.EOF Then mrsSendList.MoveFirst
                
                Do While Not mrsSendList.EOF
                    If str药品汇总 <> mrsSendList!品名 & IIf(mcondition.bln按批次汇总, IIf(IsNull(mrsSendList!批次), 0, mrsSendList!批次), "") Then
                        intRow = intRow + 1
                        .rows = intRow + 1
                        
                        str药品汇总 = mrsSendList!品名 & IIf(mcondition.bln按批次汇总, IIf(IsNull(mrsSendList!批次), 0, mrsSendList!批次), "")
                        dblSumNumber = mrsSendList!实际数量
                        dblSumMoney = Val(mrsSendList!金额)
                        
                        .TextMatrix(intRow, mIntCol汇总_当前行) = ""
                        
                        If mcondition.int药品名称编码显示 = 0 Then
                            .TextMatrix(intRow, mIntCol汇总_品名) = mrsSendList!品名
                        ElseIf mcondition.int药品名称编码显示 = 1 Then
                            .TextMatrix(intRow, mIntCol汇总_品名) = mrsSendList!药品编码
                        Else
                            .TextMatrix(intRow, mIntCol汇总_品名) = mrsSendList!药品名称
                        End If
                        
                        If mrsSendList!抗生素 <> 0 Then
                            .Cell(flexcpPicture, intRow, mIntCol汇总_品名) = Me.ImgList.ListImages(39).Picture
                            .Cell(flexcpPictureAlignment, intRow, mIntCol汇总_品名) = flexPicAlignLeftCenter
                        End If
                        
                        If mrsSendList!高危药品 > 0 Then
                            .Cell(flexcpPicture, intRow, mIntCol汇总_品名) = Me.ImgList.ListImages("高危").Picture
                            .Cell(flexcpPictureAlignment, intRow, mIntCol汇总_品名) = flexPicAlignLeftCenter
                        End If
                        
                        .TextMatrix(intRow, mIntCol汇总_规格) = IIf(IsNull(mrsSendList!规格), "", mrsSendList!规格)
                        .TextMatrix(intRow, mIntCol汇总_生产商) = IIf(IsNull(mrsSendList!产地), "", mrsSendList!产地)
                        .TextMatrix(intRow, mIntCol汇总_原产地) = IIf(IsNull(mrsSendList!原产地), "", mrsSendList!原产地)
                        .TextMatrix(intRow, mIntCol汇总_批号) = IIf(IsNull(mrsSendList!批号), "", mrsSendList!批号)
                        .TextMatrix(intRow, mIntCol汇总_效期) = IIf(IsNull(mrsSendList!效期), "", mrsSendList!效期)
                        .TextMatrix(intRow, mIntCol汇总_数量) = mrsSendList!实际数量
                        .TextMatrix(intRow, mIntCol汇总_单位) = mrsSendList!单位
                        .TextMatrix(intRow, mIntCol汇总_单价) = Format(mrsSendList!单价, "#0." & String(mintPriceDigit, "0"))
                        .TextMatrix(intRow, mIntCol汇总_金额) = Format(mrsSendList!金额, "#0." & String(mintPriceDigit, "0"))
                        
                        .TextMatrix(intRow, mIntCol汇总_药品编码和名称) = mrsSendList!品名
                        .TextMatrix(intRow, mIntCol汇总_药品编码) = mrsSendList!药品编码
                        .TextMatrix(intRow, mIntCol汇总_药品名称) = mrsSendList!药品名称
                        
                        '如果存在上一行（上一行不是固定行时），格式化上一行的数量、金额
                        If intRow - 1 > 0 Then
                            .TextMatrix(intRow - 1, mIntCol汇总_数量) = zlStr.FormatEx(Val(.TextMatrix(intRow - 1, mIntCol汇总_数量)), 5)
                            .TextMatrix(intRow - 1, mIntCol汇总_金额) = zlStr.FormatEx(Val(.TextMatrix(intRow - 1, mIntCol汇总_金额)), mintMoneyDigit, , True)
                        End If
                        
                        '特殊药品粗体显示
                        If IsInString("毒性药;麻醉药;精神I类;精神II类", zlStr.NVL(mrsSendList!毒理分类), ";") = True And zlStr.NVL(mrsSendList!毒理分类) <> "" Then
                            .Cell(flexcpFontBold, intRow, mIntCol汇总_品名, intRow, mIntCol汇总_品名) = True
                        End If
                    Else
                        dblSumNumber = dblSumNumber + mrsSendList!实际数量
                        dblSumMoney = dblSumMoney + Val(mrsSendList!金额)
                        
                        .TextMatrix(intRow, mIntCol汇总_数量) = dblSumNumber
                        .TextMatrix(intRow, mIntCol汇总_金额) = zlStr.FormatEx(dblSumMoney, mintMoneyDigit, , True)
                    End If
                    
                    mrsSendList.MoveNext
                Loop
                
                '设置数量列为加粗显示
                .Cell(flexcpFontBold, 1, mIntCol汇总_数量, .rows - 1, mIntCol汇总_数量) = True
                
                '按批次汇总（显示批号列）
                If mcondition.bln按批次汇总 = True Then
                    .ColWidth(mIntCol汇总_批号) = 1200
                Else
                    .ColWidth(mIntCol汇总_批号) = 0
                End If
                
                '设置小计，合计
                SetSubTotal vsfList(mListType.汇总), ""
            End If
            .Row = 1
            vsfChargeOff_EnterCell
        End If
        
        .Redraw = flexRDDirect
    End With
End Sub


Private Sub ResizeChargeOffList()
    On Error Resume Next
    
    vsfList(mListType.汇总).Height = mdblSumListHeight

    If vsfChargeOff.Visible = True Then
        vsfList(mListType.汇总).Height = mdblSumListHeight / 4 * 3
                        
        picHsc.Top = vsfList(mListType.汇总).Top + vsfList(mListType.汇总).Height
        picHsc.Left = vsfList(mListType.汇总).Left
        picHsc.Width = vsfList(mListType.汇总).Width
        
        vsfChargeOff.Top = picHsc.Top + picHsc.Height
        vsfChargeOff.Left = vsfList(mListType.汇总).Left
        vsfChargeOff.Height = mdblSumListHeight / 4 - picHsc.Height
        vsfChargeOff.Width = vsfList(mListType.汇总).Width
    End If
End Sub

Public Sub SetParams()
    Dim intType As Integer
    
    With mcondition
        .lng药房id = Val(zlDataBase.GetPara("发药药房", glngSys, 1342))
        .bln按科室汇总 = (Val(zlDataBase.GetPara("按科室汇总显示汇总清单", glngSys, 1342)) = 1)
        .bln启用退药销账 = (Val(zlDataBase.GetPara("发药时汇总退药销帐记录", glngSys, 1342, 0)) = 1) And IsInString(gstrprivs, "退药销帐", ";")
        .bln允许未审核处方发药 = (gtype_UserSysParms.P6_未审核记帐处方发药 = 1)
        .int药品名称编码显示 = Int(Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品部门发药管理", "药品名称显示方式", 0)))
        .int发药单格式 = Int(Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品部门发药管理", "发药单格式", 0)))
        .bln配制中心 = CheckIsCenter(.lng药房id)
        .str高危发放 = zlDataBase.GetPara("高危药品发放", glngSys, 1342, "")
        .str高危分类 = zlDataBase.GetPara("高危分类", glngSys, 1342, "")
        .int退药待发单据默认为发药状态 = Val(zlDataBase.GetPara("退药待发单据默认为发药状态", glngSys, 1342, ""))
        
        .bln显示原产地 = Is中药库房(.lng药房id)
        
        If .int药品名称编码显示 > 2 Or .int药品名称编码显示 < 0 Then .int药品名称编码显示 = 0
        If .bln药品储备 <> (Val(zlDataBase.GetPara("库房货位及库存限量提示", glngSys, 1342, 0)) = 1) Then
            .bln药品储备 = (Val(zlDataBase.GetPara("库房货位及库存限量提示", glngSys, 1342, 0)) = 1)
            
            intType = mcondition.intListType
            
            If intType = mListType.发药 Or intType = mListType.退药 Then
                If .bln药品储备 = True Then
                    mstrUnallowShow(intType) = Replace(mstrUnallowShow(intType), ";库房货位", "")
                    vsfList(intType).ColWidth(mIntCol发药_库房货位) = 1200
                Else
                    mstrUnallowShow(intType) = mstrUnallowShow(intType) & ";库房货位"
                    vsfList(intType).ColWidth(mIntCol发药_库房货位) = 0
                End If
                
                InitColSelList intType
            End If
        End If
        Call GetDrugDigit(.lng药房id, "药品处方发药", mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
    End With
    
    
End Sub

Private Sub SetPassMenuButton(ByVal intListType As Integer, ByVal lngRow As Long)
    '设置cmdAlley按钮状态
    Dim cbrControl As CommandBarControl
    Dim rsData As ADODB.Recordset
    
    If mcondition.intShowPass <> 1 Or Not IsInString(gstrprivs, "合理用药监测", ";") Then Exit Sub
    
    '判断是住院还是门诊病人，如果没有找到记录（无医嘱）就不显示cmdAlley按钮
    On Error GoTo errHandle
    gstrSQL = "Select distinct B.病人id,nvl(B.主页id,0) 主页id,nvl(C.挂号单,'') 挂号单 " & _
        " From 药品收发记录 A,住院费用记录 B,病人医嘱记录 C " & _
        " Where A.费用id=B.Id And b.医嘱序号=c.Id And nvl(B.医嘱序号,0)<>0 And C.诊疗类别 IN('5','6','7')" & _
        " And A.单据=[2] And A.no=[1] "
    Set rsData = zlDataBase.OpenSQLRecord(gstrSQL, Me.Caption, vsfList(intListType).TextMatrix(lngRow, vsfList(intListType).ColIndex("NO")), Val(vsfList(intListType).TextMatrix(lngRow, vsfList(intListType).ColIndex("单据"))))
    
    Set cbrControl = Me.cbsMain(2).Controls.Find(xtpControlButton, conMenu_Tool_ShowPlug, , True)
    
    If rsData.RecordCount = 0 Then
        If Not cbrControl Is Nothing Then cbrControl.Enabled = False
    Else
        If Not cbrControl Is Nothing Then cbrControl.Enabled = True
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetSendBillState(ByVal intChangeType As Integer, Optional ByVal intRow As Integer = 0)
    Dim intState As Integer
    Dim strState As String
    Dim lngColor As Long
    Dim i As Long
    Dim lng相关ID As Long
    Dim strNo As String
    
    If (intChangeType = mChangeState.发药 Or intChangeType = mChangeState.拒发 Or intChangeType = mChangeState.缺药 _
        Or intChangeType = mChangeState.不处理) And intRow = 0 Then Exit Sub
    
    With vsfList(mListType.发药)
        '单条记录或所选的多条记录改变状态
        If intChangeType = mChangeState.发药 Or intChangeType = mChangeState.拒发 Or intChangeType = mChangeState.不处理 Then
            Select Case intChangeType
                Case mChangeState.发药
                    intState = mState.发药
                    strState = "发药"
                    lngColor = mListColor.State_Send
                Case mChangeState.拒发
                    intState = mState.拒发
                    strState = "拒发"
                    lngColor = mListColor.State_Reject
                Case mChangeState.不处理
                    intState = mState.不处理
                    strState = "不处理"
                    lngColor = mListColor.State_UnProcess
            End Select
            
            For i = 1 To .rows - 1
                If .IsSubtotal(i) = False And .IsSelected(i) = True And Val(.TextMatrix(i, mIntCol发药_执行状态)) <> intState And Val(.TextMatrix(i, mIntCol发药_执行状态)) <> mState.缺药 Then
                    .TextMatrix(i, mIntCol发药_执行状态) = intState
                    .TextMatrix(i, mIntCol发药_状态) = strState
                    
                    .Cell(flexcpBackColor, i, 1, i, .Cols - 1) = lngColor
                    
                    mrsSendList.Filter = "收发ID=" & Val(.TextMatrix(i, mIntCol发药_收发ID))
                    
                    mrsSendList!执行状态 = intState
                    mrsSendList!状态 = strState
                    
                    mrsSendList.Update
                    
                    mblnSendChange = True
                End If
            Next
            
            '同组医嘱的药品状态都要同步改变；如果是高危药品，并且要求单独发放时不同步改变
            If mcondition.bln配制中心 = True And InStr(1, mcondition.str高危发放, .TextMatrix(intRow, mIntCol发药_高危药品)) = 0 Then
                strNo = .TextMatrix(intRow, mIntCol发药_NO)
                lng相关ID = Val(.TextMatrix(intRow, mIntCol发药_相关ID))
                If lng相关ID > 0 Then
                    For i = 1 To .rows - 1
                        If .IsSubtotal(i) = False And .TextMatrix(i, mIntCol发药_NO) = strNo And Val(.TextMatrix(i, mIntCol发药_相关ID)) = lng相关ID _
                            And Val(.TextMatrix(i, mIntCol发药_执行状态)) <> intState And Val(.TextMatrix(i, mIntCol发药_执行状态)) <> mState.缺药 _
                            And InStr(1, mcondition.str高危发放, .TextMatrix(i, mIntCol发药_高危药品)) = 0 Then
                            .TextMatrix(i, mIntCol发药_执行状态) = intState
                            .TextMatrix(i, mIntCol发药_状态) = strState
                            
                            .Cell(flexcpBackColor, i, 1, i, .Cols - 1) = lngColor
                            
                            mrsSendList.Filter = "收发ID=" & Val(.TextMatrix(i, mIntCol发药_收发ID))
                            
                            mrsSendList!执行状态 = intState
                            mrsSendList!状态 = strState
                            
                            mrsSendList.Update
                            
                            mblnSendChange = True
                        End If
                    Next
                End If
            End If
            
            SetMainComandBars mListType.发药, intRow
        Else
        '所有待发药界面的记录改变状态
            Select Case intChangeType
                Case mChangeState.全部发药
                    intState = mState.发药
                    strState = "发药"
                    lngColor = mListColor.State_Send
                Case mChangeState.全部拒发
                    intState = mState.拒发
                    strState = "拒发"
                    lngColor = mListColor.State_Reject
                Case mChangeState.全部不处理
                    intState = mState.不处理
                    strState = "不处理"
                    lngColor = mListColor.State_UnProcess
            End Select
            
            For i = 1 To .rows - 1
                If .IsSubtotal(i) = False And Val(.TextMatrix(i, mIntCol发药_执行状态)) <> mState.缺药 Then
                    .TextMatrix(i, mIntCol发药_执行状态) = intState
                    .TextMatrix(i, mIntCol发药_状态) = strState
                
                    .Cell(flexcpBackColor, i, 1, i, .Cols - 1) = lngColor
                    
                    mrsSendList.Filter = "收发ID=" & Val(.TextMatrix(i, mIntCol发药_收发ID))
                    
                    mrsSendList!执行状态 = intState
                    mrsSendList!状态 = strState
                    
                    mrsSendList.Update
                End If
            Next
            
            mblnSendChange = True
        End If
    End With
End Sub

Private Sub SetMainComandBars(ByVal intListType As Integer, ByVal lngRow As Long)
    '根据当前记录清单类型及当前记录，设置主窗体的菜单状态
    Dim cbrControl As CommandBarControl
    Dim cbrMenu As CommandBarControl
    Dim blnExists As Boolean
    
    If lngRow = 0 Then Exit Sub
    
    Select Case intListType
        Case mListType.发药
            '“拒发”状态的切换
            If Val(vsfList(intListType).TextMatrix(lngRow, mIntCol发药_收发ID)) = 0 Then Exit Sub
            
            Set cbrMenu = frm部门发药管理New.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Reject, , True)
            Set cbrControl = frm部门发药管理New.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Reject, , True)
            
            If Val(vsfList(intListType).TextMatrix(lngRow, mIntCol发药_执行状态)) = mState.拒发 Then
                If Not cbrMenu Is Nothing Then cbrMenu.Enabled = True
                If Not cbrControl Is Nothing Then cbrControl.Enabled = True
            Else
                If Not cbrMenu Is Nothing Then cbrMenu.Enabled = False
                If Not cbrControl Is Nothing Then cbrControl.Enabled = False
            End If
            
            Set cbrMenu = frm部门发药管理New.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Verify, , True)
            Set cbrControl = frm部门发药管理New.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Verify, , True)
            If Val(vsfList(intListType).TextMatrix(lngRow, mIntCol发药_执行状态)) = mState.发药 Then
                If Not cbrMenu Is Nothing Then cbrMenu.Enabled = True
                If Not cbrControl Is Nothing Then cbrControl.Enabled = True
            Else
                If Not cbrMenu Is Nothing Then cbrMenu.Enabled = False
                If Not cbrControl Is Nothing Then cbrControl.Enabled = False
            End If
        Case mListType.拒发
            '“拒发”、“恢复”状态的切换
            Set cbrMenu = frm部门发药管理New.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Reject, , True)
            Set cbrControl = frm部门发药管理New.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_Reject, , True)
            
            If Val(vsfList(intListType).TextMatrix(lngRow, mIntCol拒发_收发ID)) = 0 Then
                If Not cbrMenu Is Nothing Then cbrMenu.Enabled = False
                If Not cbrControl Is Nothing Then cbrControl.Enabled = False
            Else
                If Val(vsfList(intListType).TextMatrix(lngRow, mIntCol拒发_执行状态)) = mState.拒发 Then
                    If Not cbrMenu Is Nothing Then cbrMenu.Enabled = True
                    If Not cbrControl Is Nothing Then cbrControl.Enabled = True
                Else
                    If Not cbrMenu Is Nothing Then cbrMenu.Enabled = False
                    If Not cbrControl Is Nothing Then cbrControl.Enabled = False
                End If
            End If
            
            Set cbrMenu = frm部门发药管理New.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_RejectRestore, , True)
            Set cbrControl = frm部门发药管理New.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_RejectRestore, , True)
            
            If Val(vsfList(intListType).TextMatrix(lngRow, mIntCol拒发_收发ID)) = 0 Then
                If Not cbrMenu Is Nothing Then cbrMenu.Enabled = False
                If Not cbrControl Is Nothing Then cbrControl.Enabled = False
            Else
                If Val(vsfList(intListType).TextMatrix(lngRow, mIntCol拒发_执行状态)) = mState.拒发_恢复 Then
                    If Not cbrMenu Is Nothing Then cbrMenu.Enabled = True
                    If Not cbrControl Is Nothing Then cbrControl.Enabled = True
                Else
                    If Not cbrMenu Is Nothing Then cbrMenu.Enabled = False
                    If Not cbrControl Is Nothing Then cbrControl.Enabled = False
                End If
            End If
        Case mListType.退药
            If vsfList(intListType).TextMatrix(lngRow, mIntCol退药_收发ID) = "" Or (Not IsNumeric(vsfList(intListType).TextMatrix(lngRow, mIntCol退药_收发ID))) Then
                Exit Sub
            End If
            
            Set cbrMenu = frm部门发药管理New.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_Edit_Dept_VerifySign, , True)
            Set cbrControl = frm部门发药管理New.cbsMain(2).Controls.Find(xtpControlButton, mconMenu_Edit_Dept_VerifySign, , True)
            
            With vsfList(intListType)
                blnExists = RecipeSendWork_JudgeSign(.TextMatrix(lngRow, mIntCol退药_单据), .TextMatrix(lngRow, mIntCol退药_NO), IIf(Val(.TextMatrix(lngRow, mIntCol退药_执行状态)) = mState.退药_原始记录, 2, 3), .TextMatrix(lngRow, mIntCol退药_收发ID))
                
                If blnExists Then
                    If Not cbrMenu Is Nothing Then cbrMenu.Enabled = True
                    If Not cbrControl Is Nothing Then cbrControl.Enabled = True
                Else
                    If Not cbrMenu Is Nothing Then cbrMenu.Enabled = False
                    If Not cbrControl Is Nothing Then cbrControl.Enabled = False
                End If
            End With
    End Select
End Sub

Private Sub SetSubTotal(ByVal vsfObj As VSFlexGrid, ByVal strSub As String)
    Dim lngRow As Long
    Dim intCol As Integer
    Dim intSumType As Integer
    Dim strSum As String
    
    '设置小计，合计
    With vsfObj
        .OutlineCol = 0
        .OutlineBar = 0
        .SubtotalPosition = flexSTBelow
        
        If .index = mListType.发药 Then
            .Subtotal flexSTSum, -1, mIntCol发药_金额, "###.00", , mListColor.SumTotal, True
            If strSub = "领药部门" Then
                .Subtotal flexSTSum, mIntCol发药_领药部门, mIntCol发药_金额, "###.00", , mListColor.SumTotal, False, "", mIntCol发药_领药部门, False
                .Subtotal flexSTSum, mIntCol发药_NO, mIntCol发药_金额, "###.00", , mListColor.SumTotal, False, "", mIntCol发药_NO, False
            ElseIf strSub = "NO" Then
                .Subtotal flexSTSum, mIntCol发药_NO, mIntCol发药_金额, "###.00", , mListColor.SumTotal, False, "", mIntCol发药_NO, False
            ElseIf strSub = "药品名称" Then
                .Subtotal flexSTSum, mIntCol发药_品名, mIntCol发药_金额, "###.00", , mListColor.SumTotal, False, "", mIntCol发药_品名, False
            ElseIf strSub = "住院号" Then
                .Subtotal flexSTSum, mIntCol发药_住院号, mIntCol发药_金额, "###.00", , mListColor.SumTotal, False, "", mIntCol发药_住院号, False
            ElseIf strSub = "床号" Then
                .Subtotal flexSTSum, mIntCol发药_床号, mIntCol发药_金额, "###.00", , mListColor.SumTotal, False, "", mIntCol发药_床号, False
            ElseIf strSub = "姓名" Then
                .Subtotal flexSTSum, mIntCol发药_姓名, mIntCol发药_金额, "###.00", , mListColor.SumTotal, False, "", mIntCol发药_姓名, False
            End If
        ElseIf .index = mListType.汇总 Then
            
            If mcondition.bln按科室汇总 = True Then
                .Subtotal flexSTSum, -1, mIntCol科室汇总_应发金额, "###.00", , mListColor.SumTotal, True
                If strSub = "领药部门" Then
                    .Subtotal flexSTSum, mIntCol科室汇总_领药部门, mIntCol科室汇总_应发金额, "###.00", , mListColor.SumTotal, False, "", mIntCol科室汇总_领药部门, False
                ElseIf strSub = "药品名称" Then
                    .Subtotal flexSTSum, mIntCol科室汇总_品名, mIntCol科室汇总_应发金额, "###.00", , mListColor.SumTotal, False, "", mIntCol科室汇总_品名, False
                End If
            Else
                .Subtotal flexSTSum, -1, mIntCol汇总_金额, "###.00", , mListColor.SumTotal, True
'                If strSub = "药品名称" Then
'                    .Subtotal flexSTSum, mIntCol汇总_品名, mIntCol科室汇总_应发金额, "###.00", , mListColor.SumTotal, False, "", mIntCol汇总_品名, False
'                End If
            End If
        End If
        
        For lngRow = 1 To .rows - 1
            If .IsSubtotal(lngRow) = True Then
                '找到是哪一列合计
                If lngRow = .rows - 1 Then
                    '如果是最后一行，则是合计
                    strSum = ""
                    intSumType = mSubTotalType.SubSum
                Else
                    '如果不是，则根据列颜色判断合计的类型
                    For intCol = 1 To .Cols - 1
                        .Row = lngRow
                        .Col = intCol
                        If .CellForeColor = mListColor.SumTotal Then
                            If intCol = .ColIndex("领药部门") Then
                                intSumType = mSubTotalType.SubByDept
                            ElseIf intCol = .ColIndex("姓名") Then
                                intSumType = mSubTotalType.SubByPeople
                            ElseIf intCol = .ColIndex("NO") Then
                                intSumType = mSubTotalType.SubByNo
                            ElseIf intCol = .ColIndex("药品名称") Then
                                intSumType = mSubTotalType.SubByDrug
                            ElseIf intCol = .ColIndex("住院号") Then
                                intSumType = mSubTotalType.SubByHosNumber
                            ElseIf intCol = .ColIndex("床号") Then
                                intSumType = mSubTotalType.SubByBedNumber
                            End If

                            strSum = Trim(Replace(.TextMatrix(lngRow, intCol), "Total", ""))

                            Exit For
                        End If
                    Next
                End If

                SetGridSubTotal vsfObj, lngRow, strSum, mrsSendList, intSumType
            End If
        Next
    End With
End Sub

Private Sub SetTempOperate(ByVal intLastType As Integer, ByVal intThisType As Integer)
    '保存或提取临时的操作
    Dim strValue As String
    
    '保存上个页面的操作
    Select Case intLastType
        Case mListType.发药
            SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\发药", "显示缺药药品", IIf(mcondition.bln显示缺药, 1, 0)
            
            SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\发药", "显示退药待发单据", IIf(mcondition.bln显示退药待发单据, 1, 0)
            
            SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\发药", "显示相关信息", IIf(mcondition.bln显示扩展信息, 1, 0)
        Case mListType.汇总
            SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\汇总", "按批次汇总", IIf(mcondition.bln按批次汇总, 1, 0)
        Case mListType.退药
            SaveSetting "ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\退药", "显示过程单据", IIf(mcondition.bln显示过程单据, 1, 0)
    End Select
    
    '保存上个页面的列设置
    SaveListColState intLastType
    
    '取当前页面的操作
    Select Case intThisType
        Case mListType.发药
            strValue = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\发药", "显示缺药药品", "1")
            mcondition.bln显示缺药 = (Val(strValue) = 1)
            
            strValue = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\发药", "显示退药待发单据", "1")
            mcondition.bln显示退药待发单据 = (Val(strValue) = 1)
        Case mListType.汇总
            strValue = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\汇总", "按批次汇总", "0")
            mcondition.bln按批次汇总 = (Val(strValue) = 1)
        Case mListType.退药
            strValue = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\退药", "显示过程单据", "0")
            mcondition.bln显示过程单据 = (Val(strValue) = 1)
    End Select
End Sub

Private Sub SetGridSubTotal(ByVal vsfObj As VSFlexGrid, ByVal intRow As Integer, _
            ByVal strSub As String, ByVal rsData As ADODB.Recordset, ByVal intSubType As Integer)
    '用于产生表格小计、合计，统计的内容由合计类型决定。
    '合计：统计科室数量、病人数量、单据数量、待发药品数量、金额
    '按领药部门小计：统计病人数量、单据数量、待发药品数量、金额
    '按病人小计：统计单据数量、待发药品数量、金额
    '按单据小计：统计待发药品数量、金额
    '按品名小计：统计病人数量、单据数量、金额
    '按病人姓名，住院号，床号小计：单据数量、待发药品数量、金额
    Dim rsSub As ADODB.Recordset
    
    Dim str当前药品 As String
    Dim str当前NO As String
    Dim lng当前病人 As Long
    Dim str当前领药部门 As String
    
    Dim dbl待发药品数量 As Double
    Dim dblNO数量 As Double
    Dim dbl病人数量 As Double
    Dim dbl科室数量 As Double
    Dim Dbl金额 As Double
    Dim dbl实发金额 As Double
    Dim dbl留存金额 As Double
    Dim strTemp As String
    Dim str病人id As String
    Dim str姓名 As String
    
    
    Dim strSumText As String
    
    Dim intSumCol As Integer
    
    Dim strFilter As String
    
    '是否显示退药待发药品
    If mcondition.bln显示退药待发单据 = False Then
        strFilter = " And 记录状态=1 "
    End If
    
    Set rsSub = rsData
    
    vsfObj.MergeCells = flexMergeRestrictRows
    vsfObj.MergeRow(intRow) = True
    
    Select Case intSubType
        Case mSubTotalType.SubByNo
            '按单据号小计
            rsSub.Filter = "NO='" & strSub & "'" & strFilter
            rsSub.Sort = "品名"
            
            str当前药品 = ""
            dbl待发药品数量 = 0
            
            If rsSub.RecordCount > 0 Then
                Do While Not rsSub.EOF
                    If (vsfObj.index = mListType.汇总 And rsSub!执行状态 = mState.发药) Or _
                        (vsfObj.index <> mListType.汇总 And (mcondition.bln显示缺药 = True And rsSub!执行状态 <= 3) Or (mcondition.bln显示缺药 = False And ((rsSub!执行状态 = mState.发药 Or rsSub!执行状态 = mState.不处理 Or rsSub!执行状态 = mState.拒发)))) Then
                        If str当前药品 <> rsSub!品名 Then
                            str当前药品 = rsSub!品名
                            dbl待发药品数量 = dbl待发药品数量 + 1
                        End If
                    End If
                    rsSub.MoveNext
                Loop
                
                Dbl金额 = 0
                rsSub.MoveFirst
                Do While Not rsSub.EOF
                    If (vsfObj.index = mListType.汇总 And rsSub!执行状态 = mState.发药) Or _
                        (vsfObj.index <> mListType.汇总 And (mcondition.bln显示缺药 = True And rsSub!执行状态 <= 3) Or (mcondition.bln显示缺药 = False And ((rsSub!执行状态 = mState.发药 Or rsSub!执行状态 = mState.不处理 Or rsSub!执行状态 = mState.拒发)))) Then
                        Dbl金额 = Dbl金额 + rsSub!金额
                    End If
                    rsSub.MoveNext
                Loop
            End If
        Case mSubTotalType.SubByPeople, mSubTotalType.SubByHosNumber, mSubTotalType.SubByBedNumber
            '按病人姓名，住院号，床号汇总时
            If intSubType = mSubTotalType.SubByPeople Then
                rsSub.Filter = "姓名='" & strSub & "'" & strFilter
            ElseIf intSubType = mSubTotalType.SubByHosNumber Then
                rsSub.Filter = "住院号='" & strSub & "'" & strFilter
            ElseIf intSubType = mSubTotalType.SubByBedNumber Then
                rsSub.Filter = "床号='" & strSub & "'" & strFilter
            End If
            
            rsSub.Sort = "NO"
            
            str当前药品 = ""
            str当前NO = ""
            dblNO数量 = 0
            dbl待发药品数量 = 0
            
            If rsSub.RecordCount > 0 Then
                Do While Not rsSub.EOF
                    If (vsfObj.index = mListType.汇总 And rsSub!执行状态 = mState.发药) Or _
                        (vsfObj.index <> mListType.汇总 And (mcondition.bln显示缺药 = True And rsSub!执行状态 <= 3) Or (mcondition.bln显示缺药 = False And ((rsSub!执行状态 = mState.发药 Or rsSub!执行状态 = mState.不处理 Or rsSub!执行状态 = mState.拒发)))) Then
                        If InStr(1, "," & str当前NO, "," & rsSub!NO & ",") < 1 Then
                            str当前NO = str当前NO & rsSub!NO & ","
                            dblNO数量 = dblNO数量 + 1
                        End If
                    End If
                    rsSub.MoveNext
                Loop
                
                rsSub.Sort = "品名"
                Do While Not rsSub.EOF
                    If (vsfObj.index = mListType.汇总 And rsSub!执行状态 = mState.发药) Or _
                        (vsfObj.index <> mListType.汇总 And (mcondition.bln显示缺药 = True And rsSub!执行状态 <= 3) Or (mcondition.bln显示缺药 = False And ((rsSub!执行状态 = mState.发药 Or rsSub!执行状态 = mState.不处理 Or rsSub!执行状态 = mState.拒发)))) Then
                        If str当前药品 <> rsSub!品名 Then
                            str当前药品 = rsSub!品名
                            dbl待发药品数量 = dbl待发药品数量 + 1
                        End If
                    End If
                    rsSub.MoveNext
                Loop
                
                Dbl金额 = 0
                rsSub.MoveFirst
                Do While Not rsSub.EOF
                    If (vsfObj.index = mListType.汇总 And rsSub!执行状态 = mState.发药) Or _
                        (vsfObj.index <> mListType.汇总 And (mcondition.bln显示缺药 = True And rsSub!执行状态 <= 3) Or (mcondition.bln显示缺药 = False And ((rsSub!执行状态 = mState.发药 Or rsSub!执行状态 = mState.不处理 Or rsSub!执行状态 = mState.拒发)))) Then
                        Dbl金额 = Dbl金额 + rsSub!金额
                    End If
                    rsSub.MoveNext
                Loop
            End If
        Case mSubTotalType.SubByDept
            '按领药部门小计
            rsSub.Filter = "领药部门='" & strSub & "'" & strFilter
            rsSub.Sort = "姓名,NO"
            
            lng当前病人 = 0
            str当前药品 = ""
            str当前NO = ""
            dbl病人数量 = 0
            dblNO数量 = 0
            dbl待发药品数量 = 0
            str病人id = ""
            str当前NO = ""
            
            If rsSub.RecordCount > 0 Then
                Do While Not rsSub.EOF
                    If (vsfObj.index = mListType.汇总 And rsSub!执行状态 = mState.发药) Or _
                        (vsfObj.index <> mListType.汇总 And (mcondition.bln显示缺药 = True And rsSub!执行状态 <= 3) Or (mcondition.bln显示缺药 = False And ((rsSub!执行状态 = mState.发药 Or rsSub!执行状态 = mState.不处理 Or rsSub!执行状态 = mState.拒发)))) Then
                        If InStr(1, "," & str病人id, "," & rsSub!病人ID & ",") < 1 Or InStr(1, "," & str姓名, "," & rsSub!姓名 & ",") < 1 Then
                            str病人id = str病人id & rsSub!病人ID & ","
                            str姓名 = str姓名 & rsSub!姓名 & ","
                            dbl病人数量 = dbl病人数量 + 1
                        End If
                   
                        If InStr(1, "," & str当前NO, "," & rsSub!NO & ",") < 1 Then
                            str当前NO = str当前NO & rsSub!NO & ","
                            dblNO数量 = dblNO数量 + 1
                        End If
                    End If
                    rsSub.MoveNext
                Loop
                
                rsSub.Sort = "品名,批次"
                Do While Not rsSub.EOF
                    If (vsfObj.index = mListType.汇总 And rsSub!执行状态 = mState.发药) Or _
                        (vsfObj.index <> mListType.汇总 And (mcondition.bln显示缺药 = True And rsSub!执行状态 <= 3) Or (mcondition.bln显示缺药 = False And ((rsSub!执行状态 = mState.发药 Or rsSub!执行状态 = mState.不处理 Or rsSub!执行状态 = mState.拒发)))) Then
                        If str当前药品 <> rsSub!品名 Then
                            str当前药品 = rsSub!品名
                            dbl待发药品数量 = dbl待发药品数量 + 1
                        End If
                    End If
                    rsSub.MoveNext
                Loop
                
                dbl留存金额 = 0
                Dbl金额 = 0
                rsSub.MoveFirst
                Do While Not rsSub.EOF
                    If (vsfObj.index = mListType.汇总 And rsSub!执行状态 = mState.发药) Or _
                        (vsfObj.index <> mListType.汇总 And (mcondition.bln显示缺药 = True And rsSub!执行状态 <= 3) Or (mcondition.bln显示缺药 = False And ((rsSub!执行状态 = mState.发药 Or rsSub!执行状态 = mState.不处理 Or rsSub!执行状态 = mState.拒发)))) Then
                        Dbl金额 = Dbl金额 + rsSub!金额
                        
                        If vsfObj.index = mListType.汇总 Then
                            If mcondition.bln按批次汇总 Then
                                If InStr(1, strTemp & ",", rsSub!领药部门ID & "," & rsSub!药品ID & "," & rsSub!批次 & ",") < 1 Then
                                    dbl留存金额 = dbl留存金额 + ((rsSub!留存数量 + FormatEx(GetChargeOffCount(rsSub!领药部门ID, rsSub!药品ID, rsSub!批次), 5)) * rsSub!单价)
                                    strTemp = strTemp & rsSub!领药部门ID & "," & rsSub!药品ID & "," & rsSub!批次
                                End If
                            Else
                                If InStr(1, strTemp & ",", rsSub!领药部门ID & "," & rsSub!药品ID & ",") < 1 Then
                                
                                    dbl留存金额 = dbl留存金额 + ((rsSub!留存数量 + FormatEx(GetChargeOffCount(rsSub!领药部门ID, rsSub!药品ID, rsSub!批次), 5)) * rsSub!单价)
                                    strTemp = strTemp & rsSub!领药部门ID & "," & rsSub!药品ID
        
                                End If
                            End If
                        End If
                    End If
                    rsSub.MoveNext
                Loop
                
            End If
        Case mSubTotalType.SubByDrug
            '按药品小计
            rsSub.Filter = "品名='" & strSub & "'" & strFilter
            rsSub.Sort = "批次,姓名,NO"
            
            lng当前病人 = 0
            str当前NO = ""
            dbl病人数量 = 0
            dblNO数量 = 0
            dbl待发药品数量 = 0
            str病人id = ""
            str当前NO = ""
            
            If rsSub.RecordCount > 0 Then
                Do While Not rsSub.EOF
                    If (vsfObj.index = mListType.汇总 And rsSub!执行状态 = mState.发药) Or _
                        (vsfObj.index <> mListType.汇总 And (mcondition.bln显示缺药 = True And rsSub!执行状态 <= 3) Or (mcondition.bln显示缺药 = False And ((rsSub!执行状态 = mState.发药 Or rsSub!执行状态 = mState.不处理 Or rsSub!执行状态 = mState.拒发)))) Then
                        If InStr(1, "," & str病人id, "," & rsSub!病人ID & ",") < 1 Or InStr(1, "," & str姓名, "," & rsSub!姓名 & ",") < 1 Then
                            str病人id = str病人id & rsSub!病人ID & ","
                            str姓名 = str姓名 & rsSub!姓名 & ","
                            dbl病人数量 = dbl病人数量 + 1
                        End If
                   
                        If InStr(1, "," & str当前NO, "," & rsSub!NO & ",") < 1 Then
                            str当前NO = str当前NO & rsSub!NO & ","
                            dblNO数量 = dblNO数量 + 1
                        End If
                    End If
                    rsSub.MoveNext
                Loop
                
                Dbl金额 = 0
                dbl留存金额 = 0
                rsSub.MoveFirst
                Do While Not rsSub.EOF
                    If (vsfObj.index = mListType.汇总 And rsSub!执行状态 = mState.发药) Or _
                        (vsfObj.index <> mListType.汇总 And (mcondition.bln显示缺药 = True And rsSub!执行状态 <= 3) Or (mcondition.bln显示缺药 = False And ((rsSub!执行状态 = mState.发药 Or rsSub!执行状态 = mState.不处理 Or rsSub!执行状态 = mState.拒发)))) Then
                        Dbl金额 = Dbl金额 + rsSub!金额
                        
                        If vsfObj.index = mListType.汇总 Then
                            If mcondition.bln按批次汇总 Then
                                If InStr(1, strTemp & ",", rsSub!领药部门ID & "," & rsSub!药品ID & "," & rsSub!批次 & ",") < 1 Then
                                    dbl留存金额 = dbl留存金额 + ((rsSub!留存数量 + FormatEx(GetChargeOffCount(rsSub!领药部门ID, rsSub!药品ID, rsSub!批次), 5)) * rsSub!单价)
                                    strTemp = strTemp & rsSub!领药部门ID & "," & rsSub!药品ID & "," & rsSub!批次
                                End If
                            Else
                                If InStr(1, strTemp & ",", rsSub!领药部门ID & "," & rsSub!药品ID & ",") < 1 Then
                                
                                    dbl留存金额 = dbl留存金额 + ((rsSub!留存数量 + FormatEx(GetChargeOffCount(rsSub!领药部门ID, rsSub!药品ID, rsSub!批次), 5)) * rsSub!单价)
                                    strTemp = strTemp & rsSub!领药部门ID & "," & rsSub!药品ID
        
                                End If
                            End If
                        End If
                    End If
                    rsSub.MoveNext
                Loop
            End If
        Case mSubTotalType.SubSum
            '合计
            rsSub.Filter = IIf(mcondition.bln显示退药待发单据 = False, "记录状态=1", "")
            rsSub.Sort = "领药部门,药品id,批次,姓名,NO"
            
            str当前领药部门 = ""
            lng当前病人 = 0
            str当前药品 = ""
            str当前NO = ""
            dbl科室数量 = 0
            dbl病人数量 = 0
            dblNO数量 = 0
            dbl待发药品数量 = 0
            Dbl金额 = 0
            dbl留存金额 = 0
            str病人id = ""
            str当前NO = ""
            
            If rsSub.RecordCount > 0 Then
                Do While Not rsSub.EOF
                    If (vsfObj.index = mListType.汇总 And rsSub!执行状态 = mState.发药) Or _
                        (vsfObj.index <> mListType.汇总 And (mcondition.bln显示缺药 = True And rsSub!执行状态 <= 3) Or (mcondition.bln显示缺药 = False And ((rsSub!执行状态 = mState.发药 Or rsSub!执行状态 = mState.不处理 Or rsSub!执行状态 = mState.拒发)))) Then
                        If str当前领药部门 <> rsSub!领药部门 Then
                            str当前领药部门 = rsSub!领药部门
                            dbl科室数量 = dbl科室数量 + 1
                        End If
                        
                        If InStr(1, "," & str病人id, "," & rsSub!病人ID & ",") < 1 Or InStr(1, "," & str姓名, "," & rsSub!姓名 & ",") < 1 Then
                            str病人id = str病人id & rsSub!病人ID & ","
                            str姓名 = str姓名 & rsSub!姓名 & ","
                            dbl病人数量 = dbl病人数量 + 1
                        End If
                        
                        If InStr(1, "," & str当前NO, "," & rsSub!NO & ",") < 1 Then
                            str当前NO = str当前NO & rsSub!NO & ","
                            dblNO数量 = dblNO数量 + 1
                        End If
                        
                        Dbl金额 = Dbl金额 + rsSub!金额
                        
                        If vsfObj.index = mListType.汇总 Then
                            If mcondition.bln按批次汇总 Then
                                If InStr(1, strTemp & ",", rsSub!领药部门ID & "," & rsSub!药品ID & "," & rsSub!批次 & ",") < 1 Then
                                    dbl留存金额 = dbl留存金额 + ((rsSub!留存数量 + FormatEx(GetChargeOffCount(rsSub!领药部门ID, rsSub!药品ID, rsSub!批次), 5)) * rsSub!单价)
                                    strTemp = strTemp & rsSub!领药部门ID & "," & rsSub!药品ID & "," & rsSub!批次
                                End If
                            Else
                                If InStr(1, strTemp & ",", rsSub!领药部门ID & "," & rsSub!药品ID & ",") < 1 Then
                                
                                    dbl留存金额 = dbl留存金额 + ((rsSub!留存数量 + FormatEx(GetChargeOffCount(rsSub!领药部门ID, rsSub!药品ID, rsSub!批次), 5)) * rsSub!单价)
                                    strTemp = strTemp & rsSub!领药部门ID & "," & rsSub!药品ID
        
                                End If
                            End If
                        End If
                    End If
                    
                    rsSub.MoveNext
                Loop
                
                rsSub.Sort = "品名"
                Do While Not rsSub.EOF
                    If (vsfObj.index = mListType.汇总 And rsSub!执行状态 = mState.发药) Or _
                        (vsfObj.index <> mListType.汇总 And (mcondition.bln显示缺药 = True And rsSub!执行状态 <= 3) Or (mcondition.bln显示缺药 = False And ((rsSub!执行状态 = mState.发药 Or rsSub!执行状态 = mState.不处理 Or rsSub!执行状态 = mState.拒发)))) Then
                        If str当前药品 <> rsSub!品名 Then
                            str当前药品 = rsSub!品名
                            dbl待发药品数量 = dbl待发药品数量 + 1
                        End If
                    End If
                    rsSub.MoveNext
                Loop
            End If
    End Select
    
    If intSubType = mSubTotalType.SubSum Then
        If vsfObj.index = mListType.发药 Then
            strSumText = "合计： " & dbl科室数量 & "个科室  " & dbl病人数量 & "个病人  " & dblNO数量 & "张单据  " & dbl待发药品数量 & "种药品待发 " & "共计" & Format(Dbl金额, "#####0.00;-#####0.00; ;") & "元"
        ElseIf vsfObj.index = mListType.汇总 Then
            If mcondition.bln按科室汇总 = True Then
                strSumText = "合计： " & dbl科室数量 & "个科室  " & dbl待发药品数量 & "种药品待发 " & "应发金额共计" & Format(Dbl金额, "#####0.00;-#####0.00; ;") & "元 实发金额共计" & Format(Dbl金额 - dbl留存金额, "#####0.00;-#####0.00; ;") & "元"
            Else
                strSumText = "合计： " & dbl待发药品数量 & "种药品待发 " & "应发金额共计" & Format(Dbl金额, "#####0.00;-#####0.00; ;") & "元 实发金额共计" & Format(Dbl金额 - dbl留存金额, "#####0.00;-#####0.00; ;") & "元"
            End If
        End If
        
        vsfObj.Cell(flexcpText, intRow, 1, intRow, vsfObj.Cols - 1) = strSumText
        vsfObj.Cell(flexcpAlignment, intRow, 1, intRow, 1) = flexAlignLeftCenter
        vsfObj.Cell(flexcpForeColor, intRow, 1, intRow, vsfObj.Cols - 1) = mListColor.SumTotal
    Else
        If intSubType = mSubTotalType.SubByDept Then
            strSumText = "[领药部门：" & strSub & "]小计： "
        ElseIf intSubType = mSubTotalType.SubByNo Then
            strSumText = "[NO：" & strSub & "]小计： "
        ElseIf intSubType = mSubTotalType.SubByDrug Then
            strSumText = "[药品：" & strSub & "]小计： "
        ElseIf intSubType = mSubTotalType.SubByPeople Then
            strSumText = "[姓名：" & strSub & "]小计： "
        ElseIf intSubType = mSubTotalType.SubByHosNumber Then
            strSumText = "[住院号：" & strSub & "]小计： "
        ElseIf intSubType = mSubTotalType.SubByBedNumber Then
            strSumText = "[床位号：" & strSub & "]小计： "
        End If
        
        If vsfObj.index = mListType.发药 Then
            If dbl病人数量 > 0 Then
                strSumText = strSumText & dbl病人数量 & "个病人  "
            End If
            If dblNO数量 > 0 Then
                strSumText = strSumText & dblNO数量 & "张单据  "
            End If
        End If
        
        If intSubType = mSubTotalType.SubByDrug Then
            If vsfObj.index = mListType.汇总 Then
                strSumText = strSumText & "应发金额共计" & Format(Dbl金额, "#####0.00;-#####0.00; ;") & "元 实发金额共计" & Format(Dbl金额 - dbl留存金额, "#####0.00;-#####0.00; ;") & "元"
            Else
                strSumText = strSumText & "共计" & Format(Dbl金额, "#####0.00;-#####0.00; ;") & "元"
            End If
        ElseIf dbl待发药品数量 > 0 Then
            If vsfObj.index = mListType.汇总 Then
                strSumText = strSumText & dbl待发药品数量 & "种药品待发 " & "应发金额共计" & Format(Dbl金额, "#####0.00;-#####0.00; ;") & "元 实发金额共计" & Format(Dbl金额 - dbl留存金额, "#####0.00;-#####0.00; ;") & "元"
            Else
                strSumText = strSumText & dbl待发药品数量 & "种药品待发 " & "共计" & Format(Dbl金额, "#####0.00;-#####0.00; ;") & "元"
            End If
        End If
        
        If vsfObj.index = mListType.发药 Then
            intSumCol = mIntCol发药_审查结果 + 1
        ElseIf vsfObj.index = mListType.汇总 Then
            If mcondition.bln按科室汇总 = True Then
                intSumCol = mIntCol科室汇总_领药部门
            Else
                intSumCol = mIntCol汇总_品名
            End If
        End If
        
        vsfObj.Cell(flexcpText, intRow, intSumCol, intRow, vsfObj.Cols - 1) = strSumText
        vsfObj.Cell(flexcpAlignment, intRow, intSumCol, intRow, intSumCol) = flexAlignLeftCenter
        vsfObj.Cell(flexcpForeColor, intRow, intSumCol, intRow, vsfObj.Cols - 1) = mListColor.SumTotal
    End If
End Sub

Public Sub ShowList(ByVal intType As Integer, ByVal lng库房ID As Long)
    Dim objVSF As VSFlexGrid
    
    With mcondition
        If lng库房ID > 0 And .lng药房id <> lng库房ID Then
            .lng药房id = lng库房ID
            .bln配制中心 = CheckIsCenter(.lng药房id)
            Call GetDrugDigit(.lng药房id, "药品部门发药", mintUnit, mintCostDigit, mintPriceDigit, mintNumberDigit, mintMoneyDigit)
        End If
         
        If .intListType <> intType Then
            Call SetTempOperate(.intListType, intType)
            .intListType = intType
        End If
        Call Load配药人(.lng药房id)
        Call Load核查人(.lng药房id)
    End With
    
    For Each objVSF In vsfList
        If objVSF.index = mcondition.intListType Then
            objVSF.Visible = True
        Else
            objVSF.Visible = False
        End If
    Next
    
    mblnShowReject = (intType = mListType.拒发)
    
    If mcondition.intListType = mListType.发药 Or mcondition.intListType = mListType.汇总 Then
        picAssist.Visible = True
    Else
        picAssist.Visible = False
    End If
            
    Call SetComandBars(mcondition.intListType)
    
    Call InitColSelList(mcondition.intListType)
    
    vsfChargeOff.Visible = False
    
    If mcondition.intListType = mListType.发药 Then
        picInfo.Visible = mcondition.bln显示扩展信息
        fraH.Visible = True
        
    Else
        picInfo.Visible = False
        fraH.Visible = False
    End If

    Call Form_Resize
    '根据发药数据集的变化判断是否重新刷新列表
    If (mcondition.intListType = mListType.汇总 Or mcondition.intListType = mListType.拒发) And mblnSendChange = True Then
        Call RefreshList_Sum
        Call RefreshList_Reject
        mblnSendChange = False
    End If
End Sub

Private Sub cbo配药人_Click()
'    Exit Sub
End Sub

Private Sub cbo配药人_KeyDown(KeyCode As Integer, Shift As Integer)
    mblnDrop = False
    If KeyCode = 13 Then mblnDrop = SendMessage(cbo配药人.hWnd, CB_GETDROPPEDSTATE, 0, 0) = 1
End Sub

Private Sub cbo配药人_KeyPress(KeyAscii As Integer)
Dim i As Long, intIdx As Integer
    Dim strText As String, strResult As String, strFilter As String

    If KeyAscii = 13 Then
        strText = UCase(cbo配药人.Text)
        If cbo配药人.ListIndex <> -1 Then
            '弹出列表时,又在文本框输入了内容
            If strText <> cbo配药人.List(cbo配药人.ListIndex) Then Call zlControl.CboSetIndex(cbo配药人.hWnd, -1)
        End If
        If strText = "" Then
            cbo配药人.ListIndex = -1
        ElseIf cbo配药人.ListIndex = -1 Then
            intIdx = -1

            For i = 1 To cbo配药人.ListCount - 1
                If Mid(cbo配药人.List(i), 1, InStr(1, cbo配药人.List(i), "-") - 1) = strText _
                    Or Mid(cbo配药人.List(i), InStr(1, cbo配药人.List(i), "-")) = strText Then
                    intIdx = i
                    Exit For
                End If
            Next

            If intIdx = -1 Then
                For i = 1 To cbo配药人.ListCount - 1
                    If UCase(cbo配药人.List(i)) Like strText & "*" Then
                        intIdx = i
                    End If
                Next
            End If

            cbo配药人.ListIndex = intIdx
            SendMessage cbo配药人.hWnd, CB_SHOWDROPDOWN, True, 0
        ElseIf Not mblnDrop Then
            '回车光标经过
            Call cbo配药人_Click
            Exit Sub
        End If
        If cbo配药人.ListIndex = -1 Then
            cbo配药人.ListIndex = 0
        Else
            If intIdx <> -1 And mblnDrop Then
                '弹出回车-强行激活Click
                Call cbo配药人_Click
            ElseIf intIdx <> cbo配药人.ListIndex And intIdx <> -1 Then
                '弹出让选择-自动激活Click
                cbo配药人.SetFocus
                Exit Sub
            ElseIf intIdx <> -1 Then
                '一次性输中-强行激活Click
                Call cbo配药人_Click
            End If
        End If
    End If
End Sub


Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim Int单据 As Integer
    Dim strNo As String
    Dim objPopup As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim strsql As String
    Dim rsTmp As Recordset
    Dim lngPatiID As Long
    Dim str挂号单 As String
    Dim lng主页ID As Long
    Dim lng医嘱id As Long
    
    Select Case Control.Id
        Case conMenu_Tool_ShowShortage
            Control.Checked = Not Control.Checked
            Control.IconId = IIf(Not Control.Checked, 1, 2)
            
            mcondition.bln显示缺药 = Control.Checked
            
            Call RefreshList_Send
        Case conMenu_Tool_ShowReturnSend
            Control.Checked = Not Control.Checked
            Control.IconId = IIf(Not Control.Checked, 1, 2)
            
            mblnSendChange = True
            
            mcondition.bln显示退药待发单据 = Control.Checked
            
            Modify退药待发 mcondition.bln显示退药待发单据
            
            DoEvents
            Call RefreshList_Send
        Case conMenu_Tool_ShowInfo
            Control.Checked = Not Control.Checked
            Control.IconId = IIf(Not Control.Checked, 1, 2)
            
            mcondition.bln显示扩展信息 = Control.Checked
            picInfo.Visible = Control.Checked
            fraH.Visible = Control.Checked
            
            If picInfo.Visible Then
                vsfList(mListType.发药).Width = vsfList(mListType.发药).Width - picInfo.Width
            Else
                vsfList(mListType.发药).Width = vsfList(mListType.发药).Width + picInfo.Width
            End If
            
            Call cbsMain_Resize
        Case conMenu_Tool_SumByBatch
            Control.Checked = Not Control.Checked
            Control.IconId = IIf(Not Control.Checked, 1, 2)
            
            mcondition.bln按批次汇总 = Control.Checked
            
            Call RefreshList_Sum
        Case conMenu_Tool_ShowAllProcess
            Control.Checked = Not Control.Checked
            Control.IconId = IIf(Not Control.Checked, 1, 2)
            
            mcondition.bln显示过程单据 = Control.Checked
            
            Call RefreshList_Return
        
        Case conMenu_Tool_ShowPlug
            '功能：对病人过敏史/病生状态进行管理
            'Pass
            If vsfList(mcondition.intListType).Row = 0 Then Exit Sub
            If vsfList(mcondition.intListType).IsSubtotal(vsfList(mcondition.intListType).Row) = True Then Exit Sub
            If mcondition.intListType = mListType.发药 Then
                Int单据 = vsfList(mListType.发药).TextMatrix(vsfList(mListType.发药).Row, mIntCol发药_单据)
                strNo = vsfList(mListType.发药).TextMatrix(vsfList(mListType.发药).Row, mIntCol发药_NO)
            ElseIf mcondition.intListType = mListType.退药 Then
                Int单据 = vsfList(mListType.退药).TextMatrix(vsfList(mListType.退药).Row, mIntCol发药_单据)
                strNo = vsfList(mListType.退药).TextMatrix(vsfList(mListType.退药).Row, mIntCol发药_NO)
            End If
            
'            Call AdviceCheckWarn(Int单据, strNo, 21)
          
            '判断是住院还是门诊病人，如果没有找到记录（无医嘱）就退出
            strsql = "Select distinct B.病人id,nvl(B.主页id,0) 主页id,nvl(C.挂号单,'') 挂号单 " & _
                " From 药品收发记录 A,住院费用记录 B,病人医嘱记录 C " & _
                " Where A.费用id=B.Id And b.医嘱序号=c.Id And nvl(B.医嘱序号,0)<>0 And C.诊疗类别 IN('5','6','7')" & _
                " And A.单据=[2] And A.no=[1] " & _
                " Union All " & _
                " Select distinct B.病人id,0 主页id,nvl(C.挂号单,'') 挂号单 " & _
                " From 药品收发记录 A,门诊费用记录 B,病人医嘱记录 C " & _
                " Where A.费用id=B.Id And b.医嘱序号=c.Id And nvl(B.医嘱序号,0)<>0 And C.诊疗类别 IN('5','6','7')" & _
                " And A.单据=[2] And A.no=[1] "
            Set rsTmp = zlDataBase.OpenSQLRecord(strsql, Me.Caption, strNo, Int单据)
        
            If rsTmp.RecordCount = 0 Then
                rsTmp.Close
                Exit Sub
            End If
        
            lngPatiID = rsTmp!病人ID
            str挂号单 = NVL(rsTmp!挂号单)
            lng主页ID = rsTmp!主页ID
            
            Call gobjPass.zlPassCmdAlleyManage_YF(mlngMode, lngPatiID, lng主页ID, str挂号单)
        
        '弹出菜单：PASS命令
        Case mconMenu_PASS * 10# To mconMenu_PASS * 10# + 99
            If vsfList(mcondition.intListType).Row = 0 Then Exit Sub
            If vsfList(mcondition.intListType).IsSubtotal(vsfList(mcondition.intListType).Row) = True Then Exit Sub
            If mcondition.intListType = mListType.发药 Then
                Int单据 = vsfList(mListType.发药).TextMatrix(vsfList(mListType.发药).Row, mIntCol发药_单据)
                strNo = vsfList(mListType.发药).TextMatrix(vsfList(mListType.发药).Row, mIntCol发药_NO)
                lng医嘱id = Val(vsfList(mListType.发药).TextMatrix(vsfList(mListType.发药).Row, mIntCol发药_医嘱id))
            ElseIf mcondition.intListType = mListType.退药 Then
                Int单据 = vsfList(mListType.退药).TextMatrix(vsfList(mListType.退药).Row, mIntCol退药_单据)
                strNo = vsfList(mListType.退药).TextMatrix(vsfList(mListType.退药).Row, mIntCol退药_NO)
                lng医嘱id = Val(vsfList(mListType.退药).TextMatrix(vsfList(mListType.退药).Row, mIntCol退药_医嘱id))
            End If
            
            strsql = "Select distinct B.病人id,nvl(B.主页id,0) 主页id,nvl(C.挂号单,'') 挂号单 " & _
                " From 药品收发记录 A,住院费用记录 B,病人医嘱记录 C " & _
                " Where A.费用id=B.Id And b.医嘱序号=c.Id And nvl(B.医嘱序号,0)<>0 And C.诊疗类别 IN('5','6','7')" & _
                " And A.单据=[2] And A.no=[1] " & _
                " Union All " & _
                " Select distinct B.病人id,0 主页id,nvl(C.挂号单,'') 挂号单 " & _
                " From 药品收发记录 A,门诊费用记录 B,病人医嘱记录 C " & _
                " Where A.费用id=B.Id And b.医嘱序号=c.Id And nvl(B.医嘱序号,0)<>0 And C.诊疗类别 IN('5','6','7')" & _
                " And A.单据=[2] And A.no=[1] "
            Set rsTmp = zlDataBase.OpenSQLRecord(strsql, Me.Caption, strNo, Int单据)
        
            If rsTmp.RecordCount = 0 Then
                rsTmp.Close
                Exit Sub
            End If
        
            lngPatiID = rsTmp!病人ID
            str挂号单 = NVL(rsTmp!挂号单)
            lng主页ID = rsTmp!主页ID
            
            Call gobjPass.zlPassCommandBarExe_YF(mlngMode, Control.Id - (mconMenu_PASS * 10#), lngPatiID, lng主页ID, str挂号单, lng医嘱id)
        
        '弹出菜单：发药单据状态处理
        Case conMenu_Status_Verify
            SetSendBillState mChangeState.发药, vsfList(mListType.发药).Row
        Case conMenu_Status_Reject
            SetSendBillState mChangeState.拒发, vsfList(mListType.发药).Row
        Case conMenu_Status_Shortage
        Case conMenu_Status_NoProcess
            SetSendBillState mChangeState.不处理, vsfList(mListType.发药).Row
        Case conMenu_Status_AllSend
            SetSendBillState mChangeState.全部发药
        Case conMenu_Status_AllReject
            SetSendBillState mChangeState.全部拒发
        Case conMenu_Status_AllNoProcess
            SetSendBillState mChangeState.全部不处理
        
        '弹出菜单：退药单据状态处理
        Case conMenu_Status_AllReturn
            '全部退药
            SetAllReturn
        Case conMenu_Status_AllCancel
            '全部取消退药
            SetAllNotReturn
    End Select
End Sub


Private Function AdviceCheckWarn(ByVal Int单据 As Integer, ByVal strNo As String, ByVal lngCmd As Long, Optional ByVal lngRow As Long) As Long
'功能：调用Pass系统相关功能
'参数：lngCmd=
'        0-检测设置PASS菜单状态
'        21-病生状态/过敏史管理(只读)
'      lngRow=当前药品医嘱的行号，lngCmd=0时需要
'返回：检测PASS菜单时，返回>=0表示可以弹出菜单,其它返回-1
'说明：用药研究：涉及病人所有的医嘱(可以从数据库读,要求保存)
'      单药警告：应在用药审查过之后进行调用(有警告值)
    Dim rsTmp As New ADODB.Recordset
    Dim str药品 As String, str用法 As String, lng药品id As Long, str单量单位 As String
    Dim strsql As String, i As Long, k As Long
    Dim lngPatiID As Long
    Dim lngPassPati As Long
    Dim lng主页ID As Long
    Dim str挂号单 As String

    AdviceCheckWarn = -1

    On Error GoTo errH
    Screen.MousePointer = 11

    If strNo = "" Then Exit Function

    '检验PASS可用状态
    '-------------------------------------------------------------
    If PassGetState("PassEnable") = 0 Then
        MsgBox "当前合理用药监测系统不可用，请检查相关配置是否正确。", vbInformation, gstrSysName
        Screen.MousePointer = 0: Exit Function
    End If

    '判断是住院还是门诊病人，如果没有找到记录（无医嘱）就退出
    strsql = "Select distinct B.病人id,nvl(B.主页id,0) 主页id,nvl(C.挂号单,'') 挂号单 " & _
        " From 药品收发记录 A,住院费用记录 B,病人医嘱记录 C " & _
        " Where A.费用id=B.Id And b.医嘱序号=c.Id And nvl(B.医嘱序号,0)<>0 And C.诊疗类别 IN('5','6','7')" & _
        " And A.单据=[2] And A.no=[1] "
    Set rsTmp = zlDataBase.OpenSQLRecord(strsql, Me.Caption, strNo, Int单据)

    If rsTmp.RecordCount = 0 Then
        rsTmp.Close
        Exit Function
    End If

    lngPatiID = rsTmp!病人ID
    str挂号单 = zlStr.NVL(rsTmp!挂号单)
    lng主页ID = rsTmp!主页ID

    '传入病人就诊信息(PASS需要的基本内容,同一病人可不重复传入)
    '-------------------------------------------------------------
    If lngPatiID <> lngPassPati Then
        If str挂号单 <> "" Then               '门诊病人
            strsql = "Select 病人ID,Count(Distinct Trunc(登记时间)) as 就诊次数 From 病人挂号记录 Where 记录性质=1 And 记录状态=1 And 病人ID=[1] Group by 病人ID"
            strsql = "Select D.就诊次数,A.姓名,A.性别,A.出生日期," & _
                " C.编码 as 科室码,C.名称 as 科室名,E.编号 as 医生码,E.姓名 as 医生名" & _
                " From 病人信息 A,病人挂号记录 B,部门表 C,(" & strsql & ") D,人员表 E" & _
                " Where A.病人ID=B.病人ID And B.执行部门ID=C.ID And A.病人ID=D.病人ID" & _
                " And B.执行人=E.姓名(+) And A.病人ID=[1] And B.NO=[2]"
            Set rsTmp = zlDataBase.OpenSQLRecord(strsql, Me.Caption, lngPatiID, str挂号单)
            If rsTmp.EOF Then Screen.MousePointer = 0: Exit Function

            Call PassSetPatientInfo(lngPatiID, rsTmp!就诊次数, rsTmp!姓名, zlStr.NVL(rsTmp!性别), Format(rsTmp!出生日期, "yyyy-MM-dd"), "", "", _
                rsTmp!科室码 & "/" & rsTmp!科室名, IIf(Not IsNull(rsTmp!医生名), zlStr.NVL(rsTmp!医生码) & "/" & zlStr.NVL(rsTmp!医生名), ""), "")
        Else                                    '住院病人
            strsql = _
                " Select A.姓名,A.性别,A.出生日期,B.入院日期,B.出院日期," & _
                " C.编码 as 科室码,C.名称 as 科室名,D.编号 as 医生码,D.姓名 as 医生名" & _
                " From 病人信息 A,病案主页 B,部门表 C,人员表 D" & _
                " Where A.病人ID=B.病人ID And A.主页id=B.主页id And B.出院科室ID=C.ID" & _
                " And B.住院医师=D.姓名(+) And A.病人ID=[1] And B.主页ID=[2]"
            Set rsTmp = zlDataBase.OpenSQLRecord(strsql, Me.Caption, lngPatiID, lng主页ID)
            If rsTmp.EOF Then Screen.MousePointer = 0: Exit Function

            Call PassSetPatientInfo(lngPatiID, lng主页ID, rsTmp!姓名, zlStr.NVL(rsTmp!性别), Format(rsTmp!出生日期, "yyyy-MM-dd"), "", "", _
                rsTmp!科室码 & "/" & rsTmp!科室名, IIf(Not IsNull(rsTmp!医生名), zlStr.NVL(rsTmp!医生码) & "/" & zlStr.NVL(rsTmp!医生名), ""), _
                IIf(IsNull(rsTmp!出院日期), "", Format(rsTmp!出院日期, "yyyy-MM-dd")))
        End If
        lngPassPati = lngPatiID
    End If

    'PASS自定义菜单检测
    '-------------------------------------------------------------
    If lngCmd = 0 Then
        If mcondition.intListType = mListType.发药 Then
           '取药品名称
            str药品 = vsfList(mListType.发药).TextMatrix(lngRow, mIntCol发药_药品名称)
            lng药品id = vsfList(mListType.发药).TextMatrix(lngRow, mIntCol发药_药品ID)
            str单量单位 = vsfList(mListType.发药).TextMatrix(lngRow, mIntCol发药_单量单位)
            '取药品给药途径
            str用法 = vsfList(mListType.发药).TextMatrix(lngRow, mIntCol发药_用法)
        Else
            '取药品名称
            str药品 = vsfList(mListType.退药).TextMatrix(lngRow, mIntCol退药_药品名称)
            lng药品id = vsfList(mListType.退药).TextMatrix(lngRow, mIntCol退药_药品ID)
            str单量单位 = vsfList(mListType.退药).TextMatrix(lngRow, mIntCol退药_单量单位)
            '取药品给药途径
            str用法 = vsfList(mListType.退药).TextMatrix(lngRow, mIntCol退药_用法)
        End If
        
        '传入查询药品信息
        Call PassSetQueryDrug(lng药品id, str药品, str单量单位, str用法)

        '设置菜单可用状态
        Call SetPassMenuState

        AdviceCheckWarn = 1 '表示可以弹出菜单

        Screen.MousePointer = 0: Exit Function
    End If

    '执行相应的命令
    '-------------------------------------------------------------
    Call PassDoCommand(lngCmd)
    Screen.MousePointer = 0
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SetPassMenuState()
    '功能：设置Pass菜单可用状态
    'Pass
    Dim objPopup As CommandBarControl

    ''''一级菜单
    '药物临床信息参考
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("CPRRes") = 1

    '药品说明书
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item + 1, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("Directions") = 1

    '中国药典
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item + 2, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("Chp") = 1

    '病人用药教育
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item + 3, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("CPERes") = 1

    '检验值
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item + 4, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("CheckRes") = 1

    '专项信息
'    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item + 5, , True)
'    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("") = 1

    '医药信息中心
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item + 6, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("MEDInfo") = 1

    '药品配对信息
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item + 7, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("MATCH-DRUG") = 1

    '给药途径配对信息
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item + 8, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("MATCH-ROUTE") = 1

    '医院药品信息
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Item + 9, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("HisDrugInfo") = 1
    
    
    ''''专项信息二级菜单
    '药物-药物相互作用
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("DDIM") = 1
    
    '药物-食物相互使用
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec + 1, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("DFIM") = 1
    
    '国内注射剂体外配伍
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec + 2, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("MatchRes") = 1
    
    '国外注射剂体外配伍
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec + 3, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("TriessRes") = 1
    
    '禁忌症
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec + 4, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("DDCM") = 1
    
    '副作用
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec + 5, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("SIDE") = 1
    
    '老年人用药
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec + 6, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("GERI") = 1
    
    '儿童用药
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec + 7, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("PEDI") = 1
    
    '妊娠期用药
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec + 8, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("PREG") = 1
    
    '哺乳期用药
    Set objPopup = Me.cbsMain.ActiveMenuBar.Controls.Find(xtpControlButton, mconMenu_PASS_Spec + 9, , True)
    If Not objPopup Is Nothing Then objPopup.Enabled = PassGetState("LACT") = 1
End Sub
Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    Dim objVSF As VSFlexGrid
    
    On Error Resume Next
    
    If cbsMain.count > 1 Then
        Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
        For Each objVSF In vsfList
            If objVSF.Visible Then
                objVSF.Move lngLeft, lngTop, lngRight - lngLeft, lngBottom - lngTop - IIf(picAssist.Visible, picAssist.Height + 50, 0)
                
                fraColSel.Left = objVSF.Left + objVSF.ColWidth(0) - fraColSel.Width - 50
                fraColSel.Top = objVSF.Top + (objVSF.RowHeight(0) - fraColSel.Height) / 2 + 30
                fraColSel.ZOrder
                
                If objVSF.index = mListType.发药 Then
                    mdblSendListHeight = objVSF.Height
                    If Me.picInfo.Visible Then
                        objVSF.Width = objVSF.Width - picInfo.Width
                    End If
                    
                    With Me.picInfo
                        .Left = objVSF.Left + objVSF.Width + 50
                        .Height = objVSF.Height
                        .Top = objVSF.Top
                    End With
                    
                    fraH.Left = picInfo.Left - 20
                    fraH.Height = objVSF.Height
                    fraH.Top = objVSF.Top
                    picInfo.Top = objVSF.Top
                End If
                
                If objVSF.index = mListType.汇总 Then
                    mdblSumListHeight = objVSF.Height
                    Call ResizeChargeOffList
                End If
                Exit For
            End If
        Next
    End If
End Sub

Private Sub InitComandBars()
    Dim cbrControl As CommandBarControl
    Dim objCmdBar As CommandBar
    Dim lngCount As Integer
    
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    Me.cbsMain.VisualTheme = xtpThemeOffice2003

    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = False
'        .SetIconSize False, 24, 24
        .SetIconSize False, 16, 16
        .UseSharedImageList = False 'ImageList方式时,因同一App中共享,在AddImageList之前设置为False
    End With
    
    Me.cbsMain.EnableCustomization False
    Me.cbsMain.ActiveMenuBar.Visible = False
    Me.cbsMain.AddImageList Me.imgCheck
End Sub

Private Sub SetComandBars(ByVal intListType As Integer)
    Dim cbrControl As CommandBarControl
    Dim cbrControlSub As CommandBarControl
    Dim objCmdBar As CommandBar
    Dim lngCount As Integer
    Dim objMenu As CommandBarPopup
        
    '删除现在的工具栏及顶级菜单项
    For lngCount = cbsMain.ActiveMenuBar.Controls.count To 1 Step -1
        cbsMain.ActiveMenuBar.Controls(lngCount).Delete
    Next
    For lngCount = cbsMain.count To 2 Step -1
        cbsMain(lngCount).Delete
    Next
    
    If intListType = mListType.拒发 Or intListType = mListType.缺药 Then Exit Sub
    
    Set objCmdBar = cbsMain.Add("条件", xtpBarTop)
    objCmdBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    objCmdBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    objCmdBar.ContextMenuPresent = False
    
    '药品名称显示
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_MediPopup, "药名显示", 1, False)
    objMenu.Id = conMenu_MediPopup
    With objMenu.CommandBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Medi_CodeAddName, "药名(编码和名称)")
        cbrControl.Checked = (cbrControl.Id = 401 + mcondition.int药品名称编码显示)
        Set cbrControl = .Add(xtpControlButton, conMenu_Medi_Code, "药名(仅编码)")
        cbrControl.Checked = (cbrControl.Id = 401 + mcondition.int药品名称编码显示)
        Set cbrControl = .Add(xtpControlButton, conMenu_Medi_Name, "药名(仅名称)")
        cbrControl.Checked = (cbrControl.Id = 401 + mcondition.int药品名称编码显示)
    End With
    
    Select Case intListType
        Case mListType.发药
            '设置工具栏菜单
            Set cbrControl = objCmdBar.Controls.Add(xtpControlButton, conMenu_Tool_ShowShortage, "显示缺药药品")
            cbrControl.BeginGroup = True
            cbrControl.ToolTipText = "提示：是否显示缺药药品"
            cbrControl.Style = xtpButtonIconAndCaption
            cbrControl.IconId = IIf(mcondition.bln显示缺药, 2, 1)
            cbrControl.Checked = mcondition.bln显示缺药
            
            Set cbrControl = objCmdBar.Controls.Add(xtpControlButton, conMenu_Tool_ShowReturnSend, "显示退药待发单据")
            cbrControl.BeginGroup = True
            cbrControl.ToolTipText = "提示：是否显示退药待发单据"
            cbrControl.Style = xtpButtonIconAndCaption
            cbrControl.IconId = IIf(mcondition.bln显示退药待发单据, 2, 1)
            cbrControl.Checked = mcondition.bln显示退药待发单据
            
            Set cbrControl = objCmdBar.Controls.Add(xtpControlButton, conMenu_Tool_ShowInfo, "显示扩展信息")
            cbrControl.BeginGroup = True
            cbrControl.ToolTipText = "提示：是否显示临床诊断，抗菌药物相关信息，开单医生签名图片"
            cbrControl.Style = xtpButtonIconAndCaption
            cbrControl.IconId = IIf(mcondition.bln显示扩展信息, 2, 1)
            cbrControl.Checked = mcondition.bln显示扩展信息
            
'            Set cbrControl = objCmdBar.Controls.Add(xtpControlButton, conMenu_Tool_ShowPlug, "过敏史/病生状态")
'            cbrControl.BeginGroup = True
'            cbrControl.ToolTipText = "提示：显示过敏史/病生状态"
'            cbrControl.Style = xtpButtonIconAndCaption
'            cbrControl.IconId = 3
'            cbrControl.Visible = (mcondition.intShowPass = 1 And IsInString(gstrprivs, "合理用药监测", ";"))
            
            If Not gobjPass Is Nothing Then Call gobjPass.zlPassCommandBarAdd_YF(mlngMode, objCmdBar.Controls, conMenu_Tool_ShowPlug, 3)
            
            '设置弹出菜单
            Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_StatusPopup, "编辑(&E)", 1, False)
            objMenu.Id = conMenu_StatusPopup
            With objMenu.CommandBar.Controls
                Set cbrControl = .Add(xtpControlButton, conMenu_Status_Verify, "发药(&C)")
                Set cbrControl = .Add(xtpControlButton, conMenu_Status_Reject, "拒发(&H)")
                Set cbrControl = .Add(xtpControlButton, conMenu_Status_Shortage, "缺药(&L)")
                cbrControl.Enabled = False
                Set cbrControl = .Add(xtpControlButton, conMenu_Status_NoProcess, "不处理(&H)")
                
                Set cbrControl = .Add(xtpControlButton, conMenu_Status_AllSend, "全部发药(&S)")
                cbrControl.BeginGroup = True
                Set cbrControl = .Add(xtpControlButton, conMenu_Status_AllReject, "全部拒发(&J)")
                Set cbrControl = .Add(xtpControlButton, conMenu_Status_AllNoProcess, "全部不处理(&B)")
            End With
        Case mListType.汇总
            Set cbrControl = objCmdBar.Controls.Add(xtpControlButton, conMenu_Tool_SumByBatch, "按药品批次汇总")
            cbrControl.BeginGroup = True
            cbrControl.ToolTipText = "提示：是否按药品批次汇总"
            cbrControl.Style = xtpButtonIconAndCaption
            cbrControl.IconId = IIf(mcondition.bln按批次汇总, 2, 1)
            cbrControl.Checked = mcondition.bln按批次汇总
        Case mListType.退药
            Set cbrControl = objCmdBar.Controls.Add(xtpControlButton, conMenu_Tool_ShowAllProcess, "显示所有过程单据")
            cbrControl.BeginGroup = True
            cbrControl.ToolTipText = "提示：是否显示所有过程单据"
            cbrControl.Style = xtpButtonIconAndCaption
            cbrControl.IconId = IIf(mcondition.bln显示过程单据, 2, 1)
            cbrControl.Checked = mcondition.bln显示过程单据
            
'            Set cbrControl = objCmdBar.Controls.Add(xtpControlButton, conMenu_Tool_ShowPlug, "过敏史/病生状态")
'            cbrControl.BeginGroup = True
'            cbrControl.ToolTipText = "提示：显示过敏史/病生状态"
'            cbrControl.Style = xtpButtonIconAndCaption
'            cbrControl.IconId = 3
'            cbrControl.Visible = (mcondition.intShowPass = 1 And IsInString(gstrprivs, "合理用药监测", ";"))
            If Not gobjPass Is Nothing Then Call gobjPass.zlPassCommandBarAdd_YF(mlngMode, objCmdBar.Controls, conMenu_Tool_ShowPlug, 3)
            
            '设置弹出菜单
            Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_StatusPopup, "编辑(&E)", 1, False)
            objMenu.Id = conMenu_StatusPopup
            With objMenu.CommandBar.Controls
                Set cbrControl = .Add(xtpControlButton, conMenu_Status_AllReturn, "全部退药(&R)")
                Set cbrControl = .Add(xtpControlButton, conMenu_Status_AllCancel, "全部取消(&C)")
            End With
    End Select
    
    
    Select Case intListType
        Case mListType.发药, mListType.退药
            '设置弹出菜单，PASS
            Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, mconMenu_PASS, "PASS（&P)", 1, False)
            objMenu.Id = mconMenu_PASS
'            If Not gobjPass Is Nothing Then Call gobjPass.zlPassCommandBarAdd_YF(mlngMode, cbsMain, mconMenu_PASS, 1)
    End Select
End Sub


Private Sub chk所有诊断_Click()
    If mcondition.intListType <> mListType.发药 Then Exit Sub
    If vsfList(mcondition.intListType).Row = 0 Then Exit Sub
    If vsfList(mListType.发药).IsSubtotal(vsfList(mListType.发药).Row) = True Then Exit Sub
    
    vsf诊断.Tag = ""
    Call GetDiagnosis(vsfList(mcondition.intListType).Row)
End Sub

Private Sub Form_Load()
    With mcondition
        If Val(zlDataBase.GetPara("使用个性化风格")) = 1 Then
            .bln显示缺药 = (Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\发药", "显示缺药药品", "1")) = 1)
            .bln显示退药待发单据 = (Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\发药", "显示退药待发单据", "1")) = 1)
            .bln显示扩展信息 = (Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\发药", "显示相关信息", "1")) = 1)
            .bln按批次汇总 = (Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\汇总", "按批次汇总", "0")) = 1)
            .bln显示过程单据 = (Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\退药", "显示过程单据", "0")) = 1)
            .bln显示所有诊断 = (Val(GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\发药", "显示所有诊断", "0")) = 1)
        Else
            .bln显示缺药 = True
            .bln按批次汇总 = False
            .bln显示扩展信息 = True
            .bln显示过程单据 = False
            .bln显示退药待发单据 = True
            .bln显示所有诊断 = False
        End If
        
        .intShowPass = gintPass
        .bln医生查询 = IsInString(gstrprivs, "医生查询", ";")
        
'        .bln医生查询 = False
        
        '用于测试PASS
'        .blnShowPass = True

        .bln修改留存数量 = IsInString(gstrprivs, "修改留存数量", ";")
    End With
    
    Call SetParams
    
    Call Load发药单格式
    
    Call InitComandBars
    Call InitList(-1)
    
    vsfChargeOff.Visible = False
    Me.txt用药理由.Text = ""
    
    Me.chk所有诊断.Value = IIf(mcondition.bln显示所有诊断 = True, 1, 0)
End Sub


Private Sub Form_Resize()
    Dim objVSF As VSFlexGrid
    Dim i As Integer
    
    On Error Resume Next
    
    With picAssist
        If .Visible Then
            .Left = 0
            .Top = Me.Height - .Height
            .Width = Me.Width - 50
        End If
    End With
    
    If cbsMain.count = 1 Then
        For Each objVSF In vsfList
            If objVSF.Visible = True Then
                objVSF.Move 0, 0, Me.Width, IIf(picAssist.Visible, Me.Height - picAssist.Height - 50, Me.Height)
                
                fraColSel.Left = objVSF.Left + objVSF.ColWidth(0) - fraColSel.Width - 50
                fraColSel.Top = objVSF.Top + (objVSF.RowHeight(0) - fraColSel.Height) / 2 + 30
                fraColSel.ZOrder
                
                If objVSF.index = mListType.发药 Then
                    mdblSendListHeight = objVSF.Height
                    objVSF.Width = objVSF.Width - picInfo.Width
                    
                    With Me.picInfo
                        .Left = objVSF.Left + objVSF.Width + 50
                        .Height = objVSF.Height
                        .Top = objVSF.Top
                    End With
                    fraH.Left = picInfo.Left - 20
                    fraH.Height = objVSF.Height
                    fraH.Top = objVSF.Top
                    picInfo.Top = objVSF.Top
                End If
                
                If objVSF.index = mListType.汇总 Then
                    mdblSumListHeight = objVSF.Height
                    Call ResizeChargeOffList
                End If
                Exit For
            End If
        Next
    End If
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetTempOperate mcondition.intListType, mcondition.intListType
'    SaveListColState
    
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\药品部门发药管理", "药品名称显示方式", mcondition.int药品名称编码显示)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\药品部门发药管理", "发药单格式", cbo发药单格式.ListIndex)
    Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "部门发药管理\发药", "显示所有诊断", chk所有诊断.Value)
End Sub

Private Sub imgColSel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim i As Long
    
    If Button = 1 Then '列选择器
        '根据当前状态直接确定勾选状态
        With vsfColSel
            If .Visible Then
                .Visible = False
                vsfList(Val(.Tag)).SetFocus
            Else
                For i = .FixedRows To .rows - 1
                    If vsfList(Val(.Tag)).ColHidden(.RowData(i)) Or vsfList(Val(.Tag)).ColWidth(.RowData(i)) = 0 Then
                        .TextMatrix(i, 0) = 0
                    Else
                        .TextMatrix(i, 0) = 1
                    End If
                Next
                
                .Height = .RowHeightMin * .rows + 150
                .Top = fraColSel.Top + fraColSel.Height
                
                If .Top + .Height > Me.ScaleHeight - vsfList(Val(.Tag)).Top Then
                    .Height = Me.ScaleHeight - .Top - vsfList(Val(.Tag)).Top
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

Private Sub picInfo_Resize()
    picHscSend.Width = picInfo.Width
    Pic用药理由.Width = picInfo.Width
    picDoctor.Width = picInfo.Width
    Pic用药理由.Top = (1 / 3) * Me.picInfo.Height
    picDoctor.Top = (2 / 3) * Me.picInfo.Height
    
    vsf诊断.Height = Pic用药理由.Top - picHscSend.Top - picHscSend.Height
    vsf诊断.Top = Me.picHscSend.Height
    vsf诊断.Width = picInfo.Width
    
    txt用药理由.Top = Pic用药理由.Top + Pic用药理由.Height
    txt用药理由.Height = Me.picDoctor.Top - Pic用药理由.Top - Pic用药理由.Height
    txt用药理由.Width = picInfo.Width
    
    pic签名图片.Top = picDoctor.Height + picDoctor.Top + 20
    pic签名图片.Left = (picInfo.Width / 2) - (pic签名图片.Width / 2)
End Sub

Private Sub fraH_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        If vsfList(mListType.发药).Width + x <= 1200 Then Exit Sub
        If picInfo.Width - x < 1200 Then Exit Sub

        fraH.Left = fraH.Left + x
        picInfo.Left = picInfo.Left + x
        picInfo.Width = picInfo.Width - x
        vsfList(mListType.发药).Width = vsfList(mListType.发药).Width + x
        
        Me.Refresh
    End If
End Sub

Private Sub picHscSend_Resize()
    On Error Resume Next
    
    With lblDiag
        .Left = (picHscSend.Width - .Width) / 2
    End With
    
    With chk所有诊断
        .Left = picHscSend.Width - .Width - 50
    End With
End Sub


Private Sub Pic用药理由_Resize()
    With lbl用药理由
        .Left = (Pic用药理由.Width - .Width) / 2
    End With
End Sub

Private Sub picDoctor_Resize()
    With lblDoctor
        .Left = (picDoctor.Width - .Width) / 2
    End With
End Sub

Private Sub picAssist_Resize()
    On Error Resume Next
    
    With Lbl配药人
    End With
    
    With cbo配药人
    End With
    
    
    
    With cbo核查人
        .Left = (picAssist.Width - .Width + 400) / 2
    End With
    
    With lbl核查人
        .Left = cbo核查人.Left - 50 - .Width
    End With
    
    With cbo发药单格式
        .Left = picAssist.Width - .Width - 50
    End With
    
    With lbl发药单格式
        .Left = cbo发药单格式.Left - 50 - .Width
    End With
End Sub

Private Sub vsfChargeOff_EnterCell()
    With vsfChargeOff
        If .Row = 0 Then Exit Sub
        .Cell(flexcpPicture, 1, 0, .rows - 1, 0) = Nothing
        .Cell(flexcpPicture, .Row, 0, .Row, 0) = frmPublic.ImgList.ListImages(2).Picture

    End With
End Sub


Private Sub vsfColSel_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim lngCol As Long
    
    If Col = 0 Then
        lngCol = vsfColSel.RowData(Row)
        If Val(vsfColSel.TextMatrix(Row, 0)) <> 0 Then
            vsfList(Val(vsfColSel.Tag)).ColWidth(lngCol) = vsfList(Val(vsfColSel.Tag)).ColData(lngCol)
            vsfList(Val(vsfColSel.Tag)).ColHidden(lngCol) = False
        Else
'            vsfList(Val(vsfColSel.Tag)).ColWidth(lngCol) = 0
            vsfList(Val(vsfColSel.Tag)).ColHidden(lngCol) = True
        End If
    End If
    
    SaveListColState Val(vsfColSel.Tag)
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

Private Sub vsfList_AfterEdit(index As Integer, ByVal Row As Long, ByVal Col As Long)
    Dim i As Integer
    Dim j As Integer
    Dim intRow As Integer
    Dim dblMoney As Double
    Dim dblCurMoney As Double
    Dim strCont As String
    
    With vsfList(index)
        Select Case index
            Case mListType.退药
                If Row = 0 Then Exit Sub
                If Val(.TextMatrix(Row, mIntCol退药_收发ID)) = 0 Then Exit Sub
                If Val(.TextMatrix(Row, mIntCol退药_执行状态)) <> mState.退药 And Val(.TextMatrix(Row, mIntCol退药_执行状态)) <> mState.退药_原始记录 Then Exit Sub
                If Val(.TextMatrix(Row, mIntCol退药_准退数)) = 0 Then Exit Sub
                If Col = mIntCol退药_退药数 Then
                    If Val(.TextMatrix(Row, mIntCol退药_准退数)) >= 0 Then
                        If Val(.TextMatrix(Row, mIntCol退药_退药数)) > Val(.TextMatrix(Row, mIntCol退药_准退数)) Or Val(.TextMatrix(Row, mIntCol退药_退药数)) < 0 Then
                            .TextMatrix(Row, mIntCol退药_退药数) = Val(.TextMatrix(Row, mIntCol退药_准退数))
                        End If
                    Else
                        If Val(.TextMatrix(Row, mIntCol退药_退药数)) < Val(.TextMatrix(Row, mIntCol退药_准退数)) Or Val(.TextMatrix(Row, mIntCol退药_退药数)) >= 0 Then
                            .TextMatrix(Row, mIntCol退药_退药数) = Val(.TextMatrix(Row, mIntCol退药_准退数))
                        End If
                    End If
                    
                    If Val(.TextMatrix(Row, mIntCol退药_退药数)) = 0 Then
                        .TextMatrix(Row, mIntCol退药_退药数) = ""
                        
                        If Val(.TextMatrix(Row, mIntCol退药_执行状态)) = mState.退药_原始记录 Then
                            Exit Sub
                        ElseIf Val(.TextMatrix(Row, mIntCol退药_执行状态)) = mState.退药 Then
                            mrsReturnList.Filter = "收发ID=" & Val(.TextMatrix(Row, mIntCol退药_收发ID))
                            
                            .TextMatrix(Row, mIntCol退药_状态) = "不处理"
                            .TextMatrix(Row, mIntCol退药_执行状态) = mState.退药_原始记录
                            mrsReturnList!状态 = .TextMatrix(Row, mIntCol退药_状态)
                            mrsReturnList!执行状态 = Val(.TextMatrix(Row, mIntCol退药_执行状态))
                            mrsReturnList!退药数 = Val(.TextMatrix(Row, mIntCol退药_退药数))
                            mrsReturnList.Update
                        End If
                    Else
                        mrsReturnList.Filter = "收发ID=" & Val(.TextMatrix(Row, mIntCol退药_收发ID))
                        
                        .TextMatrix(Row, mIntCol退药_状态) = "退药"
                        .TextMatrix(Row, mIntCol退药_执行状态) = mState.退药
                        mrsReturnList!状态 = .TextMatrix(Row, mIntCol退药_状态)
                        mrsReturnList!执行状态 = Val(.TextMatrix(Row, mIntCol退药_执行状态))
                        mrsReturnList!退药数 = Val(.TextMatrix(Row, mIntCol退药_退药数))
                        mrsReturnList.Update
                    End If
                End If
            Case mListType.汇总
                If Row = 0 Then Exit Sub
                If mcondition.bln按科室汇总 = False Or mcondition.bln修改留存数量 = False Then Exit Sub
                If Val(.TextMatrix(Row, mIntCol科室汇总_领药部门id)) = 0 Then Exit Sub
                
                Dim dbl留存数 As Double
                Dim dbl应发数 As Double
                Dim dbl实发数 As Double
                
                dbl应发数 = Val(.TextMatrix(Row, mIntCol科室汇总_应发数量)) - Val(.TextMatrix(Row, mIntCol科室汇总_销帐数量))
                
                If Col = mIntCol科室汇总_实发数量 Then
                    dbl实发数 = Val(.TextMatrix(Row, mIntCol科室汇总_留存数量))
                    If dbl实发数 > dbl应发数 Or dbl实发数 < 0 Then
                        .TextMatrix(Row, mIntCol科室汇总_实发数量) = zlStr.FormatEx(dbl应发数, 5)
                        .TextMatrix(Row, mIntCol科室汇总_留存数量) = 0
                    Else
                        .TextMatrix(Row, mIntCol科室汇总_实发数量) = zlStr.FormatEx(dbl实发数, 5)
                        .TextMatrix(Row, mIntCol科室汇总_留存数量) = zlStr.FormatEx(dbl应发数 - Val(.TextMatrix(Row, mIntCol科室汇总_实发数量)), 5)
                    End If
                ElseIf Col = mIntCol科室汇总_留存数量 Then
                    dbl留存数 = Val(.TextMatrix(Row, mIntCol科室汇总_留存数量))
                    If dbl留存数 > dbl应发数 Or dbl留存数 < 0 Then
                        .TextMatrix(Row, mIntCol科室汇总_实发数量) = zlStr.FormatEx(dbl应发数, 5)
                        .TextMatrix(Row, mIntCol科室汇总_留存数量) = 0
                    Else
                        .TextMatrix(Row, mIntCol科室汇总_实发数量) = zlStr.FormatEx(dbl应发数 - Val(.TextMatrix(Row, mIntCol科室汇总_留存数量)), 5)
                        .TextMatrix(Row, mIntCol科室汇总_留存数量) = zlStr.FormatEx(dbl留存数, 5)
                    End If
                End If
                
                .TextMatrix(Row, mIntCol科室汇总_实发金额) = zlStr.FormatEx(Val(.TextMatrix(Row, mIntCol科室汇总_应发金额)) / Val(.TextMatrix(Row, mIntCol科室汇总_应发数量)) * Val(.TextMatrix(Row, mIntCol科室汇总_实发数量)), mintMoneyDigit, , True)
                        
                If mcondition.bln按科室汇总 = True Then
                    mrsSendList.Filter = "领药部门id=" & Val(.TextMatrix(Row, mIntCol科室汇总_领药部门id)) & " and 药品id=" & Val(.TextMatrix(Row, mIntCol科室汇总_药品ID))
                Else
                    mrsSendList.Filter = "药品id=" & Val(.TextMatrix(Row, mIntCol科室汇总_药品ID)) & "And 批次=" & Val(.TextMatrix(Row, mIntCol科室汇总_批次))
                End If
                
                mrsSendList!留存数量 = .TextMatrix(Row, mIntCol科室汇总_留存数量)
                
                mrsSendList.Update
                        
                DoEvents
                
                .Row = Row
                .Col = mIntCol科室汇总_实发数量
                If Val(.TextMatrix(Row, mIntCol科室汇总_实发数量)) < 0 Then
                    .CellForeColor = vbRed
                ElseIf Val(.TextMatrix(Row, mIntCol科室汇总_实发数量)) > 0 Then
                    .CellForeColor = vbBlue
                End If
                
                '获取下一个合计行，获取上一个合计行，然后重新统计实发金额
                For i = Row To .rows - 1
                    If Not IsNumeric(.TextMatrix(i, mIntCol科室汇总_应发金额)) Then
                        Exit For
                    End If
                Next
                
                
                For j = Row To 1 Step -1
                    If Not IsNumeric(.TextMatrix(j, mIntCol科室汇总_实发金额)) Then
                        Exit For
                    End If
                Next
                
                For intRow = j + 1 To i - 1
                    dblMoney = dblMoney + CDbl(.TextMatrix(intRow, mIntCol科室汇总_实发金额))
                Next
                
                '获取原始的实际金额统计值
                strCont = .TextMatrix(i, 1)
                dblCurMoney = Mid(Mid(strCont, InStr(1, strCont, "实发金额共计") + 6), 1, InStr(1, Mid(strCont, InStr(1, strCont, "实发金额共计") + 6), "元") - 1)
                
                '修改当前行统计金额
                .Cell(flexcpText, i, 1, i, .Cols - 1) = Mid(strCont, 1, InStr(1, strCont, "实发金额共计") + 5) & Format(dblMoney, "#####0.00;-#####0.00;0.00;") & "元"
                
                '获取本次编辑差价
                dblMoney = dblCurMoney - dblMoney
                
                If i <> .rows - 1 Then
                  '获取原始合计金额
                  strCont = .TextMatrix(.rows - 1, 1)
                  dblCurMoney = Mid(Mid(strCont, InStr(1, strCont, "实发金额共计") + 6), 1, InStr(1, Mid(strCont, InStr(1, strCont, "实发金额共计") + 6), "元") - 1)
                
                  
                  '修改合计统计金额
                  .Cell(flexcpText, .rows - 1, 1, .rows - 1, .Cols - 1) = Mid(strCont, 1, InStr(1, strCont, "实发金额共计") + 5) & Format(dblCurMoney - dblMoney, "#####0.00;-#####0.00;0.00;") & "元"
                End If
        End Select
    End With
    
End Sub

Private Sub vsfList_AfterMoveColumn(index As Integer, ByVal Col As Long, Position As Long)
    Dim i As Integer
    
    '重设列选择列表
    Call InitColSelList(index)
    
    '重设列顺序号
    For i = 0 To vsfList(index).Cols - 1
        Call SetColumnValue(index, vsfList(index).TextMatrix(0, i), i)
    Next
    
    '保存页面的列设置
    SaveListColState index
End Sub

Private Sub vsfList_AfterSort(index As Integer, ByVal Col As Long, Order As Integer)
    If vsfList(index).rows > 1 Then
        If vsfList(index).TextMatrix(1, vsfList(index).ColIndex("药品名称")) <> "" Then
            If index = mListType.发药 Then
                SetSubTotal vsfList(index), vsfList(index).TextMatrix(0, Col)
                SetGroup vsfList(index), Col = mIntCol发药_NO
            ElseIf index = mListType.汇总 Then
                SetSubTotal vsfList(index), vsfList(index).TextMatrix(0, Col)
            ElseIf index = mListType.退药 Then
                SetGroup vsfList(index), Col = mIntCol退药_NO And mcondition.bln显示过程单据 = False
            End If
        End If
        
        If Val(zlDataBase.GetPara("使用个性化风格")) = 1 Then
            '保存处方清单的用户排序规则
            '保存规则：
            '子项＝列表类型
            '值=列号|升/降序
            Call SaveSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\" & "药品部门发药", "处方清单排序" & index, Col & "|" & Order)
        End If
    End If
End Sub

Private Sub vsfList_AfterUserResize(index As Integer, ByVal Row As Long, ByVal Col As Long)
    If Val(zlDataBase.GetPara("使用个性化风格")) = 1 Then
        '保存页面的列设置
        SaveListColState index
    End If
    
    If index = mListType.发药 Then
        If Col = mIntCol发药_皮试结果 Then
            With vsfList(index)
                If .ColWidth(mIntCol发药_皮试结果) > 800 Then
                    .ColWidth(mIntCol发药_品名) = .ColWidth(mIntCol发药_品名) + (.ColWidth(mIntCol发药_皮试结果) - 800)
                    .ColWidth(mIntCol发药_皮试结果) = 800
                Else
                    .ColWidth(mIntCol发药_品名) = .ColWidth(mIntCol发药_品名) - (800 - .ColWidth(mIntCol发药_皮试结果))
                    .ColWidth(mIntCol发药_皮试结果) = 800
                End If
            End With
        End If
    End If
End Sub


Private Sub vsfList_BeforeMoveColumn(index As Integer, ByVal Col As Long, Position As Long)
    '设置不能移动的列
    Select Case index
        Case mListType.发药
            If Col = mIntCol发药_审查结果 Then
                Position = mIntCol发药_审查结果
            End If
            
            If Col = mIntCol发药_分组符 Then
                Position = mIntCol发药_分组符
            End If
            
            If Col = mIntCol发药_品名 Then
                Position = mIntCol发药_品名
            End If
            
            If Col = mIntCol发药_皮试结果 Then
                Position = mIntCol发药_皮试结果
            End If
            
            If (Col <> mIntCol发药_品名 And Position = mIntCol发药_品名) Or (Col <> mIntCol发药_皮试结果 And Position = mIntCol发药_皮试结果) Or (Col <> mIntCol发药_审查结果 And Position = mIntCol发药_审查结果) Or (Col <> mIntCol发药_分组符 And Position = mIntCol发药_分组符) Then
                Position = Col
            End If
    End Select

End Sub

Private Sub vsfList_BeforeSort(index As Integer, ByVal Col As Long, Order As Integer)
    If index <> mListType.发药 And index <> mListType.汇总 Then Exit Sub
    
    With vsfList(index)
        .Subtotal flexSTClear
    End With
End Sub

Private Sub vsfList_BeforeUserResize(index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '设置不能调整列宽的列
    Select Case index
        Case mListType.发药
            If Col = mIntCol发药_当前行 Or Col = mIntCol发药_分组符 Or Col = mIntCol发药_审查结果 _
                Or Col = mIntCol发药_状态 Then Cancel = True
        Case mListType.退药
            If Col = mIntCol退药_当前行 Or Col = mIntCol退药_分组符 Or Col = mIntCol退药_审查结果 Then Cancel = True
        Case Else
            If Col = 0 Then Cancel = True
    End Select
End Sub

Private Sub vsfList_DblClick(index As Integer)
    Dim lngColor As Long
    Dim i As Long
    Dim strNo As String
    Dim lng相关ID As Long
    Dim intState As Integer
    Dim strState As String
    
    With vsfList(index)
        Select Case index
            Case mListType.发药
                If .Row = 0 Then Exit Sub
                If Val(.TextMatrix(.Row, mIntCol发药_收发ID)) = 0 Then Exit Sub
                
                If .Col <> mIntCol发药_审查结果 Then
                    If Val(.TextMatrix(.Row, mIntCol发药_执行状态)) = mState.缺药 Then Exit Sub
                    
                    If Val(.TextMatrix(.Row, mIntCol发药_执行状态)) = mState.发药 Then
                        .TextMatrix(.Row, mIntCol发药_状态) = "拒发"
                        .TextMatrix(.Row, mIntCol发药_执行状态) = mState.拒发
                        lngColor = mListColor.State_Reject
                    ElseIf Val(.TextMatrix(.Row, mIntCol发药_执行状态)) = mState.拒发 Then
                        .TextMatrix(.Row, mIntCol发药_状态) = "不处理"
                        .TextMatrix(.Row, mIntCol发药_执行状态) = mState.不处理
                        lngColor = mListColor.State_UnProcess
                    ElseIf Val(.TextMatrix(.Row, mIntCol发药_执行状态)) = mState.不处理 Then
                        If mcondition.bln允许未审核处方发药 = False And Val(.TextMatrix(.Row, mIntCol发药_已收费)) = 0 Then
                            .TextMatrix(.Row, mIntCol发药_状态) = "拒发"
                            .TextMatrix(.Row, mIntCol发药_执行状态) = mState.拒发
                            lngColor = mListColor.State_Reject
                        Else
                            .TextMatrix(.Row, mIntCol发药_状态) = "发药"
                            .TextMatrix(.Row, mIntCol发药_执行状态) = mState.发药
                            lngColor = mListColor.State_Send
                        End If
                    End If
                    
                    intState = Val(.TextMatrix(.Row, mIntCol发药_执行状态))
                    strState = .TextMatrix(.Row, mIntCol发药_状态)
                    
                    .Cell(flexcpBackColor, .Row, 1, .Row, .Cols - 1) = lngColor
                    
                    mrsSendList.Filter = "收发ID=" & Val(.TextMatrix(.Row, mIntCol发药_收发ID))
                    
                    mrsSendList!状态 = .TextMatrix(.Row, mIntCol发药_状态)
                    mrsSendList!执行状态 = Val(.TextMatrix(.Row, mIntCol发药_执行状态))
                    mrsSendList.Update
                    
                    mblnSendChange = True
                
                    '同组医嘱的药品状态都要同步改变；如果是高危药品，并且要求单独发放时不同步改变
                    If mcondition.bln配制中心 And InStr(1, mcondition.str高危发放, .TextMatrix(.Row, mIntCol发药_高危药品)) = 0 Then
                        strNo = .TextMatrix(.Row, mIntCol发药_NO)
                        lng相关ID = Val(.TextMatrix(.Row, mIntCol发药_相关ID))
                        If lng相关ID > 0 Then
                            For i = 1 To .rows - 1
                                If .IsSubtotal(i) = False And .TextMatrix(i, mIntCol发药_NO) = strNo And Val(.TextMatrix(i, mIntCol发药_相关ID)) = lng相关ID _
                                    And Val(.TextMatrix(i, mIntCol发药_执行状态)) <> intState And Val(.TextMatrix(i, mIntCol发药_执行状态)) <> mState.缺药 _
                                    And InStr(1, mcondition.str高危发放, .TextMatrix(i, mIntCol发药_高危药品)) = 0 Then
                                    .TextMatrix(i, mIntCol发药_执行状态) = intState
                                    .TextMatrix(i, mIntCol发药_状态) = strState
                                    
                                    .Cell(flexcpBackColor, i, 1, i, .Cols - 1) = lngColor
                                    
                                    mrsSendList.Filter = "收发ID=" & Val(.TextMatrix(i, mIntCol发药_收发ID))
                                    
                                    mrsSendList!执行状态 = intState
                                    mrsSendList!状态 = strState
                                    
                                    mrsSendList.Update
                                End If
                            Next
                        End If
                    End If
                    
                    SetMainComandBars index, .Row
                    DoEvents
                    Call RefreshList_Sum
                Else
                    If mcondition.intShowPass = 3 And Not gobjPass Is Nothing And IsInString(gstrprivs, "合理用药监测", ";") Then
                        Call gobjPass.queryCheckResult(.TextMatrix(.Row, mIntCol发药_住院号), "2")
                    End If
                End If
            Case mListType.拒发
                If .Row = 0 Then Exit Sub
                If Val(.TextMatrix(.Row, mIntCol拒发_收发ID)) = 0 Then Exit Sub
                If .Col = mIntCol拒发_状态 Then
                    If Val(.TextMatrix(.Row, mIntCol拒发_执行状态)) = mState.拒发_不处理 Then
                        .TextMatrix(.Row, .Col) = "恢复"
                        .TextMatrix(.Row, mIntCol拒发_执行状态) = mState.拒发_恢复
                        lngColor = mListColor.State_RejectRestore
                    ElseIf Val(.TextMatrix(.Row, mIntCol拒发_执行状态)) = mState.拒发_恢复 Then
                        .TextMatrix(.Row, .Col) = "不处理"
                        .TextMatrix(.Row, mIntCol拒发_执行状态) = mState.拒发_不处理
                        lngColor = mListColor.State_RejectUnProcess
                    End If
                    
                    .Cell(flexcpBackColor, .Row, 1, .Row, .Cols - 1) = lngColor
                    
                    mrsSendList.Filter = "收发ID=" & Val(.TextMatrix(.Row, mIntCol拒发_收发ID))
                    
                    mrsSendList!状态 = .TextMatrix(.Row, mIntCol拒发_状态)
                    mrsSendList!执行状态 = Val(.TextMatrix(.Row, mIntCol拒发_执行状态))
                    mrsSendList.Update
                    
                    SetMainComandBars index, .Row
                End If
            Case mListType.退药
                If .Row = 0 Then Exit Sub
                If Val(.TextMatrix(.Row, mIntCol退药_收发ID)) = 0 Then Exit Sub
                If Val(.TextMatrix(.Row, mIntCol退药_准退数)) = 0 Then Exit Sub
                If .Col = mIntCol退药_状态 Then
                    If Val(.TextMatrix(.Row, mIntCol退药_执行状态)) = mState.退药_原始记录 Then
                        .TextMatrix(.Row, .Col) = "退药"
                        .TextMatrix(.Row, mIntCol退药_退药数) = .TextMatrix(.Row, mIntCol退药_准退数)
                        .TextMatrix(.Row, mIntCol退药_执行状态) = mState.退药
                    ElseIf Val(.TextMatrix(.Row, mIntCol退药_执行状态)) = mState.退药 Then
                        .TextMatrix(.Row, .Col) = "不处理"
                        .TextMatrix(.Row, mIntCol退药_退药数) = ""
                        .TextMatrix(.Row, mIntCol退药_执行状态) = mState.退药_原始记录
                    End If
                    
                    mrsReturnList.Filter = "收发ID=" & Val(.TextMatrix(.Row, mIntCol退药_收发ID))
                    
                    mrsReturnList!状态 = .TextMatrix(.Row, mIntCol退药_状态)
                    mrsReturnList!执行状态 = Val(.TextMatrix(.Row, mIntCol退药_执行状态))
                    mrsReturnList.Update
                End If
        End Select
    End With
End Sub

Private Sub vsfList_DrawCell(index As Integer, ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    '说明：1.OwnerDraw要设置为Over(画出单元所有内容)
    '      2.Cell的GridLine从上下左右向内都是从第1根线开始
    '      3.Cell的Border从左上是从第2根线开始,右下是从第1根线开始
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    Dim lngPen As Long, lngPenSel As Long
    Dim lngBrush As Long, lngBrushSel As Long
    
    Dim lngStateColor As Long
    
    If index <> mListType.发药 Then Exit Sub
    
    With vsfList(index)
        If Col = mIntCol发药_品名 And .IsSubtotal(Row) = False Then
            '擦除单元格右边表格线
            vRect.Left = Right - 1
            vRect.Top = Top
            vRect.Right = Right
            vRect.Bottom = Bottom - 1
            
            If Row = 0 Then
                SetBkColor hDC, SysColor2RGB(.BackColorFixed)
            Else
                If .IsSelected(Row) = True Then
                    lngStateColor = .BackColorSel
                Else
                    '设置状态的颜色
                    If .TextMatrix(Row, mIntCol发药_执行状态) <> "" Then
                        If .TextMatrix(Row, mIntCol发药_执行状态) = mState.缺药 Then
                            lngStateColor = mListColor.State_Shortage
                        ElseIf .TextMatrix(Row, mIntCol发药_执行状态) = mState.发药 Then
                            lngStateColor = mListColor.State_Send
                        ElseIf .TextMatrix(Row, mIntCol发药_执行状态) = mState.拒发 Then
                            lngStateColor = mListColor.State_Reject
                        ElseIf .TextMatrix(Row, mIntCol发药_执行状态) = mState.不处理 Then
                            lngStateColor = mListColor.State_UnProcess
                        End If
                    Else
                        lngStateColor = .BackColorSel
                    End If
                End If
                
                SetBkColor hDC, SysColor2RGB(lngStateColor)
            End If
            
            ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        End If
    End With
End Sub

Private Sub vsfList_EnterCell(index As Integer)
    Dim lng开单人id As Long
    Dim strTempFile As String
    
    If mblnOutPut = True Then Exit Sub
    If mblnRefresh = True Then Exit Sub
    
    vsfChargeOff.Visible = False
    
    With vsfList(index)
        
        If .Row = 0 Then Exit Sub
        
        .Cell(flexcpPicture, 1, 0, .rows - 1, 0) = Nothing
        .Cell(flexcpPicture, .Row, 0, .Row, 0) = frmPublic.ImgList.ListImages(2).Picture
        
        .Editable = flexEDNone
        
        Select Case index
            Case mListType.发药
                If mblnShowReject = False Then
                    If vsfList(mListType.发药).IsSubtotal(.Row) = False And (.TextMatrix(.Row, mIntCol发药_用药目的) <> "" Or .TextMatrix(.Row, mIntCol发药_用药理由) <> "") Then
                        Me.txt用药理由.Text = "用药目的：" & .TextMatrix(.Row, mIntCol发药_用药目的) & vbCrLf & "用药理由：" & .TextMatrix(.Row, mIntCol发药_用药理由)
                    Else
                        Me.txt用药理由.Text = ""
                    End If
                    
                    '显示病人诊断
                    Call GetDiagnosis(.Row)
                    
                    '获取开单人的签名图片
                    pic签名图片.Picture = Nothing
                    pic签名图片.Visible = False
                    If vsfList(mListType.发药).IsSubtotal(.Row) = False And IsNumeric(.TextMatrix(.Row, mIntCol发药_收发ID)) Then
                        lng开单人id = get开单人id(.TextMatrix(.Row, mIntCol发药_开单医生), Val(.TextMatrix(.Row, mIntCol发药_开单部门id)))
                        strTempFile = Sys.ReadLob(100, 15, lng开单人id)
                        If strTempFile <> "" Then
                            picSign.Picture = LoadPicture(strTempFile)
                            If Not picSign.Picture Is Nothing Then
                                pic签名图片.Visible = True
                                pic签名图片.PaintPicture picSign.Picture, 0, 0, pic签名图片.ScaleX(pic签名图片.Width, vbTwips, vbPixels), pic签名图片.ScaleY(pic签名图片.Height, vbTwips, vbPixels)
                            End If
                            Kill strTempFile
                        End If
                    End If
                    
                    If Not gobjPass Is Nothing Then Call gobjPass.zlPassSetDrug_YF(.TextMatrix(.Row, mIntCol发药_药品ID), .TextMatrix(.Row, mIntCol发药_药品名称))
                    If Not gobjPass Is Nothing Then Call gobjPass.zlPassClearLight_YF
                End If
            Case mListType.汇总
                If .Row = 0 Then Exit Sub
                
                If mcondition.bln按科室汇总 = False Then Exit Sub
                If Val(.TextMatrix(.Row, mIntCol科室汇总_领药部门id)) = 0 Then Exit Sub
                
                If .Col = mIntCol科室汇总_留存数量 Or .Col = mIntCol科室汇总_留存数量 Then
                    If mcondition.bln修改留存数量 = True Then .Editable = flexEDKbdMouse
                End If
                

                vsfList(mListType.汇总).Height = mdblSumListHeight
                
                If RefreshList_ChargeOff(Val(.TextMatrix(.Row, mIntCol科室汇总_领药部门id)), Val(.TextMatrix(.Row, mIntCol科室汇总_药品ID))) = True Then
                    vsfChargeOff.Visible = True
                    picHsc.Visible = True
                    
                    Call ResizeChargeOffList
                End If
            Case mListType.退药
                If .Row = 0 Then Exit Sub
                SetMainComandBars index, .Row
                If Val(.TextMatrix(.Row, mIntCol退药_收发ID)) = 0 Then Exit Sub
                
                '设置PASS按钮状态
                SetPassMenuButton index, .Row
                
                If Val(.TextMatrix(.Row, mIntCol退药_执行状态)) <> mState.退药 And Val(.TextMatrix(.Row, mIntCol退药_执行状态)) <> mState.退药_原始记录 Then Exit Sub
                If Val(.TextMatrix(.Row, mIntCol退药_准退数)) = 0 Then Exit Sub
                Select Case .Col
                    Case mIntCol退药_退药数
                        .Editable = flexEDKbdMouse
                        
                        If Val(.TextMatrix(.Row, mIntCol退药_退药数)) = 0 Then
                            .TextMatrix(.Row, mIntCol退药_退药数) = .TextMatrix(.Row, mIntCol退药_准退数)
                            .TextMatrix(.Row, mIntCol退药_状态) = "退药"
                            .TextMatrix(.Row, mIntCol退药_执行状态) = mState.退药
                            
                            mrsReturnList.Filter = "收发ID=" & Val(.TextMatrix(.Row, mIntCol退药_收发ID))
                            mrsReturnList!状态 = .TextMatrix(.Row, mIntCol退药_状态)
                            mrsReturnList!执行状态 = Val(.TextMatrix(.Row, mIntCol退药_执行状态))
                            mrsReturnList!退药数 = Val(.TextMatrix(.Row, mIntCol退药_退药数))
                            mrsReturnList.Update
                        End If
                End Select
            End Select
            
            SetMainComandBars index, .Row
    End With
End Sub

Private Function get开单人id(ByVal str姓名 As String, ByVal lng开单部门id As Long) As Long
    Dim strsql As String
    Dim rstemp As Recordset
    
    On Error GoTo errHandle
    
    strsql = "select A.ID from 人员表 A,部门人员 B where A.id=B.人员id and A.姓名=[1] and B.部门id=[2]"
    Set rstemp = zlDataBase.OpenSQLRecord(strsql, Me.Caption, str姓名, lng开单部门id)
    
    If Not rstemp.EOF Then
        get开单人id = rstemp!Id
    End If
    Exit Function
errHandle:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub vsfList_KeyPressEdit(index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strKey As String
    
    With vsfList(index)
        Select Case index
            Case mListType.退药
                If Col = mIntCol退药_退药数 Then
                    strKey = .EditText
                    If InStr("1234567890" + Chr(46) + Chr(8) + Chr(13) + Chr(Asc("-")), Chr(KeyAscii)) = 0 Then
                        KeyAscii = 0
                    ElseIf KeyAscii = Asc(".") Or (KeyAscii >= Asc("0") And KeyAscii <= Asc("9")) Then
                        If .EditSelLength = Len(strKey) Then Exit Sub
                        If InStr(strKey, ".") <> 0 And Chr(KeyAscii) = "." Then   '只能存在一个小数点
                            KeyAscii = 0
                            Exit Sub
                        End If
                        If Len(Mid(strKey, InStr(1, strKey, ".") + 1)) >= mintNumberDigit And strKey Like "*.*" Then
                            KeyAscii = 0
                            Exit Sub
                        Else
                            Exit Sub
                        End If
                    End If
                End If
            Case mListType.汇总
                If Col = mIntCol科室汇总_留存数量 Or Col = mIntCol科室汇总_实发数量 Then
                    If InStr("1234567890" + Chr(46) + Chr(8) + Chr(13), Chr(KeyAscii)) = 0 Then
                        KeyAscii = 0
                    ElseIf KeyAscii = Asc(".") Then
                        If InStr(.EditText, ".") <> 0 Then     '只能存在一个小数点
                            KeyAscii = 0
                        End If
                    End If
                End If
        End Select
    End With
End Sub

Private Sub vsfList_MouseUp(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim objPopup As CommandBarPopup
    Dim cbrControl As CommandBarControl
    Dim LngID As Long
    Dim Int单据 As Integer
    Dim strNo As String
    Dim str审查结果 As String
    Dim lng医嘱id As Long
    Dim strsql As String
    Dim rsTmp As Recordset
    Dim lngPatiID As Long
    Dim str挂号单 As String
    Dim lng主页ID As Long
    
    '其他在表单中弹出菜单
    If vsfList(index).MouseRow < 1 Then Exit Sub
    If vsfList(index).MouseCol < 1 Then Exit Sub
    If vsfList(index).IsSubtotal(vsfList(index).MouseRow) = True Then Exit Sub
    
    '发药状态弹出菜单
    If index = mListType.发药 Then
        If Val(vsfList(index).TextMatrix(vsfList(index).MouseRow, vsfList(index).ColIndex("收发ID"))) = 0 Then Exit Sub
        If vsfList(index).MouseCol <> vsfList(index).ColIndex("审查结果") Then
            If Button = 2 Then
                Select Case Val(vsfList(index).TextMatrix(vsfList(index).MouseRow, mIntCol发药_执行状态))
                    Case mState.发药
                        LngID = conMenu_Status_Verify
                    Case mState.拒发
                        LngID = conMenu_Status_Reject
                    Case mState.不处理
                        LngID = conMenu_Status_NoProcess
                End Select
            
                If Me.cbsMain Is Nothing Then Exit Sub
                Set objPopup = Me.cbsMain.ActiveMenuBar.FindControl(xtpControlPopup, conMenu_StatusPopup)
                If Not objPopup Is Nothing Then
                    For Each cbrControl In objPopup.Controls
                        If cbrControl.Id < 320 Then
                            cbrControl.Visible = True
                        Else
                            cbrControl.Visible = False
                        End If
                    Next
                    
                    Set cbrControl = objPopup.Controls.Find(xtpControlButton, conMenu_Status_Verify, , True)
                    If Not cbrControl Is Nothing Then
                        cbrControl.Checked = (cbrControl.Id = LngID)
                    End If
                    
                    Set cbrControl = objPopup.Controls.Find(xtpControlButton, conMenu_Status_Reject, , True)
                    If Not cbrControl Is Nothing Then
                        cbrControl.Checked = (cbrControl.Id = LngID)
                    End If
        
                    Set cbrControl = objPopup.Controls.Find(xtpControlButton, conMenu_Status_NoProcess, , True)
                    If Not cbrControl Is Nothing Then
                        cbrControl.Checked = (cbrControl.Id = LngID)
                    End If
                    
                    objPopup.CommandBar.ShowPopup
                End If
            End If
        End If
    End If
    
    '退药状态弹出菜单
    If index = mListType.退药 Then
        If Val(vsfList(index).TextMatrix(vsfList(index).MouseRow, vsfList(index).ColIndex("收发ID"))) = 0 Then Exit Sub
        If Button = 2 Then
            If Me.cbsMain Is Nothing Then Exit Sub
            Set objPopup = Me.cbsMain.ActiveMenuBar.FindControl(xtpControlPopup, conMenu_StatusPopup)
            If Not objPopup Is Nothing Then
                For Each cbrControl In objPopup.Controls
                    If cbrControl.Id < 320 Then
                        cbrControl.Visible = False
                    Else
                        cbrControl.Visible = True
                    End If
                Next
                
                objPopup.CommandBar.ShowPopup
            End If
        End If
    End If
    
    If Button = 1 And vsfList(index).MouseCol = vsfList(index).ColIndex("审查结果") And index = mListType.发药 Then
        If IsInString(gstrprivs, "合理用药监测", ";") And (index = mListType.发药 Or index = mListType.退药) Then
            If Val(vsfList(index).TextMatrix(vsfList(index).MouseRow, vsfList(index).ColIndex("收发ID"))) = 0 Then Exit Sub
'            If vsfList(index).Cell(flexcpPicture, vsfList(index).MouseRow, mIntCol发药_审查结果, vsfList(index).MouseRow, mIntCol发药_审查结果) Is Nothing Then Exit Sub
            Int单据 = Val(vsfList(index).TextMatrix(vsfList(index).MouseRow, vsfList(index).ColIndex("单据")))
            strNo = vsfList(index).TextMatrix(vsfList(index).MouseRow, vsfList(index).ColIndex("NO"))
            str审查结果 = vsfList(index).TextMatrix(vsfList(index).MouseRow, vsfList(index).ColIndex("结果"))
            lng医嘱id = Val(vsfList(index).TextMatrix(vsfList(index).MouseRow, vsfList(index).ColIndex("医嘱id")))
            
            '判断是住院还是门诊病人，如果没有找到记录（无医嘱）就退出
            strsql = "Select distinct B.病人id,nvl(B.主页id,0) 主页id,nvl(C.挂号单,'') 挂号单 " & _
                " From 药品收发记录 A,住院费用记录 B,病人医嘱记录 C " & _
                " Where A.费用id=B.Id And b.医嘱序号=c.Id And nvl(B.医嘱序号,0)<>0 And C.诊疗类别 IN('5','6','7')" & _
                " And A.单据=[2] And A.no=[1] " & _
                " Union All " & _
                " Select distinct B.病人id,0 主页id,nvl(C.挂号单,'') 挂号单 " & _
                " From 药品收发记录 A,门诊费用记录 B,病人医嘱记录 C " & _
                " Where A.费用id=B.Id And b.医嘱序号=c.Id And nvl(B.医嘱序号,0)<>0 And C.诊疗类别 IN('5','6','7')" & _
                " And A.单据=[2] And A.no=[1] "
            Set rsTmp = zlDataBase.OpenSQLRecord(strsql, Me.Caption, strNo, Int单据)
        
            If rsTmp.RecordCount = 0 Then
                rsTmp.Close
                Exit Sub
            End If
        
            lngPatiID = rsTmp!病人ID
            str挂号单 = NVL(rsTmp!挂号单)
            lng主页ID = rsTmp!主页ID
            
            '获取pass菜单并弹出
            Set objPopup = Me.cbsMain.ActiveMenuBar.FindControl(xtpControlPopup, mconMenu_PASS)
            
            Call gobjPass.zlPASSPopupCommandBars_YF(mlngMode, objPopup.CommandBar, mconMenu_PASS, lngPatiID, lng主页ID, str挂号单, str审查结果, lng医嘱id)
            
            objPopup.CommandBar.ShowPopup
        End If
    End If
End Sub

Private Sub SetColumnValue(ByVal intType As Integer, ByVal str列名 As String, ByVal intValue As Integer)
    Select Case intType
        Case mListType.发药
            Select Case str列名
                Case "病人科室"
                    mIntCol发药_科室 = intValue
                Case "开单医生"
                    mIntCol发药_开单医生 = intValue
                Case "状态"
                    mIntCol发药_状态 = intValue
                Case "类型"
                    mIntCol发药_类型 = intValue
                Case "发药类型"
                    mIntCol发药_发药类型 = intValue
                Case "年龄"
                    mIntCol发药_年龄 = intValue
                Case "NO"
                    mIntCol发药_NO = intValue
                Case "记帐员"
                    mIntCol发药_记帐员 = intValue
                Case "床号"
                    mIntCol发药_床号 = intValue
                Case "姓名"
                    mIntCol发药_姓名 = intValue
                Case "性别"
                    mIntCol发药_性别 = intValue
                Case "住院号"
                    mIntCol发药_住院号 = intValue
                Case "药品名称"
                    mIntCol发药_品名 = intValue
                Case "皮试结果"
                    mIntCol发药_皮试结果 = intValue
                Case "其它名"
                    mIntCol发药_其它名 = intValue
                Case "英文名"
                    mIntCol发药_英文名 = intValue
                Case "配方名称"
                    mIntCol发药_配方名称 = intValue
                Case "规格"
                    mIntCol发药_规格 = intValue
                Case "生产商"
                    mIntCol发药_生产商 = intValue
                Case "原产地"
                    mIntCol发药_原产地 = intValue
                Case "批号"
                    mIntCol发药_批号 = intValue
                Case "效期"
                    mIntCol发药_效期 = intValue
                Case "付"
                    mIntCol发药_付 = intValue
                Case "数量"
                    mIntCol发药_数量 = intValue
                Case "单价"
                    mIntCol发药_单价 = intValue
                Case "金额"
                    mIntCol发药_金额 = intValue
                Case "单量"
                    mIntCol发药_单量 = intValue
                Case "频次"
                    mIntCol发药_频次 = intValue
                Case "用法"
                    mIntCol发药_用法 = intValue
                Case "用药次数"
                    mIntCol发药_用药次数 = intValue
                Case "记帐时间"
                    mIntCol发药_记帐时间 = intValue
                Case "说明"
                    mIntCol发药_说明 = intValue
                Case "单据"
                    mIntCol发药_单据 = intValue
                 Case "退药人"
                    mIntCol发药_退药人 = intValue
                Case "库房货位"
                    mIntCol发药_库房货位 = intValue
                Case "单量单位"
                    mIntCol发药_单量单位 = intValue
                Case "领药部门"
                    mIntCol发药_领药部门 = intValue
                Case "病人类型"
                    mIntCol发药_病人类型 = intValue
                Case "禁忌药品说明"
                    mIntCol发药_禁忌药品说明 = intValue
                Case "脚注"
                    mIntCol发药_脚注 = intValue
            End Select
        Case mListType.汇总
            If mcondition.bln按科室汇总 Then
                Select Case str列名
                    Case "病人科室"
                        mIntCol科室汇总_科室 = intValue
                    Case "药品名称"
                        mIntCol科室汇总_品名 = intValue
                    Case "规格"
                        mIntCol科室汇总_规格 = intValue
                    Case "生产商"
                        mIntCol科室汇总_生产商 = intValue
                    Case "原产地"
                        mIntCol科室汇总_原产地 = intValue
                    Case "批号"
                        mIntCol科室汇总_批号 = intValue
                    Case "效期"
                        mIntCol科室汇总_效期 = intValue
                    Case "应发数量"
                        mIntCol科室汇总_应发数量 = intValue
                    Case "留存数量"
                        mIntCol科室汇总_留存数量 = intValue
                    Case "销帐数量"
                        mIntCol科室汇总_销帐数量 = intValue
                    Case "实发数量"
                        mIntCol科室汇总_实发数量 = intValue
                    Case "单位"
                        mIntCol科室汇总_单位 = intValue
                
                    Case "单价"
                        mIntCol科室汇总_单价 = intValue
                    Case "应发金额"
                        mIntCol科室汇总_应发金额 = intValue
                    Case "批次"
                        mIntCol科室汇总_批次 = intValue
                    Case "科室ID"
                        mIntCol科室汇总_科室ID = intValue
                    Case "药品ID"
                        mIntCol科室汇总_药品ID = intValue
                        
                    Case "领药部门"
                        mIntCol科室汇总_领药部门 = intValue
                    Case "领药部门id"
                        mIntCol科室汇总_领药部门id = intValue
                End Select
            Else
                Select Case str列名
                    Case "药品名称"
                        mIntCol汇总_品名 = intValue
                    Case "规格"
                        mIntCol汇总_规格 = intValue
                    Case "生产商"
                        mIntCol汇总_生产商 = intValue
                    Case "原产地"
                        mIntCol汇总_原产地 = intValue
                    Case "批号"
                        mIntCol汇总_批号 = intValue
                    Case "效期"
                        mIntCol汇总_效期 = intValue
                    Case "数量"
                        mIntCol汇总_数量 = intValue
                    Case "单位"
                        mIntCol汇总_单位 = intValue
                    Case "单价"
                        mIntCol汇总_单价 = intValue
                    Case "金额"
                        mIntCol汇总_金额 = intValue
                End Select
            End If
        Case mListType.缺药
            Select Case str列名
                Case "病人科室"
                    mIntCol缺药_科室 = intValue
                Case "NO"
                    mIntCol缺药_NO = intValue
                Case "类型"
                    mIntCol缺药_类型 = intValue
                Case "床号"
                    mIntCol缺药_床号 = intValue
                Case "姓名"
                    mIntCol缺药_姓名 = intValue
                Case "性别"
                    mIntCol缺药_性别 = intValue
                Case "发药类型"
                    mIntCol缺药_发药类型 = intValue
                Case "药品名称"
                    mIntCol缺药_品名 = intValue
                Case "规格"
                    mIntCol缺药_规格 = intValue
                Case "生产商"
                    mIntCol缺药_生产商 = intValue
                Case "原产地"
                    mIntCol缺药_原产地 = intValue
                Case "批号"
                    mIntCol缺药_批号 = intValue
                Case "效期"
                    mIntCol缺药_效期 = intValue
                Case "数量"
                    mIntCol缺药_数量 = intValue
                Case "单价"
                    mIntCol缺药_单价 = intValue
                Case "金额"
                    mIntCol缺药_金额 = intValue
                Case "脚注"
                    mIntCol缺药_脚注 = intValue
            End Select
        Case mListType.拒发
            Select Case str列名
                Case "病人科室"
                    mIntCol拒发_科室 = intValue
                Case "状态"
                    mIntCol拒发_状态 = intValue
                Case "NO"
                    mIntCol拒发_NO = intValue
                Case "类型"
                    mIntCol拒发_类型 = intValue
                Case "发药类型"
                    mIntCol拒发_发药类型 = intValue
                Case "床号"
                    mIntCol拒发_床号 = intValue
                Case "姓名"
                    mIntCol拒发_姓名 = intValue
                Case "性别"
                    mIntCol拒发_性别 = intValue
                Case "药品名称"
                    mIntCol拒发_品名 = intValue
                Case "规格"
                    mIntCol拒发_规格 = intValue
                Case "生产商"
                    mIntCol拒发_生产商 = intValue
                Case "原产地"
                    mIntCol拒发_原产地 = intValue
                Case "批号"
                    mIntCol拒发_批号 = intValue
                Case "效期"
                    mIntCol拒发_效期 = intValue
                Case "数量"
                    mIntCol拒发_数量 = intValue
                Case "单价"
                    mIntCol拒发_单价 = intValue
                Case "金额"
                    mIntCol拒发_金额 = intValue
                Case "脚注"
                    mIntCol拒发_脚注 = intValue
            End Select
        Case mListType.退药
            Select Case str列名
                Case "发送时间"
                    mIntCol退药_发送时间 = intValue
                Case "病人科室"
                    mIntCol退药_科室 = intValue
                Case "状态"
                    mIntCol退药_状态 = intValue
                Case "类型"
                    mIntCol退药_类型 = intValue
                Case "NO"
                    mIntCol退药_NO = intValue
                Case "床号"
                    mIntCol退药_床号 = intValue
                Case "姓名"
                    mIntCol退药_姓名 = intValue
                Case "性别"
                    mIntCol退药_性别 = intValue
                Case "住院号"
                    mIntCol退药_住院号 = intValue
                Case "药品名称"
                    mIntCol退药_品名 = intValue
                Case "发药类型"
                    mIntCol退药_发药类型 = intValue
                Case "其它名"
                    mIntCol退药_其它名 = intValue
                Case "英文名"
                    mIntCol退药_英文名 = intValue
                Case "规格"
                    mIntCol退药_规格 = intValue
                Case "生产商"
                    mIntCol退药_生产商 = intValue
                Case "原产地"
                    mIntCol退药_原产地 = intValue
                Case "批号"
                    mIntCol退药_批号 = intValue
                Case "效期"
                    mIntCol退药_效期 = intValue
                Case "付"
                    mIntCol退药_付 = intValue
                Case "数量"
                    mIntCol退药_数量 = intValue
                Case "已退数"
                    mIntCol退药_已退数 = intValue
                Case "准退数"
                    mIntCol退药_准退数 = intValue
                Case "退药数"
                    mIntCol退药_退药数 = intValue
        
                Case "单价"
                    mIntCol退药_单价 = intValue
                Case "金额"
                    mIntCol退药_金额 = intValue
                Case "单量"
                    mIntCol退药_单量 = intValue
                Case "频次"
                    mIntCol退药_频次 = intValue
                Case "用法"
                    mIntCol退药_用法 = intValue
                Case "操作员"
                    mIntCol退药_操作员 = intValue
                Case "发药时间"
                    mIntCol退药_发药时间 = intValue
                Case "单据"
                    mIntCol退药_单据 = intValue
                Case "医嘱id"
                    mIntCol退药_医嘱id = intValue
                Case "领/退药人"
                    mIntCol退药_领药人 = intValue
                    
                Case "库房货位"
                    mIntCol退药_库房货位 = intValue
                Case "相关ID"
                    mIntCol退药_相关ID = intValue
                Case "药品ID"
                    mIntCol退药_药品ID = intValue
                Case "单量单位"
                    mIntCol退药_单量单位 = intValue
                Case "发药号"
                    mIntCol退药_发药号 = intValue
                Case "脚注"
                    mIntCol退药_脚注 = intValue
            End Select
    End Select
End Sub

Public Sub VerifySign()
    Dim rsData As Recordset
    
    With vsfList(mListType.退药)
        If Val(.TextMatrix(.Row, mIntCol退药_收发ID)) = 0 Then Exit Sub
        
        '如果已启用了电子签名，则需要对发药/退药人进行电子签名处理
        If Val(.TextMatrix(.Row, mIntCol退药_执行状态)) = mState.退药_退药记录 Then
            '退药记录验证
            If VerifySignatureRecored_bak(EsignTache.returnStep, .TextMatrix(.Row, mIntCol退药_单据), .TextMatrix(.Row, mIntCol退药_NO), mcondition.lng药房id, .TextMatrix(.Row, mIntCol退药_收发ID)) = False Then
                Exit Sub
            End If
        Else
            '发药记录验证
            If VerifySignatureRecoredGather(EsignTache.send, .TextMatrix(.Row, mIntCol退药_收发ID)) = False Then
                Exit Sub
            End If
        End If
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub





