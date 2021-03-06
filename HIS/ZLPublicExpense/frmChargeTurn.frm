VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "{CC0839AF-B32F-436B-8884-BE2BB3B4C73F}#3.0#0"; "zlIDKind.ocx"
Begin VB.Form frmChargeTurn 
   AutoRedraw      =   -1  'True
   Caption         =   "门(急)诊费用转住院"
   ClientHeight    =   8460
   ClientLeft      =   60
   ClientTop       =   312
   ClientWidth     =   11712
   Icon            =   "frmChargeTurn.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8460
   ScaleWidth      =   11712
   StartUpPosition =   1  '所有者中心
   Begin VB.PictureBox picBill 
      Height          =   2100
      Left            =   90
      ScaleHeight     =   2052
      ScaleWidth      =   10500
      TabIndex        =   21
      Top             =   1365
      Width           =   10545
      Begin VSFlex8Ctl.VSFlexGrid mshList 
         Height          =   1470
         Left            =   75
         TabIndex        =   22
         Top             =   90
         Width           =   5490
         _cx             =   9684
         _cy             =   2593
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
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
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483633
         FloodColor      =   192
         SheetBorder     =   -2147483633
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
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
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
   End
   Begin VB.PictureBox picBalance 
      Height          =   1950
      Left            =   6285
      ScaleHeight     =   1908
      ScaleWidth      =   3000
      TabIndex        =   19
      Top             =   4035
      Width           =   3045
      Begin VSFlex8Ctl.VSFlexGrid vsBalance 
         Height          =   1335
         Left            =   0
         TabIndex        =   20
         Top             =   135
         Width           =   2565
         _cx             =   4524
         _cy             =   2355
         Appearance      =   3
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         MousePointer    =   0
         BackColor       =   -2147483633
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483633
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483633
         FloodColor      =   192
         SheetBorder     =   -2147483633
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
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
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
      Begin VB.Label lblSum 
         AutoSize        =   -1  'True
         Caption         =   "转出合计:"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   105
         TabIndex        =   23
         Top             =   1605
         Width           =   1155
      End
   End
   Begin VB.PictureBox picList 
      Height          =   1935
      Left            =   105
      ScaleHeight     =   1884
      ScaleWidth      =   5424
      TabIndex        =   17
      Top             =   3945
      Width           =   5475
      Begin VSFlex8Ctl.VSFlexGrid mshDetail 
         Height          =   1185
         Left            =   30
         TabIndex        =   18
         Top             =   165
         Width           =   5130
         _cx             =   9049
         _cy             =   2090
         Appearance      =   1
         BorderStyle     =   1
         Enabled         =   -1  'True
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
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
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483633
         FloodColor      =   192
         SheetBorder     =   -2147483633
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
         Cols            =   10
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   250
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
   End
   Begin VB.PictureBox picTop 
      BorderStyle     =   0  'None
      Height          =   885
      Left            =   0
      ScaleHeight     =   888
      ScaleWidth      =   11712
      TabIndex        =   12
      Top             =   0
      Width           =   11715
      Begin VB.Frame fraFixed 
         BorderStyle     =   0  'None
         Height          =   405
         Left            =   60
         TabIndex        =   24
         Top             =   480
         Width           =   9435
         Begin VB.CheckBox chkShow 
            Caption         =   "仅显示可转入数据"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   0
            TabIndex        =   3
            Top             =   75
            Value           =   1  'Checked
            Width           =   2280
         End
         Begin VB.CommandButton cmdRefresh 
            Caption         =   "刷新(&R)"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   390
            Left            =   7155
            TabIndex        =   6
            Top             =   0
            Width           =   1300
         End
         Begin VB.ComboBox cbo开单科室 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   360
            Left            =   4040
            Style           =   2  'Dropdown List
            TabIndex        =   5
            Top             =   15
            Width           =   2040
         End
         Begin VB.Label lbl开单科室 
            AutoSize        =   -1  'True
            Caption         =   "开单科室"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Left            =   2780
            TabIndex        =   4
            Top             =   75
            Width           =   960
         End
      End
      Begin VB.Frame fraPati 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   60
         TabIndex        =   14
         Top             =   80
         Width           =   2820
         Begin VB.TextBox txtPatient 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   360
            Left            =   1140
            MaxLength       =   64
            TabIndex        =   25
            ToolTipText     =   "热键：F11"
            Top             =   0
            Width           =   1650
         End
         Begin zlIDKind.IDKindNew IDKind 
            Height          =   345
            Left            =   510
            TabIndex        =   26
            Top             =   0
            Width           =   630
            _ExtentX        =   1101
            _ExtentY        =   614
            Appearance      =   2
            IDKindStr       =   $"frmChargeTurn.frx":058A
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            FontSize        =   12
            FontName        =   "宋体"
            IDKind          =   -1
            ShowPropertySet =   -1  'True
            NotContainFastKey=   "F1;CTRL+F1;F2;F3;CTRL+F4;F5;F6;F7;CTRL+F7;F8;F9;F10;F11;F12;CTRL+F12;CTRL+S;CTRL+A;CTRL+R;CTRL+D;CTRL+Q;ESC;ALT+?"
            AllowAutoICCard =   -1  'True
            AllowAutoIDCard =   -1  'True
            MustSelectItems =   "姓名,就诊卡"
            BackColor       =   -2147483633
         End
         Begin VB.Label lblPatient 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "病人"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000007&
            Height          =   240
            Left            =   0
            TabIndex        =   27
            Top             =   45
            Width           =   480
         End
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   345
         Left            =   4155
         TabIndex        =   1
         Top             =   90
         Width           =   2670
         _ExtentX        =   4720
         _ExtentY        =   614
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   271122435
         CurrentDate     =   36588
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   345
         Left            =   7110
         TabIndex        =   2
         Top             =   90
         Width           =   2640
         _ExtentX        =   4657
         _ExtentY        =   614
         _Version        =   393216
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
         Format          =   271122435
         CurrentDate     =   36588
      End
      Begin zlIDKind.IDKindNew IDKindTime 
         Height          =   240
         Left            =   2880
         TabIndex        =   28
         Top             =   120
         Width           =   855
         _ExtentX        =   2350
         _ExtentY        =   423
         ShowSortName    =   0   'False
         IDKindStr       =   "发生时间|发生时间|0|0|0|0|0|0|0|0|0;登记时间|登记时间|0|0|0|0|0|0|0|0|0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   12
         FontName        =   "宋体"
         IDKind          =   -1
         DefaultCardType =   "0"
         AutoSize        =   -1  'True
         BackColor       =   -2147483633
      End
      Begin VB.Label lbl至 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "-"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   240
         Left            =   6870
         TabIndex        =   13
         Top             =   135
         Width           =   120
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H00808080&
      Height          =   435
      Left            =   0
      ScaleHeight     =   432
      ScaleWidth      =   11712
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   7665
      Width           =   11715
      Begin VB.CommandButton cmdParaSet 
         Caption         =   "参数设置(&R)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   4590
         TabIndex        =   16
         Top             =   0
         Width           =   1500
      End
      Begin VB.CommandButton cmdOk 
         Caption         =   "确定(&O)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   8220
         TabIndex        =   15
         Top             =   -15
         Width           =   1300
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   150
         TabIndex        =   9
         Top             =   0
         Width           =   1300
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "全清(&C)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   1
         Left            =   3210
         TabIndex        =   7
         Top             =   0
         Width           =   1300
      End
      Begin VB.CommandButton cmdAll 
         Caption         =   "全选(&A)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Index           =   0
         Left            =   1845
         TabIndex        =   0
         Top             =   0
         Width           =   1300
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   390
         Left            =   9570
         TabIndex        =   8
         Top             =   0
         Width           =   1300
      End
   End
   Begin MSComctlLib.StatusBar sta 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   11
      Top             =   8100
      Width           =   11715
      _ExtentX        =   20659
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2350
            MinWidth        =   882
            Picture         =   "frmChargeTurn.frx":0620
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15621
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
            AutoSize        =   2
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
   Begin XtremeDockingPane.DockingPane dkpMan 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   360
      _ExtentY        =   339
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmChargeTurn"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit '要求变量声明
Private mcllNOs As Collection '要进行费用转入的单据信息,成员：
                              ' |-cllNO(Collection),成员：单据号,票据号,结帐ID,险类,单据类型,补结算单号,开单科室ID,开单人
Private mlng病人ID As Long
Private mfrmMain As Object
Private mblnOk As Boolean
Private mbln独立执行 As Boolean '是否独立执行

Private mintPatientRange As Integer
Private mobjPati As clsPatientInfo
Private mstrPrivs As String, mlngModule As Long
Private mbln门诊转住院先审核 As Boolean
Private mbln立即销帐 As Boolean
Private mblnMultiBalance As Boolean
Private mblnPrepayStrict As Boolean, mbytPrepayLen As Byte

Private Enum mObjPancel
    Pan_Search = 1
    Pan_Bill = 2
    Pan_List = 3
    Pan_Balance = 4
    Pan_Bottom = 5
End Enum
Private mstr个人帐户 As String

Private mintIDKind As Integer
Private mblnNotClick As Boolean
Private mstrTitle As String
Private mrsFeeList As ADODB.Recordset
Private mobjThirdSwap As clsThreeSwap
Private mblnRefreshData As Boolean

Private mobjExpenceSvr As zlPublicExpense.clsExpenceSvr
Private mobjOneCardComLib As zlOneCardComLib.clsOneCardComLib
Private mblnNewClinicPati As Boolean '是否为新门诊病人

Private Sub InitPancel()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:区域设置
    '编制:刘兴洪
    '日期:2011-03-25 17:37:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim panThis As Pane
    Dim panTop As Pane, panRight As Pane
    
    Set panTop = dkpMan.CreatePane(mObjPancel.Pan_Search, 200, 580, DockTopOf, Nothing)
    panTop.Title = "条件窗体"
    panTop.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable Or PaneNoCaption
    panTop.Tag = mObjPancel.Pan_Search
    panTop.Handle = picTop.hWnd
    If mbln门诊转住院先审核 Then
        panTop.MaxTrackSize.Height = 495 / Screen.TwipsPerPixelY
        panTop.MinTrackSize.Height = 495 / Screen.TwipsPerPixelY
    Else
        panTop.MaxTrackSize.Height = 850 / Screen.TwipsPerPixelY
        panTop.MinTrackSize.Height = 850 / Screen.TwipsPerPixelY
    End If
    
    Set panThis = dkpMan.CreatePane(mObjPancel.Pan_Bill, 250, 580, DockBottomOf, panTop)
    panThis.Title = "门诊转住院列表"
    panThis.Tag = mObjPancel.Pan_Bill
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panThis.Handle = picBill.hWnd
    

    Set panRight = dkpMan.CreatePane(mObjPancel.Pan_Balance, 1500 / Screen.TwipsPerPixelX, 580, DockRightOf, panThis)
    panRight.Title = "门诊转住院结算信息"
    panRight.Tag = mObjPancel.Pan_Balance
    panRight.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panRight.Handle = picBalance.hWnd
    
    Set panThis = dkpMan.CreatePane(mObjPancel.Pan_List, 250, 580, DockBottomOf, panThis)
    panThis.Title = "单据明细列表"
    panThis.Tag = mObjPancel.Pan_List
    panThis.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    panThis.Handle = picList.hWnd
 
    dkpMan.Options.ThemedFloatingFrames = True
    dkpMan.Options.HideClient = True
End Sub

Private Sub cbo开单科室_Click()
    If mblnNotClick Then Exit Sub
    Call ShowBills(mlng病人ID, dtpBegin.value, dtpEnd.value)
End Sub

Private Sub chkShow_Click()
    If mblnNotClick Then Exit Sub
    Call ShowBills(mlng病人ID, dtpBegin.value, dtpEnd.value)
End Sub

Private Sub dkpMan_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Then Cancel = True
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case Pan_Search
        Item.Handle = picTop.hWnd
    Case Pan_Bill
        Item.Handle = picBill.hWnd
    Case Pan_List
        Item.Handle = picList.hWnd
    Case Pan_Balance
        Item.Handle = picBalance.hWnd
    End Select
End Sub

Public Function ShowMe(frmMain As Object, ByVal lng病人ID As Long, Optional ByVal bln独立执行 As Boolean, _
    Optional ByVal strPrivs As String, Optional ByVal lngModule As Long, Optional ByRef blnRefreshData As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:门诊费用转住院费用
    '入参:
    '   bln独立执行:是否独立执行，如果是独立执行则会提交数据到数据库，否则由 ExecuteTurn 接口单独执行
    '出参:
    '   blnRefreshData-门诊费用转住院后是否刷新数据
    '返回:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mbln独立执行 = bln独立执行
    mlng病人ID = lng病人ID: mstrPrivs = strPrivs: mlngModule = lngModule
    mblnRefreshData = False: txtPatient.Tag = lng病人ID
    Set mfrmMain = frmMain
    
    mblnOk = False
    On Error Resume Next
    Me.Show vbModal, frmMain
    ShowMe = mblnOk
    blnRefreshData = mblnRefreshData
End Function

Private Sub ClearData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查单据中输入的负数数量及退回科室是否正确
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2010-11-09 17:30:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mshList.Redraw = flexRDNone
    mshList.Clear 1: mshList.Rows = 2
    sta.Panels(2).Text = ""
    Call setHeader: Call SetBillColor
    mshList.Redraw = flexRDBuffered
    Set mrsFeeList = Nothing
    cbo开单科室.Clear
End Sub

Private Function IsYBSingle(ByVal strNo As String, Optional blnYBAllDel_Out As Boolean, Optional ByRef blnThirdAllDel_Out As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检测医保是否全退还是分单据据
    '入参:strNo-指定单据
    '出参:blnThirdAllDel-三方卡是否必须全退
    '     blnYBAllDel_Out-医保是否必须全退
    '返回:分单据退，返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-09-13 14:16:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    On Error GoTo errHandle
    blnThirdAllDel_Out = False: blnYBAllDel_Out = False
    
    strSql = "Select 1 From 医保结算明细 Where NO = [1] And Rownum < 2 And 卡类别ID is NULL "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "检测医保是否全退还是分单据据", strNo)
    
    blnYBAllDel_Out = rsTmp.EOF
    If rsTmp.EOF Then IsYBSingle = False: Exit Function
    
    blnThirdAllDel_Out = CheckAllTurn(strNo)
    IsYBSingle = Not blnThirdAllDel_Out
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ExcuteTranSaveOver(ByVal objPati As clsPatientInfo, ByRef objBalanceInfor As clsBalanceInfo, _
    ByRef cllBillPro As Collection, Optional blnNotModify As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行转帐完成保存
    '入参:objBalanceInfor-结帐信息
    '     blnNotModify-是否不进行数据修正
    '出参:
    '返回:转帐成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-08-17 16:18:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, cllPro As Collection
    Dim blnTrans As Boolean, i As Long
    
    On Error GoTo errHandle
    
    If cllBillPro Is Nothing Then Set cllBillPro = New Collection
    
    Set cllPro = New Collection
    
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    If Not blnNotModify Then
        '    '完成执行
        '    Zl_门诊费用转住院_Modify_s
        strSql = "Zl_门诊费用转住院_Modify_s("
        '    操作类型_In   Number,  '0-仅更新校对标志:只更新关联交易ID的校对标志;1-普通退费方式:2.三方卡退费结算:;3-医保结算;4-消费卡结算:
        strSql = strSql & "1,"
        '    冲销id_In     病人预交记录.结帐id%Type,
        strSql = strSql & "" & objBalanceInfor.冲销ID & ","
        '    病人id_In     病人结帐记录.病人id%Type,
        strSql = strSql & "" & objPati.病人ID & ","
        '  姓名_In         病人预交记录.姓名%Type,
        strSql = strSql & "'" & objPati.姓名 & "' ,"
        '  性别_In         病人预交记录.性别%Type,
        strSql = strSql & "'" & objPati.性别 & "' ,"
        '  年龄_In         病人预交记录.年龄%Type,
        strSql = strSql & "'" & objPati.年龄 & "' ,"
        '  门诊号_In       病人预交记录.门诊号%Type,
        strSql = strSql & "'" & objPati.门诊号 & "' ,"
        '  住院号_In       病人预交记录.住院号%Type,
        strSql = strSql & "'" & objPati.住院号 & "' ,"
        '  付款方式名称_In 病人预交记录.付款方式名称%Type,
        strSql = strSql & "'" & objPati.医疗付款方式 & "' ,"
        '    结算方式_In   Varchar2,
        strSql = strSql & "NULL,"
        '    操作员编号_In 病人预交记录.操作员编号%Type := Null,
        strSql = strSql & "'" & UserInfo.编号 & "' ,"
        '    操作员姓名_In 病人预交记录.操作员姓名%Type := Null,
        strSql = strSql & "'" & UserInfo.姓名 & "' ,"
        '    完成退费_In   Number := 0,0-未完成退费;1-完成退费
        strSql = strSql & "1)"
        '    关联交易id_In 病人预交记录.Id%Type := Null,
        '    退款时间_In   病人预交记录.收款时间%Type := Null,
        '    校对标志_In   病人预交记录.校对标志%Type := Null,
        '    误差金额_In   病人预交记录.冲预交%Type := Null,
        '    卡类别id_In   病人预交记录.卡类别id%Type := Null,
        '    卡号_In       病人预交记录.卡号%Type := Null,
        '    交易流水号_In 病人预交记录.交易流水号%Type := Null,
        '    交易说明_In   病人预交记录.交易说明%Type := Null,
        '    清除原交易_In Number:=0
        zlAddArray cllPro, strSql
    End If
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, "完成门诊费用转住院"
    blnTrans = False
    ExcuteTranSaveOver = True
    Set cllBillPro = New Collection
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ExecuteTurn(ByVal frmMain As Object, ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByVal str住院号 As String, ByVal dat入院时间 As Date, ByVal lng入院科室ID As Long, ByVal lng入院病区ID As Long, _
    ByRef strErrmsg_Out As String, Optional ByRef blnReflashData_Out As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据指定的单据号序列,执行门诊费用转住院费用,及医保退费结算操作
    '入参:
    '   lng住院号-住院号,lng主页ID-主页ID,这两个参数仅在医保入院补充登记时才传入
    '出参:
    '   strErrMsg_Out=失败时返回错误原因
    '   blnReflashData_Out=是否有数据转入
    '返回:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, lngcnt As Long, bln医保单张退 As Boolean
    Dim strSql As String, strInvoice As String
    Dim cllPro As Collection, str已转结帐ID As String
    Dim intInsure As Integer, blnTurnAll As Boolean
    Dim objBalanceInfor As clsBalanceInfo
    Dim strSfNos As String, blnBillPrintInited As Boolean
    Dim lngStep As Long, bln存在结帐单 As Boolean
    Dim strNewNo As String, strNewNos As String, varNos As Variant, p As Integer
    Dim strDelDate As String, cllNO As Collection
    '补充结算的单据处理思路：先将费用单据转为住院费用记录，再单独处理门诊退费
    Dim strReplenishNo As String, strReplenishNos As String 'Array(补结算单据号,转费用SQL,新单据号)
    Dim cllReplenishPro As Collection, lng医疗小组ID As Long
    
    On Error GoTo errHandle
    blnReflashData_Out = False
    If mbln独立执行 = False And mblnNewClinicPati Then
        ExecuteTurn = frmChargeTurnNew.ExecuteTurn(frmMain, lng病人ID, lng主页ID, _
            str住院号, dat入院时间, lng入院科室ID, lng入院病区ID, strErrmsg_Out, blnReflashData_Out)
        Exit Function
    End If
    
    If mlng病人ID <> lng病人ID Then
        strErrmsg_Out = "本次选择转入费用所属病人与当前病人不同，不允许执行门诊费用转住院。": Exit Function
    End If
    
    If mcllNOs Is Nothing Then Exit Function
    If mcllNOs.Count = 0 Then ExecuteTurn = True: Exit Function
     
    strDelDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    If mbln立即销帐 Then Call zlBillPrint_Initialize(Val("1137-病人结帐管理"))
    
    Set mobjThirdSwap = New clsThreeSwap
    If mobjThirdSwap.Init(mobjOneCardComLib, frmMain, mlngModule, _
        mobjPati.病人ID, mobjPati.姓名, mobjPati.性别, mobjPati.年龄) = False Then Exit Function
    
    Set objBalanceInfor = New clsBalanceInfo
    With objBalanceInfor
        .结帐时间 = CDate(strDelDate)
        .结算类型 = 3  '结算类型:1-门诊结帐;2-住院结帐;3-门诊费用转住院
    End With
    
    Set cllPro = New Collection
    Set cllReplenishPro = New Collection
    
    zlCommFun.ShowFlash "正在进行门诊费用转住院处理，请稍后...", frmMain
    
    '单据号,票据号,结帐ID,险类,单据类型,补结算单号,开单科室ID,开单人
    lngStep = 0
    i = 1
    Do While i <= mcllNOs.Count
        lngStep = lngStep + 1
        Set cllNO = mcllNOs(i)
        
        '同一个发票号的一起转
        lngcnt = 1
        strInvoice = cllNO("票据号")
        If strInvoice <> "" Then
            For j = i + 1 To mcllNOs.Count
                Set cllNO = mcllNOs(j)
                If strInvoice = cllNO("票据号") Then
                    lngcnt = lngcnt + 1
                Else
                    Exit For
                End If
            Next
        End If
        
        
        '医保要求从最后一张开始退,读出的数据是按单据号倒序排列的，所以此处正序即可
        For j = i To i + lngcnt - 1
            Set cllNO = mcllNOs(j)
            bln医保单张退 = False: blnTurnAll = False
            
            strReplenishNo = cllNO("补结算单号")
            If strReplenishNo = "" Then
                If Val(cllNO("险类")) <> 0 Then '记帐单，险类为0
                    bln医保单张退 = IsYBSingle(cllNO("单据号"))
                Else
                    blnTurnAll = CheckAllTurn(cllNO("单据号"))
                    If InStr("," & str已转结帐ID & ",", "," & cllNO("结帐ID") & ",") > 0 Then blnTurnAll = True
                End If
            End If
            
            With objBalanceInfor
                .结帐ID = Val(cllNO("结帐ID"))
                .结帐单据号 = cllNO("单据号")
            End With
            intInsure = Val(cllNO("险类"))
            
            '先处理的记帐单，当前单据不是记帐单，说明记帐单已处理完
            If cllNO("单据类型") <> "记帐单" And mbln立即销帐 And Not blnBillPrintInited Then
                Call zlBillPrint_Initialize(Val("1121-门诊收费管理"))
                blnBillPrintInited = True
            End If
    
            lng医疗小组ID = ZlGetMedicalGroupID(lng病人ID, lng主页ID, cllNO("开单科室ID"), cllNO("开单人"), dat入院时间)
            
            If bln医保单张退 Or (intInsure = 0 And Not blnTurnAll) Or strReplenishNo <> "" Then
                
                If InStr("," & str已转结帐ID & ",", "," & cllNO("结帐ID") & ",") = 0 Then ' 可能一次结帐分单据的，已经转出，所以要判断
                    strNewNo = zlDatabase.NextNo(14)
                    
                    'Zl_门诊费用转住院_Insert_S
                    strSql = "Zl_门诊费用转住院_insert_S("
                    '  No_In         住院费用记录.NO%Type,
                    strSql = strSql & "'" & cllNO("单据号") & "',"
                    '  Newno_In        住院费用记录.No%Type,
                    strSql = strSql & "'" & strNewNo & "',"
                    '  住院号_In     住院费用记录.标识号%Type, --医保入院补充登记时才传入
                    strSql = strSql & "" & ZVal(str住院号) & ","
                    '  主页id_In     住院费用记录.主页id%Type, --医保入院补充登记时才传入
                    strSql = strSql & "" & ZVal(lng主页ID) & ","
                    '  入院时间_In   住院费用记录.发生时间%Type,
                    strSql = strSql & "To_Date('" & Format(dat入院时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
                    '  入院科室id_In 病人预交记录.科室id%Type,
                    strSql = strSql & "" & ZVal(lng入院科室ID) & ","
                    '  入院病区id_In 住院费用记录.病人病区id%Type,
                    strSql = strSql & "" & ZVal(lng入院病区ID) & ","
                    '  转出时间_In   住院费用记录.登记时间%Type, --多张单据转出时,每张单据的转出时间相同,都是系统当前时间
                    strSql = strSql & "To_Date('" & strDelDate & "','YYYY-MM-DD HH24:MI:SS'),"
                    '  操作员姓名_In 住院费用记录.操作员姓名%Type,
                    strSql = strSql & "'" & UserInfo.姓名 & "',"
                    '  医疗小组id_In   住院费用记录.医疗小组id%Type,
                    strSql = strSql & "" & ZVal(lng医疗小组ID) & ","
                    '  床号_In         住院费用记录.床号%Type,
                    strSql = strSql & "'" & mobjPati.床号 & "',"
                    '  单据性质_In Number := 1, --1-门诊收费单;2-门诊记帐单
                    strSql = strSql & "" & IIf(cllNO("单据类型") = "记帐单", 2, 1) & ")"
                    
                    If strReplenishNo <> "" And mbln立即销帐 Then
                        If InStr(strReplenishNos & ";", ";" & strReplenishNo & "," & cllNO("险类") & ";") = 0 Then
                            strReplenishNos = strReplenishNos & ";" & strReplenishNo & "," & cllNO("险类")
                        End If
                        'Array(补结算单据号,转费用SQL,新单据号)
                        cllReplenishPro.Add Array(strReplenishNo, strSql, strNewNo)
                    Else
                        zlAddArray cllPro, strSql
                        If cllNO("单据类型") = "记帐单" And mbln立即销帐 Then
                            'Zl_门诊转住院_记帐转出
                            strSql = "Zl_门诊转住院_记帐转出("
                            '  No_In         住院费用记录.No%Type,
                            strSql = strSql & "'" & cllNO("单据号") & "',"
                            '  操作员编号_In 住院费用记录.操作员编号%Type,
                            strSql = strSql & "'" & UserInfo.编号 & "',"
                            '  操作员姓名_In 住院费用记录.操作员姓名%Type,
                            strSql = strSql & "'" & UserInfo.姓名 & "',"
                            '  销账时间_In   住院费用记录.发生时间%Type
                            strSql = strSql & "To_Date('" & strDelDate & "','YYYY-MM-DD HH24:MI:SS'))"
                            zlAddArray cllPro, strSql
                            
                            If DelBalaceMz(mobjPati, cllPro, lng主页ID, lng入院科室ID, objBalanceInfor) = False Then
                                blnReflashData_Out = objBalanceInfor.是否保存结帐单
                                Exit Function
                            End If
                            bln存在结帐单 = True
                        ElseIf mbln立即销帐 And cllNO("单据类型") <> "记帐单" Then
                            strSfNos = "'" & cllNO("单据号") & "'"
                            If zlBillPrint_EraseBill(strSfNos, 0) = False Then Exit Function
                            
                            With objBalanceInfor
                                .结算类型 = 3 '结算类型:1-门诊结帐;2-住院结帐;3-门诊费用转住院
                                .冲销ID = zlDatabase.GetNextId("病人结帐记录")
                            End With
                            
                            'Zl_门诊转住院_收费转出_s
                            strSql = "Zl_门诊转住院_收费转出_s("
                            '  病人id_In       病人结帐记录.病人id%Type,
                            strSql = strSql & "" & mobjPati.病人ID & ","
                            '  姓名_In         病人预交记录.姓名%Type,
                            strSql = strSql & "'" & mobjPati.姓名 & "' ,"
                            '  性别_In         病人预交记录.性别%Type,
                            strSql = strSql & "'" & mobjPati.性别 & "' ,"
                            '  年龄_In         病人预交记录.年龄%Type,
                            strSql = strSql & "'" & mobjPati.年龄 & "' ,"
                            '  门诊号_In       病人预交记录.门诊号%Type,
                            strSql = strSql & "'" & mobjPati.门诊号 & "' ,"
                            '  住院号_In       病人预交记录.住院号%Type,
                            strSql = strSql & "'" & mobjPati.住院号 & "' ,"
                            '  付款方式名称_In 病人预交记录.付款方式名称%Type,
                            strSql = strSql & "'" & mobjPati.医疗付款方式 & "' ,"
                            '  No_In         住院费用记录.No%Type,
                            strSql = strSql & "'" & cllNO("单据号") & "',"
                            '  操作员编号_In 住院费用记录.操作员编号%Type,
                            strSql = strSql & "'" & UserInfo.编号 & "',"
                            '  操作员姓名_In 住院费用记录.操作员姓名%Type,
                            strSql = strSql & "'" & UserInfo.姓名 & "',"
                            '  退费时间_In   住院费用记录.发生时间%Type,
                            strSql = strSql & "To_Date('" & strDelDate & "','YYYY-MM-DD HH24:MI:SS'),"
                            '  入院科室id_In 住院费用记录.开单部门id%Type := Null,
                            strSql = strSql & "" & ZVal(lng入院科室ID) & ","
                            '  主页id_In     住院费用记录.主页id%Type := Null,
                            strSql = strSql & "" & ZVal(lng主页ID) & ","
                            '  冲销id_In     病人预交记录.结帐id%Type := Null,
                            strSql = strSql & "" & objBalanceInfor.冲销ID & ")"
                            '  原结帐id_In   病人预交记录.结帐id%Type := Null
                            zlAddArray cllPro, strSql
                            
                             '执行医保:
                            If ExcuteInsureDel(objBalanceInfor, intInsure, objBalanceInfor.结帐单据号, cllPro) = False Then
                                blnReflashData_Out = objBalanceInfor.是否保存结帐单
                                Exit Function
                            End If
                            '执行一卡通
                            If Not ExecuteThirdReturnMoneySwap(mobjPati, objBalanceInfor, cllPro) Then
                                blnReflashData_Out = objBalanceInfor.是否保存结帐单
                                Exit Function
                            End If
                            '完成
                            If ExcuteTranSaveOver(mobjPati, objBalanceInfor, cllPro) = False Then
                                blnReflashData_Out = objBalanceInfor.是否保存结帐单
                                Exit Function
                            End If
                        Else
                            '直接门诊费用转住院
                            If Not ExcuteTranSaveOver(mobjPati, objBalanceInfor, cllPro, True) Then Exit Function
                        End If
                        
                        Call mobjExpenceSvr.zlAdjustFeeData(strNewNo)
                    End If
                End If
            Else
                If InStr("," & str已转结帐ID & ",", "," & cllNO("结帐ID") & ",") = 0 Then
                    If cllNO("单据类型") = "记帐单" Then
                        varNos = Array(cllNO("单据号"))
                    Else '收费单，一次转出结算的所有单据
                        strSfNos = GetBalanceNos(1, cllNO("结帐ID"))
                        varNos = Split(strSfNos, ",")
                    End If
                    
                    strNewNos = ""
                    For p = 0 To UBound(varNos)
                        strNewNo = zlDatabase.NextNo(14)
                        strNewNos = strNewNos & "," & strNewNo
                        
                        'Zl_门诊费用转住院_Insert_S
                        strSql = "Zl_门诊费用转住院_insert_S("
                        '  No_In         住院费用记录.NO%Type,
                        strSql = strSql & "'" & varNos(p) & "',"
                        '  Newno_In        住院费用记录.No%Type,
                        strSql = strSql & "'" & strNewNo & "',"
                        '  住院号_In     住院费用记录.标识号%Type, --医保入院补充登记时才传入
                        strSql = strSql & "" & ZVal(str住院号) & ","
                        '  主页id_In     住院费用记录.主页id%Type, --医保入院补充登记时才传入
                        strSql = strSql & "" & ZVal(lng主页ID) & ","
                        '  入院时间_In   住院费用记录.发生时间%Type,
                        strSql = strSql & "To_Date('" & Format(dat入院时间, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
                        '  入院科室id_In 病人预交记录.科室id%Type,
                        strSql = strSql & "" & ZVal(lng入院科室ID) & ","
                        '  入院病区id_In 住院费用记录.病人病区id%Type,
                        strSql = strSql & "" & ZVal(lng入院病区ID) & ","
                        '  转出时间_In   住院费用记录.登记时间%Type, --多张单据转出时,每张单据的转出时间相同,都是系统当前时间
                        strSql = strSql & "To_Date('" & strDelDate & "','YYYY-MM-DD HH24:MI:SS'),"
                        '  操作员姓名_In 住院费用记录.操作员姓名%Type,
                        strSql = strSql & "'" & UserInfo.姓名 & "',"
                        '  医疗小组id_In   住院费用记录.医疗小组id%Type,
                        strSql = strSql & "" & ZVal(lng医疗小组ID) & ","
                        '  床号_In         住院费用记录.床号%Type,
                        strSql = strSql & "'" & mobjPati.床号 & "',"
                        '  单据性质_In Number := 1, --1-门诊收费单;2-门诊记帐单
                        strSql = strSql & "" & IIf(cllNO("单据类型") = "记帐单", 2, 1) & ")"
                        zlAddArray cllPro, strSql
                    Next
                    If strNewNos <> "" Then strNewNos = Mid(strNewNos, 2)
                    
                    If cllNO("单据类型") = "记帐单" And mbln立即销帐 Then
                        'Zl_门诊转住院_记帐转出
                        strSql = "Zl_门诊转住院_记帐转出("
                        '  No_In         住院费用记录.No%Type,
                        strSql = strSql & "'" & cllNO("单据号") & "',"
                        '  操作员编号_In 住院费用记录.操作员编号%Type,
                        strSql = strSql & "'" & UserInfo.编号 & "',"
                        '  操作员姓名_In 住院费用记录.操作员姓名%Type,
                        strSql = strSql & "'" & UserInfo.姓名 & "',"
                        '  销账时间_In   住院费用记录.发生时间%Type
                        strSql = strSql & "To_Date('" & strDelDate & "','YYYY-MM-DD HH24:MI:SS'))"
                        zlAddArray cllPro, strSql
                        
                        If DelBalaceMz(mobjPati, cllPro, lng主页ID, lng入院科室ID, objBalanceInfor) = False Then
                            blnReflashData_Out = objBalanceInfor.是否保存结帐单
                            Exit Function
                        End If
                        bln存在结帐单 = True
                    ElseIf mbln立即销帐 And cllNO("单据类型") <> "记帐单" Then
                        strSfNos = "'" & Replace(strSfNos, ",", "','") & "'"
                        If zlBillPrint_EraseBill(strSfNos, 0) = False Then Exit Function
                        
                        With objBalanceInfor
                            .结算类型 = 3 '结算类型:1-门诊结帐;2-住院结帐;3-门诊费用转住院
                            .冲销ID = zlDatabase.GetNextId("病人结帐记录")
                        End With
                        
                        'Zl_门诊转住院_收费转出_s
                        strSql = "Zl_门诊转住院_收费转出_s("
                        '  病人id_In       病人结帐记录.病人id%Type,
                        strSql = strSql & "" & mobjPati.病人ID & ","
                        '  姓名_In         病人预交记录.姓名%Type,
                        strSql = strSql & "'" & mobjPati.姓名 & "' ,"
                        '  性别_In         病人预交记录.性别%Type,
                        strSql = strSql & "'" & mobjPati.性别 & "' ,"
                        '  年龄_In         病人预交记录.年龄%Type,
                        strSql = strSql & "'" & mobjPati.年龄 & "' ,"
                        '  门诊号_In       病人预交记录.门诊号%Type,
                        strSql = strSql & "'" & mobjPati.门诊号 & "' ,"
                        '  住院号_In       病人预交记录.住院号%Type,
                        strSql = strSql & "'" & mobjPati.住院号 & "' ,"
                        '  付款方式名称_In 病人预交记录.付款方式名称%Type,
                        strSql = strSql & "'" & mobjPati.医疗付款方式 & "' ,"
                        '  No_In         住院费用记录.No%Type,
                        strSql = strSql & "'" & cllNO("单据号") & "',"
                        '  操作员编号_In 住院费用记录.操作员编号%Type,
                        strSql = strSql & "'" & UserInfo.编号 & "',"
                        '  操作员姓名_In 住院费用记录.操作员姓名%Type,
                        strSql = strSql & "'" & UserInfo.姓名 & "',"
                        '  退费时间_In   住院费用记录.发生时间%Type,
                        strSql = strSql & "To_Date('" & strDelDate & "','YYYY-MM-DD HH24:MI:SS'),"
                        '  入院科室id_In 住院费用记录.开单部门id%Type := Null,
                        strSql = strSql & "" & ZVal(lng入院科室ID) & ","
                        '  主页id_In     住院费用记录.主页id%Type := Null,
                        strSql = strSql & "" & ZVal(lng主页ID) & ","
                        '  冲销id_In     病人预交记录.结帐id%Type := Null,
                        strSql = strSql & "" & objBalanceInfor.冲销ID & ","
                        '  原结帐id_In   病人预交记录.结帐id%Type := Null,
                        strSql = strSql & "" & objBalanceInfor.结帐ID & ")"
                        zlAddArray cllPro, strSql
                        
                         '执行医保:
                        If ExcuteInsureDel(objBalanceInfor, intInsure, "", cllPro) = False Then
                            blnReflashData_Out = objBalanceInfor.是否保存结帐单
                            Exit Function
                        End If
                        '执行一卡通
                        If Not ExecuteThirdReturnMoneySwap(mobjPati, objBalanceInfor, cllPro) Then
                            blnReflashData_Out = objBalanceInfor.是否保存结帐单
                            Exit Function
                        End If
                        '完成
                        If ExcuteTranSaveOver(mobjPati, objBalanceInfor, cllPro) = False Then Exit Function
                    Else
                        '直接门诊费用转住院
                        If Not ExcuteTranSaveOver(mobjPati, objBalanceInfor, cllPro, True) Then Exit Function
                    End If
                    
                    Call mobjExpenceSvr.zlAdjustFeeData(strNewNos)
                End If
                str已转结帐ID = str已转结帐ID & "," & cllNO("结帐ID")
            End If
        Next
        i = i + lngcnt
    Loop
    
    sta.Panels(2).Text = ""
    
    '对补充结算单据进行退费处理
    If strReplenishNos <> "" Then
        strReplenishNos = Mid(strReplenishNos, 2)
        If ExecuteReplenishDel(strReplenishNos, cllReplenishPro, lng主页ID, lng入院科室ID, strDelDate) = False Then
            Exit Function
        End If
    End If
    
    '打印预交款部分
    Call PrintPrePayPrint(strDelDate)
    
    '显示结帐窗口
    If bln存在结帐单 And mbln立即销帐 Then
       Call ShowBalanceWindows(frmMain, strDelDate)
    End If
    
    ExecuteTurn = True
    Exit Function
errHandle:
    zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ExecuteReplenishDel(ByVal strNOs As String, ByVal cllPro As Collection, _
    ByVal lng主页ID As Long, ByVal lng入院科室ID As Long, ByVal strDelDate As String) As Boolean
    '功能:对补充结算的单据进行转费用及退费处理
    '入参:
    '   strNos 补结算单号,格式：单据号,险类;...
    '   cllPro 传入的退费过程的集合：Array(补结算单据号,转费用SQL,新单据号)
    '   strDelDate 退费时间
    Dim strSql As String, strNoTemp As String
    Dim varNos As Variant, i As Long, p As Long, blnTrans As Boolean
    Dim strNo As String, intInsure As Integer
    Dim lng结算冲销ID  As Long, lng费用冲销ID As Long, lng结算序号 As Long
    Dim lng原结帐ID As Long, strAdvance As String
    Dim strNewNos As String, strNewNo As String
    
    Err = 0: On Error GoTo errH
    If strNOs = "" Then ExecuteReplenishDel = True: Exit Function
    
    Call zlBillPrint_Initialize(Val("1124-保险补充结算"))
    varNos = Split(strNOs, ";")
    For i = 0 To UBound(varNos)
        '单据号,险类;...
        strNo = Split(varNos(i), ",")(0): intInsure = Split(varNos(i), ",")(1)
        
        If zlBillPrint_EraseBill(strNo, 0) = False Then Exit Function
        
        lng费用冲销ID = zlDatabase.GetNextId("病人结帐记录")
        lng结算冲销ID = zlDatabase.GetNextId("病人结帐记录")
        lng结算序号 = -1 * lng费用冲销ID
        
        gcnOracle.BeginTrans: blnTrans = True
        For p = 1 To cllPro.Count
            'Array(补结算单据号,转费用SQL,新单据号)
            strNoTemp = cllPro(p)(0): strSql = cllPro(p)(1): strNewNo = cllPro(p)(2)
            If strNoTemp = strNo Then
                strNewNos = strNewNos & "," & strNewNo
                zlDatabase.ExecuteProcedure strSql, "执行补结算作废"
            End If
        Next
        If strNewNos <> "" Then strNewNos = Mid(strNewNos, 2)
        
        'Zl_门诊转住院_补结算转出_s(
        strSql = "Zl_门诊转住院_补结算转出_s("
        '  No_In         费用补充记录.No%Type,
        strSql = strSql & "'" & strNo & "',"
        '  费用冲销id_In     病人预交记录.结帐id%Type,
        strSql = strSql & "" & lng费用冲销ID & ","
        '  结算冲销id_In     病人预交记录.结帐id%Type,
        strSql = strSql & "" & lng结算冲销ID & ","
        '  结算序号_In     病人预交记录.结算序号%Type,
        strSql = strSql & "" & lng结算序号 & ","
        '  退费时间_In   住院费用记录.发生时间%Type,
        strSql = strSql & "To_Date('" & strDelDate & "','yyyy-mm-dd hh24:mi:ss'),"
        '  操作员编号_In 住院费用记录.操作员编号%Type,
        strSql = strSql & "'" & UserInfo.编号 & "',"
        '  操作员姓名_In 住院费用记录.操作员姓名%Type,
        strSql = strSql & "'" & UserInfo.姓名 & "',"
        '  主页id_In     病人预交记录.主页id%Type,
        strSql = strSql & "" & lng主页ID & ","
        '  入院科室id_In 病人预交记录.科室id%Type,
        strSql = strSql & "" & lng入院科室ID & ")"
        zlDatabase.ExecuteProcedure strSql, "执行补结算转出"
        
        'Public Function ClinicDelSwap(lngStlID As Long, Optional ByVal bln退费 As Boolean = True, _
            Optional ByVal intinsure As Integer = 0, Optional ByRef strAdvance As String = "") As Boolean
            '---------------------------------------------------------------------------------------------------------------------------------------------
            '功能:将门诊退费的明细和结算数据转发送医保前置服务器确认
            '入参:lngStlID-将要退的费记录的结帐ID；，从预交记录中可以检索医保号和密码
            '     bln退费 -表明是退费交易还是改费交易在调用本接口
            '     strAdvance:格式:冲销ID|补充结算标志|…,每位|分隔
            '           第一位:传入冲销ID,医保可以根据冲销ID来进行取数
            '           第二位:补充结算标志,1-补充结算调和;0非补充结算调用
            '           第三位:NO:当前结算的NO
            '           第四位后: 待以后扩展
            '     注意：
            '           strAdvance在10.34.0以前(不含补允结算)
            '               多单据一次结算时,传入的是原结帐IDs:结帐ID1,结帐ID2,...
            '               其他，传入格式为:退费单据总张数|当前退第几张单据
            '出参:strAdvance:1.原样退回时，返回空
            '                2.退费结算方式与收费结算方式不一致时，返回格式为：结算方式|金额||结算方式|金额||…（其中，金额为负）
            '返回：交易成功返回true；否则，返回false
        strAdvance = lng结算冲销ID & "|1"
        lng原结帐ID = GetFromNOToLastBalanceID(strNo, , , , True)
        If Not gclsInsure.ClinicDelSwap(lng原结帐ID, True, intInsure, strAdvance) Then
            gcnOracle.RollbackTrans
            MsgBox "医保结算失败，无法继续进行门诊费用转住院操作。", vbInformation, gstrSysName
            Exit Function
        End If
        gcnOracle.CommitTrans: blnTrans = False
        Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, intInsure)
        
        Call mobjExpenceSvr.zlAdjustFeeData(strNewNos)
    Next
    ExecuteReplenishDel = True
    Exit Function
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        'Resume
    End If
    Call SaveErrLog
End Function

Private Function GetFromNOToLastBalanceID(ByVal strNOs As String, _
    Optional ByVal blnNOMoved As Boolean, _
    Optional ByVal bln历史表同步查 As Boolean = False, _
    Optional lng结算序号 As Long, Optional bln补结算 As Boolean = False) As Long
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据一张收费单据的NO，返回最后一次有效的结帐的ID
    '入参:blnNoMoved是否在后备表中，查询单据之前的判断需要用这个参数
    '     bln历史表同步查-是否连接历史表一起查询
    '     bln补结算-是否补充结算
    '出参:lng结算序号-返回最后一次有效的结帐序号
    '返回:结帐ID
    '编制:刘兴洪
    '日期:2014-06-17 17:06:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSql As String, strSQL1 As String
    
    On Error GoTo errHandle:
    '87975
    strSql = _
            " Select /*+cardinality(m,10)*/ Max(a.结帐id) As 结帐id" & vbNewLine & _
            " From 门诊费用记录 A, Table(f_Str2list([1])) M" & vbNewLine & _
            " Where a.No = m.Column_Value" & vbNewLine & _
            "       And a.登记时间 + 0 =" & vbNewLine & _
            "           (Select /*+cardinality(j,10)*/ Max(m.登记时间)" & vbNewLine & _
            "            From 门诊费用记录 M, Table(f_Str2list([1])) J" & vbNewLine & _
            "            Where m.No = j.Column_Value And Mod(m.记录性质, 10) = 1 And m.记录状态 In (1, 3) And Nvl(m.费用状态, 0) <> 1)" & vbNewLine & _
            "            And Mod(a.记录性质, 10) = 1 And a.记录状态 In (1, 3) And Nvl(a.费用状态, 0) <> 1"

    If bln补结算 Then
        strSql = Replace(strSql, "门诊费用记录", "费用补充记录")
        strSql = Replace(strSql, "Max(a.结帐id)", "Max(a.结算id)")
    End If

    strSql = "" & _
            "   Select A.结帐ID,B.结算序号 " & _
            "   From (" & strSql & ") A,病人预交记录 B " & _
            "   Where A.结帐ID=B.结帐ID(+) And Rownum<2"

    If Not blnNOMoved And bln历史表同步查 Then
        strSQL1 = Replace(strSql, "门诊费用记录", "H门诊费用记录")
        strSQL1 = Replace(strSql, "费用补充记录", "H费用补充记录")
        strSQL1 = Replace(strSql, "病人预交记录", "H病人预交记录")
        strSql = strSql & " Union ALL " & strSQL1
    ElseIf blnNOMoved Then
        strSql = Replace(strSql, "门诊费用记录", "H门诊费用记录")
        strSQL1 = Replace(strSql, "费用补充记录", "H费用补充记录")
        strSql = Replace(strSql, "病人预交记录", "H病人预交记录")
    End If

    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "根据单据获取最后一次正常结帐的结帐ID", strNOs)

    If rsTemp.EOF Then Exit Function

    lng结算序号 = Val(Nvl(rsTemp!结算序号))
    GetFromNOToLastBalanceID = Val(Nvl(rsTemp!结帐ID))
    
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ExcuteInsureDel(ByVal objBalanceInfor As clsBalanceInfo, _
    ByVal intInsure As Integer, ByVal strNo As String, ByRef cllBillPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行医保退费用操作
    '入参:
    '出参:
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-08-17 16:31:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, cllPro As Collection
    Dim blnTrans As Boolean, blnTransMedicare As Boolean
    Dim strAdvance As String
    
    On Error GoTo errHandle
        
    If intInsure = 0 Then ExcuteInsureDel = True: Exit Function
    If cllBillPro Is Nothing Then Set cllBillPro = New Collection
    Set cllPro = New Collection
    
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    blnTrans = True: blnTransMedicare = False
    zlExecuteProcedureArrAy cllPro, "执行医保作废", True
    
    strAdvance = objBalanceInfor.冲销ID & "|0" & IIf(strNo <> "", "|" & strNo, "")
    If Not gclsInsure.ClinicDelSwap(objBalanceInfor.结帐ID, , intInsure, strAdvance) Then
        gcnOracle.RollbackTrans
        MsgBox "医保结算失败，无法进行门诊费用转出院操作。", vbInformation, gstrSysName
        Exit Function
    End If
    gcnOracle.CommitTrans: blnTransMedicare = True: blnTrans = False
    Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, intInsure)
    objBalanceInfor.是否保存结帐单 = True
    Set cllBillPro = New Collection
    ExcuteInsureDel = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If blnTransMedicare And mbln立即销帐 Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, False, intInsure)
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetYBBalance(ByVal lng结帐ID As Long, ByVal lng病人ID As Long, _
    Optional ByVal blnDelCheck As Boolean = True, Optional ByVal blnDel As Boolean = True, _
    Optional ByVal intInsure As Integer, Optional ByVal bln门诊结算作废 As Boolean, _
    Optional ByVal str个人帐户 As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医保原结算方式和结算金额
    '返回:返回结算信息,格式:结算方式|结算金额||...
    '编制:刘兴洪
    '日期:2014-07-07 09:57:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str结算方式 As String
    Dim strSql As String, rsData As ADODB.Recordset
    
    On Error GoTo errHandle
    strSql = _
        " Select 结算方式, Sum(冲预交) As 冲预交" & _
        " From 病人预交记录 A, 结算方式 B" & _
        " Where a.结算方式 = b.名称 And a.结帐id = [1] And b.性质 In (3, 4) And a.卡类别id Is Null" & _
        " Group By 结算方式"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng结帐ID)
    Do While Not rsData.EOF
        If blnDelCheck Then
            If bln门诊结算作废 Then
                '如果这种结算方式不支持回退,要退为现金,则不用减去
                If gclsInsure.GetCapability(support门诊结算作废, lng病人ID, intInsure, Nvl(rsData!结算方式)) Then
                    str结算方式 = str结算方式 & "||" & Nvl(rsData!结算方式) & "|" & IIf(blnDel, -1, 1) * Val(Nvl(rsData!冲预交))
                End If
            Else     '不支持门诊结算作废时,只允许个帐退为现金,其它原样退,不调用医保交易
                If Nvl(rsData!结算方式) <> str个人帐户 Then
                    str结算方式 = str结算方式 & "||" & Nvl(rsData!结算方式) & "|" & IIf(blnDel, -1, 1) * Val(Nvl(rsData!冲预交))
                End If
            End If
        Else
            str结算方式 = str结算方式 & "||" & Nvl(rsData!结算方式) & "|" & IIf(blnDel, -1, 1) * Val(Nvl(rsData!冲预交))
        End If
            
        rsData.MoveNext
    Loop
    If str结算方式 <> "" Then str结算方式 = Mid(str结算方式, 3)
    GetYBBalance = str结算方式
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ExecuteInsureDel_JZ(ByVal lng结帐ID As Long, ByVal lng病人ID As Long, _
    ByVal intInsure As Integer, ByVal str个人帐户名称 As String, _
    ByRef cllBillPro As Collection, ByRef objBalanceInfor As clsBalanceInfo) As Boolean
    '功能:执行结帐医保退费用操作
    '入参:
    '   lng结帐ID - 原结帐ID
    Dim strSql As String, blnTransMedicare As Boolean
    Dim strAdvance As String, strSavedAdvance As String
    Dim bln门诊结算作废 As Boolean
    Dim blnTrans As Boolean, cllPro As Collection
    Dim i As Integer
    
    On Error GoTo errHandle
    If intInsure = 0 Then ExecuteInsureDel_JZ = True: Exit Function
    
    If cllBillPro Is Nothing Then Set cllBillPro = New Collection
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    blnTrans = True
    bln门诊结算作废 = gclsInsure.GetCapability(support门诊结算作废, lng病人ID, intInsure)
    strSavedAdvance = GetYBBalance(lng结帐ID, lng病人ID, True, True, intInsure, bln门诊结算作废, str个人帐户名称)
    
    'Zl_病人结帐作废_Modify(
    strSql = "Zl_病人结帐作废_Modify("
    '  操作类型_In      Number,
    strSql = strSql & "" & 3 & ","
    '  病人id_In        门诊费用记录.病人id%Type,
    strSql = strSql & "" & lng病人ID & ","
    '  冲销id_In        病人预交记录.结帐id%Type,
    strSql = strSql & "" & objBalanceInfor.冲销ID & ","
    '  结算方式_In      Varchar2,
    strSql = strSql & "'" & strSavedAdvance & "')"
    zlAddArray cllPro, strSql
    
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, "执行医保作废", True
          
    If bln门诊结算作废 Then
        strAdvance = objBalanceInfor.冲销ID & "|0"
        If Not gclsInsure.ClinicDelSwap(lng结帐ID, True, intInsure, strAdvance) Then
            gcnOracle.RollbackTrans
            MsgBox "医保结算失败，无法继续进行门诊费用转住院操作。", vbInformation, gstrSysName
            Exit Function
        End If
        blnTransMedicare = True
    
        '检查结算结果是否需要校对
        If zlInsureCheck(strSavedAdvance, strAdvance) Then
            'Zl_病人结帐作废_Modify(
            strSql = "Zl_病人结帐作废_Modify("
            '  操作类型_In      Number,
            strSql = strSql & "" & 3 & ","
            '  病人id_In        门诊费用记录.病人id%Type,
            strSql = strSql & "" & lng病人ID & ","
            '  冲销id_In        病人预交记录.结帐id%Type,
            strSql = strSql & "" & objBalanceInfor.冲销ID & ","
            '  结算方式_In      Varchar2,
            strSql = strSql & "'" & strAdvance & "')"
            zlDatabase.ExecuteProcedure strSql, "校对医保数据"
        End If
    End If
    
    gcnOracle.CommitTrans: blnTrans = False
    objBalanceInfor.是否保存结帐单 = True
    If blnTransMedicare Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, intInsure)
    
    Set cllBillPro = New Collection
    ExecuteInsureDel_JZ = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If blnTransMedicare Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, False, intInsure)
    If ErrCenter() = 1 Then
        'Resume
    End If
    Call SaveErrLog
End Function

Private Function ExecuteThirdReturnMoneySwap_JZ(objPati As clsPatientInfo, ByRef objBalanceInfor As clsBalanceInfo, _
    ByRef cllBillPro As Collection) As Boolean
    '功能:执行三方卡结帐退款
    '入参:objPati-当前结算的病人信息
    '     objBalanceInfor-当前的结帐信息
    '出参:
    '返回:执行成功返回true,否则返回False
    Dim strSql As String, rsTemp As ADODB.Recordset, rsBalance As ADODB.Recordset
    Dim i As Integer, lng卡类别ID As Long, lng原结帐ID As Long, lng关联交易ID As Long
    Dim objThirdDelItems As clsBalanceItems, objItem As clsBalanceItem, objItemTemp As clsBalanceItem
    Dim objItems As clsBalanceItems, blnChangeMoney As Boolean
    Dim blnFinded As Boolean, blnSaveed As Boolean
    Dim cllPro As Collection, blnTrans As Boolean
    
    On Error GoTo errHandle
    If cllBillPro Is Nothing Then Set cllBillPro = New Collection
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    '必须先执行
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, "执行费用转出", True
    
    Set cllPro = New Collection
    
    strSql = _
        " Select 卡类别id, 结算方式, 冲预交 As 结算总额, 冲预交, 交易流水号, 交易说明," & _
        "        卡号, 关联交易id, 结算号码, 摘要, 收款时间" & _
        " From 病人预交记录 A" & _
        " Where 记录性质 = 12 And a.结帐id = [1] And a.卡类别ID Is Not Null And a.校对标志 = 1"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "查询卡结算数据", objBalanceInfor.冲销ID)
    '无三方卡处理，直接退出
    If rsTemp.RecordCount = 0 Then
        gcnOracle.RollbackTrans
        ExecuteThirdReturnMoneySwap_JZ = True: Exit Function
    End If
    
    strSql = _
        " Select Distinct a.结帐id, Nvl(a.卡类别id,0) as 卡类别id,a.交易流水号,Nvl(a.关联交易id,0) as 关联交易id " & _
        " From 病人预交记录 A, " & _
        "  (Select a.ID" & _
        "   From 病人结帐记录 A, 病人结帐记录 B" & _
        "   Where a.No = b.No And a.记录状态 In (1, 3) And b.Id = [1]) B" & _
        " Where a.结帐id = b.id And Mod(a.记录性质,10)<>1"
    Set rsBalance = zlDatabase.OpenSQLRecord(strSql, "查询结算数据", objBalanceInfor.冲销ID)
    
    Set objThirdDelItems = New clsBalanceItems
    With rsTemp
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            lng卡类别ID = Val(Nvl(rsTemp!卡类别ID))
            lng关联交易ID = Val(Nvl(rsTemp!关联交易ID))
            
            lng原结帐ID = 0
            rsBalance.Filter = "卡类别ID=" & lng卡类别ID & " and 关联交易ID=" & lng关联交易ID
            If Not rsBalance.EOF Then lng原结帐ID = Val(Nvl(rsBalance!结帐ID))
            If lng原结帐ID = 0 Then
                rsBalance.Filter = "卡类别ID=" & lng卡类别ID & " and 交易流水号='" & Nvl(!交易流水号) & "'"
                If Not rsBalance.EOF Then lng原结帐ID = Val(Nvl(rsBalance!结帐ID))
                If lng原结帐ID = 0 Then
                    If blnTrans Then gcnOracle.RollbackTrans
                    MsgBox Nvl(rsTemp!结算方式) & "未找到原始结算记录 ，请检查!", vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            End If
            objBalanceInfor.结帐ID = lng原结帐ID
            
            Set objItem = New clsBalanceItem
            With objItem
                Set .objCard = mobjThirdSwap.zlGetCardFromCardType(lng卡类别ID, False, Nvl(rsTemp!结算方式))
                .冲销ID = objBalanceInfor.冲销ID
                .结算IDs = lng原结帐ID
                .结帐ID = lng原结帐ID
                .关联交易ID = lng关联交易ID
                .交易流水号 = Nvl(rsTemp!交易流水号)
                .交易说明 = Nvl(rsTemp!交易说明)
                .结算方式 = Nvl(rsTemp!结算方式)
                .结算号码 = Nvl(rsTemp!结算号码)
                .结算摘要 = Nvl(rsTemp!摘要)
                .结算金额 = Val(Nvl(rsTemp!冲预交))
                .结算类型 = 3  '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                .结算性质 = .objCard.结算性质
                .结帐时间 = Format(rsTemp!收款时间, "yyyy-mm-dd HH:MM:SS")
                .卡号 = Nvl(rsTemp!卡号)
                .卡类别ID = lng卡类别ID
                .剩余金额 = Val(Nvl(rsTemp!冲预交))
                .未退金额 = Val(Nvl(rsTemp!冲预交))
                .原始金额 = Val(Nvl(rsTemp!冲预交))
            End With
            
            blnFinded = False
            For i = 1 To objThirdDelItems.Count
                Set objItemTemp = objThirdDelItems(i)
                If objItemTemp.卡类别ID = objItem.卡类别ID And objItemTemp.关联交易ID = objItem.关联交易ID Then
                    Set objItems = objItemTemp.objTag
                    If objItems Is Nothing Then Set objItems = New clsBalanceItems
                    objItems.AddItem objItem
                    objItems.结算金额 = objItems.结算金额 + objItem.结算金额
                    Set objThirdDelItems(i).objTag = objItems
                    objThirdDelItems.结算金额 = objThirdDelItems.结算金额 + objItem.结算金额
                    blnFinded = True: Exit For
                End If
            Next
            If Not blnFinded Then
                Set objItems = objItem.objTag
                If objItems Is Nothing Then Set objItems = New clsBalanceItems
                Set objItemTemp = objItem.Clone()
                Call objItems.AddItem(objItemTemp)
                objItems.结算金额 = objItems.结算金额 + objItem.结算金额
                Set objItem.objTag = objItems
                objThirdDelItems.AddItem objItem
                objThirdDelItems.结算金额 = objThirdDelItems.结算金额 + objItem.结算金额
            End If
            
            .MoveNext
        Loop
    End With
    
    Set rsBalance = Nothing: Set rsTemp = Nothing
   '执行三方退款
    For Each objItem In objThirdDelItems
        blnSaveed = False
        'byt操作类型-0-正常调用刷卡,1-转账调用刷卡,2-退款调用刷卡
        If mobjThirdSwap.zlThird_ReturnMoney_IsValied(objItem.objCard, 2, objBalanceInfor, objItem.objTag, objItems, False) = False Then
            If blnTrans Then gcnOracle.RollbackTrans
            If objBalanceInfor.是否保存结帐单 Then
                 Call MsgBox(objItem.objCard.名称 & "退款失败，请在病人结帐窗口中进行异常重退！", vbInformation + vbOKOnly, gstrSysName)
            End If
            Exit Function
        End If
        If mobjThirdSwap.zlThird_ReturnMoney(objPati, objItem.objCard, objBalanceInfor, objItems, cllPro, False, objItems, blnSaveed, False, blnChangeMoney, False, blnTrans) = False Then
            If blnSaveed Or objBalanceInfor.是否保存结帐单 Then
                objBalanceInfor.是否保存结帐单 = True
                Call MsgBox(objItem.objCard.名称 & "退款失败，请在病人结帐窗口中进行异常重退！", vbInformation + vbOKOnly, gstrSysName)
            Else
                Call MsgBox(objItem.objCard.名称 & "退款失败，本次门诊费用转住院失败！", vbInformation + vbOKOnly, gstrSysName)
            End If
            Exit Function
        End If
        If blnSaveed And Not objBalanceInfor.是否保存结帐单 Then objBalanceInfor.是否保存结帐单 = True
    Next
    
    If blnTrans Then gcnOracle.CommitTrans
    objBalanceInfor.是否保存结帐单 = True
    Set cllBillPro = New Collection
    ExecuteThirdReturnMoneySwap_JZ = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function DelBalaceMz(objPati As clsPatientInfo, cllBillPro As Collection, _
    ByVal lng主页ID As Long, ByVal lng入院科室ID As Long, ByRef objBalanceInfor As clsBalanceInfo) As Boolean
    '功能:记账单冲销和结帐作废
    Dim strSql As String, rsData As ADODB.Recordset
    Dim blnTrans As Boolean
    Dim intInsure As Integer
    Dim lng结帐ID As Long, strNo As String, lng病人ID As Long
    Dim strBalanceIDs As String, strBalanceNos As String
    
    On Error GoTo ErrHandler
    strSql = _
        " Select /*+cardinality(j,10)*/ Distinct b.Id As 结帐ID, b.No, c.险类, b.病人ID" & _
        " From 门诊费用记录 A, 病人结帐记录 B, 保险结算记录 C" & _
        " Where a.结帐id = b.Id And a.记录性质 In (2, 12) And a.No = [1] And b.记录状态 = 1" & _
        "       And b.ID=c.记录id(+) And c.性质(+) = 1 And c.卡类别id(+) Is Null" & _
        " Order By No"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, "查询结帐数据", objBalanceInfor.结帐单据号)
    If rsData.EOF Then
        '未结账，费用转出完成
        blnTrans = True
        zlExecuteProcedureArrAy cllBillPro, "执行结帐费用转出"
        blnTrans = False
        
        objBalanceInfor.是否保存结帐单 = True
        Set cllBillPro = New Collection
        DelBalaceMz = True
        Exit Function
    End If
    
    Do While Not rsData.EOF
        strBalanceIDs = strBalanceIDs & "," & Nvl(rsData!结帐ID)
        strBalanceNos = strBalanceNos & "," & Nvl(rsData!NO)
        rsData.MoveNext
    Loop
    
    If rsData.RecordCount > 0 Then rsData.MoveFirst
    Do While Not rsData.EOF
        With objBalanceInfor
            .结算类型 = 1  '结算类型:1-门诊结帐;2-住院结帐;3-门诊费用转住院
            .冲销ID = zlDatabase.GetNextId("病人结帐记录")
        End With
        
        lng结帐ID = Val(Nvl(rsData!结帐ID))
        strNo = Nvl(rsData!NO)
        lng病人ID = Val(Nvl(rsData!病人ID))
        intInsure = Val(Nvl(rsData!险类))
        
        If zlBillPrint_EraseBill("", lng结帐ID) = False Then Exit Function
        
        'Zl_病人结帐记录_Cancel
        strSql = "Zl_病人结帐记录_Cancel("
        '  No_In         病人结帐记录.No%Type,
        strSql = strSql & "'" & strNo & "',"
        '  冲销id_In     病人结帐记录.Id%Type,
        strSql = strSql & "'" & objBalanceInfor.冲销ID & "',"
        '  操作员编号_In 病人结帐记录.操作员编号%Type,
        strSql = strSql & "'" & UserInfo.编号 & "',"
        '  操作员姓名_In 病人结帐记录.操作员姓名%Type,
        strSql = strSql & "'" & UserInfo.姓名 & "',"
        '  作废时间_In   病人结帐记录.收费时间%Type := Null
        strSql = strSql & "" & "To_Date('" & objBalanceInfor.结帐时间 & "','YYYY-MM-DD HH24:MI:SS')" & ")"
        zlAddArray cllBillPro, strSql
        
        'Zl_门诊转住院_结帐作废
        strSql = "Zl_门诊转住院_结帐作废("
        '  No_In       病人结帐记录.No%Type,
        strSql = strSql & "'" & strNo & "',"
        '  冲销id_In   病人结帐记录.Id%Type,
        strSql = strSql & "'" & objBalanceInfor.冲销ID & "',"
        '  主页id_In     病人预交记录.主页id%Type,
        strSql = strSql & "" & ZVal(lng主页ID) & ","
        '  入院科室id_In 病人预交记录.科室id%Type,
        strSql = strSql & "" & ZVal(lng入院科室ID) & ","
        '  完成作废_In Number:=0 --0-开始结帐作废;1-完成结帐作废
        strSql = strSql & "" & 0 & ")"
        zlAddArray cllBillPro, strSql
        
        '医保退款
        If ExecuteInsureDel_JZ(lng结帐ID, lng病人ID, intInsure, mstr个人帐户, cllBillPro, objBalanceInfor) = False Then Exit Function
        
        '一卡通退款
        If ExecuteThirdReturnMoneySwap_JZ(objPati, objBalanceInfor, cllBillPro) = False Then Exit Function
        
        '完成结帐作废
        'Zl_门诊转住院_结帐作废
        strSql = "Zl_门诊转住院_结帐作废("
        '  No_In       病人结帐记录.No%Type,
        strSql = strSql & "'" & Nvl(rsData!NO) & "',"
        '  冲销id_In   病人结帐记录.Id%Type,
        strSql = strSql & "'" & objBalanceInfor.冲销ID & "',"
        '  主页id_In     病人预交记录.主页id%Type,
        strSql = strSql & "" & ZVal(lng主页ID) & ","
        '  入院科室id_In 病人预交记录.科室id%Type,
        strSql = strSql & "" & ZVal(lng入院科室ID) & ","
        '  完成作废_In Number:=0 --0-开始结帐作废;1-完成结帐作废
        strSql = strSql & "" & 1 & ")"
        zlAddArray cllBillPro, strSql
        
        '完成一次结帐作废就提交
        blnTrans = True
        zlExecuteProcedureArrAy cllBillPro, "执行结帐作废"
        blnTrans = False
        
        objBalanceInfor.是否保存结帐单 = True
        Set cllBillPro = New Collection
        
        rsData.MoveNext
    Loop
    DelBalaceMz = True
    Exit Function
ErrHandler:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ShowBalanceWindows(frmMain As Object, ByVal strDelDate As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示结帐窗口
    ' 入参:strDelDate-作废日期(主要应用于再次结帐时用预交冲)
    '编制:刘兴洪
    '日期:2011-03-29 17:38:33
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objInExse As Object
    Dim lng病人ID As Long
    
   '4.创建结帐部件
    If objInExse Is Nothing Then
        Err = 0: On Error Resume Next
        Set objInExse = CreateObject("zl9InExse.clsFeeQuery")
        If Err <> 0 Then
            MsgBox "注意:" & "在创建住院费用部件时出错,可能该部件未正常注册,结帐失败,请注意重新结帐!", vbInformation + vbOKOnly, gstrSysName
            ShowBalanceWindows = True
            Exit Function
        End If
    End If
    
    On Error GoTo errHandle
    If mlng病人ID <> 0 Then
        lng病人ID = mlng病人ID
    ElseIf Not mobjPati Is Nothing Then
        lng病人ID = mobjPati.病人ID
    End If
    
    'zlPatiBalance(ByVal frmMain As Object, _
    '    ByVal cnOracle As ADODB.Connection, ByVal lngSys As Long, strDBUser As String, _
    '    ByVal lng病人ID As Long, ByVal lng主页ID As   long ) as boolean
    If objInExse.zlPatiBalance(frmMain, gcnOracle, glngSys, gstrDBUser, lng病人ID, 0, strDelDate) = False Then
        '调用结算
    End If
    ShowBalanceWindows = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub ShowBills(ByVal lngPatient As Long, ByVal DatBegin As Date, ByVal datEnd As Date, _
    Optional ByVal blnFilter As Boolean)
'功能:读取并显示病人指定天数内的门诊费用单据
    Dim DatTmp As Date, strSql As String
    Dim strWhere As String
    Dim strFilter As String
    Dim strIDs As String, lngPre开单部门ID As Long
    Dim strVerifyWhere As String
    Dim strErrWhere As String, strBalanceErrWhere As String
    
    On Error GoTo errH
    If mrsFeeList Is Nothing Or blnFilter = False Then
        zlCommFun.ShowFlash "正在读取收费单据,请稍候 ..."
        If DatBegin > datEnd Then
            DatTmp = datEnd
            datEnd = DatBegin
            DatBegin = DatTmp
        End If
        
        '排除收费异常的单据
        strErrWhere = _
            " And Not Exists (Select 1" & _
            "     From 门诊费用记录 J1, 门诊费用记录 J2, 门诊费用记录 J3" & _
            "     Where a.No = J1.No And a.序号 = J1.序号 And J1.记录性质 = 1 And J1.记录状态 In (1,3)" & _
            "           And J1.结帐id = J2.结帐id And J1.序号 =  J2.序号" & _
            "           And J2.No = J3.No And J2.序号 =  J3.序号 And Mod(J3.记录性质,10) = 1 And Nvl(J3.费用状态,0)=1)" & vbCrLf
        strErrWhere = strErrWhere & _
            " And Not Exists(Select 1 From 费用补充记录 where 收费结帐ID=a.结帐ID And 记录性质=1 And Nvl(费用状态,0)=1) " & vbCrLf
        
        '排除结帐异常的单据
        strBalanceErrWhere = _
            " And Not Exists(Select 1" & _
            "     From 门诊费用记录 J1, 病人结帐记录 J2" & _
            "     Where J1.No = a.No And J1.序号 = a.序号 And J1.记录性质 In (2,12) And J1.结帐id = J2.Id And Nvl(J2.结算状态,0)=1)"
        
        If mbln门诊转住院先审核 Then
           strWhere = " And A.病人id = [1] "
        Else
            If datEnd - DatBegin < 4 Then   '36170
                If IDKindTime.IDKind = 1 Then
                    strWhere = " And A.病人id+0 = [1] And A.发生时间 Between [2] And [3]  "
                Else
                    strWhere = " And A.病人id+0 = [1] And A.登记时间 Between [2] And [3]  "
                End If
            Else
                If IDKindTime.IDKind = 1 Then
                    strWhere = " And A.病人id = [1] And A.发生时间+0 Between [2] And [3]  "
                Else
                    strWhere = " And A.病人id = [1] And A.登记时间+0 Between [2] And [3]  "
                End If
            End If
        End If
        
        If mbln门诊转住院先审核 Then
            strVerifyWhere = _
            " And Exists (Select 1 From 门诊费用记录 M,费用审核记录 J " & _
            "             Where M.ID=J.费用ID And M.病人ID = [1] and M.NO=A.NO And M.序号 = A.序号 And Mod(M.记录性质,10)=Mod(A.记录性质,10)  " & _
            "                   And J.审核日期 is Not NULL and  nvl(J.记录状态,0)=0 and J.性质=1) " & vbNewLine
        Else
            strVerifyWhere = _
            " And Not Exists (Select 1 From 门诊费用记录 M,费用审核记录 J " & _
            "                 Where M.ID=J.费用ID And M.病人ID = [1] and M.NO=A.NO And M.序号 = A.序号 And Mod(M.记录性质,10)=Mod(A.记录性质,10) " & _
            "                       And J.审核日期 is Not NULL and  nvl(J.记录状态,0) > 0 and J.性质=1)"
        End If
        
        strSql = strSql & _
            " Select x.选择, x.类别, x.单据, Max(Decode(Nvl(z.险类, 0),0,'','√')) As 医保,Max(z.卡类别ID) As 一卡通医保," & _
            "       x.No As 单据号, x.票据号," & vbNewLine & _
            "       x.开单人, x.开单部门ID, x.应收金额, x.实收金额, x.发生时间, Max(y.结帐id) As 结帐id," & vbNewLine & _
            "       Max(Decode(z.卡类别ID,NULL,Nvl(z.险类,0),0)) As 险类" & vbNewLine & _
            " From ( Select  '√' As 选择, '可转入' As 类别, '收费单' As 单据, a.No," & vbNewLine & _
            "               a.实际票号 As 票据号, a.开单人, a.开单部门ID, Sum(a.应收金额) As 应收金额," & vbNewLine & _
            "               Sum(a.实收金额) As 实收金额, To_Char(Max(a.发生时间), 'YYYY-MM-DD HH24:MI:SS') As 发生时间" & vbNewLine & _
            "        From 门诊费用记录 A" & vbNewLine & _
            "        Where Mod(a.记录性质, 10) = 1 And nvl(a.费用状态,0)<>1 And a.记录状态 <> 0 " & strWhere & " " & strVerifyWhere & vbCrLf & strErrWhere & _
            "              And Exists (Select 1 From 门诊费用记录 K" & _
            "                          Where k.No = a.No And k.序号 = a.序号 And k.病人id = [1] And Mod(k.记录性质, 10) = Mod(a.记录性质, 10)" & _
            "                                And Nvl(k.附加标志, 0) <> 9" & _
            "                          Group By k.序号 Having Sum(k.实收金额) <> 0)" & vbNewLine & _
            "        Group By a.No, a.实际票号, a.开单人, a.开单部门ID" & _
            "      ) X, 门诊费用记录 Y," & vbNewLine & _
            "      ( Select Distinct a.记录id, a.险类,a.卡类别ID" & vbNewLine & _
            "        From 保险结算记录 A" & vbNewLine & _
            "        Where a.性质 = 1 And a.病人id = [1]) Z" & vbNewLine & _
            " Where x.No = y.No And Mod(y.记录性质, 10) = 1 And y.记录状态 In (1, 3) And y.病人ID = [1]" & _
            "        And y.登记时间 = (Select Max(登记时间) From 门诊费用记录 Where NO = x.No And Mod(记录性质, 10) = 1 And 病人ID = [1] And 记录状态 In (1, 3)) And y.结帐id = z.记录id(+)" & _
            " Group By x.选择, x.类别, x.单据, x.No, x.票据号, x.开单人, x.开单部门ID, x.应收金额, x.实收金额, x.发生时间 "
 
        strSql = strSql & vbNewLine & _
            " Union All" & vbNewLine & _
            " Select x.选择, x.类别, x.单据, Max(Decode(Nvl(z.险类, 0),0,'','√')) As 医保,Max(z.卡类别ID) As 一卡通医保," & _
            "       x.No As 单据号, x.票据号," & vbNewLine & _
            "       x.开单人, x.开单部门ID, x.应收金额, x.实收金额, x.发生时间, Max(y.结帐id) As 结帐id," & vbNewLine & _
            "       Max(Decode(z.卡类别ID,NULL,Nvl(z.险类,0),0)) As 险类" & vbNewLine & _
            " From ( " & _
            "       Select " & vbNewLine & _
            "           '' As 选择, '不可转入' As 类别, '收费单' As 单据, a.No," & vbNewLine & _
            "           a.实际票号 As 票据号, a.开单人, a.开单部门ID, Sum(a.应收金额) As 应收金额," & vbNewLine & _
            "           Sum(a.实收金额) As 实收金额, To_Char(Max(a.发生时间), 'YYYY-MM-DD HH24:MI:SS') As 发生时间" & vbNewLine & _
            "       From 门诊费用记录 A" & vbNewLine & _
            "       Where Mod(a.记录性质, 10) = 1 And nvl(a.费用状态,0)<>1 And a.记录状态 = 3 " & strWhere & " And Nvl(a.附加标志, 0) <> 9 " & vbCrLf & strErrWhere & _
            "           And Not Exists (Select 1 From 门诊费用记录 K  Where k.No = a.No And k.序号 = a.序号 And k.病人id = [1] And Mod(k.记录性质, 10) = Mod(a.记录性质, 10) And Nvl(k.附加标志, 0) <> 9 Group By k.序号  Having Sum(k.实收金额) <> 0)" & vbNewLine & _
            "       Group By a.No, a.实际票号, a.开单人, a.开单部门ID " & _
            "       ) X, 门诊费用记录 Y," & vbNewLine & _
            "       (Select Distinct a.记录id, a.险类,a.卡类别ID" & vbNewLine & _
            "        From 保险结算记录 A" & vbNewLine & _
            "        Where a.性质 = 1 And a.病人id = [1]) Z" & vbNewLine & _
            " Where x.No = y.No And Mod(y.记录性质, 10) = 1 And y.记录状态 In (1, 3) And y.病人ID = [1]" & _
            "       And y.登记时间 = (Select Max(登记时间) From 门诊费用记录 Where NO = x.No And Mod(记录性质, 10) = 1 And 病人ID = [1] And 记录状态 In (1, 3)) And y.结帐id = z.记录id(+)" & _
            " Group By x.选择, x.类别, x.单据, x.No, x.票据号, x.开单人, x.开单部门ID, x.应收金额, x.实收金额, x.发生时间"

            
        strSql = strSql & vbNewLine & _
            "Union All" & vbNewLine & _
            "Select x.选择, x.类别, x.单据, Max(Decode(Nvl(z.险类, 0),0,'','√')) As 医保,Max(z.卡类别ID) As 一卡通医保," & _
            "       x.No As 单据号, x.票据号," & vbNewLine & _
            "       x.开单人, x.开单部门ID, x.应收金额, x.实收金额, x.发生时间, Max(y.结帐id) As 结帐id," & vbNewLine & _
            "       Max(Decode(z.卡类别ID,NULL,Nvl(z.险类,0),0)) As 险类" & vbNewLine & _
            "From (Select " & vbNewLine & _
            "        '' As 选择, '不可转入' As 类别, '收费单' As 单据, a.No," & vbNewLine & _
            "        a.实际票号 As 票据号, a.开单人, a.开单部门ID, Sum(a.应收金额) As 应收金额," & vbNewLine & _
            "        Sum(a.实收金额) As 实收金额, To_Char(Max(a.发生时间), 'YYYY-MM-DD HH24:MI:SS') As 发生时间" & vbNewLine & _
            "       From 门诊费用记录 A" & vbNewLine & _
            "       Where Mod(a.记录性质, 10) = 1 And nvl(a.费用状态,0)<>1 And a.记录状态 <> 0 " & strWhere & " " & vbCrLf & strErrWhere & _
            "           And Exists (Select 1 From 门诊费用记录 M,费用审核记录 J Where M.ID=J.费用ID And M.病人ID = [1] and M.NO=A.NO And M.序号 = A.序号 And Mod(M.记录性质,10)=Mod(A.记录性质,10) And J.审核日期 is Not NULL and  nvl(J.记录状态,0) = 1 and J.性质=1)" & _
            "           And Exists　(Select 1　 From 门诊费用记录 K　Where k.No = a.No And K.序号 = a.序号 And k.病人id = [1] And Mod(k.记录性质, 10) = Mod(a.记录性质, 10) And Nvl(k.附加标志, 0) <> 9　Group By k.序号　Having Sum(k.实收金额) <> 0)" & vbNewLine & _
            "       Group By a.No, a.实际票号, a.开单人, a.开单部门ID) X, 门诊费用记录 Y," & vbNewLine & _
            "     (  Select Distinct a.记录id, a.险类,a.卡类别ID" & vbNewLine & _
            "        From 保险结算记录 A" & vbNewLine & _
            "        Where a.性质 = 1 And a.病人id = [1]) Z" & vbNewLine & _
            " Where x.No = y.No And Mod(y.记录性质, 10) = 1 And y.记录状态 In (1, 3) And y.病人ID = [1]" & _
            " And y.登记时间 = (Select Max(登记时间) From 门诊费用记录 Where NO = x.No And Mod(记录性质, 10) = 1 And 病人ID = [1] And 记录状态 In (1, 3)) And y.结帐id = z.记录id(+)" & _
            " Group By x.选择, x.类别, x.单据, x.No, x.票据号, x.开单人, x.开单部门ID, x.应收金额, x.实收金额, x.发生时间"
     
        strSql = strSql & " UNION ALL " & _
                " Select    '√' as 选择,'可转入' as 类别,'记帐单' as 单据,'' as 医保,0 As 一卡通医保," & _
                "       A.NO As 单据号, A.实际票号 As 票据号, A.开单人, a.开单部门ID," & vbNewLine & _
                "       Sum(A.应收金额) As 应收金额, Sum(A.实收金额) As 实收金额," & vbNewLine & _
                "       To_Char(Max(A.发生时间), 'YYYY-MM-DD HH24:MI:SS') As 发生时间, 0 as 结帐ID,0 as 险类" & vbNewLine & _
                " From 门诊费用记录 A" & vbNewLine & _
                " Where A.记录性质 =2 And A.记录状态 <> 0 " & strWhere & strBalanceErrWhere & vbNewLine & _
                "       And Exists (Select 1 From 门诊费用记录 K Where K.NO=A.NO And K.序号 = A.序号 And K.记录性质=A.记录性质 And Nvl(k.附加标志, 0) <> 9 Group By K.序号 Having Sum(K.数次) <> 0) " & vbNewLine & _
                        strVerifyWhere & _
                "Group By A.NO, A.实际票号, A.开单人, a.开单部门ID "
             
        strSql = strSql & " UNION ALL " & _
            " Select C.选择,C.类别,C.单据,C.医保,c.一卡通医保,C.单据号, C.票据号, C.开单人, c.开单部门ID," & vbNewLine & _
            "       Sum(D.应收金额) As 应收金额, Sum(D.实收金额) As 实收金额, C.发生时间, C.结帐ID, C.险类" & vbNewLine & _
            " From " & _
            " (Select    '' as 选择,'不可转入' as 类别,'记帐单' as 单据,'' as 医保,0 As 一卡通医保," & _
            "       A.NO As 单据号, A.实际票号 As 票据号, A.开单人, a.开单部门ID," & vbNewLine & _
            "       Sum(A.应收金额) As 应收金额, Sum(A.实收金额) As 实收金额," & vbNewLine & _
            "       To_Char(Max(A.发生时间), 'YYYY-MM-DD HH24:MI:SS') As 发生时间,0 as 结帐ID,0 as 险类" & vbNewLine & _
            "   From 门诊费用记录  A" & vbNewLine & _
            "   Where A.记录性质 = 2 And A.记录状态 In (2,3)" & strWhere & strBalanceErrWhere & vbNewLine & _
            "       And Not Exists (Select 1 From 门诊费用记录 Where NO=A.NO And 序号 = A.序号 And 记录状态=1 And 记录性质=2) " & vbNewLine & _
            "       And Not Exists (Select 1 From 门诊费用记录 K Where K.NO=A.NO And K.序号 = A.序号 And K.记录性质=A.记录性质 And Nvl(k.附加标志, 0) <> 9 Group By K.序号 Having Sum(K.实收金额) <> 0) " & vbNewLine & _
            "   Group By A.NO, A.实际票号, A.开单人, a.开单部门ID" & _
            "   Having Sum(A.实收金额)=0) C,门诊费用记录 D Where C.单据号=D.NO And D.记录性质=2 And D.记录状态=3" & vbNewLine & _
            " Group By C.选择,C.类别,C.单据,C.医保,C.单据号, C.票据号, C.开单人, c.开单部门ID,C.发生时间, C.结帐ID, C.险类 "
            
        strSql = strSql & " UNION ALL " & _
            " Select    '' as 选择,'不可转入' as 类别,'记帐单' as 单据,'' as 医保,0 As 一卡通医保, " & _
            "       A.NO As 单据号, A.实际票号 As 票据号, A.开单人, a.开单部门ID," & vbNewLine & _
            "       Sum(A.应收金额) As 应收金额, Sum(A.实收金额) As 实收金额," & vbNewLine & _
            "       To_Char(Max(A.发生时间), 'YYYY-MM-DD HH24:MI:SS') As 发生时间, 0 as 结帐ID,0 as 险类" & vbNewLine & _
            " From 门诊费用记录 A" & vbNewLine & _
            " Where A.记录性质 = 2 And A.记录状态 <> 0 " & strWhere & strBalanceErrWhere & vbNewLine & _
            "       And Exists (Select 1 From 门诊费用记录 K Where K.NO=A.NO And K.序号 = A.序号 And K.记录性质=A.记录性质 And Nvl(k.附加标志, 0) <> 9 Group By K.序号 Having Sum(K.数次) <> 0) " & vbNewLine & _
            " And  Exists (Select 1 From 门诊费用记录 M,费用审核记录 J where M.ID=J.费用ID and M.NO=A.NO And M.序号 = A.序号 And M.记录性质=A.记录性质 And J.审核日期 is Not NULL and  nvl(J.记录状态,0) = 1 and J.性质=1) " & _
            "Group By A.NO, A.实际票号, A.开单人, a.开单部门ID "
        
        strSql = _
            " Select 选择, 类别, 单据, 医保, 一卡通医保, 单据号, 票据号, 开单人, b.名称 As 开单科室,a.开单部门ID As 开单科室ID, " & _
            "        To_Char(应收金额, '" & gSysPara.Money_Decimal.strFormt_ORA & "') As 应收金额," & _
            "        To_Char(实收金额, '" & gSysPara.Money_Decimal.strFormt_ORA & "') As 实收金额, " & _
            "       发生时间, 结帐id, 险类, 开单部门id As 开单科室ID, b.编码 As 开单科室编码" & _
            " From (" & strSql & ") A,部门表 B" & _
            " Where a.开单部门ID = b.ID" & _
            " Order By 单据,类别, 票据号 Desc, 单据号 Desc"
        '注意:由于医保要求从最后一张开始退,所以排序很关键
        Set mrsFeeList = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngPatient, DatBegin, datEnd)
    
        '加载可选科室
        mblnNotClick = True
        If cbo开单科室.ListIndex <> -1 Then lngPre开单部门ID = Val(cbo开单科室.ItemData(cbo开单科室.ListIndex))
        cbo开单科室.Clear
        cbo开单科室.AddItem "所有科室"
        Do While Not mrsFeeList.EOF
            If InStr("," & strIDs & ",", "," & Nvl(mrsFeeList!开单科室ID) & ",") = 0 Then
                strIDs = strIDs & "," & Nvl(mrsFeeList!开单科室ID)
                
                cbo开单科室.AddItem IIf(zlIsShowDeptCode, Nvl(mrsFeeList!开单科室编码) & "-", "") & Nvl(mrsFeeList!开单科室)
                cbo开单科室.ItemData(cbo开单科室.NewIndex) = Nvl(mrsFeeList!开单科室ID)
                If Val(Nvl(mrsFeeList!开单科室ID)) = lngPre开单部门ID Then cbo开单科室.ListIndex = cbo开单科室.NewIndex
            End If
            mrsFeeList.MoveNext
        Loop
        cbo.SetListWidthAuto cbo开单科室
        If cbo开单科室.ListIndex = -1 Then cbo开单科室.ListIndex = 0
        mblnNotClick = False
        
        zlCommFun.StopFlash
    End If
    
    Screen.MousePointer = vbHourglass
    strFilter = ""
    If chkShow.value = vbChecked Then strFilter = strFilter & " And  类别='可转入'"
    If Val(cbo开单科室.ItemData(cbo开单科室.ListIndex)) <> 0 Then
        strFilter = strFilter & " And 开单科室ID=" & cbo开单科室.ItemData(cbo开单科室.ListIndex)
    End If
    mrsFeeList.Filter = Mid(strFilter, 5)
    
    mshList.Redraw = flexRDNone: mshList.Clear
    mshList.Rows = 2
    Set mshList.DataSource = mrsFeeList
    If mrsFeeList.EOF Then
        sta.Panels(2).Text = "没有找到指定时间范围的收费或记帐单据!"
        mshList.Rows = 2
    Else
        sta.Panels(2).Text = "共 " & mrsFeeList.RecordCount & " 张收费单据"
    End If
    Call setHeader
    Call SetInsure
    Call SetBillColor
    mshList.Redraw = flexRDBuffered
    Call mshList_AfterRowColChange(0, 0, 1, 0)
    If mshList.Rows >= 2 Then mshList.Select 1, 0
    Call SetSumMoney
    Screen.MousePointer = vbDefault
    Exit Sub
errH:
    Screen.MousePointer = vbDefault
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SetInsure()
    Dim intInsure As Integer, lngRow As Long
    Dim str单据 As String
    
    With mshList
        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, .ColIndex("类别")) = "可转入" And .TextMatrix(lngRow, .ColIndex("选择")) = "√" Then
                intInsure = Val(.TextMatrix(lngRow, .ColIndex("险类")))
                str单据 = Trim(.TextMatrix(lngRow, .ColIndex("单据")))
                If intInsure > 0 And str单据 = "收费单" Then
                    If Not gclsInsure.GetCapability(support门诊结算作废, mlng病人ID, intInsure) Then
                        .TextMatrix(lngRow, .ColIndex("选择")) = ""
                    End If
                End If
            End If
        Next lngRow
    End With
End Sub

Private Sub dtpEnd_Change()
    dtpBegin.MaxDate = dtpEnd.value
End Sub

Private Function ExecuteThirdReturnMoneySwap(objPati As clsPatientInfo, ByRef objBalanceInfor As clsBalanceInfo, ByRef cllBillPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行三方卡退款
    '入参:objPati-当前结算的病人信息
    '     cllBillPro-当前执行的过程集
    '     objBalanceInfor-当前的结帐信息
    '出参:
    '返回:执行成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2018-08-17 15:01:44
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As ADODB.Recordset, rsBalance As ADODB.Recordset
    Dim i As Integer, lng卡类别ID As Long, lng原结帐ID As Long, bln消费卡 As Boolean, lng关联交易ID As Long, lng结算卡序号 As Long
    Dim objThirdDelItems As clsBalanceItems, objSequareDelItems As clsBalanceItems, objItem As clsBalanceItem, objItemTemp As clsBalanceItem
    Dim objItems As clsBalanceItems, blnChangeMoney As Boolean
    Dim blnFinded As Boolean, blnSaveed As Boolean
    Dim cllPro As Collection, blnTrans As Boolean
    Dim rsTotal As ADODB.Recordset
    Dim str结算信息 As String
    
    On Error GoTo errHandle
    '必须先执行后才有相关数据，所以要先执行
    If cllBillPro Is Nothing Then Set cllBillPro = New Collection
    
    Set cllPro = New Collection
    For i = 1 To cllBillPro.Count
        zlAddArray cllPro, cllBillPro(i)
    Next
    
    blnTrans = True
    zlExecuteProcedureArrAy cllPro, "执行三方退款", True
    Set cllPro = New Collection
    
    strSql = _
    " Select '' as NO, 卡类别ID,结算卡序号,结算方式,冲预交 as 结算总额,冲预交,交易流水号,交易说明,卡号,关联交易ID,结算号码,摘要,收款时间" & vbNewLine & _
    " From 病人预交记录 A" & vbNewLine & _
    " Where 记录性质 = 3 And 记录状态 = 2 and 附加标志=-1  And 结帐id = [1] and 校对标志=1 " & vbNewLine & _
    "       and Not Exists(Select 1 From 医保结算明细 where 结帐ID=[1] And A.卡类别ID=卡类别ID  And a.关联交易ID=关联交易ID )" & vbNewLine & _
    " Union all " & vbNewLine & _
    " Select distinct b.NO,A.卡类别ID,A.结算卡序号,b.结算方式,A.冲预交 as 结算总额,b.金额 as 冲预交,nvl(b.交易流水号,A.交易流水号) as 交易流水号,nvl(b.交易说明,A.交易说明) as 交易说明," & vbNewLine & _
    "        A.卡号,A.关联交易ID,A.结算号码,A.摘要,A.收款时间" & vbNewLine & _
    " From 病人预交记录 A ,医保结算明细 B" & vbNewLine & _
    " Where A.记录性质 = 3 And A.记录状态 = 2 and A.附加标志=-1  And A.结帐id = [1] and A.校对标志=1 " & vbNewLine & _
    "       and A.结帐ID=B.结帐ID And A.卡类别ID=B.卡类别ID and A.关联交易ID=B.关联交易ID and a.结算方式=b.结算方式(+) " & vbNewLine & _
    " Order by 卡类别ID,关联交易ID,NO,结算方式"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "查询结算记录", objBalanceInfor.冲销ID)
    If rsTemp.RecordCount = 0 Then '无三方及消费卡处理，直接退出
        gcnOracle.RollbackTrans
        ExecuteThirdReturnMoneySwap = True: Exit Function
    End If
    
    Set rsTotal = New ADODB.Recordset
    With rsTotal
        .Fields.Append "卡类别ID", adInteger, , adFldIsNullable
        .Fields.Append "关联交易ID", adInteger, , adFldIsNullable
        .Fields.Append "单据号", adVarChar, 20, adFldIsNullable
        .Fields.Append "卡类别名称", adVarChar, 100, adFldIsNullable
        .Fields.Append "单据总额", adDouble, , adFldIsNullable
        .Fields.Append "明细总额", adDouble, , adFldIsNullable
        .CursorLocation = adUseClient
        .LockType = adLockOptimistic
        .CursorType = adOpenStatic
        .Open
    End With
    
    strSql = " " & _
    "   Select distinct a.结帐id, nvl(a.卡类别id,0) as 卡类别id,a.交易流水号,nvl(a.结算卡序号,0) as 结算卡序号,nvl(a.关联交易id,0) as 关联交易id " & _
    "   From 病人预交记录 A, " & _
    "        (Select Distinct 结帐id " & _
    "          From 门诊费用记录 " & _
    "          Where NO In (Select Distinct NO From 门诊费用记录 Where 结帐id = [1]) And Mod(记录性质, 10) = 1 And 记录状态 In (3, 1)) B " & _
    "   Where a.结帐id = b.结帐id and mod(a.记录性质,10)<>1"
    Set rsBalance = zlDatabase.OpenSQLRecord(strSql, "查询三方退款信息", objBalanceInfor.冲销ID)
    
    Set objSequareDelItems = New clsBalanceItems
    Set objThirdDelItems = New clsBalanceItems
    
    With rsTemp
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            lng卡类别ID = Val(Nvl(rsTemp!卡类别ID))
            lng结算卡序号 = Val(Nvl(rsTemp!结算卡序号))
            bln消费卡 = lng结算卡序号 <> 0
            lng关联交易ID = Val(Nvl(rsTemp!关联交易ID))
            
            rsBalance.Filter = "卡类别ID=" & lng卡类别ID & " and 关联交易ID=" & lng关联交易ID & " and 结算卡序号=" & lng结算卡序号
            lng原结帐ID = 0
            If Not rsBalance.EOF Then lng原结帐ID = Val(Nvl(rsBalance!结帐ID))
            If lng原结帐ID = 0 And Not bln消费卡 Then
                rsBalance.Filter = "卡类别ID=" & lng卡类别ID & " and 交易流水号='" & Nvl(!交易流水号) & "'"
                If Not rsBalance.EOF Then lng原结帐ID = Val(Nvl(rsBalance!结帐ID))
                If lng原结帐ID = 0 Then
                    If blnTrans Then gcnOracle.RollbackTrans
                    MsgBox Nvl(rsTemp!结算方式) & "未找到原始结算记录 ，请检查!", vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            End If
            
            Set objItem = New clsBalanceItem
            With objItem
                Set .objCard = mobjThirdSwap.zlGetCardFromCardType(lng卡类别ID, bln消费卡, Nvl(rsTemp!结算方式))
                .冲销ID = objBalanceInfor.冲销ID
                .结算IDs = lng原结帐ID
                .结帐ID = lng原结帐ID
                .关联交易ID = lng关联交易ID
                .交易流水号 = Nvl(rsTemp!交易流水号)
                .交易说明 = Nvl(rsTemp!交易说明)
                .结算方式 = Nvl(rsTemp!结算方式)
                .结算号码 = Nvl(rsTemp!结算号码)
                .结算摘要 = Nvl(rsTemp!摘要)
                .结算金额 = Val(Nvl(rsTemp!冲预交))
                .结算类型 = IIf(bln消费卡, 5, 3)  '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
                .结算性质 = .objCard.结算性质
                .结帐时间 = Format(rsTemp!收款时间, "yyyy-mm-dd HH:MM:SS")
                .卡号 = Nvl(rsTemp!卡号)
                .卡类别ID = IIf(bln消费卡, lng结算卡序号, lng卡类别ID)
                .剩余金额 = Val(Nvl(rsTemp!冲预交))
                .未退金额 = Val(Nvl(rsTemp!冲预交))
                .原始金额 = Val(Nvl(rsTemp!冲预交))
                .消费卡 = bln消费卡
                .单据号 = Nvl(rsTemp!NO)
            End With
            If objItem.单据号 <> "" And Not objItem.消费卡 And objItem.卡类别ID <> 0 Then
                rsTotal.Filter = "卡类别ID=" & objItem.卡类别ID & " and 关联交易ID=" & objItem.关联交易ID
                If rsTotal.EOF Then
                    rsTotal.AddNew
                    rsTotal!卡类别ID = objItem.卡类别ID
                    rsTotal!关联交易ID = objItem.关联交易ID
                    'rsTotal!单据号 = objItem.单据号
                    rsTotal!卡类别名称 = IIf(objItem.objCard.名称 = "", objItem.结算方式, objItem.objCard.名称)
                End If
                If InStr(str结算信息 & ",", "," & objItem.结算方式 & ",") = 0 Then
                    str结算信息 = str结算信息 & "," & objItem.结算方式
                    rsTotal!单据总额 = Val(Nvl(rsTotal!单据总额)) + Val(Nvl(rsTemp!结算总额))
                End If
                rsTotal!明细总额 = RoundEx(Val(Nvl(rsTotal!明细总额)) + objItem.结算金额, 6)
                rsTotal.Update
            End If
            
            If objItem.消费卡 Then
                objSequareDelItems.AddItem objItem
                objSequareDelItems.结算金额 = objSequareDelItems.结算金额 + objItem.结算金额
            Else
                blnFinded = False
                For i = 1 To objThirdDelItems.Count
                    Set objItemTemp = objThirdDelItems(i)
                    If objItemTemp.卡类别ID = objItem.卡类别ID And objItemTemp.关联交易ID = objItem.关联交易ID Then
                        Set objItems = objItemTemp.objTag
                        If objItems Is Nothing Then Set objItems = New clsBalanceItems
                        objItems.AddItem objItem
                        objItems.结算金额 = objItems.结算金额 + objItem.结算金额
                        Set objThirdDelItems(i).objTag = objItems
                        objThirdDelItems.结算金额 = objThirdDelItems.结算金额 + objItem.结算金额
                        blnFinded = True
                        Exit For
                    End If
                Next
                If Not blnFinded Then
                    Set objItems = objItem.objTag
                    If objItems Is Nothing Then Set objItems = New clsBalanceItems
                    Set objItemTemp = objItem.Clone
                    Call objItems.AddItem(objItemTemp)
                    objItems.结算金额 = objItems.结算金额 + objItem.结算金额
                    Set objItem.objTag = objItems
                    objThirdDelItems.AddItem objItem
                    objThirdDelItems.结算金额 = objThirdDelItems.结算金额 + objItem.结算金额
                End If
            End If
            .MoveNext
        Loop
    End With
    
    Set rsBalance = Nothing: Set rsTemp = Nothing
    '检查医保结算明细与当前的退款总额是否一致，不一致，禁止转出
    rsTotal.Filter = 0
    With rsTotal
         If .RecordCount <> 0 Then .MoveFirst
         Do While Not .EOF
            If RoundEx(Val(Nvl(!单据总额)), 6) <> RoundEx(Val(Nvl(!明细总额)), 6) Then
                If blnTrans Then gcnOracle.RollbackTrans
                MsgBox "单据号为" & !单据号 & "的退款总额与医保结算明细中的退款金额不一致，禁止门诊费用转住院!", vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
             .MoveNext
         Loop
    End With
   
   '执行三方退款
    For Each objItem In objThirdDelItems
    
        blnSaveed = False
        'byt操作类型-0-正常调用刷卡,1-转账调用刷卡,2-退款调用刷卡
        If mobjThirdSwap.zlThird_ReturnMoney_IsValied(objItem.objCard, 2, objBalanceInfor, objItem.objTag, objItems, False) = False Then
            If blnTrans Then gcnOracle.RollbackTrans
            If objBalanceInfor.是否保存结帐单 Then
                 Call MsgBox(objItem.objCard.名称 & "退款失败，请在门诊收费窗口中进行异常重退", vbInformation + vbOKOnly, gstrSysName)
            End If
            Exit Function
        End If
        If mobjThirdSwap.zlThird_ReturnMoney(objPati, objItem.objCard, objBalanceInfor, objItems, cllPro, False, objItems, blnSaveed, False, blnChangeMoney, False, blnTrans) = False Then
            If blnSaveed Or objBalanceInfor.是否保存结帐单 Then
                objBalanceInfor.是否保存结帐单 = True
                Call MsgBox(objItem.objCard.名称 & "退款失败，请在门诊收费窗口中进行异常重退！", vbInformation + vbOKOnly, gstrSysName)
            Else
                Call MsgBox(objItem.objCard.名称 & "退款失败,本次门诊费用转住院失败！", vbInformation + vbOKOnly, gstrSysName)
            End If
            Exit Function
        End If
        If blnSaveed And Not objBalanceInfor.是否保存结帐单 Then objBalanceInfor.是否保存结帐单 = True
    Next
    If objThirdDelItems.Count = 0 Then  '消费卡，在完成时一并处理
        If blnTrans Then gcnOracle.RollbackTrans
         ExecuteThirdReturnMoneySwap = True: Exit Function
    End If
    
    If blnTrans Then gcnOracle.CommitTrans
    Set cllBillPro = New Collection
    ExecuteThirdReturnMoneySwap = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub setHeader()
    Dim strHead As String, lngCardTypeID As Long, objCard As Card
    Dim i As Long
    With mshList
        strHead = "选择,4,500|类别,4,850|单据,4,800|医保,4,500|一卡通医保,1,550|单据号,4,850|票据号,4,1100|开单人,1,800|开单科室,1,1200|" & _
            "开单科室ID,1,0|应收金额,7,850|实收金额,7,850|发生时间,4,1850|结帐ID,4,0|险类,4,0|开单科室ID,4,0|开单科室编码,4,0"
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColKey(i) = Trim(.TextMatrix(0, i))
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
             If .ColKey(i) Like "*ID" Or .ColKey(i) = "险类" Or .ColKey(i) = "开单科室编码" Then
                .ColHidden(i) = True: .ColWidth(i) = 0
             End If
             .FixedAlignment(i) = flexAlignCenterCenter
        Next
        .RowHeight(0) = 320
        
        '合并标题列“医保”
        .MergeCellsFixed = flexMergeRestrictRows
        .MergeRow(0) = True
        .TextMatrix(0, .ColIndex("一卡通医保")) = .TextMatrix(0, .ColIndex("医保"))
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .Cols - 1)
        zl_vsGrid_Para_Restore 1131, mshList, Me.Caption, "门诊转住院列表", True
         
        .ColHidden(.ColIndex("一卡通医保")) = True
        For i = 1 To .Rows - 1
            '转医疗卡类别ID为名称显示
            lngCardTypeID = Val(.TextMatrix(i, .ColIndex("一卡通医保")))
            .TextMatrix(i, .ColIndex("一卡通医保")) = ""
            If lngCardTypeID > 0 Then
                .Cell(flexcpData, i, .ColIndex("一卡通医保")) = lngCardTypeID
                If GetPayCard(lngCardTypeID, objCard) Then
                    .TextMatrix(i, .ColIndex("一卡通医保")) = objCard.名称
                End If
                .ColHidden(.ColIndex("一卡通医保")) = False
            End If
        Next
        
        .Row = 1
        .Col = 0: .ColSel = .Cols - 1
    End With
End Sub

Private Function GetPayCard(ByVal lngCardTypeID As Long, ByRef objCard As Card, _
    Optional ByVal bln仅启用 As Boolean, Optional ByVal bln消费卡 As Boolean) As Boolean
    '根据卡类别ID获取卡类别信息
    On Error GoTo ErrHandler
    Set objCard = Nothing
    
    'zlGetCard:(ByVal lngCardTypeID As Long, ByVal bln消费卡 As Boolean, ByRef objCard As Card)
    If mobjOneCardComLib.zlGetCard(lngCardTypeID, bln消费卡, objCard) = False Then Exit Function
    If Not objCard Is Nothing Then
        If bln仅启用 And Not objCard.启用 Then Set objCard = Nothing
    End If
    GetPayCard = Not objCard Is Nothing
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetBillColor()
    Dim i As Long
    
    With mshList
        For i = 1 To .Rows - 1
            .Row = i
            If .TextMatrix(i, .ColIndex("类别")) = "不可转入" Then
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &H8000000C
            Else
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = .ForeColor
            End If
        Next
    End With
End Sub

Private Sub cmdParaSet_Click()
    frmChargeTurnParSet.ShowSet Me, 1131, mstrPrivs
    mbln立即销帐 = Val(zlDatabase.GetPara("费用转出立即退费", glngSys, 1131)) = 1
End Sub

Private Sub LockScreen(blnLock As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:锁定屏幕
    '编制:刘兴洪
    '日期:2018-09-12 10:54:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnEnabled As Boolean
    
    blnEnabled = Not blnLock
    cmdOk.Enabled = blnEnabled
    cmdCancel.Enabled = blnEnabled
    cmdHelp.Enabled = blnEnabled
    cmdAll(0).Enabled = blnEnabled
    cmdAll(1).Enabled = blnEnabled
    picTop.Enabled = blnEnabled
    mshList.Enabled = blnEnabled
End Sub

Private Sub cmdOK_Click()
    Dim i As Long, strNOs As String
    Dim blnThirdAllDel As Boolean, bnYBAllDel As Boolean
    Dim lng结帐ID As Long, str单据号 As String, intInsure As Long
    Dim strReplenishNo As String, strNotSelectNos As String
    Dim strTemp As String, blnErrBill As Boolean, strErrMsg As String
    Dim cllNO As Collection, cllPati As Collection
    
    Set mcllNOs = New Collection
    If mlng病人ID = 0 Then
        MsgBox "未发现病人信息，请检查！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    zlCommFun.ShowFlash "正在准备转出数据，请稍后...", Me
    
    '直接保存
    With mshList
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("类别")) = "可转入" And .TextMatrix(i, .ColIndex("选择")) = "√" Then
            
                lng结帐ID = Val(.TextMatrix(i, .ColIndex("结帐ID")))
                str单据号 = .TextMatrix(i, .ColIndex("单据号"))
                intInsure = Val(.TextMatrix(i, .ColIndex("险类")))
                strReplenishNo = "": strNotSelectNos = ""
                blnErrBill = False
                
                If .TextMatrix(i, .ColIndex("单据")) = "收费单" Then
                    If CheckBillExistReplenishData(1, , str单据号, strReplenishNo, blnErrBill) Then
                        If blnErrBill Then
                            zlCommFun.StopFlash
                            MsgBox "单据号为[" & str单据号 & "]的记录已进行医保补充结算，但正处于异常结算状态，请先到【保险补充结算】进行处理。", vbInformation, gstrSysName
                            Exit Sub
                        End If
                        If CheckReplenishAllNosIsSelected(strReplenishNo, .TextMatrix(i, .ColIndex("单据")), strNotSelectNos) = False Then
                            zlCommFun.StopFlash
                            MsgBox "单据号为[" & str单据号 & "]的记录已进行医保补充结算，以下单据也必须一起转出：" & vbCrLf & strNotSelectNos, vbInformation, gstrSysName
                            Exit Sub
                        End If
                        '获取医保险类
                        intInsure = GetReplenishInsure(strReplenishNo)
                        If intInsure = 0 Then
                            zlCommFun.StopFlash
                            MsgBox "单据号为[" & str单据号 & "]的记录已进行医保补充结算，但未获取到医保险类,不能转出！", vbInformation, gstrSysName
                            Exit Sub
                        End If
                        '检查医保是否能够原样作废
                        strTemp = CheckInsureCancel(mlng病人ID, intInsure, strReplenishNo, True)
                        If strTemp <> "" Then
                            zlCommFun.StopFlash
                            MsgBox strTemp, vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                End If
                
                If strReplenishNo = "" Then
                    If intInsure <> 0 Then
                        '检查医保单据是否全转出
                        If IsYBSingle(str单据号, bnYBAllDel, blnThirdAllDel) = False Then
                            If CheckBalanceAllNosIsSelected(lng结帐ID, .TextMatrix(i, .ColIndex("单据")), strNOs) = False Then
                                zlCommFun.StopFlash
                                MsgBox "医保单据号为[" & str单据号 & "]的记录本次未转出全部相关结算单据,不能继续!", vbInformation, gstrSysName
                                Exit Sub
                            End If
                            
                            '检查医保分单据，三方卡为全退，目前只能禁止转出
                            If InStr(strNOs, ",") > 0 And bnYBAllDel = False And blnThirdAllDel Then
                                MsgBox "暂不支持在本次门诊转住院费用中存在医保分单据结算，但一卡通必须全退的情况，不能成功转住院的单据如下：" & vbCrLf & strNOs, vbInformation + vbOKOnly, gstrSysName
                                zlCommFun.StopFlash
                                Exit Sub
                            End If
                        End If
                    Else
                        If CheckAllTurn(str单据号) Then
                            If CheckBalanceAllNosIsSelected(lng结帐ID, .TextMatrix(i, .ColIndex("单据"))) = False Then
                                zlCommFun.StopFlash
                                MsgBox "单据号为[" & str单据号 & "]的记录本次未转出全部相关结算单据,不能继续!", vbInformation, gstrSysName
                                Exit Sub
                            End If
                        End If
                    End If
                End If
                
                '单据号,票据号,结帐ID,险类,单据类型,补结算单号,开单科室ID,开单人
                Set cllNO = New Collection
                cllNO.Add str单据号, "单据号"
                cllNO.Add .TextMatrix(i, .ColIndex("票据号")), "票据号"
                cllNO.Add lng结帐ID, "结帐ID"
                cllNO.Add intInsure, "险类"
                cllNO.Add .TextMatrix(i, .ColIndex("单据")), "单据类型"
                cllNO.Add strReplenishNo, "补结算单号"
                cllNO.Add .TextMatrix(i, .ColIndex("开单科室ID")), "开单科室ID"
                cllNO.Add .TextMatrix(i, .ColIndex("开单人")), "开单人"
                mcllNOs.Add cllNO
            End If
        Next
    End With
    
    If mcllNOs.Count = 0 Then
        zlCommFun.StopFlash
        MsgBox "你还未选择要转成住院费用的单据，不能续继！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '格式：记录性质,单据号|记录性质,单据号|... 其中，记录性质：1-门诊收费，2-门诊记帐
    strNOs = ""
    For i = 1 To mcllNOs.Count
        strNOs = strNOs & IIf(strNOs = "", "", "|")
        strNOs = strNOs & IIf(mcllNOs(i)("单据类型") = "记帐单", 2, 1) & "," & mcllNOs(i)("单据号")
    Next
    
    Set cllPati = New Collection
    cllPati.Add mobjPati.病人ID, "病人ID"
    cllPati.Add mobjPati.主页ID, "主页ID"
    cllPati.Add mobjPati.姓名, "姓名"
    cllPati.Add mobjPati.审核标志, "审核标志"
    cllPati.Add mobjPati.住院状态, "住院状态"
    If mobjExpenceSvr.zlChargeTurnCheck(strNOs, cllPati, "门诊费用转转住院检查") = False Then Exit Sub
     
    strNOs = ""
    For i = 1 To mcllNOs.Count
        If i > 60 Then strNOs = strNOs & ",...": Exit For
        strNOs = strNOs & IIf(strNOs = "", "", ",")
        strNOs = strNOs & IIf(i > 0 And i Mod 6 = 0, vbCrLf, "")
        strNOs = strNOs & mcllNOs(i)("单据号")
    Next
    If MsgBox("你确定要将如下门诊费用转成住院费用吗？" & vbCrLf & _
        strNOs, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
        zlCommFun.StopFlash
        Set mcllNOs = Nothing
        Exit Sub
    End If
    
    '不需要选择病人
    If mbln独立执行 = False Then
        mblnOk = True
        Unload Me: Exit Sub
    End If
    
    Err = 0: On Error GoTo errHand:
    If mobjPati.主页ID = 0 Then
        zlCommFun.StopFlash
        MsgBox "该病人还未入院,不能门诊费用转住院费用,请检查!", vbInformation + vbOKOnly, gstrSysName
        Exit Sub
    End If
    LockScreen True
    If ExecuteTurn(Me, mobjPati.病人ID, mobjPati.主页ID, mobjPati.住院号, _
        CDate(mobjPati.入院日期), mobjPati.当前科室id, mobjPati.当前病区id, strErrMsg) = False Then
        LockScreen False
        Set mrsFeeList = Nothing
        Call cmdRefresh_Click
        zlCommFun.StopFlash
        If strErrMsg <> "" Then MsgBox strErrMsg, vbInformation, gstrSysName
        Exit Sub
    Else
        If Val(txtPatient.Tag) <> 0 And Val(txtPatient.Tag) = mobjPati.病人ID Then mblnRefreshData = True
    End If
    zlCommFun.StopFlash
    LockScreen False
    
    If mlngModule = 1137 Then
       txtPatient.Text = ""
       Set mobjPati = Nothing
       mshDetail.Clear 1
       mshDetail.Rows = 2
       mshList.Clear 1
       mshList.Rows = 2
       vsBalance.Clear 1
       vsBalance.Rows = 2
       zlControl.ControlSetFocus txtPatient
       mlng病人ID = 0
       Exit Sub
    End If
    Unload Me
    Exit Sub
errHand:
    zlCommFun.StopFlash
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    LockScreen False
End Sub

Private Function GetReplenishAllNos(ByVal strNo As String) As String
    '获取补充结算的所有费用单据
    '返回：
    '   补充结算的所有费用单据:A001,A002,...
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim strNOs As String
    
    On Error GoTo ErrHandler
    strSql = _
        " Select Distinct a.No" & vbNewLine & _
        " From 门诊费用记录 A, 门诊费用记录 B, 费用补充记录 C" & vbNewLine & _
        " Where a.No = b.No And a.序号 = b.序号 And a.记录性质 In (1, 11)" & vbNewLine & _
        "       And b.结帐id = c.收费结帐id" & vbNewLine & _
        "       And c.记录性质 = 1 And c.附加标志 = 0 And c.No = [1]" & vbNewLine & _
        " Group By a.No, a.序号" & vbNewLine & _
        " Having Nvl(Sum(Nvl(a.付数, 1) * a.数次), 0) <> 0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNo)
    Do While Not rsTmp.EOF
        strNOs = strNOs & "," & Nvl(rsTmp!NO)
        rsTmp.MoveNext
    Loop
    If strNOs <> "" Then strNOs = Mid(strNOs, 2)
    
    GetReplenishAllNos = strNOs
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckReplenishAllNosIsSelected(ByVal strNo As String, ByVal str单据 As String, _
    Optional ByRef strNotSelectNos As String) As Boolean
    '检查补充结算的所有剩余未退费用本次是否都选择了转出
    '入参：
    '   str单据 收费单/记帐单
    '出参：
    '   strNotSelectNos 没有被选择的需要一起转出的单据
    Dim i As Integer, k As Long, blnFind As Boolean
    Dim strNOs As String, varNos As Variant
    
    On Error GoTo ErrHandler
    strNotSelectNos = ""
    strNOs = GetReplenishAllNos(strNo)
    
    varNos = Split(strNOs, ",")
    With mshList
        For i = 0 To UBound(varNos)
            blnFind = False
            For k = 1 To .Rows - 1
                If .TextMatrix(k, .ColIndex("单据")) = str单据 And .TextMatrix(k, .ColIndex("单据号")) = varNos(i) Then
                    If .TextMatrix(k, .ColIndex("类别")) = "可转入" And .TextMatrix(k, .ColIndex("选择")) = "√" Then
                        blnFind = True: Exit For
                    End If
                End If
            Next
            
            If blnFind = False Then
                strNotSelectNos = strNotSelectNos & "," & varNos(i)
            End If
        Next
    End With
    
    If strNotSelectNos <> "" Then
        strNotSelectNos = Mid(strNotSelectNos, 2)
        Exit Function
    End If
    CheckReplenishAllNosIsSelected = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetReplenishInsure(ByVal strNo As String) As Long
    '获取补充结算的医保险类
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    strSql = _
        " Select Max(b.险类) As 险类" & vbNewLine & _
        " From 病人预交记录 A, 保险结算记录 B, 费用补充记录 C" & vbNewLine & _
        " Where a.结帐id = b.记录id And a.记录性质 = 6" & vbNewLine & _
        "       And a.结帐id = c.结算id And c.记录性质 = 1" & vbNewLine & _
        "       And c.记录状态 In(1,3) And c.附加标志 = 0 And c.No = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNo)
    If Not rsTmp.EOF Then GetReplenishInsure = Nvl(rsTmp!险类)
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckBalanceAllNosIsSelected(ByVal lng结帐ID As Long, ByVal str单据 As String, _
    Optional ByRef strNos_Out As String) As Boolean
    '检查一次结算的所有剩余未退费用本次是否都选择了转出
    '入参：
    '   str单据 收费单/记帐单
    '出参:
    '   strNos_Out-当前一次结帐的剩余费用单据，多个用逗号分离
    Dim strSql As String, rsTmp As ADODB.Recordset
    Dim i As Integer, blnFind As Boolean, blnNotIsSelected As Boolean
    
    On Error GoTo ErrHandler
    strNos_Out = ""
    strSql = _
        " Select Distinct a.No" & vbNewLine & _
        " From 门诊费用记录 A, 门诊费用记录 B" & vbNewLine & _
        " Where a.No = b.No And Mod(a.记录性质,10) = Mod(b.记录性质,10)" & vbNewLine & _
        "       And a.序号=b.序号 And b.结帐id = [1]" & vbNewLine & _
        " Group By a.No,a.序号" & vbNewLine & _
        " Having Nvl(Sum(Nvl(a.付数,1)*a.数次),0) <> 0"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lng结帐ID)
    Do While Not rsTmp.EOF
        With mshList
            If blnNotIsSelected = False Then
                blnFind = False
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, .ColIndex("单据")) = str单据 And .TextMatrix(i, .ColIndex("单据号")) = Nvl(rsTmp!NO) Then
                        If .TextMatrix(i, .ColIndex("类别")) = "可转入" And .TextMatrix(i, .ColIndex("选择")) = "√" Then
                            blnFind = True: Exit For
                        End If
                    End If
                Next
                If blnFind = False Then blnNotIsSelected = True
            End If
            strNos_Out = strNos_Out & "," & Nvl(rsTmp!NO)
        End With
        rsTmp.MoveNext
    Loop
    strNos_Out = Mid(strNos_Out, 2)
    CheckBalanceAllNosIsSelected = Not blnNotIsSelected
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub dtpBegin_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub Form_Activate()
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    Call picTop_Resize
End Sub

Private Function Get个人帐户名称() As String
    '功能:获取门诊个人帐户名称
    Dim rs结算方式 As ADODB.Recordset
    
    On Error GoTo errHandle
    Set rs结算方式 = Get结算方式("收费", "3")
    If rs结算方式.EOF Then Exit Function
    
    Get个人帐户名称 = Nvl(rs结算方式!名称)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Initialize()
    Call zlBillPrint_Initialize
End Sub

Private Sub Form_Load()
    Dim strTmp As String, Datsys As Date
    Dim strValue As String
    
    If zlGetOneCardComLibObject(Me, mlngModule, mobjOneCardComLib) = False Then Unload Me: Exit Sub
    If zlGetExpenceSvrObject(mobjExpenceSvr) = False Then Unload Me: Exit Sub
    Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, mobjOneCardComLib, "", txtPatient)

    Call GetRegInFor(g私有模块, Me.Name, "idkind", strTmp)
    mintIDKind = Val(strTmp)
    If mintIDKind > 0 And mintIDKind <= IDKind.ListCount Then IDKind.IDKind = mintIDKind
    mstrTitle = Me.Caption
    
    Call RestoreWinState(Me, App.ProductName)
    IDKindTime.NotAutoAppendKind = True
    IDKindTime.IDKindStr = "发生时间|发生时间|0|0|0|0|0|0|0|0|0;登记时间|登记时间|0|0|0|0|0|0|0|0|0"
    IDKindTime.IDKind = Val(zlDatabase.GetPara("上次选择时间统计类型", glngSys, 1143, 0)) + 1
    
    mintPatientRange = Val(zlDatabase.GetPara("显示结清病人", glngSys, 1137, 0))
    mbln门诊转住院先审核 = IIf(Val(zlDatabase.GetPara("门诊转住院先审核", glngSys, 1143, 0)) = 1, True, False)
    mbln立即销帐 = Val(zlDatabase.GetPara("费用转出立即退费", glngSys, 1131)) = 1
    mstr个人帐户 = Get个人帐户名称()
    '多张单据使用多种结算方式模式
    mblnMultiBalance = zlDatabase.GetPara(79, glngSys) = "1"
    
    '票据号码长度、就诊卡号长度
    strValue = zlDatabase.GetPara(20, glngSys, , "||||")
    mbytPrepayLen = Val(Split(strValue, "|")(1))
    If mbytPrepayLen = 0 Then mbytPrepayLen = 7
    '票号严格控制
    strValue = zlDatabase.GetPara(24, glngSys, , "00000")
    mblnPrepayStrict = Mid(strValue, 2, 1) = "1"
    
    mblnNotClick = True
    chkShow.value = IIf(Val(zlDatabase.GetPara("仅显示可转入数据", glngSys, 1131, 1, Array(chkShow))) = 1, 1, 0)
    mblnNotClick = False
    picBalance.BorderStyle = 0: picList.BorderStyle = 0:    picBill.BorderStyle = 0
    
    Call InitPancel
    
    Datsys = zlDatabase.Currentdate
    strTmp = GetSetting("ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "开始时间")
    If IsDate(strTmp) Then
        dtpBegin.value = CDate(strTmp)
    Else
        dtpBegin.value = Format(DateAdd("d", -3, Datsys), "yyyy-mm-dd 00:00:00")
    End If
    dtpBegin.MaxDate = Format(Datsys, "yyyy-mm-dd 23:59:59")
    dtpEnd.value = Format(Datsys, "yyyy-mm-dd 23:59:59")
    
    Call SetVisibleCtl
    Call setHeader: Call SetDetail: Call SetBalanceHead
    
    mblnNewClinicPati = False
    If mlng病人ID = 0 Then
        Call ClearData
    Else
        If GetPatient(IDKind.GetCurCard, "-" & mlng病人ID, False, True) Then
            If IsNewClinicPati(mobjPati.挂号ID) Then '新门诊病人
                Me.Hide
                If frmChargeTurnNew.ShowMe(Me, mobjPati.挂号ID, mbln独立执行) Then
                    mblnOk = True: mblnNewClinicPati = True
                End If
                Unload Me: Exit Sub
            End If
        
            Call ShowBills(mlng病人ID, dtpBegin.value, dtpEnd.value)
        End If
    End If
    
    If mbln独立执行 = False Then
        fraPati.Visible = False: cmdOk.Visible = True
    Else
        fraPati.Visible = True: cmdOk.Visible = True
    End If
    Call picTop_Resize
End Sub

Private Function IsNewClinicPati(ByVal lng挂号ID As Long) As Boolean
    '判断是否为新门诊病人
    Dim strSql As String, rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    If lng挂号ID = 0 Then Exit Function
    strSql = "Select 1 From 病人挂号记录 Where Nvl(附加标志,0) = 3 And Id = [1] And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "判断是否为新门诊病人", lng挂号ID)
    IsNewClinicPati = Not rsTemp.EOF
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetVisibleCtl()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置控件的visible属性
    '编制:刘兴洪
    '日期:2011-03-29 21:49:52
    '---------------------------------------------------------------------------------------------------------------------------------------------
    dtpBegin.Visible = Not mbln门诊转住院先审核
    dtpEnd.Visible = Not mbln门诊转住院先审核
    lbl至.Visible = Not mbln门诊转住院先审核
    IDKindTime.Visible = Not mbln门诊转住院先审核
End Sub

Private Sub cmdCancel_Click()
    Set mcllNOs = Nothing
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name
End Sub

Private Sub cmdRefresh_Click()
    If mlng病人ID = 0 Then
        MsgBox "必须选择病人，请检查！", vbInformation, gstrSysName
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Exit Sub
    End If
    Call ShowBills(mlng病人ID, dtpBegin.value, dtpEnd.value, False)
    If cmdAll(0).Visible And cmdAll(0).Enabled Then Call cmdAll(0).SetFocus
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Terminate()
    Call zlBillPrint_Terminate
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    
    mintIDKind = IDKind.IDKind
    Call SaveRegInFor(g私有模块, Me.Name, "idkind", mintIDKind)
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "开始时间", Format(dtpBegin.value, "yyyy-MM-dd HH:mm:ss")
    SaveSetting "ZLSOFT", "公共模块\" & App.ProductName & "\" & Me.Name, "结束时间", Format(dtpEnd.value, "yyyy-MM-dd HH:mm:ss")
    Call SaveWinState(Me, App.ProductName)
    
    Call zlDatabase.SetPara("仅显示可转入数据", chkShow.value, glngSys, 1131)
    zlDatabase.SetPara "上次选择时间统计类型", IDKindTime.IDKind - 1, glngSys, 1143, InStr(1, mstrPrivs, ";参数设置;") > 0
    zlSaveDockPanceToReg Me, dkpMan, "区域"
    zl_vsGrid_Para_Save 1131, mshDetail, Me.Caption, "门诊转住院明细列表", True
    zl_vsGrid_Para_Save 1131, vsBalance, Me.Caption, "门诊转住院结算列表", True
    zl_vsGrid_Para_Save 1131, mshList, Me.Caption, "门诊转住院列表", True
    Set mrsFeeList = Nothing
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlOneCardComLib.Card)
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0
    '需要清除信息,避免刷卡后,再切换,造成密文显示失去意义
    If txtPatient.Text <> "" Then txtPatient.Text = ""
    zlControl.ControlSetFocus txtPatient
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlOneCardComLib.Card, objPatiInfor As zlOneCardComLib.clsPatientInfo, blnCancel As Boolean)
    If txtPatient.Locked Then Exit Sub
    txtPatient.Text = objPatiInfor.卡号
    If txtPatient.Text = "" Then Exit Sub
    Call FindPati(objCard, False, Trim(txtPatient.Text))
End Sub

Private Sub mshDetail_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save 1131, mshDetail, Me.Caption, "门诊转住院明细列表", True
End Sub

Private Sub mshDetail_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save 1131, mshDetail, Me.Caption, "门诊转住院明细列表", True
End Sub

Private Sub mshList_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save 1131, mshList, Me.Caption, "门诊转住院列表", True
End Sub

Private Sub mshList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strNo As String, str单据 As String
    
    If NewRow = OldRow Then Exit Sub
    With mshList
        strNo = Trim(.TextMatrix(NewRow, .ColIndex("单据号")))
        str单据 = Trim(.TextMatrix(NewRow, .ColIndex("单据")))
        If NewRow = 0 Or strNo = "" Then
            mshDetail.Clear 1: mshDetail.Rows = 2
            Call SetDetail
        Else
            Call ShowDetail(str单据, strNo)
        End If
    End With
End Sub

Private Sub mshList_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save 1131, mshList, Me.Caption, "门诊转住院列表", True
End Sub

Private Sub mshList_DblClick()
    With mshList
        If .MouseRow = 0 Then Exit Sub
        If .TextMatrix(.Row, .ColIndex("单据号")) = "" Then Exit Sub
        Call SetRowSelected(.Row, Trim(.TextMatrix(.Row, .ColIndex("选择"))) = "")
    End With
    Call SetSumMoney
End Sub

Private Sub mshList_KeyPress(KeyAscii As Integer)
     If KeyAscii <> 32 Then Exit Sub
    With mshList
        If .TextMatrix(.Row, .ColIndex("单据号")) = "" Then Exit Sub
       Call SetRowSelected(.Row, Trim(.TextMatrix(.Row, .ColIndex("选择"))) = "")
    End With
    Call SetSumMoney
End Sub

Private Sub cmdAll_Click(Index As Integer)
    Dim i As Long
    
    With mshList
        .Redraw = False
        For i = 1 To .Rows - 1
            If Index = 1 Then
                .TextMatrix(i, .ColIndex("选择")) = ""
            Else
                If Not SetRowSelected(i, Index = 0) Then
                    .Row = i: .Col = 0: .ColSel = .Cols - 1
                    Call mshList_AfterRowColChange(0, 0, .Row, .Col)
                    Exit For
                End If
            End If
        Next
        .Redraw = True
    End With
    Call SetSumMoney(Index = 1)
End Sub

Private Function CheckInsureCancel(ByVal lng病人ID As Long, ByVal lngInsure As Long, _
    ByVal strNo As String, Optional ByVal bln补结算 As Long) As String
    '检查医保是否能够原样作废
    '返回：允许原样作废，则返回空；否则，返回提示信息
    Dim strTmp As String, i As Integer
    Dim arrBalanceType As Variant, strBalanceType As String
    
    On Error GoTo ErrHandler
    If Not gclsInsure.GetCapability(support门诊结算作废, lng病人ID, lngInsure) Then
        CheckInsureCancel = IIf(bln补结算, "医保补充结算", "") & "单据[" & strNo & "]的病人险类不支持门诊结算作废，不允许转出！"
        Exit Function
    Else
        '再判断该单据的每种结算方式是否支持,正常退费时,可以退为指定结算方式,此处简化规则为不允许退费
        strTmp = GetBalanceType(strNo, bln补结算)
        arrBalanceType = Split(strTmp, ",")
        For i = 0 To UBound(arrBalanceType)
            strBalanceType = arrBalanceType(i)
            If Not gclsInsure.GetCapability(support门诊结算作废, lng病人ID, lngInsure, strBalanceType) Then
                CheckInsureCancel = IIf(bln补结算, "医保补充结算", "") & "单据[" & strNo & "]的病人险类不支持" & strBalanceType & "结算作废，不允许转出！"
                Exit Function
            End If
        Next
    End If
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SetRowSelected(ByVal lngRow As Long, blnSelect As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置一行的选择状态
    '       如果是多张单据中的一张,则还需同时设置多张中的其它单据
    '编制:刘兴洪
    '日期:2011-02-21 16:10:57
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intInsure As Integer, strNo As String, strTmp As String
    Dim str单据 As String
    Dim blnAll As Boolean
    
    With mshList
        If .TextMatrix(lngRow, .ColIndex("类别")) = "可转入" And .TextMatrix(lngRow, .ColIndex("选择")) <> IIf(blnSelect, "√", "") Then
            intInsure = Val(.TextMatrix(lngRow, .ColIndex("险类")))
            str单据 = Trim(.TextMatrix(lngRow, .ColIndex("单据")))
            strNo = .TextMatrix(lngRow, .ColIndex("单据号"))
            
            If intInsure > 0 And blnSelect And str单据 = "收费单" Then
                strTmp = CheckInsureCancel(mlng病人ID, intInsure, strNo)
                If strTmp <> "" Then
                    sta.Panels(2).Text = strTmp
                    .TextMatrix(lngRow, .ColIndex("选择")) = ""
                    Exit Function
                End If
            End If
            
            .TextMatrix(lngRow, .ColIndex("选择")) = IIf(blnSelect, "√", "")
            If str单据 = "收费单" Then
                If intInsure > 0 Then      '全部选择或取消
                    blnAll = gclsInsure.GetCapability(support多单据收费必须全退, mlng病人ID, intInsure)
                    If Not blnAll Then blnAll = Not IsYBSingle(strNo)
                    If blnAll Then If Not SetMultiOther(lngRow, blnSelect, intInsure) Then Exit Function
                    
                Else '现金病人需要处理多单据收费情况
                    If Not SetMultiOther(lngRow, blnSelect, intInsure) Then Exit Function
                End If
            End If
        End If
        If .TextMatrix(lngRow, .ColIndex("类别")) = "不可转入" Then .TextMatrix(lngRow, .ColIndex("选择")) = ""
    End With
    SetRowSelected = True
End Function

Private Function CheckAllTurn(ByVal strNo As String) As Boolean
    Dim strSql As String, rsData As ADODB.Recordset, lngCardTypeID As Long
    Dim strCardTypeIDs As String, strTemp As String
    Dim strWhere As String, objCard As Card
       
    On Error GoTo errHandle
           
    strWhere = "And  Not Exists(select 1 From 医保结算明细 Where NO=[1] And A.卡类别ID=卡类别ID and A.关联交易ID=关联交易ID) "
    
    strSql = "" & _
    "   Select A.结算方式,nvl(A.卡类别ID,0) as 卡类别ID,nvl(A.结算卡序号,0) as 结算卡序号,nvl(A.关联交易ID,0) as 关联交易ID," & _
    "       max(nvl(E.是否全退,0)) as 是否全退,nvl(max(decode(nvl(C.性质,0),3,1,4,1,0)),0) as 是否医保" & vbNewLine & _
    "   From 病人预交记录 A, " & _
    "       (   Select Distinct 结帐id  " & _
    "           From 门诊费用记录 " & _
    "           Where Mod(记录性质,10) = 1 And 记录状态 <> 0  " & _
    "                 And NO In (   Select Distinct NO  From 门诊费用记录 Where 结帐id In  (Select 结帐id" & vbNewLine & _
    "                               From 病人预交记录" & vbNewLine & _
    "                               Where 结算序号 In (Select b.结算序号" & vbNewLine & _
    "                                          From 门诊费用记录 A, 病人预交记录 B" & vbNewLine & _
    "                                          Where a.No = [1] And a.记录性质 = 1 And a.记录状态 <> 0 And a.结帐id = b.结帐id))) " & vbNewLine & _
    "                 " & _
    "         ) B,结算方式 C,消费卡类别目录 E" & vbNewLine & _
    "   Where a.结帐id = b.结帐id And a.记录性质 = 3 And A.结算方式=C.名称(+) and A.结算卡序号=E.编号(+) " & vbNewLine & _
    "       " & strWhere & vbNewLine & _
    "   Group By A.结算方式,nvl(A.卡类别ID,0),nvl(A.结算卡序号,0),nvl(A.关联交易ID,0) " & vbNewLine & _
    "   Having Sum(冲预交) <> 0" & _
    "   Order by 卡类别ID,关联交易ID"

    Set rsData = zlDatabase.OpenSQLRecord(strSql, "检查是否全退", strNo)
    If rsData.EOF Then CheckAllTurn = False: Exit Function
    
    rsData.Filter = "结算卡序号<>0 And 是否全退=1"
    If Not rsData.EOF Then CheckAllTurn = True: Exit Function   '必须全退的三方卡，不允许部分退
    
    rsData.Filter = "卡类别ID<>0 "
    Do While Not rsData.EOF
        If GetPayCard(rsData!卡类别ID, objCard) Then
            If objCard.是否全退 Then CheckAllTurn = True: Exit Function   '必须全退的三方卡，不允许部分退
        End If
        rsData.MoveNext
    Loop
    If Not rsData.EOF Then CheckAllTurn = True: Exit Function   '必须全退的三方卡，不允许部分退
    
    rsData.Filter = "是否医保 =1 And 卡类别ID<>0"
    If Not rsData.EOF Then CheckAllTurn = True: Exit Function   '一卡通含有医保结算时必须全退(在SQL中排除了分单据结算的)，不允许部分退
    
    rsData.Filter = "卡类别ID<>0"
    rsData.Sort = "卡类别ID,关联交易ID"
    
    With rsData
        strCardTypeIDs = ""
        Do While Not .EOF
            lngCardTypeID = Val(Nvl(rsData!卡类别ID))
            strTemp = lngCardTypeID & ":" & Val(Nvl(rsData!关联交易ID))
            If InStr(strCardTypeIDs & ",", "," & strTemp & ",") > 0 Then    '肯定是一卡通存在多种结算方式，所以也必须全退
                CheckAllTurn = True: Exit Function   '一卡通含有医保结算时必须全退(在SQL中排除了分单据结算的)，不允许部分退
            End If
            strCardTypeIDs = strCardTypeIDs & "," & strTemp
            .MoveNext
        Loop
    End With
    CheckAllTurn = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function SetMultiOther(ByVal lngRow As Long, blnSelect As Boolean, intInsure As Integer) As Boolean
'功能:多张单据整体选择或取消
'     如果医保多张单据要求整体退费,选择其中一张时,全选多张,取消时全取消
    Dim i As Long, j As Long, k As Long, strNo As String, strTmp As String
    Dim strBalanceType As String, arrBalanceType As Variant, blnAllTurn As Boolean
    Dim str单据 As String, strReplenishNo As String
    Dim strNOs As String, varNos As Variant
    
    With mshList
        str单据 = .TextMatrix(lngRow, .ColIndex("单据"))
        If str单据 = "记帐单" Then SetMultiOther = True: Exit Function
        If intInsure = 0 Then
            '检查是否为补结算单据
            If CheckBillExistReplenishData(1, , .TextMatrix(lngRow, .ColIndex("单据号")), strReplenishNo) Then
                strNOs = GetReplenishAllNos(strReplenishNo)
                varNos = Split(strNOs, ",")
                For i = 0 To UBound(varNos)
                    For k = 1 To .Rows - 1
                        If .TextMatrix(k, .ColIndex("单据")) = str单据 And .TextMatrix(k, .ColIndex("单据号")) = varNos(i) Then
                            .TextMatrix(k, .ColIndex("选择")) = IIf(blnSelect, "√", "")
                            Exit For
                        End If
                    Next
                Next
                SetMultiOther = True
                Exit Function
            End If
        
            blnAllTurn = CheckAllTurn(.TextMatrix(lngRow, .ColIndex("单据号")))
            
            If mblnMultiBalance Or blnAllTurn Then     '   多单据,多种结算方式
                '33635:原因是多单据且多种结算方式,不能部分退
                strNo = ""
                For k = 1 To .Rows - 1
                      If .TextMatrix(k, .ColIndex("结帐ID")) = .TextMatrix(lngRow, .ColIndex("结帐ID")) _
                            And .TextMatrix(k, .ColIndex("单据")) = str单据 _
                            And Trim(.TextMatrix(lngRow, .ColIndex("结帐ID"))) <> "" Then
                            strNo = strNo & "," & .TextMatrix(k, .ColIndex("单据号"))
                      End If
                Next
                If strNo <> "" Then strNo = Mid(strNo, 2)
                If InStr(1, strNo, ",") > 0 Then    '证明为多单据
                    For k = 1 To .Rows - 1
                          If .TextMatrix(k, .ColIndex("结帐ID")) = .TextMatrix(lngRow, .ColIndex("结帐ID")) _
                            And .TextMatrix(k, .ColIndex("单据")) = str单据 _
                            And Trim(.TextMatrix(lngRow, .ColIndex("结帐ID"))) <> "" Then
                                .TextMatrix(k, .ColIndex("选择")) = IIf(blnSelect, "√", "")
                          End If
                    Next
                End If
            End If
            SetMultiOther = True
            Exit Function
        End If
        
        If IsYBSingle(.TextMatrix(lngRow, .ColIndex("单据号"))) Then SetMultiOther = True: Exit Function
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("类别")) = "可转入" _
                And .TextMatrix(i, .ColIndex("结帐ID")) = .TextMatrix(lngRow, .ColIndex("结帐ID")) _
                And i <> lngRow Then
                If .TextMatrix(i, .ColIndex("选择")) <> .TextMatrix(lngRow, .ColIndex("选择")) Then
                   If intInsure <> 0 And blnSelect Then
                        strNo = .TextMatrix(i, .ColIndex("单据号"))
                        '判断该单据的每种结算方式是否支持,正常退费时,可以退为指定结算方式,此处简化规则为不允许退费
                         strTmp = GetBalanceType(strNo)
                         If strTmp <> "" Then
                             arrBalanceType = Split(strTmp, ",")
                             For j = 0 To UBound(arrBalanceType)
                                 strBalanceType = arrBalanceType(j)
                                 If Not gclsInsure.GetCapability(support门诊结算作废, mlng病人ID, intInsure, strBalanceType) Then
                                     sta.Panels(2).Text = "单据[" & strNo & "]的病人险类不支持" & strBalanceType & "作废,此行不允许选择转入!"
                                     For k = 1 To .Rows - 1
                                        If .TextMatrix(k, .ColIndex("结帐ID")) = .TextMatrix(i, .ColIndex("结帐ID")) _
                                            And .TextMatrix(k, .ColIndex("单据")) = str单据 Then
                                            .TextMatrix(k, .ColIndex("选择")) = ""
                                        End If
                                     Next
                                     Exit Function
                                 End If
                             Next
                         End If
                    End If
                    .TextMatrix(i, .ColIndex("选择")) = IIf(blnSelect, "√", "")
                End If
            End If
        Next
    End With
    SetMultiOther = True
End Function

Private Function GetBalanceType(ByVal strNo As String, _
    Optional ByVal bln补结算 As Boolean) As String
    '功能:获取一张单据中的医保结算方式串
    Dim rsTmp As ADODB.Recordset, strSql As String
        
    On Error GoTo errH
    If bln补结算 Then
        strSql = _
            " Select Distinct a.结算方式" & vbNewLine & _
            " From 病人预交记录 A, 结算方式 B, 费用补充记录 C" & vbNewLine & _
            " Where a.结算方式 = b.名称 And a.记录性质 = 6 And b.性质 In(3,4)" & vbNewLine & _
            "       And a.结帐id = c.结算id And c.记录性质 = 1" & vbNewLine & _
            "       And c.附加标志 = 0 And Nvl(c.费用状态, 0) <> 2 And c.No = [1]"
    Else
        strSql = _
            " Select Distinct a.结算方式" & vbNewLine & _
            " From 病人预交记录 A, 结算方式 B, 门诊费用记录 C" & vbNewLine & _
            " Where a.结算方式 = b.名称 And b.性质 In(3,4)" & vbNewLine & _
            "       And a.结帐id = c.结帐ID And c.记录性质 = 1 And c.No = [1]"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNo)
    Do While Not rsTmp.EOF
        GetBalanceType = GetBalanceType & "," & rsTmp!结算方式
        rsTmp.MoveNext
    Loop
    GetBalanceType = Mid(GetBalanceType, 2)
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ShowDetail(ByVal str单据 As String, ByVal strNo As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示明细单据
    '入参:str单据:收费单(记帐单)
    '        strNO-单据号
    '编制:刘兴洪
    '日期:2011-02-22 11:14:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTmp As ADODB.Recordset, strSql As String
    
    Err = 0: On Error GoTo errH
    If mshList.Row < 0 Then Exit Sub
    If mshList.TextMatrix(mshList.Row, mshList.ColIndex("类别")) = "可转入" Then
        strSql = _
            " Select C.名称 As 类别, max(Decode(a.是否保密, 1, '***', Nvl(E.名称, B.名称))) As 名称, " & _
            "       B.规格, A.计算单位 As 单位, Sum(Nvl(A.付数, 1) * A.数次) As 数量," & _
            "       To_Char(A.标准单价, '" & gSysPara.Price_Decimal.strFormt_ORA & "') As 单价, " & _
            "       To_Char(Sum(A.应收金额), '" & gSysPara.Money_Decimal.strFormt_ORA & "') As 应收金额," & _
            "       To_Char(Sum(A.实收金额), '" & gSysPara.Money_Decimal.strFormt_ORA & "') As 实收金额, D.名称 As 执行科室, 3 As 记录状态" & _
            " From 门诊费用记录 A, 收费项目目录 B, 收费项目类别 C, 部门表 D, 收费项目别名 E" & _
            " Where A.收费细目id = B.ID And A.收费类别 = C.编码 And A.执行部门id = D.ID(+) And A.NO = [1] And Mod(A.记录性质,10) = [2]" & _
            "      And A.记录状态 In (2,3) And A.收费细目id = E.收费细目id(+) And E.码类(+) = 1 And E.性质(+) = 3 And Nvl(a.附加标志, 0) <> 9 " & _
            " Group By A.标准单价,A.序号, C.名称, B.规格, A.计算单位, D.名称" & _
            " Having Sum(A.数次) <> 0 "
        
        strSql = strSql & " Union All" & _
            " Select C.名称 As 类别,max(Decode(a.是否保密, 1, '***', Nvl(E.名称, B.名称))) As 名称," & _
            "       B.规格, A.计算单位 As 单位, Sum(Nvl(A.付数, 1) * A.数次) As 数量," & _
            "       To_Char(A.标准单价, '" & gSysPara.Price_Decimal.strFormt_ORA & "') As 单价, " & _
            "       To_Char(Sum(A.应收金额), '" & gSysPara.Money_Decimal.strFormt_ORA & "') As 应收金额," & _
            "       To_Char(Sum(A.实收金额), '" & gSysPara.Money_Decimal.strFormt_ORA & "') As 实收金额, D.名称 As 执行科室, 1 As 记录状态" & _
            " From 门诊费用记录 A, 收费项目目录 B, 收费项目类别 C, 部门表 D, 收费项目别名 E" & _
            " Where A.收费细目id = B.ID And A.收费类别 = C.编码 And A.执行部门id = D.ID(+) And A.NO = [1] And Mod(A.记录性质,10) = [2] " & _
            "      And A.记录状态=1 And A.收费细目id = E.收费细目id(+) And E.码类(+) = 1 And E.性质(+) = 3 And Nvl(A.附加标志,0) <> 9 " & _
            " Group By A.标准单价,A.序号, C.名称, B.规格, A.计算单位, D.名称" & _
            " Having Sum(A.数次) <> 0 "
    
    ElseIf mshList.TextMatrix(mshList.Row, mshList.ColIndex("类别")) = "不可转入" Then
        strSql = _
        " Select C.名称 As 类别, max(Decode(a.是否保密, 1, '***', Nvl(E.名称, B.名称))) As 名称," & _
        "       B.规格, A.计算单位 As 单位, Sum(Nvl(A.付数, 1) * A.数次) As 数量," & _
        "       To_Char(A.标准单价, '" & gSysPara.Price_Decimal.strFormt_ORA & "') As 单价, " & _
        "       To_Char(Sum(A.应收金额), '" & gSysPara.Money_Decimal.strFormt_ORA & "') As 应收金额," & _
        "       To_Char(Sum(A.实收金额), '" & gSysPara.Money_Decimal.strFormt_ORA & "') As 实收金额, D.名称 As 执行科室, 2 As 记录状态" & _
        " From 门诊费用记录 A, 收费项目目录 B, 收费项目类别 C, 部门表 D, 收费项目别名 E" & _
        " Where A.收费细目id = B.ID And A.收费类别 = C.编码 And A.执行部门id = D.ID(+) And A.NO = [1] And Mod(A.记录性质,10) = [2] " & _
        "      And A.记录状态 In (1,3) And A.收费细目id = E.收费细目id(+) And E.码类(+) = 1 And E.性质(+) = 3 And Nvl(A.附加标志,0) <> 9 " & _
        " Group By A.标准单价,A.序号, C.名称,B.规格, A.计算单位, D.名称 Having Sum(A.数次) <> 0 "
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strNo, IIf(str单据 = "记帐单", 2, 1))
    
    mshDetail.Redraw = flexRDNone
    mshDetail.Clear
    Set mshDetail.DataSource = rsTmp
    If rsTmp.EOF Then mshDetail.Rows = 2
    Call SetDetail
    mshDetail.Redraw = flexRDBuffered
    Exit Sub
errH:
    mshDetail.Redraw = flexRDBuffered
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetDetail()
    Dim strHead As String
    Dim i As Long
    strHead = "类别,1,650|名称,1,1500|规格,1,1450|单位,4,500|数量,7,500|单价,7,850|应收金额,7,850|实收金额,7,850|执行科室,4,1000|记录状态,4,0"
    With mshDetail
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColKey(i) = .TextMatrix(0, i)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .FixedAlignment(i) = 4
        Next
        .ColHidden(9) = True
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, 9)) = 1 Then .Cell(flexcpForeColor, i, 0, i, 9) = vbBlack
            'If Val(.TextMatrix(i, 9)) = 2 Then .Cell(flexcpForeColor, i, 0, i, 9) = vbRed
            If Val(.TextMatrix(i, 9)) = 3 Then .Cell(flexcpForeColor, i, 0, i, 9) = vbBlue
        Next i
        .AutoSize 0, .Cols - 1
        zl_vsGrid_Para_Restore 1131, mshDetail, Me.Caption, "门诊转住院明细列表", True
        .RowHeight(0) = 320
        .Col = 0: .ColSel = .Cols - 1
    End With
End Sub
Private Sub SetBalanceHead()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:结算列表
    '编制:刘兴洪
    '日期:2011-03-28 11:27:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strHead As String
    Dim i As Long
    strHead = "序号,4,650|标志,1,600|结算单号,1,1500|结算金额,7,1000|结算发票,1, 2600"
    With vsBalance
        .Cols = UBound(Split(strHead, "|")) + 1
        For i = 0 To UBound(Split(strHead, "|"))
            .TextMatrix(0, i) = Split(Split(strHead, "|")(i), ",")(0)
            .ColKey(i) = .TextMatrix(0, i)
            .ColAlignment(i) = Split(Split(strHead, "|")(i), ",")(1)
            If Not Visible Then .ColWidth(i) = Split(Split(strHead, "|")(i), ",")(2)
            .FixedAlignment(i) = 4
        Next
        zl_vsGrid_Para_Restore 1131, vsBalance, Me.Caption, "门诊转住院结算列表", True
        .RowHeight(0) = 320
        .Col = 0: .ColSel = .Cols - 1
    End With
End Sub
Private Sub picBill_Resize()
    Err = 0: On Error Resume Next
    With picBill
        mshList.Left = .ScaleLeft
        mshList.Top = .ScaleTop
        mshList.Width = .ScaleWidth
        mshList.Height = .ScaleHeight
    End With
End Sub
Private Sub picBalance_Resize()
    Err = 0: On Error Resume Next
    With picBalance
        vsBalance.Left = .ScaleLeft
        vsBalance.Top = .ScaleTop
        vsBalance.Width = .ScaleWidth
        lblSum.Top = .ScaleHeight - lblSum.Height
        vsBalance.Height = lblSum.Top - mshDetail.Top
    End With
End Sub

Private Sub picBottom_Resize()
    Err = 0: On Error Resume Next
    With picBottom
            cmdCancel.Left = .ScaleLeft + .ScaleWidth - cmdCancel.Width - 400
            cmdOk.Left = cmdCancel.Left - cmdOk.Width - 20
            cmdOk.Top = cmdCancel.Top
    End With
End Sub

Private Sub picList_Resize()
    Err = 0: On Error Resume Next
    With picList
        mshDetail.Left = .ScaleLeft
        mshDetail.Top = .ScaleTop
        mshDetail.Width = .ScaleWidth
        mshDetail.Height = .ScaleHeight
    End With
End Sub

Private Sub picTop_Resize()
    Err = 0: On Error Resume Next
    If mbln独立执行 Then
        fraPati.Left = picTop.ScaleLeft + 150
        IDKindTime.Left = fraPati.Left + fraPati.Width + 20
    Else
        IDKindTime.Left = picTop.ScaleLeft + 150
    End If
    dtpBegin.Left = IDKindTime.Left + IDKindTime.Width + 30
    lbl至.Left = dtpBegin.Left + dtpBegin.Width + 50
    dtpEnd.Left = lbl至.Left + lbl至.Width + 50
    
    fraFixed.Left = fraPati.Left + IIf(mbln门诊转住院先审核 And mbln独立执行, fraPati.Width + 150, 150)
    fraFixed.Top = IIf(mbln门诊转住院先审核, 80, 450)
End Sub

Private Sub txtPatient_GotFocus()
    zlControl.TxtSelAll txtPatient
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean
    
    If txtPatient.Locked Then Exit Sub
    '病人选择器
    If Not (Len(Trim(Me.txtPatient.Text)) = 0 And KeyAscii = 13) Then
       If IDKind.GetCurCard.名称 Like "姓名*" Then
            blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
        ElseIf IDKind.GetCurCard.名称 = "门诊号" Or IDKind.GetCurCard.名称 = "住院号" Then
            If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
                If InStr("0123456789-*+", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
            End If
        Else
            txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
            '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
            txtPatient.IMEMode = 0
        End If
    End If
    
    Me.Refresh
    '刷卡完毕或输入号码后回车
    If blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 Or KeyAscii = 13 And Trim(txtPatient.Text) <> "" Then
        If KeyAscii <> 13 Then
            txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
            txtPatient.SelStart = Len(txtPatient.Text)
        End If
        KeyAscii = 0
        Call FindPati(IDKind.GetCurCard, blnCard, Trim(txtPatient.Text))
    End If
End Sub

Private Sub FindPati(ByVal objCard As Card, ByVal blnCard As Boolean, ByVal strInput As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:查找病人
    '编制:刘兴洪
    '日期:2012-10-18 16:35:27
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnOutMsg As Boolean
    
    If GetPatient(objCard, strInput, blnCard, , blnOutMsg) Then
        '69526:刘尔旋,2014-02-13,出院病人无法进行门诊转住院操作
        If Val(zlDatabase.GetPara("出院病人允许门诊转住院", glngSys, 1137, "0")) = 0 Then
            If Not mobjPati.在院 Then
                MsgBox "病人" & mobjPati.姓名 & "已经出院或还未办理住院，不允许进行门诊费用转住院操作！", vbInformation, gstrSysName
                txtPatient.Text = "": mlng病人ID = 0
                Call ClearData
                Set mobjPati = Nothing
                If txtPatient.Visible And txtPatient.Enabled Then Call txtPatient.SetFocus: zlControl.TxtSelAll txtPatient
                Exit Sub
            End If
        End If
        
        If IsNewClinicPati(mobjPati.挂号ID) Then '新门诊病人
            Me.Hide
            On Error Resume Next
            Call frmChargeTurnNew.ShowMe(Me, mobjPati.挂号ID, True)
            Err = 0: On Error GoTo 0
            
            txtPatient.Text = "": mlng病人ID = 0
            Call ClearData
            Set mobjPati = Nothing
            If txtPatient.Visible And txtPatient.Enabled Then Call txtPatient.SetFocus: zlControl.TxtSelAll txtPatient
            
            Me.Show vbModal, mfrmMain
            Exit Sub
        End If
        
        '此时会先隐式调用事件Form_Load
        Call ShowBills(mlng病人ID, dtpBegin.value, dtpEnd.value)
        If mshList.TextMatrix(1, mshList.ColIndex("单据号")) <> "" Then
            If mshList.TextMatrix(1, mshList.ColIndex("选择")) <> "" Then
                If cmdOk.Visible And cmdOk.Enabled Then Call cmdOk.SetFocus
            Else
                If cmdAll(0).Visible And cmdAll(0).Enabled Then Call cmdAll(0).SetFocus
            End If
        Else
            If txtPatient.Visible And txtPatient.Enabled Then Call txtPatient.SetFocus: zlControl.TxtSelAll txtPatient
        End If
    Else
        If Not blnOutMsg Then MsgBox "没有找到该病人,请检查输入内容是否正确！", vbInformation, gstrSysName
        txtPatient.Text = "": mlng病人ID = 0
        Call ClearData
        If txtPatient.Visible And txtPatient.Enabled Then Call txtPatient.SetFocus
    End If
    txtPatient.PasswordChar = ""
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0
End Sub

Private Sub txtPatient_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        glngTXTProc = GetWindowLong(txtPatient.hWnd, GWL_WNDPROC)
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtPatient_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        Call SetWindowLong(txtPatient.hWnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txtPatient_Validate(Cancel As Boolean)
    If mobjPati Is Nothing Then Exit Sub
    If txtPatient.Text <> mobjPati.姓名 Then txtPatient.Text = mobjPati.姓名
End Sub

Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, _
    ByVal blnCard As Boolean, Optional ByVal blnFindByPatiID As Boolean, Optional ByRef blnOutMsg As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:查找病人
    '入参:
    '   objCard=卡对象
    '   strInput=输入文本
    '   blnCard=是否刷卡
    '   blnFindByPatiID=直接按病人ID查找
    '   blnOutMsg-已经提示,不用再外部再提示
    '出参:
    '   blnCancel=用于表示输入取消
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHavePassWord As Boolean
    
    On Error GoTo errH
    Call GetPatientInfo(mobjPati, objCard, strInput, blnCard, IDKind.Cards.按缺省卡查找, IDKind.DefaultCardType, _
        Me, txtPatient, blnFindByPatiID, blnHavePassWord, blnOutMsg, mintPatientRange)
    If mobjPati Is Nothing Then Set mobjPati = Nothing: Exit Function
    
    txtPatient.Text = mobjPati.姓名: mlng病人ID = mobjPati.病人ID
    If mobjPati.入院日期 <> "" Then
        '最大设置为入院日期,不能转入住院过程中的门诊费用
        dtpEnd.MaxDate = CDate(Format(mobjPati.入院日期, "yyyy-mm-dd 23:59:59"))
        dtpEnd.value = dtpEnd.MaxDate
        dtpEnd.MaxDate = dtpEnd.MaxDate + 1
        dtpBegin.MaxDate = dtpEnd.value
        '问题: 36609 比入院时间要多一天,因为可能存在病人在没有门诊结算时,先入院,再去门诊结算,从而造成门诊费用转不了的情况.
    End If
    
    GetPatient = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Set mobjPati = Nothing
End Function
  
Private Function PrintPrePayPrint(ByVal strDelDate As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打印预交款
    '入参:strDelDate-本次转出日期
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-02-16 10:30:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTemp As ADODB.Recordset, bytPrepayPrint As Byte
    Dim strNOs As String
    Dim blnNotFirst As Boolean
    
    On Error GoTo errHandle
    If zlstr.IsHavePrivs(mstrPrivs, "预交款收据打印") = False Then
       PrintPrePayPrint = True: Exit Function '不打印
    End If
    bytPrepayPrint = Val(zlDatabase.GetPara("门诊转住院预交打印", glngSys, 1131))
    If bytPrepayPrint = 0 Then PrintPrePayPrint = True: Exit Function '不打印
    
    strSql = "Select Distinct NO From 病人预交记录 Where 记录性质 = 1 And 收款时间 = [1] And 摘要 = '门诊转住院预交'"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "获取转预交单", CDate(strDelDate))
    If rsTemp.EOF Then
        '没有转为预交数据，则也不打印
        PrintPrePayPrint = True: Exit Function
    End If
    If bytPrepayPrint = 2 Then   '提示打印
        If MsgBox("本次门诊费用转住院费用时，存在现金等结算方式转为了预交款，您是否要打印预交款票据？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            PrintPrePayPrint = True: Exit Function
        End If
    End If
    
    If Val(zlDatabase.GetPara(283, glngSys)) = 1 Then  '112862
        '1-产生的所有预交单据一次性打印
        strNOs = ""
        Do While Not rsTemp.EOF
            strNOs = strNOs & "," & Nvl(rsTemp!NO)
            rsTemp.MoveNext
        Loop
        If strNOs <> "" Then
            strNOs = Mid(strNOs, 2)
            If PrintInvoice(strNOs, strDelDate) = False Then Exit Function
        End If
    Else
        '0-按生成的预交单据分别打印
        Do While Not rsTemp.EOF
            If PrintInvoice(Nvl(rsTemp!NO), strDelDate, Not blnNotFirst) = False Then Exit Function
            blnNotFirst = True
            rsTemp.MoveNext
        Loop
    End If
    PrintPrePayPrint = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub SetSumMoney(Optional blnCls As Boolean = False)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置和显示合计
    '编制:刘兴洪
    '日期:2011-03-04 14:17:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, dblSumMoney As Double
    Dim strJzNOs As String, strSfNos As String
    With mshList
        If blnCls = False Then
            For i = .FixedRows To .Rows - 1
                If Trim(.TextMatrix(i, .ColIndex("选择"))) <> "" Then
                    dblSumMoney = dblSumMoney + Val(.TextMatrix(i, .ColIndex("实收金额")))
                End If
                If .TextMatrix(i, .ColIndex("类别")) = "可转入" And .TextMatrix(i, .ColIndex("选择")) = "√" Then
                    If .TextMatrix(i, .ColIndex("单据")) = "记帐单" Then
                        strJzNOs = strJzNOs & "," & .TextMatrix(i, .ColIndex("单据号"))
                    Else
                        strSfNos = strSfNos & "," & .TextMatrix(i, .ColIndex("单据号"))
                    End If
                End If
            Next
            If strJzNOs <> "" Then strJzNOs = Mid(strJzNOs, 2)
            If strSfNos <> "" Then strSfNos = Mid(strSfNos, 2)
        Else
            dblSumMoney = 0
        End If
    End With
    lblSum.Caption = "本次转出合计:" & Format(dblSumMoney, "###0.00;-###0.00;0.00;0.00")
    
    '加载选择的数据通信
    Call LoadBalance(strJzNOs, strSfNos)
End Sub

Private Sub dkpMan_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    Bottom = sta.Height + picBottom.Height + 100
End Sub

Private Sub LoadBalance(ByVal strJzNOs As String, ByVal strSfNos As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载结算信息
    '编制:刘兴洪
    '日期:2011-03-28 11:33:27
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String
    Dim strSFTable As String, strJzTable As String
    Dim varPara() As Variant
    
    On Error GoTo errHandle
    If strJzNOs = "" And strSfNos = "" Then
        vsBalance.Clear 1: vsBalance.Rows = 2: Exit Sub
    End If
    
    ReDim Preserve varPara(0) As Variant
    If strJzNOs <> "" Then
        If zlGetVarBoundSQL(1, strJzNOs, strJzTable, varPara, UBound(varPara) + 1) = False Then Exit Sub
        strJzTable = _
            " Select A.标志, A.NO, A.结算金额, f_List2str(Cast(COLLECT(distinct C.号码) as t_Strlist)) As 发票号 " & _
            " From (Select /*+cardinality(j,10)*/ '结帐' As 标志, B.NO, To_Char(Sum(a.结帐金额),'9999990.00') As 结算金额 " & _
            "       From 门诊费用记录 A, 病人结帐记录 B, (" & strJzTable & ") J " & _
            "       Where A.NO = J.Column_Value  And A.结帐id = B.ID  And B.记录状态=1 And A.记录性质 In (2, 12) " & _
            "       Group By B.NO) A, 票据打印内容 B, 票据使用明细 C " & _
            " Where A.NO = B.NO(+) and B.数据性质(+)=3 And B.ID = C.打印id(+) And C.性质(+)=1 " & _
            " Group By A.标志, A.NO, A.结算金额"
    End If
    
    If strSfNos <> "" Then
        If zlGetVarBoundSQL(1, strSfNos, strSFTable, varPara, UBound(varPara) + 1) = False Then Exit Sub
        strSFTable = _
            IIf(strJzNOs = "", "", " Union All") & _
            " Select A.标志, A.NO, A.结算金额, f_List2str(Cast(COLLECT(distinct C.号码) as t_Strlist))  As 发票号 " & _
            " From (Select /*+cardinality(j,10)*/ '收费' As 标志, A.NO, To_Char(Sum(a.结帐金额),'9999990.00') As 结算金额 " & _
            "       From 门诊费用记录 A, (" & strSFTable & ") J " & _
            "       Where A.NO = J.Column_Value And Mod(A.记录性质,10) = 1 " & _
            "       Group By A.NO) A, 票据打印内容 B, 票据使用明细 C " & _
            " Where A.NO = B.NO(+) and B.数据性质(+)=1 And B.ID = C.打印id(+) And C.性质(+)=1 " & _
            " Group By A.标志, A.NO, A.结算金额"
    End If
    strSql = _
        " Select Rownum As 序号, 标志, NO As 结算单号, 结算金额, 发票号 " & _
        " From (" & strJzTable & strSFTable & ")"
    Set rsTemp = zlDatabase.OpenSQLRecordByArray(strSql, Me.Caption, varPara)
    
    Set vsBalance.DataSource = rsTemp
    If rsTemp.RecordCount = 0 Then
        vsBalance.Rows = 2
    End If
    Call SetBalanceHead
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsBalance_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save 1131, vsBalance, Me.Caption, "门诊转住院结算列表", True
End Sub

Private Sub vsBalance_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save 1131, vsBalance, Me.Caption, "门诊转住院结算列表", True
End Sub

Private Function PrintInvoice(ByVal strNOs As String, ByVal strDelDate As String, Optional ByVal blnFirstBill As Boolean = True) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:发票处理
    '入参：
    '   strNos 本次打印预交单据号，格式：A001,A002,A003,...
    '   blnFirstBill 是否第一张票据，后面的不再重复提示
    '编制:刘兴洪
    '日期:2011-04-02 09:48:13
    '问题:36984
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngShareUseID As Long, lng领用ID As Long, strInvoice As String
    Dim blnInput As Boolean, blnValid As Boolean
    Dim strSql As String
    Dim intInvoiceFormat As Integer
    
    '如果严格控制票据使用
    On Error GoTo errHandle
    If mblnPrepayStrict Then
        lngShareUseID = zlDatabase.GetPara("共用预交票据批次", glngSys, 1131, 0)
        '1.严格控制票据时，根据实际的票据张数,重新检查领用ID和票据号
        lng领用ID = GetInvoiceGroupID(2, 1, lng领用ID, lngShareUseID, strInvoice, "2")
        If lng领用ID <= 0 Then
            Select Case lng领用ID
                Case -1
                    MsgBox "预交单据[" & strNOs & "]共需要1张票据!" & vbCrLf & _
                        "你没有足够的自用和共用的票据,请领用一批或设置本地共用票据后重打该单据！", vbInformation, gstrSysName
                Case -2
                    MsgBox "单据[" & strNOs & "]共需要1张票据!" & vbCrLf & _
                        "你没有足够的的共用票据,请领用一批或设置本地共用票据后重打该单据！", vbInformation, gstrSysName
                Case -3
                    MsgBox "单据[" & strNOs & "]共需要1张票据!" & vbCrLf & _
                        "票据号[" & strInvoice & "]不在可用领用批次的有效票据号范围内！" & _
                        "请重新输入有效的票据号后重打该单据！", vbInformation, gstrSysName
                Case -4
                    MsgBox "单据[" & strNOs & "]共需要1张票据!" & vbCrLf & _
                        "票据号[" & strInvoice & "]所在的领用批次没有足够的票据！" & _
                        "请先打印其它票据,用完当前领用批次后,重打该单据！", vbInformation, gstrSysName
                Case Else
                    MsgBox "票据领用信息访问失败！将来，你可以重打单据[" & strNOs & "]", vbInformation, gstrSysName
            End Select
            Exit Function
        End If
        Do
            '根据票据领用读取
            blnInput = False
            strInvoice = GetNextBill(lng领用ID)
            If strInvoice = "" Then
                '如果中途换用靠后的号码,可能造成未用完,但下一号码已超出范围
                strInvoice = UCase(InputBox("无法根据票据领用情况获取将要使用的开始票据号，" & _
                                vbCrLf & "请你输入将要使用的开始票据号码：", gstrSysName, _
                                "", Me.Left + 1500, Me.Top + 1500))
                blnInput = True
            Else
                If blnFirstBill Or strInvoice = "" Then
                    strInvoice = UCase(InputBox("请确认重打使用的开始票据号码：", gstrSysName, strInvoice, Me.Left + 1500, Me.Top + 1500))
                End If
                blnInput = True
            End If
            
            '用户取消输入,不打印
            If strInvoice = "" Then Exit Function
            '检查输入有效性
            If blnInput Then
                If GetInvoiceGroupID(2, 1, lng领用ID, lngShareUseID, strInvoice, "2") = -3 Then
                    MsgBox "你输入的票据号码不在当前领用批次的有效领用范围内,请重新输入！", vbInformation, gstrSysName
                Else
                    blnValid = True
                End If
            Else
                blnValid = True
            End If
        Loop While Not blnValid
    Else
        '有可能是第一次使用
        Do
            blnInput = False
            '非严格控制时直接从本地读取
            strInvoice = UCase(zlDatabase.GetPara("当前预交票据号", glngSys, 1131, ""))
            If strInvoice = "" Then
                 strInvoice = UCase(InputBox("没有找到已用的最大票据号码，无法确定将要使用的开始票据号。" & _
                                 vbCrLf & "请输入将要使用的开始票据号码：", gstrSysName, _
                                 "", Me.Left + 1500, Me.Top + 1500))
                 blnInput = True
            Else
                strInvoice = zlCommFun.IncStr(strInvoice)
                If blnFirstBill Or strInvoice = "" Then
                    strInvoice = UCase(InputBox("请确认重打使用的开始票据号码：", gstrSysName, strInvoice, Me.Left + 1500, Me.Top + 1500))
                End If
                blnInput = True
            End If
                 
             '用户取消输入,允许打印
             If strInvoice = "" Then
                 If MsgBox("你确定不输入票据号继续打印吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                 blnValid = True
             Else
                 '检查输入有效性
                 If blnInput Then
                     If zlCommFun.ActualLen(strInvoice) <> mbytPrepayLen Then
                         MsgBox "输入的票据号码长度应该为 " & mbytPrepayLen & " 位！", vbInformation, gstrSysName
                     Else
                         blnValid = True
                     End If
                 Else
                     blnValid = True
                 End If
            End If
        Loop While Not blnValid
    End If
    
    '执行数据处理
    'Zl_病人预交记录_Reprint
    strSql = "Zl_病人预交记录_Reprint("
    '  单据号_In Varchar2,
    strSql = strSql & "'" & strNOs & "',"
    '  票据号_In 票据使用明细.号码%Type,
    strSql = strSql & "'" & strInvoice & "',"
    '  领用id_In 票据使用明细.领用id%Type,
    strSql = strSql & "" & IIf(lng领用ID = 0, "NULL", lng领用ID) & ","
    '  使用人_In 票据使用明细.使用人%Type
    strSql = strSql & "'" & UserInfo.姓名 & "')"
    zlDatabase.ExecuteProcedure strSql, Me.Caption
    
    '输出票据
    intInvoiceFormat = Val(zlDatabase.GetPara(284, glngSys, , "0"))
    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me, _
        "NO=" & strNOs, "收款时间=" & Format(strDelDate, "yyyy-mm-dd HH:MM:SS"), _
        "病人ID=" & mlng病人ID, IIf(intInvoiceFormat = 0, "", "ReportFormat=" & intInvoiceFormat), 2)
    
    '更新本地票据
    If Not mblnPrepayStrict Then
        zlDatabase.SetPara "当前预交票据号", strInvoice, glngSys, 1131
    End If
    PrintInvoice = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetBalanceNos(ByVal bytTYPE As Byte, _
    ByVal strFindValue As String, _
    Optional ByVal blnNOMoved As Boolean, _
    Optional bln补结算 As Boolean = False) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据一张收费单据的NO或结帐ID或结帐序号，返回同一次结算的NOs
    '入参:bytType-0-根据NO来查找;1-根据结帐ID来查找,2-根据结算序号来查找
    '    strFindValue-查找的值
    '    blnNOMoved-是否在后备表中，查询单据之前的判断需要用这个参数
    '    bln补结算-是否医保补结算
    '返回:格式如"AAA,BBB,CCC,..."
    '编制:刘兴洪
    '日期:2014-06-17 17:06:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSql As String, strNOs As String

    On Error GoTo errHandle:
    Select Case bytTYPE
    Case 0 '0-根据NO来查找
        If bln补结算 Then
            strSql = "" & _
            "   Select distinct A.NO " & _
            "   From 门诊费用记录 A,(Select distinct 收费结帐ID as 结帐ID From 费用补充记录 Where NO=[1] and 记录性质=1 ) B" & _
            "   Where A.结帐ID=B.结帐ID" & _
            "   Order by NO"
        Else
            strSql = _
                "Select Distinct a.No" & vbNewLine & _
                "From 门诊费用记录 A, 门诊费用记录 B" & vbNewLine & _
                "Where a.结帐id = b.结帐id And Mod(a.记录性质, 10) = 1" & vbNewLine & _
                "      And Mod(b.记录性质, 10) = 1 And b.No = [1]" & vbNewLine & _
                "Order By NO"
        End If
    Case 1  '1-根据结帐ID来查找
        If bln补结算 Then
            strSql = "" & _
            "    Select Distinct A.No " & _
            "    From 门诊费用记录 A," & _
            "        (Select distinct C1.收费结帐ID as 结帐ID " & _
            "         From 费用补充记录 A1,费用补充记录 B1,费用补充记录 C1  " & _
            "         Where A1.结算ID=[2] and A1.记录性质=1  " & _
            "               And A1.NO=B1.NO and A1.记录性质=B1.记录性质 " & _
            "               And B1.结算序号=C1.结算序号 and C1.记录状态 in (1,3) ) B " & _
            "    Where A.结帐ID=B.结帐ID    " & _
            "    Order By NO"
        Else
            strSql = _
                "Select Distinct a.No" & vbNewLine & _
                "From 门诊费用记录 A, 门诊费用记录 B, 门诊费用记录 C" & vbNewLine & _
                "Where a.No = b.No And Mod(a.记录性质, 10) = 1" & vbNewLine & _
                "      And b.结帐id = c.结帐id And c.结帐id = [2]" & vbNewLine & _
                "Order By NO"
        End If
    Case 2  '2-根据结算序号来查找
        If bln补结算 Then
            strSql = "" & _
            "    Select Distinct A.No " & _
            "    From 门诊费用记录 A," & _
            "        (Select distinct C1.收费结帐ID as 结帐ID " & _
            "         From 费用补充记录 A1,费用补充记录 B1,费用补充记录 C1  " & _
            "         Where A1.结算序号=[2] and A1.记录性质=1  " & _
            "               And A1.NO=B1.NO and A1.记录性质=B1.记录性质 " & _
            "               And B1.结算序号=C1.结算序号 and C1.记录状态 in (1,3) ) B " & _
            "    Where A.结帐ID=B.结帐ID    " & _
            "    Order By NO"
        Else
            strSql = _
                "Select Distinct a.No" & vbNewLine & _
                "From 门诊费用记录 A, 门诊费用记录 B, 门诊费用记录 C" & vbNewLine & _
                "Where a.No = b.No And Mod(a.记录性质, 10) = 1 And b.结帐id = c.结帐id" & vbNewLine & _
                "      And c.结帐id In (Select 结帐id From 病人预交记录 Where 结算序号 = [2])" & vbNewLine & _
                "Order By NO"
        End If
    End Select
    If blnNOMoved Then
        strSql = Replace(strSql, "门诊费用记录", "H门诊费用记录")
        strSql = Replace(strSql, "病人预交记录", "H病人预交记录")
        strSql = Replace(strSql, "费用补充记录", "H费用补充记录")
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSql, "根据单据获取一次结帐的单据", strFindValue, Val(strFindValue))
    
    With rsTemp
        Do While Not .EOF
            strNOs = strNOs & "," & !NO
            .MoveNext
        Loop
    End With
    If strNOs <> "" Then strNOs = Mid(strNOs, 2)
    GetBalanceNos = strNOs
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CheckBillExistReplenishData(intTYPE As Integer, _
    Optional lngBalance As Long, Optional strNOs As String, _
    Optional ByRef strReplenishNo As String, Optional ByRef blnErrBill As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查单据是否存在二次结算
    '入参:intType:0-收费数据，使用lngBalance为结算序号
    '     intType:1-收费数据，使用strNos为单据号
    '出参：
    '   strReplenishNo 补充结算单据号
    '   blnErrBill 是否异常结算单据
    '返回:True-存在二次结算数据 False-不存在二次结算数据
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSql As String, rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    strReplenishNo = ""
    If intTYPE = 0 Then
        strSql = _
            " Select Max(a.NO) As No,Max(a.费用状态) As 费用状态" & vbNewLine & _
            " From 费用补充记录 A, (Select Distinct 结帐id From 病人预交记录 Where 结算序号 = [1]) B" & vbNewLine & _
            " Where a.收费结帐id = b.结帐id And a.记录性质 = 1 And a.附加标志 = 0 And Nvl(a.费用状态,0) <> 2"
        strSql = strSql & _
            " Union All" & _
            " Select Max(a.NO) As No,Max(a.费用状态) As 费用状态 From 费用补充记录 A Where a.结算序号 = [1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "检查二次结算", lngBalance)
    Else
        strSql = _
            " Select Max(a.NO) As No,Max(a.费用状态) As 费用状态" & vbNewLine & _
            " From 费用补充记录 A," & vbNewLine & _
            "      (Select /*+cardinality(j,10)*/Distinct a.结帐id" & vbNewLine & _
            "       From 门诊费用记录 A,Table(f_Str2list([1])) J" & vbNewLine & _
            "       Where Mod(a.记录性质, 10) = 1 And a.NO=j.Column_Value) B" & vbNewLine & _
            " Where a.收费结帐id = b.结帐id And a.记录性质 = 1 And a.附加标志 = 0 And Nvl(a.费用状态,0) <> 2 "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "检查二次结算", strNOs)
    End If
    
    strReplenishNo = Nvl(rsTmp!NO)
    blnErrBill = Val(Nvl(rsTmp!费用状态)) = 1
    CheckBillExistReplenishData = strReplenishNo <> ""
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetPatientInfo(ByRef objPati As clsPatientInfo, _
    ByVal objCard As Card, ByVal strInput As String, ByVal blnCard As Boolean, _
    ByVal blnDefaultCardFind As Boolean, ByVal lngDefaultCardTypeID As Long, _
    frmMain As Object, objText As Object, _
    Optional ByVal blnFindByPatiID As Boolean, _
    Optional ByRef blnHavePassWord As Boolean, _
    Optional ByRef blnCancel As Boolean, _
    Optional ByVal intPatientRange As Integer = -1) As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '功能：获取病人信息
    '入参：
    '   objCard=用于识别的卡对象
    '   strInput=用户输入值
    '   blnCard=是否刷卡
    '   blnDefaultCardFind=是否按缺省卡查找
    '   lngDefaultCardTypeID=缺省卡类别ID
    '   frmMain=输入控件所在窗口
    '   objText=输入控件
    '   blnFindByPatiID=直接按病人ID查找
    '   lng主页ID=指定住院次数
    '   intPatientRange-按姓名查找时,是否只显示未结费用的病人,0-含已结清,1-未结清,2-体检未结清,3-住院未结清
    '出参：
    '   blnHavePassWord=是否需要密码验证
    '返回：返回病人信息
    '说明：
    '------------------------------------------------------------------------------------------------------------------------
    Dim lng病人ID As Long, lng主页ID As Long
    Dim varCardType As Variant, blnFind As Boolean
    Dim strCardPass As String, lng卡类别ID As Long
    
    On Error GoTo ErrHandler
    blnHavePassWord = False
    
    blnFind = False
    If blnCard And objCard.名称 Like "姓名*" And InStr("-+*", Left(strInput, 1)) = 0 Then '103563,缺省卡类别
        If blnDefaultCardFind And lngDefaultCardTypeID > 0 Then
            varCardType = lngDefaultCardTypeID
        Else
            varCardType = -1
        End If
        blnHavePassWord = True
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Or blnFindByPatiID Then  '病人ID
        lng病人ID = Mid(strInput, 2)
        blnFind = True
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '门诊号
        strInput = Mid(strInput, 2)
        varCardType = "门诊号"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then '住院号(病人在院)
        strInput = zlFormatID(Mid(strInput, 2))
        If mobjOneCardComLib.zlGetPatiIDFromInpatientNum(strInput, lng病人ID, , , lng主页ID) = False Then GoTo NotFoundPati:
        If lng病人ID <= 0 Then GoTo NotFoundPati:
        blnFind = True
    ElseIf Left(strInput, 1) = "^" And IsNumeric(Mid(strInput, 2)) Then '留观号
        strInput = Mid(strInput, 2)
        varCardType = "留观号"
    Else
        If Not objPati Is Nothing Then
            If objPati.姓名 = strInput Then GetPatientInfo = True: Exit Function
        End If
        
        Select Case objCard.名称
        Case "姓名", "姓名或就诊卡"
            If GetPatiIdFromPatiName(objText, strInput, lng病人ID, frmMain, blnCancel, intPatientRange) = False Then GoTo NotFoundPati:
            strInput = lng病人ID
            blnFind = True
        Case "医保号"
            strInput = UCase(strInput)
            varCardType = objCard.名称
        Case "门诊号"
            If Not IsNumeric(strInput) Then GoTo NotFoundPati:
            varCardType = objCard.名称
        Case "住院号"
            If Not IsNumeric(strInput) Then GoTo NotFoundPati:
            strInput = zlFormatID(strInput)
            If mobjOneCardComLib.zlGetPatiIDFromInpatientNum(strInput, lng病人ID, , , lng主页ID) = False Then GoTo NotFoundPati:
            If lng病人ID <= 0 Then GoTo NotFoundPati:
            blnFind = True
        Case "留观号"
            If Not IsNumeric(strInput) Then GoTo NotFoundPati:
            varCardType = objCard.名称
        Case Else
            If objCard.接口序号 > 0 Then
                varCardType = objCard.接口序号
            Else
                varCardType = objCard.名称
            End If
            blnHavePassWord = True
        End Select
    End If
    
    If blnFind = False Then
        If mobjOneCardComLib.zlGetPatiID(varCardType, strInput, , lng病人ID, strCardPass, , lng卡类别ID) = False Then GoTo NotFoundPati:
        If lng病人ID > 0 Then blnFind = True
    End If
    If blnFind = False Then GoTo NotFoundPati:
    
    Set objPati = GetPatiInfo(lng病人ID, lng主页ID)
    If objPati Is Nothing Then GoTo NotFoundPati:
    
    objPati.密码 = strCardPass
    GetPatientInfo = True
    Exit Function
NotFoundPati:
    Set objPati = Nothing
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    Set objPati = Nothing
End Function

Private Function GetPatiIdFromPatiName(ByVal objControl As Object, ByVal strName As String, ByRef lngPatiID As Long, _
    Optional frmMain As Object, Optional ByRef blnCancel As Boolean, Optional ByVal intPatientRange As Integer = -1) As Boolean
    '功能:根据病人姓名，获取病人信息
    '入参:
    '   objControl-调用的控件
    '   strName-输入的病人信息
    '   frmMain-调用的主窗体
    '   intPatientRange-按姓名查找时,是否只显示未结费用的病人,0-含已结清,1-未结清,2-体检未结清,3-住院未结清
    '出参：
    '   lngPatiId=选择的病人ID
    '   blnCancel=是否用户取消选择
    '返回:成功返回true,否则返回False
    '说明:仅结帐时调用
    Dim rsPati As ADODB.Recordset
    Dim i As Long
    Dim strSql As String, strWhere As String
    Dim cllFilter As Collection, rsPatiPageInfo As ADODB.Recordset
    Dim str病人IDs As String, strSubTable As String, varPara() As Variant
    Dim rsFee As ADODB.Recordset
    Dim str主页IDs As String
    Dim vRect As RECT, rsOutSel As ADODB.Recordset
    
    On Error GoTo ErrHandler
    If mobjOneCardComLib.zlGetPatiRecordFromPatiName(strName, rsPati) = False Then Exit Function
    If rsPati.RecordCount = 0 Then Exit Function
    
    '   rsPati-病人信息集,字段：排序ID,ID,病人ID,主页ID,姓名,性别,年龄,病人类型,险类,门诊号,住院号,科室,
    '                           床号,出生日期,身份证号,家庭地址,工作单位,在院标志,手机号,是否有医疗卡,就诊时间,入院日期,出院日期
    Do While Not rsPati.EOF
        str病人IDs = str病人IDs & "," & Nvl(rsPati!病人ID)
        If Val(Nvl(rsPati!主页ID)) <> 0 Then
            str主页IDs = str主页IDs & "," & Nvl(rsPati!病人ID) & ":" & Nvl(rsPati!主页ID)
        End If
        rsPati.MoveNext
    Loop
    rsPati.MoveFirst
    
    If str主页IDs <> "" Then
        Set cllFilter = New Collection
        cllFilter.Add Array("主页IDS", Mid(str主页IDs, 2))
        If GetPatiPageInfByRange(cllFilter, rsPatiPageInfo) = False Then Exit Function
    End If
    
    If intPatientRange >= 0 Then
        '获取存在未结费用的病人
        str病人IDs = Mid(str病人IDs, 2)
        If zlGetVarBoundSQL(0, str病人IDs, strSubTable, varPara, 0) = False Then Exit Function
        
        Select Case intPatientRange
        Case 1  '任何费用未结清病人
            strWhere = ""
        Case 2  '体检未结清的病人
            strWhere = " And a.来源途径 = 4"
        Case 3  '住院未结清的病人
            strWhere = " And a.来源途径 = 2"
        Case 4  '门诊未结清的病人
            strWhere = " And a.来源途径 = 1"
        End Select
        strSql = "Select a.病人ID" & _
                " From 病人未结费用 A,(" & strSubTable & ") B" & _
                " Where a.病人ID=b.Column_Value" & strWhere & _
                " Group By a.病人ID"
        Set rsFee = zlDatabase.OpenSQLRecordByArray(strSql, "查询存在未结费用的病人", varPara)
        
        For i = rsPati.RecordCount To 1 Step -1
            rsFee.Filter = "病人ID=" & Nvl(rsPati!病人ID)
            If rsFee.EOF Then
                rsPati.Delete adAffectCurrent
            ElseIf Not rsPatiPageInfo Is Nothing Then
                rsPatiPageInfo.Filter = "病人ID=" & Nvl(rsPati!病人ID)
                If Not rsPatiPageInfo.EOF Then
                    rsPati!入院日期 = Format(Nvl(rsPatiPageInfo!入院时间), "yyyy-MM-dd")
                    rsPati!出院日期 = Format(Nvl(rsPatiPageInfo!出院时间), "yyyy-MM-dd")
                End If
            End If
            rsPati.MoveNext
        Next
    End If
    
    rsPati.Sort = "在院标志 Desc,入院日期 Desc"
    If rsPati.RecordCount = 0 Then Exit Function
    If rsPati.RecordCount = 1 Then
        lngPatiID = Val(rsPati!病人ID)
        GetPatiIdFromPatiName = True: Exit Function
    End If
    
    vRect = zlControl.GetControlRect(objControl.hWnd)
    Set rsOutSel = zlDatabase.ShowRecSelect(frmMain, rsPati, 0, "病人选择器", _
        False, "", "请选择病人", False, False, True, vRect.Left, vRect.Top, objControl.Height, _
        blnCancel, False, True, 1, "排序ID,ID,主页ID,病人类型,险类,床号,手机号,是否有医疗卡,就诊时间")
    If rsOutSel Is Nothing Then Exit Function
    If rsOutSel.EOF Then Exit Function
    
    lngPatiID = Val(rsOutSel!病人ID)
    GetPatiIdFromPatiName = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetPatiInfoByPage(objPati As clsPatientInfo, _
    Optional ByVal lng主页ID As Long, Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:从病案主页中获取病人信息
    '入参:
    '   objPati-已有病人信息
    '   lng主页ID-主页ID，为0时，取最后一次住院的
    '出参:
    '   objPati-返回病人信息对象
    '返回:成功返回True，否则返回False
    '说明:如果传入 objPati 不为Nothing，则进行信息合并
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim objService As zlPublicExpense.clsService
    
    On Error GoTo errHandle
    If objPati Is Nothing Then Exit Function
    If objPati.病人ID = 0 Then Exit Function
    If zlGetServiceObject(objService) = False Then Exit Function
    
    If lng主页ID = 0 Then lng主页ID = objPati.主页ID '病人信息中的主页ID为最后一次主页ID
    If objService.ZlCissvr_GetPatiPageInfo(1, objPati.病人ID & ":" & lng主页ID, rsTemp, , , lngModule) = False Then Exit Function
    If rsTemp Is Nothing Then GetPatiInfoByPage = True: Exit Function
    If rsTemp.EOF Then GetPatiInfoByPage = True: Exit Function
    
    If objPati Is Nothing Then Set objPati = New clsPatientInfo
    With objPati
        .主页ID = Nvl(rsTemp!主页ID)
        .姓名 = Nvl(rsTemp!姓名)
        .性别 = Nvl(rsTemp!性别)
        .年龄 = Nvl(rsTemp!年龄)
        .费别 = Nvl(rsTemp!费别)
        .医疗付款方式 = Nvl(rsTemp!医疗付款方式名称)
        .医疗付款方式编码 = Nvl(rsTemp!医疗付款方式编码)
        .险类 = Val(Nvl(rsTemp!险类))
        .险类名称 = GetInsureName(Val(Nvl(rsTemp!险类)))
        .病人类型 = Nvl(rsTemp!病人类型)
        .当前病区id = Val(Nvl(rsTemp!当前病区id))
        .当前病区名称 = Nvl(rsTemp!当前病区名称)
        .当前科室id = Val(Nvl(rsTemp!当前科室id))
        .当前科室名称 = Nvl(rsTemp!当前科室名称)
        .床号 = Nvl(rsTemp!当前床号)
        .住院号 = Nvl(rsTemp!住院号)
        .病人性质 = Val(Nvl(rsTemp!病人性质))
        .入院日期 = Nvl(rsTemp!入院时间)
        .出院日期 = Nvl(rsTemp!出院时间)
        .住院医师 = Nvl(rsTemp!住院医师)
        .病人备注 = Nvl(rsTemp!病人备注)
        .住院状态 = Val(Nvl(rsTemp!住院状态))
        .审核标志 = Val(Nvl(rsTemp!审核标志))
        .编目日期 = Nvl(rsTemp!编目日期)
        .医保号 = Nvl(rsTemp!医保号)
        .挂号ID = Val(Nvl(rsTemp!挂号ID))
    End With
    GetPatiInfoByPage = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetPatiInfo(ByVal lng病人ID As Long, _
    Optional ByVal lng主页ID As Long, Optional ByVal lngModule As Long) As clsPatientInfo
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息，先从病人信息中获取，再从病案主页中获取进行合并
    '入参:
    '   objPati-已有病人信息
    '   lng主页ID-主页ID，为0时，取最后一次住院的
    '出参:
    '   objPati-返回病人信息对象
    '返回:成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objPati As clsPatientInfo
    
    On Error GoTo errHandle
    '读取病人信息
    If mobjOneCardComLib.zlGetPatiInforFromPatiID(lng病人ID, objPati) = False Then Exit Function
    If objPati Is Nothing Then Exit Function
    
    '2.读取病案主页
    If lng主页ID = 0 Then lng主页ID = objPati.主页ID
    If GetPatiInfoByPage(objPati, lng主页ID, lngModule) = False Then Exit Function
    
    Set GetPatiInfo = objPati
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetPatiPageInfByRange(ByVal cllFilter As Collection, _
    ByRef rsPatiPageInfo As ADODB.Recordset, _
    Optional ByVal lngModule As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:查询病人主页信息
    '入参:
    '   cllFilter 查询条件集:成员(Array(Key,Value),Array(Key,Value),,...)
    '       Key:病区IDS,科室IDS,病人IDS,主页IDS,入院开始时间,入院结束时间,出院开始时间,出院结束时间,
    '           费别,住院状态,病人性质,姓名,站点编号,查询转科病人,最后一次住院,险类,病区站点编号
    '       住院状态:0-在院病人;1-出院病人;2-在院或出院
    '       病人性质：多个用逗号分0-普通住院病人,1-门诊留观病人,2-住院留观病人，NULL-表示不区分
    '       姓名:可以代%分号表表按姓名匹配
    '       已出院天数，住院状态为1和2时有效
    '       站点编号:科室对应的站点编号
    '       险类:>0:指定险类医保病人,0:医保和普通病人,-1:普通病人,-2:医保病人
    '出参:
    '   rsPatiPageInfo 病人病案主页信息：病人ID,主页ID,姓名,性别,年龄,住院号,床号,险类,费别,病人类型,医保号,
    '                                   入院时间,出院时间,住院状态,病人性质,当前病区ID,当前病区名称,当前科室ID,当前科室名称,
    '                                   医疗付款方式名称,医疗付款方式编码,住院医师,病人备注,编目日期,护理等级,
    '                                   数据转出,审核标志,审核人,预出院时间,上次催款金额
    '       住院状态:病案主页.状态(0-正常住院；1-尚未入科；2-正在转科或正在转病区；3-已预出院)
    '       病人性质:0-普通住院病人,1-门诊留观病人,2-住院留观病人
    '       数据转出:0-未转出，1-已转出
    '       审核标志:0或空-未审核,1-已审核或开始审核;2-完成审核
    '返回:获取成功返回True，否则返回False
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objService As zlPublicExpense.clsService
    
    If zlGetServiceObject(objService) = False Then Exit Function
    GetPatiPageInfByRange = objService.ZlCissvr_GetPatiPageInfByRange(cllFilter, rsPatiPageInfo, lngModule)
End Function

