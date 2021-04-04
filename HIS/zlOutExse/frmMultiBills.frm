VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMultiBills 
   AutoRedraw      =   -1  'True
   Caption         =   "多单据退费"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11265
   Icon            =   "frmMultiBills.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   8430
   ScaleWidth      =   11265
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picInvoice 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   975
      Left            =   0
      ScaleHeight     =   975
      ScaleWidth      =   11265
      TabIndex        =   39
      Top             =   960
      Width           =   11265
      Begin VB.Frame fraSelectDownSplit 
         Height          =   30
         Left            =   -15
         TabIndex        =   41
         Top             =   900
         Width           =   11535
      End
      Begin VB.Frame fraSelectTopSplit 
         Height          =   45
         Left            =   -30
         TabIndex        =   40
         Top             =   0
         Width           =   11385
      End
      Begin VSFlex8Ctl.VSFlexGrid vsInvoice 
         Height          =   375
         Left            =   300
         TabIndex        =   42
         TabStop         =   0   'False
         Top             =   180
         Width           =   9960
         _cx             =   17568
         _cy             =   661
         Appearance      =   0
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
         ForeColorFixed  =   -2147483640
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483633
         GridColor       =   12632256
         GridColorFixed  =   -2147483641
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483633
         FocusRect       =   3
         HighLight       =   0
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   14
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   8
         FixedRows       =   0
         FixedCols       =   1
         RowHeightMin    =   360
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmMultiBills.frx":058A
         ScrollTrack     =   -1  'True
         ScrollBars      =   0
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
   End
   Begin VB.PictureBox pic退费摘要 
      Align           =   1  'Align Top
      BorderStyle     =   0  'None
      Height          =   480
      Left            =   0
      ScaleHeight     =   480
      ScaleWidth      =   11265
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   5565
      Width           =   11265
      Begin VB.TextBox txt退费摘要 
         Height          =   360
         Left            =   1035
         MaxLength       =   100
         TabIndex        =   6
         Top             =   45
         Width           =   5820
      End
      Begin VB.Label lbl摘要 
         AutoSize        =   -1  'True
         Caption         =   "退费摘要"
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
         Left            =   45
         TabIndex        =   5
         Top             =   120
         Width           =   960
      End
   End
   Begin VB.PictureBox picCmd 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   660
      Left            =   0
      ScaleHeight     =   660
      ScaleWidth      =   11265
      TabIndex        =   23
      TabStop         =   0   'False
      Top             =   7404
      Width           =   11265
      Begin VB.TextBox txtYB 
         Height          =   300
         Left            =   945
         TabIndex        =   34
         TabStop         =   0   'False
         Top             =   180
         Visible         =   0   'False
         Width           =   450
      End
      Begin VB.CommandButton cmdBillSel 
         Caption         =   "全选当前单据(&B)"
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
         Left            =   3240
         TabIndex        =   33
         ToolTipText     =   "热键：Ctrl+B"
         Top             =   135
         Width           =   2040
      End
      Begin VB.CommandButton cmdCancel 
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
         Left            =   9375
         TabIndex        =   17
         Top             =   135
         Width           =   1440
      End
      Begin VB.CommandButton cmdOK 
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
         Left            =   7845
         TabIndex        =   16
         Top             =   135
         Width           =   1440
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "全清(&R)"
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
         Left            =   1695
         TabIndex        =   25
         ToolTipText     =   "热键：Ctrl+R"
         Top             =   135
         Width           =   1440
      End
      Begin VB.CommandButton cmdSelAll 
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
         Left            =   165
         TabIndex        =   24
         ToolTipText     =   "热键：Ctrl+A"
         Top             =   135
         Width           =   1440
      End
      Begin VB.Line LineCmd_1 
         X1              =   0
         X2              =   12000
         Y1              =   0
         Y2              =   0
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   18
      Top             =   8064
      Width           =   11268
      _ExtentX        =   19870
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMultiBills.frx":0648
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12224
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Key             =   "误差"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
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
   Begin VSFlex8Ctl.VSFlexGrid vsBill 
      Align           =   1  'Align Top
      Height          =   3630
      Left            =   0
      TabIndex        =   4
      Top             =   1935
      Width           =   11265
      _cx             =   19870
      _cy             =   6403
      Appearance      =   1
      BorderStyle     =   1
      Enabled         =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   10.5
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
      BackColorSel    =   12632256
      ForeColorSel    =   0
      BackColorBkg    =   -2147483643
      BackColorAlternate=   -2147483643
      GridColor       =   -2147483636
      GridColorFixed  =   -2147483636
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   0
      FocusRect       =   1
      HighLight       =   2
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   3
      Cols            =   8
      FixedRows       =   1
      FixedCols       =   0
      RowHeightMin    =   280
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmMultiBills.frx":0EDC
      ScrollTrack     =   -1  'True
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
      AllowUserFreezing=   1
      BackColorFrozen =   0
      ForeColorFrozen =   0
      WallPaperAlignment=   9
      AccessibleName  =   ""
      AccessibleDescription=   ""
      AccessibleValue =   ""
      AccessibleRole  =   24
   End
   Begin VB.PictureBox picMoney 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   525
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   11265
      TabIndex        =   21
      TabStop         =   0   'False
      Top             =   6420
      Width           =   11265
      Begin VB.TextBox txt退款合计 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   405
         IMEMode         =   3  'DISABLE
         Left            =   5850
         Locked          =   -1  'True
         TabIndex        =   35
         TabStop         =   0   'False
         Top             =   75
         Width           =   1275
      End
      Begin VB.Frame fra退款 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   360
         Left            =   6705
         TabIndex        =   22
         Top             =   75
         Width           =   4515
         Begin VB.ComboBox cbo退款方式 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   1005
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   0
            Width           =   1620
         End
         Begin VB.TextBox txt退款金额 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00E0E0E0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   12
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   360
            IMEMode         =   3  'DISABLE
            Left            =   3195
            Locked          =   -1  'True
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   0
            Width           =   1275
         End
         Begin VB.Label lbl退款方式 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "本次退款"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Left            =   0
            TabIndex        =   12
            Top             =   75
            Width           =   960
         End
         Begin VB.Label lbl退款金额 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "金额"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000C0&
            Height          =   240
            Left            =   2685
            TabIndex        =   14
            Top             =   60
            Width           =   480
         End
      End
      Begin VB.TextBox txtAllTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   3465
         Locked          =   -1  'True
         TabIndex        =   11
         TabStop         =   0   'False
         Top             =   75
         Width           =   1275
      End
      Begin VB.TextBox txtCurTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00E0E0E0&
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   12
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         IMEMode         =   3  'DISABLE
         Left            =   1065
         Locked          =   -1  'True
         TabIndex        =   9
         TabStop         =   0   'False
         Top             =   75
         Width           =   1275
      End
      Begin VB.Label lbl退款合计 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "退款合计"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   240
         Left            =   4875
         TabIndex        =   36
         Top             =   135
         Width           =   960
      End
      Begin VB.Label lblAllTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "所属单据"
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
         Left            =   2460
         TabIndex        =   10
         Top             =   135
         Width           =   960
      End
      Begin VB.Label lblCurTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "当前单据"
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
         Left            =   75
         TabIndex        =   8
         Top             =   135
         Width           =   960
      End
   End
   Begin VB.PictureBox picPati 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   0
      ScaleHeight     =   960
      ScaleWidth      =   11265
      TabIndex        =   19
      TabStop         =   0   'False
      Top             =   0
      Width           =   11265
      Begin VB.PictureBox picPatiBack 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   540
         ScaleHeight     =   360
         ScaleWidth      =   2115
         TabIndex        =   37
         Top             =   525
         Width           =   2115
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
            ForeColor       =   &H00FF0000&
            Height          =   360
            Left            =   645
            MaxLength       =   100
            TabIndex        =   3
            ToolTipText     =   "定位:F6,输入:-病人ID,*门诊号,+住院号,.挂号单号,例如:*2536表示按门诊号查找"
            Top             =   0
            Width           =   1450
         End
         Begin zlIDKind.IDKindNew IDKind 
            Height          =   360
            Left            =   0
            TabIndex        =   38
            Top             =   0
            Width           =   660
            _ExtentX        =   1164
            _ExtentY        =   635
            Appearance      =   2
            IDKindStr       =   "姓|姓名|0;医|医保号|0;身|身份证号|0;IC|IC卡号|1;门|门诊号|0;就|就诊卡|0"
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
            MustSelectItems =   "姓名"
            BackColor       =   -2147483633
         End
      End
      Begin VB.TextBox txtPatientPrint 
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
         Left            =   9315
         MaxLength       =   64
         TabIndex        =   30
         ToolTipText     =   "热键:F11"
         Top             =   540
         Visible         =   0   'False
         Width           =   1800
      End
      Begin VB.OptionButton optNO 
         Caption         =   "票据号"
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
         Index           =   1
         Left            =   8280
         TabIndex        =   1
         Top             =   165
         Width           =   1035
      End
      Begin VB.OptionButton optNO 
         Caption         =   "单据号"
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
         Index           =   0
         Left            =   7245
         TabIndex        =   0
         Top             =   165
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.PictureBox pic退 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         Enabled         =   0   'False
         ForeColor       =   &H80000008&
         Height          =   390
         Left            =   10470
         ScaleHeight     =   360
         ScaleWidth      =   615
         TabIndex        =   28
         Top             =   45
         Visible         =   0   'False
         Width           =   645
         Begin VB.Label lbl退 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "退"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   18
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   360
            Left            =   90
            TabIndex        =   29
            Top             =   0
            Width           =   405
         End
      End
      Begin VB.Frame fraInfo_1 
         Height          =   120
         Left            =   -120
         TabIndex        =   27
         Top             =   390
         Width           =   12000
      End
      Begin VB.TextBox txtNO 
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
         IMEMode         =   3  'DISABLE
         Left            =   9315
         TabIndex        =   2
         Top             =   120
         Width           =   1800
      End
      Begin VB.Label lblPatiName 
         AutoSize        =   -1  'True
         Caption         =   "姓名"
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
         Height          =   240
         Left            =   8760
         TabIndex        =   31
         Top             =   600
         Visible         =   0   'False
         Width           =   480
      End
      Begin VB.Label lblTitle 
         AutoSize        =   -1  'True
         Caption         =   "病人收费单"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   18
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00800000&
         Height          =   360
         Left            =   75
         TabIndex        =   26
         ToolTipText     =   "清除:F6"
         Top             =   45
         Width           =   1875
      End
      Begin VB.Label lblPati 
         AutoSize        =   -1  'True
         Caption         =   "病人: "
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
         Height          =   240
         Left            =   45
         TabIndex        =   20
         Top             =   585
         Width           =   720
      End
   End
   Begin VSFlex8Ctl.VSFlexGrid vsBalance 
      Align           =   1  'Align Top
      Height          =   375
      Left            =   0
      TabIndex        =   7
      TabStop         =   0   'False
      Top             =   6045
      Width           =   11265
      _cx             =   19870
      _cy             =   661
      Appearance      =   0
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
      BackColorFixed  =   8421504
      ForeColorFixed  =   16777215
      BackColorSel    =   -2147483635
      ForeColorSel    =   -2147483634
      BackColorBkg    =   -2147483633
      BackColorAlternate=   -2147483633
      GridColor       =   12632256
      GridColorFixed  =   -2147483633
      TreeColor       =   -2147483632
      FloodColor      =   192
      SheetBorder     =   -2147483633
      FocusRect       =   3
      HighLight       =   0
      AllowSelection  =   0   'False
      AllowBigSelection=   0   'False
      AllowUserResizing=   1
      SelectionMode   =   1
      GridLines       =   1
      GridLinesFixed  =   1
      GridLineWidth   =   1
      Rows            =   2
      Cols            =   8
      FixedRows       =   0
      FixedCols       =   1
      RowHeightMin    =   360
      RowHeightMax    =   0
      ColWidthMin     =   0
      ColWidthMax     =   0
      ExtendLastCol   =   0   'False
      FormatString    =   $"frmMultiBills.frx":0F56
      ScrollTrack     =   -1  'True
      ScrollBars      =   0
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
End
Attribute VB_Name = "frmMultiBills"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mbytMode As Integer '0-多张单据查看,1-多张单据退费,2-退异常的退费单进行重新退费
Private mstrNo As String '要查看或退费的多张单据中的某张NO,退费时可以没有
Private mblnNOMoved As Boolean '操作的单据是否在后备数据表中
Private mstrPrivs As String
Private mlng领用ID As Long
Private mblnOneCard As Boolean
Private mblnSingleBlance As Boolean '单种结算方式
Private mrsALL As ADODB.Recordset  '所有单据的明细记录
Private mstrDelTime As String '查看退费单据的登记时间(yyyy-MM-dd HH:mm:ss) '只有查看退费单据时才传入时间,以区别正常单据
Private mstrNOs As String '实际读出可以退费的单据号
Private mstrNOsOverFlow As String '超出金额上限的单据号
Private mrsBalance As ADODB.Recordset '记录每张单据的结算情况
Private mcolError As Collection '记录每张单据的误差金额
Private mstrDelNOs As String '已经退完的单据或执行不能退的单据
Private mstr个人帐户 As String   '医保个人帐户的名称
Private mintInsure As Integer   '医保单据的险类
Private mblnYB结算作废 As Boolean '医保是否支持门诊结算作废
Private mint退费回单打印 As Integer '退费回单打印方式 0-不打印,1-自动打印,2-选择是否打印
Private mblnOK As Boolean
Private mblnPrintView As Boolean    '打印前查看调用
Private mintReturnMode As Integer   '用于退费时,全退禁用结算方式时恢复初始的结算方式
Private mrs收费对照 As ADODB.Recordset '收费对照 :问题:33634
Private Const mlngModule = 1121
Private mstrNOsPatiDel As String    '记录部分退费的单据
Private Type TYPE_MedicarePAR
    允许不设置医保项目 As Boolean
    门诊必须传递明细 As Boolean
    医保接口打印票据 As Boolean
    多单据一次结算 As Boolean
    分币处理 As Boolean
    实时监控 As Boolean
    退费后打印回单 As Boolean
    多单据调一次交易 As Boolean
    多单据收费必须全退 As Boolean
    
End Type
Private mstr现金结算方式 As String
Private MCPAR As TYPE_MedicarePAR
Private mobjSquare As Object
Private mlngShareUseID As Long '共享领用批次ID
Private mstrUseType As String '使用类别
Private mintInvoiceFormat As Integer  '打印的发票格式,发票格式序号
Private mintOldInvoiceFormat As Integer '旧发票格式
Private mintInvoicePrint As Integer '0-不打印;1-自动打印;2-提示打印
Private mblnNotClick As Boolean
'关于消费卡的处理变量
Private Type Ty_SquareCard
    blnExistsObjects As Boolean '安装了消费卡的
    rsSquare As ADODB.Recordset
    dbl刷卡总额 As Double
    bln卡结算 As Boolean '当前读取的单据是卡结算
End Type
Private mtySquareCard As Ty_SquareCard
Private mlng病人ID As Long
Private Type Ty_Pati
    病人ID As Long
    姓名 As String
    性别 As String
    年龄 As String
End Type
Private mtyPati As Ty_Pati
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1
Private mrsInfo As ADODB.Recordset
Private Enum EM_BillDelType
        EM_多张全退 = 0
        EM_单张全退 = 1
        EM_单张部分退 = 2
End Enum
Private mBillDelType As EM_BillDelType
Private mblnHaveExcuteData As Boolean '是否医嘱计价中存在数据:60735
Private Type tyBillType
    bln多单据 As Boolean
    bln单种结算方式 As Boolean
    bln单张部分退费 As Boolean  '存在单张部分退
    bln多张部分退费 As Boolean  '存在多张部分退
    bln存在卡结算 As Boolean '存在卡结算
    bln按结算序号退 As Boolean  '按结算序号退费
    strNos As String '本次收费单据
    str结算方式 As String '当前结算方式:多张时,用逗号分隔
    blnSingleBalance As Boolean
    bln存在医疗卡结算 As Boolean
    bln三方卡全退 As Boolean
End Type
Private mCurBillType As tyBillType  '当前单据类型
Private mrsDelInvoice As ADODB.Recordset
Private mobjDrugPacker  As Object ' 自动发药机(更新发药窗口)
Private mblnDrugPacker As Boolean
Private mblnFromInNewDel As Boolean ' 是否是从新退费窗口进来的

Public Function ShowMe(frmParent As Object, _
    ByVal bytMode As Byte, ByVal strPrivs As String, _
    ByVal strNo As String, ByVal strTime As String, _
    Optional blnPrintView As Boolean, _
    Optional lng领用ID As Long = 0, _
    Optional blnOneCard As Boolean = False, _
    Optional blnNOMoved As Boolean = False, _
    Optional blnFromInNewDel As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:多单据退费入口
    '入参:bytMode-0-多张单据查看,1-多张单据退费,2-退异常的退费单进行重新退费
    '       strPrivs-权限串
    '       blnNOMoved-是否转到后备数据表
    '       blnFromInNewDel-是否是从新退费窗口进来的
    '出参:
    '返回:退费成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-09-04 10:16:43
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mblnNOMoved = blnNOMoved: mstrPrivs = strPrivs
    mlng领用ID = lng领用ID: mblnOneCard = blnOneCard
    mbytMode = bytMode: mstrNo = strNo: mblnFromInNewDel = blnFromInNewDel
    mstrDelTime = strTime              '只有查看退费单据时才传入时间,以区别正常单据
    mblnPrintView = blnPrintView
    mblnOK = False
    On Error Resume Next
    If frmParent Is Nothing Then
        '医保调试调用
        Me.Show 0
    Else
        Me.Show 1, frmParent
    End If
    On Error GoTo 0
    ShowMe = mblnOK
End Function
Private Sub cbo退款方式_Click()
    If mblnNotClick Then Exit Sub
    Call ReCalcDelMoney
End Sub

Private Sub cmdBillSel_Click()
    Dim i As Long
    With vsBill
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("单据号")) <> "" And _
               .TextMatrix(i, .ColIndex("单据号")) = .TextMatrix(.Row, .ColIndex("单据号")) And InStr(1, mstrNOsOverFlow, vsBill.TextMatrix(i, .ColIndex("单据号"))) <= 0 Then
                .TextMatrix(i, .ColIndex("选择")) = -1
            End If
        Next
    End With
    Call LoadDelBalanceInfor
    Call ReCalcDelMoney
    Call FromNoSelectInvoice
    Call ShowAndHideDelBillRow
End Sub

Private Sub cmdCancel_Click()
    If mblnPrintView Then
        If txtPatientPrint.Visible Then
            If txtPatientPrint.Text = "" Then
                MsgBox "姓名为空,请输入姓名！", vbInformation, gstrSysName
                If txtPatientPrint.Enabled Then txtPatientPrint.SetFocus
                Exit Sub
            End If
            
            If zlCommFun.ActualLen(txtPatientPrint.Text) > txtPatientPrint.MaxLength Then
                MsgBox "病人姓名输入过长，只允许输入 " & txtPatientPrint.MaxLength & " 个字符或 " & txtPatientPrint.MaxLength \ 2 & " 个汉字。", vbInformation, gstrSysName
                If txtPatientPrint.Enabled Then txtPatientPrint.SetFocus
                Exit Sub
            End If
            
            If txtPatientPrint.Text <> txtPatientPrint.Tag Then
                
                Call ExecuteModifyPatiName
            End If
        End If
        mblnOK = True
    End If
    If mstrNOs <> "" And txtNO.Visible Then
        Call ClearFace
        txtNO.SetFocus
    Else
        Unload Me
    End If
End Sub

Private Function SetNOBill(ByVal strNo As String, ByVal blnSel As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:按单据全选或全清单据
    '入参:strNO-指定的NO
    '        blnSel:true表示全选,否则全清
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2011-01-24 10:47:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo errHandle
    With vsBill
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("单据号")) <> "" _
                And .TextMatrix(i, .ColIndex("单据号")) = strNo Then
                .TextMatrix(i, .ColIndex("选择")) = IIf(blnSel, -1, 0)
            End If
        Next
    End With
    SetNOBill = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Sub ExecuteModifyPatiName()
    Dim arrSQL As Variant, arrNo As Variant
    Dim i As Long, blnTrans As Boolean
    
    On Error GoTo errH
    arrNo = Split(mstrNOs, ",")
    arrSQL = Array()
    
    For i = 0 To UBound(arrNo)
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_病人费用记录_Update('" & arrNo(i) & "',1,null,null,'" & txtPatientPrint.Text & "')"
    Next
    
    gcnOracle.BeginTrans: blnTrans = True
        For i = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
        Next
    gcnOracle.CommitTrans: blnTrans = False
    
    Exit Sub
errH:
    If Err.Number <> 0 Then
        If blnTrans Then gcnOracle.RollbackTrans
        If ErrCenter() = 1 Then Resume
    End If
     
    If Err.Number <> 0 Then Call SaveErrLog
End Sub

Private Sub cmdClear_Click()
    Dim i As Long
    With vsBill
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("选择")) = 0
        Next
    End With
    Call LoadDelBalanceInfor
    Call ReCalcDelMoney
  '62492
    If vsInvoice.Visible Then
        If vsInvoice.Rows - 1 >= 1 And vsInvoice.COLS - 1 >= 1 Then
            vsInvoice.Cell(flexcpChecked, 0, 1, vsInvoice.Rows - 1, vsInvoice.COLS - 1) = 2
        End If
    End If
    Call ShowAndHideDelBillRow
End Sub
Private Function GetErrBillPartDelFee() As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:确定异常单据是否部分退费
    '出参:bln单张部分退
    '返回:返回退费类型(0-全退;1-单张部分退;2-多单据按单据部分退)
    '编制:刘兴洪
    '日期:2011-09-04 14:31:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim varData As Variant, i As Long, strNo As String
    On Error GoTo errHandle
    strSQL = "" & _
    "   Select /*+ rule */ A.NO,Max(decode(A.记录状态,2,1,0)) as 部分, " & _
    "           nvl(sum(A.实收金额),0)-nvl(sum(A.结帐金额),0) as 未结金额 " & _
    "   From 门诊费用记录 A,Table(f_str2List([1])) J" & _
    "   Where A.NO=J.Column_Value and A.记录性质=1  and A.执行状态<>9 " & _
    "   Having  nvl(sum(A.实收金额),0)-nvl(sum(A.结帐金额),0) <>0 " & _
    "   Group by A.NO"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrNOs)
    
    If rsTemp.RecordCount = 0 Then GetErrBillPartDelFee = 0: Exit Function
    '检查退费单据是否将所有单据退完
    With rsTemp
        varData = Split(mstrNo, ",")
        Do While Not .EOF
            strNo = Nvl(rsTemp!NO)
            If Val(Nvl(!部分)) = 1 Then GetErrBillPartDelFee = 1: Exit Function
            For i = 0 To UBound(varData)
                If varData(i) = strNo Then strNo = "HAVE": Exit For
            Next
            If strNo <> "HAVE" Then GetErrBillPartDelFee = 2
            .MoveNext
        Loop
    End With
    GetErrBillPartDelFee = 0
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function ErrBillReDelFee() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:对异常单据重新退费
    '返回:退费成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-09-04 10:42:39
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln退现 As Boolean, str退结算方式 As Boolean
    Dim blnYbComit As Boolean, blnCommited As Boolean, blnOneCardComit As Boolean
    Dim arrNo As Variant, i As Long, int医保 As Integer
    Dim lngPages As Long, lngPage, cllYB As Collection
    Dim rs原结帐 As ADODB.Recordset, str医保结算 As String
    Dim rs医保 As ADODB.Recordset, strInvoices As String
    Dim str冲销ID As String, strNo As String, lng结帐ID As Long
    Dim blnPrint As Boolean, strAllNOs As String
    Dim varTemp As Variant, blnTrans As Boolean
    Dim strReclaimInvoice  As String, intInvoiceFormat As Integer
    Dim strReturn As String, strReturnRecipt As String '退费处方信息，格式：NO,药房ID|NO,药房ID|…
    Dim rs药品记录 As ADODB.Recordset
    
    On Error GoTo errHandle
    Dim strSQL As String, blnAll部份退费 As Boolean, bln单张部分退 As Boolean
    '并发检查
    If zlIsCheckExistErrBill(0, False, mstrNOs) = False Then
        MsgBox "当前异常单据已被处理，你不能继续！", vbInformation, gstrSysName
        Exit Function
    End If
    If zlCheckOtherSessionDoing(0, mstrNOs) Then
        MsgBox "当前单据正在其它收费窗口中进行处理，你不能继续！", vbInformation, gstrSysName
        Exit Function
    End If
    '0-全退;1-单张部分退;2-多单据按单据部分退
    bln单张部分退 = False: blnAll部份退费 = False
    Select Case GetErrBillPartDelFee
    Case 1
        blnAll部份退费 = True: bln单张部分退 = True
    Case 2
        bln单张部分退 = False: blnAll部份退费 = True
    Case Else
    End Select
    
    If blnAll部份退费 Then
        If InStr(mstrPrivs, ";部份退费;") = 0 Then
            MsgBox "你没有权限执行部份退费操作！", vbInformation, gstrSysName
            vsBill.SetFocus: Exit Function
        End If
        '刘兴洪 问题:27352 日期:2010-01-13 10:26:08
        If InStr(1, mstrPrivs, ";退费核收发票;") > 0 Then
            If frmReInvoice.ShowMe(Me, mstrNo, Val(txtAllTotal.Text), Val(txt退款金额.Text), strInvoices) = False Then
                vsBill.SetFocus: Exit Function
            End If
        End If
    End If
      
    With mrsBalance
        .Filter = 0
        str冲销ID = ""
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If InStr(1, str冲销ID & ",", "," & Val(Nvl(!结帐ID)) & ",") = 0 Then
                    str冲销ID = str冲销ID & "," & Val(Nvl(!结帐ID))
            End If
            .MoveNext
        Loop
        If str冲销ID <> "" Then str冲销ID = Mid(str冲销ID, 2)
        If str冲销ID = "" Then
            MsgBox "未找到冲销数据,请检查!", vbOKOnly + vbInformation, gstrSysName
            Exit Function
        End If
    End With
    varTemp = Split(str冲销ID, ",")
    '问题:43347
    For i = 0 To UBound(varTemp)
        '一卡通检查
        If Not CheckOnCardValied(bln单张部分退, Val(varTemp(i))) Then Exit Function
        '三方交易检查
        If Not CheckThreeSwapValied(bln单张部分退, Val(varTemp(i)), InStr(1, mstrNOs, ",") > 0, True) Then Exit Function
    Next
    
    strSQL = "" & _
    "   Select /*+ rule */ distinct  A.结帐ID,A.No  " & _
    "   From  门诊费用记录 A,Table(f_str2List([1])) J" & _
    "   Where A.No=J.Column_Value and A.记录性质=1 And  A.记录状态 in (1,3)"
    Set rs原结帐 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstrNOs)
 
    
    strSQL = "Select /*+ Rule*/ Distinct a.No, a.结帐id, Decode(Nvl(b.险类, 0), 0, 0, 1) As 医保作废" & vbNewLine & _
            " From 门诊费用记录 A," & vbNewLine & _
            "      (Select Distinct j.Column_Value As 结帐id, m.险类" & vbNewLine & _
            "        From Table(f_Num2list([1])) J, 保险结算记录 M" & vbNewLine & _
            "        Where j.Column_Value = m.记录id(+) And m.性质(+) = 1) B" & vbNewLine & _
            " Where a.结帐id = b.结帐id"
    Set rs医保 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str冲销ID)
    arrNo = Split(mstrNOs, ",")
'    bln退现 = False
'    If cbo退款方式.ListIndex >= 0 Then
'        bln退现 = cbo退款方式.ItemData(cbo退款方式.ListIndex) = 1
'        str退结算方式 = zlStr.NeedName(cbo退款方式.Text)
'    End If

    
    If gblnBillPrint Then
        If gobjBillPrint.zlEraseBill(strAllNOs, 0) = False Then Exit Function
    End If
    blnTrans = False
    blnCommited = False: blnYbComit = False
    '--------------------------------------------------------------------------
    '医保退费:
    If mintInsure <> 0 Then
        '多单据退费是放在一个事务中的,肯定退费成功
        ' 不存在医保作废时,是原样退的,未调接口.因此也不可能出错
        If Not (MCPAR.多单据一次结算 Or MCPAR.多单据调一次交易 Or Not mblnYB结算作废) Then
            '-------------------------------------------------------------------------------------------------------
            '刘兴洪:医保的strAdvancey计算:本次退费总张数|当前退费第几张:27231
            Set cllYB = New Collection
            lngPage = 0: lngPages = 0
            For i = 0 To UBound(arrNo)
                strNo = arrNo(i)
                lngPage = UBound(arrNo) + 1 - i
                lngPages = lngPages + 1
                rs原结帐.Filter = "NO='" & arrNo(i) & "'"
                If rs原结帐.EOF Then
                    MsgBox "未找到单据号为" & arrNo(i) & "的原始结帐数据,请检查!", vbOKOnly + vbInformation, gstrSysName
                    Exit Function
                End If
                mrsBalance.Filter = "NO='" & arrNo(i) & "' And 性质=2"
                If mrsBalance.RecordCount <> 0 Then mrsBalance.MoveFirst
                With mrsBalance
                    str医保结算 = ""
                    Do While Not .EOF
                        str医保结算 = str医保结算 & "," & Nvl(!结算方式)
                        .MoveNext
                    Loop
                    If str医保结算 <> "" Then str医保结算 = Mid(str医保结算, 2)
                End With
                rs医保.Filter = "NO='" & arrNo(i) & "' and 医保作废=1"
                int医保 = IIf(rs医保.RecordCount <> 0, 1, 0)
                lng结帐ID = Val(Nvl(rs原结帐!结帐ID))
                cllYB.Add Array(lngPage, lng结帐ID, str医保结算, int医保), "_" & strNo
            Next
            '医保
            gcnOracle.BeginTrans: blnTrans = True
             For i = 0 To UBound(arrNo)
                strNo = arrNo(i)
                If Val(cllYB("_" & strNo)(3)) = 0 Then
                   'strAdace
                    lngPage = Val(cllYB("_" & strNo)(0)):  lng结帐ID = Val(cllYB("_" & strNo)(1))
                    If Not DelInsureOneBill(str医保结算, True, lng结帐ID, i + 1, lngPages, blnCommited) Then
                        If blnCommited = False Then gcnOracle.RollbackTrans: blnTrans = False
                        Exit Function
                    End If
                    If blnCommited = False Then gcnOracle.CommitTrans
                    gcnOracle.BeginTrans: blnTrans = True
                End If
            Next
        End If
    End If
    
    If blnTrans = False Then gcnOracle.BeginTrans: blnTrans = True
    '------------------------------------------------------------------------------------------
    '退一卡通
    blnCommited = False
    If Not DelOneCardPay(arrNo, blnCommited) Then
        If blnCommited = False Then gcnOracle.RollbackTrans: blnTrans = False
        Exit Function
    End If
    If blnCommited Then gcnOracle.BeginTrans: blnTrans = True
    '------------------------------------------------------------------------------------------
    '退一卡通等的三方交易
    blnCommited = False
    If Not DelThreeSwapFee(arrNo, blnCommited) Then
        If blnCommited = False Then gcnOracle.RollbackTrans: blnTrans = False
        Exit Function
    End If
    If blnCommited Then gcnOracle.BeginTrans: blnTrans = True
 
    '------------------------------------------------------------------------------------------
    '完成退费
    blnCommited = False
    If OverFeeDel(str冲销ID, mtyPati.病人ID, blnCommited) = False Then
        If blnCommited = False Then gcnOracle.RollbackTrans
        Exit Function
    End If
    If blnCommited = False Then gcnOracle.CommitTrans
    
    '81190,冉俊明,退费业务向发药机上传退费信息
    On Error Resume Next
    If mblnDrugPacker Then
        strSQL = "Select NO, 执行部门id" & _
            "   From 门诊费用记录" & _
            "   Where 结帐id In (Select Column_Value From Table(f_Str2list([1]))) And 收费类别 In ('5', '6', '7')"
        Set rs药品记录 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str冲销ID)
        
        If rs药品记录.RecordCount <> 0 Then
            Do While Not rs药品记录.EOF
                If InStr(strReturnRecipt & "|", "|" & Nvl(rs药品记录!NO) & "," & Nvl(rs药品记录!执行部门ID) & "|") = 0 Then
                    strReturnRecipt = strReturnRecipt & "|" & Nvl(rs药品记录!NO) & "," & Nvl(rs药品记录!执行部门ID)
                End If
                rs药品记录.MoveNext
            Loop
        End If

        If strReturnRecipt <> "" Then
            strReturnRecipt = Mid(strReturnRecipt, 2)
            Call mobjDrugPacker.DYEY_MZ_TransRecipeReturn(1, UserInfo.编号, UserInfo.姓名, strReturnRecipt, strReturn)
        End If
    End If
    Err.Clear: On Error GoTo errHandle
    
   '部分退费时收回并重打,包括单张部分退和退多张中的某几张
    If blnAll部份退费 Then
        'If InStr(1, mstrPrivs, "退费核收发票") > 0 Then strInvoices = frmReInvoice.ShowMe(Me, strNO, Val(txtAllTotal.Text), Val(txt退款金额.Text))
        If strInvoices = "" Then 'a.收回并重新打印门诊收据
            blnPrint = True
            If mintInvoicePrint = 0 Then
                blnPrint = False
            Else
                If mintInvoicePrint = 2 Then
                    If MsgBox("是否打印票据？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then
                        blnPrint = False
                    End If
                End If
            End If
            If blnPrint Then
                intInvoiceFormat = IIf(strReclaimInvoice = "" And gTy_Module_Para.byt票据分配规则 <> 0, mintOldInvoiceFormat, mintInvoiceFormat)
                Dim strPriceGrade As String
                If gintPriceGradeStartType >= 2 Then
                    strPriceGrade = GetPriceGradeFromNos(strAllNOs)
                Else
                    strPriceGrade = gstr普通价格等级
                End If
                Call RePrintCharge(1, strAllNOs, Me, mlng领用ID, strReclaimInvoice, True, CDate(mstrDelTime), _
                     intInvoiceFormat, , , mlngShareUseID, mstrUseType, , strPriceGrade)
            End If
        Else
            If strInvoices = "无可退票据" Then
                'b.收费或上一次退时没有打印票据
            Else
                'c.只收回票据
                strSQL = "zl_门诊收费记录_RePrint('" & strNo & "',Null,0,'" & UserInfo.姓名 & "'," & _
                        "To_Date('" & mstrDelTime & "','YYYY-MM-DD HH24:MI:SS'),1,0,'" & strInvoices & "')"
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            End If
        End If
        
        '打印费用清单
        If InStr(mstrPrivs, ";打印清单;") > 0 Then
            If gint收费清单 = 1 Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me, "NO=" & strAllNOs, "药品单位=" & IIf(gbln药房单位, 1, 0), 2)
            ElseIf gint收费清单 = 2 Then
                If MsgBox("要打印收费清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me, "NO=" & strAllNOs, "药品单位=" & IIf(gbln药房单位, 1, 0), 2)
                End If
            End If
        End If
    Else
         '税控部件全退时收回处理(全退时，zl_门诊收费记录_DELETE中已收回票据)
        If Not gobjTax Is Nothing And gblnTax Then
            gstrTax = gobjTax.zlTaxOutErase(gcnOracle, strAllNOs)
            If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
        End If
    End If
    If mintInsure <> 0 And MCPAR.退费后打印回单 And InStr(1, mstrPrivs, ";医保退费回单;") > 0 Then
        '问题:35248
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_4", Me, "NO=" & mstrNOs, 2)
    End If
    If mint退费回单打印 = 1 Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_5", Me, "NO=" & mstrNOs, 2)
    ElseIf mint退费回单打印 = 2 Then
        If MsgBox("是否打印退费回单？", vbYesNo + vbQuestion, gstrSysName) = vbYes Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_5", Me, "NO=" & mstrNOs, 2)
        End If
    End If
    ErrBillReDelFee = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub cmdOK_Click()
    If mbytMode = 2 Then
        '异常单据重新退费
        If ErrBillReDelFee = False Then Exit Sub
        mblnOK = True
        Unload Me: Exit Sub
    End If
    If ExecDelete Then
        mblnOK = True
        Call ClearFace(True, False)
        If txtNO.Visible Then
            txtNO.SetFocus
        Else
            Unload Me
            Exit Sub
        End If
    End If
 
End Sub

Private Sub cmdSelAll_Click()
    Dim i As Long
    With vsBill
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("单据号")) <> "" And InStr(1, mstrNOsOverFlow, vsBill.TextMatrix(i, .ColIndex("单据号"))) <= 0 Then
                .TextMatrix(i, .ColIndex("选择")) = -1
            End If
        Next
    End With
    If mbytMode <> 0 Then
        Call LoadDelBalanceInfor
    End If
    Call ReCalcDelMoney
    Call FromNoSelectInvoice
    Call ShowAndHideDelBillRow
End Sub

Private Sub Form_Activate()
    If txtNO.Visible And txtNO.Text = "" Then
        txtNO.SetFocus
    Else
        If vsBill.Enabled And vsBill.Visible Then vsBill.SetFocus
    End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF1 Then
        '###
    ElseIf KeyCode = vbKeyE And Shift = vbCtrlMask Then
        If cmdOK.Visible Then Call cmdOK_Click
    ElseIf KeyCode = vbKeyA And Shift = vbCtrlMask Then
        If cmdSelAll.Visible Then Call cmdSelAll_Click
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        If cmdClear.Visible Then Call cmdClear_Click
    ElseIf KeyCode = vbKeyEscape Then
        If mblnPrintView Then
            Unload Me
        Else
            If cmdCancel.Visible Then Call cmdCancel_Click
        End If
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("',|~:：;；?？" & Chr(22), Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub
Private Sub ClearVar()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除相关变量
    '编制:刘兴洪
    '日期:2012-09-17 13:23:35
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With mCurBillType
        .bln单张部分退费 = False
        .bln单种结算方式 = False
        .bln多单据 = False
        .bln多张部分退费 = False
        .strNos = ""
        .str结算方式 = ""
        .blnSingleBalance = False
        .bln存在医疗卡结算 = False
        .bln三方卡全退 = False
    End With
End Sub

Private Sub Form_Load()

    Call SetpicInvoiceVisible
    Call InitBillHead
    Call RestoreWinState(Me, App.ProductName)
    If Val(zlDatabase.GetPara("退费号码输入模式", glngSys, 1121, 0)) = 0 Then
        optNO(0).Value = True
    Else
        optNO(1).Value = True
    End If
    
    Call NewCardObject
    Call ClearFace
    
     lblTitle.Caption = gstrUnitName & "病人收费单"
    mint退费回单打印 = Val(zlDatabase.GetPara("退费回单打印方式", glngSys, mlngModule, "0"))
    If mbytMode = 0 Then '查看单据
        Caption = "查看多张单据"
        fra退款.Visible = False
        vsBill.ColHidden(0) = True
        cmdSelAll.Visible = False
        cmdClear.Visible = False
        cmdOK.Visible = False
        cmdBillSel.Visible = False
        If mblnPrintView Then cmdCancel.Caption = "确定(&X)"
        pic退.Visible = mstrDelTime <> ""
        lbl退款合计.Visible = False: txt退款合计.Visible = False
    ElseIf mbytMode = 2 Then
        '异常单据退费
        Caption = "异常退费单重新退费"
        Call initCardSquareData
        vsBill.ColHidden(0) = True
        cmdSelAll.Visible = False
        cmdClear.Visible = False
        cmdBillSel.Visible = False
        fra退款.Visible = False
        cmdOK.Visible = True
        pic退.Visible = mstrDelTime <> ""
        vsBill.Editable = flexEDNone
    Else
        Caption = "多张单据退费"
        Call initCardSquareData
    End If
    
    If mstrNo <> "" Then '指定了单据
        txtNO.Visible = False
        optNO(0).Visible = False
        optNO(1).Visible = False
        picPatiBack.Visible = False
        If Not ReadBills(mstrNo) Then Unload Me: Exit Sub
    End If

    '81190,冉俊明,退费业务向发药机上传退费信息
    mblnDrugPacker = False
    If mobjDrugPacker Is Nothing And (mbytMode = 1 Or mbytMode = 2) Then
        Err = 0: On Error Resume Next
        Set mobjDrugPacker = CreateObject("zlDrugPacker.clsDrugPacker")
        If Err <> 0 Then
            mblnDrugPacker = False
        Else
            mblnDrugPacker = mobjDrugPacker.DYEY_MZ_IniSoap
        End If
    End If
End Sub

Private Function Load结算方式() As Boolean
'说明:1-现金结算方式,2-其他非医保结算,3-医保个人帐户,4-医保各类统筹
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, i As Long
    
    cbo退款方式.Clear
    
    On Error GoTo errH
    Set rsTmp = Get结算方式("收费")
    For i = 1 To rsTmp.RecordCount
        If rsTmp!性质 = 3 Then
            mstr个人帐户 = rsTmp!名称
        ElseIf InStr(",1,2,7,", "," & rsTmp!性质 & ",") > 0 And Val(Nvl(rsTmp!应付款)) = 0 Then
            cbo退款方式.AddItem rsTmp!编码 & "-" & rsTmp!名称
            cbo退款方式.ItemData(cbo退款方式.NewIndex) = rsTmp!性质
            If rsTmp!性质 = 1 Then
                mstr现金结算方式 = Nvl(rsTmp!名称)
            End If
            If rsTmp!名称 = gstr结算方式 Then
                Call zlControl.CboSetIndex(cbo退款方式.hWnd, cbo退款方式.NewIndex)
            End If
            If rsTmp!缺省 = 1 And cbo退款方式.ListIndex = -1 Then
                Call zlControl.CboSetIndex(cbo退款方式.hWnd, cbo退款方式.NewIndex)
            End If
        End If
        
        rsTmp.MoveNext
    Next
    If mstr现金结算方式 = "" Then mstr现金结算方式 = "现金"
    If cbo退款方式.ListIndex = -1 And cbo退款方式.ListCount > 0 Then Call zlControl.CboSetIndex(cbo退款方式.hWnd, 0)
    Load结算方式 = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Form_Resize()
    Dim staH As Long

    On Error Resume Next
    
    staH = IIf(stbThis.Visible, stbThis.Height, 0)
    
    vsBill.Height = Me.ScaleHeight - picCmd.Height - staH - picPati.Height - picMoney.Height - pic退费摘要.Height - vsBalance.Height - IIf(picInvoice.Visible, picInvoice.Height, 0)
    
    If picMoney.ScaleWidth - fra退款.Width - 45 < txtAllTotal.Left + txtAllTotal.Width + 90 Then
        fra退款.Left = txtAllTotal.Left + txtAllTotal.Width + 90
    Else
        fra退款.Left = picMoney.ScaleWidth - fra退款.Width - 45
    End If
    
    If Me.ScaleWidth - cmdSelAll.Left - cmdCancel.Width > 5500 Then
        cmdCancel.Left = Me.ScaleWidth - cmdSelAll.Left - cmdCancel.Width
    Else
        cmdCancel.Left = 5500
    End If
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 90
    
    fraInfo_1.Width = Me.ScaleWidth + 300
    LineCmd_1.x2 = Me.ScaleWidth + 300
    
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbytMode = 0
    mstrNo = ""
    mstrDelTime = ""
    mstrNOsOverFlow = ""
    mblnNOMoved = False   '查看时,可能传入true
    Call initCardSquareData
    Call CloseIDCard
    zlDatabase.SetPara "退费号码输入模式", IIf(optNO(0).Value, "0", "1"), glngSys, 1121, InStr(1, mstrPrivs, ";参数设置;") > 0
    Call SaveWinState(Me, App.ProductName)
    
    If Not mobjDrugPacker Is Nothing Then
        '81190
        Set mobjDrugPacker = Nothing
    End If
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    '问题:50885
    Dim lng卡类别ID As Long, strOutCardNO As String, strExpand As String
    Dim strOutPatiInforXML As String
    If txtPatient.Locked Or txtPatient.Enabled = False Or txtPatient.Visible = False Then Exit Sub   '
    
    If objCard.名称 Like "IC卡*" And objCard.系统 Then
        If mobjICCard Is Nothing Then
            On Error Resume Next
            Set mobjICCard = CreateObject("zlICCard.clsICCard")
            If Err <> 0 Then
                Err = 0: On Error GoTo 0
                Exit Sub
            End If
        End If
        txtPatient.Text = mobjICCard.Read_Card()
        If txtPatient.Text <> "" Then Call FindPati(objCard, False, txtPatient.Text)
        Exit Sub
    End If
    lng卡类别ID = objCard.接口序号
    
    If lng卡类别ID = 0 Then Exit Sub
    '    zlReadCard(frmMain As Object, _
    '    ByVal lngModule As Long, _
    '    ByVal lngCardTypeID As Long, _
    '    ByVal blnOlnyCardNO As Boolean, _
    '    ByVal strExpand As String, _
    '    ByRef strOutCardNO As String, _
    '    ByRef strOutPatiInforXML As String) As Boolean
    '    '---------------------------------------------------------------------------------------------------------------------------------------------
    '    '功能:读卡接口
    '    '入参:frmMain-调用的父窗口
    '    '       lngModule-调用的模块号
    '    '       strExpand-扩展参数,暂无用
    '    '       blnOlnyCardNO-仅仅读取卡号
    '    '出参:strOutCardNO-返回的卡号
    '    '       strOutPatiInforXML-(病人信息返回.XML串)
    '    '返回:函数返回    True:调用成功,False:调用失败\
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModule, lng卡类别ID, True, strExpand, strOutCardNO, strOutPatiInforXML) = False Then Exit Sub
    txtPatient.Text = strOutCardNO
    If txtPatient.Text <> "" Then Call FindPati(objCard, False, txtPatient.Text)
 
End Sub

Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0
    Set gobjSquare.objCurCard = objCard
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    zlControl.TxtSelAll txtPatient
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    If txtPatient.Locked Then Exit Sub
    txtPatient.Text = objPatiInfor.卡号
    If txtPatient.Text <> "" Then Call FindPati(objCard, False, txtPatient.Text)
End Sub

Private Sub optNO_Click(Index As Integer)
    If Visible Then txtNO.SetFocus
    If Index = 0 Then
        If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZ", Mid(txtNO.Text, 2, 1)) > 0 Then
            txtNO.Text = ""
        End If
    End If
End Sub
Private Sub picInvoice_Resize()
    Err = 0: On Error Resume Next
    With fraSelectTopSplit
        .Top = picInvoice.ScaleTop
        .Left = picInvoice.ScaleLeft
        .Width = picInvoice.ScaleWidth
    End With
    With vsInvoice
        .Top = fraSelectTopSplit.Top + fraSelectTopSplit.Height + 50
        .Left = picInvoice.ScaleLeft + 50
        .Width = picInvoice.ScaleWidth - .Left * 2
    End With
    Call SetInvoceSizeAndShowTittle
    With fraSelectDownSplit
        .Top = vsInvoice.Top + vsInvoice.Height + 50
        .Left = picInvoice.ScaleLeft
        .Width = picInvoice.ScaleWidth
    End With
    picInvoice.Height = fraSelectDownSplit.Top + fraSelectDownSplit.Height + 50
End Sub

Private Sub picPati_Resize()
    txtNO.Left = picPati.ScaleWidth - txtNO.Width - 45
    optNO(1).Left = txtNO.Left - optNO(1).Width - 30
    optNO(0).Left = optNO(1).Left - optNO(0).Width - 15
    pic退.Left = picPati.ScaleWidth - pic退.Width - 45
    
    If txtPatientPrint.Visible Then
        txtPatientPrint.Left = picPati.Left + picPati.Width - txtPatientPrint.Width - 50
        lblPatiName.Left = txtPatientPrint.Left - lblPatiName.Width - 50
        txtPatientPrint.Top = txtNO.Top - 50
        lblPatiName.Top = txtNO.Top
    End If
End Sub

Private Sub pic退费摘要_Resize()
    Err = 0: On Error Resume Next
    With pic退费摘要
        txt退费摘要.Width = .ScaleWidth - txt退费摘要.Left - 50
    End With
End Sub

Private Sub txtAllTotal_GotFocus()
    zlControl.TxtSelAll txtAllTotal
End Sub

Private Sub txtCurTotal_GotFocus()
    zlControl.TxtSelAll txtCurTotal
End Sub

Private Sub txtNO_GotFocus()
    zlControl.TxtSelAll txtNO
End Sub

Private Sub txtNO_KeyPress(KeyAscii As Integer)
    Dim strAbc As String, str1 As String, str2 As String
    If KeyAscii = 13 Then
        KeyAscii = 0
        If txtNO.Text <> "" Then
            If optNO(0).Value Then
                txtNO.Text = GetFullNO(txtNO.Text, 13)
            End If
            Call zlControl.TxtSelAll(txtNO)
            If ReadBills(txtNO.Text) Then vsBill.SetFocus
        ElseIf txtPatient.Visible And txtPatient.Enabled Then
            txtPatient.SetFocus
        End If
    Else
        Call SetNOInputLimit(txtNO, KeyAscii, IIf(optNO(0).Value, 0, 1))
    End If
End Sub
Private Sub InitBillHead()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化退费的表头列信息
    '返回: 成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-09-11 09:47:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim arrHead As Variant, strHead As String, i As Long
    Dim varTemp As Variant, intCol As Integer
    
    strHead = "选择,300,4;单据号,1000,1;类别,720,1;项目,2800,1;商品名,2000,1;数量,750,7;单位,550,1;单价,1100,7;" & _
        "应收金额,1100,7;实收金额,1100,7;开单科室,1000,1;执行科室,1000,1;操作员,850,1;时间,1260,1;结帐ID;医嘱,1560,1;" & _
        "原始数量,0,4;准退数量,0,4;医嘱序号,0,4;执行科室ID,0,1"
    
    arrHead = Split(strHead, ";")
    With vsBill
        .Clear
        .FixedRows = 1: .FixedCols = 0
        .COLS = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        For i = 0 To UBound(arrHead)
            varTemp = Split(arrHead(i) & ",,,", ",")
            intCol = .FixedCols + i
            .ColKey(intCol) = varTemp(0)
            .TextMatrix(.FixedRows - 1, intCol) = varTemp(0)
            If UBound(varTemp) > 0 Then
                .ColHidden(intCol) = False
                .ColWidth(intCol) = Val(varTemp(1))
                If .ColWidth(intCol) = 0 Then .ColHidden(intCol) = True
                .ColAlignment(intCol) = Val(varTemp(2))
            Else
                .ColHidden(intCol) = True
            End If
        Next
         .TextMatrix(.FixedRows - 1, .ColIndex("选择")) = ""
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .COLS - 1) = 4
        .ColHidden(.ColIndex("商品名")) = gTy_System_Para.byt药品名称显示 <> 2
        .FrozenCols = 2
        .Editable = flexEDKbdMouse
        .ColDataType(0) = flexDTBoolean
    End With
End Sub

Private Sub txtPatient_KeyDown(KeyCode As Integer, Shift As Integer)
    If txtPatient.Locked Or txtPatient.Enabled = False Or txtPatient.Text <> "" Then Exit Sub
    If IDKind.ActiveFastKey = True Then Exit Sub
End Sub

Private Sub txtPatient_LostFocus()
    '问题:60010
    If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
    IDKind.SetAutoReadCard (False)
End Sub

Private Sub txtPatientPrint_Validate(Cancel As Boolean)
    txtPatientPrint.Text = Trim(txtPatientPrint.Text)
End Sub
Private Sub txt退费摘要_KeyPress(KeyAscii As Integer)
    zlControl.TxtCheckKeyPress txt退费摘要, KeyAscii, m文本式
End Sub

Private Sub vsBalance_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Row = 1 Then
        '问题:43403
        With vsBalance
            If Val(.TextMatrix(Row, Col)) <> 0 Then
                .Cell(flexcpForeColor, Row, Col - 1, Row, Col) = vbRed
                .Cell(flexcpFontBold, Row, .Col - 1, Row, .Col) = True
            Else
                .Cell(flexcpForeColor, Row, Col - 1, Row, Col) = Me.ForeColor
                .Cell(flexcpFontBold, Row, .Col - 1, Row, .Col) = False
            End If
        End With
    End If
    Call ReCalcDelMoney
End Sub

Private Sub vsBalance_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mbytMode = 2 Or mbytMode = 0 Then Cancel = True: Exit Sub
    With vsBalance
        '        '问题:43403
        If Col Mod 2 <> 0 Then Cancel = True: Exit Sub
        If Row <> 1 Then Cancel = True: Exit Sub
        If Val(.ColData(Col)) = 0 Or (mCurBillType.bln三方卡全退 And .RowHidden(1)) Then Cancel = True: Exit Sub
        '1.允许退现且支持部分退，可选择退回三方卡的金额
        '2.允许退现不支持部分退，单据全退时可选择退回三方卡的金额
        .ColComboList(Col) = " ||" & FormatEx(Val(.Cell(flexcpData, Row, Col)), 2)
    End With
End Sub

Private Sub vsBalance_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If vsBalance.MouseCol > 0 Then vsBalance.ToolTipText = vsBalance.ColData(vsBalance.MouseCol)  '显示结算摘要
End Sub
Private Sub zlSet诊疗固定关系(ByVal lngRow As Long, ByVal Col As Long, Optional lngNotCheckRow As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:住院费用记录
    '编制:刘兴洪
    '日期:2010-12-31 15:49:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, varTemp As Variant, bln固定 As Boolean, i As Long, j As Long
    
    If vsBill.Cell(flexcpData, lngRow, vsBill.ColIndex("结帐ID")) = "" Then Exit Sub
    If mrs收费对照 Is Nothing Then Exit Sub
     '问题:33634:如果是固定的项目(诊疗收费关系):即医嘱产生的才判断
    varData = Split(vsBill.Cell(flexcpData, lngRow, vsBill.ColIndex("结帐ID")) & ",", ",")
    If Val(varData(0)) = 0 Then Exit Sub
    
    mrs收费对照.Filter = "医嘱序号=" & Val(varData(0)) & " And 收费细目ID=" & Val(varData(1))
    If Not mrs收费对照.EOF Then
        bln固定 = Val(Nvl(mrs收费对照!固有对照)) = 1
    Else
        bln固定 = False
    End If
    mrs收费对照.Filter = 0
    If bln固定 = False Then Exit Sub
    With vsBill
        For i = 1 To .Rows - 1
            If i <> lngRow And lngNotCheckRow <> i Then
                varTemp = Split(vsBill.Cell(flexcpData, i, .ColIndex("结帐ID")) & ",", ",")
                If varData(0) = varTemp(0) Then    '是相同的医嘱序号
                     mrs收费对照.Filter = "医嘱序号=" & Val(varTemp(0)) & " And 收费细目ID=" & Val(varTemp(1))
                    If Not mrs收费对照.EOF Then
                        bln固定 = Val(Nvl(mrs收费对照!固有对照)) = 1
                    Else
                        bln固定 = False
                    End If
                    If bln固定 Then
                         .Cell(flexcpChecked, i, .ColIndex("选择")) = .Cell(flexcpChecked, lngRow, .ColIndex("选择"))
                         .TextMatrix(i, .ColIndex("选择")) = .TextMatrix(lngRow, .ColIndex("选择"))
                        '如果是主项,需要检查重项
                        If Val(.Cell(flexcpData, i, .ColIndex("项目"))) = 0 Then  '肯定为父项,因此,需要找从项内容
                            For j = i + 1 To vsBill.Rows - 1
                                 If .RowData(i) = Val(.Cell(flexcpData, j, .ColIndex("项目"))) Then
                                        .Cell(flexcpChecked, j, .ColIndex("选择")) = .Cell(flexcpChecked, i, .ColIndex("选择"))
                                         .TextMatrix(j, .ColIndex("选择")) = .TextMatrix(i, .ColIndex("选择"))
                                 End If
                            Next
                        End If
                    End If
                 End If
            End If
        Next
    End With
End Sub

Private Sub vsBill_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long, varData As Variant, bln固定 As Boolean
    Dim varTemp As Variant, j As Long
    
    With vsBill
        If Col = .ColIndex("选择") Then
            If mBillDelType = EM_多张全退 Then
                vsBill.TextMatrix(Row, .ColIndex("选择")) = 1
                    '根据单据选择发票
                    Call FromNoSelectInvoice
                Exit Sub
            ElseIf mBillDelType = EM_单张全退 Then
                    Call SetNOBill(vsBill.TextMatrix(Row, .ColIndex("单据号")), Val(vsBill.TextMatrix(Row, .ColIndex("选择"))) <> 0)
                    Call LoadBalanceInfor
                    Call LoadDelBalanceInfor
                    Call ReCalcDelMoney
                    '根据单据选择发票
                    Call FromNoSelectInvoice
            
                    Exit Sub
            End If
             stbThis.Panels(2).Text = ""
            '29201
            If Val(vsBill.Cell(flexcpData, Row, .ColIndex("项目"))) = 0 Then
                For i = Row + 1 To vsBill.Rows - 1
                     If Val(vsBill.RowData(Row)) = Val(vsBill.Cell(flexcpData, i, .ColIndex("项目"))) Then
                           vsBill.TextMatrix(i, .ColIndex("选择")) = vsBill.TextMatrix(Row, .ColIndex("选择"))
                     Else
                        Exit For
                     End If
                Next
                Call zlSet诊疗固定关系(Row, Col)
            Else
                Call zlSet诊疗固定关系(Row, Col)
                '需要检查主项是否已经被
                    For i = Row - 1 To 1 Step -1
                        If Val(vsBill.RowData(i)) = Val(vsBill.Cell(flexcpData, Row, .ColIndex("项目"))) Then
                            If vsBill.TextMatrix(i, .ColIndex("选择")) <> 0 Then
                                vsBill.TextMatrix(i, .ColIndex("选择")) = vsBill.TextMatrix(Row, .ColIndex("选择"))
                            End If
                            Call zlSet诊疗固定关系(i, Col, Row)
                             Exit For
                        End If
                    Next
            End If
            Call LoadBalanceInfor
            Call LoadDelBalanceInfor
            Call ReCalcDelMoney
            '根据单据选择发票
            Call FromNoSelectInvoice
        End If
    End With
End Sub

Private Sub vsBill_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim cur合计 As Currency, i As Long
        
    If NewRow <> OldRow Then
        With vsBill
            If .TextMatrix(NewRow, .ColIndex("单据号")) <> "" Then
                For i = NewRow - 1 To .FixedRows Step -1
                    If .TextMatrix(i, .ColIndex("单据号")) = .TextMatrix(NewRow, .ColIndex("单据号")) Then
                        cur合计 = cur合计 + Val(.TextMatrix(i, .ColIndex("实收金额")))
                    Else
                        Exit For
                    End If
                Next
                For i = NewRow To .Rows - 1
                    If .TextMatrix(i, .ColIndex("单据号")) = .TextMatrix(NewRow, .ColIndex("单据号")) Then
                        cur合计 = cur合计 + Val(.TextMatrix(i, .ColIndex("实收金额")))
                    Else
                        Exit For
                    End If
                Next
            End If
            txtCurTotal.Text = Format(cur合计, gstrDec)
        End With
    End If
End Sub

Private Sub vsBill_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
       With vsBill
            Select Case Col
            Case .ColIndex("选择")
                If mBillDelType = EM_多张全退 Then Cancel = True: Exit Sub
            End Select
        End With
End Sub

Private Sub vsBill_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vsBill.ColIndex("选择") Then Cancel = True
End Sub

Private Sub GetBillRow(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long)
    Dim i As Long
    
    lngBegin = lngRow: lngEnd = lngRow
    With vsBill
        For i = lngRow - 1 To .FixedRows Step -1
            If .TextMatrix(i, .ColIndex("单据号")) = .TextMatrix(lngRow, .ColIndex("单据号")) Then
                lngBegin = i
            Else
                Exit For
            End If
        Next
        For i = lngRow To .Rows - 1
            If .TextMatrix(i, .ColIndex("单据号")) = .TextMatrix(lngRow, .ColIndex("单据号")) Then
                lngEnd = i
            Else
                Exit For
            End If
        Next
    End With
End Sub

Private Sub vsBill_CellChanged(ByVal Row As Long, ByVal Col As Long)
    
    With vsBill
        If .ColIndex("单据号") < 0 Then Exit Sub
        If .TextMatrix(Row, .ColIndex("单据号")) <> "" And InStr(1, mstrNOsOverFlow, .TextMatrix(Row, .ColIndex("单据号"))) > 0 Then
             .TextMatrix(Row, .ColIndex("选择")) = 0
        End If
    End With
End Sub

Private Sub vsBill_DrawCell(ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'说明：1.OwnerDraw要设置为Over(画出单元所有内容)
'      2.Cell的GridLine从上下左右向内都是从第1根线开始
'      3.Cell的Border从左上是从第2根线开始,右下是从第1根线开始
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT
    
    With vsBill
        '擦除一并给药相关行列的边线及内容
        lngLeft = .ColIndex("单据号"): lngRight = .ColIndex("单据号")
        If Not Between(Col, lngLeft, lngRight) Then Exit Sub
        Call GetBillRow(Row, lngBegin, lngEnd)
        If lngBegin = lngEnd Then Exit Sub
        
        vRect.Left = Left '擦除左边表格线
        vRect.Right = Right - 1 '保留右边表格线
        If Row = lngBegin Then
            vRect.Top = Bottom - 1 '首行保留文字内容
            vRect.Bottom = Bottom
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 1 '底行保留下边线
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
        Else
            SetBkColor hDC, OS.SysColor2RGB(.BackColor)
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub
Private Sub vsBill_KeyPress(KeyAscii As Integer)
    With vsBill
        Select Case KeyAscii
        Case 32 '空格
            If .ColHidden(.ColIndex("选择")) Then Exit Sub
            KeyAscii = 0
            If .TextMatrix(.Row, .ColIndex("单据号")) <> "" Then
                If mBillDelType = EM_多张全退 Then Exit Sub
                If .TextMatrix(.Row, .ColIndex("选择")) = 0 _
                    And InStr(1, mstrNOsOverFlow, .TextMatrix(.Row, .ColIndex("单据号"))) <= 0 Then
                     .TextMatrix(.Row, .ColIndex("选择")) = -1
                Else
                     .TextMatrix(.Row, .ColIndex("选择")) = 0
                End If
                If mBillDelType = EM_单张全退 Then
                    Call SetNOBill(.TextMatrix(.Row, .ColIndex("单据号")), Val(.TextMatrix(.Row, .ColIndex("选择"))) <> 0)
                    Call LoadDelBalanceInfor
                    Call ReCalcDelMoney
                    Exit Sub
                End If
                 Call ReCalcDelMoney
            End If
            '87675,需要手动触发AfterEdit事件
            Call vsBill_AfterEdit(.Row, .ColIndex("选择"))
        Case 13 '回车
            KeyAscii = 0
            If .Row + 1 <= .Rows - 1 Then
                 .Row = .Row + 1: .ShowCell .Row, .Col
            End If
        End Select
 
    End With
End Sub

Private Sub vsBill_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsBill
        If Col <> .ColIndex("选择") Then Cancel = True
    End With
End Sub

Private Sub ClearFace(Optional ByVal blnNO As Boolean = True, Optional blnCllAllData As Boolean = True)
    Set mrsBalance = Nothing
    Set mcolError = Nothing
    Call ClearBalance
    vsBill.Rows = vsBill.FixedRows
    vsBill.Rows = vsBill.FixedRows + 1
    vsBill.Row = vsBill.FixedRows: vsBill.Col = vsBill.ColIndex("项目")
    If blnCllAllData = True Then
        vsBalance.COLS = 1
        vsBalance.TextMatrix(0, 0) = IIf(mstrDelTime = "", "收款结算", "退款结算")
    End If
    mstrNOs = ""
    mintInsure = 0: mstrDelNOs = ""
    mblnYB结算作废 = False  '不同的医保可能支持不一样,所以要清除
    txt退款金额.ToolTipText = ""    '记录的误差金额
    
    lblPati.Caption = "病人:"
    If blnNO Then txtNO.Text = ""
    If blnCllAllData Then
        txtCurTotal.Text = ""
        txtAllTotal.Text = ""
        txt退款金额.Text = "":
        stbThis.Panels(2).Text = ""
    End If
    If (mbytMode = 1 Or mbytMode = 2) And blnCllAllData Then
        Call Load结算方式
        If mbytMode = 1 Then
            cmdSelAll.Visible = True
            cmdClear.Visible = True
            cmdBillSel.Visible = True
        End If
    End If
End Sub
Private Sub initInsurePara(ByVal strNo As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化医保参数
    '入参:strNo-单据号
    '编制:刘兴洪
    '日期:2012-09-12 09:35:26
    '问题:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lng病人ID As Long, lng结帐ID As Long
    mintInsure = ChargeExistInsure(strNo, lng病人ID, lng结帐ID)
    If mintInsure = 0 Then Exit Sub
    MCPAR.多单据收费必须全退 = gclsInsure.GetCapability(support多单据收费必须全退, lng病人ID, mintInsure)
    mblnYB结算作废 = gclsInsure.GetCapability(support门诊结算作废, lng病人ID, mintInsure)
    MCPAR.多单据一次结算 = gclsInsure.GetCapability(support多单据一次结算, lng病人ID, mintInsure)
    MCPAR.医保接口打印票据 = gclsInsure.GetCapability(support医保接口打印票据, lng病人ID, mintInsure, CStr(lng结帐ID))
    MCPAR.退费后打印回单 = gclsInsure.GetCapability(support退费后打印回单, lng病人ID, mintInsure)
    MCPAR.多单据调一次交易 = gclsInsure.GetCapability(support门诊_不分单据结算, lng病人ID, mintInsure)
End Sub
Private Function CheckPrivsIsValied() As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查操作员是否具备操作退费单
    '返回:具备返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-09-12 09:46:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If Not (mbytMode = 1 Or mbytMode = 2) Then CheckPrivsIsValied = True: Exit Function
    
    '检查权限是否满足
    If mintInsure > 0 Then
        '保险退费权限检查
        If InStr(mstrPrivs, ";保险收费;") = 0 Then
            Screen.MousePointer = 0
            MsgBox "你没有权限对医保病人的单据退费！", vbInformation, gstrSysName
            Exit Function
        End If
        CheckPrivsIsValied = True: Exit Function
    End If
    '普通病人的处理
    '是否有非医保病人的退费权限
    If InStr(mstrPrivs, ";允许非医保病人;") = 0 Then
        Screen.MousePointer = 0
        MsgBox "你没有权限对非医保病人进行退费操作！", vbInformation, gstrSysName
        Exit Function
    End If
    CheckPrivsIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetBillDelType(ByVal strNos As String) As EM_BillDelType
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取本单据的退费类型
    '入参:strNos-本次操作的单据
    '返回:根据相关退费规则,返回本次退费的类型(多张全退;单张全退;单张部分退)
    '编制:刘兴洪
    '日期:2012-09-12 09:59:23
    '退费规则说明如下:
    '    普通病人:
    '        1.如果是多单据多种结算方式
    '            a)如果只有一种结算方式,则允许退其中的一笔,但要选择退的结算方式
    '            b)如果有多种结算方式,则只能退其中的一张单据,不能退一笔.
    '        2. 如果存在结算卡:
    '            a)如果是否退现=0的话,则根据是否全退来确定是否全部退或者单张退
    '            b)如果是否退现=1的话,则允许部分退,单笔退时,退成指定的结算方式
    '        3. 如果存在医疗卡:
    '           a.如果只有一种结算方式
    '            a)如果三方卡为"全退"且不支持退现(是否退现=0),且所有单据必须全退为原结算方式
    '            b)如果三方卡为"全退"且支持退现(是否退现=1),则允许选择单笔退,退成指定的结算方式
    '            c)如果三方卡为"部分退"且不支持退现(是否退现=0),则允许选择单笔退,但只能为原结算方式
    '            d)如果三方卡为"部分退"且支持退现(是否退现=1),则允许选择单笔退,退成指定的结算方式
    '           b.如果存在多种结算方式
    '            a)如果是否退现=0的话,则根据是否全退来确定是否单张全退或所有全退
    '            b)如果是否退现=1的话,则允许部分退,单笔退时,退成指定的结算方式
    '        4. 如果存在一卡通:应该按单张单据全退
    '   医保病人:
    '        a.support多单据收费必须全退:为true时,不能部分退
    '        b.其他的与普通病人的处理方式一致
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnDelSingleNO As Boolean
    On Error GoTo errHandle
    mblnSingleBlance = False
   '35461
    If mintInsure > 0 Then        '医保处理
        If InStr(strNos, ",") > 0 And MCPAR.多单据收费必须全退 Then
            mBillDelType = EM_多张全退: Exit Function
        End If
        blnDelSingleNO = True
    End If
    
    If mintInsure = 0 Then
        '普通病人的处理
        mBillDelType = EM_单张部分退
        '如果所有单据都只使用了一种结算,则允许部分退费(退其中的一张,或一张中的几笔,或几张中的几笔)
        mblnSingleBlance = CheckSingleBalance(Replace(strNos, "'", ""))
        '多种结算方式,必须单张全退
        'If Not mblnSingleBlance Then blnDelSingleNO = True
    End If
    
    '检查三方交易数据
    '       1.一卡通存在必须全退,直接返回
    '       2.三方交易,只要在医疗卡类别.是否全退为true且不能退现,则返回true
    '性质:1-预存款,2-医保,3-医疗卡(一卡通),4-结算卡,5-一卡通,0-其他类
    mrsBalance.Filter = "性质>2"
    With mrsBalance
        Do While Not .EOF
            '是否全退
            Select Case mrsBalance!性质
            Case 5:  '1.一卡通存在,必须全退,直接返回
                If InStr(strNos, ",") > 0 Then
                    GetBillDelType = EM_多张全退: Exit Function
                End If
                GetBillDelType = EM_单张全退: Exit Function
            Case 3:  '医疗卡
                If Val(Nvl(!是否全退)) = 1 And Val(Nvl(!是否退现)) = 0 Then '全退不退现
                    If InStr(strNos, ",") > 0 Then
                        GetBillDelType = EM_多张全退: Exit Function
                    End If
                    GetBillDelType = EM_单张全退: Exit Function
                End If
                If Not mCurBillType.blnSingleBalance Then
                    If Val(Nvl(!是否退现)) = 0 Then
                        '单张必须全退,不支持全退的,则必须单张全退
                        blnDelSingleNO = True
                    End If
                End If
            Case 4: '结算卡
                If Val(Nvl(!是否全退)) = 1 And Val(Nvl(!是否退现)) = 0 Then
                    If InStr(strNos, ",") > 0 Then
                        GetBillDelType = EM_多张全退: Exit Function
                    End If
                    GetBillDelType = EM_单张全退: Exit Function
                End If
                If Val(Nvl(!是否退现)) = 0 Then
                    '单张必须全退,不支持全退的,则必须单张全退
                    blnDelSingleNO = True
                End If
            End Select
            .MoveNext
        Loop
    End With
    GetBillDelType = IIf(blnDelSingleNO, EM_单张全退, EM_单张部分退)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function CheckDelIsValied(ByVal strNos As String, ByRef strNotCanDelNOs As String, ByRef strCanDelNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查退费单据是否合法
    '出参:strNotCanDelNOs-不能退的单据(已经执行及不能嫁的单据)
    '        strCanDelNos-能退的单号
    '返回:合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-09-12 15:12:40
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim arrNo As Variant, i As Long, intTmp As Integer
    Dim strInfo As String, strFlagPrintInfor As String
    Dim blnFlagPrint As Boolean, strNo As String, strCurNO As String
    Dim strOper As String, vDate As Date, blnHaveExe As Boolean
    
    On Error GoTo errHandle
    
    '问题:54728
    If Not mbytMode = 1 Then CheckDelIsValied = True: Exit Function   '退费时判断
    
    arrNo = Split(strNos, ",")
    strNotCanDelNOs = ""
     '是否已执行
    strCanDelNos = ""   '记录可以退的单据号
    strInfo = ""        '检查结果提示信息
    strFlagPrintInfor = ""
    For i = 0 To UBound(arrNo)
        strCurNO = Replace(arrNo(i), "'", "")
        If strNo = "" Then strNo = strCurNO
        If i = 0 Then
            If Not ReadBillInfo(1, strCurNO, 1, strOper, vDate) Then
                strInfo = "单据[" & strCurNO & "]不存在!"
                Exit For
            End If
            If InStr(mstrPrivs, "所有操作员") <= 0 And UserInfo.姓名 <> strOper Then
                strInfo = "你没有""所有操作员""权限,不能对" & strOper & "的单据进行退费!"
                Exit For
            End If
            If Not BillOperCheck(2, strOper, vDate, "退费", strCurNO, , 1) Then
                Screen.MousePointer = 0:  Exit Function
            End If
        End If
        
        blnHaveExe = False: blnFlagPrint = False
        intTmp = BillCanDelete(strCurNO, 1, blnHaveExe, , blnFlagPrint)
        If intTmp <> 0 Then
            strNotCanDelNOs = strNotCanDelNOs & "," & strCurNO  '不能退的单据
            If Not mCurBillType.bln多张部分退费 Then mCurBillType.bln多张部分退费 = True
            Select Case intTmp
                Case 1 '该单据不存在
                    strInfo = strInfo & "指定的单据不存在！" & vbCrLf
                    Exit For
                Case 2 '已经全部完全执行(收费不考虑退费自动退药)
                    strInfo = strInfo & "[" & strCurNO & "]中的项目已经全部完全执行,不能退费!" & vbCrLf
                Case 3 '未完全执行部分剩余数量为0
                    strInfo = strInfo & "[" & strCurNO & "]中未完全执行的项目剩余数量为零,没有可退费用！" & vbCrLf
            End Select
            
        ElseIf blnHaveExe Then
            '存在已执行项目
            If mintInsure > 0 Then '收费医保退费
                strInfo = strInfo & "[" & strCurNO & "]中属于医保病人的收费单,存在已经执行的项目,不能退费！" & vbCrLf
                strNotCanDelNOs = strNotCanDelNOs & "," & strCurNO  '不能退的单据
            ElseIf mBillDelType <> EM_单张部分退 Then
                strInfo = strInfo & "[" & strCurNO & "]中存在已执行的项目,不能退费。" & vbCrLf
                strNotCanDelNOs = strNotCanDelNOs & "," & strCurNO  '不能退的单据
            Else
                strInfo = strInfo & "[" & strCurNO & "]中存在已执行的项目，此单据将执行的是部分退费。" & vbCrLf
                strCanDelNos = strCanDelNos & "," & strCurNO
            End If
            If Not mCurBillType.bln多张部分退费 Then mCurBillType.bln多张部分退费 = True
        Else
            strCanDelNos = strCanDelNos & "," & strCurNO
        End If
        
        If blnFlagPrint Then
            '检查对应的条码是否已打印(检验医嘱中的采集方式已执行)
            strFlagPrintInfor = strFlagPrintInfor & "[" & strCurNO & "]检验医嘱的条码已打印。" & vbCrLf
        End If
    Next
    
    If strNotCanDelNOs <> "" Then strNotCanDelNOs = Mid(strNotCanDelNOs, 2)
    strCanDelNos = Mid(strCanDelNos, 2)
    
    If strFlagPrintInfor <> "" Then
        If MsgBox("注意:" & vbCrLf & strFlagPrintInfor & vbCrLf & " 是否继续退费？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Screen.MousePointer = 0
            Exit Function
        End If
    End If
    
    If strNotCanDelNOs <> "" And mBillDelType = EM_多张全退 Then
        Screen.MousePointer = 0
        MsgBox "本次收费必须全退:" & vbCrLf & strInfo, vbInformation, gstrSysName
        Exit Function
    End If
    
    If strCanDelNos = "" Then
        '多张单据因为登记日期一样,必然是一起转出或都没有转出
        '是否已转入后备数据表中
        If zlDatabase.NOMoved("门诊费用记录", strNo, , "1") Then
            If Not ReturnMovedExes(strNo, 1, Me.Caption) Then
                Screen.MousePointer = 0
                Exit Function
            End If
        End If
        Screen.MousePointer = 0
        MsgBox strInfo, vbInformation, gstrSysName
        Exit Function
    End If

    If strInfo <> "" Then
        MsgBox strInfo, vbInformation, gstrSysName
    End If
    strNos = strCanDelNos
    
    '多张单据因为登记日期一样,必然是一起转出或都没有转出
    '是否已转入后备数据表中
    If zlDatabase.NOMoved("门诊费用记录", strNo, , "1") Then
        If Not ReturnMovedExes(strNo, 1, Me.Caption) Then
            Screen.MousePointer = 0
            Exit Function
        End If
    End If
    CheckDelIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub InitBillVar(ByVal strNos As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化单据
    '编制:刘兴洪
    '日期:2012-09-17 13:29:53
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str结算方式 As String
    With mCurBillType
        .bln多单据 = InStr(strNos, ",") > 0
        .strNos = strNos
        .bln存在卡结算 = False
        .bln存在医疗卡结算 = False
        .bln三方卡全退 = False
    End With
    '检查三方交易数据
    '       1.一卡通存在必须全退,直接返回
    '       2.三方交易,只要在医疗卡类别.是否全退为true且不能退现,则返回true
    '性质:1-预存款,2-医保,3-医疗卡(一卡通),4-结算卡,5-一卡通,0-其他类
    mrsBalance.Filter = "性质<>2 And 性质<>1"
    str结算方式 = ""
    mrsBalance.Sort = "NO,性质"
    'W.NO,A.结帐ID
    With mrsBalance
        Do While Not .EOF
            If InStr(str结算方式 & ",", "," & Nvl(!结算方式) & ",") = 0 Then
                str结算方式 = str结算方式 & "," & Nvl(!结算方式)
            End If
            If Val(Nvl(!性质)) = 4 Or Val(Nvl(!性质)) = 3 Then mCurBillType.bln存在卡结算 = True
            If Val(Nvl(!性质)) = 3 Then mCurBillType.bln存在医疗卡结算 = True
             
            .MoveNext
        Loop
    End With
    If str结算方式 <> "" Then str结算方式 = Mid(str结算方式, 2)
    mCurBillType.bln单种结算方式 = InStr(str结算方式, ",") = 0
    mCurBillType.str结算方式 = str结算方式
    
    str结算方式 = ""
    mrsBalance.Filter = 0
    With mrsBalance
        Do While Not .EOF
            If InStr(str结算方式 & ",", "," & Nvl(!结算方式) & ",") = 0 Then
                str结算方式 = str结算方式 & "," & Nvl(!结算方式)
            End If
            .MoveNext
        Loop
    End With
    If str结算方式 <> "" Then str结算方式 = Mid(str结算方式, 2)
    mCurBillType.blnSingleBalance = InStr(str结算方式, ",") = 0
End Sub



Private Function ReadBills(ByVal strNo As String) As Boolean
'功能：根据当前输入的单据号或票据号,读取并显示多张单据
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strSQL1 As String, strSQL2 As String, strSQL3 As String
    Dim strNos As String, strBalance As String, str结算 As String
    Dim strSub As String, strTmp As String, strCurNO As String, strCanDelNos As String
    Dim blnDelSingleNO As Boolean
    Dim i As Long, intTmp As Integer, intSign As Integer, j As Integer
    Dim arrNOs As Variant, cur合计 As Currency, arrNo As Variant
    Dim strTemp As String, str医嘱序号 As String
    Dim blnNotFind As Boolean
    Dim lng病人ID As Long, cllInvoiceNoInfor As Collection
    Dim str结算序号 As String
    Dim strInvoiceNO As String, strOldNO As String
    
    On Error GoTo errH
    Set mrsDelInvoice = Nothing
    mblnSingleBlance = False
    Call ClearFace(False)
    strOldNO = strNo
    
    Call ClearVar   '清除当前单据的相关变量
    Screen.MousePointer = 11
    Set cllInvoiceNoInfor = New Collection
    '确定一起收费的多张单据
    '如果是多张退费状态mbytMode = 1,是否在在线表不能确定，需要联接查询
    '----------------------------------------------------------------------------------
    '56963
    strInvoiceNO = ""
    If Not (mstrNo <> "" Or optNO(0).Value) Then
         '按票据号:可能不同批次票号重复
        strInvoiceNO = strNo
        strNos = zlInvoiceFromNOs(strInvoiceNO, True, str结算序号, cllInvoiceNoInfor)
        If InStr(str结算序号, ",") > 0 Then
            '证明有多个结算序号,需要操作员确定哪次费用
            strNo = ""
            Screen.MousePointer = 0
            If frmMulitChargeSelect.zlShowSelect(Me, mlngModule, strNos, strInvoiceNO, strNo) = False Then
                Screen.MousePointer = 0
                Exit Function
            End If
            If strNo = "" Then
                Screen.MousePointer = 0
                Exit Function
            End If
            Screen.MousePointer = 99
        Else
            strNo = Replace(Split(strNos & ",", ",")(0), "'", "")
        End If
        strOldNO = ""
        strNos = GetMultiNOs(strNo, , , True, True)
    End If
    
    Dim strTempNos As String, intInsure As Integer
    strTempNos = GetMultiNOs(strNo, , , True, True)
    strNos = GetMultiNOs(strNo, , , False, True)
    mCurBillType.bln按结算序号退 = False
    If mbytMode = 0 Then
        strNos = strTempNos
    Else
        If InStr(strTempNos, ",") > 0 And InStr(strNos, ",") = 0 Then
            '肯定是按单据分别打印的
            '要多单据退费,就必须满足以下条件
            '1.医保多单据必须全退时,必须按结算序号进行退费
            '2.三方账户全退时,必须按结算序号进行退费
            intInsure = ChargeExistInsure(strNo)
            If intInsure <> 0 Then
                If gclsInsure.GetCapability(support多单据收费必须全退, , intInsure) Then
                    strNos = strTempNos: mCurBillType.bln按结算序号退 = True
                End If
            ElseIf zlIsExistsSquareCard(strTempNos, True) Then
                '检查一卡通结算部分是否存在全退的
                strNos = strTempNos: mCurBillType.bln按结算序号退 = True
            Else
                If mbytMode = 1 And mstrNo <> "" And Not mblnFromInNewDel Then
                    If frmMulitChargeSelect.zlShowSelect(Me, mlngModule, strTempNos, "", strNos, True) = False Then
                        Screen.MousePointer = 0
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    
    If strNos = "" Then
        If optNO(1).Value Then
            Screen.MousePointer = 0
            MsgBox "没有找到与号码""" & strNo & """相关的收费记录。", vbInformation, gstrSysName
            Exit Function
        End If
        '可能因为未用票据而读不出来
        strNos = strNo
    End If
    '需要加引号
    If InStr(1, strNos, "'") = 0 Then
        strNos = "'" & Replace(strNos, ",", "','") & "'"
    End If
    '冉俊明:选择的单据进行了医保补充结算，则不允许退费
    If mbytMode <> 0 Then
        If CheckBillExistReplenishData(1, , Replace(strNos, "'", "")) = True Then
            Screen.MousePointer = 0
            MsgBox "当前单据进行了医保补充结算，不允许进行退费操作！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    arrNo = Split(strNos, ",")
    
    If gbln退费申请模式 And mbytMode = 1 Then
        Set rsTmp = GetApply(strNo, 1)
        rsTmp.Filter = "状态<>2"
        If rsTmp.RecordCount = 0 Then
            Screen.MousePointer = 0
            MsgBox "请先对该单据进行退费申请！", vbInformation, gstrSysName
            Exit Function
        End If
        If IsNull(rsTmp!审核人) Then
            Screen.MousePointer = 0
            MsgBox "该单据未进行退费审核，不能进行退费！", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    '性质:1-预存款,2-医保,3-医疗卡,4-结算卡,5-一卡通,0-其他类
    intSign = IIf(mstrDelTime <> "", -1, 1) '数量,金额正负符号
    Set mrsBalance = GetChargeBalance(strNos, , , , mstrDelTime, intSign)
    Call initInsurePara(strNo)
    If CheckPrivsIsValied = False Then Exit Function    '操作权限检查
    
    Call InitBillVar(strNos)    '初始化当前退费的单据信息变量
    
    '确定退费类型
    mBillDelType = GetBillDelType(strNos)
    
    '退费相关检查
    If CheckDelIsValied(strNos, mstrDelNOs, strCanDelNos) = False Then
        Screen.MousePointer = 0:  Exit Function
        Exit Function
    End If
    If strCanDelNos <> "" Then strNos = strCanDelNos
        
    '读取病人信息
    '----------------------------------------------------------------------------------
    strSQL = "" & _
    " Select A.病人ID,A.姓名,A.性别,A.年龄,A.标识号,A.费别,C.名称 as 付款方式,B.险类,E.病人类型" & _
    " From " & IIf(mblnNOMoved, "H", "") & "门诊费用记录 A,病人信息 E,保险结算记录 B,医疗付款方式 C,人员表 D" & _
    " Where A.病人ID=E.病人ID(+) And A.付款方式=C.编码(+) And A.结帐ID=B.记录ID(+) And B.性质(+)=1 And A.操作员姓名=D.姓名" & _
    "       And (D.站点='" & gstrNodeNo & "' Or D.站点 is Null)" & vbNewLine & _
    "       And A.记录性质=1 And A.记录状态 IN(1,3) And A.NO=[1] And Rownum < 2"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo)
    If rsTmp.EOF Then
        Screen.MousePointer = 0
        MsgBox "没有找到与号码""" & strNo & """相关的收费记录。", vbInformation, gstrSysName
        mlng病人ID = 0
        Exit Function
    End If
    txtPatient.Text = Nvl(rsTmp!姓名)
    
    lblPati.Caption = "病人:" & IIf(txtPatient.Visible, "       ", rsTmp!姓名) & _
        "　性别:" & Nvl(rsTmp!性别) & _
        "　年龄:" & Nvl(rsTmp!年龄) & _
        "　门诊号:" & Nvl(rsTmp!标识号) & _
        "　费别:" & Nvl(rsTmp!费别) & _
        "　付款方式:" & rsTmp!付款方式
        
    mlng病人ID = Val(Nvl(rsTmp!病人ID))
    With mtyPati
        .病人ID = mlng病人ID
        .性别 = Nvl(rsTmp!性别)
        .年龄 = Nvl(rsTmp!年龄)
        .姓名 = Nvl(rsTmp!姓名)
    End With
    
    If Not IsNull(rsTmp!险类) Then
        lblPati.ForeColor = vbRed
        txtYB.Text = Val(Nvl(rsTmp!险类))   '问题:41760
        txtPatient.ForeColor = vbRed
    Else
        lblPati.ForeColor = &HC00000
        txtYB.Text = ""
        txtPatient.ForeColor = &HC00000
    End If
    '75259：李南春,2014-7-10，病人姓名显示颜色处理
    Call SetPatiColor(txtPatient, Nvl(rsTmp!病人类型), txtPatient.ForeColor)
    lblPati.ForeColor = txtPatient.ForeColor
     If mblnPrintView And InStr(1, mstrPrivs, "修改姓名重打") > 0 And IsNull(rsTmp!病人ID) Then
        txtPatientPrint.Text = "" & rsTmp!姓名
        txtPatientPrint.Tag = txtPatientPrint.Text
        txtPatientPrint.Visible = True
        lblPatiName.Visible = True
    End If
    
    '读取结算内容:原始或退费的,结算方式为空指冲预交的记录
    '----------------------------------------------------------------------------------
    Call LoadBalanceInfor
      mintReturnMode = cbo退款方式.ListIndex  '用于退费时,全退禁用结算方式时恢复初始的结算方式
    '读取单据内容
    '----------------------------------------------------------------------------------
    
    If mbytMode = 1 Then
        '0-多张单据查看,1-多张单据退费
        '退费时不用考虑后备表,前面的操作已禁用
        '读取准退数,并计算应收金额,实收金额(金额=剩余金额*(准退数/剩余数))
        '读取单据中原始记录的费用ID
        Dim strTableNo As String
        mblnHaveExcuteData = zlCheckIsExcuteData(Replace(strNos, "'", ""), 1)     '60735
        
        '刘兴宏45685,58077
        strTableNo = "" & _
        "   With 门诊单据 as (Select  Column_Value as No From Table(f_Str2list([2])))," & _
        "            门诊费用  as (" & _
        "           Select A.ID,A.记录性质,A.NO,A.记录状态,A.序号,A.从属父号,A.价格父号,A.收费细目ID, " & _
        "                      nvl(A.付数,1) as 付数, nvl(A.数次,0) as 数次, " & _
        "                      nvl(A.应收金额,0) as 应收金额 ,nvl(A.实收金额,0) as 实收金额,nvl(A.结帐金额,0) as 结帐金额," & _
        "                      Nvl(A.付数,1)*A.数次 as 数量, nvl(标准单价,0)  as 标准单价," & _
                               IIf(gbln药房单位, "nvl(B." & gstr药房包装 & ",1)", "1") & " as 换算系数, " & _
                               IIf(gbln药房单位, " decode(B.药品ID,NULL,A.计算单位,B." & gstr药房单位 & ")", "A.计算单位 ") & " as 计算单位," & _
        "                      A.开单部门ID,A.执行部门ID,A.医嘱序号, " & _
        "                      A.执行状态,A.费用类型,A.费用状态 ,A.附加标志,A.费别,A.收费类别,A.操作员姓名,A.登记时间,A.结帐ID," & _
        "                      B.药品ID" & _
        "           From 门诊费用记录 A,药品规格 B,门诊单据 J  " & _
        "           Where A.记录性质=1 And A.NO=J.NO and A.记录状态<>0" & _
        "                       And A.收费细目ID=B.药品ID(+)" & _
        "              )," & _
        ""
        '求准退费(卫材,药品,其他治疗类)
        strTableNo = strTableNo & vbCrLf & _
        "            准退数  as ( " & _
        "            Select  A.费用ID,Sum(Nvl(A.付数,1)*A.实际数量" & IIf(gbln药房单位, "/Nvl(B." & gstr药房包装 & ",1)", "") & ") as 准退数量" & _
        "            From 药品收发记录 A,药品规格 B, 门诊单据 J" & _
        "           Where A.药品ID=B.药品ID(+) And Mod(A.记录状态,3)=1  " & _
        "                       And (A.单据 =8 or a.单据=24) And A.审核人 is NULL And A.NO =J.NO" & _
        "           Group by A.费用ID"
        
        '求诊疗相关的准退数
        If mblnHaveExcuteData Then
            '60735:在医嘱执行计价中存在数据时,则按医嘱执行计价中取数
            strTableNo = strTableNo & " Union ALL  " & _
            " Select Max(ID) As 费用id, Decode(Sign(Sum(数量)), -1, 0, Sum(数量)) As 准退数" & vbNewLine & _
            " From ( Select Decode(a.记录状态, 2, 0, a.Id) As ID, a.医嘱序号 As 医嘱id, a.收费细目id, Nvl(a.付数, 1) * Nvl(a.数次, 1) As 数量," & vbNewLine & _
            "              Decode(a.记录状态, 2, 0, Nvl(a.付数, 1) * Nvl(a.数次, 1)) As 原始数量" & vbNewLine & _
            "       From 门诊费用 A, 病人医嘱记录 M" & vbNewLine & _
            "       Where a.医嘱序号 = m.Id And Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0 And Instr('5,6,7', a.收费类别) = 0 And" & vbNewLine & _
            "            　a.记录状态 In (1, 2, 3)　and  a.价格父号 is null " & vbNewLine & _
            "          And Not Exists" & _
            "                (Select 1 From 病人医嘱附费 Where a.医嘱序号 = 医嘱id and a.No = NO and Mod(a.记录性质, 10) = 记录性质)" & _
            "       Union All" & vbNewLine & _
            "       Select a.Id, a.医嘱序号 As 医嘱id, a.收费细目id, -1 * b.数量 As 已执行, 0 原始数量" & vbNewLine & _
            "       From 门诊费用 A, 医嘱执行计价 B, 病人医嘱记录 M" & vbNewLine & _
            "       Where a.医嘱序号 = b.医嘱id And a.收费细目id = b.收费细目id + 0 And a.医嘱序号 = m.Id And Instr(',C,D,F,G,K,', ',' || m.诊疗类别 || ',') = 0" & vbNewLine & _
            "           And Instr('5,6,7', a.收费类别) = 0" & vbNewLine & _
            "           And (Exists (Select 1  From 病人医嘱执行  Where b.医嘱id = 医嘱id And b.发送号 = 发送号 And b.要求时间 = 要求时间 And Nvl(执行结果, 0) = 1)" & vbNewLine & _
            "                Or Exists (Select 1 From 病人医嘱发送 Where b.医嘱id = 医嘱id And b.发送号 = 发送号 And Nvl(执行状态, 0) = 1))" & vbNewLine & _
            "          And a.记录状态 In (1, 3)　and a.价格父号 Is Null " & vbNewLine & _
            "          And Not Exists" & _
            "                (Select 1 From 病人医嘱附费 Where a.医嘱序号 = 医嘱id and a.No = NO and Mod(a.记录性质, 10) = 记录性质)" & _
            "       ) Q1" & vbNewLine & _
            " Where Not Exists (Select 1 From 药品收发记录 Where 费用id = Q1.Id)" & vbNewLine & _
            " Group by 医嘱ID,收费细目ID  Having Max(ID)<>0 )"
        Else
            '     And A.费用性质=0 :61879,经与张永康确认,费用性质在门诊只有0-基础费用
            strTableNo = strTableNo & " Union ALL  " & _
             " Select Max(ID) as 费用ID,decode(sign(Sum(数量)),-1,0,Sum(数量)) as 准退数 " & _
             " From (  Select decode(J.记录状态,2,0,J.ID) as ID,J.医嘱序号 as 医嘱ID,J.收费细目ID,nvl(J.付数,1)*nvl(J.数次,1) as 数量 " & _
             "              From  门诊费用 J,病人医嘱记录 M " & _
             "              Where  J.医嘱序号=M.ID  " & _
             "                      And Exists(Select 1 From   病人医嘱发送 where 医嘱ID=J.医嘱序号 and  Nvl( 执行状态, 0) <> 1 And No =J.NO  ) " & _
            "                       And Exists(Select 1 From   病人医嘱计价 A Where   A.医嘱ID=J.医嘱序号 and A.收费细目ID=J.收费细目ID And A.费用性质=0  And  Nvl( A.收费方式, 0) =0 ) " & _
             "                      And J.记录状态 in (1,2,3) and J.价格父号 is null   " & _
             "                      And Instr('5,6,7', j.收费类别) = 0 And  Not Exists  (Select 1  From 材料特性  Where 材料id = j.收费细目id And Nvl(跟踪在用, 0) = 1)  " & _
             "                      And  instr(',C,D,F,G,K,',','||M.诊疗类别||',')=0 " & _
             "              Union all  " & _
             "              Select j.id, A.医嘱ID,a.收费细目ID,-1*nvl(a.数量,1)*nvl(C.本次数次,1) as 数量 " & _
             "              From 病人医嘱计价 A,病人医嘱发送 B,病人医嘱执行 C,门诊费用 J,病人医嘱记录 M " & _
             "              Where  A.医嘱ID=b.医嘱id And A.费用性质=0  and  Nvl( A.收费方式, 0) =0  and b.医嘱id=c.医嘱id and b.发送号=c.发送号 And a.医嘱id=M.ID " & _
             "                      And Nvl(C.执行结果, 1) =1 And Nvl(b.执行状态, 0) <> 1 And B.NO=J.No and B.记录性质=1 " & _
             "                      And a.医嘱id=J.医嘱序号 and a.收费细目id=j.收费细目id  " & _
             "                      And  J.记录状态 in (1,3) and J.价格父号 is null   " & _
             "                      And Instr('5,6,7', j.收费类别) = 0 And  Not Exists  (Select 1  From 材料特性  Where 材料id = j.收费细目id And Nvl(跟踪在用, 0) = 1)  " & _
             "                      And  instr(',C,D,F,G,K,',','||M.诊疗类别||',')=0  " & _
              "       ) " & _
             " group by 医嘱ID,收费细目ID  Having Max(ID) <>0)"
        End If
        '整张单据汇总结果(明细到收费细目)
        '执行状态应该在原始记录上判断(部分退药且部份退费的记录)
        '当退两次以上时"记录状态,序号"重复,AVG有问题,所以要用"执行状态"

        
   '58077:需要排开医嘱计划中不为正常收取的费用:
        '   0-正常收取，1-检验试管费用；2-一次发送只收取一次；3-当天只收取一次；4-当天未执行收取一次；5-当天只收取一次，排斥其他项目；6-当天未执行收取一次，排斥其他项目；7-每天首次不收取
        '
        Dim strSQLIn As String
        
        strSQLIn = "" & _
            "  Select NO,Nvl(价格父号,序号) as 序号 From 门诊费用  " & _
            "  Where 记录性质=1 And 记录状态 IN( 1,3)  And Nvl(执行状态,0)<>1     " & _
            "   Minus " & _
            "  Select NO,Nvl(价格父号,序号) as 序号 " & _
            "  From 门诊费用 A1,病人医嘱计价 B1 " & _
            "  Where A1.医嘱序号=B1.医嘱id And A1.收费细目ID=B1.收费细目ID And B1.费用性质=0  And Nvl( B1.收费方式, 0) <>0  " & _
            "           And A1.记录性质=1 And A1.记录状态 IN(1,3)  And Nvl(A1.执行状态,0)=2 " & _
            "           And Instr('5,6,7', a1.收费类别) = 0 And  Not Exists  (Select 1  From 材料特性  Where 材料id = a1.收费细目id And Nvl(跟踪在用, 0) = 1)  " & _
            "           And Not Exists (Select 1 From 药品收发记录 Where 费用id =a1.Id) "
        
        
        strSQL = _
        " Select A.NO,A.记录状态,A.执行状态,Nvl(A.价格父号,A.序号) as 序号,A.从属父号," & _
        "       A.费别,C.编码 as 类别码,C.名称 as 类别名,A.收费细目ID,B.编码,B.名称,B.规格,Nvl(A.费用类型,B.费用类型) 费用类型," & _
        "       A.计算单位,Max(A.医嘱序号) as 医嘱序号, " & _
        "       Avg(Nvl(A.付数,1)) as 付数,Avg(A.数次/A.换算系数) as 数次," & _
        "       Sum(A.标准单价*A.换算系数) as 单价," & _
        "       Sum(A.应收金额) as 应收金额,Sum(A.实收金额) as 实收金额, " & _
        "       D.名称 as 执行科室,A.执行部门ID,E.名称 as 开单科室" & _
        " From 门诊费用 A,收费项目目录 B,收费项目类别 C,部门表 D,部门表 E" & _
        " Where A.收费细目ID=B.ID And C.编码=A.收费类别" & _
        "       And A.执行部门ID=D.ID(+) And A.开单部门ID=E.ID(+)" & _
        "       And (A.NO,Nvl(A.价格父号,A.序号)) IN( " & strSQLIn & ")  " & _
        "       And A.NO IN( Select NO From 门诊费用 where  记录性质=1 and 记录状态 in (1,3) )" & _
        " Group by A.NO,A.记录状态,A.执行状态,Nvl(A.价格父号,A.序号),A.费别,A.从属父号," & _
        "       C.编码,C.名称,A.收费细目ID,B.编码,B.名称,B.规格,Nvl(A.费用类型,B.费用类型),A.计算单位," & _
        "       D.名称,A.执行部门ID,E.名称,A.药品ID "
        
        '最后计算结果
        '当"准退数量=原始数量"时,付数才保留
        '排开已经全部退费的行(执行状态=0的一种可能)
        '有剩余数量无准退数量的有两种情况：
            '1.无对应的收发记录(如普通费用或不跟踪在用的卫材),这时应用剩余数量
            '2.收发记录中已全部发放,即已全部执行,SQL已排除这种记录
        strSQL = strTableNo & vbCrLf & _
        " Select A.NO,A.序号,A.从属父号,A.费别,A.类别码,A.类别名,A.收费细目ID,A.编码,A.名称,A.规格,A.费用类型,A.计算单位, Max(A.医嘱序号) as 医嘱序号," & _
        "       Decode(Sign(Nvl(C.准退数量,Sum(A.付数*A.数次))-B.原始数量),0,Avg(A.付数),1) as 准退付数," & _
        "       Decode(Sign(Nvl(C.准退数量,Sum(A.付数*A.数次))-B.原始数量),0,Sum(A.数次),Nvl(C.准退数量,Sum(A.付数*A.数次))) as 准退数次," & _
        "       Nvl(C.准退数量,Sum(A.付数*A.数次)) as 准退数量,Sum(A.付数*A.数次) as 剩余数量," & _
        "       A.单价,Sum(A.应收金额) as 剩余应收,Sum(A.实收金额) as 剩余实收,max(q1.记录标志) as 记录标志," & _
        "       A.执行科室,A.执行部门ID,A.开单科室,B.操作员姓名,B.登记时间,B.结帐ID,Max(M.医嘱内容) as 医嘱内容,b.原始数量" & _
        " From (" & strSQL & ") A, 准退数 C,病人医嘱记录 M," & _
        "          ( Select  ID, NO,序号, 收费细目ID,Nvl( 数量,0)/NVL(换算系数,1) as 原始数量,操作员姓名,登记时间,结帐ID" & _
        "            From 门诊费用   " & _
        "            Where  记录状态 IN(1,3) And Nvl( 附加标志,0)<>9 And  价格父号 is NULL )B, " & _
        "            ( Select NO,Max(记录状态) as 记录标志 From 门诊费用  Where 记录状态 in (1,3) Group by NO) Q1" & _
        " Where A.NO=B.NO And A.序号=B.序号 And A.收费细目ID=B.收费细目ID+0  And B.ID=C.费用ID(+)" & _
        "            and A.医嘱序号=M.ID(+) and A.NO=q1.NO(+) " & _
        " Group by A.NO,A.序号,A.从属父号,A.费别,A.类别码,A.类别名,A.收费细目ID,A.编码,A.名称,A.规格,A.费用类型," & _
        "       A.计算单位,A.单价,B.原始数量,C.准退数量,A.执行科室,A.执行部门ID,A.开单科室,B.操作员姓名,B.登记时间,B.结帐ID" & _
        " Having Sum(A.付数*A.数次)<>0"
            
        strSQL = _
        " Select /*+ rule */  A.NO,A.序号,A.从属父号,A.费别,A.类别码,A.类别名,A.编码,Nvl(B.名称,A.名称) as 名称,E1.名称 as 商品名," & _
        "       A.规格,A.费用类型,A.计算单位,A.收费细目ID,A.准退付数 as 付数,A.准退数次 as 数次,A.单价, A.医嘱序号 ," & _
        "       A.剩余应收*(A.准退数量/A.剩余数量) as 应收金额," & _
        "       A.剩余实收*(A.准退数量/A.剩余数量) as 实收金额," & _
        "       A.执行科室,A.执行部门ID,A.开单科室,A.操作员姓名,A.登记时间,A.结帐ID,A.医嘱内容,A.记录标志, " & _
        "       A.原始数量,A.准退数量,A.剩余数量" & _
        " From (" & strSQL & ") A,收费项目别名 B,收费项目别名 E1" & _
        " Where     A.收费细目ID=B.收费细目ID(+) And B.码类(+)=1 And B.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
        "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3 " & _
        " Order by A.NO,A.序号"
    Else
        '读取单据原始内容
        intSign = IIf(mstrDelTime <> "", -1, 1) '数量,金额正负符号
        
        strSQL = "" & _
        " Select A.NO " & _
        " From " & IIf(mblnNOMoved, "H", "") & "门诊费用记录 A,Table(f_Str2list([2])) J" & _
        " Where A.记录性质=1 And A.记录状态 IN(1,3) And A.NO=J.Column_Value"
        
        strSQL = _
        " Select A.结帐ID,A.NO,Nvl(A.价格父号,A.序号) as 序号,A.从属父号,A.费别," & _
        "        A.收费细目ID,C.编码 as 类别码,C.名称 as 类别名,B.编码,B.名称,B.规格,Nvl(A.费用类型,B.费用类型) 费用类型," & _
                IIf(gbln药房单位, "Decode(X.药品ID,NULL,A.计算单位,X." & gstr药房单位 & ")", "A.计算单位") & " as 计算单位," & _
        "       Max(A.医嘱序号) as 医嘱序号,Avg(Nvl(A.付数,1)) as 付数," & _
        "       Avg(" & intSign & "*A.数次" & IIf(gbln药房单位, "/Nvl(X." & gstr药房包装 & ",1)", "") & ") as 数次," & _
        "       Sum(A.标准单价" & IIf(gbln药房单位, "*Nvl(X." & gstr药房包装 & ",1)", "") & ") as 单价," & _
        "       Sum(" & intSign & "*A.应收金额) as 应收金额,Sum(" & intSign & "*A.实收金额) as 实收金额," & _
        "       D.名称 as 执行科室,A.执行部门ID,E.名称 as 开单科室,A.操作员姓名,A.登记时间,Max(A.摘要) as 摘要" & _
        " From " & IIf(mblnNOMoved, "H", "") & "门诊费用记录 A,收费项目目录 B,收费项目类别 C,部门表 D,部门表 E,药品规格 X" & _
        " Where A.收费细目ID=B.ID And C.编码=A.收费类别 And A.收费细目ID=X.药品ID(+)" & _
        "       And A.执行部门ID=D.ID(+) And A.开单部门ID=E.ID(+) And A.记录性质=1" & _
        "       And A.记录状态" & IIf(mstrDelTime <> "", "=2", " IN(1,3)") & " And A.NO IN(" & strSQL & ")" & _
                IIf(mstrDelTime <> "", " And A.登记时间=[1]", "") & _
                IIf(Not gblnShowErr, " And Nvl(A.附加标志,0)<>9", "") & _
        " Group by A.结帐ID,A.NO,Nvl(A.价格父号,A.序号),A.从属父号,A.费别,A.收费细目ID,C.编码,C.名称,B.编码,B.名称," & _
        "       B.规格,Nvl(A.费用类型,B.费用类型),A.计算单位,D.名称,A.执行部门ID,E.名称,X.药品ID,X." & gstr药房单位 & ",A.操作员姓名,A.登记时间"
            
        strSQL = "Select /*+ rule */ " & _
            "       A.结帐ID,A.NO,A.序号,A.从属父号,A.费别,A.类别码,A.类别名,A.编码,Nvl(B.名称,A.名称) as 名称,E1.名称 as 商品名,A.规格,A.费用类型," & _
            "       A.计算单位,A.医嘱序号 ,A.收费细目ID,A.付数,A.数次,A.单价,A.应收金额,A.实收金额,A.执行科室,A.执行部门ID,A.开单科室,A.操作员姓名,A.登记时间,A.摘要,M.医嘱内容, " & _
            "       1 as 记录标志,0 as 原始数量,0 as 准退数量,0 as 剩余数量" & _
            " From (" & strSQL & ") A,收费项目别名 B,收费项目别名 E1,病人医嘱记录 M" & _
            " Where     A.收费细目ID=B.收费细目ID(+) And B.码类(+)=1 And B.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
            "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3 " & _
            "       And A.医嘱序号=M.ID(+) " & _
            " Order by A.NO,A.序号"
    End If
    
    If mstrDelTime <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(mstrDelTime), Replace(strNos, "'", ""))
    Else
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate("1991-01-01"), Replace(strNos, "'", ""))
    End If
    
    Call LoadInvoiceData(Replace(strNos, "'", ""))
    
    str医嘱序号 = ""
    If rsTmp.EOF Then
        Screen.MousePointer = 0
        MsgBox "没有找到与号码""" & strNo & """相关的可以退费的记录。" & _
            vbCrLf & "这些收费记录可能已经退费或已经完全执行。", vbInformation, gstrSysName
        Call ClearFace(False)
        Exit Function
    End If
    
    mstrNOsOverFlow = ""
    If mbytMode = 1 Then
        strTmp = ""
        For i = 0 To UBound(Split(strNos, ","))
            strTmp = Replace(Split(strNos, ",")(i), "'", "")
            '检查是否金额超过上限
            If Not BillOperCheck(2, rsTmp!操作员姓名, rsTmp!登记时间, "退费", strTmp, , 1, True) Then
                mstrNOsOverFlow = mstrNOsOverFlow & " ," & strTmp
            End If
        Next
        If mstrNOsOverFlow <> "" Then mstrNOsOverFlow = Mid(mstrNOsOverFlow, 2)
        If mBillDelType = EM_多张全退 And mstrNOsOverFlow <> "" Then
            Screen.MousePointer = 0
            MsgBox "多张单据使用一卡通模式或医保退费要求整体退，不允许部分退费！", vbInformation, gstrSysName
            Call ClearFace(False)
            Exit Function
        End If
    End If
    
    If mbytMode = 0 Or mbytMode = 2 Then
        pic退费摘要.Enabled = False
        txt退费摘要.Text = Nvl(rsTmp!摘要)
    End If
    
    mCurBillType.bln单张部分退费 = False
    With vsBill
        .Redraw = flexRDNone
        .Rows = .FixedRows + rsTmp.RecordCount
        mstrNOs = ""
        For i = 1 To rsTmp.RecordCount
            '问题:29201
            .Cell(flexcpData, i, .ColIndex("项目")) = Nvl(rsTmp!从属父号)
            '问题:33634
            .Cell(flexcpData, i, .ColIndex("结帐ID")) = Nvl(rsTmp!医嘱序号) & "," & Nvl(rsTmp!收费细目ID)
            If mbytMode = 1 Then
                If Val(Nvl(rsTmp!医嘱序号)) <> 0 And InStr(str医嘱序号 & ",", "," & Nvl(rsTmp!医嘱序号) & ",") = 0 Then
                    str医嘱序号 = str医嘱序号 & "," & Nvl(rsTmp!医嘱序号)
                End If
            End If
            strTemp = ""
            If Val(Nvl(rsTmp!从属父号)) <> 0 Then
                rsTmp.MoveNext
                strTemp = "┣"
                If rsTmp.EOF Then
                    strTemp = "┗"
                ElseIf Val(.Cell(flexcpData, i, .ColIndex("项目"))) <> Nvl(rsTmp!从属父号) Then
                    strTemp = "┗"
                End If
                rsTmp.MovePrevious
                strTemp = "  " & strTemp & " "
            End If
    
            .RowData(i) = CLng(rsTmp!序号)
            .TextMatrix(i, .ColIndex("选择")) = 0
            
            .TextMatrix(i, .ColIndex("单据号")) = rsTmp!NO
            .TextMatrix(i, .ColIndex("类别")) = rsTmp!类别名
            .TextMatrix(i, .ColIndex("项目")) = strTemp & rsTmp!名称 & IIf(IsNull(rsTmp!规格), "", " " & rsTmp!规格)
            .TextMatrix(i, .ColIndex("商品名")) = strTemp & Nvl(rsTmp!商品名)
            .TextMatrix(i, .ColIndex("数量")) = FormatEx(Nvl(rsTmp!付数, 1) * rsTmp!数次, 5)
            .TextMatrix(i, .ColIndex("单位")) = Nvl(rsTmp!计算单位)
            .TextMatrix(i, .ColIndex("单价")) = Format(rsTmp!单价, gstrFeePrecisionFmt)
            .TextMatrix(i, .ColIndex("应收金额")) = Format(rsTmp!应收金额, gstrDec)
            .TextMatrix(i, .ColIndex("实收金额")) = Format(rsTmp!实收金额, gstrDec)
            .TextMatrix(i, .ColIndex("开单科室")) = Nvl(rsTmp!开单科室)
            .TextMatrix(i, .ColIndex("执行科室")) = Nvl(rsTmp!执行科室)
            .TextMatrix(i, .ColIndex("操作员")) = rsTmp!操作员姓名
            .TextMatrix(i, .ColIndex("时间")) = Format(rsTmp!登记时间, "MM-dd HH:mm")
            .TextMatrix(i, .ColIndex("结帐ID")) = rsTmp!结帐ID
            .TextMatrix(i, .ColIndex("医嘱")) = Nvl(rsTmp!医嘱内容)
            .TextMatrix(i, .ColIndex("原始数量")) = Nvl(rsTmp!原始数量)
            .TextMatrix(i, .ColIndex("准退数量")) = Nvl(rsTmp!准退数量)
            .TextMatrix(i, .ColIndex("医嘱序号")) = Nvl(rsTmp!医嘱序号)
            .TextMatrix(i, .ColIndex("执行科室ID")) = Nvl(rsTmp!执行部门ID)
            If Not mCurBillType.bln单张部分退费 Then mCurBillType.bln单张部分退费 = RoundEx(Val(Nvl(rsTmp!原始数量)), 7) <> RoundEx(Val(Nvl(rsTmp!准退数量)), 7)
            .Cell(flexcpData, i, .ColIndex("选择")) = Val(Nvl(rsTmp!记录标志))    '用于判断是否被销帐过,>1表示已销帐
            If Val(Nvl(rsTmp!记录标志)) > 1 And InStr(1, mstrNOsPatiDel & ",", "," & rsTmp!NO & "") = 0 Then mstrNOsPatiDel = mstrNOsPatiDel & "," & rsTmp!NO
            If InStr(mstrNOs & ",", "," & rsTmp!NO & ",") = 0 Then
                '画出分隔线
                If mstrNOs <> "" Then
                    .Select i, .FixedCols, i, .COLS - 1
                    .CellBorder vbBlue, 0, 1, 0, 0, 0, 0
                End If
                mstrNOs = mstrNOs & "," & rsTmp!NO
            End If
            cur合计 = cur合计 + rsTmp!实收金额
            rsTmp.MoveNext
        Next
        .Row = .FixedRows: .Col = .ColIndex("项目")
        Call vsBill_AfterRowColChange(-1, -1, .Row, .Col)
        .Redraw = flexRDDirect
    End With
    If str医嘱序号 <> "" Then
        Set mrs收费对照 = zlGet诊疗收费对照(Mid(str医嘱序号, 2))
    Else
        Set mrs收费对照 = Nothing
    End If
    
    Call SetpicInvoiceVisible   '设置发票控件的显示
    If mbytMode = 1 Or mbytMode = 2 Then
        '--问题:31179:主要解决医保病人先收费的后退费的处理.为了保持以前的歉容，因此没有直接在SQL中排序(界面上还是体现的按单据号进行排序处理)
        rsTmp.Sort = "结帐ID,NO,序号"
        If rsTmp.RecordCount <> 0 Then rsTmp.MoveFirst
        With rsTmp
                mstrNOs = ""
                Do While Not .EOF
                        If InStr(1, mstrNOs & ",", "," & Nvl(rsTmp!NO) & ",") = 0 Then
                            mstrNOs = mstrNOs & "," & Nvl(rsTmp!NO)
                            If Not mCurBillType.bln单张部分退费 Then mCurBillType.bln单张部分退费 = Not BillDeleteAll(Nvl(rsTmp!NO), 1, mblnHaveExcuteData)
                        End If
                        .MoveNext
                Loop
        End With
    End If
     If mstrNOs <> "" Then mstrNOs = Mid(mstrNOs, 2)
    txtAllTotal.Text = Format(cur合计, gstrDec)
    If mbytMode = 1 Then
        If strInvoiceNO <> "" And gTy_Module_Para.byt票据分配规则 <> 0 Then
            If mBillDelType = EM_单张部分退 Or mBillDelType = EM_单张全退 Then
                '只有单张部分退,才会存在部分选的问题
                vsBill.Cell(flexcpChecked, 1, vsBill.ColIndex("选择"), vsBill.Rows - 1, vsBill.ColIndex("选择")) = 0
                Call FromInvoiceSelectNO(strInvoiceNO)
            End If
            If mBillDelType <> EM_多张全退 Then
                Call SelectRelatingInvoice(strInvoiceNO, True)
                '仅显示被勾选的发票
                'Call OlnyShowSelectedInvoice
                Call ShowAndHideDelBillRow
            End If
        Else
            '78569,冉俊明,2014-10-14,默认勾选单据
            If mBillDelType = EM_单张部分退 Or mBillDelType = EM_单张全退 Then
                With vsBill
                    For i = .FixedRows To .Rows - 1
                        If .TextMatrix(i, .ColIndex("单据号")) = strOldNO Then
                            .Row = i: Exit For
                        End If
                    Next
                End With
                Call cmdBillSel_Click
            End If
        End If
        '40391
        If mBillDelType = EM_单张部分退 Or mBillDelType = EM_单张全退 Then
            Call LoadBalanceInfor
            Call LoadDelBalanceInfor
            Call ReCalcDelMoney
            Call FromNoSelectInvoice
        End If
        If mBillDelType = EM_多张全退 Then
            Call cmdSelAll_Click
        End If
        '78569,冉俊明,2014-10-14,默认勾选单据
        If InStr(";" & mstrPrivs & ";", ";部份退费;") = 0 Then Call cmdSelAll_Click
    Else
         Call cmdSelAll_Click
    End If
 
    If mbytMode = 1 Then
        cmdSelAll.Visible = mBillDelType <> EM_多张全退
        cmdClear.Visible = mBillDelType <> EM_多张全退
        cmdBillSel.Visible = mBillDelType <> EM_多张全退
    End If
    Screen.MousePointer = 0
    Call ReInitPatiInvoice
    ReadBills = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub CalcDelMoney()
'功能：根据当前界面退费选择情况，计算退款金额和误差金额
    Dim cur单据合计 As Currency, cur选择合计 As Currency
    Dim cur退费合计 As Currency, cur误差金额 As Currency, cur误差合计 As Currency
    Dim bln完全退费 As Boolean, bln现金结算 As Boolean
    Dim curTotal As Currency, strNo As String
    Dim i As Long, j As Long, k As Long, bln原样退 As Boolean
    Dim colAllReturn As Collection
    
    If mbytMode = 0 Then Exit Sub
    If mrsBalance Is Nothing Then Exit Sub
        
    Set mcolError = New Collection
    Set colAllReturn = New Collection
    
    If mBillDelType = EM_多张全退 Then
        '多张单据一起退,无误差
        For i = 0 To UBound(Split(mstrNOs, ","))
            strNo = CStr(Split(mstrNOs, ",")(i))
            mcolError.Add 0, "_" & strNo
        Next
        If cbo退款方式.ListIndex = -1 And cbo退款方式.ListCount > 0 Then cbo退款方式.ListIndex = 0
        cbo退款方式.Enabled = False
        cbo退款方式.Locked = True
        
        curTotal = 0
        ''不包括医保和预交款金额
         '性质:1-预存款,2-医保,3-医疗卡,4-结算卡,5-一卡通,0-其他类
         mrsBalance.Filter = "性质<>1 And 性质<>2"
        With mrsBalance
             If .RecordCount <> 0 Then .MoveFirst
             Do While Not .EOF
                Select Case Val(Nvl(!性质))
                Case 3, 4, 5    '3-医疗卡,4-结算卡,5-一卡通
                Case Else
                    curTotal = curTotal + !结算金额
                End Select
                 .MoveNext
            Loop
            .Filter = 0
        End With
        txt退款金额.Text = curTotal
        Exit Sub
    End If
    
    '1.先判断整个是否是原样退,以决定是否禁用结算方式选择,以及分币误差的生成
    bln原样退 = True
    For i = 0 To UBound(Split(mstrNOs, ","))
        strNo = CStr(Split(mstrNOs, ",")(i))
        cur单据合计 = 0: cur选择合计 = 0
        With vsBill
            k = vsBill.FindRow(strNo, , .ColIndex("单据号"))
            For j = k To vsBill.Rows - 1
                If vsBill.TextMatrix(j, .ColIndex("单据号")) <> strNo Then Exit For
                cur单据合计 = cur单据合计 + Val(vsBill.TextMatrix(j, .ColIndex("实收金额")))
                If Val(vsBill.TextMatrix(j, .ColIndex("选择"))) <> 0 Then
                    cur选择合计 = cur选择合计 + Val(vsBill.TextMatrix(j, .ColIndex("实收金额")))
                End If
            Next
        End With
        bln完全退费 = Not BillExistDelete(strNo, 1) And BillDeleteAll(strNo, 1, mblnHaveExcuteData) And (cur单据合计 = cur选择合计)
        colAllReturn.Add Array(IIf(bln完全退费, 1, 0), strNo, cur单据合计, cur选择合计), "_" & strNo   '保存用于后面的判断
        If Not bln完全退费 Then bln原样退 = False
        
        If bln完全退费 Then
            mrsBalance.Filter = "NO='" & strNo & "'"
            With mrsBalance
                If .RecordCount <> 0 Then .MoveFirst
                Do While Not .EOF
                    If Val(Nvl(!性质)) = 2 Then '医保
                        If Not mblnYB结算作废 Then bln原样退 = False
                        If mblnYB结算作废 Then
                            If Not gclsInsure.GetCapability(support门诊结算作废, mlng病人ID, mintInsure, !结算方式) Then
                               bln原样退 = False
                            End If
                        End If
                    ElseIf InStr("3,4,5", Val(Nvl(!性质))) > 0 Then
                        '一卡通相关
                        'If Nvl(!是否退现) = 1 Then bln原样退 = False
                    End If
                    .MoveNext
                Loop
            End With
        End If
    Next
    
    '收费时全部用预交(结算方式为空),退费时,不允许指定退费方式
    '性质:1-预存款,2-医保,3-医疗卡(一卡通),4-结算卡,5-一卡通,0-其他类
    mrsBalance.Filter = "性质<>1"
    If mrsBalance.RecordCount = 0 Then bln原样退 = True
    mrsBalance.Filter = 0
    If mBillDelType = EM_单张全退 Then bln原样退 = True
    
    If bln原样退 Then
        zlControl.CboSetIndex cbo退款方式.hWnd, mintReturnMode
    End If
    
    cbo退款方式.Enabled = Not bln原样退
    cbo退款方式.Locked = bln原样退
    
    '2.计算退款金额及误差
    If cbo退款方式.ListIndex <> -1 Then
        If cbo退款方式.ItemData(cbo退款方式.ListIndex) = 1 Then
            bln现金结算 = True
        End If
    End If
    Dim varTemp As Variant
    
    For i = 0 To colAllReturn.Count     ' UBound(Split(mstrNOs, ","))
        '0-是否完全退费;1-NO,2-单据合计,3-选择合计
        varTemp = colAllReturn(i)
        strNo = varTemp(1)
        cur单据合计 = Val(varTemp(2)): cur选择合计 = Val(varTemp(3))
        cur退费合计 = 0: cur误差金额 = 0
        '完全退费时排开医保结算及冲预交金额
        bln完全退费 = IIf(Val(varTemp(0)) = 1, True, False)
        If bln完全退费 Then
            mrsBalance.Filter = "NO='" & strNo & "'"
            '性质:1-预存款,2-医保,3-医疗卡,4-结算卡,5-一卡通,0-其他类
            With mrsBalance
                Do While Not .EOF
                    Select Case Val(Nvl(!性质))
                    Case 1 '预交款
                         cur选择合计 = cur选择合计 - Nvl(!结算金额, 0)
                    Case 2 '医保
                        '如果这种结算方式不支持回退,要退为现金,则不用减去
                        If mblnYB结算作废 Then
                            If gclsInsure.GetCapability(support门诊结算作废, mlng病人ID, mintInsure, !结算方式) Then
                                cur选择合计 = cur选择合计 - Nvl(!结算金额, 0)
                            End If
                        Else     '不支持门诊结算作废时,只允许个帐退为现金,其它原样退,不调用医保交易
                            If !结算方式 <> mstr个人帐户 Then
                                cur选择合计 = cur选择合计 - Nvl(!结算金额, 0)
                            End If
                        End If
                    Case 3, 4 '医疗卡和结算卡
                        If Val(Nvl(!是否退现)) = 0 Then
                            cur选择合计 = cur选择合计 - Nvl(!结算金额, 0)
                        End If
                    Case 5 '一卡通
                            cur选择合计 = cur选择合计 - Nvl(!结算金额, 0)
                    Case Else
                    End Select
                    .MoveNext
                Loop
            End With
        End If
        
        '费用金额保留位数,及现金结算时处理分币
        If bln现金结算 Then
            If mintInsure > 0 Then
                If gclsInsure.GetCapability(support分币处理, mlng病人ID, mintInsure) Then
                    cur退费合计 = CentMoney(cur选择合计)
                Else
                    cur退费合计 = Format(cur选择合计, "0.00")
                End If
            Else
                cur退费合计 = CentMoney(cur选择合计)
            End If
        Else
            cur退费合计 = Format(cur选择合计, "0.00")
        End If
        
        '误差金额,部分退,或医保全退时因为结算方式不支持回退而退为现金,可能产生误差
        '非现金结算时,也可能有误差,这个误差是费用金额保留位数引起的
        If Not bln原样退 Then
            cur误差金额 = cur退费合计 - cur选择合计
        End If
        
        curTotal = curTotal + cur退费合计
        mcolError.Add cur误差金额, "_" & strNo
        cur误差合计 = cur误差合计 + cur误差金额
    Next
    
    txt退款金额.ToolTipText = "退费误差金额:" & Format(cur误差合计, gstrDec)
    txt退款金额.Text = Format(curTotal, "0.00")
    Call Show退款方式(cbo退款方式.Enabled)
    
End Sub

Private Sub Show退款方式(ByVal blnVisible As Boolean)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示退款方式
    '入参:blnVisible-true,显示,否则隐藏
    '编制:刘兴洪
    '日期:2012-09-12 11:27:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    cbo退款方式.Visible = blnVisible
    lbl退款方式.Visible = blnVisible
    lbl退款金额.Visible = blnVisible
    txt退款金额.Visible = blnVisible
End Sub


Private Function CheckOnCardValied(ByVal blnCur部分退费 As Boolean, ByVal lng结帐ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查一卡通是否合法
    '返回:合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-31 12:00:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCardNo As String
    On Error GoTo errHandle
    If Not mblnOneCard Then CheckOnCardValied = True: Exit Function
    mrsBalance.Filter = "结帐ID=" & lng结帐ID & " And 性质=5"
    If mrsBalance.RecordCount = 0 Then CheckOnCardValied = True: Exit Function
    If blnCur部分退费 Then
         MsgBox "当前单据使用了一卡通结算,不能进行部分退费！", vbInformation, gstrSysName
        Exit Function
    End If
    If mobjICCard Is Nothing Then
        On Error Resume Next
        Set mobjICCard = CreateObject("zlICCard.clsICCard")
        On Error GoTo 0
        If mobjICCard Is Nothing Then
            MsgBox "一卡通接口创建失败,不能进行退费!请检查接口文件.", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    With mrsBalance
        'mObjSquare.objSquareCard
        'strCardNo = objICCard.Read_Card(Me)
        '弹出刷卡界面
        'zlBrushCard(frmMain As Object, _
        'ByVal lngModule As Long, _
        'ByVal rsClassMoney As ADODB.Recordset, _
        'ByVal lngCardTypeID As Long, _
        'ByVal bln消费卡 As Boolean, _
        'ByVal strPatiName As String, ByVal strSex As String, _
        'ByVal strOld As String, ByVal dbl金额 As Double, _
        'Optional ByRef strCardNo As String, _
        'Optional ByRef strPassWord As String) As Boolean
        If mobjSquare.zlBrushCard(Me, mlngModule, Nothing, 0, False, _
          mtyPati.姓名, mtyPati.性别, mtyPati.年龄, 0, strCardNo, "") = False Then Exit Function
        If strCardNo = "" Then Exit Function
        If strCardNo <> Nvl(!卡号) Then
            MsgBox "当前卡号与扣款卡号不一致,不能进行退费.", vbInformation, gstrSysName
            Exit Function
        End If
    End With
    CheckOnCardValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckThreeSwapValied(ByVal blnCur部分退费 As Boolean, _
    ByVal lng结帐ID As Long, ByVal blnMulitNo As Boolean, Optional bln异常单据 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:三方交易检查
    '入参:blnCur部分退费
    '        lng结帐ID
    '       blnMulitNo-是否多单据
    '       bln异常单据-异常单据重新收费
    '出参:
    '返回:数据合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-31 15:35:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim bln退现 As Boolean, intCol As Integer
    Dim str三方结算方式 As String
    On Error GoTo errHandle
    If cbo退款方式.Visible Then
        mrsBalance.Filter = "结帐ID=" & lng结帐ID & " And 性质>=3 and 性质<=4  and 结算方式='" & zlStr.NeedName(cbo退款方式.Text) & "'"
        With mrsBalance
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                 str三方结算方式 = str三方结算方式 & "," & Nvl(!结算方式)
                .MoveNext
            Loop
        End With
        mrsBalance.Filter = 0
    End If
    mrsBalance.Filter = "结帐ID=" & lng结帐ID & " And 性质>=3 and 性质<=4 "
    If mrsBalance.RecordCount = 0 Then CheckThreeSwapValied = True: Exit Function
    With mrsBalance
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            bln退现 = False
            If Val(Nvl(!是否退现)) = 1 Then
                '排队交现金
                For intCol = 1 To vsBalance.COLS - 1 Step 2
                    If vsBalance.Cell(flexcpData, 1, intCol) = !结算方式 And _
                        (Val(vsBalance.TextMatrix(1, intCol + 1)) = 0) Or vsBalance.RowHidden(1) Then
                         bln退现 = True
                    End If
                Next
            End If
        
            If blnCur部分退费 And Val(Nvl(!是否全退)) = 1 And blnMulitNo Then
                If Val(Nvl(!是否退现)) <> 1 Then
                    MsgBox "当前单据使用了第三方结算交易,所有单据必须全退！", vbInformation, gstrSysName
                    Exit Function
                ElseIf bln退现 = False Then
                    MsgBox "当前单据使用了第三方结算交易,所有单据必须全退！", vbInformation, gstrSysName
                    Exit Function
                ElseIf InStr(str三方结算方式 & ",", "," & zlStr.NeedName(cbo退款方式.Text) & ",") > 0 Then
                    MsgBox "当前单据使用了第三方结算交易,所有单据必须全退！", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
            
            If cbo退款方式.Visible And cbo退款方式.Enabled Then
                If bln退现 And blnCur部分退费 And InStr(str三方结算方式 & ",", "," & zlStr.NeedName(cbo退款方式.Text) & ",") > 0 Then
                    MsgBox "当前单据使用了第三方结算交易,单张单据必须全退！", vbInformation, gstrSysName
                    Exit Function
                End If
                If InStr(str三方结算方式 & ",", "," & zlStr.NeedName(cbo退款方式.Text) & ",") > 0 Then
                    MsgBox "部分退费时,不允许选择三方结算", vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                End If
            End If
            
            If Not bln退现 And Not blnCur部分退费 Then
                '1.未退现金
                '2.不是部分退的(单张)
                If zlCheckDelValied(Val(Nvl(!卡类别ID)), Nvl(!名称), Val(Nvl(!性质)) = 4, Nvl(!卡号), Nvl(!交易流水号), Nvl(!交易说明), lng结帐ID, Val(Nvl(!结算金额)), bln异常单据) = False Then Exit Function
                If Val(Nvl(!是否退款验卡)) = 1 And Val(Nvl(!性质)) = 4 Then
                    '需要验卡
                    If CheckBrushCard(Val(Nvl(!卡类别ID)), Val(Nvl(!性质)) = 4, Val(Nvl(!结算金额)), Nvl(!卡号), "", bln退现) = False Then Exit Function
                End If
            End If
            .MoveNext
        Loop
    End With
    CheckThreeSwapValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function DelInsureMulitOneBalance(ByVal blnExistThreeSwap As Boolean, _
     ByVal arrNo As Variant, ByVal lng结帐ID As Long, ByVal strAllBalance As String, _
     ByVal str医保结算 As String, ByVal str退结算方式 As String, ByVal bln退现 As Boolean, _
    ByRef blnCommited As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:多单据一次结算退费
    '入参:arrNO-本次退的单据号
    '       str退结算方式-退的结算方式
    '       bln退现-退的结算方式是否现金
    '出参:
    '返回:成功或非医保或非多单据一次结算,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-31 23:38:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strAdvance As String, arrBalance As Variant, str结算方式 As String
    Dim dbl结算金额 As Double, dbl可分配额 As Double, dbl余额 As Double
    Dim strBalance As String, dbl退款合计 As Double, str退回结算 As String
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Dim i As Long, k As Long, j As Long, cur误差金额 As Double
    Dim lng冲销ID As Long
    
    On Error GoTo errHandle
    If mintInsure = 0 Then DelInsureMulitOneBalance = True: Exit Function
    If Not (MCPAR.多单据一次结算 Or MCPAR.多单据调一次交易) Then DelInsureMulitOneBalance = True: Exit Function
    
    strAdvance = strAllBalance
    If blnExistThreeSwap Then
        
        ' Zl_门诊结算_较对标志_Update
        strSQL = "Zl_门诊结算_较对标志_Update("
        '  结帐id_In     门诊费用记录.结帐id%Type,
        strSQL = strSQL & "" & lng结帐ID & ","
        '  结算序号id_In 病人预交记录.结算序号%Type,
        strSQL = strSQL & "NULL,"
        '  收费结算_In   Varchar2,
        strSQL = strSQL & "'" & str医保结算 & "',"
        '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
        strSQL = strSQL & "NULL,"
        '  消费卡_In     Integer := 0,
        strSQL = strSQL & "0,"
        '  卡号_In       病人预交记录.卡号%Type := Null,
        strSQL = strSQL & "NULL,"
        '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
        strSQL = strSQL & "NULL,"
        '  交易说明_In   病人预交记录.交易说明%Type := Null,
        strSQL = strSQL & "NULL,"
        '  校对标志_In   病人预交记录.校对标志%Type := 0
        strSQL = strSQL & "2)"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    If mintInsure <> 0 And mblnYB结算作废 Then
        If Not gclsInsure.ClinicDelSwap(lng结帐ID, , mintInsure, strAdvance) Then Exit Function
    Else
        strAdvance = ""
    End If
    
    If strAdvance = strAllBalance Or strAdvance = "" Then
        gcnOracle.CommitTrans: blnCommited = True
        If mblnYB结算作废 Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, mintInsure)
        DelInsureMulitOneBalance = True: Exit Function
    End If
    
    '根据返回的结算信息，修正预交记录，strAdvance返回格式:结算方式1|金额||结算方式2:金额...
    '先分摊到每张单据上
    Set rsTmp = GetBalanceSet
    arrBalance = Split(strAdvance, "||")
    For i = 0 To UBound(arrBalance)
        str结算方式 = Split(arrBalance(i), "|")(0)
        dbl结算金额 = -1 * Val(Split(arrBalance(i), "|")(1))
        For k = 0 To UBound(arrNo)
            dbl可分配额 = Get实收金额(arrNo(k))
            rsTmp.Filter = "单据序号=" & k
            For j = 1 To rsTmp.RecordCount
                dbl可分配额 = dbl可分配额 - rsTmp!结算金额
                rsTmp.MoveNext
            Next
            If dbl可分配额 > 0 Then
                If dbl可分配额 <= dbl结算金额 Then
                    dbl结算金额 = dbl结算金额 - dbl可分配额
                Else
                    dbl可分配额 = dbl结算金额
                    dbl结算金额 = 0
                End If
                rsTmp.AddNew
                rsTmp!单据序号 = k
                rsTmp!结算方式 = str结算方式
                rsTmp!结算金额 = dbl可分配额
                rsTmp.Update
                If dbl结算金额 = 0 Then Exit For
            End If
        Next
    Next
    For k = 0 To UBound(arrNo)
        strBalance = "": cur误差金额 = 0
        dbl余额 = Get实收金额(arrNo(k))
        rsTmp.Filter = "单据序号=" & k
        For i = 1 To rsTmp.RecordCount
            strBalance = IIf(strBalance = "", "", strBalance & "||") & rsTmp!结算方式 & "|" & -1 * rsTmp!结算金额
            dbl余额 = dbl余额 - rsTmp!结算金额
            rsTmp.MoveNext
        Next
        '退为指定的结算方式，如果是现金，可能产生新的误差金额
        dbl结算金额 = dbl余额
        If bln退现 Then
            dbl结算金额 = Format(CentMoney(dbl余额), "0.00")
            cur误差金额 = dbl结算金额 - dbl余额
        End If
        dbl退款合计 = dbl退款合计 + dbl结算金额
        str退回结算 = str退结算方式 & "|" & -1 * dbl结算金额 & "| "
        lng结帐ID = GetDelBalanceID(arrNo(k))
        If Not blnExistThreeSwap Then
            strSQL = "zl_门诊收费结算_Update(" & lng结帐ID & ",'" & str退回结算 & "',0,'" & _
                strBalance & "'," & -1 * cur误差金额 & ",NULL,NULL,NULL,1)"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        Else
            'Zl_医保结算校对_Update
             strSQL = "Zl_医保结算校对_Update("
             '  结帐id_In   门诊费用记录.结帐id%Type,
             strSQL = strSQL & "" & lng冲销ID & ","
             '  保险结算_In Varchar2
             strSQL = strSQL & "'" & strBalance & "')"
             Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        End If
    Next
    gcnOracle.CommitTrans: blnCommited = True
    Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, mintInsure)
    If Not blnExistThreeSwap Then
        If Not (strAdvance = strAllBalance Or strAdvance = "") Then
            MsgBox "应退金额" & vbCrLf & str退结算方式 & "：" & Format(dbl退款合计, "0.00") & "元", vbInformation + vbOKOnly, gstrSysName
        End If
    End If
    DelInsureMulitOneBalance = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans: blnCommited = True
    Call ErrCenter
    Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, False, mintInsure)
End Function
Private Function DelInsureOneBill(ByVal str医保结算 As String, ByVal blnExistThreeSwap As Boolean, _
     ByVal lng结帐ID As Long, _
     ByVal lngPage As Long, ByVal lngPages As Long, _
     ByRef blnCommited As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用医保接口退费
    '入参:blnExistThreeSwap-存在第三方接口
    '        lng结帐ID-结帐ID
    '       lngPage(lngPages)-当前页(当前页数)
    '出参:
    '返回:退费成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-09-01 01:08:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strAdvance As String
    Dim blnTransMedicare As Boolean
    On Error GoTo errHandle
    blnTransMedicare = False
    If mintInsure = 0 Then DelInsureOneBill = True: Exit Function
    If Not mblnYB结算作废 Or lng结帐ID = 0 Then DelInsureOneBill = True: Exit Function
    If blnExistThreeSwap Then
        ' Zl_门诊结算_较对标志_Update
        strSQL = "Zl_门诊结算_较对标志_Update("
        '  结帐id_In     门诊费用记录.结帐id%Type,
        strSQL = strSQL & "" & lng结帐ID & ","
        '  结算序号id_In 病人预交记录.结算序号%Type,
        strSQL = strSQL & "NULL,"
        '  收费结算_In   Varchar2,
        strSQL = strSQL & "'" & str医保结算 & "',"
        '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
        strSQL = strSQL & "NULL,"
        '  消费卡_In     Integer := 0,
        strSQL = strSQL & "0,"
        '  卡号_In       病人预交记录.卡号%Type := Null,
        strSQL = strSQL & "NULL,"
        '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
        strSQL = strSQL & "NULL,"
        '  交易说明_In   病人预交记录.交易说明%Type := Null,
        strSQL = strSQL & "NULL,"
        '  校对标志_In   病人预交记录.校对标志%Type := 0
        strSQL = strSQL & "2)"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    strAdvance = lngPages & "|" & lngPage
    'strAdvance = CStr(UBound(arrNO) + 1) & "|" & CStr(UBound(arrNO) + 1 - i)
    If Not gclsInsure.ClinicDelSwap(lng结帐ID, , mintInsure, strAdvance) Then GoTo errHandle:
    blnTransMedicare = True
    gcnOracle.CommitTrans: blnCommited = True
    Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, mintInsure)
    DelInsureOneBill = True
    Exit Function
errHandle:
    '50134
    gcnOracle.RollbackTrans: blnCommited = True
    If blnTransMedicare Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, False, mintInsure)
    If Err.Number <> 0 Then
        Call ErrCenter
    End If
 End Function
Private Function DelOneCardPay(ByVal varNO As Variant, _
     ByRef blnCommited As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:一卡通退费
    '返回:
    '编制:刘兴洪
    '日期:2011-09-01 01:47:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCardNo As String, strSwap As String, strHsptCode As String '医院编码
    Dim i As Long, dblMoney As Double, strNos As String, strSQL As String

    If mblnOneCard = False Then DelOneCardPay = True: Exit Function
    mrsBalance.Filter = "性质=5"
    If mrsBalance.RecordCount = 0 Then
        mrsBalance.Filter = 0
        DelOneCardPay = True: Exit Function
    End If
    For i = 0 To UBound(varNO)
        strNos = strNos & "," & varNO(i)
    Next
    If strNos <> "" Then strNos = Mid(strNos, 2)
    With mrsBalance
        .MoveFirst
        dblMoney = 0
        Do While Not .EOF
            If InStr(1, "," & strNos & ",", "," & Nvl(!NO) & ",") > 0 Then
                strCardNo = Nvl(!卡号): strSwap = Nvl(!交易流水号): strHsptCode = Nvl(!医院编码)
                dblMoney = dblMoney + Nvl(!结算金额)
            End If
            .MoveNext
        Loop
    End With
    mrsBalance.Filter = 0
    If dblMoney = 0 Then DelOneCardPay = True: Exit Function
    On Error GoTo errHandle
    If Not mobjICCard.ReturnSwap(strCardNo, strHsptCode, strSwap, dblMoney) Then
        gcnOracle.RollbackTrans
        MsgBox "一卡通退费交易调用失败！退费操作失败！", vbExclamation, gstrSysName
        Exit Function
    End If
    ' Zl_门诊收费_完成校对
    strSQL = "Zl_门诊收费_完成校对("
    '  No_In       Varchar2,
    strSQL = strSQL & "'" & strNos & "',"
    '  操作类型_In Number,
    '  --操作类型_In:0-一卡通;1-消费卡;2-医疗卡
    strSQL = strSQL & "0,"
    '  卡类别id_In 病人预交记录.卡类别id%Type,
    strSQL = strSQL & "NULL,"
    '  卡号_In     病人预交记录.卡号%Type
    strSQL = strSQL & "'" & strCardNo & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    blnCommited = True: gcnOracle.CommitTrans
    DelOneCardPay = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    Call ErrCenter
    MsgBox "一卡通退费交易调用失败！退费操作失败！", vbExclamation, gstrSysName
End Function

Private Function DelThreeSwapFeeSingle(ByVal varNO As Variant, colThreeBalance As Collection, _
    colOrder As Collection, ByVal str冲销IDs As String, ByRef blnCommited As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:退第三方接口交易,退费成功,返回true,否则返回False
    '参数:
    '       varNO - 本次结算单据号
    '       colThreeBalance - 三方退费信息，NO|退费金额
    '       colOrder - 退费行信息，序号
    '返回:三方交易退费成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-09-01 02:45:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, dblMoney As Double
    Dim i As Long, strNos As String, strSQL As String
    Dim strSelNos As String, str结帐IDs As String, strErrMsg As String
    Dim cllUpdate As Collection, cllThreeSwap As Collection
     
    On Error GoTo errHandle
    If Left(str冲销IDs, 1) = "," Then str冲销IDs = Mid(str冲销IDs, 2)
    
    For i = 0 To UBound(varNO)
        strNos = strNos & "," & varNO(i)
        If colOrder("_" & varNO(i)) <> "未选择" Then
            strSelNos = strSelNos & "," & varNO(i)
            mrsBalance.Filter = "NO='" & varNO(i) & "'"
            If mrsBalance.RecordCount <> 0 Then
                mrsBalance.MoveFirst
                str结帐IDs = str结帐IDs & "," & Nvl(mrsBalance!结帐ID)
            End If
        End If
        If colThreeBalance("_" & varNO(i)) <> "" Then
            '计算总金额
            dblMoney = dblMoney + Val(Split(colThreeBalance("_" & varNO(i)) & "|", "|")(1))
        End If
    Next
    If strNos <> "" Then strNos = Mid(strNos, 2)
    
    mrsBalance.Filter = "性质=3"
    With mrsBalance
        If .RecordCount <> 0 Then
            .MoveFirst
            varData = Array(0, 0, "", "", "", 0, "")
            varData(0) = Val(Nvl(!性质))
            varData(1) = Val(Nvl(!卡类别ID))
            varData(2) = Nvl(!卡号)
            varData(3) = Nvl(!交易流水号)
            varData(4) = Nvl(!交易说明)
            varData(5) = dblMoney
        End If
    End With
    mrsBalance.Filter = 0
    
    If strSelNos = "" Or RoundEx(dblMoney, 5) = 0 Then DelThreeSwapFeeSingle = True: Exit Function
    strSelNos = Mid(strSelNos, 2)
    If str结帐IDs = "" Then str结帐IDs = ",0"
    str结帐IDs = Mid(str结帐IDs, 2)
    'varData = Array(Val(Nvl(!性质)), Val(Nvl(!卡类别ID)), _
                    CStr(Nvl(!卡号)), CStr(Nvl(!交易流水号)), CStr(Nvl(!交易说明)), dblMoney)
    ' Zl_门诊收费_完成校对
    strSQL = "Zl_门诊收费_完成校对("
    '  No_In       Varchar2,
    strSQL = strSQL & "'" & strSelNos & "',"
    '  操作类型_In Number,
    '  --操作类型_In:0-一卡通;1-消费卡;2-医疗卡
    strSQL = strSQL & "2,"
    '  卡类别id_In 病人预交记录.卡类别id%Type,
    strSQL = strSQL & "" & varData(1) & ","
    '  卡号_In     病人预交记录.卡号%Type
    strSQL = strSQL & "'" & varData(2) & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    Set cllUpdate = New Collection: Set cllThreeSwap = New Collection
    If CallBackBalanceInterface(varData(1), False, varData(2), _
        varData(3), varData(4), str结帐IDs, str冲销IDs, varData(5), cllUpdate, cllThreeSwap, strErrMsg) = False Then
        gcnOracle.RollbackTrans: blnCommited = True
        If strErrMsg <> "" Then
            MsgBox strErrMsg, vbExclamation, gstrSysName
        Else
            MsgBox "三方退费交易调用失败！退费操作失败！", vbExclamation, gstrSysName
        End If
        Exit Function
    End If
    ' zlExecuteProcedureArrAy cllUpdate, Me.Caption
    gcnOracle.CommitTrans: blnCommited = True
    On Error GoTo Errhand:
    zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
    
    DelThreeSwapFeeSingle = True
    Exit Function
errHandle:
    If Not blnCommited Then gcnOracle.RollbackTrans
    Call ErrCenter
    MsgBox "三方退费交易调用失败！", vbExclamation, gstrSysName
    blnCommited = True
    Exit Function
Errhand:
    If Not blnCommited Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        gcnOracle.BeginTrans: blnCommited = False
    End If
End Function

Private Function DelThreeSwapFee(ByVal varNO As Variant, _
     ByRef blnCommited As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:退第三方接口交易,退费成功,返回true,否则返回False
    '返回:三方交易退费成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-09-01 02:45:49
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCardNo As String, strSwap As String, strHsptCode As String '医院编码
    Dim varData As Variant, cllBlance As Collection, dblMoney As Double
    Dim i As Long, strNos As String, strSQL As String
    Dim strSelNos As String, str结帐ID As String, strErrMsg As String
    Dim cllUpdate As Collection, cllThreeSwap As Collection
    Dim dbl退款合计 As Double
    Dim bln退指定金额  As Boolean   '退
     
    On Error GoTo errHandle
    Dim blnHaveData As Boolean
    
    DelThreeSwapFee = False
    mrsBalance.Filter = "性质>=3 and 性质<=4 and 校对标志<>2"
    If mrsBalance.RecordCount = 0 Then
        mrsBalance.Filter = 0
        DelThreeSwapFee = True: Exit Function
    End If
    '退现
    bln退指定金额 = False
    Set cllBlance = New Collection
     If cbo退款方式.Enabled And cbo退款方式.Visible Then
        mrsBalance.Filter = "性质>=3 and  性质<=4 and 结算方式='" & zlStr.NeedName(cbo退款方式.Text) & "'"
        If mrsBalance.RecordCount <> 0 Then
            With mrsBalance
                bln退指定金额 = True
                varData = Array(0, 0, "", "", "", 0, "")
                varData(0) = Val(Nvl(!性质)): varData(1) = Val(Nvl(!卡类别ID))
                varData(2) = Nvl(!卡号): varData(3) = Nvl(!交易流水号)
                varData(4) = Nvl(!交易说明): varData(6) = Nvl(!结算方式)
                varData(5) = Val(txt退款金额.Text)
                cllBlance.Add varData
            End With
        Else
            '退给指定的结算方式
            mrsBalance.Filter = "性质>=3 and  性质<=4  "
            If mrsBalance.RecordCount = 0 Then
                mrsBalance.Filter = 0: DelThreeSwapFee = True: Exit Function
            End If
        End If
    End If
     
    mrsBalance.Filter = "性质=3 or 性质=4"
    dbl退款合计 = 0
    For i = 0 To UBound(varNO)
        strNos = strNos & "," & varNO(i)
    Next
    If strNos <> "" Then strNos = Mid(strNos, 2)
    
    With mrsBalance
        .Sort = "性质,卡类别ID,卡号,交易流水号,交易说明"
        .MoveFirst
        varData = Array(0, 0, "", "", "", 0, "")
        dblMoney = 0
        If .RecordCount <> 0 Then .MoveFirst
        
        Do While Not .EOF
            If InStr(1, "," & strNos & ",", "," & Nvl(!NO) & ",") > 0 Then
                If Not bln退指定金额 Then
                        If Not (varData(0) = Val(Nvl(!性质)) And varData(1) = Val(Nvl(!卡类别ID)) _
                            And varData(2) = Nvl(!卡号) _
                            And varData(3) = Nvl(!交易流水号) _
                            And varData(4) = Nvl(!交易说明)) Then
                            
                            If varData(0) <> 0 Then
                               ' varData = Array(Val(Nvl(!性质)), Val(Nvl(!卡类别ID)), _
                                CStr(Nvl(!卡号)), CStr(Nvl(!交易流水号)), CStr(Nvl(!交易说明)), dblMoney)
                                blnHaveData = False
                                For i = 1 To vsBalance.COLS - 1 Step 2
                                    If vsBalance.Cell(flexcpData, 1, i) = Nvl(!结算方式) _
                                        And vsBalance.RowHidden(1) = False _
                                        And Val(vsBalance.TextMatrix(1, i + 1)) <> 0 Then
                                        blnHaveData = True: Exit For
                                    End If
                                Next
                                If blnHaveData Then cllBlance.Add varData
                            End If
                            varData(5) = 0
                            varData(0) = Val(Nvl(!性质)): varData(1) = Val(Nvl(!卡类别ID))
                            varData(2) = Nvl(!卡号): varData(3) = Nvl(!交易流水号)
                            varData(4) = Nvl(!交易说明): varData(6) = Nvl(!结算方式)
                        End If
                        dblMoney = dblMoney + Nvl(!结算金额)
                        varData(5) = varData(5) + Nvl(!结算金额)
                End If
                If InStr(1, "," & strSelNos & ",", "," & Nvl(!NO) & ",") = 0 Then
                        strSelNos = strSelNos & "," & Nvl(!NO)
                        str结帐ID = str结帐ID & "," & Nvl(!结帐ID)
                End If
            End If
            .MoveNext
        Loop
    End With
    If varData(0) <> 0 And Not bln退指定金额 Then
        blnHaveData = False
        For i = 1 To vsBalance.COLS - 1 Step 2
            If vsBalance.Cell(flexcpData, 1, i) = varData(6) _
               And vsBalance.RowHidden(1) = False And Val(vsBalance.TextMatrix(1, i + 1)) <> 0 Then
                blnHaveData = True: Exit For
            End If
        Next
        If blnHaveData Then cllBlance.Add varData
    End If
    
    mrsBalance.Filter = 0
    If strSelNos = "" Or cllBlance.Count = 0 Then DelThreeSwapFee = True: Exit Function
    strSelNos = Mid(strSelNos, 2)
    If str结帐ID = "" Then str结帐ID = ",0"
    str结帐ID = Mid(str结帐ID, 2)
    For i = 1 To cllBlance.Count
      'varData = Array(Val(Nvl(!性质)), Val(Nvl(!卡类别ID)), _
                        CStr(Nvl(!卡号)), CStr(Nvl(!交易流水号)), CStr(Nvl(!交易说明)), dblMoney)
        ' Zl_门诊收费_完成校对
        strSQL = "Zl_门诊收费_完成校对("
        '  No_In       Varchar2,
        strSQL = strSQL & "'" & strSelNos & "',"
        '  操作类型_In Number,
        '  --操作类型_In:0-一卡通;1-消费卡;2-医疗卡
        strSQL = strSQL & "" & IIf(cllBlance(i)(0) = 3, 2, 1) & ","
        '  卡类别id_In 病人预交记录.卡类别id%Type,
        strSQL = strSQL & "" & cllBlance(i)(1) & ","
        '  卡号_In     病人预交记录.卡号%Type
        strSQL = strSQL & "'" & cllBlance(i)(2) & "')"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
        Set cllUpdate = New Collection: Set cllThreeSwap = New Collection
        If CallBackBalanceInterface(cllBlance(i)(1), cllBlance(i)(0) = 4, cllBlance(i)(2), _
            cllBlance(i)(3), cllBlance(i)(4), str结帐ID, "", cllBlance(i)(5), cllUpdate, cllThreeSwap, strErrMsg) = False Then
            gcnOracle.RollbackTrans: blnCommited = True
            If strErrMsg <> "" Then
                    MsgBox strErrMsg, vbExclamation, gstrSysName
            Else
                   MsgBox "三方退费交易调用失败！退费操作失败！", vbExclamation, gstrSysName
            End If
            Exit Function
        End If
       ' zlExecuteProcedureArrAy cllUpdate, Me.Caption
        gcnOracle.CommitTrans: blnCommited = True
        On Error GoTo Errhand:
        zlExecuteProcedureArrAy cllThreeSwap, Me.Caption
        gcnOracle.BeginTrans: blnCommited = False
    Next
    gcnOracle.CommitTrans: blnCommited = True
    DelThreeSwapFee = True
    Exit Function
errHandle:
    If Not blnCommited Then gcnOracle.RollbackTrans
    Call ErrCenter
    MsgBox "三方退费交易调用失败！", vbExclamation, gstrSysName
      blnCommited = True
    Exit Function
Errhand:
     If Not blnCommited Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        gcnOracle.BeginTrans: blnCommited = False
        Resume
    End If
    
End Function
Private Function ExecDelete() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行多单据退费
    '返回:退费成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-30 18:52:56
    '说明:因为医保的原因，多张单据退费时，分次提交
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim colOrder As New Collection, colBalanceID As New Collection
    Dim colBalance As New Collection    '保险结算方式,金额
    Dim colThreeBalance As New Collection    '三方结算,金额
    Dim colOtherBalance As New Collection    '其他结算,金额
    Dim colSQL As New Collection
    Dim arrSQL As Variant, strSQL As String, strInvoices As String, strInvoice As String
    Dim blnCur部份退费 As Boolean, blnAll部份退费 As Boolean, blnTrans As Boolean, blnTransMedicare As Boolean, blnPrint As Boolean
    Dim strBalance As String, strAllBalance As String, strTmp As String, strAdvance As String
    Dim strNo As String, str序号 As String, strDelNOs As String, strOtherNOs As String, strAllNOs As String
    Dim cur误差金额 As Currency, DateDel As Date
    Dim i As Long, j As Long, k As Long, lngCount As Long, arrNo As Variant, lng领用ID As Long
    Dim strThreeSwapBanace As String '三方交易
    Dim objICCard As Object, strCardNo As String, rsOneCard As ADODB.Recordset
    Dim colOneCard As New Collection, blnTransOneCard As Boolean
    Dim str医保结算 As String, rsTmp As ADODB.Recordset
    Dim arrBalance() As String, str结算方式 As String, lng结帐ID As Long
    Dim cur可分配额 As Currency, cur结算金额 As Currency, cur余额 As Currency, cur退款合计 As Currency
    Dim strCurDelNOs As String '用逗号分离,如:'J0002','J00023'
    Dim blnRllTrans As Boolean  '是否回退
    Dim strCurSelNos As String '当前选中的单号
    Dim str退结算方式 As String, bln退现 As Boolean
    Dim blnThreeSwapComit As Boolean
    Dim lng冲销ID As Long, lng结算序号 As Long  '43395
    Dim strThreeBalance As String, intCol As Integer
    Dim strOtherBalance As String '其他退费方式
    Dim blnNotFind As Boolean, str冲销IDs As String
    Dim blnExistThreeSwap As Boolean, blnExistOneCardSwap As Boolean, bln全退 As Boolean
    Dim blnYbComit As Boolean, blnCommited As Boolean, blnOneCardComit As Boolean
    Dim varTemp As Variant, strReclaimInvoice As String, intInvoiceFormat As Integer '回收票据:56963
    Dim cll退费结帐ID As Collection, str成功退费ID As String
    Dim strCmdCaptions As String, bln药品 As Boolean, blnSel As Boolean
    Dim strYPNos As String, strPrintNOInfor As String  '当前打印的单据信息:NO:序号;
    Dim strReturn As String, strReturnRecipt As String '退费处方信息，格式：NO,药房ID|NO,药房ID|…
    Dim dblDelMoney As Double, bln完全退费 As Boolean
    
    str冲销IDs = ""
    '检查输入是否正确
    If mstrNOs = "" Then
        MsgBox "请先输入要退费的单据。", vbInformation, gstrSysName
        If txtNO.Visible Then txtNO.SetFocus: Exit Function
    End If
    If CheckBillExistReplenishData(1, , mstrNOs) Then
        MsgBox "选择的退费记录进行了医保补充结算，不允许进行退费操作！", vbInformation, gstrSysName
        Exit Function
    End If
    arrNo = Split(mstrNOs, ",")
    strYPNos = ""
    bln药品 = False: blnSel = False
    For i = 1 To vsBill.Rows - 1
        If Val(vsBill.TextMatrix(i, vsBill.ColIndex("选择"))) <> 0 Then
            blnSel = True
            If vsBill.ColIndex("类别") <> -1 Then     '47400
                If vsBill.TextMatrix(i, vsBill.ColIndex("类别")) Like "*西*药*" _
                    Or vsBill.TextMatrix(i, vsBill.ColIndex("类别")) Like "*中*药*" _
                    Or vsBill.TextMatrix(i, vsBill.ColIndex("类别")) Like "*卫材*" Then
                    strYPNos = strYPNos & "," & vsBill.TextMatrix(i, vsBill.ColIndex("单据号"))
                    bln药品 = True
                    '81190,冉俊明,退费业务向发药机上传退费信息
                    '格式：NO,药房ID|NO,药房ID|…
                    If Not vsBill.TextMatrix(i, vsBill.ColIndex("类别")) Like "*卫材*" Then
                        If InStr(strReturnRecipt & "|", _
                            "|" & vsBill.TextMatrix(i, vsBill.ColIndex("单据号")) & "," & vsBill.TextMatrix(i, vsBill.ColIndex("执行科室ID")) & "|") = 0 Then
                            strReturnRecipt = strReturnRecipt & "|" & vsBill.TextMatrix(i, vsBill.ColIndex("单据号")) & "," & vsBill.TextMatrix(i, vsBill.ColIndex("执行科室ID"))
                        End If
                    End If
                End If
            End If
        End If
    Next
    
    If blnSel = False Then
        MsgBox "请在单据中至少选择一个要退费的项目。", vbInformation, gstrSysName
        vsBill.SetFocus: Exit Function
    End If
    
    '47400
    If bln药品 Then
        If strYPNos <> "" Then strYPNos = Mid(strYPNos, 2)
        If zlCheckDrugIsPutDrug(strYPNos) = False Then Exit Function
    End If
    
    '刘兴洪:28947
    If mintInsure <> 0 Then
        If gclsInsure.CheckInsureValid(mintInsure) = False Then
            Exit Function
        End If
    End If
    '先判断所有单据是否部份退费,以决定票据的处理方式
    blnAll部份退费 = False
    
    '用于判断仅使用医疗卡结算时，是否为部分退
    Dim strCurNO As String
    For j = 1 To vsBill.Rows - 1
        If Val(vsBill.TextMatrix(j, vsBill.ColIndex("选择"))) = 0 Then bln完全退费 = False: Exit For
        If strCurNO = "" Or strCurNO <> vsBill.TextMatrix(j, vsBill.ColIndex("单据号")) Then
            strCurNO = vsBill.TextMatrix(j, vsBill.ColIndex("单据号"))
            bln完全退费 = BillDeleteAll(strCurNO, 1, mblnHaveExcuteData)
            If bln完全退费 Then bln完全退费 = Not BillExistDelete(strCurNO, 1)
            If bln完全退费 = False Then Exit For
        End If
    Next
    
    '一起收费的其它单据:mstrNOs只是可以退的,不是所有的
    strAllNOs = GetMultiNOs(CStr(arrNo(0)), , , mCurBillType.bln按结算序号退)
    
    strOtherNOs = strAllNOs
    If zlCheckIsMzToZY(mstrNOs, 1) Then
          MsgBox "注意:" & vbCrLf & _
            "    该单据已经被门诊费用转住院费用 " & vbCrLf & _
            "    或已经审核了门诊费用转住院费用,不能再退费", vbInformation + vbOKOnly, gstrSysName
          Exit Function
    End If
      
    strCurSelNos = ""
    
    For i = 0 To UBound(arrNo)
        strNo = arrNo(i)
        str序号 = "": strBalance = "": lngCount = 0: strThreeBalance = ""
        dblDelMoney = 0
                       
        '收集当前单据要退费的行号
        With vsBill
            k = vsBill.FindRow(strNo, , .ColIndex("单据号"))
            For j = k To vsBill.Rows - 1
                If vsBill.TextMatrix(j, .ColIndex("单据号")) <> strNo Then Exit For
                If Val(vsBill.TextMatrix(j, .ColIndex("选择"))) <> 0 Then
                    str序号 = str序号 & "," & CLng(vsBill.RowData(j))
                    If InStr(1, strCurSelNos & ",", "," & strNo & ",") = 0 Then
                        strCurSelNos = strCurSelNos & "," & strNo
                    End If
                    dblDelMoney = dblDelMoney + Val(vsBill.TextMatrix(j, .ColIndex("实收金额")))
                End If
                lngCount = lngCount + 1
            Next
        End With
        str序号 = Mid(str序号, 2)
        
        If str序号 <> "" Then
            blnCur部份退费 = Not BillDeleteAll(strNo, 1, mblnHaveExcuteData)
            strPrintNOInfor = strPrintNOInfor & ";" & strNo & ":" & str序号
            
            If UBound(Split(str序号, ",")) + 1 = lngCount And blnCur部份退费 = False Then str序号 = ""
            If mintInsure <> 0 Then
                strAllBalance = Get医保结算方式(strNo)
                For j = 0 To UBound(Split(strAllBalance, ","))
                    strTmp = Split(strAllBalance, ",")(j)
                    If Not mblnYB结算作废 Then
                          '不支持门诊结算作废时,只允许个帐退为现金,其它原样退,不调用医保交易
                        If strTmp = mstr个人帐户 Then strBalance = "," & strTmp
                    End If
                    If mblnYB结算作废 Then
                        If Not gclsInsure.GetCapability(support门诊结算作废, mlng病人ID, mintInsure, strTmp) Then
                            strBalance = strBalance & "," & strTmp
                        End If
                    End If
                Next
            End If
            '医保部份退费检查
            'Or BillExistDelete(strNO, 1):如果多次退费,最后一次退费的没有发票收回,并提示重打,也没重打票据,所以不应该加:Or BillExistDelete(strNO, 1)
            blnCur部份退费 = Not (Not blnCur部份退费 And str序号 = "")
            If blnCur部份退费 Then blnAll部份退费 = True '这张单据为部份退费,则所有单据为部份退费
           '三方交易
            If mCurBillType.blnSingleBalance And mCurBillType.bln存在医疗卡结算 And Not bln完全退费 Then
                If vsBalance.RowHidden(1) = False And Val(vsBalance.TextMatrix(1, 2)) <> 0 Then '不退现
                    strThreeBalance = strThreeBalance & "," & vsBalance.Cell(flexcpData, 1, 1) & "|" & dblDelMoney
                End If
            Else
                mrsBalance.Filter = "NO='" & strNo & "' and 性质>=3"
                With mrsBalance
                     Do While Not .EOF
                         If blnCur部份退费 Then
                            strThreeBalance = strThreeBalance & "," & Nvl(!结算方式)
                        End If
                        If Val(Nvl(!是否退现)) = 1 Then
                            '排队交现金
                            For intCol = 1 To vsBalance.COLS - 1 Step 2
                                If vsBalance.Cell(flexcpData, 1, intCol) = !结算方式 And _
                                    Val(vsBalance.TextMatrix(1, intCol + 1)) = 0 Then
                                    strThreeBalance = strThreeBalance & "," & Nvl(!结算方式)
                                End If
                            Next
                        End If
                        .MoveNext
                    Loop
               End With
               
                strOtherBalance = ""
               '其他结算方式:部分退费
                mrsBalance.Filter = "NO='" & strNo & "' "
                With mrsBalance
                     Do While Not .EOF
                         If InStr(",1,2,", "," & Val(Nvl(!结算性质)) & ",") > 0 Then
                             blnNotFind = True
                             For intCol = 1 To vsBalance.COLS - 1 Step 2
                                 If vsBalance.Cell(flexcpData, 1, intCol) = !结算方式 And vsBalance.TextMatrix(1, intCol) <> "" Then
                                     blnNotFind = False: Exit For
                                 End If
                             Next
                             If blnNotFind Then
                                 strOtherBalance = strOtherBalance & "," & Nvl(!结算方式)
                             End If
                         End If
                         .MoveNext
                     Loop
                End With
            End If
            colOrder.Add str序号, "_" & strNo
            lng结帐ID = Val(vsBill.TextMatrix(k, vsBill.ColIndex("结帐ID")))
        
        '一卡通检查
        If Not CheckOnCardValied(blnCur部份退费, lng结帐ID) Then Exit Function
        '三方交易检查
            If mCurBillType.blnSingleBalance And mCurBillType.bln存在医疗卡结算 And Not bln完全退费 Then
                If mCurBillType.bln三方卡全退 Then
                    If Val(vsBalance.ColData(2)) = 0 Then '不退现
                        MsgBox "当前单据使用了第三方结算交易，所有单据必须全退！", vbInformation, gstrSysName
                        Exit Function
                    ElseIf cbo退款方式.Visible = False Then
                        MsgBox "当前单据使用了第三方结算交易，所有单据必须全退，或者你可以选择退为其它结算方式！", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
                
                If strThreeBalance <> "" Then
                    mrsBalance.Filter = "结帐ID=" & lng结帐ID & " And 性质=3"
                    If mrsBalance.RecordCount = 0 Then
                        MsgBox "当前单据 " & strNo & " 使用了第三方结算交易，但未发现原始结算数据，请核查！", vbInformation, gstrSysName
                        Exit Function
                    End If
                    With mrsBalance
                        If .RecordCount <> 0 Then
                            .MoveFirst
                            If zlCheckDelValied(Val(Nvl(!卡类别ID)), Nvl(!名称), False, Nvl(!卡号), Nvl(!交易流水号), _
                                Nvl(!交易说明), lng结帐ID, dblDelMoney) = False Then Exit Function
                        End If
                    End With
                End If
            Else
                If Not CheckThreeSwapValied(blnCur部份退费, lng结帐ID, InStr(1, mstrNOs, ",") > 0) Then Exit Function
            End If
            If mintInsure <> 0 And blnCur部份退费 Then      '医保支持退费时,每一张要求全退
                If str序号 <> "" Then
                    MsgBox "单据""" & strNo & """包含保险结算费用，而其中一些项目可能已经执行，不允许部份退费。", vbInformation, gstrSysName
                Else
                    MsgBox "单据""" & strNo & """包含保险结算费用，不允许部份退费。", vbInformation, gstrSysName
                End If
                vsBill.SetFocus: Exit Function
            End If
            
            '判断本次是否退完时，排开这张单据
            strOtherNOs = Mid(Replace("," & strOtherNOs, ",'" & strNo & "'", ""), 2)
        Else
            blnAll部份退费 = True                       '这张单据不退费,则所有单据为部份退费
            colOrder.Add "未选择", "_" & strNo
        End If
        
        '医保不允许退的结算方式,非医保时为空
        If strBalance <> "" Then strBalance = Mid(strBalance, 2)
        If strThreeBalance <> "" Then strThreeBalance = Mid(strThreeBalance, 2)
        If strOtherBalance <> "" Then strOtherBalance = Mid(strOtherBalance, 2)
        colBalance.Add strBalance, "_" & strNo
        colThreeBalance.Add strThreeBalance, "_" & strNo
        colOtherBalance.Add strOtherBalance, "_" & strNo
        
        '医保退费结算要用的结帐ID,非医保时为0,不支持作废的这种医保,不调用医保交易
        If mblnYB结算作废 And mintInsure <> 0 Then
            colBalanceID.Add Val(vsBill.TextMatrix(k, vsBill.ColIndex("结帐ID"))), "_" & strNo
        Else
            colBalanceID.Add 0, "_" & strNo
        End If
    Next
    
    '根据其它单据是否未退完,则可判断出所有单据是否部份退费
    If (Not blnAll部份退费) And strOtherNOs <> "" Then
        If BillExistMoney(strOtherNOs, 1) Then blnAll部份退费 = True
    End If
    
    '预交相关检查和验证
    If strCurSelNos <> "" Then
        strCurSelNos = Mid(strCurSelNos, 2)
        If Not zlCheckPrepayBack(mlng病人ID, strCurSelNos) Then
            Exit Function
        End If
    End If
    
    If blnAll部份退费 Then
        '56963
        If gTy_Module_Para.byt票据分配规则 <> 0 And strReclaimInvoice = "" Then
            strReclaimInvoice = zlGetReclaimInvoice(Mid(strPrintNOInfor, 2))
        End If
        If Not (gTy_Module_Para.byt票据分配规则 <> 0 And strReclaimInvoice <> "") Then
                If InStr(mstrPrivs, "部份退费") = 0 Then
                    MsgBox "你没有权限执行部份退费操作！", vbInformation, gstrSysName
                    vsBill.SetFocus: Exit Function
                End If
                If gTy_Module_Para.bln工本费 Then
                    MsgBox "自动收取工本费时不允许部份退费。", vbInformation, gstrSysName: vsBill.SetFocus: Exit Function
                End If
                
                '刘兴洪 问题:27352 日期:2010-01-13 10:26:08
                If InStr(1, mstrPrivs, "退费核收发票") > 0 Then
                    If frmReInvoice.ShowMe(Me, strNo, Val(txtAllTotal.Text), Val(txt退款金额.Text), strInvoices) = False Then
                        vsBill.SetFocus: Exit Function
                    End If
                End If
        End If
    End If
    
    If mBillDelType = EM_多张全退 And blnAll部份退费 Then
        MsgBox "多张单据使用一卡通结算模式或医保退费要求整体退，不允许部分退费！", vbInformation, gstrSysName
        vsBill.SetFocus: Exit Function
    End If
                      
    If mintInsure <> 0 And MCPAR.医保接口打印票据 Then
        If zlGetInvoiceGroupUseID(lng领用ID) = False Then Exit Function
        strInvoice = GetNextBill(lng领用ID)
    End If
    
    
    DateDel = zlDatabase.Currentdate
    Set cll退费结帐ID = New Collection
    
    '生成要执行的SQL
    lng结算序号 = 0
    '医保要处理为倒序,因此,按最后一张先冲销
    For i = UBound(arrNo) To 0 Step -1
        arrSQL = Array(): strNo = arrNo(i)
        If colOrder("_" & strNo) <> "未选择" Then
            cur误差金额 = Val(mcolError("_" & strNo))
           '60974
            If mintInsure <> 0 And colBalance("_" & strNo) = "" And MCPAR.多单据收费必须全退 Then cur误差金额 = 0    '医保本张全退且结算全支持作废时无误差
            lng冲销ID = zlDatabase.GetNextId("病人结帐记录")
            str冲销IDs = str冲销IDs & "," & lng冲销ID
            If lng结算序号 = 0 Then lng结算序号 = lng冲销ID
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            
            'Zl_门诊收费记录_Delete
            strSQL = "zl_门诊收费记录_DELETE("
            '  No_In           门诊费用记录.NO%Type,
            strSQL = strSQL & "'" & strNo & "',"
            '  操作员编号_In   门诊费用记录.操作员编号%Type,
            strSQL = strSQL & "'" & UserInfo.编号 & "',"
            '  操作员姓名_In   门诊费用记录.操作员姓名%Type,
            strSQL = strSQL & "'" & UserInfo.姓名 & "',"
            '  医保结算方式_In Varchar2 := Null,
            strSQL = strSQL & "'" & colBalance("_" & strNo) & "',"
            '  序号_In         Varchar2 := Null,
            strSQL = strSQL & "'" & colOrder("_" & strNo) & "',"
            '  结算方式_In     病人预交记录.结算方式%Type := Null,
            strSQL = strSQL & "'" & zlStr.NeedName(cbo退款方式.Text) & "',"
            '  误差_In         门诊费用记录.实收金额%Type := 0,
            strSQL = strSQL & "" & cur误差金额 & ","
            '  退费时间_In     门诊费用记录.登记时间%Type := Null,
            strSQL = strSQL & "To_Date('" & Format(DateDel, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
            '  回收票据_In     Number := 0,
            strSQL = strSQL & "" & IIf(blnAll部份退费, "1", "0") & ","
            '  退费摘要_In     门诊费用记录.摘要%Type := Null
            strSQL = strSQL & "" & IIf(Trim(txt退费摘要.Text) = "", "NULL", "'" & Trim(txt退费摘要.Text) & "'") & ","
            '     校对标志_In: 0-不需要较对;1-需较对(不处理人员缴款余额,不回收票据)
            strSQL = strSQL & "1,"
            '  结帐id_In       病人预交记录.结帐id%Type := Null,
            strSQL = strSQL & lng冲销ID & ","
            '  结算序号_In     病人预交记录.结算序号%Type := Null
            strSQL = strSQL & lng结算序号 & ","
              '一卡通结算_In   Varchar2 := Null
             strOtherBalance = colOtherBalance("_" & strNo)
            'If Not blnAll部份退费 Then strOtherBalance = ""
             strSQL = strSQL & "'" & colThreeBalance("_" & strNo) & _
                IIf(colThreeBalance("_" & strNo) <> "" And strOtherBalance <> "", ",", "") & strOtherBalance & "',"
             '退款操作_In:1-进行部分退(将退款方式退到指定的结算方式<结算方式_In>中,0-不指定退款方式)
             If (blnAll部份退费 Or mCurBillType.bln多张部分退费) And mintInsure = 0 Then
                '普通病人
                '检查是否退到指定的结算方式<结算方式_In>中
                blnNotFind = True
                For intCol = 1 To vsBalance.COLS - 1 Step 2
                    If Val(vsBalance.TextMatrix(1, intCol + 1)) <> 0 Then blnNotFind = False: Exit For
                Next
                strSQL = strSQL & IIf(cbo退款方式.Visible And (vsBalance.RowHidden(1) Or blnNotFind), "1", "0") & ","
             Else
                strSQL = strSQL & "0,"
             End If
             '多单据全退_IN=1-多单据全退(多张单据全退,原样退);0-非原样退:60974
              'strSQL = strSQL & IIf(Not vsBalance.RowHidden(1), "1", "0") & ")"
              If mintInsure <> 0 And colBalance("_" & strNo) = "" And MCPAR.多单据收费必须全退 Then
                 strSQL = strSQL & "1)"
              Else
                 strSQL = strSQL & IIf(cbo退款方式.Visible Or blnAll部份退费 Or cur误差金额 <> 0, "0", "1") & ")"
              End If
            arrSQL(UBound(arrSQL)) = strSQL
            '60974
'            If cur误差金额 <> 0 Then
'                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
'                arrSQL(UBound(arrSQL)) = "zl_门诊收费误差_Insert('" & strNO & "'," & cur误差金额 & ",1)"
'            End If
            cll退费结帐ID.Add lng冲销ID, "_" & strNo
            strCurDelNOs = strCurDelNOs & ",'" & strNo & "'"
        End If
        colSQL.Add arrSQL, "_" & strNo '当前单据的SQL集
    Next
    
    bln退现 = False
    If cbo退款方式.ListIndex >= 0 Then
        bln退现 = cbo退款方式.ItemData(cbo退款方式.ListIndex) = 1
        str退结算方式 = zlStr.NeedName(cbo退款方式.Text)
    Else
        bln退现 = True
        str退结算方式 = IIf(mstr现金结算方式 = "", "现金", mstr现金结算方式)
    End If
    '56963
    If strPrintNOInfor <> "" Then strPrintNOInfor = Mid(strPrintNOInfor, 2)
    strReclaimInvoice = zlGetReclaimInvoice(strPrintNOInfor)
    
    If gTy_Module_Para.byt票据分配规则 <> 0 And strReclaimInvoice <> "" Then
        If InStr(1, mstrPrivs, "退费核收发票") > 0 Then
            If MsgBox("注意:" & vbCrLf & " 当前退费的单据中包含如下收费票据，是否回收这些票据?" & vbCrLf & strReclaimInvoice, vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
        End If
    End If
    If gblnBillPrint Then
        If gobjBillPrint.zlEraseBill(strAllNOs, 0) = False Then Exit Function
    End If
    
    '1-预存款,2-医保,3-医疗卡,4-结算卡,5-一卡通,0-其他类
    mrsBalance.Filter = "性质=3 or 性质=4"  '消费卡和银行卡
    blnExistThreeSwap = mrsBalance.RecordCount <> 0
    mrsBalance.Filter = "是否全退=1"
    bln全退 = mrsBalance.RecordCount <> 0
    mrsBalance.Filter = "性质=5"
    blnExistOneCardSwap = mrsBalance.RecordCount <> 0
    mrsBalance.Filter = 0
    
    '执行退费的SQL
    On Error GoTo errH
    strDelNOs = ""
    blnCommited = False: blnYbComit = False
    If mintInsure <> 0 And (MCPAR.多单据一次结算 Or MCPAR.多单据调一次交易) Then
        '多张单据医保一次结算
        gcnOracle.BeginTrans: blnTrans = True
        strAllBalance = "": strBalance = ""
        For i = 0 To UBound(arrNo)
            strNo = arrNo(i)          '从最后一张开始退
            For j = 0 To UBound(colSQL("_" & strNo))
                Call zlDatabase.ExecuteProcedure(CStr(colSQL("_" & strNo)(j)), Me.Caption)
            Next
            strAllBalance = IIf(strAllBalance = "", "", strAllBalance & ",") & colBalanceID("_" & strNo)
            If i = 0 Then strBalance = colBalanceID("_" & strNo)
        Next
        '先产生票据，医保接口才能取到
        If MCPAR.医保接口打印票据 _
            And Not (gTy_Module_Para.byt票据分配规则 <> 0 And strReclaimInvoice <> "") Then
            '56963
            strSQL = "zl_门诊收费记录_RePrint('" & strNo & "','" & strInvoice & "'," & ZVal(lng领用ID) & ",'" & UserInfo.姓名 & "'," & _
                "To_Date('" & Format(DateDel, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),1,1)"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        End If
        
        If DelInsureMulitOneBalance(blnExistThreeSwap, arrNo, Val(strBalance), strAllBalance, str医保结算, str退结算方式, bln退现, blnCommited) = False Then
            If Not blnCommited Then gcnOracle.RollbackTrans
            Exit Function
        End If
        If blnCommited = True Then
            blnYbComit = True: gcnOracle.BeginTrans: blnTrans = True
        End If
    Else
        '-------------------------------------------------------------------------------------------------------
        '刘兴洪:医保的strAdvancey计算:本次退费总张数|当前退费第几张:27231
        Dim lngPages As Long, lngPage, cllYB As Collection
        Set cllYB = New Collection
        lngPage = 0: lngPages = 0
        For i = UBound(arrNo) To 0 Step -1
            strNo = arrNo(i)
            If UBound(colSQL("_" & strNo)) >= 0 And mintInsure <> 0 Then
                '医保的
                 If mblnYB结算作废 And colBalanceID("_" & strNo) <> 0 Then
                    lngPage = lngPage + 1
                    lngPages = lngPages + 1
                    cllYB.Add lngPage, "_" & strNo
                 End If
            End If
        Next
        
        '-------------------------------------------------------------------------------------------------------
        '先处理医保
        If blnExistThreeSwap And bln全退 Then
               '先将所有单据退费,然后按医保单张退
               gcnOracle.BeginTrans '启动事务
               blnTrans = True
                For i = 0 To UBound(arrNo)
                    strNo = arrNo(UBound(arrNo) - i)        '从最后一张开始退
                    If UBound(colSQL("_" & strNo)) >= 0 Then
                        For j = 0 To UBound(colSQL("_" & strNo))
                            Call zlDatabase.ExecuteProcedure(CStr(colSQL("_" & strNo)(j)), Me.Caption)
                        Next
                    End If
                Next
                '分单据处理医保
                For i = 0 To UBound(arrNo)
                
                    strNo = arrNo(UBound(arrNo) - i)        '从最后一张开始退
                    If UBound(colSQL("_" & strNo)) >= 0 Then
                        '退医保
                        blnCommited = False
                        If mintInsure <> 0 And mblnYB结算作废 Then
                            lngPage = Val(cllYB("_" & strNo)): lng结帐ID = Val(colBalanceID("_" & strNo))
                            If Not DelInsureOneBill(str医保结算, blnExistThreeSwap, lng结帐ID, lngPage, lngPages, blnCommited) Then
                                If blnCommited = False Then gcnOracle.RollbackTrans
                                blnTrans = False
                                '显示退费成功单据提示
                                Call ShowErrBill(strDelNOs, strNo, 3): Exit Function
                            End If
                        End If
                        If blnCommited Then
                            blnTrans = False
                            gcnOracle.BeginTrans    '只有提交后，才能重新
                            blnYbComit = True: blnTrans = True '只要有一种，就会提交
                        End If
                        strDelNOs = strDelNOs & IIf(strDelNOs = "", "", ",") & strNo
                    End If
                Next
        Else
               str成功退费ID = ""
                gcnOracle.BeginTrans '启动事务
                blnTrans = True
                For i = 0 To UBound(arrNo)
                    strNo = arrNo(UBound(arrNo) - i)       '医保要求从最后一张开始退
                    If UBound(colSQL("_" & strNo)) >= 0 Then
                        For j = 0 To UBound(colSQL("_" & strNo))
                            Call zlDatabase.ExecuteProcedure(CStr(colSQL("_" & strNo)(j)), Me.Caption)
                        Next
                        '退医保
                        blnCommited = False
                        If mintInsure <> 0 And mblnYB结算作废 Then
                            lngPage = Val(cllYB("_" & strNo)): lng结帐ID = Val(colBalanceID("_" & strNo))
                            If Not DelInsureOneBill(str医保结算, blnExistThreeSwap, lng结帐ID, lngPage, lngPages, blnCommited) Then
                                If Not blnCommited Then gcnOracle.RollbackTrans: blnTrans = False
                                gcnOracle.BeginTrans:  blnTrans = True
                                If strDelNOs <> "" And blnExistThreeSwap And blnExistOneCardSwap Then
                                    varTemp = Split(strDelNOs, ",")
                                    If Not DelSawpSpecifyNOs(varTemp, blnExistThreeSwap, blnExistOneCardSwap, strNo, blnCommited) Then
                                        If Not blnCommited Then gcnOracle.RollbackTrans
                                        Exit Function
                                    End If
                                    If Not blnCommited Then gcnOracle.CommitTrans
                                    gcnOracle.BeginTrans:   blnTrans = True
                                End If
                                If strDelNOs <> "" Then
                                    If OverFeeDel(str成功退费ID, mtyPati.病人ID, blnCommited) = False Then
                                        If blnOneCardComit = False And blnYbComit = False And blnThreeSwapComit = False Then
                                            gcnOracle.RollbackTrans: Exit Function
                                        End If
                                        Exit Function
                                    End If
                                    If blnCommited Then blnTrans = False
                                End If
                                If blnTrans Then gcnOracle.RollbackTrans: blnTrans = True
                                '对成功部分进行完成收费
                                '显示退费成功单据提示
                                Call ShowErrBill(strDelNOs, strNo): Exit Function
                            End If
                        End If
                        If blnCommited Then
                            gcnOracle.BeginTrans    '只有提交后，才能重新
                            blnYbComit = True: blnTrans = True '只要有一种，就会提交
                        End If
                        strDelNOs = strDelNOs & IIf(strDelNOs = "", "", ",") & strNo
                        str成功退费ID = str成功退费ID & IIf(str成功退费ID = "", "", ",") & cll退费结帐ID("_" & strNo)
                    End If
                Next
            End If
    End If
    
    If Not blnTrans Then gcnOracle.BeginTrans: blnTrans = True
    If strDelNOs <> "" Then
        varTemp = Split(strDelNOs, ",")
    Else
        varTemp = arrNo
    End If
    '------------------------------------------------------------------------------------------
 
    '退一卡通
ReDOOneCard:
    blnCommited = False
    If Not DelOneCardPay(varTemp, blnCommited) Then
        If blnCommited = False Then gcnOracle.RollbackTrans: blnTrans = False
        If blnYbComit Then
            strCmdCaptions = "异常单据(&C)|表示不进行一卡通调用,数据将以异常形式体现,但必须在今后进行处理"
            strCmdCaptions = strCmdCaptions & ";重退(&R)|表示重新调用一卡通结算交易"
            If frmVerfyCodeInput.ShowMsg(Me, "单据[" & strDelNOs & "]已经退费成功,但一卡通交易失败,[异常单据]必须输入验证码,建议不进行异常单据保存", strCmdCaptions) = False Then
                 gcnOracle.BeginTrans: blnTrans = True
                GoTo ReDOOneCard:
            End If
        End If
        Call ClearFace(True, False)
        Exit Function
    End If
    
    If blnCommited Then
        blnOneCardComit = True: blnTrans = False
        gcnOracle.BeginTrans: blnTrans = True
    End If
    '------------------------------------------------------------------------------------------
    '退一卡通等的三方交易
ReDOThreeSwap:
    blnCommited = False
    If mCurBillType.blnSingleBalance And mCurBillType.bln存在医疗卡结算 And Not bln完全退费 Then
        If Not DelThreeSwapFeeSingle(varTemp, colThreeBalance, colOrder, str冲销IDs, blnCommited) Then
            If blnCommited = False Then gcnOracle.RollbackTrans: blnTrans = True
            
            If blnOneCardComit Or blnYbComit Then
                strCmdCaptions = "异常单据(&C)|表示不进行三方交易,数据将以异常形式体现,但必须在今后进行处理"
                strCmdCaptions = strCmdCaptions & ";重退(&R)|表示重新调用三方结算交易进行退费"
                If frmVerfyCodeInput.ShowMsg(Me, "单据[" & "4243;432432" & "]已经" & IIf(blnYbComit, "医保", "") & IIf(blnOneCardComit, IIf(blnYbComit, "及", "") & "一卡通", "") & "退费成功,但三方交易退费失败,[异常单据]必须输入验证码,建议不进行异常单据保存", strCmdCaptions) = False Then
                  If blnCommited Then gcnOracle.BeginTrans: blnTrans = True
                  GoTo ReDOThreeSwap:
                End If
            End If
            Call ClearFace(True, False)
            Exit Function
        End If
    Else
        If Not DelThreeSwapFee(varTemp, blnCommited) Then
            If blnCommited = False Then gcnOracle.RollbackTrans: blnTrans = True
            
            If blnOneCardComit Or blnYbComit Then
                strCmdCaptions = "异常单据(&C)|表示不进行三方交易,数据将以异常形式体现,但必须在今后进行处理"
                strCmdCaptions = strCmdCaptions & ";重退(&R)|表示重新调用三方结算交易进行退费"
                If frmVerfyCodeInput.ShowMsg(Me, "单据[" & "4243;432432" & "]已经" & IIf(blnYbComit, "医保", "") & IIf(blnOneCardComit, IIf(blnYbComit, "及", "") & "一卡通", "") & "退费成功,但三方交易退费失败,[异常单据]必须输入验证码,建议不进行异常单据保存", strCmdCaptions) = False Then
                  If blnCommited Then gcnOracle.BeginTrans: blnTrans = True
                  GoTo ReDOThreeSwap:
                End If
            End If
            Call ClearFace(True, False)
            Exit Function
        End If
    End If
    If blnCommited Then
        blnThreeSwapComit = True: blnTrans = False
        gcnOracle.BeginTrans: blnTrans = True
    End If
    '------------------------------------------------------------------------------------------
    '完成收费
    blnCommited = False
    If OverFeeDel(str冲销IDs, mtyPati.病人ID, blnCommited) = False Then
        If blnCommited = False Then
            gcnOracle.RollbackTrans: Exit Function
        End If
        Call ClearFace(True, False)
        Exit Function
    End If
    
    If Not blnCommited Then         '普通病人,现收,直接提交
        gcnOracle.CommitTrans: Exit Function
    End If
    
    '81190,冉俊明,退费业务向发药机上传退费信息
    On Error Resume Next
    If mblnDrugPacker Then
        If strReturnRecipt <> "" Then
            strReturnRecipt = Mid(strReturnRecipt, 2)
            Call mobjDrugPacker.DYEY_MZ_TransRecipeReturn(1, UserInfo.编号, UserInfo.姓名, strReturnRecipt, strReturn)
        End If
    End If
    Err.Clear: On Error GoTo errH
    
    '打印退费单据
    Call PrintDelBill(strAllNOs, strCurDelNOs, strNo, mtyPati.病人ID, DateDel, blnAll部份退费, strInvoices, strReclaimInvoice)
    ExecDelete = True
    Exit Function
errH:
    blnRllTrans = False
    If Err.Number <> 0 Then
        If blnTrans Then gcnOracle.RollbackTrans: blnRllTrans = True
        If ErrCenter() = 1 Then
            Resume
        End If
    End If
        
    If blnTrans Then
        If Not blnRllTrans Then gcnOracle.RollbackTrans: blnRllTrans = True
        '医保和HIS不是同一个事务,HIS事务失败,但医保可能已上传,所以需要调"取消交易"接口
        If blnTransMedicare Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, False, mintInsure)
        If blnTransOneCard Then MsgBox "一卡通退费交易调用失败！退费操作失败！", vbExclamation, gstrSysName
    End If
    
    If Err.Number <> 0 Then Call SaveErrLog
    '中断提示,不打印，重新退费后再打印或自己选择重打
    Call ShowErrBill(strDelNOs, strNo)
    Exit Function
ErrRquare:
    blnRllTrans = False
    If Err.Number <> 0 Then
        If blnTrans Then gcnOracle.RollbackTrans: blnRllTrans = True
        If ErrCenter() = 1 Then
            Resume
        End If
    End If
    If blnTrans Then
        If Not blnRllTrans Then gcnOracle.RollbackTrans: blnRllTrans = True
         MsgBox "结算卡退费交易调用失败！", vbExclamation, gstrSysName
    End If
    If Err.Number <> 0 Then Call SaveErrLog
    If txtNO.Visible Then txtNO.SetFocus
End Function

Private Sub PrintDelBill(ByVal strAllNOs As String, ByVal strCurDelNOs As String, _
    ByVal strNo As String, _
    ByVal lng病人ID As Long, _
    ByVal dtDateDel As Date, ByVal blnAll部分退费 As Boolean, _
    ByVal strInvoices As String, ByVal strReclaimInvoice As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打印相关票据
    '入参: strAllNOs-当前涉及的所有单据
    '       strCurDelNOs-当前退费的单据
    '       dtDateDel-退费日期
    '       strInvoices-选择的发票号(旧模式)
    '       strReclaimInvoice-回收的发票号
    '出参:
    '编制:刘兴洪
    '日期:2013-05-27 16:41:06
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intInvoiceFormat As Integer, blnPrint As Integer
    Dim str发票号 As String, int票据张数 As Integer
    Dim strSQL As String
    Dim strPriceGrade As String
    
    On Error GoTo errHandle
    If Not blnAll部分退费 Then
         '税控部件全退时收回处理(全退时，zl_门诊收费记录_DELETE中已收回票据)
        If Not gobjTax Is Nothing And gblnTax Then
            gstrTax = gobjTax.zlTaxOutErase(gcnOracle, strAllNOs)
            If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
        End If
        GoTo PrintList:
        Exit Sub
    End If
    
    '部分退费时收回并重打,包括单张部分退和退多张中的某几张
    If gTy_Module_Para.byt票据分配规则 <> 0 And strReclaimInvoice <> "" Then
        '按新票据分配规则打印
        '先预算,看票据是否充足
        str发票号 = strReclaimInvoice
        If zlExeCuteBillNoSplit(True, 4, mlng领用ID, strAllNOs, lng病人ID, "", dtDateDel, 1, str发票号, int票据张数) = False Then GoTo PrintList:
        If int票据张数 = 0 Then
            '只回收票据,但不打印
            str发票号 = strReclaimInvoice
            Call zlExeCuteBillNoSplit(False, 4, mlng领用ID, strAllNOs, lng病人ID, "", dtDateDel, 1, str发票号, int票据张数)
            GoTo PrintList:
        End If
        blnPrint = True
        ''0-不打印;1-自动打印;2-提示打印
        If mintInvoicePrint = 0 Then blnPrint = False   '自动打印
        If mintInvoicePrint = 2 Then
            If MsgBox("是否打印票据？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then blnPrint = False
        End If
        '重打收回发票
        If blnPrint Then
            intInvoiceFormat = IIf(strReclaimInvoice = "" And gTy_Module_Para.byt票据分配规则 <> 0, mintOldInvoiceFormat, mintInvoiceFormat)
            If gintPriceGradeStartType >= 2 Then
                strPriceGrade = GetPriceGradeFromNos(strAllNOs)
            Else
                strPriceGrade = gstr普通价格等级
            End If
            Call RePrintCharge(1, strAllNOs, Me, mlng领用ID, strReclaimInvoice, True, dtDateDel, _
                intInvoiceFormat, , , mlngShareUseID, mstrUseType, , strPriceGrade)
        End If
        GoTo PrintList:
        Exit Sub
    End If
    If strInvoices = "" Then 'a.收回并重新打印门诊收据
        blnPrint = True
        ''0-不打印;1-自动打印;2-提示打印
        If mintInvoicePrint = 0 Then blnPrint = False   '自动打印
        If mintInvoicePrint = 2 Then
            If MsgBox("是否打印票据？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then blnPrint = False
        End If
        
        If blnPrint Then
            intInvoiceFormat = IIf(strReclaimInvoice = "" And gTy_Module_Para.byt票据分配规则 <> 0, mintOldInvoiceFormat, mintInvoiceFormat)
            If gintPriceGradeStartType >= 2 Then
                strPriceGrade = GetPriceGradeFromNos(strAllNOs)
            Else
                strPriceGrade = gstr普通价格等级
            End If
            Call RePrintCharge(1, strAllNOs, Me, mlng领用ID, strReclaimInvoice, True, dtDateDel, _
            intInvoiceFormat, , , mlngShareUseID, mstrUseType, , strPriceGrade)
        End If
        GoTo PrintList:
        Exit Sub
    End If
    
    'b.收费或上一次退时没有打印票据
    If strInvoices <> "无可退票据" Then
        'c.只收回票据
        strSQL = "zl_门诊收费记录_RePrint('" & strNo & "',Null,0,'" & UserInfo.姓名 & "'," & _
            "To_Date('" & Format(dtDateDel, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),1,0,'" & strInvoices & "')"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    End If
PrintList:
    If blnAll部分退费 Then
        '打印费用清单
        If InStr(mstrPrivs, "打印清单") > 0 Then
            If gint收费清单 = 1 Then
                Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me, "NO=" & strAllNOs, "药品单位=" & IIf(gbln药房单位, 1, 0), 2)
            ElseIf gint收费清单 = 2 Then
                If MsgBox("要打印收费清单吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_3", Me, "NO=" & strAllNOs, "药品单位=" & IIf(gbln药房单位, 1, 0), 2)
                End If
            End If
        End If
    End If
    '79448,冉俊明,2014-11-10,打印回单时传入参数错误,传为了",'O0000678','O0000679'"，应该去掉第一个逗号","
    If strCurDelNOs <> "" Then strCurDelNOs = Mid(strCurDelNOs, 2)
    If mintInsure <> 0 And MCPAR.退费后打印回单 And InStr(1, mstrPrivs, "医保退费回单") > 0 Then
        '问题:35248
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_4", Me, "NO=" & strCurDelNOs, 2)
    End If
    If mint退费回单打印 = 1 Then
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_5", Me, "NO=" & strCurDelNOs, 2)
    ElseIf mint退费回单打印 = 2 Then
        If MsgBox("是否打印退费回单？", vbYesNo + vbQuestion, gstrSysName) = vbYes Then
            Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_5", Me, "NO=" & strCurDelNOs, 2)
        End If
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function DelSawpSpecifyNOs(ByVal varNO As Variant, _
    ByVal blnExistThreeSwap As Boolean, _
    ByVal blnExistOneCardSwap As Boolean, _
     Optional strNOLost As String, Optional ByRef blnCommited As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:退三方交易和一卡通的指定单据
    '入参:varTemp():单据集
    '出参:blnCommited-是否已经处理了事务
    '返回:退费成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-01-11 15:47:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strDelNOs As String, i As Long
    
    blnCommited = False
    If UBound(varNO) < 0 Then DelSawpSpecifyNOs = True: Exit Function
    If blnExistOneCardSwap = False And blnExistThreeSwap = False Then DelSawpSpecifyNOs = True:  Exit Function
    '存在一卡通
    If Not DelOneCardPay(varNO, blnCommited) Then
        '显示错误信息
        If Not blnCommited = False Then gcnOracle.RollbackTrans: blnCommited = True
        For i = 0 To UBound(varNO)
            strDelNOs = strDelNOs & "," & varNO(i)
        Next
        If strDelNOs <> "" Then strDelNOs = Mid(strDelNOs, 2)
        Call ShowErrBill(strDelNOs, strNOLost, 1)
        Exit Function
    End If
    blnCommited = False
    '第三方接口交易
    If Not DelThreeSwapFee(varNO, blnCommited) Then
        If blnCommited = False Then gcnOracle.RollbackTrans: blnCommited = True
        For i = 0 To UBound(varNO)
            strDelNOs = strDelNOs & "," & varNO(i)
        Next
        If strDelNOs <> "" Then strDelNOs = Mid(strDelNOs, 2)
        '显示错误信息
        Call ShowErrBill(strDelNOs, strNOLost, 1)
        Exit Function
    End If
    DelSawpSpecifyNOs = True
End Function

Private Function ShowErrBill(ByVal strDelSucceedNos As String, _
    ByVal strDelLost As String, Optional bytType As Byte = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:部分成功单据提示信息
    '入参:strDelSucceedNos-退费成功的单据;
    '       strDelLost-退费失败的单据
    '       bytType-0-医保失败;1-一卡通失败;2-第三方交易失败;3-医保退费成功,但第三方交易未进行
    '编制:刘兴洪
    '返回:重试返回true,否则返回False
    '日期:2012-01-11 13:58:53
    '说明:中断提示,不打印，重新退费后再打印或自己选择重打
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If strDelSucceedNos = "" Then Exit Function
    
    If bytType = 1 Then
        MsgBox "单据[" & strDelLost & "]退费失败。但是单据[" & strDelSucceedNos & "]已成功进行医保退费, " & vbCrLf & _
            "   但一卡通退费失败, ，请对执行失败的单据重新退费", vbExclamation, gstrSysName
        GoTo GoClear:
    ElseIf bytType = 2 Then
        MsgBox "单据[" & strDelLost & "]退费失败。但是单据[" & strDelSucceedNos & "]已成功进行医保退费, " & vbCrLf & _
            "   但三方接口交易退费失败, ，请对执行失败的单据重新退费", vbExclamation, gstrSysName
        GoTo GoClear:
    ElseIf bytType = 3 Then
        MsgBox "单据[" & strDelLost & "]退费失败。但是单据[" & strDelSucceedNos & "]已成功进行医保退费。" & vbCrLf & _
            "但三方交易还未进行退费，请重新退费！", vbExclamation, gstrSysName
    Else
        MsgBox "单据[" & strDelLost & "]退费失败。但是单据[" & strDelSucceedNos & "]已成功退费。" & vbCrLf & _
            "单据未打印，请对执行失败的单据重新退费！", vbInformation, gstrSysName
    End If
GoClear:
    Call ClearFace
    If txtNO.Visible Then txtNO.SetFocus
End Function

Public Function Get医保结算方式(ByVal strNo As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定单据的医保结算方式
    '返回:返回结算方式,用逗号分隔:个人帐户,医保基金...
    '编制:刘兴洪
    '日期:2011-08-30 18:54:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strBalance As String
    On Error GoTo errH
    With mrsBalance
         .Filter = "NO='" & strNo & "' and 性质=2"
         Do While Not .EOF
            strBalance = strBalance & "," & !结算方式
            .MoveNext
         Loop
    End With
    If strBalance <> "" Then strBalance = Mid(strBalance, 2)
    Get医保结算方式 = strBalance
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Public Function Get三方交易结算方式(ByVal strNo As String) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取指定单据的三方交易的相关结算方式
    '返回:返回结算方式,用逗号分隔:建行,一卡通...
    '编制:刘兴洪
    '日期:2011-08-30 18:54:51
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strBalance As String
    On Error GoTo errH
    With mrsBalance
         .Filter = "NO='" & strNo & "' and 性质>=5"
         Do While Not .EOF
            strBalance = strBalance & "," & !结算方式
            .MoveNext
         Loop
    End With
    If strBalance <> "" Then strBalance = Mid(strBalance, 2)
    Get三方交易结算方式 = strBalance
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Public Function Get实收金额(ByVal strNo As String) As Currency
    Dim i As Long
    With vsBill
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("单据号")) = strNo Then Get实收金额 = Get实收金额 + Val(.TextMatrix(i, .ColIndex("实收金额")))
        Next
    End With
End Function
Private Sub txt退费摘要_KeyDown(KeyCode As Integer, Shift As Integer)
    '选择退费原因
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If Trim(txt退费摘要.Tag) <> "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If Trim(txt退费摘要.Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    If zl_SelectAndNotAddItem(Me, txt退费摘要, Trim(txt退费摘要.Text), "常用退费原因", "常用退费原因选择", True, True) = False Then
        If zlCommFun.IsCharChinese(Trim(txt退费摘要.Text)) = False Then
            Exit Sub
        Else
            zlCommFun.PressKey vbKeyTab
        End If
    End If
End Sub
Private Sub txt退费摘要_GotFocus()
    zlCommFun.OpenIme True
    zlControl.TxtSelAll txt退费摘要
End Sub
Private Sub txt退费摘要_LostFocus()
    zlCommFun.OpenIme False
End Sub

Private Sub txt退费摘要_Change()
    txt退费摘要.Tag = ""
End Sub
Private Sub initCardSquareData()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:创建或关闭结算卡对象
    '编制:刘兴洪
    '日期:2010-01-05 14:51:23
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strExpend As String
    Set mobjSquare = gobjSquare.objSquareCard
    If mbytMode = 0 Then Exit Sub
    If gobjSquare.objSquareCard Is Nothing Then
        '创建对象
        Call CreateSquareCardObject(gfrmMain, mlngModule)
    End If
    If gobjSquare.objSquareCard Is Nothing Then Exit Sub
    Call IDKind.zlInit(Me, glngSys, glngModul, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, "", txtPatient)
      
    Dim objCard As Card
    Set objCard = IDKind.GetfaultCard
    Set gobjSquare.objDefaultCard = objCard
    If IDKind.Cards.按缺省卡查找 And Not objCard Is Nothing Then
        gobjSquare.bln缺省卡号密文 = objCard.卡号密文规则 <> ""
        gobjSquare.int缺省卡号长度 = objCard.卡号长度
    Else
        gobjSquare.bln缺省卡号密文 = IDKind.Cards.加密显示
        gobjSquare.int缺省卡号长度 = 100
    End If
    Set mobjSquare = gobjSquare.objSquareCard
End Sub


Private Function CheckBillIsAllDels(ByVal strNo As String, Optional ByRef strSel序号 As String) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查指定的单据是否全部选中退费
    '入参:strNO-单据号
    '出参:strSel序号-返回选中的序号
    '返回:0-全部未选择;1-全部选择;2-选择了一部分
    '编制:刘兴洪
    '日期:2011-01-24 16:43:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim k As Long, j As Long, lngCount As Long, str序号 As String
    With vsBill
        k = vsBill.FindRow(strNo, , vsBill.ColIndex("单据号"))
         For j = k To vsBill.Rows - 1
             If vsBill.TextMatrix(j, .ColIndex("单据号")) <> strNo Then Exit For
             If Val(vsBill.TextMatrix(j, .ColIndex("选择"))) <> 0 Then
                 str序号 = str序号 & "," & CLng(vsBill.RowData(j))
             End If
             lngCount = lngCount + 1
         Next
     End With
     
     If str序号 <> "" Then str序号 = Mid(str序号, 2)
     strSel序号 = str序号
     If str序号 = "" Then CheckBillIsAllDels = 0: Exit Function
     If lngCount = UBound(Split(str序号, ",")) + 1 Then
        If InStr(1, mstrNOsPatiDel & ",", "," & strNo & ",") > 0 Then
            CheckBillIsAllDels = 2: Exit Function
        End If
        CheckBillIsAllDels = 1: Exit Function
     End If
    CheckBillIsAllDels = 2
End Function
Private Function zlCheckPrepayBack(ByVal lng病人ID As Long, ByVal strSelNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查是否存在预交,如果存在预交,则根据消费确认原则让用户刷卡消费
    '入参:
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-04-15 14:46:41
    '问题:37307
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTmp As Variant, dblMoney As Double
    Dim strFilter As String, i As Long
    '没有选择的单据,返回true
    If strSelNos = "" Then zlCheckPrepayBack = True: Exit Function
    If lng病人ID = 0 Then zlCheckPrepayBack = True: Exit Function
    If gbyt预存款退费验卡 = 0 Then zlCheckPrepayBack = True: Exit Function
    varTmp = Split(strSelNos, ","): strFilter = ""
    For i = 0 To UBound(varTmp)
        strFilter = strFilter & " or NO='" & varTmp(i) & "'"
    Next
    strFilter = Mid(strFilter, 4)
    On Error GoTo errHandle
    mrsBalance.Filter = strFilter
    If mrsBalance.RecordCount <> 0 Then mrsBalance.MoveFirst
    Do While Not mrsBalance.EOF
        If Nvl(mrsBalance!性质) = 1 Then
            dblMoney = dblMoney + Val(Nvl(mrsBalance!结算金额))
        End If
        mrsBalance.MoveNext
    Loop
    mrsBalance.Filter = 0
    '问题:37307
    If dblMoney = 0 Then zlCheckPrepayBack = True: Exit Function
    If Not zlDatabase.PatiIdentify(Me, glngSys, lng病人ID, dblMoney, , , , , , , , (gbyt预存款退费验卡 = 2)) Then Exit Function
    zlCheckPrepayBack = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub ReInitPatiInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新初始化病人发票信息
    '编制:刘兴洪
    '日期:2011-04-29 14:17:33
    '问题:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInvoiceFormat As String
    If mbytMode = 0 Then Exit Sub
    mstrUseType = "": mlngShareUseID = 0: mintInvoiceFormat = 0
    mstrUseType = zl_GetInvoiceUserType(mlng病人ID, 0, mintInsure)
    mlngShareUseID = zl_GetInvoiceShareID(mlngModule, mstrUseType)
    mintInvoiceFormat = zl_GetInvoicePrintFormat(mlngModule, mstrUseType, mintOldInvoiceFormat)
    mintInvoicePrint = zl_GetInvoicePrintMode(mlngModule, mstrUseType)
    
End Sub
Private Function zlGetInvoiceGroupUseID(ByRef lng领用ID As Long, _
    Optional intNum As Integer = 1, Optional strInvoiceNO As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取票据的领用ID
    '入参:lng领用ID-领用id
    '       intNum-页数
    '       strInvoiceNO-输入的发票号
    '出参:lng领用ID-领用ID
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-04-29 15:36:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    lng领用ID = GetInvoiceGroupID(1, intNum, lng领用ID, mlngShareUseID, strInvoiceNO, mstrUseType)
    If lng领用ID <= 0 Then
        Select Case lng领用ID
            Case 0 '操作失败
            Case -1
                If Trim(mstrUseType) = "" Then
                    MsgBox "你没有自用和共用的收费票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                Else
                    MsgBox "你没有自用和共用的『" & mstrUseType & "』收费票据,请先领用一批票据或设置本地共用票据！", vbInformation, gstrSysName
                End If
                Exit Function
            Case -2
                If Trim(mstrUseType) = "" Then
                    MsgBox "本地的共用票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                Else
                    MsgBox "本地的共用票据的『" & mstrUseType & "』收费票据已经用完,请先领用一批票据或重新设置本地共用票据！", vbInformation, gstrSysName
                End If
                Exit Function
        End Select
    End If
    zlGetInvoiceGroupUseID = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function



Private Function zlCheckDelValied(ByVal lng卡类别ID As Long, _
     ByVal strName As String, _
     ByVal bln消费卡 As Boolean, ByVal strCardNo As String, _
     ByVal strSwapNO As String, strSwapMemo As String, _
     ByRef lng结帐ID As Long, _
    ByVal dbl退款金额 As Double, Optional bln异常单据 As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查退费交易接口
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2010-02-08 16:40:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strXMLExend As String
    If lng卡类别ID = 0 Then zlCheckDelValied = True: Exit Function
    If Not mobjSquare Is Nothing Then
        Call initCardSquareData
    End If
    If mobjSquare Is Nothing Then
    
        MsgBox "注意:" & vbCrLf & _
                     "      当前收费是按" & strName & " 收费的,但不存在操作的相关部件,不能退款,请与系统管理员联系!", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
        Exit Function
    End If
    
'zlReturnCheck(frmMain As Object, ByVal lngModule As Long, _
    ByVal lngCardTypeID As Long, bln消费卡 As Boolean, ByVal strCardNo As String, _
    ByVal strBalanceIDs As String, _
    ByVal dblMoney As Double, ByVal strSwapNo As String, _
    ByVal strSwapMemo As String, ByRef strXMLExpend As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:帐户回退交易前的检查
    '入参:frmMain-调用的主窗体
    '       lngModule-调用的模块号
    '       lngCardTypeID-卡类别ID
    '       strCardNo-卡号
    '       strBalanceIDs   String  In  本次支付所涉及的结算ID 格式:收费类型|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '                                   收费类型: 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
    '       dblMoney-退款金额
    '       strSwapNo-交易流水号(退款时检查)
    '       strSwapMemo-交易说明(退款时传入)
    '       strXMLExpend    XML IN  可选参数:异常单据重新退费(1)
    '返回:退款合法,返回true,否则返回Flase
    strXMLExend = IIf(bln异常单据, 1, "")
      If mobjSquare.zlReturnCheck(Me, mlngModule, lng卡类别ID, bln消费卡, strCardNo, _
        "3|" & lng结帐ID, dbl退款金额, strSwapNO, strSwapMemo, strXMLExend) = False Then
          zlCheckDelValied = False
          Exit Function
     End If
goEnd:
    zlCheckDelValied = True
    Exit Function
End Function

Private Function CheckBrushCard(ByVal lng卡类别ID As Long, ByVal bln消费卡 As Boolean, _
    ByVal dbl退费额 As Double, ByRef strBrushCardNo As String, ByRef strbrPassWord As String, Optional ByRef bln退现 As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查刷卡
    '返回:合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-18 14:52:47
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsMoney As ADODB.Recordset
    On Error GoTo errHandle
    Dim dblMoney As Double
     '弹出刷卡界面
    'zlBrushCard(frmMain As Object, _
    'ByVal lngModule As Long, _
    'ByVal rsClassMoney As ADODB.Recordset, _
    'ByVal lngCardTypeID As Long, _
    'ByVal bln消费卡 As Boolean, _
    'ByVal strPatiName As String, ByVal strSex As String, _
    'ByVal strOld As String, ByVal dbl金额 As Double, _
    'Optional ByRef strCardNo As String, _
    'Optional ByRef strPassWord As String, _
    Optional ByRef bln退费 As Boolean = False, _
    Optional ByRef blnShowPatiInfor As Boolean = False, _
    Optional ByRef bln退现 As Boolean) As Boolean
    If mobjSquare.zlBrushCard(Me, mlngModule, Nothing, lng卡类别ID, bln消费卡, mtyPati.姓名, mtyPati.性别, mtyPati.年龄, dbl退费额, strBrushCardNo, strbrPassWord, _
        True, True, bln退现) = False Then Exit Function
    CheckBrushCard = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function CallBackBalanceInterface(ByVal lng卡类别ID As Long, ByVal bln消费卡 As Boolean, _
    ByVal strCardNo As String, ByVal strSwapGlideNO As String, ByVal strSwapMemo As String, _
    ByVal str结帐IDs As String, str冲销IDs As String, _
    ByVal dblMoney As Double, _
    ByRef cllUpdate As Collection, _
    ByRef cllThreeSwap As Collection, ByRef strErrMsg As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用回退接口
    '入参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-13 10:33:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str卡号 As String, str结算信息 As String, strSwapExtendInfor As String
    Dim varData As Variant, varTemp As Variant, i As Long, cllPro As Collection
    Dim strSQL As String, rsTemp As ADODB.Recordset, strTemp As String
    
    Err = 0: On Error GoTo Errhand:
    If lng卡类别ID = 0 Then CallBackBalanceInterface = True: Exit Function
    
    If str结帐IDs <> "" Then str结算信息 = str结算信息 & "||3|" & str结帐IDs
    If str结算信息 <> "" Then str结算信息 = Mid(str结算信息, 3)
    
    If str冲销IDs = "" Then
    strSQL = "" & _
    "   Select /*+ RULE */ distinct   A.结帐ID  " & _
    "   From  门诊费用记录 A,门诊费用记录 B,table(f_num2list([1])) P " & _
    "   Where A.NO=B.NO and A.记录性质=1 And A.记录状态=2  " & _
    "           And B.结帐ID=P.Column_Value " & _
    "             "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str结帐IDs)
    With rsTemp
            str冲销IDs = ""
            Do While Not .EOF
                str冲销IDs = str冲销IDs & "," & Val(Nvl(!结帐ID))
                .MoveNext
            Loop
        End With
        If str冲销IDs <> "" Then str冲销IDs = Mid(str冲销IDs, 2)
    End If
    If str冲销IDs = "" Then str冲销IDs = "0"
    '81489,冉俊明,2015-1-22,退费传入冲销ID
    strSwapExtendInfor = "3|" & str冲销IDs: strTemp = strSwapExtendInfor
    
    'zlReturnMoney(frmMain As Object, ByVal lngModule As Long, _
        ByVal lngCardTypeID As Long, ByVal strCardNo As String, ByVal strBalanceIDs As String, _
        ByVal dblMoney As Double, _
        ByRef strSwapGlideNO As String, ByRef strSwapMemo As String, _
        ByRef strSwapExtendInfor As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:帐户扣款回退交易
    '入参:frmMain-调用的主窗体
    '       lngModule-调用的模块号
    '       lngCardTypeID-卡类别ID:医疗卡类别.ID
    '       strCardNo-卡号
    '       strBalanceIDs-本次支付所涉及的结算ID(这是原结帐ID):
    '                           格式:收费类型(|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '                           收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
    '       dblMoney-退款金额
    '       strSwapNo-交易流水号(扣款时的交易流水号)
    '       strSwapMemo-交易说明(扣款时的交易说明)
    '       strSwapExtendInfor-传入，本次退费的冲销ID：
    '                           格式:收费类型1|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
    '                           收费类型:1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款
    '       strSwapExtendInfor-传出，交易的扩展信息
    '           格式为:项目名称1|项目内容2||…||项目名称n|项目内容n 每个项目中不能包含|字符
    If mobjSquare.zlReturnMoney(Me, mlngModule, lng卡类别ID, bln消费卡, strCardNo, str结算信息, dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then Exit Function
    Set cllUpdate = New Collection: Set cllThreeSwap = New Collection
    Call zlAddUpdateSwapSQL(False, str冲销IDs, lng卡类别ID, bln消费卡, str卡号, strSwapGlideNO, strSwapMemo, cllUpdate, 2)
    If strTemp <> strSwapExtendInfor Then
        Call zlAddThreeSwapSQLToCollection(False, str冲销IDs, lng卡类别ID, bln消费卡, strCardNo, strSwapExtendInfor, cllThreeSwap)
    End If
    CallBackBalanceInterface = True
Errhand:
End Function

Private Function OverFeeDel(ByVal str冲销IDs As String, ByVal lng病人ID As Long, ByRef blnCommited As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:完成退费收费
    '入参:str冲销IDs-完成收费的单据(可以为多张的结帐ID,但目前只有一张单据)
    '出参:
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-29 14:50:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    If Left(str冲销IDs, 1) = "," Then str冲销IDs = Mid(str冲销IDs, 2)

    On Error GoTo errHandle
    ' Zl_门诊收费结算_完成退费
    strSQL = "Zl_门诊收费结算_完成退费("
    '  病人id_In       门诊费用记录.病人id%Type,
    strSQL = strSQL & "" & lng病人ID & ","
    '  退费结算序号_In 病人预交记录.结算序号%Type,
    strSQL = strSQL & "NULL,"
    '  冲销ids_In      Varchar2,
    strSQL = strSQL & "'" & str冲销IDs & "',"
    '  操作员姓名_In   病人预交记录.操作员姓名%Type := Null
    strSQL = strSQL & "'" & UserInfo.姓名 & "')"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    gcnOracle.CommitTrans: blnCommited = True
    OverFeeDel = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
    blnCommited = True
End Function
Private Sub ClearBalance()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除结算数据
    '编制:刘兴洪
    '日期:2011-11-22 15:59:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsBalance
        .Clear 1: .COLS = 1
'        .Cell(flexcpData, 0, 1, .Rows - 1, .COLS - 1) = ""
'        .Cell(flexcpText, 0, 1, .Rows - 1, .COLS - 1) = ""
        .Editable = flexEDKbdMouse
    End With
End Sub
Private Sub LoadBalanceInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载应收款结算
    '编制:刘兴洪
    '日期:2011-11-22 15:45:46
    '问题:43403
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNotNos  As String, lngRow As Long, strFilter As String, str结算方式 As String
    Dim strBalance As String, str缺省结算方式 As String
    
    strNotNos = Replace(mstrDelNOs, "'", "")
    lngRow = 0
    mrsBalance.Filter = 0
    If strNotNos <> "" Then
         strFilter = Replace(strNotNos, ",", "' and  NO<>'")
         strFilter = " NO<>'" & strFilter & "'"
         mrsBalance.Filter = strFilter
    End If
    mrsBalance.Sort = "结算性质,应付款,结算方式"
    With vsBalance
        .Redraw = flexRDNone
        Call ClearBalance
         If mrsBalance.RecordCount <> 0 Then mrsBalance.MoveFirst
         Do While Not mrsBalance.EOF
            '--问题:52530
            If InStr(1, ",1,2,3,4,5,", "," & Val(Nvl(mrsBalance!性质)) & ",") = 0 Then
                 '--1预交 ,2  '医保 ,3, 4 '医疗卡和结算卡,5 '一卡通
                 str缺省结算方式 = Nvl(mrsBalance!结算方式, "  ")
            End If
             str结算方式 = Nvl(mrsBalance!结算方式, "  ")
             If str结算方式 <> strBalance Then
                 strBalance = str结算方式: .COLS = .COLS + 2
                  .ColAlignment(.COLS - 2) = 7: .ColAlignment(.COLS - 1) = 1
             End If
             
             .TextMatrix(lngRow, .COLS - 2) = strBalance & ":"
             .Cell(flexcpData, lngRow, .COLS - 2) = strBalance
             .TextMatrix(lngRow, .COLS - 1) = Val(.TextMatrix(lngRow, .COLS - 1)) + Nvl(mrsBalance!结算金额, 0)
             .Cell(flexcpData, lngRow, .COLS - 1, lngRow, .COLS - 1) = Val(Nvl(mrsBalance!是否退现))
             mCurBillType.bln三方卡全退 = Val(Nvl(mrsBalance!是否全退)) = 1
             
             '多单据使用多种结算时,单笔结算金额看没有进行分币处理,所以不能用format取两位数
             .ColData(.COLS - 2) = "摘要:" & mrsBalance!摘要
             .ColData(.COLS - 1) = "结算号码:" & mrsBalance!结算号码
             
             If mrsBalance!结算性质 <> 1 Then
                .Cell(flexcpForeColor, lngRow, .COLS - 1, lngRow, .COLS - 2) = vbBlue
                .Cell(flexcpForeColor, 1, .COLS - 1, 1, .COLS - 2) = vbRed
                .Cell(flexcpFontBold, 1, .COLS - 1, 1, .COLS - 2) = True    '粗体
            End If
             mrsBalance.MoveNext
            .Redraw = flexRDBuffered
         Loop
         vsBalance.AutoSizeMode = flexAutoSizeColWidth
         Call vsBalance.AutoSize(0, .COLS - 1)
         
         If mblnSingleBlance And str缺省结算方式 <> "" Then
            mblnNotClick = True
            zlControl.CboSetText cbo退款方式, str缺省结算方式
            mblnNotClick = False
         End If
         If mbytMode = 0 Then
            .RowHidden(1) = True: ControlResize
         End If
    End With
End Sub
Private Sub LoadPart预存款(ByVal strNos As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载部分退未费的预存款
    '编制:刘兴洪
    '日期:2011-12-01 11:26:48
    '问题:43403
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varNos  As Variant, i As Long, strNo As String, j As Long, k As Long
    Dim dbl选择合计 As Double
    If strNos = "" Then Exit Sub
    If Left(strNos, 1) = "," Then strNos = Mid(strNos, 2)
    '未选择,但可能存在部分选择,又只用了一种预存款的,则只能退预存款
    varNos = Split(strNos, ",")
    With vsBill
        For i = 0 To UBound(varNos)
            strNo = varNos(i)
            mrsBalance.Filter = " NO='" & strNo & "' and 性质<>1 "
            If mrsBalance.RecordCount = 0 Then
                '只有一种结算方式
                k = vsBill.FindRow(strNo, , .ColIndex("单据号"))
                For j = k To vsBill.Rows - 1
                    If vsBill.TextMatrix(j, .ColIndex("单据号")) <> strNo Then Exit For
                    If Val(vsBill.TextMatrix(j, .ColIndex("选择"))) <> 0 Then
                        dbl选择合计 = RoundEx(dbl选择合计 + Val(vsBill.TextMatrix(j, .ColIndex("实收金额"))), 6)
                    End If
                Next
            End If
        Next
    End With
    With vsBalance
        For i = 1 To .COLS - 1 Step 2
            If .Cell(flexcpData, 0, i) = "预存款" Then
                If Trim(.TextMatrix(1, i)) = "" Then
                    .Cell(flexcpData, 1, i) = "预存款"
                    .TextMatrix(1, i) = "预存款"
                End If
                .TextMatrix(1, i + 1) = Val(.TextMatrix(1, i + 1)) + dbl选择合计
                .Cell(flexcpData, 1, i + 1) = Val(.TextMatrix(1, i + 1))
                .Cell(flexcpFontBold, 1, i + 1) = True
                .Cell(flexcpForeColor, 1, i + 1) = vbRed
                .RowHidden(1) = False
                Exit For
            End If
        Next
        txt退款金额.Tag = dbl选择合计
    End With
End Sub
Private Sub LoadDelBalanceInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载退款结算方式
    '编制:刘兴洪
    '日期:2011-11-22 15:45:46
    '问题:43403
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strBalance As String, varNO As Variant, i As Long, str结算 As String
    Dim strSelNo As String, bln全选 As Boolean, bln未选 As Boolean
    Dim strFilter As String, bln退款 As Boolean, intCol As Integer
    Dim str结算方式 As String, strNotNos As String
    Dim bln部分选择 As Boolean, lngRow As Long
    Dim blnYb As Boolean, strPartNO As String, varPatiNo As Variant
    Dim strTemp As String
    Dim bln预存 As Boolean
    Dim blnThreeSwap As Boolean
    Dim strTempBalance As String
    Dim str普通结算 As String
    
    Err = 0: On Error GoTo Errhand:
    strNotNos = Replace(mstrDelNOs, "'", "")
    lngRow = 1
    If mstrNOs = "" Then Exit Sub
    varNO = Split(mstrNOs, ",")
    blnThreeSwap = False
    bln全选 = True: bln未选 = True: strSelNo = ""
    strPartNO = ""
    For i = 0 To UBound(varNO)
       Select Case CheckBillIsAllDels(varNO(i))
       Case 0   '未选择
            bln全选 = False
       Case 1   '全选择
            strSelNo = strSelNo & "," & varNO(i): bln未选 = False
       Case Else    '部分选择
            bln全选 = False: bln未选 = False: bln部分选择 = True
            strPartNO = strPartNO & "," & varNO(i)
       End Select
    Next
    
    '未选择
    vsBalance.RowHidden(lngRow) = bln部分选择
    If bln部分选择 And mCurBillType.bln存在医疗卡结算 And mCurBillType.blnSingleBalance Then
        With vsBalance
            .RowHidden(lngRow) = mCurBillType.bln三方卡全退
            .Redraw = flexRDNone
            .TextMatrix(1, 1) = .TextMatrix(0, 1)
            .Cell(flexcpData, 1, 1) = .Cell(flexcpData, 0, 1)
            .TextMatrix(1, 2) = .Cell(flexcpData, 1, 2)
            .ColData(2) = .Cell(flexcpData, 0, 2) '是否退现
            Call ControlResize
            .Redraw = flexRDDirect
        End With
        Exit Sub
    End If
    If strSelNo = "" Then
        vsBalance.Redraw = flexRDNone
        Call LoadPart预存款(strPartNO)
        If bln未选 And vsBalance.COLS > 1 Then
            vsBalance.Cell(flexcpText, 1, 1, 1, vsBalance.COLS - 1) = ""
            vsBalance.Cell(flexcpData, 1, 1, 1, vsBalance.COLS - 1) = ""
        End If
        Call ControlResize
        vsBalance.Redraw = flexRDDirect
        Exit Sub
    End If
       
    strSelNo = Mid(strSelNo, 2): strTempBalance = ""
    strFilter = Replace(strSelNo, ",", "' Or NO='")
    strFilter = " NO='" & strFilter & "'"
    mrsBalance.Filter = strFilter
    mrsBalance.Sort = "结算性质,应付款,结算方式"
    With vsBalance
        .Redraw = flexRDNone
        .Cell(flexcpData, 1, 0, 1, .COLS - 1) = ""
        If .COLS - 1 > 0 Then
            .Cell(flexcpText, 1, 1, 1, .COLS - 1) = ""
        End If
         If mrsBalance.RecordCount <> 0 Then mrsBalance.MoveFirst
         intCol = 1: strBalance = ""
         Do While Not mrsBalance.EOF
             str结算方式 = Nvl(mrsBalance!结算方式, "  ")
             If str结算方式 <> strBalance Then
                 For intCol = 1 To .COLS - 1 Step 2
                    If .Cell(flexcpData, 0, intCol) = str结算方式 Then
                        Select Case Val(Nvl(mrsBalance!性质))
                        Case 1  '预交
                            bln预存 = True
                            Exit For
                        Case 2 '医保
                             '如果这种结算方式不支持回退,要退为现金,则不用减去
                            If mblnYB结算作废 Then
                                If gclsInsure.GetCapability(support门诊结算作废, mlng病人ID, mintInsure, str结算方式) Then Exit For
                                 blnYb = True
                            Else     '不支持门诊结算作废时,只允许个帐退为现金,其它原样退,不调用医保交易
                                If str结算方式 <> mstr个人帐户 Then Exit For
                                blnYb = True
                            End If
                        Case 3, 4 '医疗卡和结算卡
                            blnThreeSwap = True
                            Exit For
                        Case 5 '一卡通
                            blnThreeSwap = True
                            Exit For
                        Case Else
                            str普通结算 = str普通结算 & "," & str结算方式
                            If mBillDelType = EM_多张全退 Then
                                '49155:问题 '全退,不允许单张退
'                                If InStr(",1,", "," & mrsBalance!结算性质 & ",") = 0 Then Exit For
                                 Exit For
                            End If
                            '单张作全退时
                            If mBillDelType = EM_单张全退 Then Exit For
                            strTempBalance = strTempBalance & "," & str结算方式
                            If InStr(",1,2,", "," & mrsBalance!结算性质 & ",") = 0 Then Exit For
                            If bln全选 And mstrDelNOs = "" Then Exit For
                        End Select
                    End If
                 Next
                  strBalance = str结算方式
             End If
             If intCol < .COLS - 1 Then
                .TextMatrix(lngRow, intCol) = strBalance & ":"
                .TextMatrix(lngRow, intCol + 1) = Val(.TextMatrix(lngRow, intCol + 1)) + Nvl(mrsBalance!结算金额, 0)
                .Cell(flexcpData, lngRow, intCol, lngRow, intCol) = strBalance
                .Cell(flexcpData, lngRow, intCol + 1, lngRow, intCol + 1) = .TextMatrix(lngRow, intCol + 1)
                .ColData(intCol + 1) = Val(Nvl(mrsBalance!是否退现))
            End If
             mrsBalance.MoveNext
         Loop
         
         If blnYb Then
            For i = 1 To .COLS - 1 Step 2
               If InStr(strTempBalance & ",", "," & .Cell(flexcpData, lngRow, i)) > 0 Then
                   .TextMatrix(lngRow, i) = "": .Cell(flexcpData, lngRow, i) = ""
                   .TextMatrix(lngRow, i + 1) = "": .Cell(flexcpData, lngRow, i + 1) = ""
               End If
            Next
         End If
         If blnThreeSwap And (mCurBillType.bln多张部分退费 Or Not bln全选) And bln部分选择 = False Then
            For i = 1 To .COLS - 1 Step 2
               If InStr(str普通结算 & ",", "," & .Cell(flexcpData, lngRow, i)) > 0 Then
                   .TextMatrix(lngRow, i) = "": .Cell(flexcpData, lngRow, i) = ""
                   .TextMatrix(lngRow, i + 1) = "": .Cell(flexcpData, lngRow, i + 1) = ""
               End If
            Next
         End If
        Call LoadPart预存款(strPartNO)
        If .COLS - 1 > 0 Then
            .Cell(flexcpForeColor, lngRow, 1, lngRow, .COLS - 1) = vbRed
        End If
        .Cell(flexcpFontBold, lngRow, 0, lngRow, .COLS - 1) = True  '粗体
         Call vsBalance.AutoSize(0, .COLS - 1)
         If .COLS - 1 > 0 Then
            .Row = .FixedRows: .Col = .FixedCols
        End If
        
        .RowHidden(lngRow) = (bln部分选择 Or mCurBillType.bln单张部分退费 _
                            And Not (mCurBillType.bln存在医疗卡结算 And mCurBillType.blnSingleBalance)) And bln预存 = False
         
        If Not mblnSingleBlance Then
            '不是单种结算方式
            If mintInsure = 0 And Not mCurBillType.bln存在卡结算 Then
                .RowHidden(lngRow) = .RowHidden(lngRow) Or Not bln全选
            End If
              
        End If
         ControlResize
        .Redraw = flexRDBuffered
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
Private Sub InitRecErrCurStruct(ByRef rsErr As ADODB.Recordset)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始异常单据的数据结构
    '入参:
    '出参:
    '返回:
    '编制:刘兴洪
    '日期:2012-01-16 15:09:28
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set rsErr = New ADODB.Recordset
    rsErr.Fields.Append "NO", adVarChar, 20, adFldIsNullable
    rsErr.Fields.Append "结算金额", adDouble, , adFldIsNullable
    rsErr.CursorLocation = adUseClient
    rsErr.LockType = adLockOptimistic
    rsErr.CursorType = adOpenStatic
    rsErr.Open
End Sub

Private Sub zlALLNosBack()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:所有单据一起退
    '编制:刘兴洪
    '日期:2011-11-24 09:53:09
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varNO As Variant, i As Long
    Dim dblToTal As Double, strNo As String
    Dim dblMoney As Double, intCol As Integer
    Dim bln退现 As Boolean, bln现金结算 As Boolean
    Dim cur退费合计 As Double, cur误差金额 As Double
    Dim dblTemp As Double, dbl误差合计 As Double
    Dim rsErr As ADODB.Recordset
    Dim bln原样退 As Boolean
    bln现金结算 = False
    If cbo退款方式.ListIndex <> -1 Then
        If cbo退款方式.ItemData(cbo退款方式.ListIndex) = 1 Then
            bln现金结算 = True
        End If
    End If
    
    Call InitRecErrCurStruct(rsErr)
    varNO = Split(mstrNOs, ",")

    For i = 0 To UBound(varNO)
        strNo = CStr(varNO(i))
        mcolError.Add 0, "_" & strNo
    Next
    
    If cbo退款方式.ListIndex = -1 And cbo退款方式.ListCount > 0 Then cbo退款方式.ListIndex = 0
    cbo退款方式.Enabled = False
    cbo退款方式.Locked = True
    
    dblToTal = 0
    ''不包括医保和预交款金额
    '性质:1-预存款,2-医保,3-医疗卡,4-结算卡,5-一卡通,0-其他类

    mrsBalance.Filter = 0 '  "性质<>1"
    mrsBalance.Sort = "性质 DESC"
    bln原样退 = True
    With mrsBalance
         If .RecordCount <> 0 Then .MoveFirst
         Do While Not .EOF
            bln退现 = False
            Select Case Val(Nvl(!性质))
            Case 1 '预存
            Case 2
                '49155:问题
                If mblnYB结算作废 Then
                    If Not gclsInsure.GetCapability(support门诊结算作废, mlng病人ID, mintInsure, Nvl(!结算方式)) Then
                        bln退现 = True: bln原样退 = False
                    End If
                Else
                    '不支持门诊结算作废时,只允许个帐退为现金,其它原样退,不调用医保交易
                    If !结算方式 = mstr个人帐户 Then
                        bln退现 = True: bln原样退 = False
                    End If
                End If
            Case 3, 4, 5    '3-医疗卡,4-结算卡,5-一卡通
                If Val(Nvl(!是否退现)) = 1 Then
                        For intCol = 1 To vsBalance.COLS - 1 Step 2
                            If vsBalance.Cell(flexcpData, 1, intCol) = !结算方式 Then
                                If Val(vsBalance.TextMatrix(1, intCol + 1)) = 0 And vsBalance.RowHidden(1) = False Then
                                    bln退现 = True: bln原样退 = False: Exit For
                                End If
                            End If
                        Next
                End If
            Case Else
                '49155:问题
               ' If Val(Nvl(mrsBalance!结算性质)) = 1 Then
               If !结算方式 = zlStr.NeedName(cbo退款方式) And bln原样退 = False Then
                    bln退现 = True
               End If
        
                For intCol = 1 To vsBalance.COLS - 1 Step 2
                    If 1 = 1 _
                        And vsBalance.Cell(flexcpData, 1, intCol) = !结算方式 Then
                        If bln退现 And Not bln原样退 Then
                            vsBalance.TextMatrix(1, intCol) = ""
                            vsBalance.TextMatrix(1, intCol + 1) = ""
                        ElseIf vsBalance.TextMatrix(1, intCol) = "" Then
                            vsBalance.TextMatrix(1, intCol) = vsBalance.Cell(flexcpData, 1, intCol) & ":"
                            vsBalance.TextMatrix(1, intCol + 1) = vsBalance.Cell(flexcpData, 1, intCol + 1)
                        End If
                    End If
                Next
            End Select
            
            If bln退现 Then
                rsErr.Find "NO='" & Nvl(!NO) & "'"
                If rsErr.EOF Then rsErr.AddNew
                rsErr!NO = CStr(Nvl(!NO))
                rsErr!结算金额 = RoundEx(Val(Nvl(rsErr!结算金额)) + Val(Nvl(!结算金额)), 6)
                rsErr.Update
                dblMoney = RoundEx(dblMoney + !结算金额, 6)
            End If
            dblToTal = RoundEx(dblToTal + !结算金额, 6)
             .MoveNext
        Loop
    End With
    
    dbl误差合计 = 0: dblMoney = 0:   dblTemp = 0
    rsErr.Filter = 0
    With rsErr
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            dblTemp = dblTemp + Val(Nvl(!结算金额))
            cur退费合计 = cur退费合计 + Format(Val(Nvl(!结算金额)), "0.00")
            .MoveNext
        Loop
    End With
    
    If bln现金结算 And dblTemp <> 0 Then
          dblTemp = dblTemp - Get误差费(mstrNOs)
          If mintInsure > 0 Then
              If gclsInsure.GetCapability(support分币处理, mlng病人ID, mintInsure) Then
                  cur退费合计 = CentMoney(dblTemp)
              End If
          Else
              cur退费合计 = CentMoney(dblTemp)
          End If
    End If
    cur误差金额 = RoundEx(cur退费合计 - dblTemp, 6)
    dbl误差合计 = dbl误差合计 + cur误差金额
    dblMoney = dblMoney + cur退费合计
    mcolError.Remove "_" & strNo
    mcolError.Add dbl误差合计, "_" & strNo

    txt退款金额.ToolTipText = ""
    If dbl误差合计 <> 0 Then
        txt退款金额.ToolTipText = "退费误差金额:" & Format(dbl误差合计, gstrDec)
    End If

    txt退款合计.ToolTipText = txt退款金额.ToolTipText
    txt退款金额.Text = Format(dblMoney, "0.00")
    txt退款合计.Text = Format(dblToTal, "0.00")
    
    If mBillDelType = EM_多张全退 And bln原样退 Then dblMoney = 0      '原样退
         
    '费用金额保留位数,及现金结算时处理分币
    cbo退款方式.Locked = dblMoney = 0
    cbo退款方式.Enabled = dblMoney <> 0
    cbo退款方式.Visible = dblMoney <> 0
End Sub

Private Sub ReCalcDelMoney()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新计算退费金额
    '编制:刘兴洪
    '日期:2011-11-22 16:50:38
    '说明:根据当前界面退费选择情况，计算退款金额和误差金额
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cur单据合计 As Currency, cur选择合计 As Currency
    Dim cur退费合计 As Currency, cur误差金额 As Currency, cur误差合计 As Currency
    Dim bln完全退费 As Boolean, bln现金结算 As Boolean
    Dim curTotal As Currency, strNo As String
    Dim i As Long, j As Long, k As Long, bln原样退 As Boolean
    Dim colAllReturn As Collection, bln退现 As Boolean
    Dim intCol As Long, bln其他 As Boolean
    Dim bln未选 As Boolean, varNO As Variant
    Dim bln全退 As Boolean, blnFind As Boolean
    Dim dbl退款合计 As Double, dblBalanceSum As Double
    Dim dblCashSum As Double '现金合计
    Dim blnHaveSelected As Boolean, blnHaveNotSelected As Boolean
    
    If mbytMode = 0 Then Exit Sub
    If mrsBalance Is Nothing Then Exit Sub
        
    Set mcolError = New Collection
    Set colAllReturn = New Collection
    
       
    If mBillDelType = EM_多张全退 Then
        '多张单据一起退,无误差
        Call zlALLNosBack: curTotal = Val(txt退款合计): GoTo GoSetVisible: Exit Sub
    End If
    
    '非一次退费
   varNO = Split(mstrNOs, ",")
    
    '1.先判断整个是否是原样退,以决定是否禁用结算方式选择,以及分币误差的生成
    bln原样退 = True: bln退现 = False: bln其他 = False
    bln全退 = True
    dblBalanceSum = 0: dblCashSum = 0
    For i = 0 To UBound(varNO)
        strNo = CStr(varNO(i))
        cur单据合计 = 0: cur选择合计 = 0
        blnHaveNotSelected = False
        With vsBill
            k = vsBill.FindRow(strNo, , .ColIndex("单据号"))
            For j = k To vsBill.Rows - 1
                If vsBill.TextMatrix(j, .ColIndex("单据号")) <> strNo Then Exit For
                cur单据合计 = cur单据合计 + Val(vsBill.TextMatrix(j, .ColIndex("实收金额")))
                If Val(vsBill.TextMatrix(j, .ColIndex("选择"))) <> 0 Then
                    cur选择合计 = cur选择合计 + Val(vsBill.TextMatrix(j, .ColIndex("实收金额")))
                    blnHaveSelected = True
                Else
                    blnHaveNotSelected = True
                End If
            Next
        End With
        
        bln完全退费 = BillDeleteAll(strNo, 1, mblnHaveExcuteData)
        bln完全退费 = bln完全退费 And Not BillExistDelete(strNo, 1) And (cur单据合计 = cur选择合计 And blnHaveNotSelected = False) '零费用，blnHaveNotSelected

        If mCurBillType.bln存在卡结算 And mCurBillType.bln单种结算方式 And bln完全退费 Then
        
        ElseIf mCurBillType.bln存在卡结算 = False And mintInsure = 0 And bln完全退费 Then
            If InStr(mstrNOs, ",") > 0 Then
                bln完全退费 = Not mCurBillType.bln多张部分退费
            End If
        End If
                
        If bln全退 And Not bln完全退费 Then bln全退 = False
        
        colAllReturn.Add Array(IIf(bln完全退费, 1, 0), strNo, cur单据合计, cur选择合计), "_" & strNo    '保存用于后面的判断
        If Not bln完全退费 Then bln原样退 = False
        
        dbl退款合计 = RoundEx(dbl退款合计 + cur选择合计, 5)
        If bln完全退费 Then
            mrsBalance.Filter = "NO='" & strNo & "'"
            With mrsBalance
                If .RecordCount <> 0 Then .MoveFirst
                Do While Not .EOF
                    If Val(Nvl(!性质)) = 2 Then '医保
                        If mblnYB结算作废 Then
                            If Not gclsInsure.GetCapability(support门诊结算作废, mlng病人ID, mintInsure, !结算方式) Then
                               bln原样退 = False
                            End If
                        ElseIf !结算方式 = mstr个人帐户 Then
                             bln原样退 = False
                        End If
                    ElseIf InStr("3,4,5", Val(Nvl(!性质))) > 0 Then
                        '一卡通相关
                        If Val(Nvl(!是否退现)) = 1 Then
                            For intCol = 1 To vsBalance.COLS - 1 Step 2
                                If vsBalance.Cell(flexcpData, 1, intCol) = !结算方式 And Val(vsBalance.TextMatrix(1, intCol + 1)) = 0 Then
                                    bln退现 = True: bln原样退 = False
                                End If
                            Next
                        End If
                    Else
                        If mCurBillType.bln多张部分退费 Then bln原样退 = False  '55675
                        bln其他 = True
                    End If
                    dblBalanceSum = dblBalanceSum + Nvl(!结算金额, 0)
                    If !结算方式 = mstr现金结算方式 And Val(Nvl(!性质)) <> 1 Then dblCashSum = dblCashSum + Nvl(!结算金额, 0)
                    .MoveNext
            Loop
            End With
        End If
    Next
    
    If bln全退 And mstrDelNOs <> "" Then bln全退 = False
    
    '收费时全部用预交(结算方式为空),退费时,不允许指定退费方式
    '性质:1-预存款,2-医保,3-医疗卡(一卡通),4-结算卡,5-一卡通,0-其他类
    mrsBalance.Filter = "性质<>1"
    If mrsBalance.RecordCount = 0 Then bln原样退 = True
    mrsBalance.Filter = 0
   ' If mBillDelType = EM_单张全退 Then bln原样退 = True
         
 
    '需要确定多张单据中的单张退
    If bln原样退 Then
        '可能存在部分退的情况
        If dblBalanceSum <> dbl退款合计 Then bln原样退 = False
    End If
    
    txt退款合计.Text = Format(dbl退款合计, "0.00")
    If bln原样退 Then
        '可能存在多单据中将误差放在最后一张单据,造成单张退费时,现金存在误差项
         If mintInsure > 0 Then
            If gclsInsure.GetCapability(support分币处理, mlng病人ID, mintInsure) Then
                cur单据合计 = CentMoney(dblCashSum)
            Else
                cur单据合计 = Format(dblCashSum, "0.00")
            End If
        Else
            cur单据合计 = CentMoney(dblCashSum)
        End If
        If cur单据合计 <> dblCashSum Then bln原样退 = False
    End If
    
    If bln原样退 Then
        zlControl.CboSetIndex cbo退款方式.hWnd, mintReturnMode
    End If
    cbo退款方式.Enabled = Not bln原样退
    cbo退款方式.Locked = bln原样退
    '2.计算退款金额及误差
    If cbo退款方式.ListIndex <> -1 Then
        If cbo退款方式.ItemData(cbo退款方式.ListIndex) = 1 Then
            bln现金结算 = True
        End If
    End If
    Dim varTemp As Variant
    
    For i = 1 To colAllReturn.Count
        '0-是否完全退费;1-NO,2-单据合计,3-选择合计
        varTemp = colAllReturn(i)
        strNo = varTemp(1)
        cur单据合计 = Val(varTemp(2)): cur选择合计 = Val(varTemp(3))
        cur退费合计 = 0: cur误差金额 = 0
        '完全退费时排开医保结算及冲预交金额
        bln完全退费 = IIf(Val(varTemp(0)) = 1, True, False)
        
        If bln完全退费 Or bln全退 Then
            mrsBalance.Filter = "NO='" & strNo & "'"
            '性质:1-预存款,2-医保,3-医疗卡,4-结算卡,5-一卡通,0-其他类
            With mrsBalance
                Do While Not .EOF
                    Select Case Val(Nvl(!性质))
                    Case 1 '预交款
                         cur选择合计 = cur选择合计 - Nvl(!结算金额, 0)
                    Case 2 '医保
                        '如果这种结算方式不支持回退,要退为现金,则不用减去
                        If mblnYB结算作废 Then
                            If gclsInsure.GetCapability(support门诊结算作废, mlng病人ID, mintInsure, !结算方式) Then
                                cur选择合计 = cur选择合计 - Nvl(!结算金额, 0)
                            End If
                        Else     '不支持门诊结算作废时,只允许个帐退为现金,其它原样退,不调用医保交易
                            If !结算方式 <> mstr个人帐户 Then
                                cur选择合计 = cur选择合计 - Nvl(!结算金额, 0)
                            End If
                        End If
                    Case 3, 4 '医疗卡和结算卡
                            If Val(Nvl(!是否退现)) = 0 Then
                                cur选择合计 = cur选择合计 - Nvl(!结算金额, 0)
                            End If
                            If Val(Nvl(!是否退现)) = 1 Then
                                For intCol = 1 To vsBalance.COLS - 1 Step 2
                                    If vsBalance.Cell(flexcpData, 1, intCol) = !结算方式 Then
                                        If Val(vsBalance.TextMatrix(1, intCol + 1)) <> 0 And vsBalance.RowHidden(1) = False Then
                                            cur选择合计 = cur选择合计 - Nvl(!结算金额, 0)
                                        End If
                                    End If
                                Next
                            End If
                    Case 5 '一卡通
                            cur选择合计 = cur选择合计 - Nvl(!结算金额, 0)
                    Case Else
                        blnFind = False
                        For intCol = 1 To vsBalance.COLS - 1 Step 2
                            If vsBalance.Cell(flexcpData, 1, intCol) = zlStr.NeedName(cbo退款方式.Text) _
                                And vsBalance.Cell(flexcpData, 1, intCol) = !结算方式 Then
                                If bln退现 Or Not bln原样退 Then    'Or Not bln原样退:55187
                                    vsBalance.TextMatrix(1, intCol) = ""
                                    vsBalance.TextMatrix(1, intCol + 1) = ""
                                ElseIf vsBalance.TextMatrix(1, intCol) = "" Then
                                    vsBalance.TextMatrix(1, intCol) = vsBalance.Cell(flexcpData, 1, intCol) & ":"
                                    vsBalance.TextMatrix(1, intCol + 1) = vsBalance.Cell(flexcpData, 1, intCol + 1)
                                End If
                                blnFind = True: Exit For
                            End If
                        Next
                        
                        blnFind = False
                        For intCol = 1 To vsBalance.COLS - 1 Step 2
                            If zlStr.NeedName(cbo退款方式.Text) <> vsBalance.Cell(flexcpData, 1, intCol) _
                                And vsBalance.TextMatrix(1, intCol) = "" _
                                And vsBalance.Cell(flexcpData, 1, intCol) <> "" Then
                                    vsBalance.TextMatrix(1, intCol) = vsBalance.Cell(flexcpData, 1, intCol) & ":"
                                    vsBalance.TextMatrix(1, intCol + 1) = vsBalance.Cell(flexcpData, 1, intCol + 1)
                            End If
                            If vsBalance.Cell(flexcpData, 1, intCol) = !结算方式 And vsBalance.TextMatrix(1, intCol) <> "" Then
                                blnFind = True: Exit For
                            End If
                        Next
                        If bln全退 And blnFind Then
                            cur选择合计 = cur选择合计 - Nvl(!结算金额, 0)
                        End If

                    End Select
                    .MoveNext
                Loop
            End With
        Else
            '部分退费，检查是否部分退费
            mrsBalance.Filter = "NO='" & strNo & "'  and 性质<>1 "
            If mrsBalance.RecordCount = 0 Then
                cur选择合计 = 0
            Else
              
'                '可能存在预交及三方卡的退费,因此,需要排除该数据
'                If vsBalance.RowHidden(1) = False Then
'                    For intCol = 1 To vsBalance.COLS - 1 Step 2
'                        If vsBalance.Cell(flexcpData, 1, intCol) <> "" Then
'                            cur选择合计 = cur选择合计 - Val(vsBalance.TextMatrix(1, intCol + 1))
'                        End If
'                    Next
'                End If
                If mCurBillType.bln存在医疗卡结算 And mCurBillType.blnSingleBalance Then
                    vsBalance.Cell(flexcpData, 1, 2) = IIf(blnHaveSelected, dbl退款合计, "")
                    If vsBalance.TextMatrix(1, 2) = "" And mCurBillType.bln三方卡全退 = False Then
                        vsBalance.TextMatrix(1, 2) = IIf(blnHaveSelected, FormatEx(dbl退款合计, 2), "")
                    End If
                    If Val(vsBalance.TextMatrix(1, 2)) = 0 Or vsBalance.RowHidden(1) Then
                        cbo退款方式.Enabled = True
                        cbo退款方式.Locked = False
                    Else
                        vsBalance.TextMatrix(1, 2) = FormatEx(dbl退款合计, 2)
                        cbo退款方式.Enabled = False
                        cbo退款方式.Locked = True
                        cur选择合计 = 0 ' cur选择合计 - dbl退款合计
                    End If
                    If cbo退款方式.Visible And cbo退款方式.ListIndex <> -1 Then
                        If cbo退款方式.ItemData(cbo退款方式.ListIndex) = 1 Then
                            bln现金结算 = True
                        End If
                    End If
                End If
            End If
            
        End If
        
        '费用金额保留位数,及现金结算时处理分币
        If bln现金结算 Then
            If mintInsure > 0 Then
                If gclsInsure.GetCapability(support分币处理, mlng病人ID, mintInsure) Then
                    cur退费合计 = CentMoney(cur选择合计)
                Else
                    cur退费合计 = Format(cur选择合计, "0.00")
                End If
            Else
                cur退费合计 = CentMoney(cur选择合计)
            End If
        Else
            cur退费合计 = Format(cur选择合计, "0.00")
        End If
        
        '误差金额,部分退,或医保全退时因为结算方式不支持回退而退为现金,可能产生误差
        '非现金结算时,也可能有误差,这个误差是费用金额保留位数引起的
        If Not bln原样退 Then
            cur误差金额 = cur退费合计 - cur选择合计
        End If
        
        curTotal = curTotal + cur退费合计
        mcolError.Add cur误差金额, "_" & strNo
        cur误差合计 = cur误差合计 + cur误差金额
    Next
    txt退款金额.ToolTipText = "退费误差金额:" & Format(cur误差合计, gstrDec)
    
    txt退款金额.Text = Format(curTotal, "0.00")
    vsBalance.AutoSizeMode = flexAutoSizeColWidth
    Call vsBalance.AutoSize(0, vsBalance.COLS - 1)
    
GoSetVisible:
        Call Show退款方式(cbo退款方式.Enabled And curTotal <> 0)
End Sub

Private Sub ControlResize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调整控件位置
    '编制:刘兴洪
    '日期:2011-11-23 14:21:25
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim blnFind As Boolean
    With vsBalance
        For i = 1 To .COLS - 1 Step 2
            If .Cell(flexcpData, 1, i) <> "" Then
                blnFind = True: Exit For
            End If
        Next
        If blnFind = False Then .RowHidden(1) = True
        .Height = IIf(.RowHidden(1), 375, 735)
    End With
    Form_Resize
End Sub

 
 
Private Sub txtPatient_Change()
    '问题:50885
    If Me.ActiveControl Is txtPatient Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
        IDKind.SetAutoReadCard (txtPatient.Text = "")
    End If
End Sub
Private Sub txtPatient_GotFocus()
    '问题:50885
    If txtPatient.Locked Or Not txtPatient.Visible Then Exit Sub
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtPatient.Text = "")
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (txtPatient.Text = "")
    IDKind.SetAutoReadCard (txtPatient.Text = "")
End Sub
Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean, blnICCard As Boolean, blnCancel As Boolean
    '问题:50885
    
    On Error GoTo errH
    If txtPatient.Locked Then Exit Sub
     
    If IDKind.GetCurCard.名称 Like "姓名*" Then
        '103563,只要输入的第一个字符是“-+*”，后面是全数字，都认为不是刷卡
        If Not (InStr("-+*", Left(txtPatient.Text, 1)) > 0 And IsNumeric(Mid(txtPatient.Text, 2))) Then
            blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, IDKind.ShowPassText)
        End If
    ElseIf IDKind.GetCurCard.名称 = "门诊号" Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If InStr("0123456789-*+", Chr(KeyAscii)) = 0 Then KeyAscii = 0: Exit Sub
        End If
    Else
        txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
        '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
        txtPatient.IMEMode = 0
    End If
 
    If Not (blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 _
        Or KeyAscii = 13 And Trim(txtPatient.Text) <> "") Then Exit Sub
        
    If KeyAscii <> 13 Then
        txtPatient.Text = txtPatient.Text & Chr(KeyAscii): txtPatient.SelStart = Len(txtPatient.Text)
    End If
    KeyAscii = 0
    Call FindPati(IDKind.GetCurCard, blnCard, Trim(txtPatient.Text))
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub FindPati(ByVal objCard As Card, ByVal blnCard As Boolean, ByVal strInput As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:查找病人
    '编制:刘兴洪
    '日期:2012-09-03 09:46:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnICCard As Boolean, blnCancel As Boolean
    
    If objCard.名称 Like "IC卡*" And objCard.系统 Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    'a.根据输入读取病人信息失败
    If Not GetPatient(objCard, Trim(txtPatient.Text), blnCancel, blnCard) Then
        If blnCancel Then '取消输入
            If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
            txtPatient.Text = ""
            Exit Sub
        End If
        stbThis.Panels(2) = "未找到该病人，请检查输入内容!"
        If blnCard = True Then
            txtPatient.PasswordChar = "": txtPatient.Text = ""
            '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
            txtPatient.IMEMode = 0
        Else
            txtPatient.SelStart = 0: txtPatient.SelLength = Len(txtPatient.Text)
        End If
        Set mrsInfo = New ADODB.Recordset
        If txtPatient.Visible And txtPatient.Enabled Then txtPatient.SetFocus
        Exit Sub
    End If
    mlng病人ID = Val("" & mrsInfo!病人ID)
    txtPatient = Nvl(mrsInfo!姓名)

    lblPati.Caption = "病人:" & "                 " & _
        "　性别:" & Nvl(mrsInfo!性别) & _
        "　年龄:" & Nvl(mrsInfo!年龄) & _
        "　门诊号:" & Nvl(mrsInfo!门诊号) & _
        "　费别:" & Nvl(mrsInfo!费别) & _
        "　付款方式:" & mrsInfo!医疗付款方式
    With mtyPati
        .病人ID = mlng病人ID
        .性别 = Nvl(mrsInfo!性别)
        .年龄 = Nvl(mrsInfo!年龄)
        .姓名 = Nvl(mrsInfo!姓名)
    End With
    If SelectNO(mlng病人ID) = False Then Exit Sub
    If vsBill.Enabled And vsBill.Visible Then vsBill.SetFocus
End Sub

Private Function GetPatient(ByVal objCard As Card, ByVal strInput As String, blnCancel As Boolean, Optional blnCard As Boolean = False) As Boolean
    '功能：读取病人信息
    '参数：strInput=[刷卡]|[A病人ID]|[B住院号]
    '说明：
    '     1.适用于病人预交款
    '     2.自动识别病人在院状态,读出(病人ID,主页ID,姓名,性别,年龄,住院号,床号,在院标志)
    '返回:是否读取成功,成功时mrsInfo中包含病人信息,失败时mrsInfo=Close
    '问题:50885
    Dim rsTmp As ADODB.Recordset, strPati As String, strSQL As String
    Dim vRect As RECT, i As Integer, lng卡类别ID As Long, bln存在帐户 As Boolean, lng病人ID As Long, strPassWord As String, strErrMsg As String
    Dim strWhere As String, blnICCard As Boolean
    Dim blnHavePassWord As Boolean
    blnCancel = False
    strWhere = ""
    If blnCard And objCard.名称 Like "姓名*" And InStr("-+*", Left(strInput, 1)) = 0 Then  '103563
        lng卡类别ID = IDKind.GetDefaultCardTypeID
        '短名|完成名|刷卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户);…
        If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg, lng卡类别ID) = False Then GoTo NotFoundPati:
        If lng病人ID <= 0 Then GoTo NotFoundPati:
        strWhere = strWhere & " And A.病人ID=[1]"
        strInput = "-" & lng病人ID
        blnHavePassWord = True
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then  '病人ID
        strWhere = strWhere & " And A.病人ID=[1]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then  '住院号(对住(过)院的病人)
        strWhere = strWhere & " And A.住院号=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '门诊号(仅对门诊病人)
        strWhere = strWhere & " And A.门诊号=[1]"
        '75087,冉俊明,2014-7-29,门诊病人收费时,不需要输入完整的门诊号,只需要输入门诊号的最后顺序号即能找到当天就诊的病人信息、费用
        strInput = "*" & zlCommFun.GetFullNO(Mid(strInput, 2), 3)
    Else '当作姓名
        Select Case objCard.名称
            Case "姓名", "姓名或就诊卡"
                strPati = _
                " Select /*+Rule */A.病人ID as ID,A.病人ID,A.姓名,A.性别,A.年龄," & _
                "           A.住院号,B.名称 as 科室,A.当前床号 as 床号," & _
                "           A.出生日期,A.身份证号,A.家庭地址,A.卡验证码 " & _
                " From 病人信息 A,部门表 B" & _
                " Where A.停用时间 is NULL And A.当前科室ID=B.ID(+) And A.姓名 Like [1]" & _
                "   Order by A.姓名"
                vRect = zlControl.GetControlRect(txtPatient.hWnd)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "病人查找", 1, "", "请选择病人", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput & "%", "bytSize=1")
                If Not rsTmp Is Nothing Then
                    strInput = rsTmp!病人ID
                    strWhere = strWhere & " And A.病人ID=[2]"
                Else
                    Set mrsInfo = New ADODB.Recordset: Exit Function
                End If
            Case "医保号"
                strInput = UCase(strInput)
                strWhere = strWhere & " And A.医保号=[2]"
            Case "身份证号", "二代身份证", "身份证"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("身份证", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                strInput = "-" & lng病人ID
                strWhere = strWhere & " And A.病人ID=[1]"
            Case "IC卡号", "IC卡"
                strInput = UCase(strInput)
                If gobjSquare.objSquareCard.zlGetPatiID("IC卡", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                strInput = "-" & lng病人ID
                strWhere = strWhere & " And A.病人ID=[1]"
                blnICCard = (InStr(1, "-+*.", Left(strInput, 1)) = 0)
            Case "门诊号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strWhere = strWhere & " And A.门诊号=[2]"
                '75087,冉俊明,2014-7-29,门诊病人收费时,不需要输入完整的门诊号,只需要输入门诊号的最后顺序号即能找到当天就诊的病人信息、费用
                strInput = zlCommFun.GetFullNO(strInput, 3)
            Case "住院号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strWhere = strWhere & " And A.住院号=[2]"
            Case Else
                '其他类别的,获取相关的病人ID
                If objCard.接口序号 > 0 Then
                    lng卡类别ID = objCard.接口序号
                    bln存在帐户 = objCard.是否存在帐户 = 1
                    If gobjSquare.objSquareCard.zlGetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    If lng病人ID = 0 Then GoTo NotFoundPati:
                Else
                    If gobjSquare.objSquareCard.zlGetPatiID(objCard.名称, strInput, False, lng病人ID, _
                        strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                End If
                If lng病人ID <= 0 Then GoTo NotFoundPati:
                strWhere = strWhere & " And A.病人ID=[1]"
                strInput = "-" & lng病人ID
                blnHavePassWord = True
        End Select
    End If
    '76451,冉俊明,2014-8-19
    strSQL = _
    " Select A.病人ID,Nvl(C.主页ID,0) as 主页ID,A.门诊号,Nvl(C.当前病区ID,0) as 病区ID,Nvl(c.出院科室ID,0) as 科室ID,Nvl(A.当前科室ID,0) as 当前科室ID, Nvl(a.在院,0) as 在院," & _
    "           Decode(Nvl(A.主页ID,0),0,A.医疗付款方式,C.医疗付款方式) 医疗付款方式,Nvl(A.病人类型,C.病人类型) as 病人类型," & _
    "           A.姓名,A.性别,A.年龄,Nvl(A.住院号,0) as 住院号,Nvl(C.出院病床,0) as 床号,A.家庭地址,A.卡验证码," & _
    "           B.险类,B.卡号,Nvl(B.医保号,A.医保号) 医保号,B.密码,Nvl(C.费别,A.费别) 费别,A.担保人,A.担保额,Nvl(A.担保性质,0) as 担保性质, C.备注 " & _
    " From 病人信息 A,医保病人档案 B,病案主页 C,医保病人关联表 E " & _
    " Where A.停用时间 is NULL" & _
    "       And A.病人ID=C.病人ID(+) And Nvl(A.主页ID,0)=C.主页ID(+)" & _
    "       And C.病人ID=E.病人ID(+) And E.标志(+)=1  " & _
    "       And E.医保号=B.医保号(+) And E.险类=B.险类(+) And E.中心 = B.中心(+) " & strWhere
    
    On Error GoTo errH
    '75259：李南春,2014-7-10，病人姓名颜色处理
    txtPatient.ForeColor = &HC00000
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
    If mrsInfo.EOF Then
        Set mrsInfo = New ADODB.Recordset: Exit Function
    End If
    Call SetPatiColor(txtPatient, Nvl(mrsInfo!病人类型), IIf(IsNull(mrsInfo!险类), &HC00000, vbRed))
    GetPatient = True
    Exit Function
errH:
     If ErrCenter() = 1 Then Resume
    Call SaveErrLog
NotFoundPati:
    Set mrsInfo = New ADODB.Recordset
End Function
Private Function SelectNO(ByVal lng病人ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据病人ID选择合适的退费单据
    '入参:
    '出参:
    '返回: 成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2012-07-04 10:32:40
    '问题:50885
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, blnCancel As Boolean
    On Error GoTo errHandle
    '问题号:50885
    strSQL = "" & _
        "  With 收费单 as ( " & _
        "           Select Max(a.ID) as ID,a.No as 单据号,  B.名称 as 开单部门, a.开单人, a.操作员编号, a.操作员姓名, a.实际票号, To_Char(a.发生时间, 'yyyy-mm-dd hh24:mi:ss') As 发生时间, " & vbCrLf & _
        "                   To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间 " & vbCrLf & _
        "           From 门诊费用记录 A,部门表 B " & vbCrLf & _
        "           Where a.记录性质 = 1 And nvl(A.附加标志,0)<>9 and A.开单部门ID=B.ID(+) And a.记录状态 =1 " & vbCrLf & _
        "                       And Nvl(a.执行状态, 0) <> 1 And Nvl(a.费用状态, 0) <> 1 And a.病人id = [1] " & vbCrLf & _
        "          Group by   a.No,  a.开单人, B.名称,a.操作员编号, a.操作员姓名, a.实际票号, To_Char(a.发生时间, 'yyyy-mm-dd hh24:mi:ss'),To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') " & vbCrLf & _
        "           )"
        
     strSQL = strSQL & vbCrLf & _
     "  Select J.*  " & vbCrLf & _
     "  From 收费单 J," & vbCrLf & _
     "           (Select A.NO,sum(nvl(A.付数,1)*nvl(A.数次,1)) 数次" & vbCrLf & _
     "             From 门诊费用记录 A,收费单 B  " & vbCrLf & _
     "             Where A.NO=B.单据号 And A.记录性质=1 And a.价格父号 is null  " & vbCrLf & _
     "             Group by A.NO " & vbCrLf & _
     "              Having sum(nvl(A.付数,1)*nvl(A.数次,1))>0 ) M" & vbCrLf & _
     "  Where J.单据号=M.NO " & vbCrLf
     
     strSQL = "Select * From (" & strSQL & ") Order by 登记时间 desc,单据号"
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "退费单据", 1, "", "请选择需要退费的单据", False, False, False, 0, 0, 0, blnCancel, False, False, lng病人ID, "bytSize=1")
    If blnCancel = True Then Exit Function
    If rsTemp Is Nothing Then Exit Function
    If rsTemp.State <> 1 Then Exit Function
    
    Dim strNo As String
    mstrNo = Nvl(rsTemp!单据号)
    mblnOneCard = GetOneCard.RecordCount > 0
    If Not ReadBills(mstrNo) Then
        ClearFace True, True
        Exit Function
    End If
    SelectNO = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub mobjICCard_ShowICCardInfo(ByVal strNo As String)
    '问题:50885
    Dim lngPreIDKind As Long
    If Not txtPatient.Locked And txtPatient.Text = "" And Me.ActiveControl Is txtPatient And strNo <> "" Then
        lngPreIDKind = IDKind.IDKind
        mblnNotClick = True
        Dim intIndex As Integer
        intIndex = IDKind.GetKindIndex("IC卡号")
        If intIndex <= 0 Then mblnNotClick = False: Exit Sub
        IDKind.IDKind = intIndex: Call IDKind_Click(IDKind.GetCurCard)
        txtPatient.Text = strNo
        Call txtPatient_KeyPress(vbKeyReturn)
        If txtPatient.Text = "" Then Call mobjICCard.SetEnabled(False)
        IDKind.IDKind = lngPreIDKind
        mblnNotClick = False
        If Not txtPatient.Locked And Me.ActiveControl Is txtPatient Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
    End If
End Sub
Private Sub mobjIDCard_ShowIDCardInfo(ByVal strID As String, ByVal strName As String, ByVal strSex As String, _
                            ByVal strNation As String, ByVal datBirthday As Date, ByVal strAddress As String)
    Dim lngPreIDKind As Long
    '问题:50885
    If Not txtPatient.Locked And txtPatient.Text = "" And Me.ActiveControl Is txtPatient Then
        mblnNotClick = True
        lngPreIDKind = IDKind.IDKind
        IDKind.IDKind = IDKind.GetKindIndex("身份证号")
        txtPatient.Text = strID
        Call txtPatient_KeyPress(vbKeyReturn)
        IDKind.IDKind = lngPreIDKind
        mblnNotClick = False
    End If
End Sub

Private Sub CloseIDCard()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:关闭自助读卡功能
    '编制:刘兴洪
    '日期:2012-03-09 16:26:40
    '问题:50885
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not mobjIDCard Is Nothing Then
        mobjIDCard.SetEnabled (False)
        Set mobjIDCard = Nothing
    End If
    If Not mobjICCard Is Nothing Then
        mobjICCard.SetEnabled (False)
        Set mobjICCard = Nothing
    End If
End Sub
Private Sub NewCardObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化新的卡对象
    '编制:刘兴洪
    '日期:2012-03-09 16:28:23
    '问题:50885
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mbytMode = 0 Then Exit Sub
    
    If mobjIDCard Is Nothing Then
        Set mobjIDCard = New clsIDCard
        Call mobjIDCard.SetParent(Me.hWnd)
    End If
    If mobjICCard Is Nothing Then
        Set mobjICCard = New clsICCard
        Call mobjICCard.SetParent(Me.hWnd)
    End If
End Sub

Private Sub SetInvoceSizeAndShowTittle()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调整发票显示控件的大小和显示
    '编制:刘兴洪
    '日期:2013-05-07 16:14:02
    '问题:25187
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllInvoice As New Collection
    Dim r As Long, c As Long
    Dim bytSel As Byte '1-选择;2-不选择,3-不能取消的选择(关联发票)
    Dim strInvoice As String '发票号
    Dim sngColWidth As Single
    Dim i As Long
    Err = 0: On Error GoTo Errhand:
    Set cllInvoice = New Collection
    With vsInvoice
        If .Rows = 1 And .Cell(flexcpLeft, 0, .COLS - 1) + .ColWidth(.COLS - 1) <= .Width Then Exit Sub
        For r = 0 To .Rows - 1
            For c = 1 To .COLS - 1
                bytSel = .Cell(flexcpChecked, r, c)
                strInvoice = Trim(.Cell(flexcpData, r, c))
                sngColWidth = .ColWidth(c)
                If strInvoice <> "" Then
                    cllInvoice.Add Array(bytSel, strInvoice, sngColWidth)
                End If
            Next
        Next
        .Redraw = flexRDNone
        .Rows = 1
        .COLS = 1
        .Clear
        .TextMatrix(0, 0) = "发票号"
        sngColWidth = .ColWidth(0)
        For i = 1 To cllInvoice.Count
            If sngColWidth + cllInvoice(i)(2) * 0.5 > .Width Then
                If .COLS <= 1 Then
                    .COLS = .COLS + 1
                    .ColWidth(.COLS - 1) = cllInvoice(i)(2)
                End If
                Exit For
            End If
            .COLS = .COLS + 1
            .ColWidth(.COLS - 1) = cllInvoice(i)(2)
            sngColWidth = sngColWidth + .ColWidth(.COLS - 1)
        Next
        .Cell(flexcpChecked, 0, .COLS - 1, .Rows - 1, .COLS - 1) = 0
        c = 0: r = 0
        For i = 1 To cllInvoice.Count
            If c >= .COLS - 1 Then
                .Rows = .Rows + 1
                r = r + 1
                c = 1
            Else
                c = c + 1
            End If
            .TextMatrix(r, c) = cllInvoice(i)(1)
            .Cell(flexcpData, r, c) = cllInvoice(i)(1)
            .Cell(flexcpChecked, r, c) = cllInvoice(i)(0)
            .ColWidth(c) = cllInvoice(i)(2)
        Next
        .Height = (.RowHeight(0) + 90) * (.Rows)
        Call MergeFixedCol
        .Redraw = flexRDBuffered
    End With
    Exit Sub
Errhand:
    vsInvoice.Redraw = flexRDBuffered
End Sub
Private Sub SetpicInvoiceVisible()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置发票控件的显示
    '编制:刘兴洪
    '日期:2013-05-09 11:30:35
    '---------------------------------------------------------------------------------------------------------------------------------------------
    picInvoice.Visible = False
    If gTy_Module_Para.byt票据分配规则 = 0 Then GoTo ReSizing:
    If mbytMode <> 1 Then GoTo ReSizing:
    If mrsDelInvoice Is Nothing Then GoTo ReSizing:
    mrsDelInvoice.Filter = 0
    If mrsDelInvoice.RecordCount = 0 Then GoTo ReSizing:
    picInvoice.Visible = True
ReSizing:
    '重新调整大小
    Form_Resize
    picInvoice_Resize
    picPati_Resize
End Sub
Private Sub LoadInvoiceData(ByVal strNos As String, Optional ByVal strInvoiceNO As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载发票信息
    '入参:strNos-单据号,多个用逗号分隔
    '       strInvoiceNo-发票号(按指定的发票号发票号查找)
    '编制:刘兴洪
    '日期:2013-05-07 17:07:38
    '问题:25187
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str序号 As String, varTemp As Variant
    Dim i As Long, str发票号 As String
    If gTy_Module_Para.byt票据分配规则 = 0 Then Exit Sub
    If mbytMode <> 1 Then Exit Sub
    If mrsDelInvoice Is Nothing Then
        Set mrsDelInvoice = zlGetFromNoTOInvoice(strNos)
    End If
    If mrsDelInvoice Is Nothing Then Exit Sub
    If mrsDelInvoice.RecordCount = 0 Then Exit Sub
    mrsDelInvoice.Sort = "票号"
    str发票号 = ""
    With mrsDelInvoice
        Do While Not .EOF
            If InStr(str发票号 & ",", "," & Nvl(!票号) & ",") = 0 Then
                str发票号 = str发票号 & "," & Nvl(!票号)
            End If
            .MoveNext
        Loop
    End With
    If str发票号 <> "" Then str发票号 = Mid(str发票号, 2)
      '加载发票号
    varTemp = Split(str发票号, ",")
    With vsInvoice
        .Clear
        .Rows = 1: .COLS = 1
        .FixedCols = 1
        .TextMatrix(0, 0) = "发票号"
        .Redraw = flexRDNone
        .COLS = .COLS + UBound(varTemp) + 1
        For i = 0 To UBound(varTemp)
            If i + 1 > .COLS - 1 Then
                .COLS = .COLS + 1
            End If
            .TextMatrix(0, i + 1) = CStr(varTemp(i))
            .Cell(flexcpData, 0, i + 1) = CStr(varTemp(i))
            .Cell(flexcpChecked, 0, i + 1) = 2
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .COLS - 1)
        Call picInvoice_Resize
        Call Form_Resize
        
        .Editable = flexEDKbdMouse
        .Redraw = flexRDBuffered
    End With
End Sub
Private Sub FromNoSelectInvoice()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据单据选择发票
    '编制:刘兴洪
    '日期:2013-05-08 15:52:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str发票号 As String, strNo As String
    Dim strNos As String, i As Long, j As Long
    If mbytMode <> 1 Or gTy_Module_Para.byt票据分配规则 = 0 Then Exit Sub
    
    On Error GoTo errHandle
    With vsBill
        str发票号 = ""
        For i = 1 To .Rows - 1
              If Abs(Val(.TextMatrix(i, .ColIndex("选择")))) = 1 Then
                    strNo = .TextMatrix(i, .ColIndex("单据号"))
                    If strNo <> "" Then
                        str发票号 = str发票号 & "," & GetFromNumToInvoiceNo(strNo, CStr(.RowData(i)))
                    End If
              End If
        Next
    End With
    With vsInvoice
        For i = 0 To .Rows - 1
            For j = 1 To .COLS - 1
                If InStr(1, str发票号 & ",", "," & .Cell(flexcpData, i, j) & ",") > 0 Then
                    .Cell(flexcpChecked, i, j) = 1
                ElseIf Trim(.Cell(flexcpData, i, j)) <> "" Then
                    .Cell(flexcpChecked, i, j) = 2
                Else
                End If
            Next
        Next
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function GetFromNumToInvoiceNo(ByVal strNo As String, ByVal str序号 As String) As String
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据序号获取对应的发票号
    '入参:strNO-单据号
    '       str序号-序号,可以为多个,多个用逗号分离
    '       strNotInvoice-不包含的发票号
    '返回:成功返回true,否则返回False
    '编制:刘兴洪
    '日期:2013-05-07 17:38:24
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str发票号 As String, str关联序号 As String
    Dim varTemp As Variant, i As Long, strTemp As String
    On Error GoTo errHandle
    If mrsDelInvoice Is Nothing Then Exit Function
    If mrsDelInvoice.State <> 1 Then Exit Function
    If mrsDelInvoice.RecordCount = 0 Then Exit Function
    With mrsDelInvoice
        str关联序号 = "": str发票号 = ""
        varTemp = Split(str序号, ",")
        .Filter = "NO='" & strNo & "'"
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
                strTemp = "," & Nvl(!序号) & ","
                For i = 0 To UBound(varTemp)
                    If InStr(1, strTemp, "," & varTemp(i) & ",") > 0 _
                        And InStr(str发票号 & ",", "," & Nvl(!票号) & ",") = 0 Then
                        str发票号 = str发票号 & "," & Nvl(!票号)
                        If Val(Nvl(!关联票号序号)) <> 0 Then
                            str关联序号 = str关联序号 & "," & Val(Nvl(!关联票号序号))
                        End If
                    End If
                Next
            .MoveNext
        Loop
        .Filter = 0: .MoveFirst
        If str关联序号 = "" Then GoTo GoSort:
        '需要查找关联票号
       varTemp = Split(Mid(str关联序号, 2), ",")
        Do While Not .EOF
                For i = 0 To UBound(varTemp)
                    If Val(varTemp(i)) = Val(Nvl(!关联票号序号)) _
                        And InStr(str发票号 & ",", "," & Nvl(!票号) & ",") = 0 Then
                        str发票号 = str发票号 & "," & Nvl(!票号)
                    End If
                Next
            .MoveNext
        Loop
    End With
    '进行排序处理
GoSort:
    If str发票号 = "" Then Exit Function
    str发票号 = Mid(str发票号, 2)
    GetFromNumToInvoiceNo = zlStringSort(str发票号)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub MergeFixedCol()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:合并固定列
    '编制:刘兴洪
    '日期:2013-05-08 15:38:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, c As Long
    On Error GoTo errHandle
    If mbytMode <> 1 Or gTy_Module_Para.byt票据分配规则 = 0 Then Exit Sub
    With vsInvoice
        If .FixedCols = 0 Then Exit Sub
        For i = 0 To .Rows - 1
            .MergeRow(c) = True
            For c = 0 To .FixedCols - 1
                .TextMatrix(i, c) = "发票号"
            Next
        Next
        .MergeCellsFixed = flexMergeRestrictRows
        For c = 0 To .FixedCols - 1
            .MergeCol(c) = True
        Next
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub vsInvoice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strInvoice As String
    With vsInvoice
        strInvoice = Trim(.Cell(flexcpData, Row, Col))
        If strInvoice <> "" Then
            '同时选择关联发票
            Call SelectRelatingInvoice(strInvoice, Abs(Val(.Cell(flexcpChecked, Row, Col))) = 1)
        End If
    End With
    Call FromAllInvoiceSelectNO
    '主要选择中关联收回的发票
    Call FromNoSelectInvoice
    Call ShowAndHideDelBillRow
End Sub

Private Sub vsInvoice_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = 0 Then Cancel = True: Exit Sub
    With vsInvoice
'        Select Case Val(.Cell(flexcpChecked, Row, Col))
'        Case 3
'            Cancel = True
'        End Select
    End With
End Sub

Private Sub SelectRelatingInvoice(ByVal strInvoiceNO As String, ByVal blnSel As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:选中指定发票的关联发票
    '入参:strInvoiceNo-发票号
    '编制:刘兴洪
    '日期:2013-05-09 10:41:45
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim r As Long, c As Long, lng关联序号 As Long
    Dim str发票号 As String
    On Error GoTo errHandle
    If mbytMode <> 1 Or gTy_Module_Para.byt票据分配规则 = 0 Then Exit Sub
    '选择本身发票
    With vsInvoice
        For r = 0 To .Rows - 1
            For c = 1 To .COLS - 1
                str发票号 = .Cell(flexcpData, r, c)
                If str发票号 = strInvoiceNO Then
                        .Cell(flexcpChecked, r, c) = IIf(blnSel, 1, 2)
                End If
            Next
        Next
    End With
    If mrsDelInvoice Is Nothing Then Exit Sub
    If mrsDelInvoice.RecordCount = 0 Then Exit Sub
    lng关联序号 = 0
    mrsDelInvoice.Filter = "票号='" & strInvoiceNO & "'"
    If mrsDelInvoice.RecordCount <> 0 Then
        lng关联序号 = Val(Nvl(mrsDelInvoice!关联票号序号))
    End If
    If lng关联序号 = 0 Then
        mrsDelInvoice.Filter = 0: Exit Sub
    End If
    mrsDelInvoice.Filter = "关联票号序号=" & lng关联序号
    If mrsDelInvoice.RecordCount = 0 Then
        mrsDelInvoice.Filter = 0: Exit Sub
    End If
    
    With mrsDelInvoice
        .MoveFirst
        Do While Not .EOF
            With vsInvoice
                For r = 0 To .Rows - 1
                    For c = 1 To .COLS - 1
                        str发票号 = .Cell(flexcpData, r, c)
                        If str发票号 = Nvl(mrsDelInvoice!票号) Or str发票号 = strInvoiceNO Then
                            .Cell(flexcpChecked, r, c) = IIf(blnSel, 1, 2)
                        End If
                    Next
                Next
            End With
            .MoveNext
        Loop
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub FromInvoiceSelectNO(ByVal strInvoiceNO As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据指定的发票号,选择具体的单据
    '编制:刘兴洪
    '日期:2013-05-08 16:23:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNo As String, str序号 As String
    Dim k As Long, j As Long
    If mbytMode <> 1 Or gTy_Module_Para.byt票据分配规则 = 0 Then Exit Sub
    If mrsDelInvoice Is Nothing Then Exit Sub
    If mrsDelInvoice.RecordCount = 0 Then Exit Sub
    On Error GoTo errHandle
    mrsDelInvoice.Filter = "票号='" & strInvoiceNO & "'"
    
    With mrsDelInvoice
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            strNo = Nvl(!NO): str序号 = "," & Nvl(!序号) & ","
              With vsBill
                  k = vsBill.FindRow(strNo, , .ColIndex("单据号"))
                  For j = k To .Rows - 1
                      If .TextMatrix(j, .ColIndex("单据号")) <> strNo Then Exit For
                      If InStr(1, str序号, "," & .RowData(j) & ",") > 0 Then
                            .Cell(flexcpChecked, j, .ColIndex("选择")) = 1
                      End If
                      '同步选择相关组合项目
                      Call SynchronizationSelect(j)
                  Next
              End With
             .MoveNext
        Loop
    End With
    mrsDelInvoice.Filter = 0
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
   mrsDelInvoice.Filter = 0
End Sub
Private Sub FromAllInvoiceSelectNO()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据所有选择的发票号,选择具体的单据
    '编制:刘兴洪
    '日期:2013-05-08 16:23:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str发票号 As String, i As Long, c As Long, k As Long, j As Long
    Dim strNo As String, str序号 As String
    On Error GoTo errHandle
    If mrsDelInvoice Is Nothing Then Exit Sub
    If mrsDelInvoice.RecordCount = 0 Then Exit Sub
    If mbytMode <> 1 Or gTy_Module_Para.byt票据分配规则 = 0 Then Exit Sub
    
    With vsBill
        .Cell(flexcpChecked, .FixedRows, .ColIndex("选择"), .Rows - 1, .ColIndex("选择")) = 0
    End With
    With vsInvoice
        For i = 0 To .Rows - 1
            For c = 1 To .COLS - 1
                str发票号 = Trim(.Cell(flexcpData, i, c))
                If Abs(Val(.Cell(flexcpChecked, i, c))) = 1 Then
                    Call FromInvoiceSelectNO(str发票号)
                End If
            Next
        Next
    End With
    '显示相关的结算信息
    Call LoadBalanceInfor
    Call LoadDelBalanceInfor
    Call ReCalcDelMoney
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    mrsDelInvoice.Filter = 0
End Sub
Private Sub SynchronizationSelect(ByVal lngRow As Long)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:同步选择
    '入参:lngRow-当前选择的行
    '编制:刘兴洪
    '日期:2013-05-08 16:54:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo errHandle
    With vsBill
        If mBillDelType = EM_多张全退 Then
           .TextMatrix(lngRow, .ColIndex("选择")) = 1
          Exit Sub
        End If
        If mBillDelType = EM_单张全退 Then
          Call SetNOBill(.TextMatrix(lngRow, .ColIndex("单据号")), Val(.TextMatrix(lngRow, .ColIndex("选择"))) <> 0)
          Exit Sub
        End If
        '29201
        If Val(.Cell(flexcpData, lngRow, .ColIndex("项目"))) = 0 Then
            For i = lngRow + 1 To vsBill.Rows - 1
                 If Val(vsBill.RowData(lngRow)) = Val(vsBill.Cell(flexcpData, i, .ColIndex("项目"))) Then
                       vsBill.TextMatrix(i, .ColIndex("选择")) = vsBill.TextMatrix(lngRow, .ColIndex("选择"))
                 Else
                    Exit For
                 End If
            Next
            Call zlSet诊疗固定关系(lngRow, .ColIndex("选择"))
            Exit Sub
        End If
        Call zlSet诊疗固定关系(lngRow, .ColIndex("选择"))
        '需要检查主项是否已经被
        For i = lngRow - 1 To 1 Step -1
            If Val(.RowData(i)) = Val(.Cell(flexcpData, lngRow, .ColIndex("项目"))) Then
                If .TextMatrix(i, .ColIndex("选择")) <> 0 Then
                     .TextMatrix(i, .ColIndex("选择")) = .TextMatrix(lngRow, .ColIndex("选择"))
                End If
                Call zlSet诊疗固定关系(i, .ColIndex("选择"), lngRow)
                 Exit For
            End If
        Next
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub ShowAndHideDelBillRow()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:显示或隐藏退费行
    '编制:刘兴洪
    '日期:2013-05-09 10:14:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnSeled As Boolean, r As Long, c As Long
    On Error GoTo errHandle
    If mbytMode <> 1 Or gTy_Module_Para.byt票据分配规则 = 0 Then Exit Sub
    If mBillDelType = EM_多张全退 Then Exit Sub
    blnSeled = False
    With vsInvoice
        For r = 0 To .Rows - 1
            For c = 1 To .COLS - 1
                    If Trim(.Cell(flexcpData, r, c)) <> "" Then
                        If Abs(Val(.Cell(flexcpChecked, r, c))) = 1 Then
                            blnSeled = True: Exit For
                        End If
                    End If
            Next
            If blnSeled Then Exit For
        Next
    End With
    With vsBill
        '隐藏未选择的行
        For r = 1 To .Rows - 1
            .RowHidden(r) = False
            If Abs(Val(.Cell(flexcpChecked, r, .ColIndex("选择")))) <> 1 And blnSeled Then
                .RowHidden(r) = True
            End If
        Next
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Sub OlnyShowSelectedInvoice()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:仅显示被勾选的发票,未勾选的,删除
    '编制:刘兴洪
    '日期:2013-05-09 10:23:34
    '说明:只有通过发票提取单据时显示
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str发票号 As String, r As Long, c As Long, i As Long
    Dim varTemp As Variant
    If mbytMode <> 1 Or gTy_Module_Para.byt票据分配规则 = 0 Then Exit Sub
    On Error GoTo errHandle
    With vsInvoice
        For r = 0 To .Rows - 1
            For c = 1 To .COLS - 1
                If .Cell(flexcpData, r, c) <> "" Then
                    If Abs(Val(.Cell(flexcpChecked, r, c))) = 1 Then
                        str发票号 = str发票号 & "," & .Cell(flexcpData, r, c)
                    End If
                End If
            Next
        Next
        '加载发票号
        If str发票号 = "" Then Exit Sub
        str发票号 = Mid(str发票号, 2)
        varTemp = Split(str发票号, ",")
        .Clear
        .Rows = 1: .COLS = 1
        .FixedCols = 1
        .TextMatrix(0, 0) = "发票号"
        .Redraw = flexRDNone
        .COLS = .COLS + UBound(varTemp) + 1
        For i = 0 To UBound(varTemp)
            If i + 1 > .COLS - 1 Then
                .COLS = .COLS + 1
            End If
            .TextMatrix(0, i + 1) = CStr(varTemp(i))
            .Cell(flexcpData, 0, i + 1) = CStr(varTemp(i))
            .Cell(flexcpChecked, 0, i + 1) = 1
        Next
        .AutoSizeMode = flexAutoSizeColWidth
        Call .AutoSize(0, .COLS - 1)
        Call picInvoice_Resize
        Call Form_Resize
        .Redraw = flexRDBuffered
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsInvoice_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If vsBill.Enabled Then vsBill.SetFocus
    End If
End Sub

Private Function Get误差费(ByVal strNos As String) As Double
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取原始单据的误差费,以便退费时误差处理
    '入参:strNos-单据号(多个用逗号分离)
    '返回:成功返回误差金额
    '编制:刘兴洪
    '日期:2013-11-29 15:06:11
    '说明:针对多单据全退时,需要减去误差费
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo errHandle
    
    strSQL = "" & _
    "   Select /*+ rule */ nvl(Sum(实收金额),0) as 误差费 " & _
    "   From 门诊费用记录 A,table(f_str2List([1])) J " & _
    "   where A.NO=J.Column_value and A.记录性质=1 And A.记录状态 in (1,3) And nvl(A.附加标志,0)=9 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNos)
    Get误差费 = RoundEx(Val(Nvl(rsTemp!误差费)), 6)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
