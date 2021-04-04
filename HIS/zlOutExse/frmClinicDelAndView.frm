VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#9.0#0"; "zlIDKind.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmClinicDelAndView 
   AutoRedraw      =   -1  'True
   Caption         =   "病人退费管理"
   ClientHeight    =   8430
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11265
   Icon            =   "frmClinicDelAndView.frx":0000
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
      TabIndex        =   33
      Top             =   960
      Width           =   11265
      Begin VB.Frame fraSelectDownSplit 
         Height          =   30
         Left            =   -15
         TabIndex        =   35
         Top             =   900
         Width           =   11535
      End
      Begin VB.Frame fraSelectTopSplit 
         Height          =   45
         Left            =   -30
         TabIndex        =   34
         Top             =   0
         Width           =   11385
      End
      Begin VSFlex8Ctl.VSFlexGrid vsInvoice 
         Height          =   375
         Left            =   300
         TabIndex        =   32
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
         FormatString    =   $"frmClinicDelAndView.frx":058A
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
      TabIndex        =   26
      TabStop         =   0   'False
      Top             =   5565
      Width           =   11265
      Begin VB.TextBox txt退费摘要 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1035
         MaxLength       =   100
         TabIndex        =   6
         Top             =   60
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
      TabIndex        =   18
      TabStop         =   0   'False
      Top             =   7404
      Width           =   11265
      Begin VB.CommandButton cmdRefuseApply 
         Caption         =   "拒绝(&N)"
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
         Left            =   6300
         TabIndex        =   37
         Top             =   150
         Visible         =   0   'False
         Width           =   1440
      End
      Begin VB.TextBox txtYB 
         Height          =   300
         Left            =   945
         TabIndex        =   28
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
         TabIndex        =   27
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
         TabIndex        =   13
         Top             =   150
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
         TabIndex        =   20
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
         TabIndex        =   19
         ToolTipText     =   "热键：Ctrl+A"
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
         TabIndex        =   12
         Top             =   150
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
      TabIndex        =   14
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
            Picture         =   "frmClinicDelAndView.frx":0648
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
      FormatString    =   $"frmClinicDelAndView.frx":0EDC
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
      TabIndex        =   17
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
         Left            =   10020
         Locked          =   -1  'True
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   75
         Width           =   1275
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
         Left            =   9000
         TabIndex        =   30
         Top             =   132
         Width           =   960
      End
      Begin VB.Label lblAllTotal 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "收费合计"
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
         Top             =   132
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
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   0
      Width           =   11265
      Begin VB.PictureBox picPatiBack 
         BorderStyle     =   0  'None
         Height          =   360
         Left            =   540
         ScaleHeight     =   360
         ScaleWidth      =   2640
         TabIndex        =   31
         Top             =   525
         Width           =   2640
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
            Top             =   -15
            Width           =   1980
         End
         Begin zlIDKind.IDKindNew IDKind 
            Height          =   360
            Left            =   0
            TabIndex        =   36
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
         TabIndex        =   24
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
         TabIndex        =   22
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
            TabIndex        =   23
            Top             =   0
            Width           =   405
         End
      End
      Begin VB.Frame fraInfo_1 
         Height          =   120
         Left            =   -120
         TabIndex        =   21
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
         TabIndex        =   25
         Top             =   600
         Visible         =   0   'False
         Width           =   480
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
         TabIndex        =   16
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
      FormatString    =   $"frmClinicDelAndView.frx":0F56
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
Attribute VB_Name = "frmClinicDelAndView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Public Enum gEM_ChargeDelType
    EM_MULTI_查看 = 0
    EM_MULTI_退费 = 1
    EM_MULTI_异常重退 = 2
    EM_MULTI_退费申请 = 3
    EM_MULTI_取消申请 = 4
    EM_MULTI_退费审核 = 5
    EM_MULTI_拒绝申请 = 6
    EM_MULTI_取消审核 = 7
End Enum
'----------------------------------------------------------------
'接口变量
Private mstrPrivs As String
Private mbytMode As gEM_ChargeDelType  '0-多张单据查看,1-多张单据退费,2-退异常的退费单进行重新退费;3-退费申请
Private mlng结算序号 As Long  '要查看或退费的多张单据中结帐序号
Private mblnNOMoved As Boolean '操作的单据是否在后备数据表中
Private mstrDelTime As String '查看退费单据的登记时间(yyyy-MM-dd HH:mm:ss) '只有查看退费单据时才传入时间,以区别正常单据
Private mstrApplyTime As String
'----------------------------------------------------------------
Private mlngModule  As Long
Private mlng领用ID As Long
Private mstr个人帐户 As String   '医保个人帐户的名称
Private mdbl个帐余额 As Double   '当前病人个人帐户余额,重收费用用
Private mdbl个帐透支 As Double   '个人帐户允许透支金额,重收费用用

Private mblnOK As Boolean
Private mlngShareUseID As Long '共享领用批次ID
Private mstrUseType As String '使用类别
Private mintInvoiceFormat As Integer  '打印的发票格式,发票格式序号
Private mintOldInvoiceFormat As Integer '旧发票格式
Private mintInvoicePrint As Integer '0-不打印;1-自动打印;2-提示打印
Private mint退费回单打印 As Integer '退费回单打印方式 0-不打印,1-自动打印,2-选择是否打印
Private mintInvoiceFormatDel As Integer  '退费打印的发票格式,发票格式序号(91998)
Private mintInvoicePrintDel As Integer '0-不打印;1-自动打印;2-提示打印
Private mblnPrintView As Boolean    '打印前查看调用
Private mblnOneCard As Boolean
Private mblnFirst As Boolean
Private mblnNotClick As Boolean
Private mstrTittle As String
Private mstrNo As String '要查看或退费的多张单据中的某张NO,退费时可以没有

Private mrs结算方式 As ADODB.Recordset
Private mrs收费对照 As ADODB.Recordset '收费对照 :问题:33634
Private mrsDelInvoice As ADODB.Recordset
Private mrsBalance As ADODB.Recordset '记录每张单据的结算情况
Private mrsInsureBalance As ADODB.Recordset '记录每张单据的医保结算明细
Private mrsInfo As ADODB.Recordset

Private mstrOnePatiPrintNos As String, mblnOnePatiPrint As Boolean

Private Type tyBillType
    bln单种结算方式 As Boolean
    strNos As String '实际读出可以退费的单据号
    strAllNOs As String '所有单据号(一次收费的所有单据)
    strDelNOs As String '当前选中要退的单据
    strNosOverFlow As String '超出金额上限的单据号
    strNosPatiDel As String '记录部分退费的单据
    strNotCanDelNOs As String  '(不能退的单据)已经退完的单据或执行不能退的单据
    str结算方式 As String '当前结算方式:多张时,用逗号分隔
    bln存在卡结算 As Boolean
    intInsure  As Integer   '医保单据的险类
    bln单张部分退费 As Boolean
    blnExistOnCard As Boolean '是否存在一卡通结算
    blnExistThreeAllDel As Boolean '是否存在一卡通全退的
    strInvoice As String '当前发票号
    lng原结帐ID As Long
    lng结帐ID As Long '重新结帐ID
    lng冲销ID As Long '冲销ID
    lng结算序号 As Long
    
    lng病人ID As Long
    str姓名 As String
    str性别 As String
    str年龄 As String
    str费别 As String
End Type
Private mCurBillType As tyBillType  '当前单据类型

Private mobjSquare As Object
Private WithEvents mobjIDCard As clsIDCard
Attribute mobjIDCard.VB_VarHelpID = -1
Private WithEvents mobjICCard As clsICCard
Attribute mobjICCard.VB_VarHelpID = -1

Private mobjDrugPacker  As Object ' 自动发药机(更新发药窗口)
Private mblnDrugPacker As Boolean
Private mobjDrugMachine  As Object
Private mblnDrugMachine As Boolean

Private mblnHistoryData As Boolean '是否为“多单据分单据结算”或“一次结算分单据退费”时的历史数据
Private mblnDelByNo As Boolean '是否分单据退费 = (多单据分单据结算=True Or 一次结算分单据退费=True) And Not mblnHistoryData
Private mcllForceDelToCash As Collection '强制退现信息：Array(操作员,卡类别名称)
'-------------------------------------------------------------------------------
'医保相关定义:参数
Private Type TYPE_MedicarePAR
    医保接口打印票据 As Boolean
    退费后打印回单 As Boolean
    医保不走票号  As Boolean        '预结算时有效
    门诊结算作废 As Boolean             '医保是否支持门诊结算作废
    门诊预结算 As Boolean
    先自付 As Boolean
    全自付 As Boolean
    按单据全退 As Boolean
    多单据分单据结算 As Boolean '86321
    一次结算分单据退费 As Boolean '91602
End Type
Private MCPAR As TYPE_MedicarePAR
'-------------------------------------------------------------------------------
'Api定义
Private Const VK_RETURN = &HD
Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer

Private Function CheckDepend() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:数据关联性检查
    '返回:数据关联检查合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-07 11:41:01
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If mbytMode = EM_MULTI_查看 Then CheckDepend = True: Exit Function
    
    Set mrs结算方式 = Get结算方式("收费")
    mrs结算方式.Filter = "性质=3"
    If Not mrs结算方式.EOF Then
       mstr个人帐户 = mrs结算方式!名称
    End If
    mrs结算方式.Filter = 0
    If mrs结算方式.RecordCount = 0 Then
        MsgBox "收费场合没有可用的结算方式，请先到结算方式管理中设置。", vbInformation, gstrSysName
        Exit Function
    End If
    mrs结算方式.MoveFirst
    
    CheckDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Function ShowMe(frmMain As Object, ByVal bytMode As gEM_ChargeDelType, _
    ByVal strPrivs As String, lng结算序号 As Long, _
    Optional blnPrintView As Boolean, _
    Optional lng领用ID As Long = 0, _
    Optional blnNOMoved As Boolean = False, _
    Optional strDelTime As String = "", _
    Optional strApplyTime As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:单据查看,退费
    '入参:bytMode-0-多张单据查看,1-多张单据退费,2-退异常的退费单进行重新退费
    '     strPrivs-权限串
    '     mblnPrintView-打印前查看调用
    '     blnNOMoved-是否转到后备数据表
    '     strDelTime-查看退费单据的登记时间(yyyy-MM-dd HH:mm:ss) '只有查看退费单据时才传入时间,以区别正常单据
    '     strApplyTime-申请时间(yyyy-MM-dd HH:mm:ss)，退费申请模式时必须传入
    '出参:
    '返回:退费成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-06-24 14:34:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mblnNOMoved = blnNOMoved: mstrPrivs = strPrivs
    mlng领用ID = lng领用ID: mlng结算序号 = lng结算序号
    mlngModule = 1121: mblnPrintView = blnPrintView
    mbytMode = bytMode:
    mstrDelTime = strDelTime              '只有查看退费单据时才传入时间,以区别正常单据
    mstrApplyTime = strApplyTime
    mblnOK = False
    If CheckDepend = False Then Exit Function
    On Error Resume Next
    
    Me.Show IIf(gfrmMain Is Nothing, 0, 1), frmMain
    On Error GoTo 0
    ShowMe = mblnOK
End Function

Private Sub cmdRefuseApply_Click()
    If SaveDelApplied(EM_MULTI_拒绝申请) = False Then Exit Sub
    mblnOK = True
    Unload Me: Exit Sub
End Sub

Private Sub Form_Load()
    mblnFirst = True
    mstrTittle = "病人退费管理"
    Call InitFace
    Call RestoreWinState(Me, App.ProductName, mstrTittle)
    
    '81190,冉俊明,退费业务向发药机上传退费信息
    Call CreateDrugPacker
End Sub
Private Sub CreateDrugPacker()
    '功能:创建自助发药机(自动化药房)
    Dim objComLib As New zl9ComLib.clsComLib
    Dim strPrivs As String
    Dim strMessage As String
    
    '81190,冉俊明,退费业务向发药机上传退费信息
    mblnDrugPacker = False: mblnDrugMachine = False
    If Not (mbytMode = EM_RBDTY_退费 Or mbytMode = EM_RBDTY_异常重退) Then Exit Sub

    Err = 0: On Error Resume Next
    If Val(zlDatabase.GetPara("启用药品自动化设备接口", glngSys, Val("9010-药品自动化设备接口"))) = 1 Then
        '优先新接口
        Set mobjDrugMachine = CreateObject("zlDrugMachine.clsDrugMachine")
        If Err = 0 Then mblnDrugMachine = True
    End If
    
    If mblnDrugMachine = False Then
        '旧部件
        Err = 0
        Set mobjDrugPacker = CreateObject("zlDrugPacker.clsDrugPacker")
        If Err = 0 Then mblnDrugPacker = True
    End If
    
    Err = 0: On Error GoTo 0
    If mblnDrugMachine Then
        '权限检查
        strPrivs = GetPrivFunc(glngSys, Val("9010-药品自动化设备接口"))
        If InStr(";" & strPrivs & ";", ";基本;") > 0 Then
            mblnDrugMachine = mobjDrugMachine.Init(1, objComLib, strMessage)
        Else
            mblnDrugMachine = False
        End If
    ElseIf mblnDrugPacker Then
        mblnDrugPacker = mobjDrugPacker.DYEY_MZ_IniSoap
    End If
End Sub

Private Sub InitFace()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化界面
    '编制:刘兴洪
    '日期:2014-06-24 14:36:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim TY_Temp As tyBillType
    
    Call SetpicInvoiceVisible '设置发票控件显示
    Call InitBillHead   '设置单据列头
    
    If Val(zlDatabase.GetPara("退费号码输入模式", glngSys, 1121, 0)) = 0 Then
        optNO(0).Value = True
    Else
        optNO(1).Value = True
    End If
    mCurBillType = TY_Temp
    mint退费回单打印 = Val(zlDatabase.GetPara("退费回单打印方式", glngSys, mlngModule, "0"))
    Call NewCardObject
    Call ClearFace
    Call SetFunCtrlVisible
    
    Select Case mbytMode
    Case EM_MULTI_查看
        mstrTittle = "病人收费单据查阅"
        Caption = mstrTittle
        vsBill.ColHidden(0) = True
        cmdCancel.Caption = "退出(&X)"
        If mblnPrintView Then cmdCancel.Caption = "确定(&X)"
        pic退.Visible = mstrDelTime <> ""
        
        mblnOneCard = False
    Case EM_MULTI_异常重退
        mstrTittle = "病人退费管理-异常退费单重新退费"
        Caption = mstrTittle
        vsBill.ColHidden(0) = True
        pic退.Visible = mstrDelTime <> ""
        vsBill.Editable = flexEDNone
        mblnOneCard = GetOneCard.RecordCount <> 0
        Call initCardSquareData
    Case EM_MULTI_退费申请, EM_MULTI_取消申请, EM_MULTI_退费审核, EM_MULTI_取消审核
        mstrTittle = "病人退费管理-" & Switch(mbytMode = EM_MULTI_退费申请, "退费申请", mbytMode = EM_MULTI_取消申请, "取消申请", _
                                            mbytMode = EM_MULTI_退费审核, "退费审核", mbytMode = EM_MULTI_取消审核, "取消审核")
        Caption = mstrTittle
        Call initCardSquareData
    Case Else 'EM_MULTI_退费
        mstrTittle = "病人退费管理"
        Caption = mstrTittle
        Call initCardSquareData
        mblnOneCard = GetOneCard.RecordCount <> 0
    End Select
    
    If mlng结算序号 <> 0 Then
        picPatiBack.Top = txtNO.Top
        lblPati.Top = picPatiBack.Top + (picPatiBack.Height - lblPati.Height) \ 2
        txtPatientPrint.Top = txtNO.Top
        lblPatiName.Top = txtPatientPrint.Top + (txtPatientPrint.Height - lblPatiName.Height) \ 2
        picPati.Height = 480
    End If
    
    
End Sub

Private Sub SetpicInvoiceVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置发票控件的显示
    '编制:刘兴洪
    '日期:2014-06-24 14:36:26
    '---------------------------------------------------------------------------------------------------------------------------------------------
    picInvoice.Visible = False
    If gTy_Module_Para.byt票据分配规则 = 0 And gTy_Module_Para.bln分别打印 = False Then GoTo ReSizing:
    If mbytMode <> EM_MULTI_退费 Then GoTo ReSizing:
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

Private Sub InitBillHead()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化退费的表头列信息
    '编制:刘兴洪
    '日期:2014-06-24 14:39:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim arrHead As Variant, strHead As String, i As Long
    Dim varTemp As Variant, intCol As Integer

    strHead = "" & _
    "选择,300,4;单据号,1000,1;类别,720,1;项目,2800,1;商品名,2000,1;数量,750,7;单位,550,1;单价,1100,7;" & _
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

Private Sub ClearFace(Optional ByVal blnNO As Boolean = True)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:清除界面的信息
    '入参:blnNo=清除单据号
    '编制:刘兴洪
    '日期:2014-06-24 15:19:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mrsBalance = Nothing: Set mrsInsureBalance = Nothing
    Set mrsDelInvoice = Nothing

    With vsBill
        .Rows = .FixedRows '对非固定行的第一行被隐藏时恢复可见
        .Rows = .FixedRows + 1
        .Row = .FixedRows: .Col = .ColIndex("项目")
        .Clear 1
    End With
    mCurBillType.strNos = ""
    mCurBillType.intInsure = 0
    lblPati.Caption = "病人:"
    If blnNO Then txtNO.Text = ""
    Call SetpicInvoiceVisible
    
    Call ClearBalance
    With vsBalance
         .COLS = 1
         .TextMatrix(0, 0) = IIf(mstrDelTime = "", "收款结算", "退款结算")
    End With
    txtCurTotal.Text = ""
    txtAllTotal.Text = ""
    txt退款合计.Text = ""
    stbThis.Panels(2).Text = ""
    Call SetFunCtrlVisible
End Sub

Private Sub NewCardObject()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化新的卡对象
    '编制:刘兴洪
    '日期:2014-06-24 14:43:56
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If mbytMode <> EM_MULTI_查看 Then Exit Sub
   
    If mobjIDCard Is Nothing Then
        Set mobjIDCard = New clsIDCard
        Call mobjIDCard.SetParent(Me.hWnd)
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
    End If
    If mobjICCard Is Nothing Then
        Set mobjICCard = New clsICCard
        Call mobjICCard.SetParent(Me.hWnd)
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
    End If
    IDKind.SetAutoReadCard (False)
End Sub
Private Sub CloseIDCard()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:关闭自助读卡功能
    '编制:刘兴洪
    '日期:2014-06-24 14:43:35
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

Private Function LoadViewBills(ByVal lng结算序号 As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据结算序号来加载数据(只针对查看或异常退费)
    '入参:lng结算序号-结算序号
    '返回:加载或读取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-06-24 16:17:58
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllInvoiceNoInfor As Collection
    Dim intSign As Integer, strSQL As String, rsTemp As ADODB.Recordset
    Dim str结帐ID As String, strNos As String, strAllNOs As String, intInsure As Integer
    Dim strWhere As String, lng结帐ID As Long, str医嘱序号 As String, lng原结算序号 As Long
    Dim lng医嘱序号 As Long, rsAdvice As ADODB.Recordset, i As Long
    Dim strTemp As String, j As Long, dbl合计 As Double, lng冲销ID As Long
    Dim varData As Variant, strInfos As String
    Dim varNos As Variant, blnNotDelAll As Boolean, blnHaveExe As Boolean
    
    If mbytMode = EM_MULTI_退费 Then
        '退费需要重新算
        LoadViewBills = ReadBills("")
        Exit Function
    End If
    
    Screen.MousePointer = 11
    intSign = IIf(mstrDelTime <> "", -1, 1) '数量,金额正负符号
    On Error GoTo errHandle
    
    str结帐ID = zlGet结帐ID(lng结算序号, strNos, intInsure, mblnNOMoved, lng冲销ID)
    
    mCurBillType.lng冲销ID = lng冲销ID
    varData = Split(str结帐ID & ",,", ",")
     If Val(varData(0)) = lng冲销ID Then
         mCurBillType.lng结帐ID = Val(varData(1))
    End If
    
    If InStr(str结帐ID, ",") > 0 Then
        strWhere = "And A.结帐ID IN (Select Column_Value From table(f_num2List([1])))"
        lng结帐ID = Split(str结帐ID & ",", ",")(0)
    Else
        strWhere = "And A.结帐ID=[2]"
        lng结帐ID = Val(str结帐ID)
    End If

        
    'bytType-0-根据NO来查找;1-根据结帐ID来查找,2-根据结算序号来查找
    strAllNOs = zlGetBalanceNos(1, lng结帐ID, mblnNOMoved)
    mCurBillType.strAllNOs = strAllNOs
    mCurBillType.intInsure = intInsure
    mCurBillType.strNos = strNos
     
     
    'bytType-查找类型:0-根据结帐ID查找;1-根据结算序号查找,2-根据NOs来获取
    If mbytMode = EM_MULTI_异常重退 Then
        mCurBillType.lng原结帐ID = zlGetFromNOToLastBalanceID(strAllNOs, mblnNOMoved, False, lng原结算序号)
    End If
    Set mrsBalance = zlFromIDGetChargeBalance(1, lng结算序号, mblnNOMoved)
    Set mrsInsureBalance = zlGetInsureBalanceDetail(1, lng结算序号, mblnNOMoved)
    
    strSQL = "" & _
    " Select A.病人ID,A.姓名,A.性别,A.年龄,A.标识号,A.费别,C.名称 as 付款方式,B.病人类型,B.险类 " & _
    " From 门诊费用记录 A,医疗付款方式 C,人员表 D,病人信息 B" & _
    " Where A.付款方式=C.编码(+)  And  A.操作员姓名=D.姓名 And A.病人ID=B.病人ID(+) " & _
    "       And (D.站点='" & gstrNodeNo & "' Or D.站点 is Null)" & vbNewLine & _
    "       And mod(A.记录性质,10)=1 And Rownum <2 " & strWhere
    If mblnNOMoved Then
        strSQL = Replace(strSQL, "门诊费用记录", "H门诊费用记录")
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str结帐ID, lng结帐ID)
    If rsTemp.EOF Or strAllNOs = "" Then
        Screen.MousePointer = 0
        MsgBox "没有找到与结算相关的收费记录。", vbInformation, gstrSysName
        mCurBillType.lng病人ID = 0
        Exit Function
    End If
    
    txtPatient.Text = Nvl(rsTemp!姓名)
    lblPati.Caption = "病人:" & IIf(txtPatient.Visible, "       ", rsTemp!姓名) & _
        "　性别:" & Nvl(rsTemp!性别) & _
        "　年龄:" & Nvl(rsTemp!年龄) & _
        "　门诊号:" & Nvl(rsTemp!标识号) & _
        "　费别:" & Nvl(rsTemp!费别) & _
        "　付款方式:" & rsTemp!付款方式
    
    With mCurBillType
        .lng病人ID = Val(Nvl(rsTemp!病人ID))
        .str性别 = Nvl(rsTemp!性别)
        .str年龄 = Nvl(rsTemp!年龄)
        .str姓名 = Nvl(rsTemp!姓名)
    End With
    
    If mbytMode <> EM_MULTI_查看 Then
        Call initInsurePara(mCurBillType.intInsure, mCurBillType.lng病人ID, lng结帐ID)
        mblnDelByNo = MCPAR.多单据分单据结算 Or MCPAR.一次结算分单据退费
    End If
    If CheckPrivsIsValied = False Then Screen.MousePointer = 0: Exit Function   '操作权限检查
    
    If mCurBillType.intInsure <> 0 Then
        lblPati.ForeColor = vbRed
        txtYB.Text = mCurBillType.intInsure
        txtPatient.ForeColor = vbRed
    Else
        lblPati.ForeColor = &HC00000
        txtYB.Text = ""
        txtPatient.ForeColor = &HC00000
    End If
    Call SetPatiColor(txtPatient, Nvl(rsTemp!病人类型), txtPatient.ForeColor)
    lblPati.ForeColor = txtPatient.ForeColor
    If mblnPrintView And zlStr.IsHavePrivs(mstrPrivs, "修改姓名重打") _
        And mCurBillType.lng病人ID = 0 Then
        txtPatientPrint.Text = "" & rsTemp!姓名
        txtPatientPrint.Tag = txtPatientPrint.Text
        txtPatientPrint.Visible = True
        lblPatiName.Visible = True
    End If
    
    If mblnDelByNo Then
        If mrsBalance.RecordCount > 0 Then
            mblnHistoryData = zlGetInsureBalanceDetail(1, Val(Nvl(mrsBalance!结算序号))).RecordCount = 0
        End If
        mblnDelByNo = Not mblnHistoryData
    End If
    
    '加载结算方式
    Call LoadBalanceInfor
    If mbytMode = EM_MULTI_退费申请 Then
        strWhere = strWhere & " And Not exists(select 1 From 病人退费申请 where NO=A.NO And 记录性质=1 And Nvl(状态,0) In(0,1) ) "
    ElseIf mbytMode = EM_MULTI_取消申请 Or mbytMode = EM_MULTI_退费审核 Then
        strWhere = strWhere & " And Exists(select 1 From 病人退费申请 where NO=A.NO And 记录性质=1 And Nvl(状态, 0) = 0 " & _
                              " And 申请时间=To_Date('" & mstrApplyTime & "','yyyy-mm-dd hh24:mi:ss')) "
    ElseIf mbytMode = EM_MULTI_取消审核 Then
        strWhere = strWhere & " And Exists(select 1 From 病人退费申请 where NO=A.NO And 记录性质=1 And Nvl(状态, 0) = 1 " & _
                              " And 申请时间=To_Date('" & mstrApplyTime & "','yyyy-mm-dd hh24:mi:ss')) "
        '已退过费的不允许取消审核
        strWhere = strWhere & " And Not Exists(select 1 From 门诊费用记录 where NO=A.NO And 记录性质=A.记录性质 And 记录状态=2)"
    End If
    'InStr(str结帐ID, ",") > 0:表示可能存在重收的情况，所以肯定是查的退费记录，所以摘要应该以退费的摘要为准
    '104788，不单独计算付数，直接计算数次，因为医保病人多单据一次结算时部分退，查看退费单据显示的数量翻倍了
    '" Avg(Nvl(A.付数,1)) as 付数,Avg(A.数次) as 数次," 改为 " Avg(Nvl(A.付数,1)*A.数次) as 数次,"
    strSQL = "" & _
    "   Select A.NO,Nvl(A.价格父号,A.序号) as 序号,A.从属父号,A.开单部门ID,A.执行部门ID,A.收费类别,A.费别,A.收费细目ID," & vbNewLine & _
    "          A.费用类型,A.计算单位,max(A.医嘱序号) as 医嘱序号," & vbNewLine & _
    "          Avg(Nvl(A.付数,1)*A.数次) as 数次," & vbNewLine & _
    "          Sum(A.标准单价) as 单价, Sum(A.应收金额) as 应收金额,Sum(A.实收金额) as 实收金额," & vbNewLine & _
    "          Max(A.操作员姓名) as 操作员姓名,max(A.登记时间) as 登记时间," & _
    "           " & IIf(InStr(str结帐ID, ",") > 0, "Max(Decode(A.记录状态,2,A.摘要,NULL))", "Max(A.摘要)") & " as 摘要,A.结帐ID" & vbNewLine & _
    "   From 门诊费用记录 A" & vbNewLine & _
    "   Where Mod(A.记录性质,10)=1  " & strWhere & vbNewLine & _
    "   Group by A.结帐ID,A.NO,Nvl(A.价格父号,A.序号),A.从属父号,A.开单部门ID,A.执行部门ID,A.收费类别,A.费别,A.收费细目ID,A.费用类型,A.计算单位,A.结帐ID"
    
    If mblnNOMoved Then
        strSQL = Replace(strSQL, "门诊费用记录", "H门诊费用记录")
    End If
    
    strSQL = _
    " Select /*+ Rule*/ A.NO,A.序号,A.从属父号,A.费别,A.收费细目ID,C.编码 as 类别码,C.名称 as 类别名,B.编码, " & _
    "       Nvl(M1.名称,B.名称) as 名称,E1.名称 as 商品名 ,B.规格,Nvl(A.费用类型,B.费用类型) 费用类型," & _
            IIf(gbln药房单位, "Decode(X.药品ID,NULL,A.计算单位,X." & gstr药房单位 & ")", "A.计算单位") & " as 计算单位," & _
    "       Max(A.医嘱序号) as 医嘱序号," & _
    "       sum(A.数次" & IIf(gbln药房单位, "/Nvl(X." & gstr药房包装 & ",1)", "") & ") as 数次," & _
    "       Max(A.单价" & IIf(gbln药房单位, "*Nvl(X." & gstr药房包装 & ",1)", "") & ") as 单价," & _
    "       Sum(A.应收金额) as 应收金额,Sum(A.实收金额) as 实收金额, 1 as 记录标志,0 as 原始数量,0 as 准退数量," & _
    "       D.名称 as 执行科室,A.执行部门ID,E.名称 as 开单科室,A.操作员姓名,A.登记时间, " & _
    "       Max(A.摘要) as 摘要" & _
    " From (" & strSQL & ") A,收费项目目录 B,收费项目类别 C,部门表 D,部门表 E,药品规格 X," & _
    "       收费项目别名 M1,收费项目别名 E1" & _
    " Where A.收费细目ID=B.ID And C.编码=A.收费类别 And A.收费细目ID=X.药品ID(+)" & _
    "       And A.执行部门ID=D.ID(+) And A.开单部门ID=E.ID(+) " & _
    "       And A.收费细目ID=M1.收费细目ID(+) And M1.码类(+)=1 And M1.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
    "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3 " & _
    " Group by A.NO,A.序号,A.从属父号,A.费别,A.收费细目ID,C.编码,C.名称,B.编码,Nvl(M1.名称,B.名称)," & _
    "       E1.名称,B.规格,Nvl(A.费用类型,B.费用类型),A.计算单位,D.名称,A.执行部门ID,E.名称,X.药品ID,X." & gstr药房单位 & ",A.操作员姓名,A.登记时间" & _
    " Having Sum(A.数次)<>0 " & _
    " Order by NO,序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str结帐ID, lng结帐ID)
    
    If rsTemp.EOF Then
        Screen.MousePointer = 0
        If mbytMode = EM_MULTI_查看 Then
            MsgBox "没有找到指定结算信息的费用记录,可能因并发原因被他人操作或操作了错误的结算单据。", vbInformation, gstrSysName
        ElseIf mbytMode = EM_MULTI_退费申请 Then
            MsgBox "没有找到需要退费申请的单据，可能已被他人申请了。", vbInformation, gstrSysName
        ElseIf mbytMode = EM_MULTI_取消申请 Then
            MsgBox "没有找到需要取消申请的单据，可能已被他人取消申请或审核了。", vbInformation, gstrSysName
        ElseIf mbytMode = EM_MULTI_退费审核 Then
            MsgBox "没有找到需要退费审核的单据，可能已被他人审核了。", vbInformation, gstrSysName
        ElseIf mbytMode = EM_MULTI_取消审核 Then
            MsgBox "没有找到需要取消审核的单据，可能已退费或被他人取消审核了。", vbInformation, gstrSysName
        Else
            MsgBox "没有找到与结算信息相关的可以退费的记录。" & _
                vbCrLf & "这些收费记录可能已经退费或已经完全执行。", vbInformation, gstrSysName
        End If
        Call ClearFace(False)
        Exit Function
    End If
    
    If mbytMode <> EM_MULTI_退费 Then
        If mbytMode = EM_MULTI_退费申请 Or mbytMode = EM_MULTI_取消申请 Or mbytMode = EM_MULTI_退费审核 Or mbytMode = EM_MULTI_取消审核 Then
            pic退费摘要.Enabled = True
            pic退费摘要.Visible = True
            lbl摘要.Caption = Switch(mbytMode = EM_MULTI_退费申请, "申请原因", mbytMode = EM_MULTI_取消申请, "申请原因", mbytMode = EM_MULTI_退费审核, "审核/拒绝原因", _
                                    mbytMode = EM_MULTI_取消审核, "审核原因")
            txt退费摘要.Text = ""
            If mbytMode = EM_MULTI_退费申请 Or mbytMode = EM_MULTI_取消申请 Then
                lbl退款合计.Caption = "申请合计"
            ElseIf mbytMode = EM_MULTI_退费审核 Or mbytMode = EM_MULTI_取消审核 Then
                lbl退款合计.Caption = "审核合计"
            End If
        Else
            pic退费摘要.Enabled = mbytMode = EM_MULTI_异常重退
            txt退费摘要.Text = Nvl(rsTemp!摘要)
        End If
    End If
    
    With rsTemp
        str医嘱序号 = ""
        Do While Not .EOF
            lng医嘱序号 = Val(Nvl(!医嘱序号))
            If InStr(str医嘱序号 & ",", "," & lng医嘱序号 & ",") = 0 And lng医嘱序号 <> 0 Then
                str医嘱序号 = str医嘱序号 & "," & Val(Nvl(!医嘱序号))
            End If
            .MoveNext
        Loop
        .MoveFirst
    End With
    
    Set rsAdvice = Nothing
    If str医嘱序号 <> "" Then
        str医嘱序号 = Mid(str医嘱序号, 2)
        Set rsAdvice = zlGetAdviceFromID(str医嘱序号)
    End If
    Call LoadInvoiceData(Replace(strAllNOs, "'", ""))
    With vsBill
        .Redraw = flexRDNone
        .Rows = .FixedRows + rsTemp.RecordCount
        mCurBillType.strDelNOs = ""
        For i = 1 To rsTemp.RecordCount
            .RowData(i) = Val(Nvl(rsTemp!序号))
            .TextMatrix(i, .ColIndex("选择")) = 0
            .Cell(flexcpData, i, .ColIndex("项目")) = Val(Nvl(rsTemp!从属父号))
            .Cell(flexcpData, i, .ColIndex("结帐ID")) = Nvl(rsTemp!医嘱序号) & "," & Nvl(rsTemp!收费细目ID)
            strTemp = ""
            If Val(Nvl(rsTemp!从属父号)) <> 0 Then
                rsTemp.MoveNext
                strTemp = "┣"
                If rsTemp.EOF Then
                    strTemp = "┗"
                ElseIf Val(.Cell(flexcpData, i, .ColIndex("项目"))) <> Nvl(rsTemp!从属父号) Then
                    strTemp = "┗"
                End If
                rsTemp.MovePrevious
                strTemp = "  " & strTemp & " "
            End If
            .TextMatrix(i, .ColIndex("单据号")) = Nvl(rsTemp!NO)
            .TextMatrix(i, .ColIndex("类别")) = Nvl(rsTemp!类别名)
            .TextMatrix(i, .ColIndex("项目")) = strTemp & rsTemp!名称 & IIf(IsNull(rsTemp!规格), "", " " & rsTemp!规格)
            .TextMatrix(i, .ColIndex("商品名")) = strTemp & Nvl(rsTemp!商品名)
            .TextMatrix(i, .ColIndex("数量")) = FormatEx(intSign * rsTemp!数次, 5)
            .Cell(flexcpData, i, .ColIndex("数量")) = intSign * rsTemp!数次
            .TextMatrix(i, .ColIndex("单位")) = Nvl(rsTemp!计算单位)
            .TextMatrix(i, .ColIndex("单价")) = Format(rsTemp!单价, gstrFeePrecisionFmt)
            .TextMatrix(i, .ColIndex("应收金额")) = Format(intSign * Val(Nvl(rsTemp!应收金额)), gstrDec)
            .TextMatrix(i, .ColIndex("实收金额")) = Format(intSign * Val(Nvl(rsTemp!实收金额)), gstrDec)
            .TextMatrix(i, .ColIndex("开单科室")) = Nvl(rsTemp!开单科室)
            .TextMatrix(i, .ColIndex("执行科室")) = Nvl(rsTemp!执行科室)
            .TextMatrix(i, .ColIndex("操作员")) = rsTemp!操作员姓名
            .TextMatrix(i, .ColIndex("时间")) = Format(rsTemp!登记时间, "MM-dd HH:mm")
            .TextMatrix(i, .ColIndex("结帐ID")) = str结帐ID
            lng结算序号 = Val(Nvl(rsTemp!医嘱序号))
            If Not rsAdvice Is Nothing And str医嘱序号 <> "" And lng结算序号 <> 0 Then
                rsAdvice.Filter = "医嘱ID=" & lng结算序号
                If rsAdvice.EOF = False Then
                    .TextMatrix(i, .ColIndex("医嘱")) = Nvl(rsAdvice!医嘱内容)
                End If
            End If
            .TextMatrix(i, .ColIndex("原始数量")) = Nvl(rsTemp!原始数量)
            .TextMatrix(i, .ColIndex("准退数量")) = Nvl(rsTemp!准退数量)
            .TextMatrix(i, .ColIndex("医嘱序号")) = Nvl(rsTemp!医嘱序号)
            .TextMatrix(i, .ColIndex("执行科室ID")) = Nvl(rsTemp!执行部门ID)
            .Cell(flexcpData, i, .ColIndex("选择")) = Val(Nvl(rsTemp!记录标志))    '用于判断是否被销帐过,>1表示已销帐
            If Val(Nvl(rsTemp!记录标志)) > 1 And InStr(1, mCurBillType.strNosPatiDel & ",", "," & rsTemp!NO & "") = 0 Then
                mCurBillType.strNosPatiDel = mCurBillType.strNosPatiDel & "," & rsTemp!NO
            End If
            If InStr(mCurBillType.strDelNOs & ",", "," & rsTemp!NO & ",") = 0 Then
                '画出分隔线
                If mCurBillType.strDelNOs <> "" Then
                    .Select i, .FixedCols, i, .COLS - 1
                    .CellBorder vbBlue, 0, 1, 0, 0, 0, 0
                End If
                mCurBillType.strDelNOs = mCurBillType.strDelNOs & "," & rsTemp!NO
            End If
            dbl合计 = dbl合计 + Val(Nvl(rsTemp!实收金额))
            rsTemp.MoveNext
        Next
        .Row = .FixedRows: .Col = .ColIndex("项目")
        Call vsBill_AfterRowColChange(-1, -1, .Row, .Col)
        .Redraw = flexRDDirect
    End With
    
    Call SetpicInvoiceVisible   '设置发票控件的显示
    txtAllTotal.Text = Format(intSign * dbl合计, gstrDec)
    Call ReInitPatiInvoice
    txt退款合计.Text = Format(GetDelMoney, "0.00")
    
    If mbytMode = EM_MULTI_退费申请 Or mbytMode = EM_MULTI_退费审核 Then '存在已执行的项目时进行提示
        varNos = Split(mCurBillType.strDelNOs, ",")
        For i = 0 To UBound(varNos)
            Call BillCanDelete(varNos(i), 1, blnHaveExe)
            If blnHaveExe Then
                strInfos = strInfos & "," & varNos(i)
            End If
        Next
        If strInfos <> "" Then
            strInfos = Mid(strInfos, 2)
            If MsgBox("单据[" & strInfos & "]中存在已执行的项目，你确认要继续吗？", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                Call ClearFace(False): Screen.MousePointer = 0: Exit Function
            End If
        End If
    End If
    
    Screen.MousePointer = 0
    LoadViewBills = True
    Exit Function
errHandle:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ReadBills(ByVal strNo As String, Optional blnCheckMulitBalance As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据当前输入的单据号或票据号,读取并显示多张单据
    '入参:strNO-指定的单据号或发票号
    '     blnCheckMulitBalance-已经检查了多单据一次结算,不用再在内部检查
    '返回:读取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-06-24 15:41:24
    '说明:
    '   只有退费模式下才进入该模块
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim strNos As String, strAllNOs As String
    Dim blnNOMoved As Boolean
    Dim strTmp As String, strCanDelNos As String
    Dim i As Long, j As Integer
    Dim dbl合计 As Currency, arrNo As Variant
    Dim strTemp As String, str医嘱序号 As String
    Dim blnNotFind As Boolean
    Dim lng病人ID As Long, cllInvoiceNoInfor As Collection
    Dim str结算序号 As String, blnFind As Boolean
    Dim strInvoiceNO As String, strOldNO As String
    On Error GoTo errH
    If mbytMode <> EM_MULTI_退费 Then Exit Function
    
    Screen.MousePointer = 11
    
    Call ClearFace(False)
    strOldNO = strNo
    Set cllInvoiceNoInfor = New Collection
    mblnOnePatiPrint = False: mstrOnePatiPrintNos = ""
    If mlng结算序号 = 0 And strNo <> "" Then
        strInvoiceNO = ""
        If Not (mstrNo <> "" Or optNO(0).Value) Then
             '按票据号:可能不同批次票号重复
            strInvoiceNO = strNo
            blnNOMoved = zlDatabase.NOMoved("票据打印明细", "票号 =", "1")
            strNos = zlInvoiceGetNOs(strInvoiceNO, cllInvoiceNoInfor, blnNOMoved)
            strNo = Split(strNos & ",", ",")(0)
            If zlIsOnePatiPrint(strNo, mstrOnePatiPrintNos, mblnOnePatiPrint, blnNOMoved) = False Then Exit Function
            
            If mblnOnePatiPrint Then    '多次结算时，需要选择指定的结算方式
                 If SelectMulitBalance(mstrOnePatiPrintNos, strNo) = False Then Exit Function
            End If
        Else
            If zlIsOnePatiPrint(strNo, mstrOnePatiPrintNos, mblnOnePatiPrint, blnNOMoved) = False Then Exit Function

        End If
        
        blnNOMoved = zlDatabase.NOMoved("门诊费用记录", strNo, , "1")
        strNos = zlGetBalanceNos(0, strNo, blnNOMoved)
        If blnCheckMulitBalance = False Then
            '78663,冉俊明,2014-10-15,无收费单据时弹出提示
            If Trim(strNos) = "" Then
                Screen.MousePointer = 0
                MsgBox "没有找到与号码""" & strOldNO & """相关的收费记录。", vbInformation, gstrSysName
                Exit Function
            End If
            If Not zlIsMulitOneBalance(strNos) Then
                '非多单据一次结算,需采用34以前版本处理规则处理
                'bytMode-0-多张单据查看,1-多张单据退费,2-退异常的退费单进行重新退费
                frmMultiBills.ShowMe Me, 1, mstrPrivs, strNo, "", False, mlng领用ID, mblnOneCard, False, True
                Exit Function
            End If
        End If
    Else
        'bytType-0-根据NO来查找;1-根据结帐ID来查找,2-根据结算序号来查找
        strNos = zlGetBalanceNos(2, mlng结算序号, mblnNOMoved)
        strNo = Split(strNos & ",", ",")(0)
        If zlIsOnePatiPrint(strNo, mstrOnePatiPrintNos, mblnOnePatiPrint, blnNOMoved) = False Then Exit Function '按结算次数来的，肯定不用再去选择单据
    End If
    
    strAllNOs = strNos
    
    If strNos = "" Then
        If optNO(1).Value Then
            Screen.MousePointer = 0
            MsgBox "没有找到与号码""" & strNo & """相关的收费记录。", vbInformation, gstrSysName
            Exit Function
        End If
        '可能因为未用票据而读不出来
        strNos = strNo
    End If
    mCurBillType.lng原结帐ID = zlGetFromNOToLastBalanceID(strNos, blnNOMoved)
    mCurBillType.strAllNOs = strAllNOs
    '执行补结算的单据不能进行退费
    If CheckBillExistReplenishData(1, , strNos) = True Then
        Screen.MousePointer = 0
        MsgBox "选择的退费记录进行了医保补充结算，不允许进行退费操作！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '升级医嘱执行计价.执行状态
    If Upgrade医嘱执行计价执行状态(strNos) = False Then
        Screen.MousePointer = 0
        MsgBox "医嘱执行计价数据修正失败，不能继续！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '需要加引号
    If InStr(1, strNos, "'") = 0 Then
        strNos = "'" & Replace(strNos, ",", "','") & "'"
    End If
    arrNo = Split(strNos, ",")
    
    '获取结算方式
    '0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡)
    Set mrsBalance = zlFromIDGetChargeBalance(2, strAllNOs, mblnNOMoved)
    Set mrsInsureBalance = zlGetInsureBalanceDetail(2, strAllNOs, mblnNOMoved)
    mCurBillType.intInsure = zlGetChargeInsure(mCurBillType.lng原结帐ID, lng病人ID, mblnNOMoved)
    
    '初始化结算方式相关变量
    Call InitBalanceVar
    
    
    Call initInsurePara(mCurBillType.intInsure, lng病人ID, mCurBillType.lng原结帐ID)
    mblnDelByNo = MCPAR.多单据分单据结算 Or MCPAR.一次结算分单据退费
    
    If CheckPrivsIsValied = False Then Exit Function    '操作权限检查
    
    '退费相关检查
    If CheckDelIsValied(strNos, mCurBillType.strNotCanDelNOs, strCanDelNos) = False Then
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
    "       And mod(A.记录性质,10)=1 And A.记录状态 IN(1,3) And A.NO=[1] And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNo)
    If rsTemp.EOF Then
        Screen.MousePointer = 0
        MsgBox "没有找到与号码""" & strNo & """相关的收费记录。", vbInformation, gstrSysName
        mCurBillType.lng病人ID = 0
        Exit Function
    End If
    txtPatient.Text = Nvl(rsTemp!姓名)

    lblPati.Caption = "病人:" & IIf(txtPatient.Visible, "                       ", rsTemp!姓名) & _
        "　性别:" & Nvl(rsTemp!性别) & _
        "　年龄:" & Nvl(rsTemp!年龄) & _
        "　门诊号:" & Nvl(rsTemp!标识号) & _
        "　费别:" & Nvl(rsTemp!费别) & _
        "　付款方式:" & rsTemp!付款方式

    With mCurBillType
        .lng病人ID = Val(Nvl(rsTemp!病人ID))
        .str性别 = Nvl(rsTemp!性别)
        .str年龄 = Nvl(rsTemp!年龄)
        .str姓名 = Nvl(rsTemp!姓名)
    End With

    If Not IsNull(rsTemp!险类) Then
        lblPati.ForeColor = vbRed
        txtYB.Text = Val(Nvl(rsTemp!险类))   '问题:41760
        txtPatient.ForeColor = vbRed
    Else
        lblPati.ForeColor = &HC00000
        txtYB.Text = ""
        txtPatient.ForeColor = &HC00000
    End If
    Call SetPatiColor(txtPatient, Nvl(rsTemp!病人类型), txtPatient.ForeColor)
    lblPati.ForeColor = txtPatient.ForeColor
    If mblnPrintView And zlStr.IsHavePrivs(mstrPrivs, "修改姓名重打") _
        And mCurBillType.lng病人ID = 0 Then
        txtPatientPrint.Text = "" & rsTemp!姓名
        txtPatientPrint.Tag = txtPatientPrint.Text
        txtPatientPrint.Visible = True
        lblPatiName.Visible = True
    End If

    If mblnDelByNo Then
        If mrsBalance.RecordCount > 0 Then
            mblnHistoryData = zlGetInsureBalanceDetail(1, Val(Nvl(mrsBalance!结算序号))).RecordCount = 0
        End If
        mblnDelByNo = Not mblnHistoryData
    End If

    '读取结算内容:原始或退费的,结算方式为空指冲预交的记录
    '----------------------------------------------------------------------------------
    Call LoadBalanceInfor
    Call LoadInvoiceData(strNos)

    If GetFeeListData(strNos, rsTemp) = False Then
        Call ClearFace(False)
        Exit Function
    End If

    str医嘱序号 = ""
    If rsTemp.EOF Then
        Screen.MousePointer = 0
        MsgBox "没有找到与号码""" & strNo & """相关的可以退费的记录。" & _
            vbCrLf & "这些收费记录可能已经退费或已经完全执行。", vbInformation, gstrSysName
        Call ClearFace(False)
        Exit Function
    End If

    mCurBillType.strNosOverFlow = ""
    strTmp = ""
    For i = 0 To UBound(Split(strNos, ","))
        strTmp = Replace(Split(strNos, ",")(i), "'", "")
        '检查是否金额超过上限
        If Not BillOperCheck(2, rsTemp!操作员姓名, rsTemp!登记时间, "退费", strTmp, , 1, True) Then
            mCurBillType.strNosOverFlow = mCurBillType.strNosOverFlow & " ," & strTmp
        End If
    Next
    If mCurBillType.strNosOverFlow <> "" Then mCurBillType.strNosOverFlow = Mid(mCurBillType.strNosOverFlow, 2)
    
    With vsBill
        .Redraw = flexRDNone
        .Rows = .FixedRows + rsTemp.RecordCount
        mCurBillType.strNos = ""
        For i = 1 To rsTemp.RecordCount
            '问题:29201
            .Cell(flexcpData, i, .ColIndex("项目")) = Nvl(rsTemp!从属父号)
            '问题:33634
            .Cell(flexcpData, i, .ColIndex("结帐ID")) = Nvl(rsTemp!医嘱序号) & "," & Nvl(rsTemp!收费细目ID)
            If Val(Nvl(rsTemp!医嘱序号)) <> 0 And InStr(str医嘱序号 & ",", "," & Nvl(rsTemp!医嘱序号) & ",") = 0 Then
                str医嘱序号 = str医嘱序号 & "," & Nvl(rsTemp!医嘱序号)
            End If
            
            strTemp = ""
            If Val(Nvl(rsTemp!从属父号)) <> 0 Then
                rsTemp.MoveNext
                strTemp = "┣"
                If rsTemp.EOF Then
                    strTemp = "┗"
                ElseIf Val(.Cell(flexcpData, i, .ColIndex("项目"))) <> Nvl(rsTemp!从属父号) Then
                    strTemp = "┗"
                End If
                rsTemp.MovePrevious
                strTemp = "  " & strTemp & " "
            End If

            .RowData(i) = CLng(rsTemp!序号)
            .TextMatrix(i, .ColIndex("选择")) = 0
            For j = 1 To cllInvoiceNoInfor.Count
                If cllInvoiceNoInfor(j)(0) = Nvl(rsTemp!NO) Then
                    If InStr(1, "," & cllInvoiceNoInfor(j)(1) & ",", "," & Nvl(rsTemp!序号) & ",") > 0 Then
                         .TextMatrix(i, .ColIndex("选择")) = 1: Exit For
                    End If
                End If
            Next
            .TextMatrix(i, .ColIndex("单据号")) = rsTemp!NO
            .TextMatrix(i, .ColIndex("类别")) = rsTemp!类别名
            .Cell(flexcpData, i, .ColIndex("类别")) = Nvl(rsTemp!类别码)
            .TextMatrix(i, .ColIndex("项目")) = strTemp & rsTemp!名称 & IIf(IsNull(rsTemp!规格), "", " " & rsTemp!规格)
            .TextMatrix(i, .ColIndex("商品名")) = strTemp & Nvl(rsTemp!商品名)
            .TextMatrix(i, .ColIndex("数量")) = FormatEx(Nvl(rsTemp!付数, 1) * rsTemp!数次, 5)
            .Cell(flexcpData, i, .ColIndex("数量")) = Nvl(rsTemp!付数, 1) * Val(Nvl(rsTemp!数次))
            
            .TextMatrix(i, .ColIndex("单位")) = Nvl(rsTemp!计算单位)
            .TextMatrix(i, .ColIndex("单价")) = Format(Val(Nvl(rsTemp!单价)), gstrFeePrecisionFmt)
            .TextMatrix(i, .ColIndex("应收金额")) = Format(Val(Nvl(rsTemp!应收金额)), gstrDec)
            .TextMatrix(i, .ColIndex("实收金额")) = Format(Val(Nvl(rsTemp!实收金额)), gstrDec)
            .TextMatrix(i, .ColIndex("开单科室")) = Nvl(rsTemp!开单科室)
            .TextMatrix(i, .ColIndex("执行科室")) = Nvl(rsTemp!执行科室)
            .TextMatrix(i, .ColIndex("操作员")) = rsTemp!操作员姓名
            .TextMatrix(i, .ColIndex("时间")) = Format(rsTemp!登记时间, "MM-dd HH:mm")
            .TextMatrix(i, .ColIndex("结帐ID")) = rsTemp!结帐ID
            .TextMatrix(i, .ColIndex("医嘱")) = Nvl(rsTemp!医嘱内容)
            .TextMatrix(i, .ColIndex("原始数量")) = Val(Nvl(rsTemp!原始数量))
            .TextMatrix(i, .ColIndex("准退数量")) = Val(Nvl(rsTemp!准退数量))
            .TextMatrix(i, .ColIndex("医嘱序号")) = Nvl(rsTemp!医嘱序号)
            .TextMatrix(i, .ColIndex("执行科室ID")) = Nvl(rsTemp!执行部门ID)
            
            If Not mCurBillType.bln单张部分退费 Then mCurBillType.bln单张部分退费 = RoundEx(Val(Nvl(rsTemp!原始数量)), 7) <> RoundEx(Val(Nvl(rsTemp!准退数量)), 7)
            If Not mCurBillType.bln单张部分退费 Then mCurBillType.bln单张部分退费 = Val(Nvl(rsTemp!记录标志)) > 1
            
            .Cell(flexcpData, i, .ColIndex("选择")) = Val(Nvl(rsTemp!记录标志))    '用于判断是否被销帐过,>1表示已销帐
            
            If Val(Nvl(rsTemp!记录标志)) > 1 And InStr(1, mCurBillType.strNosPatiDel & ",", "," & rsTemp!NO & "") = 0 Then mCurBillType.strNosPatiDel = mCurBillType.strNosPatiDel & "," & rsTemp!NO
            If InStr(mCurBillType.strNos & ",", "," & rsTemp!NO & ",") = 0 Then
                '画出分隔线
                If mCurBillType.strNos <> "" Then
                    .Select i, .FixedCols, i, .COLS - 1
                    .CellBorder vbBlue, 0, 1, 0, 0, 0, 0
                End If
                mCurBillType.strNos = mCurBillType.strNos & "," & rsTemp!NO
            End If
            dbl合计 = dbl合计 + Val(Nvl(rsTemp!实收金额))
            rsTemp.MoveNext
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
    
    If Not mCurBillType.bln单张部分退费 Then
        mCurBillType.bln单张部分退费 = zlExistDelFeeChargeBill(mCurBillType.strAllNOs)
    End If
    
    Call SetpicInvoiceVisible   '设置发票控件的显示
    
    If mCurBillType.strNos <> "" Then mCurBillType.strNos = Mid(mCurBillType.strNos, 2)
    
    
    txtAllTotal.Text = Format(dbl合计, gstrDec)
    If strInvoiceNO <> "" Then
        vsBill.Cell(flexcpChecked, 1, vsBill.ColIndex("选择"), vsBill.Rows - 1, vsBill.ColIndex("选择")) = 0
        Call FromInvoiceSelectNO(strInvoiceNO)
        Call SelectRelatingInvoice(strInvoiceNO, True)
        '仅显示被勾选的发票
        Call ShowAndHideDelBillRow
    End If
    '78596,冉俊明,2014-10-15,默认勾选单据
    If mlng结算序号 = 0 Then
        '87489
        If gTy_Module_Para.byt退费缺省选择方式 = 0 Then
            blnFind = False
            With vsBill
                For i = .FixedRows To .Rows - 1
                    If .TextMatrix(i, .ColIndex("单据号")) = strOldNO Then
                        .Row = i: blnFind = True
                        Exit For
                    End If
                Next
            End With
            If blnFind Then Call cmdBillSel_Click
        End If
    End If
'    '78596,冉俊明,2014-10-14,默认勾选单据
'    If InStr(";" & mstrPrivs & ";", ";部份退费;") = 0 Then Call cmdSelAll_Click
    If gTy_Module_Para.byt退费缺省选择方式 = 1 Then
        Call cmdSelAll_Click
    End If
    
    Call LoadBalanceInfor
    Call ReCalcDelMoney
    Call FromNoSelectInvoice
    
    Call SetFunCtrlVisible
    
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

Private Function CheckPrivsIsValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查操作员是否具备操作退费单
    '返回:具备返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-06-26 16:31:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    If Not (mbytMode = EM_MULTI_退费 Or mbytMode = EM_MULTI_异常重退) Then CheckPrivsIsValied = True: Exit Function

    '检查权限是否满足
    If mCurBillType.intInsure > 0 Then
        '保险退费权限检查
        If zlStr.IsHavePrivs(mstrPrivs, "保险收费") = False Then
            Screen.MousePointer = 0
            MsgBox "你没有权限对医保病人的单据退费！", vbInformation, gstrSysName
            Exit Function
        End If
        CheckPrivsIsValied = True: Exit Function
    End If
    
    '普通病人的处理
    '是否有非医保病人的退费权限
    If zlStr.IsHavePrivs(mstrPrivs, "允许非医保病人") = False Then
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
 
Private Sub cmdBillSel_Click()
    Dim i As Long
    With vsBill
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("单据号")) <> "" And _
               .TextMatrix(i, .ColIndex("单据号")) = .TextMatrix(.Row, .ColIndex("单据号")) And InStr(1, mCurBillType.strNosOverFlow, vsBill.TextMatrix(i, .ColIndex("单据号"))) <= 0 Then
                .TextMatrix(i, .ColIndex("选择")) = -1
            End If
        Next
    End With
    Call FromNoSelectInvoice
    Call ShowAndHideDelBillRow
    Call ReCalcDelMoney
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
    
    If mCurBillType.strNos <> "" And txtNO.Visible Then
        Call ClearFace
        txtNO.SetFocus
    Else
        Unload Me
    End If
End Sub
Private Function FromNOSelect(ByVal strNo As String, ByVal blnSel As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:按单据全选或全清单据
    '入参:strNO-指定的NO
    '     blnSel:true表示全选,否则全清
    '返回:设置成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-08-05 11:06:51
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
    FromNOSelect = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Sub ExecuteModifyPatiName()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行病人信息修改
    '编制:刘兴洪
    '日期:2014-07-03 17:00:55
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPro As New Collection
    Dim strSQL As String, arrNo As Variant, i As Long
    
    arrNo = Split(mCurBillType.strNos, ",")
    For i = 0 To UBound(arrNo)
        strSQL = "Zl_病人费用记录_Update('" & arrNo(i) & "',1,null,null,'" & txtPatientPrint.Text & "')"
        zlAddArray cllPro, strSQL
    Next

    On Error GoTo errHandle:
    zlExecuteProcedureArrAy cllPro, Me.Caption
    Exit Sub
errHandle:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Sub

Private Sub cmdClear_Click()
    Dim i As Long, j As Long
    
    With vsBill
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, .ColIndex("选择")) = 0
        Next
    End With
    
    With vsInvoice
        If .Visible Then
            If .Rows - 1 >= 0 And .COLS - 1 >= 1 Then
                For i = .FixedRows To .Rows - 1
                    For j = .FixedCols To .COLS - 1
                        If Trim(.TextMatrix(i, j)) <> "" Then .Cell(flexcpChecked, i, j) = 2
                    Next
                Next
            End If
        End If
    End With
    
    Call ShowAndHideDelBillRow
    Call ReCalcDelMoney
End Sub
 


Private Sub cmdOK_Click()
    If mbytMode = EM_MULTI_查看 Then Unload Me: Exit Sub
    
    If mbytMode = EM_MULTI_异常重退 Then
        '异常单据重新退费
        If ExecuteReDelFee = False Then
            '重新加载异常数据,以便读取正确的结帐数据
            Call LoadViewBills(mlng结算序号)
            Exit Sub
        End If
        mblnOK = True
        Unload Me: Exit Sub
    End If
    If mbytMode = EM_MULTI_退费申请 Or mbytMode = EM_MULTI_取消申请 _
        Or mbytMode = EM_MULTI_退费审核 Or mbytMode = EM_MULTI_取消审核 Then
        If SaveDelApplied(mbytMode) = False Then Exit Sub
        mblnOK = True
        Unload Me: Exit Sub
    End If
    Call ExecDelete
End Sub
Private Sub cmdSelAll_Click()
    Dim i As Long
    With vsBill
        For i = .FixedRows To .Rows - 1
            If .TextMatrix(i, .ColIndex("单据号")) <> "" And InStr(1, mCurBillType.strNosOverFlow, vsBill.TextMatrix(i, .ColIndex("单据号"))) <= 0 Then
                .TextMatrix(i, .ColIndex("选择")) = -1
            End If
        Next
    End With
    Call ReCalcDelMoney
    Call FromNoSelectInvoice
    Call ShowAndHideDelBillRow
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    
    If mlng结算序号 <> 0 Then  '指定了结算数据的
        If LoadViewBills(mlng结算序号) = False Then Unload Me: Exit Sub
    End If
    
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
 

Private Sub Form_Resize()
    Dim staH As Long

    On Error Resume Next

    staH = IIf(stbThis.Visible, stbThis.Height, 0)
    
    vsBill.Height = Me.ScaleHeight - picCmd.Height - staH - picPati.Height - picMoney.Height - pic退费摘要.Height - vsBalance.Height - IIf(picInvoice.Visible, picInvoice.Height, 0)
    
    If Me.ScaleWidth - cmdSelAll.Left - cmdCancel.Width > 5500 Then
        cmdCancel.Left = Me.ScaleWidth - cmdSelAll.Left - cmdCancel.Width
    Else
        cmdCancel.Left = 5500
    End If
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 90
    If mbytMode = EM_MULTI_退费审核 Then
        cmdRefuseApply.Left = cmdOK.Left - cmdRefuseApply.Width - 90
    End If


    fraInfo_1.Width = Me.ScaleWidth + 300
    LineCmd_1.x2 = Me.ScaleWidth + 300

    With txt退款合计
        .Left = Me.ScaleWidth - .Width - 100
        lbl退款合计.Left = .Left - lbl退款合计.Width - 20
    End With
    Me.Refresh
End Sub

Private Sub Form_Unload(Cancel As Integer)

    mbytMode = EM_MULTI_查看
    mstrNo = "": mstrDelTime = "": mCurBillType.strNosOverFlow = ""
    mstrApplyTime = ""
    mblnHistoryData = False: mblnDelByNo = False
    mblnNOMoved = False   '查看时,可能传入true
    Call initCardSquareData
    Call CloseIDCard
    zlDatabase.SetPara "退费号码输入模式", IIf(optNO(0).Value, "0", "1"), glngSys, 1121, InStr(1, mstrPrivs, ";参数设置;") > 0
    Call SaveWinState(Me, App.ProductName, mstrTittle)
    
    If Not mrs结算方式 Is Nothing Then Set mrs结算方式 = Nothing
    If Not mrs收费对照 Is Nothing Then Set mrs收费对照 = Nothing
    If Not mrsDelInvoice Is Nothing Then Set mrsDelInvoice = Nothing
    If Not mrsBalance Is Nothing Then Set mrsBalance = Nothing
    If Not mrsInsureBalance Is Nothing Then Set mrsInsureBalance = Nothing
    If Not mrsInfo Is Nothing Then Set mrsInfo = Nothing
    If Not mobjDrugPacker Is Nothing Then Set mobjDrugPacker = Nothing
    If Not mobjDrugMachine Is Nothing Then Set mobjDrugMachine = Nothing
    If Not mcllForceDelToCash Is Nothing Then Set mcllForceDelToCash = Nothing
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
        If mbytMode = EM_MULTI_退费审核 Then txt退费摘要.Left = 1600
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
        mstrNo = "" '78663,冉俊明,2014-10-15,输入病人ID方式查找单据成功后，再通过输入“票据号”查找不出单据
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
        If Val(.ColData(Col)) = 0 Then Cancel = True: Exit Sub
        .ColComboList(Col) = " ||" & Val(.Cell(flexcpData, Row, Col))
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
                                If .RowData(i) <> Val(.Cell(flexcpData, j, .ColIndex("项目"))) Then Exit For
                                .Cell(flexcpChecked, j, .ColIndex("选择")) = .Cell(flexcpChecked, i, .ColIndex("选择"))
                                .TextMatrix(j, .ColIndex("选择")) = .TextMatrix(i, .ColIndex("选择"))
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
        If Col <> .ColIndex("选择") Then Exit Sub
        If mbytMode = EM_MULTI_退费申请 Or mbytMode = EM_MULTI_取消申请 Or mbytMode = EM_MULTI_退费审核 Or mbytMode = EM_MULTI_取消审核 Then
            Call FromNOSelect(vsBill.TextMatrix(Row, .ColIndex("单据号")), Val(vsBill.TextMatrix(Row, .ColIndex("选择"))) <> 0)
            Call ReCalcDelMoney
            Exit Sub
        End If
        If mCurBillType.intInsure <> 0 And (MCPAR.按单据全退 Or mblnDelByNo) Then '86176
            Call FromNOSelect(vsBill.TextMatrix(Row, .ColIndex("单据号")), Val(vsBill.TextMatrix(Row, .ColIndex("选择"))) <> 0)
            Call ReCalcDelMoney
            '根据单据选择发票
            Call FromNoSelectInvoice
            Exit Sub
        End If
        
        stbThis.Panels(2).Text = ""
        If Val(.Cell(flexcpData, Row, .ColIndex("项目"))) = 0 Then
            For i = Row + 1 To .Rows - 1
                 If Val(.RowData(Row)) <> Val(.Cell(flexcpData, i, .ColIndex("项目"))) Then Exit For
                .TextMatrix(i, .ColIndex("选择")) = vsBill.TextMatrix(Row, .ColIndex("选择"))
            Next
            Call zlSet诊疗固定关系(Row, Col)
        Else
            Call zlSet诊疗固定关系(Row, Col)
            '需要检查主项是否已经被
            For i = Row - 1 To 1 Step -1
                If Val(.RowData(i)) = Val(.Cell(flexcpData, Row, .ColIndex("项目"))) Then
                    If .TextMatrix(i, .ColIndex("选择")) <> 0 Then
                         .TextMatrix(i, .ColIndex("选择")) = .TextMatrix(Row, .ColIndex("选择"))
                    End If
                    Call zlSet诊疗固定关系(i, Col, Row)
                     Exit For
                End If
            Next
        End If
        Call ReCalcDelMoney
        '根据单据选择发票
        Call FromNoSelectInvoice
    End With
End Sub
Private Sub vsBill_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim dbl合计 As Currency, i As Long
    If NewRow = OldRow Then Exit Sub
    With vsBill
        If Trim(.TextMatrix(NewRow, .ColIndex("单据号"))) = "" Then
            txtCurTotal.Text = Format(dbl合计, gstrDec)
            Exit Sub
        End If
        For i = NewRow - 1 To .FixedRows Step -1
            If .TextMatrix(i, .ColIndex("单据号")) <> .TextMatrix(NewRow, .ColIndex("单据号")) Then Exit For
            dbl合计 = dbl合计 + Val(.TextMatrix(i, .ColIndex("实收金额")))
        Next
        For i = NewRow To .Rows - 1
            If .TextMatrix(i, .ColIndex("单据号")) <> .TextMatrix(NewRow, .ColIndex("单据号")) Then Exit For
            dbl合计 = dbl合计 + Val(.TextMatrix(i, .ColIndex("实收金额")))
        Next
        txtCurTotal.Text = Format(dbl合计, gstrDec)
    End With
End Sub

Private Sub vsBill_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsBill
       If .Col = .ColIndex("选择") Then
            If .ColIndex("单据号") < 0 Then Cancel = True: Exit Sub
            If Trim(.TextMatrix(Row, .ColIndex("单据号"))) = "" Then Cancel = True
       Else
            Cancel = True
       End If
    End With
End Sub

Private Sub vsBill_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vsBill.ColIndex("选择") Then Cancel = True
End Sub

Private Sub GetBillRow(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据单据的指定行，获取单据的开始行和结速行
    '入参:lngRow-当前行
    '出参:lngBegin-单据的开始行
    '     lngEnd-单据的结束行
    '编制:刘兴洪
    '日期:2014-07-03 17:39:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    lngBegin = lngRow: lngEnd = lngRow
    With vsBill
        For i = lngRow - 1 To .FixedRows Step -1
            If .TextMatrix(i, .ColIndex("单据号")) <> .TextMatrix(lngRow, .ColIndex("单据号")) Then Exit For
            lngBegin = i
        Next
        For i = lngRow To .Rows - 1
            If .TextMatrix(i, .ColIndex("单据号")) <> .TextMatrix(lngRow, .ColIndex("单据号")) Then Exit For
            lngEnd = i
        Next
    End With
End Sub

Private Sub vsBill_CellChanged(ByVal Row As Long, ByVal Col As Long)
    With vsBill
        If .ColIndex("单据号") < 0 Then Exit Sub
        '超出限额的设置
        If .TextMatrix(Row, .ColIndex("单据号")) <> "" _
            And InStr(1, mCurBillType.strNosOverFlow, .TextMatrix(Row, .ColIndex("单据号"))) > 0 Then
             .TextMatrix(Row, .ColIndex("选择")) = 0
        End If
    End With
End Sub

Private Sub vsBill_DrawCell(ByVal hDC As Long, ByVal Row As Long, _
    ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, _
    ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:画线
    '编制:刘兴洪
    '日期:2014-07-03 17:41:52
    '说明：1.OwnerDraw要设置为Over(画出单元所有内容)
    '      2.Cell的GridLine从上下左右向内都是从第1根线开始
    '      3.Cell的Border从左上是从第2根线开始,右下是从第1根线开始
    '---------------------------------------------------------------------------------------------------------------------------------------------
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
            If Trim(.TextMatrix(.Row, .ColIndex("单据号"))) = "" Then Exit Sub
            
            If .TextMatrix(.Row, .ColIndex("选择")) = 0 _
                And InStr(1, mCurBillType.strNosOverFlow, .TextMatrix(.Row, .ColIndex("单据号"))) <= 0 Then
                 .TextMatrix(.Row, .ColIndex("选择")) = -1
            Else
                 .TextMatrix(.Row, .ColIndex("选择")) = 0
            End If
            Call ReCalcDelMoney
            Call FromNoSelectInvoice
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
        Select Case Col
        Case .ColIndex("选择")
        Case Else
             Cancel = True
        End Select
    End With
End Sub

Private Function CheckDelIsValied(ByVal strNos As String, _
    ByRef strNotCanDelNOs As String, _
    ByRef strCanDelNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查退费单据是否合法
    '出参:strNotCanDelNOs-不能退的单据(已经执行及不能退的单据)
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
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHandle

    '问题:54728
    If Not mbytMode = EM_MULTI_退费 Then CheckDelIsValied = True: Exit Function   '退费时判断

    arrNo = Split(strNos, ","): strNotCanDelNOs = ""
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
            Select Case intTmp
                Case 1 '该单据不存在
                    strInfo = strInfo & "指定的单据不存在！" & vbCrLf
                    Exit For
                Case 2 '已经全部完全执行(收费不考虑退费自动退药)
                    strInfo = strInfo & "[" & strCurNO & "]中的项目已经全部完全执行，不能退费！" & vbCrLf
                Case 3 '未完全执行部分剩余数量为0
                    strInfo = strInfo & "[" & strCurNO & "]中未完全执行的项目剩余数量为零，没有可退费用！" & vbCrLf
            End Select
        ElseIf blnHaveExe Then
            '存在已执行项目
            If mCurBillType.intInsure > 0 And (MCPAR.按单据全退 Or mblnDelByNo) Then '收费医保退费
                strInfo = strInfo & "[" & strCurNO & "]属于医保病人的收费单，存在已经执行的项目，不能退费！" & vbCrLf
                strNotCanDelNOs = strNotCanDelNOs & "," & strCurNO  '不能退的单据
            ElseIf gbln退费申请模式 Then
                '未申请或未审核的单据不能退费
                Set rsTemp = GetApply(strCurNO, 1)
                rsTemp.Filter = "状态<>2"
                If rsTemp.RecordCount = 0 Then
                    strInfo = strInfo & "[" & strCurNO & "]未进行退费申请及审核，不能进行退费！" & vbCrLf
                    strNotCanDelNOs = strNotCanDelNOs & "," & strCurNO  '不能退的单据
                ElseIf IsNull(rsTemp!审核人) Then
                    strInfo = strInfo & "[" & strCurNO & "]未进行退费审核，不能进行退费！" & vbCrLf
                    strNotCanDelNOs = strNotCanDelNOs & "," & strCurNO  '不能退的单据
                Else
                    strInfo = strInfo & "[" & strCurNO & "]中存在已执行的项目，此单据将执行的是部分退费。" & vbCrLf
                    strCanDelNos = strCanDelNos & "," & strCurNO
                End If
            Else
                strInfo = strInfo & "[" & strCurNO & "]中存在已执行的项目，此单据将执行的是部分退费。" & vbCrLf
                strCanDelNos = strCanDelNos & "," & strCurNO
            End If
        ElseIf gbln退费申请模式 Then
            '未申请或未审核的单据不能退费
            Set rsTemp = GetApply(strCurNO, 1)
            rsTemp.Filter = "状态<>2"
            If rsTemp.RecordCount = 0 Then
                strInfo = strInfo & "[" & strCurNO & "]未进行退费申请及审核，不能进行退费！" & vbCrLf
                strNotCanDelNOs = strNotCanDelNOs & "," & strCurNO  '不能退的单据
            ElseIf IsNull(rsTemp!审核人) Then
                strInfo = strInfo & "[" & strCurNO & "]未进行退费审核，不能进行退费！" & vbCrLf
                strNotCanDelNOs = strNotCanDelNOs & "," & strCurNO  '不能退的单据
            Else
                strCanDelNos = strCanDelNos & "," & strCurNO
            End If
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

Private Sub InitBalanceVar()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化单据
    '编制:刘兴洪
    '日期:2014-07-04 10:02:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str结算方式 As String
    
    If mrsBalance Is Nothing Then Exit Sub
    If mrsBalance.State <> 1 Then Exit Sub
    
    mrsBalance.Filter = "类型<>2 And 类型<>1"
    '       字段:类型,记录性质,结算方式,摘要,卡类别ID,卡类别名称,自制卡,结算卡序号,结算号码,卡号,交易流水号, 交易说明,结算序号,校对标志,医保,消费卡id
    '            是否密文,是否全退,是否退现,冲预交
    '       类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡)
    str结算方式 = ""
    mrsBalance.Sort = "类型,结算性质"
    With mrsBalance
        Do While Not .EOF
            If InStr(str结算方式 & ",", "," & Nvl(!结算方式) & ",") = 0 Then
                str结算方式 = str结算方式 & "," & Nvl(!结算方式)
            End If
            If Val(Nvl(!类型)) = 3 Or Val(Nvl(!类型)) = 4 Then mCurBillType.bln存在卡结算 = True
            .MoveNext
        Loop
    End With
    If str结算方式 <> "" Then str结算方式 = Mid(str结算方式, 2)
    mCurBillType.bln单种结算方式 = InStr(str结算方式, ",") = 0
    mCurBillType.str结算方式 = str结算方式
    
    '4-一卡通(老)
    mrsBalance.Filter = "类型=4"
    mCurBillType.blnExistOnCard = mrsBalance.EOF = False
    
    '3.一卡通
    mrsBalance.Filter = "类型=3 And  是否全退=1 and 是否退现=0"
    mCurBillType.blnExistThreeAllDel = mrsBalance.EOF = False
    mrsBalance.Filter = 0
End Sub

Private Function ExecuteClinicDelNo(ByVal lng病人ID As Long, ByVal intInsure As Integer, _
    ByVal lng冲销ID As Long, ByVal lng原结帐ID As Long, ByRef strAdvance As String, _
    Optional ByVal blnReDelete As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:按单据执行医保退费交易
    '入参:lng病人ID-病人ID
    '     intInsure-险类
    '     lng冲销ID-冲销ID
    '     lng原结帐ID-原始结帐ID
    '     strAdvance - 结算方式
    '     blnReDelete - 是否重新退费
    '出参:
    '返回:医保退费交易成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-31 23:38:11
    '说明:
    '   调用接口前,必须先打开事务,完成后,会自动提交事务;失败时,会回退事务
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strAllBalance As String, strSQL As String
    Dim varData As Variant, varTemp As Variant, i As Long, p As Integer, strAdvanceOld As String
    Dim colBalance As Collection '记录各张单据保险结算
    Dim strNo As String, strDel结算方式 As String
    Dim rsCharge As ADODB.Recordset, str结算方式 As String
    On Error GoTo errHandle
    
    strAdvance = lng冲销ID & "|" & "0"
    Set colBalance = New Collection
    strAdvanceOld = strAdvance
    
    '93337,退费时按单据号倒序进行接口调用
    strSQL = "Select Distinct NO From 门诊费用记录 Where 结帐id = [1] Order By No Desc"
    Set rsCharge = zlDatabase.OpenSQLRecord(strSQL, "获取本次冲销费用单据号", lng冲销ID)
    
    p = 1
    Do While Not rsCharge.EOF
        colBalance.Add Array()
        strDel结算方式 = ""
        strNo = Nvl(rsCharge!NO)
        '检查该单据是否已医保退费
        '如果调用成功过接口，但没有任何医保作废，则会再一次调用医保接口，因为无法确定是否调用成功过
        If blnReDelete Then
            strDel结算方式 = zlGetYBBalanceNo(lng冲销ID, strNo)
            Call SetBalanceVal(colBalance, p, strDel结算方式)
        End If
        
        str结算方式 = zlGetYBBalanceNo(lng原结帐ID, strNo, lng病人ID, intInsure, True)
        'str结算方式 为空，表示医保不支持医保作废
        If str结算方式 <> "" And strDel结算方式 = "" Then
            '    Zl_医保结算明细_Insert(
            strSQL = "Zl_医保结算明细_Insert("
            '      结帐id_In   医保结算明细.结帐id%Type,
            strSQL = strSQL & "" & lng冲销ID & ","
            '      No_In       医保结算明细.No%Type,
            strSQL = strSQL & "'" & strNo & "',"
            '      结算方式_In Varchar2,
            strSQL = strSQL & "'" & str结算方式 & "')"
            '      备注_In     医保结算明细.备注%Type := Null
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
            
            strAdvance = strAdvanceOld & "|" & strNo
            '因为参数固定为医保基金,所以名称固定为医保基金(多种统筹不好确定),以后应去掉该参数
            If Not gclsInsure.ClinicDelSwap(lng原结帐ID, True, intInsure, _
                                            strAdvance) Then Exit Function
            If strAdvance = strAdvanceOld & "|" & strNo Then strAdvance = ""
            
            If zlInsureCheck(str结算方式, strAdvance) Then
                str结算方式 = strAdvance
                '    Zl_医保结算明细_Insert(
                strSQL = "Zl_医保结算明细_Insert("
                '      结帐id_In   医保结算明细.结帐id%Type,
                strSQL = strSQL & "" & lng冲销ID & ","
                '      No_In       医保结算明细.No%Type,
                strSQL = strSQL & "'" & strNo & "',"
                '      结算方式_In Varchar2,
                strSQL = strSQL & "'" & strAdvance & "')"
                '      备注_In     医保结算明细.备注%Type := Null
                zlDatabase.ExecuteProcedure strSQL, Me.Caption
            End If
            gcnOracle.CommitTrans
            
            Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, intInsure)
            Call SetBalanceVal(colBalance, p, str结算方式)
            
            gcnOracle.BeginTrans
        End If
        
        p = p + 1
        rsCharge.MoveNext
    Loop

    '全部成功，返回总的结算方式
    strAdvance = GetMedicareStr(colBalance)
    
    ExecuteClinicDelNo = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ExecuteClinicDelSwap(ByVal lng病人ID As Long, ByVal intInsure As Integer, _
    ByVal lng冲销ID As Long, ByVal lng原结帐ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行医保退费交易
    '入参:lng病人ID-病人ID
    '     intInsure-险类
    '     lng冲销ID-冲销ID
    '     lng原结帐ID-原始结帐ID
    '出参:
    '返回:医保退费交易成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-31 23:38:11
    '说明:
    '   调用接口前,必须先打开事务,完成后,会自动提交事务;失败时,会回退事务
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strAdvance As String, strAllBalance As String, strSQL As String
    Dim varData As Variant, varTemp As Variant, i As Long, p As Integer, strAdvanceOld As String
    Dim colBalance As Collection '记录各张单据保险结算
    Dim rsBalance As ADODB.Recordset
    Dim blnDo As Boolean, strTemp As String, strNo As String
    Dim rsCharge As ADODB.Recordset
    On Error GoTo errHandle
    
    If intInsure = 0 Then ExecuteClinicDelSwap = True: gcnOracle.CommitTrans: Exit Function
    strAllBalance = GetYBOldBalance(lng病人ID, intInsure, lng原结帐ID)
    
    strAdvance = ""
    If MCPAR.门诊结算作废 Then
        
        If Not mblnDelByNo Then
            strAdvance = lng冲销ID & "|" & "0"
            'ClinicDelSwap (医保退费结算)
            '参数名  参数类型    入/出   原参数说明  现调整说明
            'lngStlID    long    IN  将要退费的费用记录的结帐ID(原结帐ID)
            'bln退费 Boolean IN  表明是退费交易还是改费交易在调用本接口
            'intInsure   Intger  In  险类
            'strAdvance  String  In  NULL    冲销ID:增加传入冲销ID
            '医保可以根据冲销ID来进行取数
            '        Out 退费结算：结算方式1|金额||结算方式2|金额...
            '    Boolean 函数返回    True:调用成功,False:调用失败
            If Not gclsInsure.ClinicDelSwap(lng原结帐ID, , intInsure, strAdvance) Then
                gcnOracle.RollbackTrans
                Exit Function
            End If
            If strAdvance = CStr(lng冲销ID) & "|" & "0" Then strAdvance = ""
        Else
            If ExecuteClinicDelNo(lng病人ID, intInsure, lng冲销ID, lng原结帐ID, strAdvance) = False Then
                gcnOracle.RollbackTrans
                Exit Function
            End If
        End If
    Else
        strAdvance = strAllBalance
        varData = Split(strAdvance, "||")
        strAdvance = ""
        For i = 0 To UBound(varData)
            varTemp = Split(varData(i) & "|||", "|")
            strAdvance = strAdvance & "||" & varTemp(0) & "|" & -1 * Val(varTemp(1))
        Next
        If strAdvance <> "" Then strAdvance = Mid(strAdvance, 3)
    End If
    
    If MCPAR.门诊结算作废 Then
        If Not zlInsureCheck(strAllBalance, strAdvance) Then
            '修改校对标志
            ' Zl_病人门诊收费_医保更新
            strSQL = "Zl_病人门诊收费_医保更新("
            '  结帐id_In   门诊费用记录.结帐id%Type,
            strSQL = strSQL & lng冲销ID & ","
            '  结算序号_In 病人预交记录.结算序号%Type,
            strSQL = strSQL & "Null,"
            '  保险结算_In Varchar2
            strSQL = strSQL & "Null)"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            gcnOracle.CommitTrans
            If Not mblnDelByNo Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, intInsure)
            ExecuteClinicDelSwap = True: Exit Function
        End If
        gcnOracle.CommitTrans: gcnOracle.BeginTrans
    End If
    '退费和收费不一致时,需要效对
        '增加结算方式为空的记录
        ' Zl_门诊退费结算_Modify
        strSQL = "Zl_门诊退费结算_Modify("
        '  操作类型_In   Number,
        '  --   0-原样退
        '  --      原样结算一起全退,所有校对标志都为1,医保调用成功后,调整为2,完成后变成0
        '  --   1-普通退费方式:
        '  --     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
        '  --     ②冲预交_In:如果涉及预交款,则传入本次的退预交,非正常收费时,传入零(<0 表示退预交款;>0 表示将剩余款生成预交记录
        '  --   2.三方卡退费结算:
        '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
        '  --     ②退预交_In: 传入零
        '  --     ③卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
        '  --   3-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
        '  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
        '  --     ②退预交_In: 传入零
        '  --     ③退支票额_In:传入零
        '  --   4-消费卡结算:
        '  --     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."
        '  --     ②退预交_In: 传入零
        '  --     ③退支票额_In:传入零
        strSQL = strSQL & "" & 3 & ","
        '  病人id_In     门诊费用记录.病人id%Type,
        strSQL = strSQL & "" & lng病人ID & ","
        '  冲销id_In     病人预交记录.结帐id%Type,
        strSQL = strSQL & "" & lng冲销ID & ","
        '  结算方式_In   Varchar2,
        strSQL = strSQL & "'" & strAdvance & "')"
        '  退预交_In     病人预交记录.冲预交%Type := Null,
        '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
        '  卡号_In       病人预交记录.卡号%Type := Null,
        '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
        '  交易说明_In   病人预交记录.交易说明%Type := Null,
        '  缴款_In       病人预交记录.缴款%Type := Null,
        '  找补_In       病人预交记录.找补%Type := Null,
        '  误差金额_In   门诊费用记录.实收金额%Type := Null,
        '  完成退费_In   Number := 0,
        '  原结帐id_In   病人预交记录.结帐id%Type := Null
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    '修改校对标志
    ' Zl_病人门诊收费_医保更新
    strSQL = "Zl_病人门诊收费_医保更新("
    '  结帐id_In   门诊费用记录.结帐id%Type,
    strSQL = strSQL & lng冲销ID & ","
    '  结算序号_In 病人预交记录.结算序号%Type,
    strSQL = strSQL & "Null,"
    '  保险结算_In Varchar2
    strSQL = strSQL & "Null)"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    gcnOracle.CommitTrans
    If MCPAR.门诊结算作废 Then
        If Not mblnDelByNo Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, True, intInsure)
    End If
    ExecuteClinicDelSwap = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, False, mCurBillType.intInsure)
End Function

Private Function ExecuteOneCardDelInterface(ByVal rsBalance As ADODB.Recordset, _
        ByVal lng冲销ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:一卡通退费(旧)接口
    '入参:lng冲销ID-冲销ID
    '     rsBalance-原结算记录集
    '返回:执行成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-08-01 15:45:26
    '说明:调用本接口前，必须开通事务,完成或异常都会终止事务
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCardNo As String, strSwap As String, strHsptCode As String '医院编码
    Dim i As Long, dblMoney As Double, strNos As String, strSQL As String
    Dim str结算方式 As String
    
    '字段:类型,记录性质,结算方式,摘要,卡类别ID,卡类别名称,自制卡,结算卡序号,结算号码,卡号,交易流水号, 交易说明,结算序号,校对标志,医保,消费卡id
    '     是否密文,是否全退,是否退现,冲预交
    '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡)
    On Error GoTo errHandle
    rsBalance.Filter = "性质=4"
    If rsBalance.RecordCount = 0 Then
        rsBalance.Filter = 0
        gcnOracle.CommitTrans
        ExecuteOneCardDelInterface = True: Exit Function
    End If
    
    '一卡通(旧):只能使用一种
    With rsBalance
        .MoveFirst
        Do While Not .EOF
            dblMoney = dblMoney + Val(Nvl(!冲预交))
            .MoveNext
        Loop
        dblMoney = RoundEx(dblMoney, 6)
        .MoveFirst
        If dblMoney = 0 Then
            rsBalance.Filter = 0: gcnOracle.CommitTrans
            ExecuteOneCardDelInterface = True: Exit Function
        End If
        strCardNo = Nvl(!卡号)
        str结算方式 = Nvl(!结算方式)
        '结算方式|结算金额|结算号码|结算摘要||..
        str结算方式 = str结算方式 & "|" & -1 * dblMoney
        str结算方式 = str结算方式 & "|" & IIf(Trim(Nvl(!结算号码)) = "", " ", Trim(Nvl(!结算号码)))
        str结算方式 = str结算方式 & "| "
        
        'Zl_门诊退费结算_Modify
        '--操作类型_In:
        '--   0-原样退
        '--      原样结算一起全退,所有校对标志都为1,医保调用成功后,调整为2,完成后变成0
        '--   1-普通退费方式:
        '--     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
        '--   2.三方卡退费结算:
        '--     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
        '--     ②卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
        '--   3-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
        '--     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
        '--     ②退支票额_In:传入零
        '--   4-消费卡结算:
        '--     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."
        '--     ②退支票额_In:传入零
        strSQL = "Zl_门诊退费结算_Modify("
        '  操作类型_In   Number,
        strSQL = strSQL & "" & 2 & ","
        '  病人id_In     门诊费用记录.病人id%Type,
        strSQL = strSQL & "" & mCurBillType.lng病人ID & ","
        '  冲销id_In     病人预交记录.结帐id%Type,
        strSQL = strSQL & "" & lng冲销ID & ","
        '  结算方式_In   Varchar2,
        strSQL = strSQL & "'" & str结算方式 & "',"
        '  退预交_In     病人预交记录.冲预交%Type := Null,
        strSQL = strSQL & "NULL,"
        '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
        strSQL = strSQL & "NULL,"
        '  卡号_In       病人预交记录.卡号%Type := Null,
        strSQL = strSQL & "'" & strCardNo & "',"
        '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
        strSQL = strSQL & "'" & Nvl(!交易流水号) & "',"
        '  交易说明_In   病人预交记录.交易说明%Type := Null,
        strSQL = strSQL & "'" & Nvl(!交易说明) & "')"
        '  缴款_In       病人预交记录.缴款%Type := Null,
        '  找补_In       病人预交记录.找补%Type := Null,
        '  误差金额_In   门诊费用记录.实收金额%Type := Null,
        '  完成退费_In   Number := 0,
        '  原结帐id_In   病人预交记录.结帐id%Type := Null
    End With
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    If Not mobjICCard.ReturnSwap(strCardNo, strHsptCode, strSwap, dblMoney) Then
        gcnOracle.RollbackTrans
        MsgBox "一卡通退费交易调用失败,不能继续退费操作！", vbExclamation, gstrSysName
        Exit Function
    End If
    gcnOracle.CommitTrans
    ExecuteOneCardDelInterface = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call SaveErrLog
End Function

Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:进行退费是否合法
    '返回:合法返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-04 11:23:03
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim arrNo As Variant, strYPNos As String, bln药品 As Boolean, blnSel As Boolean
    Dim i As Long, strDelNOs As String, strNo As String, str操作员姓名 As String
    Dim varTemp As Variant, rsTemp As ADODB.Recordset
    On Error GoTo errHandle
 
    '检查输入是否正确
    If mCurBillType.strNos = "" Then
        MsgBox "请先输入要退费的单据。", vbInformation, gstrSysName
        If txtNO.Visible Then txtNO.SetFocus: Exit Function
    End If
    
    '检查本次结算单据中是否存在退费异常单据，若存在，则不允许继续退费
    If CheckIsExistDelErrBill(mCurBillType.strNos, str操作员姓名) Then
        MsgBox "注意：" & vbCrLf & _
            "    本次结算中存在异常的退费记录，请先对其进行重新退费！" & _
            IIf(str操作员姓名 <> UserInfo.姓名, vbCrLf & "    提示：异常单据是操作员【" & str操作员姓名 & "】收取的。", ""), vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    arrNo = Split(mCurBillType.strNos, ",")
    strYPNos = "": strDelNOs = ""
    bln药品 = False: blnSel = False
    With vsBill
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("选择"))) <> 0 Then
                blnSel = True
                strNo = .TextMatrix(i, .ColIndex("单据号"))
                If InStr(strDelNOs & ",", "," & strNo & ",") = 0 Then
                    strDelNOs = strDelNOs & "," & strNo
                End If
                
                If .ColIndex("类别") <> -1 And bln药品 = False Then     '47400
                    If .TextMatrix(i, .ColIndex("类别")) Like "*西*药*" _
                        Or .TextMatrix(i, .ColIndex("类别")) Like "*中*药*" _
                        Or .TextMatrix(i, .ColIndex("类别")) Like "*卫材*" Then
                        If InStr(strYPNos & ",", "," & strNo & ",") = 0 Then
                            strYPNos = strYPNos & "," & strNo
                        End If
                        bln药品 = True
                    End If
                End If
            End If
        Next
    End With
    If strDelNOs <> "" Then strDelNOs = Mid(strDelNOs, 2)
    
    If strDelNOs <> "" And gbln退费申请模式 Then
        Set rsTemp = GetApply(strDelNOs, 1)
        varTemp = Split(strDelNOs, ",")
        For i = 0 To UBound(varTemp)
            strNo = varTemp(i)
            rsTemp.Filter = "NO='" & strNo & "' And 状态<>2"
            If rsTemp.RecordCount = 0 Then
                Screen.MousePointer = 0
                MsgBox "请先对单据:" & strNo & " 进行退费申请！", vbInformation, gstrSysName
                Exit Function
            End If
            If IsNull(rsTemp!审核人) Then
                Screen.MousePointer = 0
                MsgBox "单据:" & strNo & " 未进行退费审核，不能进行退费！", vbInformation, gstrSysName
                Exit Function
            End If
        Next
    End If
    
    If blnSel = False Then
        MsgBox "请在单据中至少选择一个要退费的项目。", vbInformation, gstrSysName
        vsBill.SetFocus: Exit Function
    End If
    
    If bln药品 Then
        If strYPNos <> "" Then strYPNos = Mid(strYPNos, 2)
        If zlCheckDrugIsPutDrug(strYPNos) = False Then Exit Function
    End If
    
    '医保检查
    If mCurBillType.intInsure <> 0 Then
        If gclsInsure.CheckInsureValid(mCurBillType.intInsure) = False Then Exit Function
    End If
    
    If zlCheckIsMzToZY(strDelNOs, 1) Then
          MsgBox "注意:" & vbCrLf & _
            "    该单据已经被门诊费用转住院费用 " & vbCrLf & _
            "    或已经审核了门诊费用转住院费用,不能再退费", vbInformation + vbOKOnly, gstrSysName
          Exit Function
    End If
    
    If CheckBillExistReplenishData(0, mlng结算序号) = True Then
        MsgBox "选择的退费记录进行了医保补充结算，不允许进行退费操作！", vbInformation, gstrSysName
        Exit Function
    End If
    
    '105432,三方卡结算方式有效性检查
    If ThreeBalanceCheck(mrsBalance, mrs结算方式, mcllForceDelToCash) = False Then Exit Function
    
    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function ThreeBalanceCheck(ByVal rsBalance As ADODB.Recordset, ByVal rs结算方式 As ADODB.Recordset, _
    ByRef cllForceDelToCash As Collection) As Boolean
    '三方卡结算方式有效性检查
    '入参：
    '   rsBalance 结算数据
    '   rs结算方式 “收费”场合的所有结算方式
    '出参：
    '   cllForceDelToCash 强制退现信息：Array(操作员,卡类别名称)
    '返回：检查通过，返回True；否则，返回False
    '105432
    Dim objCards As Cards, objCard As Card
    Dim cllFeeBalance As New Collection, i As Integer
    Dim blnFind As Boolean, blnQuestion As Boolean
    Dim str操作员 As String, strKey As String
    Dim dblMoney  As Double
    
    On Error GoTo errHandler
    Set cllForceDelToCash = New Collection
    If rsBalance Is Nothing Then ThreeBalanceCheck = True: Exit Function
    
    '类型：0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
    rsBalance.Filter = "类型= 3"
    '去重
    With rsBalance
        Do While Not .EOF
            strKey = "_" & Val(Nvl(!卡类别ID))
            If CollectionExitsValue(cllFeeBalance, strKey) Then
                dblMoney = cllFeeBalance(strKey)(4) + Val(Nvl(!冲预交))
                cllFeeBalance.Remove strKey
            Else
                dblMoney = Val(Nvl(!冲预交))
            End If
            If RoundEx(dblMoney, 6) > 0 Then '全部退完的就不再加入
                'Array(结算方式,卡类别ID,是否退现,卡类别名称,冲预交)
                cllFeeBalance.Add Array(Nvl(!结算方式), Val(Nvl(!卡类别ID)), Val(Nvl(!是否退现)), Nvl(!卡类别名称), dblMoney), strKey
            End If
            .MoveNext
        Loop
    End With
    If cllFeeBalance.Count = 0 Then ThreeBalanceCheck = True: Exit Function
    
    If Not gobjSquare Is Nothing Then
        ' zlGetCards(ByVal BytType As Byte)
        '   入参:bytType-  0-所有医疗卡;
        '                    1-启用的医疗卡,
        '                    2-所有存在三方账户的三方卡
        '                    3-启用的三方账户的医疗卡
        Set objCards = gobjSquare.objSquareCard.zlGetCards(3)
    End If
    For i = 1 To cllFeeBalance.Count
        blnQuestion = False
        '结算方式检查
        If rs结算方式 Is Nothing Then
            If MsgBox("结算方式『" & cllFeeBalance(i)(0) & "』未启用，该结算方式支付的金额将被退为其它结算方式，是否继续？", _
                vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            blnQuestion = True
        Else
            rs结算方式.Filter = "名称='" & cllFeeBalance(i)(0) & "'" '结算方式要设置了"费用"应用场合才能使用
            If rs结算方式.EOF Then
                If MsgBox("结算方式『" & cllFeeBalance(i)(0) & "』未启用，该结算方式支付的金额将被退为其它结算方式，是否继续？", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                blnQuestion = True
            End If
        End If
        
        If blnQuestion = False Then
            '医疗卡检查
            If objCards Is Nothing Then
                If MsgBox("『" & cllFeeBalance(i)(3) & "』未启用，该医疗卡支付的金额将被退为其它结算方式，是否继续？", _
                    vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                blnQuestion = True
            Else
                blnFind = False
                For Each objCard In objCards
                    If objCard.接口序号 = cllFeeBalance(i)(1) Then blnFind = True: Exit For
                Next
                If blnFind = False Then
                    If MsgBox("『" & cllFeeBalance(i)(3) & "』未启用，该医疗卡支付的金额将被退为其它结算方式，是否继续？", _
                        vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                    blnQuestion = True
                End If
            End If
        End If
        
        If blnQuestion And cllFeeBalance(i)(2) = 0 Then '强制退现
            If str操作员 = "" Then '多种卡类别时只验证一次
                If zlStr.IsHavePrivs(GetPrivFunc(glngSys, 1151), "三方退款强制退现") Then
                    str操作员 = UserInfo.姓名
                Else
                    str操作员 = zlDatabase.UserIdentifyByUser(Me, "『" & cllFeeBalance(i)(3) & "』强制退现，权限验证：", _
                        glngSys, mlngModule, "三方退款强制退现", , True)
                    If str操作员 = "" Then Exit Function
                End If
                'Array(操作员,卡类别名称)
                cllForceDelToCash.Add Array(str操作员, cllFeeBalance(i)(3))
            End If
        End If
    Next
    ThreeBalanceCheck = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function InputFactNo(ByRef lng领用ID As Long, ByRef strInvoice As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:输入有效的发票号
    '入参:
    '     lng领用ID-当前的领用ID
    '出参:返回的发票号
    '返回:输入成功，返回true,否则返回False
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnValid As Boolean
    
    On Error GoTo errHandle
    Do
        '根据票据领用读取
        blnValid = False
        
        If gblnStrictCtrl Then
            If zlCheckInvoiceValied(lng领用ID, 1, , mlngShareUseID, mstrUseType) = False Then Exit Function
            strInvoice = GetNextBill(lng领用ID)
        Else
            strInvoice = zlStr.Increase(UCase(zlDatabase.GetPara("当前收费票据号", glngSys, mlngModule)))
        End If
        
        If strInvoice = "" Then
            If frmInputBox.InputBox(Me, "开始发票号", "" & _
                 "请你输入将要使用的开始票据号码：", 30, 1, False, False, strInvoice, _
                False, Me.Left + 1500, Me.Top + 1500) = False Then Exit Function
        End If
                    
        '用户取消输入,终止操作
        If strInvoice = "" Then Exit Function
        If gblnStrictCtrl Then
            If zlCheckInvoiceValied(lng领用ID, 1, strInvoice, mlngShareUseID, mstrUseType) Then blnValid = True
        Else
            blnValid = True
        End If
    Loop While Not blnValid
    
    InputFactNo = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
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
    Dim strTemp As String, varTemp As Variant, cllPro As Collection
    Dim arrNo As Variant, k As Long
    Dim i As Long, j As Long, dtDelDate As Date, lngCount As Long
    Dim strNo As String, blnTrans As Boolean
    Dim colOrder As New Collection
    Dim lng结帐ID As Long, lng冲销ID As Long, lng结算序号 As Long
    Dim lng原结帐ID As Long, blnAll部份退费 As Boolean, blnCur部份退费 As Boolean
    Dim bln全退 As Boolean, lngCheck病人ID As Long, intCheckInsure As Integer
    Dim strYBPati As String, strPrintNOInfor As String, strInvoice As String
    Dim str序号  As String, strCurSelNos As String, strReclaimInvoice As String
    Dim strInvoices As String, lng领用ID As Long
    Dim strSQL As String, strCmdCaptions As String
    Dim bln原样退 As Boolean
    Dim cur个帐透支 As Currency, str保险金额 As String 'cur实收合计;cur进入统筹;cur全自付;cur先自付
    Dim strReturn As String, strReturnRecipt As String '退费处方信息，格式：NO,药房ID|NO,药房ID|…
    Dim strPartSelectNos As String '部分选择的单据
    Dim strPartDoNos As String '全选但存在部分执行的单据
    Dim bln分别打印 As Boolean

    If isValied = False Then Exit Function
 
    lng原结帐ID = mCurBillType.lng原结帐ID
    bln分别打印 = gTy_Module_Para.bln分别打印 And mblnOnePatiPrint = False
      
    On Error GoTo Errhand:
    '先判断所有单据是否部份退费,以决定票据的处理方式
    arrNo = Split(mCurBillType.strNos, ",")
    
    blnAll部份退费 = False
    strCurSelNos = ""
    Set cllPro = New Collection
    For i = 0 To UBound(arrNo)
        strNo = arrNo(i)
        str序号 = "":   lngCount = 0
        
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
                    '81190,冉俊明,退费业务向发药机上传退费信息
                    '格式：NO,药房ID|NO,药房ID|…
                    If vsBill.TextMatrix(j, vsBill.ColIndex("类别")) Like "*西*药*" _
                        Or vsBill.TextMatrix(j, vsBill.ColIndex("类别")) Like "*中*药*" Then
                        If InStr(strReturnRecipt & "|", _
                            "|" & vsBill.TextMatrix(j, vsBill.ColIndex("单据号")) & "," & vsBill.TextMatrix(j, vsBill.ColIndex("执行科室ID")) & "|") = 0 Then
                            strReturnRecipt = strReturnRecipt & "|" & vsBill.TextMatrix(j, vsBill.ColIndex("单据号")) & "," & vsBill.TextMatrix(j, vsBill.ColIndex("执行科室ID"))
                        End If
                    End If
                End If
                lngCount = lngCount + 1
            Next
        End With
        str序号 = Mid(str序号, 2)
        If str序号 <> "" Then
            strPrintNOInfor = strPrintNOInfor & ";" & strNo & ":" & str序号
            blnCur部份退费 = Not BillDeleteAllNew(strNo, 1)
            If blnCur部份退费 Then strPartDoNos = strPartDoNos & "," & strNo '存在部分执行的单据
            
            If UBound(Split(str序号, ",")) + 1 = lngCount And blnCur部份退费 = False Then str序号 = ""
            blnCur部份退费 = Not (Not blnCur部份退费 And str序号 = "")
            If blnCur部份退费 And str序号 <> "" Then strPartSelectNos = strPartSelectNos & "," & strNo '部分选择的单据
            
            If blnCur部份退费 Then blnAll部份退费 = True '这张单据为部份退费,则所有单据为部份退费
            colOrder.Add str序号, "_" & strNo
        Else
            blnAll部份退费 = True                       '这张单据不退费,则所有单据为部份退费
            colOrder.Add "未选择", "_" & strNo
        End If
    Next
    If strPartSelectNos <> "" Then strPartSelectNos = Mid(strPartSelectNos, 2)
    If strPartDoNos <> "" Then strPartDoNos = Mid(strPartDoNos, 2)
    
    '根据其它单据是否未退完,则可判断出所有单据是否部份退费
    If Not blnAll部份退费 Then
        varTemp = Split(mCurBillType.strAllNOs, ",")
        strTemp = ""
        For i = 0 To UBound(varTemp)
            If InStr(1, "," & mCurBillType.strNos & ",", "," & varTemp(i) & ",") = 0 Then
                strTemp = strTemp & "," & varTemp(i)
                 blnAll部份退费 = True: Exit For
            End If
        Next
    End If
    
    If CheckSelectItemCanDel(strCurSelNos) = False Then Exit Function
    
    If blnAll部份退费 Then
        If mCurBillType.intInsure > 0 And (MCPAR.按单据全退 Or mblnDelByNo) Then '86176
            If strPartSelectNos <> "" Then
                MsgBox "单据[" & strPartSelectNos & "]包含保险结算费用，不允许部份退费。", vbInformation, gstrSysName
                Exit Function
            ElseIf strPartDoNos <> "" Then
                MsgBox "单据[" & strPartDoNos & "]包含保险结算费用，而其中一些项目已经执行，不允许部份退费。", vbInformation, gstrSysName
                Exit Function
            End If
        End If
        
        '56963
        If gTy_Module_Para.byt票据分配规则 <> 0 And strReclaimInvoice = "" Then
            strReclaimInvoice = zlGetReclaimInvoice(Mid(strPrintNOInfor, 2))
        End If
        
        If Not (gTy_Module_Para.byt票据分配规则 <> 0 And strReclaimInvoice <> "") Then
            If zlStr.IsHavePrivs(mstrPrivs, "部份退费") = False Then
                MsgBox "你没有权限执行部份退费操作！", vbInformation, gstrSysName
                vsBill.SetFocus: Exit Function
            End If
            If gTy_Module_Para.bln工本费 Then
                MsgBox "自动收取工本费时不允许部份退费。", vbInformation, gstrSysName: vsBill.SetFocus: Exit Function
            End If
            '刘兴洪 问题:27352 日期:2010-01-13 10:26:08
            If zlStr.IsHavePrivs(mstrPrivs, "退费核收发票") Then
                
                If frmReInvoice.ShowMe(Me, strCurSelNos, Val(txtAllTotal.Text), Val(txt退款合计.Text), strInvoices) = False Then
                    vsBill.SetFocus: Exit Function
                End If
            End If
        End If
    End If
    
    If mCurBillType.intInsure <> 0 And MCPAR.医保接口打印票据 Then
        If InputFactNo(lng领用ID, strInvoice) = False Then Exit Function
    End If
    dtDelDate = zlDatabase.Currentdate
    bln全退 = CheckIsAllDel(mCurBillType.strAllNOs)
    bln原样退 = bln全退
    If bln原样退 Then
        bln原样退 = Not zlExistDelFeeChargeBill(mCurBillType.strAllNOs)
    End If
    
    '生成要执行的SQL
    lng冲销ID = zlDatabase.GetNextId("病人结帐记录")
    lng结算序号 = -1 * lng冲销ID
    mCurBillType.strDelNOs = ""
    For i = UBound(arrNo) To 0 Step -1
        strNo = arrNo(i)
        If bln分别打印 And gTy_Module_Para.byt票据分配规则 = 0 Then
            bln全退 = CheckIsAllDel(strNo)
        End If
        If colOrder("_" & strNo) <> "未选择" Then
            ' Zl_门诊收费记录_销帐
            strSQL = "Zl_门诊收费记录_销帐("
            '  No_In         门诊费用记录.No%Type,
            strSQL = strSQL & "'" & strNo & "',"
            '  操作员编号_In 门诊费用记录.操作员编号%Type,
            strSQL = strSQL & "'" & UserInfo.编号 & "',"
            '  操作员姓名_In 门诊费用记录.操作员姓名%Type,
            strSQL = strSQL & "'" & UserInfo.姓名 & "',"
            '  序号_In       Varchar2 := Null,
            strSQL = strSQL & "'" & colOrder("_" & strNo) & "',"
            '  退费时间_In   门诊费用记录.登记时间%Type := Null,
            strSQL = strSQL & "To_Date('" & Format(dtDelDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
            '  退费摘要_In   门诊费用记录.摘要%Type := Null,
            strSQL = strSQL & "" & IIf(Trim(txt退费摘要.Text) = "", "NULL", "'" & Trim(txt退费摘要.Text) & "'") & ","
            '  结帐id_In     病人预交记录.结帐id%Type := Null,
            strSQL = strSQL & lng冲销ID & ","
            '  回收票据_In Number:=0
            If bln全退 And mblnOnePatiPrint And gTy_Module_Para.byt票据分配规则 <> 0 Then
                '根据报表实际打印，需要回收票据(就不单独处理了,而其他方式，需要在后续单项奖独处理
                strSQL = strSQL & "0)" '按病人打印不进行回收票据,在后面处理
            Else
                strSQL = strSQL & "" & IIf(bln全退, "1", "0") & ")"
            End If
            zlAddArray cllPro, strSQL
            mCurBillType.strDelNOs = mCurBillType.strDelNOs & "," & strNo
        End If
    Next
    bln全退 = CheckIsAllDel(mCurBillType.strAllNOs)
    If mCurBillType.intInsure <> 0 And MCPAR.门诊结算作废 Then
        If Not mblnDelByNo Then
            If Not bln全退 Then lng结帐ID = zlDatabase.GetNextId("病人结帐记录")
            '只有医保结算,才会存在重新收取的情况
            'Zl_门诊收费记录_重收
            strSQL = "Zl_门诊收费记录_重收("
            '  原结帐id_In 门诊费用记录.结帐id%Type,
            strSQL = strSQL & "" & lng原结帐ID & ","
            '  冲销id_In   门诊费用记录.结帐id%Type,
            strSQL = strSQL & "" & lng冲销ID & ","
            '  重收结帐id_In 门诊费用记录.结帐id%Type
            strSQL = strSQL & "" & IIf(lng结帐ID = 0, "NULL", lng结帐ID) & ","
            '  排开医保结算_In Varchar2:=Null
            strSQL = strSQL & "'" & GetYBTOCash(mCurBillType.lng病人ID, mCurBillType.intInsure) & "')"
            zlAddArray cllPro, strSQL
            '调用医保接口
            '先回收票据，预结算之后再产生票据
            If MCPAR.医保接口打印票据 Then '81684
                If Not bln全退 Then '预结算之后再发出票据
                    '56963,77058
                    strSQL = "zl_门诊收费记录_RePrint('" & strNo & "',NULL," & ZVal(lng领用ID) & ",'" & UserInfo.姓名 & "'," & _
                        "To_Date('" & Format(dtDelDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),1,1)"
                    zlAddArray cllPro, strSQL
                ElseIf Not (gTy_Module_Para.byt票据分配规则 <> 0 And strReclaimInvoice <> "") Then  '全退费也要生存票据号，北京医保
                    strSQL = "zl_门诊收费记录_RePrint('" & strNo & "','" & strInvoice & "'," & ZVal(lng领用ID) & ",'" & UserInfo.姓名 & "'," & _
                        "To_Date('" & Format(dtDelDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),1,1)"
                    zlAddArray cllPro, strSQL
                End If
            End If
        Else
            ' Zl_门诊退费结算_Modify
            strSQL = "Zl_门诊退费结算_Modify("
            '  操作类型_In   Number,
            '  --   0-原样退
            '  --      原样结算一起全退,所有校对标志都为1,医保调用成功后,调整为2,完成后变成0
            '  --   1-普通退费方式:
            '  --     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
            '  --     ②冲预交_In:如果涉及预交款,则传入本次的退预交,非正常收费时,传入零(<0 表示退预交款;>0 表示将剩余款生成预交记录
            '  --   2.三方卡退费结算:
            '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
            '  --     ②退预交_In: 传入零
            '  --     ③卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
            '  --   3-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
            '  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
            '  --     ②退预交_In: 传入零
            '  --     ③退支票额_In:传入零
            '  --   4-消费卡结算:
            '  --     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."
            '  --     ②退预交_In: 传入零
            '  --     ③退支票额_In:传入零
            strSQL = strSQL & "" & 3 & ","
            '  病人id_In     门诊费用记录.病人id%Type,
            strSQL = strSQL & "" & mCurBillType.lng病人ID & ","
            '  冲销id_In     病人预交记录.结帐id%Type,
            strSQL = strSQL & "" & lng冲销ID & ","
            '  结算方式_In   Varchar2,
            strSQL = strSQL & "'" & zlGetYBBalanceNo(lng原结帐ID, mCurBillType.strDelNOs, mCurBillType.lng病人ID, mCurBillType.intInsure, True) & "')"
            '  退预交_In     病人预交记录.冲预交%Type := Null,
            '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
            '  卡号_In       病人预交记录.卡号%Type := Null,
            '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
            '  交易说明_In   病人预交记录.交易说明%Type := Null,
            '  缴款_In       病人预交记录.缴款%Type := Null,
            '  找补_In       病人预交记录.找补%Type := Null,
            '  误差金额_In   门诊费用记录.实收金额%Type := Null,
            '  完成退费_In   Number := 0,
            '  原结帐id_In   病人预交记录.结帐id%Type := Null
            zlAddArray cllPro, strSQL
        End If
    Else
        '增加结算方式为空的记录
        ' Zl_门诊退费结算_Modify
        strSQL = "Zl_门诊退费结算_Modify("
        '  操作类型_In   Number,
        '  --   0-原样退
        '  --      原样结算一起全退,所有校对标志都为1,医保调用成功后,调整为2,完成后变成0
        '  --   1-普通退费方式:
        '  --     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
        '  --     ②冲预交_In:如果涉及预交款,则传入本次的退预交,非正常收费时,传入零(<0 表示退预交款;>0 表示将剩余款生成预交记录
        '  --   2.三方卡退费结算:
        '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
        '  --     ②退预交_In: 传入零
        '  --     ③卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
        '  --   3-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
        '  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
        '  --     ②退预交_In: 传入零
        '  --     ③退支票额_In:传入零
        '  --   4-消费卡结算:
        '  --     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."
        '  --     ②退预交_In: 传入零
        '  --     ③退支票额_In:传入零
        strSQL = strSQL & "" & 1 & ","
        '  病人id_In     门诊费用记录.病人id%Type,
        strSQL = strSQL & "" & mCurBillType.lng病人ID & ","
        '  冲销id_In     病人预交记录.结帐id%Type,
        strSQL = strSQL & "" & lng冲销ID & ","
        '  结算方式_In   Varchar2,
        strSQL = strSQL & "" & "NULL" & ")"
        '  退预交_In     病人预交记录.冲预交%Type := Null,
        '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
        '  卡号_In       病人预交记录.卡号%Type := Null,
        '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
        '  交易说明_In   病人预交记录.交易说明%Type := Null,
        '  缴款_In       病人预交记录.缴款%Type := Null,
        '  找补_In       病人预交记录.找补%Type := Null,
        '  误差金额_In   门诊费用记录.实收金额%Type := Null,
        '  完成退费_In   Number := 0,
        '  原结帐id_In   病人预交记录.结帐id%Type := Null
        zlAddArray cllPro, strSQL
    End If
    
    '先退医保
    If mCurBillType.intInsure <> 0 And MCPAR.门诊结算作废 Then
        If Not bln全退 And Not mblnDelByNo Then
            '可能存在重新收费,因此,需要调用身份验证接口(Identifiy)
            'strAdvace:医保部分退时:传入1,表示医保部分退后再重新收费的身份验证;其他传入: 空
            lngCheck病人ID = mCurBillType.lng病人ID
            intCheckInsure = mCurBillType.intInsure
            strYBPati = gclsInsure.Identify(0, lngCheck病人ID, intCheckInsure, 1)
            
            If strYBPati = "" Then
                MsgBox "医保身份验证失败,不允许继续退费!", vbOKOnly + vbDefaultButton1 + vbExclamation, gstrSysName
                Exit Function
            End If
            
            If Val(CLng(Split(strYBPati, ";")(8))) <> mCurBillType.lng病人ID Then
                MsgBox "医保验证的病人与退费的病人不是同一个病人!", vbInformation, gstrSysName
                Call ExecuteYBIdentifyCancel(mCurBillType.lng病人ID, mCurBillType.intInsure)
                Exit Function
            End If
        End If
        blnTrans = True
        zlExecuteProcedureArrAy cllPro, Me.Caption, True
        
        If Not bln全退 And Not mblnDelByNo Then
            '更新重收记录的保险信息
            '77951,冉俊明,2014-9-16，
            If ExecuteInsureInfoUpdate(lng结帐ID, str保险金额) = False Then Exit Function
            '读取个帐余额
            cur个帐透支 = mdbl个帐透支
            mdbl个帐余额 = gclsInsure.SelfBalance(mCurBillType.lng病人ID, CStr(Split(strYBPati, ";")(1)), 10, cur个帐透支, mCurBillType.intInsure)
            mdbl个帐透支 = cur个帐透支
        End If
        If ExecuteClinicDelSwap(mCurBillType.lng病人ID, mCurBillType.intInsure, lng冲销ID, lng原结帐ID) = False Then Exit Function
        Set cllPro = New Collection
        
        '重新进行收费处理
        '77058
        If Not bln全退 And Not mblnDelByNo Then
            gcnOracle.BeginTrans
            If ExcuteInsureReCharge(mCurBillType.lng病人ID, mCurBillType.intInsure, lng结帐ID, lng结算序号, str保险金额, _
                        strNo, lng领用ID, strInvoice, dtDelDate) = False Then Exit Function
        End If
        blnTrans = False
    End If
    
    '2.再退一卡通(老版本)
    If mCurBillType.blnExistOnCard Then
ReDOOneCard:
        If CheckOnCardValied(mrsBalance) = False Then Exit Function
        blnTrans = True
        zlExecuteProcedureArrAy cllPro, Me.Caption, True
        If Not ExecuteOneCardDelInterface(mrsBalance, lng冲销ID) Then
            If mCurBillType.intInsure <> 0 Then
                strCmdCaptions = "异常单据(&C)|表示不进行一卡通调用,数据将以异常形式体现,但必须在今后进行处理"
                strCmdCaptions = strCmdCaptions & ";重退(&R)|表示重新调用一卡通结算交易"
                If frmVerfyCodeInput.ShowMsg(Me, "单据[" & mCurBillType.strDelNOs & "]已经退费成功,但一卡通交易失败,[异常单据]必须输入验证码,建议不进行异常单据保存", strCmdCaptions) = False Then
                     gcnOracle.BeginTrans: blnTrans = True
                    GoTo ReDOOneCard:
                End If
            End If
            Exit Function
        End If
        Set cllPro = New Collection: blnTrans = False
    End If
    
    '4.显示结算界面
    Dim frmBalance As New frmClinicDelBalance, objDelBalance As New clsCliniDelBalance
    
    Set objDelBalance.rsBalance = mrsBalance
    Set objDelBalance.rs结算方式 = mrs结算方式
    If strPrintNOInfor <> "" Then strPrintNOInfor = Mid(strPrintNOInfor, 2)
    mCurBillType.lng结算序号 = lng结算序号 '记录用于打印红票
    
    With objDelBalance
        .intInsure = mCurBillType.intInsure
        .CurDelNos = mCurBillType.strDelNOs
        .AllNos = mCurBillType.strAllNOs
        .PrintNOs = strPrintNOInfor
        
        .PatiUseType = mstrUseType
        .SaveBilled = cllPro.Count = 0
        .ShareUserID = mlngShareUseID
        .病人ID = mCurBillType.lng病人ID
        .冲销ID = lng冲销ID
        .当前发票号 = strInvoice
        .回收发票 = strInvoices
        .结算序号 = lng结算序号
        .结帐ID = lng结帐ID
        .缺省结算方式 = mCurBillType.str结算方式
        .退费合计 = -1 * GetDelMoney
        .费别 = mCurBillType.str费别
        .年龄 = mCurBillType.str年龄
        .性别 = mCurBillType.str性别
        .姓名 = mCurBillType.str姓名
        .医保不走票号 = MCPAR.医保不走票号
        .原结帐ID = mCurBillType.lng原结帐ID
        .退费时间 = dtDelDate
        .部分退费 = Not bln全退
        .原样退 = bln原样退
        .blnOnePatiPrint = mblnOnePatiPrint
        .strOnePatiPrintNos = mstrOnePatiPrintNos
    End With
    Call GetAsyncKeyState(VK_RETURN)
    If frmBalance.zlDelCharge(Me, EM_FUN_退费, mlngModule, mstrPrivs, objDelBalance, cllPro, , mcllForceDelToCash) = False Then
        Exit Function
    End If
    
    '81190,冉俊明,退费业务向发药机上传退费信息
    On Error Resume Next
    If mblnDrugMachine Then
        Dim rsTemp As ADODB.Recordset, strData As String '门诊处方退药格式：费用ID1,退药数量1;费用ID2,退药数量2;...
        '本次退的减去重收的就是实际退的
        strSQL = "Select Max(Decode(a.记录状态, 2, a.Id, 0)) As 费用id, -1 * Nvl(Sum(a.付数 * a.数次), 0) As 退药数量" & vbNewLine & _
                " From 门诊费用记录 A,(Select Distinct 结帐ID From 病人预交记录 Where 结算序号 = [1]) B" & vbNewLine & _
                " Where a.结帐id = b.结帐ID And Mod(a.记录性质, 10) = 1 And a.收费类别 In ('5', '6', '7')" & vbNewLine & _
                " Group By NO, Nvl(价格父号, 序号)" & vbNewLine & _
                " Having Nvl(Sum(a.付数 * a.数次), 0) <> 0"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询本次退费项目", objDelBalance.结算序号)
        Do While Not rsTemp.EOF
            strData = strData & ";" & Nvl(rsTemp!费用id) & "," & Nvl(rsTemp!退药数量)
            rsTemp.MoveNext
        Loop
        If strData <> "" Then
            strData = Mid(strData, 2)
            Call mobjDrugMachine.Operation(gstrDBUser, Val("24-处方退药(完整/部分)"), strData, strReturn)
        End If
    ElseIf mblnDrugPacker Then
        If strReturnRecipt <> "" Then
            strReturnRecipt = Mid(strReturnRecipt, 2)
            Call mobjDrugPacker.DYEY_MZ_TransRecipeReturn(1, UserInfo.编号, UserInfo.姓名, strReturnRecipt, strReturn)
        End If
    End If
    Err.Clear: On Error GoTo Errhand
    
    If Not gfrmMain Is Nothing Then
        Call zlExeBalanceWinRefrshData(True, objDelBalance)
    End If
    ExecDelete = True
    Exit Function
Errhand:
    If blnTrans Then
        gcnOracle.RollbackTrans
        Call ErrCenter
    Else
        If ErrCenter = 1 Then
            Resume
        End If
    End If
    If Err.Number <> 0 Then Call SaveErrLog
    
    '中断提示,不打印，重新退费后再打印或自己选择重打
    Call ShowErrBill(mCurBillType.strDelNOs, strNo)
End Function



Private Sub PrintDelBill(ByVal strAllNOs As String, ByVal strCurDelNOs As String, _
    ByVal strNo As String, _
    ByVal lng病人ID As Long, _
    ByVal dtDateDel As Date, ByVal blnAll部分退费 As Boolean, _
    ByVal strInvoices As String, ByVal strReclaimInvoice As String, _
    Optional blnOnePatiPrint As Boolean, Optional strOnePatiPrintNos As String)
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:打印相关票据
    '入参: strAllNOs-当前涉及的所有单据
    '       strCurDelNOs-当前退费的单据
    '       dtDateDel-退费日期
    '       strInvoices-选择的发票号(旧模式)
    '       strReclaimInvoice-回收的发票号
    '       blnOnePatiPrint-是否按病人打印票据
    '       strOnePatiPrintNos-按病人打印的单据号(多个用逗号分离,如:a,b,c
    '出参:
    '编制:刘兴洪
    '日期:2013-05-27 16:41:06
    '说明:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intInvoiceFormat As Integer, blnPrint As Boolean
    Dim str发票号 As String, int票据张数 As Integer
    Dim strSQL As String, strTempNO As String, i As Integer
    Dim lng打印ID As Long
    Dim strBillPrintNos As String
    Dim varNos As Variant, strNotAllDelNos As String '当前部分退单据，按实际分配票号分别打印时记录部分退单据
    Dim strPriceGrade As String

    On Error GoTo errHandle
    If InStr(strAllNOs, "'") = 0 Then
        strAllNOs = "'" & Replace(strAllNOs, ",", "','") & "'"
    End If
    strBillPrintNos = strAllNOs

    If InStr(strCurDelNOs, "'") = 0 Then
        strCurDelNOs = Mid(strCurDelNOs, 2)
        strCurDelNOs = "'" & Replace(strCurDelNOs, ",", "','") & "'"
    End If

    If blnOnePatiPrint Then
        '按病人补打票据，需要生产生临时表数据
        Dim blnAllDel As Boolean
        If zlSaveTempPrintData(strOnePatiPrintNos, mlng领用ID, "", lng打印ID) = False Then GoTo PrintList
        If zlChargeBillIsAllDel("", lng打印ID, blnAllDel, strBillPrintNos) = False Then GoTo PrintList
        
        If InStr(strBillPrintNos, "'") = 0 Then strBillPrintNos = "'" & Replace(strBillPrintNos, ",", "','") & "'"

        If blnAllDel Then
            If gTy_Module_Para.byt票据分配规则 <> 0 Then
                '全退了，则直接打印清单(回收票据已经在退费单据中处理了)
                str发票号 = strReclaimInvoice
                Call zlExeCuteBillNoSplit(False, 4, mlng领用ID, strAllNOs, lng病人ID, "", dtDateDel, 1, str发票号, int票据张数, lng打印ID)
'            Else
'                'Zl_门诊收费记录_Reprint
'                strSQL = "Zl_门诊收费记录_Reprint("
'                '  No_In         门诊费用记录.No%Type,
'                strSQL = strSQL & "'" & Split(strBillPrintNos & ",", ",")(0) & "',"
'                '  票据号_In     票据使用明细.号码%Type,
'                strSQL = strSQL & "NULL,"   '全部收回，没有发出票号
'                '  领用id_In     票据使用明细.领用id%Type,
'                strSQL = strSQL & "" & 0 & ","
'                '  使用人_In     票据使用明细.使用人%Type,
'                strSQL = strSQL & "'" & UserInfo.姓名 & "',"
'                '  使用时间_In   票据使用明细.使用时间%Type,
'                strSQL = strSQL & "to_date('" & Format(dtDateDel, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
'                '  退费_In       Number := 0,
'                strSQL = strSQL & "1,"
'                '  票据张数_In   Number := 0,
'                strSQL = strSQL & "0,"
'                '  收回票据号_In Varchar2 := Null,
'                strSQL = strSQL & "NULL,"
'                '  票种_In Number:=1
'                strSQL = strSQL & "1)"
'                zlDatabase.ExecuteProcedure strSQL, Me.Caption
            End If
            GoTo PrintList:
            Exit Sub
        End If
    End If
    
    If (Not blnAll部分退费 And blnOnePatiPrint = False) Or (blnAllDel And blnOnePatiPrint) Then
         '税控部件全退时收回处理(全退时，zl_门诊收费记录_DELETE中已收回票据)
        If Not gobjTax Is Nothing And gblnTax Then
            gstrTax = gobjTax.zlTaxOutErase(gcnOracle, strAllNOs)
            If gstrTax <> "" Then MsgBox gstrTax, vbExclamation, gstrSysName
        End If
        GoTo PrintList:
        Exit Sub
    End If
    '77058
    If blnAll部分退费 And mCurBillType.intInsure <> 0 And MCPAR.医保接口打印票据 And blnOnePatiPrint = False Then GoTo PrintList
    
    '部分退费时收回并重打,包括单张部分退和退多张中的某几张
    If gTy_Module_Para.byt票据分配规则 <> 0 And strReclaimInvoice <> "" Then
        '按新票据分配规则打印
        '先预算,看票据是否充足
        str发票号 = strReclaimInvoice
        If zlExeCuteBillNoSplit(True, 4, mlng领用ID, strAllNOs, lng病人ID, "", dtDateDel, 1, str发票号, int票据张数, , , lng打印ID) = False Then GoTo PrintList:
        If int票据张数 = 0 Then
            '只回收票据,但不打印
            str发票号 = strReclaimInvoice
            Call zlExeCuteBillNoSplit(False, 4, mlng领用ID, strAllNOs, lng病人ID, "", dtDateDel, 1, str发票号, int票据张数, , , lng打印ID)
            GoTo PrintList:
        End If
        
        '0-不打印;1-自动打印;2-提示打印
        Select Case mintInvoicePrint
        Case 0
            blnPrint = False
        Case 1
            blnPrint = True
        Case 2
            blnPrint = MsgBox("是否打印票据？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes
        End Select
        
        '重打收回发票
        If blnPrint Then
            intInvoiceFormat = IIf(strReclaimInvoice = "" And gTy_Module_Para.byt票据分配规则 <> 0 And blnOnePatiPrint = False, mintOldInvoiceFormat, mintInvoiceFormat)
            If gintPriceGradeStartType >= 2 Then
                strPriceGrade = GetPriceGradeFromNos(strAllNOs)
            Else
                strPriceGrade = gstr普通价格等级
            End If
            Call RePrintCharge(1, strBillPrintNos, Me, mlng领用ID, strReclaimInvoice, True, dtDateDel, _
                intInvoiceFormat, , , mlngShareUseID, mstrUseType, blnOnePatiPrint, strPriceGrade)
        End If
        GoTo PrintList:
        Exit Sub
    End If
    
    If strInvoices = "无可退票据" Or strInvoices = "" Then 'a.收回并重新打印门诊收据
        '0-不打印;1-自动打印;2-提示打印
        If gTy_Module_Para.bln分别打印 And blnOnePatiPrint = False Then
            If mintInvoicePrint = 0 Then
                blnPrint = False
            Else
                strNotAllDelNos = ""
                varNos = Split(Replace(strCurDelNOs, "'", ""), ",")
                For i = 0 To UBound(varNos)
                    If CheckIsAllDel(varNos(i), True) = False Then
                        strNotAllDelNos = strNotAllDelNos & ",'" & varNos(i) & "'"
                    End If
                Next
                If strNotAllDelNos <> "" Then strNotAllDelNos = Mid(strNotAllDelNos, 2)
                
                '存在部分退的单据，需要重打
                If strNotAllDelNos = "" Then
                    blnPrint = False
                Else
                    If mintInvoicePrint = 1 Then
                        blnPrint = True
                    Else
                        blnPrint = MsgBox("是否打印票据？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes
                    End If
                End If
            End If
        Else
            '0-不打印;1-自动打印;2-提示打印
            Select Case mintInvoicePrint
            Case 0
                blnPrint = False
            Case 1
                blnPrint = True
            Case 2
                blnPrint = MsgBox("是否打印票据？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes
            End Select
        End If

        If blnPrint Then
            intInvoiceFormat = IIf(strReclaimInvoice = "" And gTy_Module_Para.byt票据分配规则 <> 0 And blnOnePatiPrint = False, mintOldInvoiceFormat, mintInvoiceFormat)
            If gintPriceGradeStartType >= 2 Then
                strPriceGrade = GetPriceGradeFromNos(strAllNOs)
            Else
                strPriceGrade = gstr普通价格等级
            End If
            If gTy_Module_Para.bln分别打印 = True And blnOnePatiPrint = False Then
                Call RePrintCharge(1, strCurDelNOs, Me, mlng领用ID, strReclaimInvoice, True, dtDateDel, _
                intInvoiceFormat, , , mlngShareUseID, mstrUseType, , strPriceGrade)
            Else
                Call RePrintCharge(1, strBillPrintNos, Me, mlng领用ID, strReclaimInvoice, True, dtDateDel, _
                intInvoiceFormat, , , mlngShareUseID, mstrUseType, blnOnePatiPrint, strPriceGrade)
            End If
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
    '退费发票(红票)打印，91998
    '0-不打印;1-自动打印;2-提示打印
    If mintInvoicePrintDel = 1 Then
        Call PrintDelCharge(mCurBillType.lng结算序号, Me, mlng领用ID, True, dtDateDel, mintInvoiceFormatDel, , , mlngShareUseID, mstrUseType)
    ElseIf mintInvoicePrintDel = 2 Then
        If MsgBox("是否打印退费票据(红票)？", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then
            Call PrintDelCharge(mCurBillType.lng结算序号, Me, mlng领用ID, True, dtDateDel, mintInvoiceFormatDel, , , mlngShareUseID, mstrUseType)
        End If
    End If
    
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
    If mCurBillType.intInsure <> 0 And MCPAR.退费后打印回单 And InStr(1, mstrPrivs, "医保退费回单") > 0 Then
        '问题:35248
        'If strCurDelNOs <> "" Then strCurDelNOs = Mid(strCurDelNOs, 2) '冉俊明,2014-9-10,分隔符已在前面去掉，这里不能再处理
        Call ReportOpen(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1121_4", Me, "NO=" & strCurDelNOs, 2)
    End If
    '77058
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
    '77635:李南春,2014/9/9,申请原因长度控制
    If zlCommFun.ActualLen(txt退费摘要.Text) > 100 Then
        MsgBox "申请原因最多允许输入 " & 100 & " 个字符或 " & 50 & " 个汉字！", vbInformation, gstrSysName
        If txt退费摘要.Visible And txt退费摘要.Enabled Then txt退费摘要.SetFocus
    End If
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
        If InStr(1, mCurBillType.strNosPatiDel & ",", "," & strNo & ",") > 0 Then
            CheckBillIsAllDels = 2: Exit Function
        End If
        CheckBillIsAllDels = 1: Exit Function
     End If
    CheckBillIsAllDels = 2
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
    mstrUseType = zl_GetInvoiceUserType(mCurBillType.lng病人ID, 0, mCurBillType.intInsure)
    mlngShareUseID = zl_GetInvoiceShareID(mlngModule, mstrUseType)
    mintInvoiceFormat = zl_GetInvoicePrintFormat(mlngModule, mstrUseType, mintOldInvoiceFormat, mblnOnePatiPrint)
    mintInvoiceFormatDel = zl_GetInvoicePrintFormat(mlngModule, mstrUseType, , , True)
    mintInvoicePrint = zl_GetInvoicePrintMode(mlngModule, mstrUseType)
    mintInvoicePrintDel = zl_GetInvoicePrintMode(mlngModule, mstrUseType, True)
End Sub

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
        .Editable = flexEDKbdMouse
    End With
End Sub
Private Sub LoadBalanceInfor()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载收款结算方式
    '编制:刘兴洪
    '日期:2014-07-02 14:46:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str结算方式 As String, lngRow As Long
    Dim lngCol As Long, i As Long, intSign As Integer

    If mrsBalance Is Nothing Then Exit Sub
    If mrsBalance.State <> 1 Then Exit Sub
    intSign = IIf(mstrDelTime <> "", -1, 1) '数量,金额正负符号
    '字段:类型 ,结帐ID, 记录性质, 结算方式, 摘要, 卡类别ID, 卡类别名称, 自制卡, 结算卡序号, 结算号码, 卡号, 交易流水号, 交易说明, 结算序号, 校对标志, 医保, 消费卡id
    '            是否密文,是否全退,是否退现,冲预交
    '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
    lngRow = 0
    mrsBalance.Filter = 0
    mrsBalance.Sort = "类型,结算方式"
    With vsBalance
        .Redraw = flexRDNone
        Call ClearBalance
        If mrsBalance.RecordCount <> 0 Then mrsBalance.MoveFirst
        Do While Not mrsBalance.EOF
            '--问题:52530
            If Val(Nvl(mrsBalance!类型)) = 1 Then
                str结算方式 = "冲预存款"
            Else
                str结算方式 = Nvl(mrsBalance!结算方式, "未结金额")
            End If
            If str结算方式 <> "" Then
                '先查找是否存在相同的结算方式,存在直接汇总
                lngCol = -1
                For i = 1 To .COLS - 1 Step 2
                    If str结算方式 = .Cell(flexcpData, lngRow, i) Then
                        lngCol = i: Exit For
                    End If
                Next
                If lngCol = -1 Then
                    .COLS = .COLS + 2
                    .ColAlignment(.COLS - 2) = 7: .ColAlignment(.COLS - 1) = 1
                    lngCol = .COLS - 2
                End If
                .TextMatrix(lngRow, lngCol) = str结算方式 & ":"
                .Cell(flexcpData, lngRow, lngCol) = str结算方式
                .TextMatrix(lngRow, lngCol + 1) = zlFormatNum(Val(.TextMatrix(lngRow, .COLS - 1)) + intSign * Val(Nvl(mrsBalance!冲预交, 0)))
                
                .Cell(flexcpData, lngRow, lngCol + 1, lngRow, lngCol + 1) = Val(Nvl(mrsBalance!是否退现))
                If mbytMode = EM_MULTI_退费 Then
                    .ColData(lngCol) = "摘要:" & IIf(Val(Nvl(mrsBalance!退费)) = 1, "", Nvl(mrsBalance!摘要))
                    .ColData(lngCol + 1) = "结算号码:" & IIf(Val(Nvl(mrsBalance!退费)) = 1, "", Nvl(mrsBalance!结算号码))
                ElseIf mbytMode = EM_MULTI_异常重退 Then
                    .ColData(lngCol) = "摘要:" & IIf(Val(Nvl(mrsBalance!退费)) = 1, "", Nvl(mrsBalance!摘要))
                    .ColData(lngCol + 1) = "结算号码:" & IIf(Val(Nvl(mrsBalance!退费)) = 1, "", Nvl(mrsBalance!结算号码))
                Else
                    If mstrDelTime <> "" Then
                        .ColData(lngCol) = "摘要:" & IIf(Val(Nvl(mrsBalance!退费)) = 1, "", Nvl(mrsBalance!摘要))
                        .ColData(lngCol + 1) = "结算号码:" & IIf(Val(Nvl(mrsBalance!退费)) = 1, "", Nvl(mrsBalance!结算号码))
                    Else
                        .ColData(lngCol) = "摘要:" & IIf(Val(Nvl(mrsBalance!退费)) = 1, Nvl(mrsBalance!摘要), "")
                        .ColData(lngCol + 1) = "结算号码:" & IIf(Val(Nvl(mrsBalance!退费)) = 1, Nvl(mrsBalance!结算号码), "")
                    End If
                End If
                If Val(Nvl(mrsBalance!结算性质)) <> 1 Then
                   .Cell(flexcpForeColor, lngRow, .COLS - 1, lngRow, .COLS - 2) = IIf(mrsBalance!结算性质 = 9, vbRed, vbBlue)
                   .Cell(flexcpForeColor, 1, .COLS - 1, 1, .COLS - 2) = vbRed
                   .Cell(flexcpFontBold, 1, .COLS - 1, 1, .COLS - 2) = True    '粗体
                
                End If
            End If
             mrsBalance.MoveNext
            .Redraw = flexRDBuffered
         Loop
         '77210,冉俊明,2014-8-27,部分退费后再退费,不显示金额为零的结算方式信息
         i = 1
         Do While i < .COLS - 1
            If .TextMatrix(lngRow, i + 1) = "0" Then
                For lngCol = i To .COLS - 3
                    .TextMatrix(lngRow, lngCol) = .TextMatrix(lngRow, lngCol + 2)
                    .Cell(flexcpData, lngRow, lngCol) = .Cell(flexcpData, lngRow, lngCol + 2)
                    .ColData(lngCol) = .ColData(lngCol + 2)
                    .Cell(flexcpForeColor, lngRow, lngCol) = .Cell(flexcpForeColor, lngRow, lngCol + 2)
                    .Cell(flexcpForeColor, 1, lngCol) = .Cell(flexcpForeColor, 1, lngCol + 2)
                    .Cell(flexcpFontBold, 1, lngCol) = .Cell(flexcpFontBold, 1, lngCol + 2)
                Next
                .COLS = .COLS - 2
            Else
                i = i + 2
            End If
         Loop
         vsBalance.AutoSizeMode = flexAutoSizeColWidth
         Call vsBalance.AutoSize(0, .COLS - 1)
         If mbytMode = EM_MULTI_查看 Or mbytMode = EM_MULTI_退费申请 Or mbytMode = EM_MULTI_取消申请 _
            Or mbytMode = EM_MULTI_退费审核 Or mbytMode = EM_MULTI_取消审核 Then
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
'        txt退款金额.Tag = dbl选择合计
    End With
End Sub

Private Sub ReCalcDelMoney()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:重新计算退款合计
    '编制:刘兴洪
    '日期:2014-07-03 17:24:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnAll As Boolean, blnShowDel As Boolean
    Dim i As Long, strTemp As String
    Dim blnSeled As Boolean
    Dim bln全退 As Boolean
    Dim strSelNos As String, varSelNos As Variant
    Dim strFilter As String, dblMoneyNo As Double, strBalances As String
    
    txt退款合计 = Format(GetDelMoney, gstrDec)
    bln全退 = IsFeeAllDel
    
    blnShowDel = bln全退 Or mCurBillType.intInsure <> 0 Or mCurBillType.blnExistThreeAllDel
    blnShowDel = blnShowDel And Not (mbytMode = EM_MULTI_退费申请 Or mbytMode = EM_MULTI_取消申请 _
                Or mbytMode = EM_MULTI_退费审核 Or mbytMode = EM_MULTI_取消审核) '退费申请模式不显示退费行
    
    blnSeled = False
    With vsBill
        For i = 1 To .Rows - 1
            If .TextMatrix(i, .ColIndex("单据号")) <> "" And Abs(Val(.TextMatrix(i, .ColIndex("选择")))) = 1 Then
                blnSeled = True
                If InStr(strSelNos, "," & .TextMatrix(i, .ColIndex("单据号"))) = 0 Then
                    strSelNos = strSelNos & "," & .TextMatrix(i, .ColIndex("单据号"))
                End If
            End If
        Next
        If strSelNos <> "" Then strSelNos = Mid(strSelNos, 2)
    End With
    blnShowDel = blnShowDel And blnSeled
    
    vsBalance.RowHidden(1) = Not blnShowDel
    If vsBalance.RowHidden(1) Then
        If vsBalance.COLS > 1 Then
            vsBalance.Cell(flexcpData, 1, 1, 1, vsBalance.COLS - 1) = ""
            vsBalance.Cell(flexcpText, 1, 1, 1, vsBalance.COLS - 1) = ""
        End If
        Call ControlResize
        Exit Sub
    End If

    With vsBalance
        If bln全退 And MCPAR.门诊结算作废 Then
            For i = 1 To .COLS - 1
                .TextMatrix(1, i) = .TextMatrix(0, i)
                .Cell(flexcpData, 1, i) = .Cell(flexcpData, 0, i)
            Next
            Call ControlResize
            Exit Sub
        End If
        '加载医保的
        '字段:类型 ,结帐ID, 记录性质, 结算方式, 摘要, 卡类别ID, 卡类别名称, 自制卡, 结算卡序号, 结算号码, 卡号, 交易流水号, 交易说明, 结算序号, 校对标志, 医保, 消费卡id
         '            是否密文,是否全退,是否退现,冲预交
         '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
         mrsBalance.Filter = 0
         mrsBalance.Sort = "类型,结算方式"
         If vsBalance.COLS > 1 Then
            .Cell(flexcpText, 1, 1, 1, .COLS - 1) = ""
            .Cell(flexcpData, 1, 1, 1, .COLS - 1) = ""
         End If
         With vsBalance
             .Redraw = flexRDNone
             
             If mrsBalance.RecordCount <> 0 Then mrsBalance.MoveFirst
             Do While Not mrsBalance.EOF
                strTemp = ""
                Select Case Val(Nvl(mrsBalance!类型))
                Case 2 '医保
                    If MCPAR.门诊结算作废 Then
                        strTemp = mrsBalance!结算方式
                        
                        If mblnDelByNo Then
                            '每一种结算只加载一次
                            If InStr(strBalances, "," & strTemp) = 0 Then
                                '按选择单据计算医保结算方式退费金额
                                dblMoneyNo = 0: strFilter = ""
                                varSelNos = Split(strSelNos, ",")
                                For i = 0 To UBound(varSelNos)
                                    If UBound(varSelNos) = 0 Then '只有一张单据
                                        strFilter = " or No='" & varSelNos(i) & "' and 结算方式='" & strTemp & "'"
                                    Else    '多张单据
                                        strFilter = strFilter & " or (No='" & varSelNos(i) & "' and 结算方式='" & strTemp & "')"
                                    End If
                                Next
                                If strFilter <> "" Then strFilter = Mid(strFilter, 4)
                                mrsInsureBalance.Filter = strFilter
                                Do While Not mrsInsureBalance.EOF
                                    dblMoneyNo = dblMoneyNo + Val(Nvl(mrsInsureBalance!金额))
                                    mrsInsureBalance.MoveNext
                                Loop
                                For i = 1 To .COLS - 1 Step 2
                                    If .Cell(flexcpData, 0, i) = strTemp Then
                                        .TextMatrix(1, i) = .TextMatrix(0, i)
                                        .Cell(flexcpData, 1, i) = strTemp
                                        .TextMatrix(1, i + 1) = zlFormatNum(Val(.TextMatrix(1, i + 1)) + dblMoneyNo)
                                        Exit For
                                    End If
                                Next
                                strBalances = strBalances & "," & strTemp
                            End If
                            strTemp = "" '清空，标记下面不再加载
                        End If
                    End If
                Case 4 '一卡通(老)
                   strTemp = mrsBalance!结算方式
                End Select
                If strTemp <> "" Then
                    For i = 1 To .COLS - 1 Step 2
                        If .Cell(flexcpData, 0, i) = strTemp Then
                            .TextMatrix(1, i) = .TextMatrix(0, i)
                            .Cell(flexcpData, 1, i) = strTemp
                            .TextMatrix(1, i + 1) = zlFormatNum(Val(.TextMatrix(1, i + 1)) + Val(Nvl(mrsBalance!冲预交)))
                            Exit For
                        End If
                    Next
                End If
                mrsBalance.MoveNext
            Loop
            If vsBalance.COLS > 1 Then .AutoSize 1, .COLS - 1
            
            .Redraw = flexRDBuffered
         End With
        
    End With
    Call ControlResize
End Sub
Private Function GetDelMoney() As Double
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取退款合计
    '返回:获取退款合计
    '编制:刘兴洪
    '日期:2014-07-03 17:24:06
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim dbl退款合计 As Double, i As Long
    With vsBill
        For i = 1 To .Rows - 1
            If Val(vsBill.TextMatrix(i, .ColIndex("选择"))) <> 0 Or mbytMode = EM_MULTI_异常重退 Or mbytMode = EM_MULTI_查看 Then
                dbl退款合计 = dbl退款合计 + Val(vsBill.TextMatrix(i, .ColIndex("实收金额")))
            End If
        Next
    End With

    GetDelMoney = RoundEx(dbl退款合计, 6)
End Function

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
    If Me.ActiveControl Is txtPatient And txtPatient.Visible Then
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(txtPatient.Text = "")
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(txtPatient.Text = "")
        IDKind.SetAutoReadCard (txtPatient.Text = "" And txtPatient.Visible)
    Else
        If Not mobjIDCard Is Nothing Then Call mobjIDCard.SetEnabled(False)
        If Not mobjICCard Is Nothing Then Call mobjICCard.SetEnabled(False)
        IDKind.SetAutoReadCard (False)
    End If
End Sub
Private Sub txtPatient_GotFocus()
    '问题:50885
    If txtPatient.Locked Or Not txtPatient.Visible Then Exit Sub
    If Not mobjIDCard Is Nothing Then mobjIDCard.SetEnabled (txtPatient.Text = "")
    If Not mobjICCard Is Nothing Then mobjICCard.SetEnabled (txtPatient.Text = "")
    IDKind.SetAutoReadCard (txtPatient.Text = "" And txtPatient.Visible)
    zlControl.TxtSelAll txtPatient
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
        Call ClearFace
        Exit Sub
    End If
    mCurBillType.lng病人ID = Val("" & mrsInfo!病人ID)
    txtPatient = Nvl(mrsInfo!姓名)

    lblPati.Caption = "病人:" & "                 " & _
        "　性别:" & Nvl(mrsInfo!性别) & _
        "　年龄:" & Nvl(mrsInfo!年龄) & _
        "　门诊号:" & Nvl(mrsInfo!门诊号) & _
        "　费别:" & Nvl(mrsInfo!费别) & _
        "　付款方式:" & mrsInfo!医疗付款方式
        
    With mCurBillType
        .str性别 = Nvl(mrsInfo!性别)
        .str年龄 = Nvl(mrsInfo!年龄)
        .str姓名 = Nvl(mrsInfo!姓名)
        .str费别 = Nvl(mrsInfo!费别)
    End With
    If SelectNO(mCurBillType.lng病人ID) = False Then
        If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
        Call ClearFace
        Exit Sub
    End If
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
    txtPatient.ForeColor = &HC00000: lblPati.ForeColor = txtPatient.ForeColor
    Set mrsInfo = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Mid(strInput, 2), strInput)
    If mrsInfo.EOF Then
        Set mrsInfo = New ADODB.Recordset: Exit Function
    End If
    Call SetPatiColor(txtPatient, Nvl(mrsInfo!病人类型), IIf(IsNull(mrsInfo!险类), &HC00000, vbRed))
    lblPati.ForeColor = txtPatient.ForeColor
    GetPatient = True
    Exit Function
errH:
     If ErrCenter() = 1 Then Resume
    Call SaveErrLog
NotFoundPati:
    Set mrsInfo = New ADODB.Recordset
    txtPatient.ForeColor = &HC00000
    lblPati.ForeColor = txtPatient.ForeColor
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
    '77198,冉俊明,2014-8-27,在病人退费时病人输入后,弹出的收费单选择只能提取有效挂号天数的收费单进行退费
    strSQL = "" & _
        "  With 收费单 as ( " & _
        "           Select Max(a.ID) as ID,max(M.结算序号) as 结算ID ,max(A.结帐ID) as 结帐ID,a.No as 单据号,  B.名称 as 开单部门, a.开单人, a.操作员编号, a.操作员姓名, a.实际票号, To_Char(a.发生时间, 'yyyy-mm-dd hh24:mi:ss') As 发生时间, " & vbCrLf & _
        "                   To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') As 登记时间 " & vbCrLf & _
        "           From 门诊费用记录 A,部门表 B,病人预交记录 M" & vbCrLf & _
        "           Where a.记录性质 = 1 And nvl(A.附加标志,0)<>9 and A.开单部门ID=B.ID(+) And a.记录状态 in (1,3) " & vbCrLf & _
        "                and A.结帐ID=M.结帐ID And Nvl(a.执行状态, 0) <> 1 And Nvl(a.费用状态, 0) <> 1 And a.病人id = [1] " & vbCrLf & _
        "                And a.登记时间 Between Sysdate - " & gTy_System_Para.Sy_Reg.bytNODaysGeneral & " And Sysdate " & vbCrLf & _
        "          Group by   a.No,  a.开单人, B.名称,a.操作员编号, a.操作员姓名, a.实际票号, To_Char(a.发生时间, 'yyyy-mm-dd hh24:mi:ss'),To_Char(a.登记时间, 'yyyy-mm-dd hh24:mi:ss') " & vbCrLf & _
        "           )"

     strSQL = strSQL & vbCrLf & _
     "  Select J.*  " & vbCrLf & _
     "  From 收费单 J," & vbCrLf & _
     "           (Select A.NO,sum(nvl(A.付数,1)*nvl(A.数次,1)) 数次" & vbCrLf & _
     "             From 门诊费用记录 A,收费单 B  " & vbCrLf & _
     "             Where A.NO=B.单据号 And mod(A.记录性质,10)=1 And a.价格父号 is null  " & vbCrLf & _
     "             Group by A.NO " & vbCrLf & _
     "              Having sum(nvl(A.付数,1)*nvl(A.数次,1))>0 ) M" & vbCrLf & _
     "  Where J.单据号=M.NO " & vbCrLf

     strSQL = "Select * From (" & strSQL & ") Order by 登记时间 desc,单据号"
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "退费单据", 1, "", "请选择需要退费的单据", False, False, False, 0, 0, 0, blnCancel, False, False, lng病人ID, "bytSize=1")
    If blnCancel = True Then Exit Function
    If rsTemp Is Nothing Then Exit Function
    If rsTemp.State <> 1 Then Exit Function

    Dim strNo As String
    strNo = Nvl(rsTemp!单据号)
    mblnOneCard = GetOneCard.RecordCount > 0
    mstrNo = strNo
    
    If Val(Nvl(rsTemp!结算ID)) >= 0 Then
        'bytMode-0-多张单据查看,1-多张单据退费,2-退异常的退费单进行重新退费
        frmMultiBills.ShowMe Me, 1, mstrPrivs, strNo, "", False, mlng领用ID, mblnOneCard, False, True
        Call ClearFace: Exit Function
    End If
    
    If Not ReadBills(mstrNo, True) Then
        Call ClearFace: Exit Function
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
        If .Rows = 1 And .Cell(flexcpLeft, 0, .COLS - 1) + .ColWidth(.COLS - 1) <= .Width Then
            .Height = .RowHeight(0) + 90
            Exit Sub
        End If
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

Private Sub LoadInvoiceData(ByVal strNos As String, Optional ByVal strInvoiceNO As String)

   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载发票信息
    '入参:strNos-单据号,多个用逗号分隔
    '       strInvoiceNo-发票号(按指定的发票号发票号查找)
    '编制:刘兴洪
    '日期:2013-05-07 17:07:38
    '问题:25187
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str序号 As String, varTemp As Variant, strSQL As String
    Dim i As Long, str发票号 As String
    If mbytMode <> EM_MULTI_退费 Then Exit Sub
    
    If gTy_Module_Para.byt票据分配规则 = 0 And gTy_Module_Para.bln分别打印 = True Then
        strSQL = "Select b.No, a.号码 As 票号, Null As 序号, Null As 关联票号序号" & vbNewLine & _
                " From 票据使用明细 A, 票据打印内容 B" & vbNewLine & _
                " Where b.数据性质 = 1 And b.No In (Select Column_Value From Table(f_Str2list([1])))" & vbNewLine & _
                "       And b.Id = a.打印id And a.票种 = 1 And a.性质 = 1 And a.原因<>6" & vbNewLine & _
                "       And Not Exists (Select 1 From 票据使用明细 Where 打印id = b.Id And 性质 = 2)"

        Set mrsDelInvoice = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNos)
        GoTo LoadIntoVS
    End If
    If gTy_Module_Para.byt票据分配规则 = 0 Then Exit Sub
    If mrsDelInvoice Is Nothing Then
        Set mrsDelInvoice = zlGetFromNoTOInvoice(strNos)
    End If
LoadIntoVS:
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

Private Function GetInvoiceNo(ByVal strNo As String) As String
    Dim str发票号 As String
    On Error GoTo errHandle
    If mrsDelInvoice Is Nothing Then Exit Function
    If mrsDelInvoice.State <> 1 Then Exit Function
    If mrsDelInvoice.RecordCount = 0 Then Exit Function
    With mrsDelInvoice
        .Filter = "NO='" & strNo & "'"
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            str发票号 = str发票号 & "," & Nvl(!票号)
            .MoveNext
        Loop
        .Filter = 0: .MoveFirst
    End With
    '进行排序处理
    If str发票号 = "" Then Exit Function
    str发票号 = Mid(str发票号, 2)
    GetInvoiceNo = zlStringSort(str发票号)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub FromNoSelectInvoice()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据单据选择发票
    '编制:刘兴洪
    '日期:2013-05-08 15:52:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str发票号 As String, strNo As String
    Dim strNos As String, i As Long, j As Long
    If mbytMode <> 1 Or (gTy_Module_Para.byt票据分配规则 = 0 And gTy_Module_Para.bln分别打印 = False) Then Exit Sub

    On Error GoTo errHandle
    If gTy_Module_Para.byt票据分配规则 = 0 And gTy_Module_Para.bln分别打印 = True Then
        With vsBill
            str发票号 = ""
            For i = 1 To .Rows - 1
                  If Abs(Val(.TextMatrix(i, .ColIndex("选择")))) = 1 Then
                        strNo = .TextMatrix(i, .ColIndex("单据号"))
                        If strNo <> "" Then
                            str发票号 = str发票号 & "," & GetInvoiceNo(strNo)
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
    Else
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
    End If
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
    If mbytMode <> 1 Or (gTy_Module_Para.byt票据分配规则 = 0 And gTy_Module_Para.bln分别打印 = False) Then Exit Sub
    If mrsDelInvoice Is Nothing Then Exit Sub
    If mrsDelInvoice.RecordCount = 0 Then Exit Sub
    On Error GoTo errHandle
    mrsDelInvoice.Filter = "票号='" & strInvoiceNO & "'"
    If gTy_Module_Para.byt票据分配规则 = 0 And gTy_Module_Para.bln分别打印 = True Then
        With mrsDelInvoice
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                strNo = Nvl(!NO)
                  With vsBill
                      k = vsBill.FindRow(strNo, , .ColIndex("单据号"))
                      For j = k To .Rows - 1
                          If .TextMatrix(j, .ColIndex("单据号")) <> strNo Then Exit For
                          .Cell(flexcpChecked, j, .ColIndex("选择")) = 1
                          '同步选择相关组合项目
                          Call SynchronizationSelect(j)
                      Next
                  End With
                 .MoveNext
            Loop
        End With
    Else
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
    End If
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
    If mbytMode <> 1 Or (gTy_Module_Para.byt票据分配规则 = 0 And gTy_Module_Para.bln分别打印 = False) Then Exit Sub

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
Private Function ExistsBalance(ByVal str结算方式 As String, ByRef intCol As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:存在指定的结算方式
    '入参:
    '出参:intCol-指定结算方式列(-1表示未找到)
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-20 13:40:40
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer

    On Error GoTo errHandle
    intCol = -1
    With vsBalance
        For i = 1 To .COLS - 1 Step 2
            If .Cell(flexcpData, 1, i) = str结算方式 Then
                intCol = i
                ExistsBalance = True: Exit Function
            End If
        Next
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Public Function CheckDiff(strNos As String, strDiffNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:比较个值的单据号是否一致
    '入参:
    '出参:
    '返回:全部一致,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-03-21 17:19:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varTemp As Variant, varData As Variant, i As Long

    On Error GoTo errHandle
    varTemp = Split(Replace(strDiffNos, "'", ""), ",")
    varData = Split(Replace(strNos, "'", ""), ",")
    If UBound(varTemp) <> UBound(varData) Then Exit Function
    For i = 0 To UBound(varData)
        If InStr(1, "," & strDiffNos & ",", "," & varData(i) & ",") = 0 Then Exit Function
    Next
    For i = 0 To UBound(varTemp)
        If InStr(1, "," & strNos & ",", "," & varTemp(i) & ",") = 0 Then Exit Function
    Next
    CheckDiff = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub initInsurePara(ByVal intInsure As Integer, ByVal lng病人ID As Long, ByVal lng结帐ID As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化医保参数
    '编制:刘兴洪
    '日期:2014-06-26 16:25:17
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If intInsure = 0 Then Exit Sub
    
    MCPAR.门诊结算作废 = gclsInsure.GetCapability(support门诊结算作废, lng病人ID, intInsure)
    MCPAR.医保接口打印票据 = gclsInsure.GetCapability(support医保接口打印票据, lng病人ID, intInsure, CStr(lng结帐ID))
    MCPAR.退费后打印回单 = gclsInsure.GetCapability(support退费后打印回单, lng病人ID, intInsure)
    MCPAR.门诊预结算 = gclsInsure.GetCapability(support门诊预算, lng病人ID, intInsure)
    MCPAR.先自付 = gclsInsure.GetCapability(support收费帐户首先自付, lng病人ID, intInsure)
    MCPAR.全自付 = gclsInsure.GetCapability(support收费帐户全自费, lng病人ID, intInsure)
    MCPAR.医保不走票号 = False
    MCPAR.按单据全退 = gclsInsure.GetCapability(support按单据全退, lng病人ID, intInsure) '86176
    MCPAR.多单据分单据结算 = gclsInsure.GetCapability(support多单据分单据结算, lng病人ID, intInsure)
    MCPAR.一次结算分单据退费 = gclsInsure.GetCapability(support一次结算分单据退费, lng病人ID, intInsure)
End Sub

Private Sub SetFunCtrlVisible()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置功能控件的visible属性
    '编制:刘兴洪
    '日期:2014-07-03 16:41:02
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnVisible As Boolean
    
    blnVisible = mbytMode = EM_MULTI_退费 Or mbytMode = EM_MULTI_退费申请 Or mbytMode = EM_MULTI_取消申请 _
                Or mbytMode = EM_MULTI_退费审核 Or mbytMode = EM_MULTI_取消审核
    cmdSelAll.Visible = blnVisible
    cmdClear.Visible = blnVisible
    cmdBillSel.Visible = mbytMode = EM_MULTI_退费
    cmdRefuseApply.Visible = mbytMode = EM_MULTI_退费审核
    cmdOK.Visible = Not mbytMode = EM_MULTI_查看
    If mlng结算序号 <> 0 Then   '外面传入时,不用手工输入
        txtNO.Visible = False
        optNO(0).Visible = False
        optNO(1).Visible = False
        picPatiBack.Visible = False
        fraInfo_1.Visible = False
    End If
End Sub
Private Function GetYBTOCash(ByVal lng病人ID As Long, ByVal intInsure As Integer) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医保使用现金支付的结算方式(多个用逗号分隔)
    '返回:返回结算方式,多个用逗号分隔:个人帐户,医保基金...
    '编制:刘兴洪
    '日期:2014-07-07 09:57:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str结算方式 As String
    
    On Error GoTo errHandle
    If intInsure = 0 Then Exit Function
    
    '字段:类型,记录性质,结算方式,摘要,卡类别ID,卡类别名称,自制卡,结算卡序号,结算号码,卡号,交易流水号, 交易说明,结算序号,校对标志,医保,消费卡id
    '     是否密文,是否全退,是否退现,冲预交
    '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡)
    mrsBalance.Filter = "类型=2"
    If mrsBalance.RecordCount = 0 Then Exit Function
    With mrsBalance
        Do While Not .EOF
            '如果这种结算方式不支持回退,要退为现金,则不用减去
            If MCPAR.门诊结算作废 Then
                If Not gclsInsure.GetCapability(support门诊结算作废, lng病人ID, intInsure, !结算方式) Then
                    str结算方式 = str结算方式 & "," & !结算方式
                End If
            Else     '不支持门诊结算作废时,只允许个帐退为现金,其它原样退,不调用医保交易
                If !结算方式 = mstr个人帐户 Then
                    str结算方式 = str结算方式 & "," & !结算方式
                End If
            End If
            .MoveNext
        Loop
    End With
    If str结算方式 <> "" Then str结算方式 = Mid(str结算方式, 2)
    GetYBTOCash = str结算方式
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetYBOldBalance(ByVal lng病人ID As Long, ByVal intInsure As Integer, ByVal lng原结帐ID As Long) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取医保原结算方式和结算金额
    '返回:返回结算信息,格式:结算方式|结算金额||...
    '编制:刘兴洪
    '日期:2014-07-07 09:57:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str结算方式 As String
    
    On Error GoTo errHandle
    If intInsure = 0 Then Exit Function
    
    '字段:类型,记录性质,结算方式,摘要,卡类别ID,卡类别名称,自制卡,结算卡序号,结算号码,卡号,交易流水号, 交易说明,结算序号,校对标志,医保,消费卡id
    '     是否密文,是否全退,是否退现,冲预交
    '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡)
    mrsBalance.Filter = "类型=2 and 结帐ID=" & lng原结帐ID
    If mrsBalance.RecordCount = 0 Then Exit Function
    With mrsBalance
        Do While Not .EOF
            '如果这种结算方式不支持回退,要退为现金,则不用减去
            If MCPAR.门诊结算作废 Then
                If gclsInsure.GetCapability(support门诊结算作废, lng病人ID, intInsure, !结算方式) Then
                    str结算方式 = str结算方式 & "||" & !结算方式 & "|" & Val(Nvl(!冲预交))
                End If
            Else     '不支持门诊结算作废时,只允许个帐退为现金,其它原样退,不调用医保交易
                If !结算方式 <> mstr个人帐户 Then
                    str结算方式 = str结算方式 & "||" & !结算方式 & "|" & Val(Nvl(!冲预交))
                End If
            End If
            .MoveNext
        Loop
    End With
    If str结算方式 <> "" Then str结算方式 = Mid(str结算方式, 3)
    GetYBOldBalance = str结算方式
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ExcuteInsureReCharge(ByVal lng病人ID As Long, ByVal intInsure As Integer, _
    ByVal lng结帐ID As Long, ByVal lng结算序号 As Long, ByVal str保险金额 As String, _
    ByVal strNo As String, ByVal lng领用ID As Long, ByVal strInvoice As String, ByVal dtDelDate As Date) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行医保重新收费
    '返回:执行成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-31 23:38:11
    '说明:参数strNO,lng领用ID,strInvoice,dtDelDate用于医保接口打印票据
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strAdvance As String, arrBalance As Variant, str结算方式 As String
    Dim dbl结算金额 As Double, dbl可分配额 As Double, dbl余额 As Double
    Dim strBalance As String, dbl退款合计 As Double, str退回结算 As String
    Dim strSQL As String, rsTmp As ADODB.Recordset, strYbInvoice As String
    Dim i As Long, k As Long, j As Long, cur误差金额 As Double
    Dim strNone As String, strNos As String, varTemp As Variant, cur个帐 As Currency
    
    On Error GoTo errHandle
    If mCurBillType.intInsure = 0 Then
        ExcuteInsureReCharge = False
        gcnOracle.RollbackTrans: Exit Function
    End If
    
    strBalance = ""
    If Not MCPAR.门诊预结算 Then '计算个人帐户支付金额
        varTemp = Split(str保险金额, ";") 'cur实收合计;cur进入统筹;cur全自付;cur先自付
        If intInsure <> 0 And mstr个人帐户 <> "" And mdbl个帐余额 > -1 * mdbl个帐透支 Then
            If RoundEx(Val(varTemp(0)), 6) >= 0 Then
                cur个帐 = RoundEx(Val(varTemp(1)), 6) + IIf(MCPAR.先自付, RoundEx(Val(varTemp(3)), 6), 0) + IIf(MCPAR.全自付, RoundEx(Val(varTemp(2)), 6), 0)
                If mdbl个帐余额 - cur个帐 >= -1 * mdbl个帐透支 Then
                    strBalance = mstr个人帐户 & "|" & cur个帐   '在允许透支范围内足够(允许透支0为特例)
                Else
                    If mdbl个帐透支 = 0 And mdbl个帐余额 > 0 Then
                        strBalance = mstr个人帐户 & "|" & mdbl个帐余额  '不允许透支且有余额
                    Else
                        '超过允许透支范围或不允许透支时无余额
                        If mdbl个帐透支 <> 0 Then
                            strBalance = mstr个人帐户 & "|" & mdbl个帐余额 + mdbl个帐透支 '在允许透支范围内支付
                        Else
                            strBalance = mstr个人帐户 & "|0"
                        End If
                    End If
                End If
            Else
                strBalance = mstr个人帐户 & "|0"
            End If
        End If
    Else
        If ExecuteClinicPreSwap(intInsure, lng结帐ID, lng病人ID, strBalance, strNone, strYbInvoice, strNos) = False Then
            gcnOracle.RollbackTrans
            If strNone <> "" Then
                MsgBox "当前保险结算使用的结算方式" & vbCrLf & vbCrLf & vbTab & strNone & vbCrLf & vbCrLf & _
                    "在门诊未设置，请先到结算方式管理中设置这些结算方式！", vbInformation, gstrSysName
            End If
            Exit Function
        End If
    End If
    ' Zl_门诊收费结算_Modify
    strSQL = "Zl_门诊收费结算_Modify("
    '  操作类型_In   Number,
    '  --操作类型_In:
    '  --   0-普通收费方式:
    '  --     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
    '  --     ②冲预交_In:如果涉及预交款,则传入本次的冲预交,非正常收费时,传入零
    '  --     ③退支票额_In:如果涉及退支票,则传入本次的退支票额,非正常收费时,传入零
    '  --   1.三方卡结算:
    '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
    '  --     ②冲预交_In: 传入零
    '  --     ③退支票额_In:传入零
    '  --     ④卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
    '  --   2-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
    '  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
    '  --     ②冲预交_In: 传入零
    '  --     ③退支票额_In:传入零
    '  --   3-消费卡结算:
    '  --     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."
    '  --     ②冲预交_In: 传入零
    '  --     ③退支票额_In:传入零
    strSQL = strSQL & "" & 2 & ","
    '  病人id_In     门诊费用记录.病人id%Type,
    strSQL = strSQL & "" & lng病人ID & ","
    '  结帐id_In     病人预交记录.结帐id%Type,
    strSQL = strSQL & "" & lng结帐ID & ","
    '  结算方式_In   Varchar2,
    strSQL = strSQL & "'" & strBalance & "')"
    '  冲预交_In     病人预交记录.冲预交%Type := Null,
    '  退支票额_In   病人预交记录.冲预交%Type := Null,
    '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
    '  卡号_In       病人预交记录.卡号%Type := Null,
    '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
    '  交易说明_In   病人预交记录.交易说明%Type := Null,
    '  缴款_In       病人预交记录.缴款%Type := Null,
    '  找补_In       病人预交记录.找补%Type := Null,
    '  误差金额_In   门诊费用记录.实收金额%Type := Null,
    '  完成结算_In Number:=0
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    If MCPAR.医保接口打印票据 And MCPAR.医保不走票号 = False Then
        '38821,77058
        '票据数据生成(因为不调HIS的打印，医保接口打印，所以先填票据数据)
        strSQL = "zl_门诊收费记录_RePrint('" & strNo & "','" & strInvoice & "'," & ZVal(lng领用ID) & ",'" & UserInfo.姓名 & "'," & _
                  "To_Date('" & Format(dtDelDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),1,1)"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    
    '调用医保结算接口
    If ExecuteClinicSwap(lng病人ID, intInsure, lng结帐ID, lng结算序号, strBalance, strNos, str保险金额) = False Then Exit Function
    ExcuteInsureReCharge = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    Call ErrCenter
    Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicDelSwap, False, mCurBillType.intInsure)
End Function
Private Function ExecuteClinicPreSwap(ByVal intInsure As Integer, _
    ByVal lng结帐ID As Long, ByVal lng病人ID As Long, ByRef strBalance As String, _
    ByRef strNone As String, ByRef strYbInvoice As String, ByRef strNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行门诊预结算
    '入参:intInsure-险类
    '     lng结帐ID-重新收费的结帐ID
    '出参:strNone-不存在的结算方式
    '     strBalance-返回结算方式(结算方式|金额||...)
    '     strYbInvoice-医保返回的发票号
    '     strNOs-返回本次结算的NOs
    '返回:预结算成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-07 11:18:54
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strInvoice As String, varData As Variant
    Dim rsTemp As ADODB.Recordset, strAdvance As String
    Dim i As Long, str结算方式 As String
    Dim varTemp As Variant
    
    
    On Error GoTo errHandle
    
    strInvoice = mCurBillType.strInvoice
    Set rsTemp = zlMakeClinicPreSwapData(strInvoice, lng结帐ID, strNos)
RePreSwap:
    strAdvance = "1": strBalance = ""
    If Not gclsInsure.ClinicPreSwap(rsTemp, strBalance, intInsure, strAdvance) Then
        Screen.MousePointer = 0
        If MsgBox("重新进行医保收费时,单据预结算失败,是否重新进行预结算?", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbYes Then GoTo RePreSwap:
        Exit Function
    End If
    If strAdvance <> "" And InStr(1, strAdvance, "|") = 0 Then '医保票据号
        strYbInvoice = strAdvance
    End If
    
    MCPAR.医保不走票号 = False
    If InStr(1, strAdvance, ";") > 0 Then
        varData = Split(strAdvance & ";", ";")
        strYbInvoice = Trim(varData(0))
        '38821:strAdvance:发票号;是否不走票据号
        MCPAR.医保不走票号 = Val(varData(1)) = 1
    End If
    '结算方式;原始(最大)金额;可否修改;改后金额
    varData = Split(strBalance, "|")
    
    '结算方式|结算金额||..
    strBalance = "": strNone = ""
    For i = 0 To UBound(varData)
        varTemp = Split(varData(i), ";")
        str结算方式 = varTemp(0)
        mrs结算方式.Filter = "名称='" & str结算方式 & "' And  性质>=3 and 性质<= 4"
        If mrs结算方式.EOF Then
            strNone = strNone & "," & str结算方式
        End If
        strBalance = strBalance & "||" & varTemp(0) & "|" & Val(varTemp(1))
    Next
    If strBalance <> "" Then strBalance = Mid(strBalance, 3)
    If strNone <> "" Then
        strNone = Mid(strNone, 2): Exit Function
    End If
    
    ExecuteClinicPreSwap = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function ExecuteClinicSwap(ByVal lng病人ID As Long, _
    ByVal intInsure As Integer, ByVal lng结帐ID As Long, _
    ByVal lng结算序号 As Long, ByVal str预结算 As String, ByVal strNos As String, Optional ByVal str保险金额 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:医保结算接口
    '入参:  lng结帐ID:本次结帐的ID
    '出参:
    '返回:医保调用成功或非医保,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-20 17:15:15
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varNos As Variant
    Dim strBillNO As String, blnTrans As Boolean, blnTransMedicare As Boolean
    Dim p As Integer, strAdvance As String
    Dim strTmp As String, i As Long, strSQL As String
    Dim cur个帐支付 As Currency, cur医保基金 As Currency, varTemp As Variant
    
    On Error GoTo errHandle
     
    blnTrans = True
    '北京联合医保,要保存为划价单,所以暂未处理(需了解业务)
''    '1. 保存为划价单
''    If mblnSavePrice Then
''        '保存为划价单
''        '如果是联合医保,收费确定时实际却保存为划价单:传划价单明细,不在Oracle事务中执行
''        varNos = Split(mobjChargeInfor.Nos, ",")
''        For p = 1 To UBound(varNos)
''            strBillNO = mobjChargeInfor(p)
''            If Not gclsInsure.TranChargeDetail(1, strBillNO, 1, 0, "", , mobjChargeInfor.intInsure) Then
''                '删除划价单(继续处理)
''                Call DelMedicareTempNO(True, strBillNO)
''                gcnOracle.RollbackTrans: Exit Function
''            End If
''        Next
''        mblnYbBalanced = True   '医保已经结算
''        ExecuteClinicSwap = True
''        Exit Function
''    End If
      
    If MCPAR.医保接口打印票据 And MCPAR.医保不走票号 = False Then
        '不严格控制票据时保存当前票号
        If Not gblnStrictCtrl Then
            zlDatabase.SetPara "当前收费票据号", mCurBillType.strInvoice, glngSys, 1121, zlStr.IsHavePrivs(mstrPrivs, "参数设置")
        End If
    End If
    
    cur个帐支付 = 0: cur医保基金 = 0
    If str预结算 <> "" Then
        varTemp = Split(str预结算, "||")
        For i = 0 To UBound(varTemp)
            If Split(varTemp(i), "|")(0) = mstr个人帐户 Then
                cur个帐支付 = cur个帐支付 + CCur(Val(Split(varTemp(i), "|")(1)))
            ElseIf Split(varTemp(i), "|")(0) = "医保基金" Then
                cur医保基金 = cur医保基金 + CCur(Val(Split(varTemp(i), "|")(1)))
            End If
        Next
    End If
    varTemp = Split(str保险金额, ";") 'cur实收合计;cur进入统筹;cur全自付;cur先自付
    
    strAdvance = CStr(lng结算序号)
    If Not gclsInsure.ClinicSwap(lng结帐ID, cur个帐支付, cur医保基金, _
                        CCur(Val(varTemp(2))), CCur(Val(varTemp(3))), intInsure, strAdvance) Then
        gcnOracle.RollbackTrans:  Exit Function
    End If
    
    blnTransMedicare = True
    
    If strAdvance = CStr(lng结算序号) Then strAdvance = ""
     
    If strAdvance = "" Then
       gcnOracle.CommitTrans
       Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicSwap, True, intInsure)
       ExecuteClinicSwap = True: Exit Function
    End If
    
    If Not zlInsureCheck(str预结算, strAdvance) Then
        '修改校对标志
        ' Zl_病人门诊收费_医保更新
        strSQL = "Zl_病人门诊收费_医保更新("
        '  结帐id_In   门诊费用记录.结帐id%Type,
        strSQL = strSQL & lng结帐ID & ","
        '  结算序号_In 病人预交记录.结算序号%Type,
        strSQL = strSQL & "Null,"
        '  保险结算_In Varchar2
        strSQL = strSQL & "Null)"
        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
       gcnOracle.CommitTrans
       Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicSwap, True, intInsure)
       ExecuteClinicSwap = True: Exit Function
    End If
    
    '进行数据重新修正
    '需要修正结算数据
    'Zl_门诊收费结算_Modify
    strSQL = "Zl_门诊收费结算_Modify("
    '  操作类型_In   Number,
    strSQL = strSQL & "" & 2 & ","
    '  病人id_In     门诊费用记录.病人id%Type,
    strSQL = strSQL & "" & lng病人ID & ","
    '  结帐id_In     病人预交记录.结帐id%Type,
    strSQL = strSQL & "" & lng结帐ID & ","
    '  结算方式_In   Varchar2,
    strSQL = strSQL & "'" & strAdvance & "')"
    '  冲预交_In     病人预交记录.冲预交%Type,
    '  退支票额_In   病人预交记录.冲预交%Type,
    '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
    '  卡号_In       病人预交记录.卡号%Type := Null,
    '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
    '  交易说明_In   病人预交记录.交易说明%Type := Null,
    '  缴款_In       病人预交记录.缴款%Type := Null,
    '  找补_In       病人预交记录.找补%Type := Null,
    '  误差金额_In   门诊费用记录.实收金额%Type := Null,
    '  完成结算_In Number:=0
    ') As
    '  ------------------------------------------------------------------------------------------------------------------------------
    '  --功能:收费结算时,修改结算的相关信息
    '  --操作类型_In:
    '  --   0-普通收费方式:
    '  --     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
    '  --     ②冲预交_In:如果涉及预交款,则传入本次的冲预交,非正常收费时,传入零
    '  --     ③退支票额_In:如果涉及退支票,则传入本次的退支票额,非正常收费时,传入零
    '  --   1.三方卡结算:
    '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
    '  --     ②冲预交_In: 传入零
    '  --     ③退支票额_In:传入零
    '  --     ④卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
    '  --   2-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
    '  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
    '  --     ②冲预交_In: 传入零
    '  --     ③退支票额_In:传入零
    '  --   3-消费卡结算:
    '  --     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."
    '  --     ②冲预交_In: 传入零
    '  --     ③退支票额_In:传入零
    '  -- 误差金额_In:存在误差费时,传入
    '  -- 完成结算_In:1-完成收费;0-未完成收费
    '  ------------------------------------------------------------------------------------------------------------------------------
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    
    '修改校对标志
    ' Zl_病人门诊收费_医保更新
    strSQL = "Zl_病人门诊收费_医保更新("
    '  结帐id_In   门诊费用记录.结帐id%Type,
    strSQL = strSQL & lng结帐ID & ","
    '  结算序号_In 病人预交记录.结算序号%Type,
    strSQL = strSQL & "Null,"
    '  保险结算_In Varchar2
    strSQL = strSQL & "Null)"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    gcnOracle.CommitTrans: blnTrans = False
    Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicSwap, True, intInsure)
    ExecuteClinicSwap = True
    Exit Function
errHandle:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    '医保和HIS不是同一个事务,HIS事务失败,但医保可能已上传,所以需要调"取消交易"接口
    If blnTransMedicare Then Call gclsInsure.BusinessAffirm(交易Enum.Busi_ClinicSwap, False, intInsure)
    Call SaveErrLog
End Function

Private Function ExecuteYBIdentifyCancel(ByVal lng病人ID As Long, ByVal intInsure As Integer) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:取消医保病人身份验证
    '返回:返回假时不退出界面或清除操作
    '编制:刘兴洪
    '日期:2014-06-09 14:37:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    ExecuteYBIdentifyCancel = True
    If mbytMode = EM_MULTI_查看 Then Exit Function
    If lng病人ID = 0 Then Exit Function
    On Error GoTo errHandle
    ExecuteYBIdentifyCancel = gclsInsure.IdentifyCancel(0, lng病人ID, intInsure)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Sub zlExeBalanceWinRefrshData(ByVal blnSaveOK As Boolean, _
    ByRef objDelBalance As clsCliniDelBalance)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:执行退费结算操作后的刷新操作
    '入参:blnSaveOK-是否保存成功
    '     objChargeInfor-结算信息
    '编制:刘兴洪
    '日期:2014-06-17 10:50:41
    '说明:之所要独立出来,主要原因是解决医保调试的问题(模态窗体不好调试)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strPrintNos As String, strReclaimInvoice As String
    Dim strNo As String, lng打印ID As Long
    
    On Error GoTo errHandle
    
    If blnSaveOK = False Then Exit Sub
    
    If objDelBalance.blnOnePatiPrint Then
        strPrintNos = "'" & Replace(objDelBalance.strOnePatiPrintNos, ",", "';'") & "'"
    Else
        strPrintNos = objDelBalance.PrintNOs
    End If
 
    If Mid(objDelBalance.CurDelNos, 1, 1) = "," Then
        strNo = Split(objDelBalance.CurDelNos, ",")(1)
    Else
        strNo = Split(objDelBalance.CurDelNos, ",")(0)
    End If
    
    strReclaimInvoice = zlGetReclaimInvoice(strPrintNos)
    If gTy_Module_Para.byt票据分配规则 <> 0 And strReclaimInvoice <> "" Then
        If InStr(1, mstrPrivs, "退费核收发票") > 0 Then
            If MsgBox("注意:" & vbCrLf & " 当前退费的单据中包含如下收费票据，是否回收这些票据?" & vbCrLf & strReclaimInvoice, _
                vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then GoTo Completed '不打印退费正常退出
        End If
    End If
    
    If gblnBillPrint Then
        If objDelBalance.blnOnePatiPrint Then
            If gobjBillPrint.zlEraseBill(objDelBalance.strOnePatiPrintNos, 0) = False Then Exit Sub
        Else
            If gobjBillPrint.zlEraseBill(mCurBillType.strAllNOs, 0) = False Then Exit Sub
        End If

    End If
    
   '打印退费单据
    Call PrintDelBill(objDelBalance.AllNos, objDelBalance.CurDelNos, strNo, objDelBalance.病人ID, _
        objDelBalance.退费时间, objDelBalance.部分退费, objDelBalance.回收发票, strReclaimInvoice, objDelBalance.blnOnePatiPrint, objDelBalance.strOnePatiPrintNos)


Completed:
    mblnOK = True: Call ClearFace
    
    If txtNO.Visible Then
        txtNO.SetFocus: Exit Sub
    End If
    Unload Me
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub
Private Function IsFeeAllDel() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:判断是否全退费
    '返回:合退费返回成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-14 16:36:48
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnDel As Boolean, blnAllDel As Boolean
    Dim j As Long
    On Error GoTo errHandle
    '1.看是否为全选，全选就原样退
    If mCurBillType.bln单张部分退费 Then Exit Function
    With vsBill
        For j = 1 To .Rows - 1
            If .TextMatrix(j, .ColIndex("单据号")) <> "" And Abs(Val(.TextMatrix(j, .ColIndex("选择")))) <> 1 Then
                IsFeeAllDel = False: Exit Function
            End If
        Next
    End With
    
    '2.当前退费与本次收费单据完全一致
    If CheckDiff(Replace(mCurBillType.strAllNOs, "'", ""), Replace(mCurBillType.strNos, "'", "")) = False Then Exit Function
    
    
    IsFeeAllDel = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function
Private Function GetFeeDelNumRecord(ByVal strAllNOs As String) As ADODB.Recordset
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取费用的剩余数量集
    '入参:strAllNos-所有单据
    '出参:
    '返回:记录集(NO,序号,原始数量,剩余数量)
    '编制:刘兴洪
    '日期:2014-07-15 11:35:41
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    '78004,冉俊明,2014-9-16,在参数“药品单位”为“门诊单位”时出错
    strSQL = "" & _
    "   Select A.NO,nvl(A.价格父号,A.序号) as 序号,a.收费细目ID,A.记录性质,A.结帐ID, " & _
    "         Decode(A.记录性质,1, 1,0)*decode(A.记录状态,1,1,3,1,0)*Avg(nvl(A.付数,1) *数次) as 原始数量," & _
    "         Avg(nvl(A.付数,1) *数次) as 数量" & _
    "   From 门诊费用记录 A" & _
    "   Where A.NO in (select J.Column_value From  Table(f_str2List([1])) J )  " & _
    "       And mod(a.记录性质,10)=1 And nvl(A.费用状态,0)<>1" & _
    "   Group by A.NO,nvl(A.价格父号,A.序号),A.记录性质,A.记录状态,A.结帐ID,a.收费细目ID"
    
    strSQL = "" & _
    "   Select /*+ Rule */ A.NO,A.序号,A.收费细目ID," & _
    "      sum(A.原始数量/" & IIf(gbln药房单位, "nvl(B." & gstr药房包装 & ",1)", "1") & ") as 原始数量, " & _
    "      sum(A.数量/" & IIf(gbln药房单位, "nvl(B." & gstr药房包装 & ",1)", "1") & ")  as 剩余数量 " & _
    "   From (" & strSQL & ") A,药品规格 B" & _
    "   Where A.收费细目ID=B.药品ID(+) " & _
    "   Group by A.NO,A.序号,a.收费细目ID" & _
    "   Order by NO,序号"

    On Error GoTo errHandle
    Set GetFeeDelNumRecord = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strAllNOs)
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckIsAllDel(ByVal strAllNOs As String, _
    Optional ByVal blnBillSaved As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查所有费用是否全退
    '入参:strAllNos-所有单据,多个用逗号分隔
    '出参:
    '返回:所有全退时,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-15 11:28:38
    
    '修改：104573
    '入参：
    '   blnBillSaved - 费用数据是否已保存，已保存的只要读出的数据有剩余数量就表示未退完，主要是票据打印和异常重退调用
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim strNo As String, int序号 As Integer
    Dim blnFind As Boolean, dbl剩余数量 As Double
    Dim j As Long, k As Long
    
    On Error GoTo errHandle
    If mbytMode = EM_MULTI_退费 Then
        With vsBill
            For j = 1 To vsBill.Rows - 1
                If Abs(Val(.TextMatrix(j, .ColIndex("选择")))) <> 1 And InStr(strAllNOs, .TextMatrix(j, .ColIndex("单据号"))) > 0 Then
                   CheckIsAllDel = False: Exit Function
                End If
            Next
        End With
    End If
    
    Set rsTemp = GetFeeDelNumRecord(strAllNOs)
    If blnBillSaved = False Then
        Do While Not rsTemp.EOF
            strNo = Nvl(rsTemp!NO): int序号 = Val(Nvl(rsTemp!序号))
            dbl剩余数量 = Val(Nvl(rsTemp!剩余数量))
            If dbl剩余数量 <> 0 Then
                With vsBill
                    k = vsBill.FindRow(strNo, , .ColIndex("单据号"))
                    If k <= 0 Then Exit Function
                    blnFind = False
                    For j = k To vsBill.Rows - 1
                        If .TextMatrix(j, .ColIndex("单据号")) <> strNo Then Exit For
                        If Abs(Val(.TextMatrix(j, .ColIndex("选择")))) <> 1 _
                            And mbytMode <> EM_MULTI_异常重退 Then
                            CheckIsAllDel = False: Exit Function
                        End If
                        If Val(.RowData(j)) = int序号 Then
                            If dbl剩余数量 <> Val(.Cell(flexcpData, j, .ColIndex("数量"))) Then
                               CheckIsAllDel = False: Exit Function
                            End If
                            blnFind = True: Exit For
                        End If
                    Next
                End With
                If blnFind = False Then Exit Function
            End If
            rsTemp.MoveNext
        Loop
    Else
        If rsTemp.RecordCount = 0 Then
            CheckIsAllDel = True: Exit Function
        End If
        Do While Not rsTemp.EOF
            If RoundEx(Val(Nvl(rsTemp!剩余数量)), 6) <> 0 Then
                '票据打印时，退费数据已保存，此时只要有剩余数量不等于零就表示没退完
                CheckIsAllDel = False: Exit Function
            End If
            rsTemp.MoveNext
         Loop
    End If
    CheckIsAllDel = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function ExecuteReDelFee() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:对异常单据重新退费
    '返回:退费成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-07-17 15:43:07
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim frmBalance As frmClinicDelBalance, objDelBalance As clsCliniDelBalance
    Dim bln全退 As Boolean, str结帐ID As String, lng结帐ID As Long, lng冲销ID As Long
    Dim strNos As String, varData As Variant, strCmdCaptions As String
    Dim cllPro  As New Collection, strInvoices As String, strInvoice As String
    Dim lngCheck病人ID As Long, intCheckInsure   As Integer, strYBPati As String
    Dim dtDelDate As Date, blnTrans As Boolean, strNo As String
    Dim str序号 As String, j As Long, strPrintNOInfor As String
    Dim strSQL As String, rsTemp As New ADODB.Recordset
    Dim cur个帐透支 As Currency, str保险金额 As String 'cur实收合计;cur进入统筹;cur全自付;cur先自付
    Dim strReturn As String, strReturnRecipt As String '退费处方信息，格式：NO,药房ID|NO,药房ID|…
    Dim rs药品记录 As ADODB.Recordset, lng领用ID As Long
    Dim strAllBalance As String, strAdvance As String
    
    On Error GoTo errHandle
    '并发检查
    If zlIsCheckExistErrBill(mlng结算序号) = False Then
        MsgBox "当前异常单据已被处理，你不能继续！", vbInformation, gstrSysName
        Exit Function
    End If
    If zlCheckOtherSessionDoing(mlng结算序号) Then
        MsgBox "当前单据正在其它收费窗口中进行处理，你不能继续！", vbInformation, gstrSysName
        Exit Function
    End If
    '105432,三方卡结算方式有效性检查
    If ThreeBalanceCheck(zlFromIDGetChargeBalance(2, mCurBillType.strAllNOs, mblnNOMoved, , True), mrs结算方式, mcllForceDelToCash) = False Then Exit Function
    
    bln全退 = CheckIsAllDel(mCurBillType.strAllNOs, True)
    If Not bln全退 Then
        If zlStr.IsHavePrivs(mstrPrivs, "部份退费") = False Then
            MsgBox "你没有权限执行部份退费操作！", vbInformation, gstrSysName
            vsBill.SetFocus: Exit Function
        End If
        If zlStr.IsHavePrivs(mstrPrivs, "退费核收发票") Then
            If frmReInvoice.ShowMe(Me, mstrNo, Val(txtAllTotal.Text), 0, strInvoices) = False Then
                vsBill.SetFocus: Exit Function
            End If
        End If
    End If
    With vsBill
        str序号 = "": strNo = ""
        For j = 1 To vsBill.Rows - 1
            If strNo <> Trim(.TextMatrix(j, .ColIndex("单据号"))) Then
                If str序号 <> "" Then
                    strPrintNOInfor = strPrintNOInfor & ";" & strNo & ":" & Mid(str序号, 2)
                End If
                strNo = .TextMatrix(j, .ColIndex("单据号"))
                str序号 = ""
            End If
            str序号 = str序号 & "," & CLng(vsBill.RowData(j))
        Next
        If strNo <> "" And str序号 <> "" Then
            strPrintNOInfor = strPrintNOInfor & ";" & strNo & ":" & str序号
        End If
    End With
    
    Set objDelBalance = New clsCliniDelBalance
    'bytType-查找类型:0-根据结帐ID查找;1-根据结算序号查找,2-根据单据号来获取结算方式
    Set objDelBalance.rsBalance = zlFromIDGetChargeBalance(2, mCurBillType.strAllNOs, False)
    Set objDelBalance.rs结算方式 = mrs结算方式
    
    lng结帐ID = mCurBillType.lng结帐ID
    lng冲销ID = mCurBillType.lng冲销ID
    
    If mCurBillType.intInsure <> 0 And lng结帐ID <> 0 And MCPAR.医保接口打印票据 Then
        If InputFactNo(lng领用ID, strInvoice) = False Then Exit Function
    End If
    
    dtDelDate = zlDatabase.Currentdate
    
    '将日期更新为当前日期
    '重新收费时，将收费的登记时按新时间进行登记处理
    'Zl_门诊收费异常_Update
    strSQL = "Zl_门诊收费异常_Update("
    '  No_In       门诊费用记录.No%Type,
    strSQL = strSQL & "NULL,"
    '  登记时间_In 门诊费用记录.登记时间%Type,
    strSQL = strSQL & "to_date('" & Format(dtDelDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
    '  结帐id_In   门诊费用记录.结帐id%Type := Null
    strSQL = strSQL & "" & mCurBillType.lng冲销ID & ")"
    zlAddArray cllPro, strSQL
    If mCurBillType.lng结帐ID <> 0 Then
        'Zl_门诊收费异常_Update
        strSQL = "Zl_门诊收费异常_Update("
        '  No_In       门诊费用记录.No%Type,
        strSQL = strSQL & "NULL,"
        '  登记时间_In 门诊费用记录.登记时间%Type,
        strSQL = strSQL & "to_date('" & Format(dtDelDate, "yyyy-mm-dd HH:MM:SS") & "','yyyy-mm-dd hh24:mi:ss'),"
        '  结帐id_In   门诊费用记录.结帐id%Type := Null
        strSQL = strSQL & "" & mCurBillType.lng结帐ID & ")"
        zlAddArray cllPro, strSQL
    End If
    
    '多单据分单据结算时，没有重收记录，lng结帐ID肯定等于0
    '先退医保
    If mCurBillType.intInsure <> 0 And lng结帐ID <> 0 And MCPAR.门诊结算作废 Then
        '如果是医保,出现异常,肯定是只有重收部分才出现异常
        '字段:类型 ,结帐ID, 记录性质, 结算方式, 摘要, 卡类别ID, 卡类别名称, 自制卡, 结算卡序号, 结算号码, 卡号, 交易流水号, 交易说明, 结算序号, 校对标志, 医保, 消费卡id
        '            是否密文,是否全退,是否退现,冲预交
        '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
        mrsBalance.Filter = "结帐ID=" & lng结帐ID & " And 类型=2 "
        If mrsBalance.EOF Then
            '79237,冉俊明,2014-11-5
            '有可能已经成功进行了医保结算，但是医保报销金额为零
            strSQL = "" & _
                "   Select 1" & _
                "   From 病人预交记录 A, 保险结算记录 B" & _
                "   Where a.结帐id = b.记录id And a.记录性质 = 3 And a.记录状态 = 1 And b.性质 = 1 " & _
                "         And a.结算序号 = [1] And Rownum < 2"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查是否已调用医保接口", mlng结算序号)
            If rsTemp.EOF Then
                '未进行医保预结算,因此,需要重新预结,然后结算
                '可能存在重新收费,因此,需要调用身份验证接口(Identifiy)
                'strAdvace:医保部分退时:传入1,表示医保部分退后再重新收费的身份验证;其他传入: 空
                lngCheck病人ID = mCurBillType.lng病人ID
                intCheckInsure = mCurBillType.intInsure
                strYBPati = gclsInsure.Identify(0, lngCheck病人ID, intCheckInsure, 1)
                If strYBPati = "" Then
                     MsgBox "医保身份验证失败,不允许继续退费!", vbOKOnly + vbDefaultButton1 + vbExclamation, gstrSysName
                     Exit Function
                End If
                 
                If Val(CLng(Split(strYBPati, ";")(8))) <> mCurBillType.lng病人ID Then
                    MsgBox "医保验证的病人与退费的病人不是同一个病人!", vbInformation, gstrSysName
                    Call ExecuteYBIdentifyCancel(mCurBillType.lng病人ID, mCurBillType.intInsure)
                    Exit Function
                End If
                blnTrans = True
                zlExecuteProcedureArrAy cllPro, Me.Caption, True
                
                '更新重收记录的保险信息，因为在退费时可能未更新，为了保险起见，再重新更新一遍
                '77951,冉俊明,2014-9-16
                If ExecuteInsureInfoUpdate(lng结帐ID, str保险金额) = False Then Exit Function
                '读取个帐余额
                cur个帐透支 = mdbl个帐透支
                mdbl个帐余额 = gclsInsure.SelfBalance(mCurBillType.lng病人ID, CStr(Split(strYBPati, ";")(1)), 10, cur个帐透支, mCurBillType.intInsure)
                mdbl个帐透支 = cur个帐透支
                '77058
                If ExcuteInsureReCharge(mCurBillType.lng病人ID, mCurBillType.intInsure, lng结帐ID, mlng结算序号, str保险金额, _
                            strNo, lng领用ID, strInvoice, dtDelDate) = False Then Exit Function
                blnTrans = False
                Set cllPro = New Collection
            End If
        End If
    ElseIf mCurBillType.intInsure <> 0 And mblnDelByNo Then
        strAllBalance = GetYBOldBalance(mCurBillType.lng病人ID, mCurBillType.intInsure, mCurBillType.lng冲销ID)
        
        blnTrans = True
        zlExecuteProcedureArrAy cllPro, Me.Caption, True
        If ExecuteClinicDelNo(mCurBillType.lng病人ID, mCurBillType.intInsure, lng冲销ID, mCurBillType.lng原结帐ID, strAdvance, True) = False Then
            gcnOracle.RollbackTrans
            Exit Function
        End If
        gcnOracle.CommitTrans: blnTrans = False
        
        If zlInsureCheck(strAllBalance, strAdvance) And strAdvance <> "" Then
            '退费和收费不一致时,需要效对
            ' Zl_门诊退费结算_Modify
            strSQL = "Zl_门诊退费结算_Modify("
            '  操作类型_In   Number,
            '  --   0-原样退
            '  --      原样结算一起全退,所有校对标志都为1,医保调用成功后,调整为2,完成后变成0
            '  --   1-普通退费方式:
            '  --     ①结算方式_IN:允许传入多个,格式为:"结算方式|结算金额|结算号码|结算摘要||.." ;也允许传入空.
            '  --     ②冲预交_In:如果涉及预交款,则传入本次的退预交,非正常收费时,传入零(<0 表示退预交款;>0 表示将剩余款生成预交记录
            '  --   2.三方卡退费结算:
            '  --     ①结算方式_IN:只能传入一个结算方式,但允许包含一些辅助信息,格式为:"结算方式|结算金额|结算号码|结算摘要"
            '  --     ②退预交_In: 传入零
            '  --     ③卡类别ID_IN,卡号_IN,交易流水号_IN,交易说明_In:需要传入
            '  --   3-医保结算(如果存在医保的结算,则要先删除原医保结算,后按新传入的更新)
            '  --     ①结算方式_IN:允许传入多个,格式为:结算方式|结算金额||.."
            '  --     ②退预交_In: 传入零
            '  --     ③退支票额_In:传入零
            '  --   4-消费卡结算:
            '  --     ①结算方式_IN:允许一次刷多张卡,格式为:卡类别ID|卡号|消费卡ID|消费金额||."
            '  --     ②退预交_In: 传入零
            '  --     ③退支票额_In:传入零
            strSQL = strSQL & "" & 3 & ","
            '  病人id_In     门诊费用记录.病人id%Type,
            strSQL = strSQL & "" & mCurBillType.lng病人ID & ","
            '  冲销id_In     病人预交记录.结帐id%Type,
            strSQL = strSQL & "" & mCurBillType.lng冲销ID & ","
            '  结算方式_In   Varchar2,
            strSQL = strSQL & "'" & strAdvance & "')"
            '  退预交_In     病人预交记录.冲预交%Type := Null,
            '  卡类别id_In   病人预交记录.卡类别id%Type := Null,
            '  卡号_In       病人预交记录.卡号%Type := Null,
            '  交易流水号_In 病人预交记录.交易流水号%Type := Null,
            '  交易说明_In   病人预交记录.交易说明%Type := Null,
            '  缴款_In       病人预交记录.缴款%Type := Null,
            '  找补_In       病人预交记录.找补%Type := Null,
            '  误差金额_In   门诊费用记录.实收金额%Type := Null,
            '  完成退费_In   Number := 0,
            '  原结帐id_In   病人预交记录.结帐id%Type := Null
            zlDatabase.ExecuteProcedure strSQL, Me.Caption
        End If
    
        '修改校对标志
        ' Zl_病人门诊收费_医保更新
        strSQL = "Zl_病人门诊收费_医保更新("
        '  结帐id_In   门诊费用记录.结帐id%Type,
        strSQL = strSQL & mCurBillType.lng冲销ID & ","
        '  结算序号_In 病人预交记录.结算序号%Type,
        strSQL = strSQL & "Null,"
        '  保险结算_In Varchar2
        strSQL = strSQL & "Null)"
        zlDatabase.ExecuteProcedure strSQL, Me.Caption
    End If
    
    '2.再退一卡通(老版本)
    '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡
    mrsBalance.Filter = "类型=4 "
    objDelBalance.rsBalance.Filter = "类型=4 "
    If mrsBalance.EOF And objDelBalance.rsBalance.EOF = False Then
ReDOOneCard:
        If CheckOnCardValied(objDelBalance.rsBalance) = False Then Exit Function
        
        blnTrans = True
        zlExecuteProcedureArrAy cllPro, Me.Caption, True
        If Not ExecuteOneCardDelInterface(objDelBalance.rsBalance, lng冲销ID) Then
            mrsBalance.Filter = 0
            If mCurBillType.intInsure <> 0 Then
                If frmVerfyCodeInput.ShowMsg(Me, "单据[" & mCurBillType.strDelNOs & "]已经退费成功,但一卡通交易失败,[异常单据]必须输入验证码,建议不进行异常单据保存", strCmdCaptions) = False Then
                    gcnOracle.BeginTrans: blnTrans = True
                    GoTo ReDOOneCard:
                End If
            End If
            Exit Function
        End If
        blnTrans = False
        Set cllPro = New Collection
    End If
    
    '4.显示结算界面
    mCurBillType.lng结算序号 = mlng结算序号 '记录用于打印红票
    If strPrintNOInfor <> "" Then strPrintNOInfor = Mid(strPrintNOInfor, 2)
    With objDelBalance
        .intInsure = mCurBillType.intInsure
        .CurDelNos = mCurBillType.strDelNOs
        .AllNos = mCurBillType.strAllNOs
        .PrintNOs = strPrintNOInfor
        
        .PatiUseType = mstrUseType
        .SaveBilled = True
        
        .ShareUserID = mlngShareUseID
        .病人ID = mCurBillType.lng病人ID
        .冲销ID = lng冲销ID
        .当前发票号 = strInvoice
        .回收发票 = strInvoices
        .结算序号 = mlng结算序号
        .结帐ID = lng结帐ID
        .缺省结算方式 = mCurBillType.str结算方式
        .退费合计 = -1 * GetDelMoney
        .费别 = mCurBillType.str费别
        .年龄 = mCurBillType.str年龄
        .性别 = mCurBillType.str性别
        .姓名 = mCurBillType.str姓名
        .医保不走票号 = MCPAR.医保不走票号
        .原结帐ID = mCurBillType.lng原结帐ID
        .退费时间 = dtDelDate
        .部分退费 = Not bln全退
    End With
    
    Set frmBalance = New frmClinicDelBalance
    If frmBalance.zlDelCharge(Me, EM_FUN_重退, mlngModule, mstrPrivs, objDelBalance, cllPro, , mcllForceDelToCash) = False Then Exit Function
    
    '81190,冉俊明,退费业务向发药机上传退费信息
    On Error Resume Next
    If mblnDrugMachine Then
        Dim strData As String '门诊处方退药格式：费用ID1,退药数量1;费用ID2,退药数量2;...
        '本次退的减去重收的就是实际退的
        strSQL = "Select Max(Decode(a.记录状态, 2, a.Id, 0)) As 费用id, -1 * Nvl(Sum(a.付数 * a.数次), 0) As 退药数量" & vbNewLine & _
                " From 门诊费用记录 A,(Select Distinct 结帐ID From 病人预交记录 Where 结算序号 = [1]) B" & vbNewLine & _
                " Where a.结帐id = b.结帐ID And Mod(a.记录性质, 10) = 1 And a.收费类别 In ('5', '6', '7')" & vbNewLine & _
                " Group By NO, Nvl(价格父号, 序号)" & vbNewLine & _
                " Having Nvl(Sum(a.付数 * a.数次), 0) <> 0"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询本次退费项目", objDelBalance.结算序号)
        Do While Not rsTemp.EOF
            strData = strData & ";" & Nvl(rsTemp!费用id) & "," & Nvl(rsTemp!退药数量)
            rsTemp.MoveNext
        Loop
        If strData <> "" Then
            strData = Mid(strData, 2)
            Call mobjDrugMachine.Operation(gstrDBUser, Val("24-处方退药(完整/部分)"), strData, strReturn)
        End If
    ElseIf mblnDrugPacker Then
        strSQL = "Select a.No, a.执行部门id" & _
            "   From 门诊费用记录 A, 病人预交记录 B" & _
            "   Where a.结帐id = b.结帐id And a.记录状态=2  And a.收费类别 In ('5', '6', '7') And b.结算序号 = [1]"
        Set rs药品记录 = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mlng结算序号))
        
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
    
    If Not gfrmMain Is Nothing Then
        Call zlExeBalanceWinRefrshData(True, objDelBalance)
    End If
    ExecuteReDelFee = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function


Private Function CheckOnCardValied(ByVal rsBalance As ADODB.Recordset) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查一卡通是否合法
    '入参:rsBalance-原始的结帐数据
    '返回:合法,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-08-31 12:00:30
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strCardNo As String
    On Error GoTo errHandle
    
    '字段:类型,记录性质,结算方式,摘要,卡类别ID,卡类别名称,自制卡,结算卡序号,结算号码,卡号,交易流水号, 交易说明,结算序号,校对标志,医保,消费卡id
    '     是否密文,是否全退,是否退现,冲预交
    '类型:0-普通结算;1-预交款;2-医保,3-一卡通;4-一卡通(老);5-消费卡)
    mrsBalance.Filter = "类型=4"
    If rsBalance.RecordCount = 0 Then CheckOnCardValied = True: Exit Function
    If mobjICCard Is Nothing Then
        On Error Resume Next
        Set mobjICCard = CreateObject("zlICCard.clsICCard")
        On Error GoTo 0
        If mobjICCard Is Nothing Then
            MsgBox "一卡通接口创建失败,不能进行退费!请检查接口文件.", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    strCardNo = mobjICCard.Read_Card(Me)
    If strCardNo <> Nvl(rsBalance!卡号) Then
        MsgBox "当前卡号与扣款卡号不一致,不能进行退费.", vbInformation, gstrSysName
        Exit Function
    End If
    CheckOnCardValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function CheckDelAppliedValied(ByVal bytMode As gEM_ChargeDelType, ByRef strNos As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查退费申请的合法性
    '入参:
    '出参:strNos-本次退费申请的单据号,多个用逗号分离
    '返回:退费申请合法成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-08-05 11:20:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strNo As String, j As Long, i As Long
    Dim rsTemp As ADODB.Recordset, varTemp As Variant
    
    On Error GoTo errHandle
    strNos = ""
    With vsBill
        strNo = ""
        For j = 1 To vsBill.Rows - 1
            If strNo <> Trim(.TextMatrix(j, .ColIndex("单据号"))) Then
                strNo = .TextMatrix(j, .ColIndex("单据号"))
                If InStr(strNos & ",", "," & strNo & ",") = 0 Then
                    If Abs(Val(.TextMatrix(j, .ColIndex("选择")))) = 1 Then
                        strNos = strNos & "," & strNo
                    End If
                End If
            End If
        Next
    End With
    If strNos <> "" Then strNos = Mid(strNos, 2)
    If strNos = "" Then
        MsgBox "未选择单据，请选择！", vbInformation + vbOKOnly, gstrSysName
        vsBill.SetFocus: Exit Function
    End If
    
    If bytMode = EM_MULTI_拒绝申请 Then
        If Trim(txt退费摘要.Text) = "" Then
            MsgBox "必须输入拒绝原因！", vbInformation + vbOKOnly, gstrSysName
            If txt退费摘要.Visible And txt退费摘要.Enabled Then txt退费摘要.SetFocus
            Exit Function
        End If
    End If
    
    Set rsTemp = GetApply(strNos, 1)
    varTemp = Split(strNos, ",")
    For i = 0 To UBound(varTemp)
        strNo = varTemp(i)
        Select Case bytMode
            Case EM_MULTI_退费申请
                rsTemp.Filter = "NO='" & strNo & "' And 状态=0" '已申请
                If rsTemp.RecordCount <> 0 Then
                    MsgBox "单据:" & strNo & " 已被退费申请，不用再进行申请！", vbInformation, gstrSysName
                    Exit Function
                End If
                rsTemp.Filter = "NO='" & strNo & "' And 状态=1" '已审核
                If rsTemp.RecordCount <> 0 Then
                    MsgBox "单据:" & strNo & " 已被退费申请并进行了审核，不能再进行申请！", vbInformation, gstrSysName
                    Exit Function
                End If
            Case EM_MULTI_取消申请
                rsTemp.Filter = "NO='" & strNo & "' And 状态=0" '已申请
                If rsTemp.RecordCount = 0 Then
                    MsgBox "单据:" & strNo & " 已被取消申请，不用再进行取消申请！", vbInformation, gstrSysName
                    Exit Function
                End If
            Case EM_MULTI_退费审核, EM_MULTI_拒绝申请, EM_MULTI_取消审核
                rsTemp.Filter = "NO='" & strNo & "' And 申请时间=#" & mstrApplyTime & "#" '已申请
                If rsTemp.RecordCount = 0 Then
                    MsgBox "单据:" & strNo & " 已被取消申请，不能进行" & _
                            IIf(bytMode = EM_MULTI_退费审核, "退费审核", IIf(bytMode = EM_MULTI_拒绝申请, "拒绝申请", "取消审核")) & "！", vbInformation, gstrSysName
                    Exit Function
                End If
                If bytMode = EM_MULTI_退费审核 Then
                    rsTemp.Filter = "(NO='" & strNo & "' And 状态=1) " & _
                                    "Or (NO='" & strNo & "' And 状态=2 And 申请时间=#" & mstrApplyTime & "#)" '已审核或拒绝
                    If rsTemp.RecordCount <> 0 Then
                        MsgBox "单据:" & strNo & " 已被退费审核或拒绝申请，不能再进行退费审核！", vbInformation, gstrSysName
                        Exit Function
                    End If
                ElseIf bytMode = EM_MULTI_拒绝申请 Then
                    rsTemp.Filter = "NO='" & strNo & "' And 状态=2 And 申请时间=#" & mstrApplyTime & "#"
                    If rsTemp.RecordCount <> 0 Then
                        MsgBox "单据:" & strNo & " 已被拒绝申请，不能再进行拒绝申请！", vbInformation, gstrSysName
                        Exit Function
                    End If
                Else
                    rsTemp.Filter = "NO='" & strNo & "' And 状态=0 And 申请时间=#" & mstrApplyTime & "#"
                    If rsTemp.RecordCount <> 0 Then
                        MsgBox "单据:" & strNo & " 已被取消审核，不能再进行取消审核！", vbInformation, gstrSysName
                        Exit Function
                    End If
                    If BillExistDelete(strNo, 1) Then
                        MsgBox "单据:" & strNo & " 已退费，不能取消审核。", vbInformation, gstrSysName
                        Exit Function
                    End If
                End If
        End Select
    Next
    
    CheckDelAppliedValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If

End Function

Private Function SaveDelApplied(ByVal bytMode As gEM_ChargeDelType) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:保存退费申请
    '入参:strNos-申请的单据号,多个用逗号分离
    '返回:退费申请成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2014-08-05 11:14:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim cllPro As Collection, strSQL As String
    Dim strDate As String, varNO As Variant
    Dim str原因 As String, strNos As String
    Dim i As Long
    
    On Error GoTo errHandle
    
    If CheckDelAppliedValied(bytMode, strNos) = False Then Exit Function
      
    strDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    str原因 = txt退费摘要.Text
    Set cllPro = New Collection
    varNO = Split(strNos, ",")
    For i = 0 To UBound(varNO)
        Select Case bytMode
            Case EM_MULTI_退费申请
                'Zl_病人退费申请_Apply
                strSQL = "Zl_病人退费申请_Apply("
                '  操作类型_In Number,
                strSQL = strSQL & "" & "0" & ","
                '  No_In       病人退费申请.No%Type,
                strSQL = strSQL & "'" & varNO(i) & "',"
                '  记录性质_In 病人退费申请.记录性质%Type,
                strSQL = strSQL & "" & "1" & ","
                '  申请人_In   病人退费申请.申请人%Type,
                strSQL = strSQL & "'" & UserInfo.姓名 & "',"
                '  申请时间_In 病人退费申请.申请时间%Type,
                strSQL = strSQL & "" & "To_Date('" & strDate & "','YYYY-MM-DD HH24:MI:SS')" & ","
                '  申请原因_In 病人退费申请.申请原因%Type := Null
                strSQL = strSQL & "'" & str原因 & "')"
            Case EM_MULTI_取消申请
                'Zl_病人退费申请_Apply
                strSQL = "Zl_病人退费申请_Apply("
                '  操作类型_In Number,
                strSQL = strSQL & "" & "1" & ","
                '  No_In       病人退费申请.No%Type,
                strSQL = strSQL & "'" & varNO(i) & "',"
                '  记录性质_In 病人退费申请.记录性质%Type,
                strSQL = strSQL & "" & "1" & ","
                '  申请人_In   病人退费申请.申请人%Type,
                strSQL = strSQL & "'" & "" & "',"
                '  申请时间_In 病人退费申请.申请时间%Type,
                strSQL = strSQL & "" & "To_Date('" & mstrApplyTime & "','YYYY-MM-DD HH24:MI:SS')" & ")"
                '  申请原因_In 病人退费申请.申请原因%Type := Null
            Case EM_MULTI_退费审核
                'Zl_病人退费申请_Audit
                strSQL = "Zl_病人退费申请_Audit("
                '  No_In       病人退费申请.No%Type,
                strSQL = strSQL & "'" & varNO(i) & "',"
                '  记录性质_In 病人退费申请.记录性质%Type,
                strSQL = strSQL & "" & "1" & ","
                '  申请时间_In 病人退费申请.申请时间%Type,
                strSQL = strSQL & "" & "To_Date('" & mstrApplyTime & "','YYYY-MM-DD HH24:MI:SS')" & ","
                '  审核人_In   病人退费申请.审核人%Type,
                strSQL = strSQL & "'" & UserInfo.姓名 & "',"
                '  审核时间_In 病人退费申请.审核时间%Type,
                strSQL = strSQL & "" & "To_Date('" & strDate & "','YYYY-MM-DD HH24:MI:SS')" & ","
                '  审核原因_In 病人退费申请.审核原因%Type := Null,
                strSQL = strSQL & "'" & str原因 & "',"
                '  状态_In     病人退费申请.状态%Type := 1
                '--       状态_In：1-审核通过，2-拒绝申请(审核不通过)，3-取消审核
                strSQL = strSQL & "" & "1" & ")"
            Case EM_MULTI_拒绝申请
                'Zl_病人退费申请_Audit
                strSQL = "Zl_病人退费申请_Audit("
                '  No_In       病人退费申请.No%Type,
                strSQL = strSQL & "'" & varNO(i) & "',"
                '  记录性质_In 病人退费申请.记录性质%Type,
                strSQL = strSQL & "" & "1" & ","
                '  申请时间_In 病人退费申请.申请时间%Type,
                strSQL = strSQL & "" & "To_Date('" & mstrApplyTime & "','YYYY-MM-DD HH24:MI:SS')" & ","
                '  审核人_In   病人退费申请.审核人%Type,
                strSQL = strSQL & "'" & UserInfo.姓名 & "',"
                '  审核时间_In 病人退费申请.审核时间%Type,
                strSQL = strSQL & "" & "To_Date('" & strDate & "','YYYY-MM-DD HH24:MI:SS')" & ","
                '  审核原因_In 病人退费申请.审核原因%Type := Null,
                strSQL = strSQL & "'" & str原因 & "',"
                '  状态_In     病人退费申请.状态%Type := 1
                '--       状态_In：1-审核通过，2-拒绝申请(审核不通过)，3-取消审核
                strSQL = strSQL & "" & "2" & ")"
            Case EM_MULTI_取消审核
                'Zl_病人退费申请_Audit
                strSQL = "Zl_病人退费申请_Audit("
                '  No_In       病人退费申请.No%Type,
                strSQL = strSQL & "'" & varNO(i) & "',"
                '  记录性质_In 病人退费申请.记录性质%Type,
                strSQL = strSQL & "" & "1" & ","
                '  申请时间_In 病人退费申请.申请时间%Type,
                strSQL = strSQL & "" & "To_Date('" & mstrApplyTime & "','YYYY-MM-DD HH24:MI:SS')" & ","
                '  审核人_In   病人退费申请.审核人%Type,
                strSQL = strSQL & "" & "NULL" & ","
                '  审核时间_In 病人退费申请.审核时间%Type,
                strSQL = strSQL & "" & "NULL" & ","
                '  审核原因_In 病人退费申请.审核原因%Type := Null,
                strSQL = strSQL & "" & "NULL" & ","
                '  状态_In     病人退费申请.状态%Type := 1
                '--       状态_In：1-审核通过，2-拒绝申请(审核不通过)，3-取消审核
                strSQL = strSQL & "" & "3" & ")"
        End Select
        zlAddArray cllPro, strSQL
    Next
    On Error GoTo ErrRoll:
    zlExecuteProcedureArrAy cllPro, Me.Caption
    SaveDelApplied = True
    Exit Function
ErrRoll:
    gcnOracle.RollbackTrans
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function CheckIsExistDelErrBill(ByVal strNos As String, Optional ByRef str操作员姓名 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:根据单据号,检查是否存在退费异常记录
    '入参:
    '     strNOs=单据号,格式 NO1,NO2,NO3,...
    '出参:
    '     strUser=产生退费异常单据的操作员姓名
    '返回:存在退费异常单据,返回true,否则返回False
    '编制:冉俊明
    '日期:2014-08-18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String
    
    str操作员姓名 = ""
    If strNos = "" Then Exit Function
   
    On Error GoTo Errhand
    strSQL = "" & _
            " Select 操作员姓名" & _
            " From 门诊费用记录 A" & _
            " Where Nvl(费用状态, 0) = 1 And 记录性质 = 1 And 记录状态 = 2" & _
            "       And a.No In (Select Column_Value From Table(f_Str2list([1])))" & _
            "       And Not Exists (Select 1 From 病人预交记录 B Where a.结帐id = b.结帐id And Nvl(b.校对标志, 0) = 0)" & _
            "       And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "检查是否存在退费异常记录", strNos)
    
    If Not rsTemp.EOF Then
        str操作员姓名 = Nvl(rsTemp!操作员姓名)
        CheckIsExistDelErrBill = True
    End If
    
    Exit Function
Errhand:
    If ErrCenter() = 1 Then Resume
End Function

Public Sub zlGetClassMoney(ByRef rsClass As ADODB.Recordset)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取分类汇总金额
    '编制:刘兴洪
    '日期:2011-12-26 13:19:04
    '问题:44944
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo errHandle
    Set rsClass = New ADODB.Recordset
    rsClass.Fields.Append "收费类别", adVarChar, 10, adFldIsNullable
    rsClass.Fields.Append "金额", adDouble, , adFldIsNullable
    rsClass.CursorLocation = adUseClient
    rsClass.LockType = adLockOptimistic
    rsClass.CursorType = adOpenStatic
    rsClass.Open
    With vsBill
        For i = 1 To .Rows - 1
'            If .TextMatrix(i, .ColIndex("选择")) <> 0 Then
                rsClass.Find "收费类别='" & .Cell(flexcpData, i, .ColIndex("类别")) & "'", , adSearchForward, 1
                If rsClass.EOF Then rsClass.AddNew
                rsClass!收费类别 = .Cell(flexcpData, i, .ColIndex("类别"))
                rsClass!金额 = Val(Nvl(rsClass!金额)) + .TextMatrix(i, .ColIndex("实收金额"))
                rsClass.Update
'            End If
        Next
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then Resume
End Sub

Private Function ExecuteInsureInfoUpdate(ByVal lng结帐ID As Long, ByRef str保险金额 As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:更新重收记录的保险信息
    '参数:
    '   str保险金额-"实收合计;进入统筹;全自付;先自付"
    '返回:所有重收记录的保险信息更新成功返回True，否则返回False
    '编制:冉俊明
    '日期:2014-9-16
    '问题:77951
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Dim rsReCharge As ADODB.Recordset, strBXInfo As String, cllReChargePro As Collection
    Dim blnTrans As Boolean
    Dim cur实收合计 As Currency, cur进入统筹 As Currency
    Dim cur全自付 As Currency, cur先自付 As Currency
    Dim cur实收金额 As Currency, cur统筹金额 As Currency, bln保险项目 As Boolean
    
    On Error GoTo Errhand
    str保险金额 = ""
    strSQL = " Select a.Id, a.病人id, a.收费细目id, Nvl(a.付数, 1) * Nvl(a.数次, 0) As 数量, Nvl(a.实收金额, 0) As 实收金额, a.摘要, " & _
            " Nvl(a.保险项目否, 0) As 保险项目否, a.保险大类id, Nvl(a.统筹金额, 0) As 统筹金额, a.保险编码, a.费用类型" & _
            " From 门诊费用记录 A" & _
            " Where a.记录性质 = 11 And a.结帐id = [1]"
    Set rsReCharge = zlDatabase.OpenSQLRecord(strSQL, "获取重收费用记录", lng结帐ID)
    With rsReCharge
        If .RecordCount > 0 Then
            Set cllReChargePro = New Collection
            Do While Not .EOF
                '保险项目否(0/1);保险大类ID;进入统筹金额;保险项目编码;摘要;费用类型
                strBXInfo = gclsInsure.GetItemInsure(Nvl(!病人ID), Nvl(!收费细目ID), Val(Nvl(!实收金额)), True, mCurBillType.intInsure, _
                        Nvl(!摘要) & "||" & Val(Nvl(!数量)))
                If strBXInfo <> "" Then
                    '  Zl_门诊收费记录_Update
                    strSQL = "Zl_门诊收费记录_Update("
                    '  Id_In         In 门诊费用记录.Id%Type,
                    strSQL = strSQL & Nvl(!ID) & ","
                    '  保险大类id_In In 门诊费用记录.保险大类id%Type,
                    strSQL = strSQL & ZVal(Split(strBXInfo, ";")(1)) & ","
                    '  保险项目否_In In 门诊费用记录.保险项目否%Type,
                    strSQL = strSQL & Val(Split(strBXInfo, ";")(0)) & ","
                    '  保险编码_In   In 门诊费用记录.保险编码%Type,
                    strSQL = strSQL & "'" & CStr(Split(strBXInfo, ";")(3)) & "',"
                    '  费用类型_In   In 门诊费用记录.费用类型%Type,
                    strSQL = strSQL & "'" & CStr(Split(strBXInfo, ";")(5)) & "',"
                    '  统筹金额_In   In 门诊费用记录.统筹金额%Type,
                    strSQL = strSQL & Format(Val(Split(strBXInfo, ";")(2)), gstrDec) & ","
                    '  摘要_In       In 门诊费用记录.摘要%Type
                    strSQL = strSQL & "'" & CStr(Split(strBXInfo, ";")(4)) & "')"
                    zlAddArray cllReChargePro, strSQL
                    
                    cur统筹金额 = CCur(Val(Split(strBXInfo, ";")(2)))
                    bln保险项目 = Val(Split(strBXInfo, ";")(0)) = 1
                Else
                    cur统筹金额 = Val(Nvl(!统筹金额))
                    bln保险项目 = Val(Nvl(!保险项目否)) = 1
                End If
                
                '统计保险金额
                cur实收金额 = Val(Nvl(!实收金额))
                If cur统筹金额 = 0 Or Not bln保险项目 Then
                    '以原始金额为准,不管分币处理
                    cur全自付 = cur全自付 + cur实收金额
                Else
                    cur进入统筹 = cur进入统筹 + cur统筹金额
                    '以原始金额为准,不管分币处理
                    cur先自付 = cur先自付 + cur实收金额 - cur统筹金额
                End If
                cur实收合计 = cur实收合计 + CCur(Val(Nvl(!实收金额)))
                rsReCharge.MoveNext
            Loop
            '执行过程
            blnTrans = True
            zlExecuteProcedureArrAy cllReChargePro, Me.Caption, True, True
            blnTrans = False
        End If
    End With
    '保险金额信息
    str保险金额 = cur实收合计 & ";" & cur进入统筹 & ";" & cur全自付 & ";" & cur先自付
    ExecuteInsureInfoUpdate = True
    Exit Function
Errhand:
    If blnTrans Then
        gcnOracle.RollbackTrans
        Call ErrCenter
    Else
        If ErrCenter = 1 Then Resume
    End If
End Function

Private Function SelectMulitBalance(ByVal strNos As String, ByRef strNo As String) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:选择多次结算的一次结算单据
    '入参:strNos-单据号,多个用逗号
    '     strNo -当前输入的单据号,再只一次结算时，直接返回
    '出参:strNo-返回当前选中的单据号
    '返回:选择成功，返回true,否则返回False
    '编制:刘兴洪
    '日期:2016-05-04 17:16:56
    '说明：
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, rsSel As ADODB.Recordset
    Dim strWithTable As String, cllPro As Collection, varData() As Variant
    On Error GoTo errHandle
    
    '先检查本次按病人打印是否只有一次结算的，如果只有一次结算，就直接退出,不用再选择
   If Len(strNos) <= 4000 Then  '大于4000,多半有几次结算
       strSQL = "" & _
       " Select /*+cardinality(b,10)*/ Count(Distinct nvl(C.结算序号,C.结帐ID)) as 次数 " & vbNewLine & _
       " From 门诊费用记录 A,病人预交记录 C, Table(f_Str2list([1])) B" & vbNewLine & _
       " where A.结帐ID=C.结帐ID And Mod(A.记录性质, 10) = 1  And A.记录状态 in (1,3) And A.NO=B.Column_Value "
      Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNos)
      If Nvl(rsTemp!次数, 0) <= 1 Then
            SelectMulitBalance = True
            Exit Function
      End If
   End If

    If Len(strNos) <= 4000 Then
        strSQL = "" & _
        " Select A.NO as 单据号,max(C.名称) as 开单科室, " & _
        "       A.序号,sum(nvl(A.付数,1)*nvl(A.数次,0)) as 数量, " & _
        "       max(decode(A.记录状态,1,A.操作员编号,3,A.操作员编号,NULL)) as 操作员编号,max(decode(A.记录状态,1,A.操作员姓名,3,A.操作员姓名,NULL)) as 操作员姓名, " & _
        "       to_char(max(decode(A.记录状态,1,a.登记时间,3,a.登记时间,NULL )),'yyyy-mm-dd hh24:mi:ss') as 收款时间" & vbNewLine & _
        " From 门诊费用记录 A, Table(f_Str2list([1])) B,部门表 C" & vbNewLine & _
        " where Mod(A.记录性质, 10) = 1 And A.价格父号 is null  And A.NO=B.Column_Value " & _
        " AND A.开单部门id=C.id " & _
        " Group by A.NO,A.序号" & _
        " Having sum(nvl(A.付数,1)*nvl(A.数次,0)) <>0"
        strSQL = "" & _
        "   Select distinct 单据号,开单科室,操作员编号,操作员姓名,收款时间 " & _
        "   From (" & strSQL & ")" & _
        "   Order by 收款时间,单据号"

        
        strSQL = "" & _
        "  Select Rownum as ID,单据号,开单科室,操作员编号,操作员姓名,收款时间 " & _
        "  From (" & strSQL & ") "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNos)
    Else
        
        If zlGetSplitString4000(strNos, cllPro) = False Then Exit Function
        If zlFromCollectBulidSQL(cllPro, strSQL, varData) = False Then Exit Function
        
        strSQL = "With 单据信息 as (" & strSQL & ")" & vbCrLf
        strSQL = strSQL & vbCrLf & _
        " Select A.NO as 单据号,max(C.名称) as 开单科室, " & _
        "       A.序号,sum(nvl(A.付数,1)*nvl(A.数次,0)) as 数量, " & _
        "       max(decode(A.记录状态,1,A.操作员编号,3,A.操作员编号,NULL)) as 操作员编号,max(decode(A.记录状态,1,A.操作员姓名,3,A.操作员姓名,NULL)) as 操作员姓名, " & _
        "       to_char(max(decode(A.记录状态,1,a.登记时间,3,a.登记时间,NULL )),'yyyy-mm-dd hh24:mi:ss') as 收款时间" & vbNewLine & _
        " From 门诊费用记录 A, 单据信息 B,部门表 C" & vbNewLine & _
        " where Mod(A.记录性质, 10) = 1 And A.价格父号 is null  And A.NO=B.NO " & _
        " AND A.开单部门id=C.id " & _
        " Group by A.NO,A.序号" & _
        " Having sum(nvl(A.付数,1)*nvl(A.数次,0)) <>0"
       
        strSQL = "" & _
        "   Select distinct 单据号,开单科室,操作员编号,操作员姓名,收款时间 " & _
        "   From (" & strSQL & ")" & _
        "   Order by 收款时间,单据号"
        
        strSQL = "" & _
        "  Select Rownum as ID,单据号,开单科室,操作员编号,操作员姓名,收款时间 " & _
        "  From (" & strSQL & ") "
        Set rsTemp = zlDatabase.OpenSQLRecordByArray(strSQL, "检查二次结算", varData)
    End If

    If rsTemp Is Nothing Then Exit Function
    If rsTemp.State <> 1 Then Exit Function
    If rsTemp.EOF Then Exit Function
    If rsTemp.RecordCount = 1 Then
        strNo = Nvl(rsSel!单据号)
        SelectMulitBalance = True: Exit Function
    End If
    
    If zlDatabase.zlShowListSelect(Me, glngSys, mlngModule, txtNO, rsTemp, True, "选择指定单据", , rsSel) = False Then Exit Function
    If rsSel Is Nothing Then Exit Function
    If rsSel.State <> 1 Then Exit Function
    If rsSel.EOF Then Exit Function
    strNo = Nvl(rsSel!单据号)
    SelectMulitBalance = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Public Function CheckSelectItemCanDel(ByVal strNos As String) As Boolean
    '功能：判断选择的退费项目是否可以正常退费，主要检查并发，可能有的项目在提出单据出来后又被执行了
    '参数：
    '   strNos - 本次选择的退费单据号
    '返回：
    '   检查通过，返回True；否则，返回False
    '问题号：105429
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String
    Dim i As Long, j As Long, k As Long
    Dim arrNo As Variant
    Dim dbl剩余数量 As Double, dbl本次数量 As Double
    
    On Error GoTo errHandler
    If Left(strNos, 1) = "," Then strNos = Mid(strNos, 2)
    strNos = Replace(strNos, "'", "")
    If GetFeeListData(strNos, rsTemp) = False Then Exit Function
    If rsTemp.EOF Then
        MsgBox "单据:" & strNos & " 中没有可退费的项目，不能退费！", vbExclamation, gstrSysName
        Exit Function
    End If
    
    arrNo = Split(strNos, ",")
    For i = 0 To UBound(arrNo)
        With vsBill
            k = .FindRow(arrNo(i), , .ColIndex("单据号"))
            For j = k To vsBill.Rows - 1
                If .TextMatrix(j, .ColIndex("单据号")) <> arrNo(i) Then Exit For
                If Val(.TextMatrix(j, .ColIndex("选择"))) <> 0 Then
                    rsTemp.Filter = "NO='" & arrNo(i) & "' And 序号=" & .RowData(j)
                    If rsTemp.EOF Then
                        MsgBox "单据:" & arrNo(i) & " 中第 " & (j - k + 1) & " 行项目的剩余未退数量为零，不能退费！" & _
                            "请重新获取费用数据！", vbExclamation, gstrSysName
                        If .Visible And .Enabled Then .Row = j: .SetFocus
                        Exit Function
                    ElseIf Val(Nvl(rsTemp!原始数量)) > 0 Then
                        '负数收费的不检查
                        dbl剩余数量 = Val(Nvl(rsTemp!付数, 1)) * Val(Nvl(rsTemp!数次))
                        dbl本次数量 = Val(.TextMatrix(j, .ColIndex("数量")))
                        If RoundEx(dbl本次数量, 6) > RoundEx(dbl剩余数量, 6) Then
                            MsgBox "单据:" & arrNo(i) & " 中第 " & (j - k + 1) & " 行项目的本次退费数量(" & _
                                FormatEx(dbl本次数量, 5) & ")大于了剩余未退数量(" & FormatEx(dbl剩余数量, 5) & ")，" & _
                                "不能退费！请重新获取费用数据！", vbExclamation, gstrSysName
                            If .Visible And .Enabled Then .Row = j: .SetFocus
                            Exit Function
                        End If
                    End If
                End If
            Next
        End With
    Next
    CheckSelectItemCanDel = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetFeeListData(ByVal strNos As String, ByRef rsFeeList As ADODB.Recordset) As Boolean
    '读取可退费单据内容
    '出参:rsFeeList-返回准退费集
    '返回:获取成功,返回true,否则返回False
    '----------------------------------------------------------------------------------
    '退费时不用考虑后备表,前面的操作已禁用
    '读取准退数,并计算应收金额,实收金额(金额=剩余金额*(准退数/剩余数))
    '读取单据中原始记录的费用ID
    Dim strSQL As String
    Dim strTableNo As String, strSQLIn As String
    Dim strSqlSub As String
    
    On Error GoTo errHandler
    strSqlSub = _
        " Select /*+cardinality(j,10)*/ a.Id, a.记录性质, a.No, a.记录状态, a.序号, a.从属父号, a.价格父号, a.收费细目id," & vbNewLine & _
        "        Nvl(a.付数, 1) As 付数, Nvl(a.数次, 0) As 数次," & vbNewLine & _
        "        Nvl(a.应收金额, 0) As 应收金额, Nvl(a.实收金额, 0) As 实收金额, Nvl(a.结帐金额, 0) As 结帐金额," & vbNewLine & _
        "        Nvl(a.付数, 1) * a.数次 As 数量, Nvl(标准单价, 0) As 标准单价," & vbNewLine & _
                 IIf(gbln药房单位, "Nvl(b." & gstr药房包装 & ",1)", "1") & " As 换算系数, " & vbNewLine & _
                 IIf(gbln药房单位, "Decode(B.药品ID,NULL,A.计算单位,B." & gstr药房单位 & ")", "A.计算单位 ") & " As 计算单位," & vbNewLine & _
        "        a.开单部门id, a.执行部门id, a.医嘱序号, " & vbNewLine & _
        "        a.执行状态,a.费用类型, a.费用状态, a.附加标志,a.费别, a.收费类别, a.操作员姓名, a.登记时间, a.结帐id," & vbNewLine & _
        "        b.药品id" & vbNewLine & _
        " From 门诊费用记录 A, 药品规格 B, Table(f_Str2list([1])) J" & vbNewLine & _
        " Where Mod(a.记录性质, 10) = 1 And a.No = j.Column_Value And a.记录状态 <> 0" & _
        "       And a.收费细目id = b.药品id(+)"
    '求准退费(卫材,药品,其他治疗类)
    strTableNo = _
        " With 门诊费用 As (" & strSqlSub & ")," & vbNewLine & _
        "      准退数 As (Select /*+cardinality(j,10)*/ A.费用ID," & _
        "                        Sum(Nvl(A.付数,1)*A.实际数量" & IIf(gbln药房单位, "/Nvl(B." & gstr药房包装 & ",1)", "") & ") as 准退数量" & vbNewLine & _
        "                 From 药品收发记录 A,药品规格 B, Table(f_Str2list([1])) J" & vbNewLine & _
        "                 Where A.药品ID=B.药品ID(+) And Mod(A.记录状态,3)=1  " & vbNewLine & _
        "                       And (A.单据 =8 or a.单据=24) And A.审核人 is NULL And A.NO =J.Column_Value" & vbNewLine & _
        "                 Group by A.费用ID"

    '求诊疗相关的准退数
    '*在医嘱执行计价中存在数据时,则按医嘱执行计价中取数
    '*病人医嘱发送.执行状态=1（完成执行）时，准退数为0，不再根据医嘱执行计价来统计准退数,112447
    strTableNo = strTableNo & vbNewLine & _
        "   Union ALL " & vbNewLine & _
        "   Select Max(ID) As 费用ID, Nvl(Sum(数量), 0) As 准退数" & vbNewLine & _
        "   From(Select a.Id, a.医嘱序号 As 医嘱id, a.收费细目id, Decode(b.执行状态, 1, 0, Decode(c.执行状态, 0, 1, 0)) * c.数量 As 数量" & vbNewLine & _
        "        From (" & strSqlSub & ") A, 病人医嘱发送 B, 医嘱执行计价 C, 病人医嘱记录 M" & vbNewLine & _
        "        Where a.医嘱序号 = b.医嘱id And a.No = b.No And b.医嘱id = c.医嘱id And b.医嘱ID = m.id" & vbNewLine & _
        "              And b.发送号 = c.发送号 And a.收费细目id = c.收费细目id + 0 And a.价格父号 Is Null" & vbNewLine & _
        "              And a.记录性质 = 1 And a.记录状态 in (1, 3) And Instr(',5,6,7,', ',' || a.收费类别 || ',') = 0" & vbNewLine & _
        "              And Not Exists(Select 1 From 材料特性 C Where a.收费细目id = c.材料id And c.跟踪在用 = 1)" & vbNewLine & _
        "              And Instr(',C,D,F,G,K,',','||m.诊疗类别||',')=0 And b.记录性质 = 1" & vbNewLine & _
        "        )" & vbNewLine & _
        "   Group By 医嘱ID, 收费细目ID" & vbNewLine & _
        "   Having Max(ID) <> 0" & vbNewLine & _
        "  )"
    
    '整张单据汇总结果(明细到收费细目)
    '执行状态应该在原始记录上判断(部分退药且部份退费的记录)
    '当退两次以上时"记录状态,序号"重复,AVG有问题,所以要用"执行状态"
    '   *无医嘱执行计价的部分退费无法判断准退数量，不允许退费
    strSQLIn = "" & _
        " Select NO, Nvl(价格父号, 序号) As 序号" & vbNewLine & _
        " From 门诊费用" & vbNewLine & _
        " Where 记录性质 = 1 And 记录状态 In (1, 3) And Nvl(执行状态, 0) <> 1" & vbNewLine & _
        " Minus" & vbNewLine & _
        " Select NO, Nvl(价格父号, 序号) As 序号" & vbNewLine & _
        " From 门诊费用 A1" & vbNewLine & _
        " Where A1.记录性质 = 1 And A1.记录状态 In (1, 3) And Nvl(A1.执行状态, 0) = 2" & vbNewLine & _
        "       And Not Exists(Select 1" & vbNewLine & _
        "                      From 病人医嘱发送 B, 医嘱执行计价 C" & vbNewLine & _
        "                      Where b.医嘱id = A1.医嘱序号 And b.No = A1.No" & vbNewLine & _
        "                            And b.医嘱id = c.医嘱id And b.发送号 = c.发送号" & vbNewLine & _
        "                            And c.收费细目id + 0 = A1.收费细目id And b.记录性质 = 1)" & vbNewLine & _
        "       And Instr('5,6,7', A1.收费类别) = 0" & vbNewLine & _
        "       And Not Exists(Select 1 From 材料特性 Where 材料id = A1.收费细目id And Nvl(跟踪在用, 0) = 1)"
    
    strSQL = _
        " Select A.NO,A.记录状态,A.记录性质,A.执行状态,Nvl(A.价格父号,A.序号) as 序号,A.从属父号," & _
        "       A.费别,C.编码 as 类别码,C.名称 as 类别名,A.收费细目ID,B.编码,B.名称,B.规格," & _
        "       Max(Nvl(A.费用类型,B.费用类型)) 费用类型," & _
        "       A.计算单位,Max(A.医嘱序号) as 医嘱序号, " & _
        "       Avg(Nvl(A.付数,1)) as 付数,Avg(A.数次/A.换算系数) as 数次," & _
        "       Sum(A.标准单价*A.换算系数) as 单价," & _
        "       Sum(A.应收金额) as 应收金额,Sum(A.实收金额) as 实收金额, " & _
        "       D.名称 as 执行科室,A.执行部门ID,E.名称 as 开单科室" & _
        " From  门诊费用 A,收费项目目录 B,收费项目类别 C,部门表 D,部门表 E" & _
        " Where A.收费细目ID=B.ID And C.编码=A.收费类别" & _
        "       And A.执行部门ID=D.ID(+) And A.开单部门ID=E.ID(+)" & _
        "       And (A.NO,Nvl(A.价格父号,A.序号)) IN( " & strSQLIn & ")  " & _
        "       And A.NO IN( Select NO From 门诊费用 where  记录性质=1 and 记录状态 in (1,3) )" & _
        " Group by A.NO,A.记录性质,A.记录状态,A.执行状态,Nvl(A.价格父号,A.序号),A.费别,A.从属父号," & _
        "       C.编码,C.名称,A.收费细目ID,B.编码,B.名称,B.规格,A.计算单位," & _
        "       D.名称,A.执行部门ID,E.名称,A.药品ID,a.结帐ID "

        '最后计算结果
        '当"准退数量=原始数量"时,付数才保留
        '排开已经全部退费的行(执行状态=0的一种可能)
        '有剩余数量无准退数量的有两种情况：
            '1.无对应的收发记录(如普通费用或不跟踪在用的卫材),这时应用剩余数量
            '2.收发记录中已全部发放,即已全部执行,SQL已排除这种记录
    strSQL = strTableNo & vbCrLf & _
        " Select A.NO,A.序号,A.从属父号,A.费别,A.类别码,A.类别名,A.收费细目ID,A.编码,A.名称,A.规格," & _
        "       Max(A.费用类型) As 费用类型,A.计算单位, Max(A.医嘱序号) as 医嘱序号," & _
        "       Decode(Sign(Nvl(C.准退数量,Sum(A.付数*A.数次))-B.原始数量),0,avg(A.付数),1) as 准退付数," & _
        "       Decode(Sign(Nvl(C.准退数量,Sum(A.付数*A.数次))-B.原始数量),0,Sum(A.数次),Nvl(C.准退数量,Sum(A.付数*A.数次))) as 准退数次," & _
        "       Nvl(C.准退数量,Sum(A.付数*A.数次)) as 准退数量,Sum(A.付数*A.数次) as 剩余数量," & _
        "       A.单价,Sum(A.应收金额) as 剩余应收,Sum(A.实收金额) as 剩余实收,max(q1.记录标志) as 记录标志," & _
        "       A.执行科室,A.执行部门ID,A.开单科室,B.操作员姓名,B.登记时间,B.结帐ID,Max(M.医嘱内容) as 医嘱内容,b.原始数量" & _
        " From (" & strSQL & ") A, 准退数 C,病人医嘱记录 M," & _
        "          ( Select  ID, NO,序号, 收费细目ID,Nvl( 数量,0)/NVL(换算系数,1) as 原始数量,操作员姓名,登记时间,结帐ID" & _
        "            From 门诊费用   " & _
        "            Where  记录状态 IN(1,3) and 记录性质=1 And Nvl( 附加标志,0)<>9 And  价格父号 is NULL )B, " & _
        "            ( Select NO,Max(记录状态) as 记录标志 From 门诊费用  Where 记录状态 in (1,3) Group by NO) Q1" & _
        " Where A.NO=B.NO And A.序号=B.序号 And A.收费细目ID=B.收费细目ID+0  And B.ID=C.费用ID(+)" & _
        "            and A.医嘱序号=M.ID(+) and A.NO=q1.NO(+) " & _
        " Group by A.NO,A.序号,A.从属父号,A.费别,A.类别码,A.类别名,A.收费细目ID,A.编码,A.名称,A.规格," & _
        "       A.计算单位,A.单价,B.原始数量,C.准退数量,A.执行科室,A.执行部门ID,A.开单科室,B.操作员姓名,B.登记时间,B.结帐ID" & _
        " Having Sum(A.付数*A.数次)<>0"

    strSQL = _
        " Select A.NO,A.序号,A.从属父号,A.费别,A.类别码,A.类别名,A.编码,Nvl(B.名称,A.名称) as 名称,E1.名称 as 商品名," & _
        "       A.规格,A.费用类型,A.计算单位,A.收费细目ID,A.准退付数 as 付数,A.准退数次 as 数次,A.单价, A.医嘱序号 ," & _
        "       A.剩余应收*(A.准退数量/A.剩余数量) as 应收金额," & _
        "       A.剩余实收*(A.准退数量/A.剩余数量) as 实收金额," & _
        "       A.执行科室,A.执行部门ID,A.开单科室,A.操作员姓名,A.登记时间,A.结帐ID,A.医嘱内容,A.记录标志, " & _
        "       A.原始数量,A.准退数量,A.剩余数量" & _
        " From (" & strSQL & ") A,收费项目别名 B,收费项目别名 E1" & _
        " Where     A.收费细目ID=B.收费细目ID(+) And B.码类(+)=1 And B.性质(+)=" & IIf(gTy_System_Para.byt药品名称显示 = 1, 3, 1) & _
        "       And A.收费细目ID=E1.收费细目ID(+) And E1.码类(+)=1 And E1.性质(+)=3 " & _
        " Order by A.NO,A.序号"

    Set rsFeeList = zlDatabase.OpenSQLRecord(strSQL, "读取可退费项目", strNos)
    GetFeeListData = True
    Exit Function
errHandler:
    If ErrCenter() = 1 Then
       Resume
    End If
    Call SaveErrLog
End Function

Public Function GetDelXMLExpend() As String
    '获取传入三方卡退费接口zlRetuenCheck中strXMLExpend参数值
    If mbytMode = EM_MULTI_退费 Then
        GetDelXMLExpend = ZlGetDelXMLExpendByGrid(Me.vsBill)
    ElseIf mbytMode = EM_MULTI_异常重退 Then
        GetDelXMLExpend = ZlGetDelXMLExpend(mlng结算序号, True)
    End If
End Function
