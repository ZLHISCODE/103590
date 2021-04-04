VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.9#0"; "zlIDKind.ocx"
Begin VB.Form frmPatiBooks 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "退病历费"
   ClientHeight    =   6330
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8910
   Icon            =   "frmPatiBooks.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6330
   ScaleWidth      =   8910
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Height          =   2235
      Left            =   0
      TabIndex        =   25
      Top             =   4065
      Width           =   7575
      Begin VSFlex8Ctl.VSFlexGrid vsPay 
         Height          =   1740
         Left            =   120
         TabIndex        =   26
         ToolTipText     =   "允许退现的情况下，不选择退费方式将以""现金""方式进行退费"
         Top             =   465
         Width           =   7425
         _cx             =   13097
         _cy             =   3069
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
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPatiBooks.frx":6852
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
         AllowUserFreezing=   2
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.PictureBox picPay 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         ForeColor       =   &H80000008&
         Height          =   1725
         Left            =   5130
         ScaleHeight     =   1695
         ScaleWidth      =   2250
         TabIndex        =   32
         Top             =   480
         Width           =   2280
         Begin VB.TextBox txt应退1 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            IMEMode         =   3  'DISABLE
            Left            =   855
            MaxLength       =   10
            TabIndex        =   35
            Top             =   90
            Width           =   1290
         End
         Begin VB.TextBox txt缴款 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            IMEMode         =   3  'DISABLE
            Left            =   855
            MaxLength       =   12
            TabIndex        =   34
            Top             =   615
            Width           =   1290
         End
         Begin VB.TextBox txt找补 
            Alignment       =   1  'Right Justify
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   14.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   450
            IMEMode         =   3  'DISABLE
            Left            =   855
            Locked          =   -1  'True
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   1155
            Width           =   1290
         End
         Begin VB.Label lbl冲预交 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "应退"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   120
            TabIndex        =   38
            Top             =   195
            Width           =   660
         End
         Begin VB.Label lblPayType 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "退款"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   2
            Left            =   120
            TabIndex        =   37
            Top             =   705
            Width           =   660
         End
         Begin VB.Label lbl找补 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "找补"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   15.75
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   135
            TabIndex        =   36
            Top             =   1230
            Width           =   660
         End
      End
      Begin VB.TextBox txt未退 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   450
         IMEMode         =   3  'DISABLE
         Left            =   6255
         MaxLength       =   10
         TabIndex        =   30
         Top             =   -15
         Width           =   1290
      End
      Begin VB.TextBox txt应退 
         Alignment       =   1  'Right Justify
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   450
         IMEMode         =   3  'DISABLE
         Left            =   4185
         MaxLength       =   10
         TabIndex        =   28
         Top             =   0
         Width           =   1290
      End
      Begin VB.Label lbl未退 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "未退"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   5565
         TabIndex        =   31
         Top             =   75
         Width           =   660
      End
      Begin VB.Label lbl应退 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "应退"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   15.75
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   3480
         TabIndex        =   29
         Top             =   75
         Width           =   660
      End
      Begin VB.Label Label4 
         Caption         =   "本次退费情况"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   14.25
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   120
         TabIndex        =   27
         Top             =   135
         Width           =   2145
      End
   End
   Begin VB.Frame fraSplit3 
      Height          =   6435
      Left            =   7620
      TabIndex        =   24
      Top             =   -90
      Width           =   45
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   45
      Left            =   0
      TabIndex        =   23
      Top             =   0
      Width           =   7620
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消(&C)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   7710
      TabIndex        =   11
      Top             =   705
      Width           =   1095
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   10.5
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   350
      Left            =   7710
      TabIndex        =   10
      Top             =   210
      Width           =   1095
   End
   Begin VB.Frame fraSplit1 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   45
      Left            =   0
      TabIndex        =   15
      Top             =   1410
      Width           =   7620
   End
   Begin VB.Frame fraSplit2 
      BackColor       =   &H00C0C0C0&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   45
      Left            =   0
      TabIndex        =   12
      Top             =   3990
      Width           =   7620
   End
   Begin VB.Frame fraGroup 
      BorderStyle     =   0  'None
      Height          =   2505
      Left            =   0
      TabIndex        =   14
      Top             =   1470
      Width           =   7620
      Begin VSFlex8Ctl.VSFlexGrid vsBooks 
         Height          =   2385
         Left            =   120
         TabIndex        =   9
         Top             =   60
         Width           =   7440
         _cx             =   13123
         _cy             =   4207
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
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   0
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   1
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPatiBooks.frx":6936
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
         AllowUserFreezing=   2
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.Frame fraPatiCard 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1395
      Left            =   0
      TabIndex        =   13
      Top             =   0
      Width           =   7575
      Begin VB.TextBox txt手机 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         IMEMode         =   3  'DISABLE
         Left            =   4395
         MaxLength       =   18
         TabIndex        =   8
         Tag             =   "门诊号"
         Top             =   975
         Width           =   1290
      End
      Begin VB.TextBox txtPatient 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1285
         TabIndex        =   0
         Tag             =   "姓名"
         Top             =   120
         Width           =   2205
      End
      Begin VB.TextBox txtNation 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   6300
         TabIndex        =   6
         Top             =   555
         Width           =   1170
      End
      Begin VB.TextBox txt身份证号 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         IMEMode         =   3  'DISABLE
         Left            =   1290
         MaxLength       =   18
         TabIndex        =   7
         Tag             =   "身份证号"
         Top             =   975
         Width           =   2205
      End
      Begin VB.TextBox txt门诊号 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         IMEMode         =   3  'DISABLE
         Left            =   6315
         MaxLength       =   18
         TabIndex        =   2
         Tag             =   "门诊号"
         Top             =   120
         Width           =   1170
      End
      Begin VB.TextBox txt年龄 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   4395
         TabIndex        =   5
         Top             =   555
         Width           =   990
      End
      Begin VB.TextBox txt性别 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   4395
         TabIndex        =   1
         Top             =   120
         Width           =   1005
      End
      Begin MSMask.MaskEdBox txt出生时间 
         Height          =   345
         Left            =   2505
         TabIndex        =   4
         Tag             =   "出生时间"
         Top             =   555
         Width           =   645
         _ExtentX        =   1138
         _ExtentY        =   609
         _Version        =   393216
         BackColor       =   -2147483633
         Enabled         =   0   'False
         MaxLength       =   5
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "hh:mm"
         Mask            =   "##:##"
         PromptChar      =   "_"
      End
      Begin MSMask.MaskEdBox txt出生日期 
         Height          =   345
         Left            =   1305
         TabIndex        =   3
         Tag             =   "出生日期"
         Top             =   555
         Width           =   1185
         _ExtentX        =   2090
         _ExtentY        =   609
         _Version        =   393216
         BackColor       =   -2147483633
         AutoTab         =   -1  'True
         Enabled         =   0   'False
         MaxLength       =   10
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Format          =   "YYYY-MM-DD"
         Mask            =   "####-##-##"
         PromptChar      =   "_"
      End
      Begin zlIDKind.IDKindNew IDKind 
         Height          =   360
         Left            =   670
         TabIndex        =   39
         Top             =   120
         Width           =   600
         _ExtentX        =   1058
         _ExtentY        =   635
         Appearance      =   2
         IDKindStr       =   $"frmPatiBooks.frx":6A14
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
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         BackColor       =   -2147483633
      End
      Begin VB.Label lbl手机号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "手机号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3735
         TabIndex        =   40
         Top             =   1035
         Width           =   630
      End
      Begin VB.Label lbl民族 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "民族"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5865
         TabIndex        =   22
         Top             =   600
         Width           =   420
      End
      Begin VB.Label lbl身份证号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "身份证号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   420
         TabIndex        =   21
         Top             =   1050
         Width           =   840
      End
      Begin VB.Label lbl门诊号 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "门诊号"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   5660
         TabIndex        =   20
         Top             =   180
         Width           =   630
      End
      Begin VB.Label lbl年龄 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "年龄"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3930
         TabIndex        =   19
         Top             =   615
         Width           =   420
      End
      Begin VB.Label lbl出生日期 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "出生日期"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   420
         TabIndex        =   18
         Top             =   615
         Width           =   840
      End
      Begin VB.Label lbl性别 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "性别"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   3915
         TabIndex        =   17
         Top             =   180
         Width           =   420
      End
      Begin VB.Label lbl姓名 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   165
         TabIndex        =   16
         Top             =   210
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmPatiBooks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mbytFunc As Long    '1-查看;2-编辑
Private mlngModule As Long
Private mobjKeyboard As Object
Private mrsBooks As ADODB.Recordset
Private mobjPayCards As Cards
Private mobjPayCard As Card '病历费原来的支付方式
Private mobjDelObjects  As clsCardObjects
Private mobjDelObject  As clsCardObject
Private mblnUnLoad As Boolean
Private mblnNotClick As Boolean
Private mstr结算方式 As String
Private mlng医疗卡长度  As Long
Private Type T_Pati
    病人ID As Long
    姓名 As String
    性别 As String
    年龄 As String
    卡号 As String
    密码 As String
    民族 As String
    出生日期 As String
    门诊号 As Long
    身份证号 As String
    手机号 As String
    病人类型 As String
End Type
Private mPati As T_Pati

Private Type M_PayInfo
    bln全退 As Boolean
    bln退现 As Boolean
    bln医保 As Boolean
End Type
Private mPayInfo As M_PayInfo

Private Const C_BookInfoColumHeader = "单据号,905,1;场合,605,4;病历费,905,7;记账,605,4;发生时间,1205,4;操作员,905,1;备注,605,1" '格式:"列名","对齐","列宽"(其中对齐取值为:1-左对齐 4-居中 7-右对齐)
Private Const C_PayInfoColumHeader = "选择,605,4;支付方式,1205,1;支付金额,1005,7;卡号,605,1;交易流水号,1505,1;交易说明,1205,1;备注,605,1" '格式:"列名","对齐","列宽"(其中对齐取值为:1-左对齐 4-居中 7-右对齐)
Private Const C_COLOR_背景 = &H80000005
Private Const C_COLOR_白色 = &H80000005
Private Const C_COLOR_蓝色 = &H8000000D

Public Sub ShowMe(ByRef frmMain As Object, ByVal bytFunc As Byte, ByVal lngModul As Long, ByVal lng病人ID As Long, ByVal str卡号 As String)
'功能:显示主窗体
'参数: FrmMain-主窗体
'      bytFunc=1-查看,2-编辑
'      lng病人ID-查看时传人 （bytFunc=1时传入,bytFunc=2时刷卡获取）
'      lngModul 模块号

    mbytFunc = bytFunc
    mlngModule = lngModul
    mPati.病人ID = lng病人ID
    mPati.卡号 = str卡号
    Me.Show 1, frmMain
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdDel_Click()
    '三方卡支持退现的情况下，删除三方卡退款方式，增加现金退费方式，如果有则在现金方式上增加退款金额
    Dim dblMoney As Double, i As Integer, blnFind As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo Errhand
    strSQL = "Select 名称 From 结算方式 Where 性质=1 Order By 缺省标志 Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTemp.RecordCount = 0 Then Exit Sub
    With vsPay
        dblMoney = CDbl(.TextMatrix(.RowSel, .ColIndex("支付金额")))
        .RemoveItem
        
        For i = 1 To .Rows - 1
            If .Cell(flexcpData, i, .ColIndex("支付方式")) = 1 Then
                blnFind = True
                dblMoney = dblMoney + CDbl(.TextMatrix(i, .ColIndex("支付方式")))
                .TextMatrix(i, .ColIndex("支付金额")) = Format(dblMoney, "0.00")
            End If
            If blnFind Then Exit For
        Next
        
        If Not blnFind Then
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, .ColIndex("支付方式")) = Nvl(rsTemp!名称)
            .Cell(flexcpData, .Rows - 1, .ColIndex("支付方式")) = 1
            .TextMatrix(.Rows - 1, .ColIndex("支付金额")) = Format(dblMoney, "0.00")
            .RowData(.Rows - 1) = 1
        End If
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Sub

Private Sub cmdOK_Click()
    If IsCheckCancelValied = False Then Exit Sub
    If IsNoCanc = False Then Exit Sub
    If SaveData = False Then Exit Sub
    Call ClearPatiInfo
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0: Exit Sub
End Sub

Private Sub Form_Load()
    '初始化
    Call InitFace
    Call InitVsFlex
    Call Init退费方式
    Call LoadPatiBooks
    Call IDKind.zlInit(Me, glngSys, mlngModule, gcnOracle, gstrDBUser, gobjSquare.objSquareCard, , txtPatient)

    If mPati.病人ID <> 0 Then Call LoadPati(mPati.病人ID)
    
End Sub

Private Sub Form_Resize()
    Dim lngW As Long
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set mobjKeyboard = Nothing
    Set mrsBooks = Nothing
    Set mobjPayCards = Nothing
    Set mobjPayCard = Nothing
    Set mobjDelObject = Nothing
End Sub

Private Sub InitFace()
    Dim objCtl As Object
    
    For Each objCtl In Me.Controls
        Select Case UCase(TypeName(objCtl))
            Case "TEXTBOX", "MASKEDBOX"
                objCtl.Enabled = False
        End Select
    Next
    txtPatient.Enabled = True
    txt应退.Enabled = True: txt应退.Locked = True
    txt未退.Enabled = True: txt未退.Locked = True
End Sub

Private Sub txtPatient_Change()
    If Trim(txtPatient.Text) = "" Then
        Call ClearPatiInfo
    End If
End Sub

Private Sub LoadPati(ByVal lng病人ID As Long)
    '功能:加载病人信息
    Dim strSQL As String, rsTemp As ADODB.Recordset
    strSQL = "Select A.姓名,A.性别,A.民族,A.出生日期,A.年龄,A.身份证号,A.门诊号,A.手机号," & vbNewLine & _
    "         Nvl(Nvl(A.病人类型,B.病人类型),Decode(Nvl(A.险类,B.险类),Null,'普通病人','医保病人')) 病人类型 " & vbNewLine & _
    " From 病人信息 A,病案主页 B " & vbNewLine & _
    " Where A.病人ID=B.病人ID(+) And A.主页ID=B.主页ID(+) And A.停用时间 is NULL And A.病人ID=[1] "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "病人信息", lng病人ID)
    With rsTemp
        If .RecordCount = 0 Then Exit Sub
        txtPatient.Text = Nvl(!姓名)
        txt性别.Text = Nvl(!性别)
        txtNation.Text = Nvl(!民族)
        txt年龄.Text = Nvl(!年龄)
        txt出生日期.Text = Format(Nvl(!出生日期), "YYYY-MM-DD")
        txt出生时间.Text = Format(Nvl(!出生日期), "MM:SS")
        txt门诊号.Text = Nvl(!门诊号)
        txt身份证号.Text = Nvl(!身份证号)
        txt手机.Text = Nvl(!手机号)
        
        mPati.病人ID = lng病人ID
        mPati.姓名 = Nvl(!姓名)
        mPati.密码 = ""
        mPati.年龄 = Nvl(!年龄)
        mPati.性别 = Nvl(!性别)
        mPati.出生日期 = Format(Nvl(!出生日期), "YYYY-MM-DD MM:SS")
        mPati.民族 = Nvl(!民族)
        mPati.门诊号 = Nvl(!门诊号)
        mPati.身份证号 = Nvl(!身份证号)
        mPati.手机号 = Nvl(!手机号)
        mPati.病人类型 = Nvl(!病人类型)
    End With
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Sub

Private Sub IDKind_Click(objCard As zlIDKind.Card)
    Dim lng卡类别ID As Long, strOutCardNo As String, strExpand
    Dim strOutPatiInforXml As String

    lng卡类别ID = IDKind.GetCurCard.接口序号
    If lng卡类别ID <= 0 Then Exit Sub
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
    'Call InitInterFacel(Me, mlngModule, lng卡类别ID, False, mobjCardObject)
    strExpand = lng卡类别ID
    If gobjSquare.objSquareCard.zlReadCard(Me, mlngModule, lng卡类别ID, True, strExpand, strOutCardNo, strOutPatiInforXml) = False Then Exit Sub
    txtPatient.Text = strOutCardNo
    If txtPatient.Text <> "" Then Call txtPatient_KeyPress(vbKeyReturn)
    Exit Sub
 
End Sub

 
Private Sub IDKind_ItemClick(Index As Integer, objCard As zlIDKind.Card)
     '短名|全名|读卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|
    '是否存在帐户(1-存在帐户;0-不存在帐户)|卡号密文(第几位至第几位加密,空为不加密)
    Set gobjSquare.objCurCard = objCard
    mlng医疗卡长度 = objCard.卡号长度
    '105667:李南春，2017/5/23，卡号加密导致第一个汉字拼音不能触发输入法
    txtPatient.PasswordChar = ""
    '需要清除信息,避免刷卡后,再切换,造成密文显示失去意义
    If txtPatient.Text <> "" And Not mblnNotClick Then txtPatient.Text = ""
    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0
End Sub

Private Sub IDKind_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    Dim lngPreIDKind As Long, intIndex As Integer
    Dim dtDate As Date
    Dim blnNew As Boolean

    If txtPatient.Locked Then Exit Sub  'Or Not Me.ActiveControl Is txtPatient Or txtPatient.Text <> ""
    mblnNotClick = True
    intIndex = IDKind.GetKindIndex(objCard.名称)
    lngPreIDKind = IDKind.IDKind
    If intIndex > 0 Then IDKind.IDKind = intIndex

    txtPatient.Text = objPatiInfor.卡号
    Call txtPatient_KeyPress(vbKeyReturn)
    IDKind.IDKind = lngPreIDKind
    mblnNotClick = False
End Sub

Private Sub InitVsFlex()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:初始化VSGrid控件
    '编制:56599
    '日期:2012-12-05 11:39:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strTmp As String
    Dim i As Integer
    On Error GoTo Errhand
    
    With vsBooks
        .Redraw = False
        .Cols = UBound(Split(C_BookInfoColumHeader, ";")) + 1
        For i = 0 To UBound(Split(C_BookInfoColumHeader, ";"))
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColKey(i) = Split(Split(C_BookInfoColumHeader, ";")(i), ",")(0)
            .TextMatrix(0, i) = Split(Split(C_BookInfoColumHeader, ";")(i), ",")(0)
            .ColWidth(i) = Split(Split(C_BookInfoColumHeader, ";")(i), ",")(1)
            .ColAlignment(i) = Split(Split(C_BookInfoColumHeader, ";")(i), ",")(2)
           ' .ColAlignmentFixed(i) = 4
        Next
      '   .ColHidden(getColNum("记录状态")) = True
        .RowHeight(0) = 320
        .ExtendLastCol = True
        
        .ForeColorSel = C_COLOR_蓝色
        .Redraw = True
    End With
    
    With vsPay
        .Redraw = False
        .Cols = UBound(Split(C_PayInfoColumHeader, ";")) + 1
        For i = 0 To UBound(Split(C_PayInfoColumHeader, ";"))
            .FixedAlignment(i) = flexAlignCenterCenter
            .ColKey(i) = Split(Split(C_PayInfoColumHeader, ";")(i), ",")(0)
            .TextMatrix(0, i) = Split(Split(C_PayInfoColumHeader, ";")(i), ",")(0)
            .ColWidth(i) = Split(Split(C_PayInfoColumHeader, ";")(i), ",")(1)
            .ColAlignment(i) = Split(Split(C_PayInfoColumHeader, ";")(i), ",")(2)
           ' .ColAlignmentFixed(i) = 4
        Next
        '增加一个隐藏的卡类别ID
        .Cols = .Cols + 1
        .ColKey(.Cols - 1) = "卡类别ID"
        .TextMatrix(0, .Cols - 1) = "卡类别ID"
        .ColHidden(.Cols - 1) = True
      '   .ColHidden(getColNum("记录状态")) = True
        .RowHeight(0) = 320
        .ExtendLastCol = True
        .ColDataType(0) = flexDTBoolean
        .Editable = flexEDKbd
        .ForeColorSel = C_COLOR_蓝色
        .Redraw = True
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub LoadPatiBooks()
'功能:加载病人家属信息
    '病人家属
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Long
    
    On Error GoTo errH
    
    If mPati.病人ID = 0 Then Exit Sub
    strSQL = strSQL & _
    "   Select A.ID, '发卡' as 场合,A.no as 单据号,A.实收金额,A.记帐费用,A.发生时间,A.操作员姓名,A.摘要,A.序号,A.记录状态, " & vbNewLine & _
    "          A.结帐ID,Decode(B.记录性质 , 11,'冲预交',B.结算方式) as 结算方式,Sum(B.冲预交) as 金额, B.卡类别ID,B.卡号,B.交易说明,B.结算卡序号,B.交易流水号, " & vbNewLine & _
    "          Decode(B.记录性质 , 11,0,D.性质) as 性质,Nvl(C.是否退现,1) as 是否退现,Nvl(C.是否全退,0) as 是否全退 " & vbNewLine & _
    "   From 住院费用记录 A,病人预交记录 B, 医疗卡类别 C,结算方式 D " & vbNewLine & _
    "   Where A.NO = B.NO(+) and A.记录性质  = B.记录性质(+) And B.卡类别ID = C.ID(+) And B.结算方式 = D.名称(+) And A.记录性质=5 And A.记录状态 = 1 and A.附加标志=8 And A.病人ID=[1]" & vbNewLine & _
    "   Group by A.ID,A.no,A.实收金额,A.记帐费用,A.发生时间,A.操作员姓名,A.摘要,A.序号,A.记录状态,A.结帐ID,Decode(B.记录性质 , 11,'冲预交',B.结算方式), " & vbNewLine & _
    "            B.卡类别ID , B.卡号, B.交易说明, B.结算卡序号, B.交易流水号, Decode(B.记录性质, 11, 0, D.性质), Nvl(C.是否退现, 1), Nvl(C.是否全退, 0) " & vbNewLine & _
    "   Order by A.ID Desc,场合"
'    "   Union All" & vbNewLine & _
'    "   Select A.ID, '挂号' as 场合,A.no as 单据号,A.实收金额,A.记帐费用,A.发生时间,A.操作员姓名,A.摘要,A.序号,A.记录状态, " & _
'    "          B.结帐ID,Decode(B.记录性质 , 11,'冲预交',B.结算方式) as 结算方式,B.冲预交 as 金额, B.卡类别ID,B.卡号,B.交易说明,B.结算序号,B.结算卡序号,B.交易流水号, " & _
'    "          Decode(B.记录性质 , 11,0,D.性质) as 性质,Nvl(C.是否退现,1) as 是否退现,Nvl(C.是否全退,0) as 是否全退 " & vbNewLine & _
'    "   From 门诊费用记录 A,病人预交记录 B, 医疗卡类别 C,结算方式 D " & vbNewLine & _
'    "   Where A.结帐ID = B.结帐ID(+) And B.卡类别ID = C.ID(+) And B.结算方式 = D.名称 And A.记录性质=4 And A.记录状态 = 1 And A.附加标志=1 And A.病人ID=[1]" & vbNewLine & _
'    ") Order by ID Desc,场合"
    Set mrsBooks = zlDatabase.OpenSQLRecord(strSQL, "病人病历费", mPati.病人ID)

    With vsBooks
       .Rows = 1 '缺省显示一行
        If mrsBooks Is Nothing Then Exit Sub
        Do While Not mrsBooks.EOF
            '如果单据号和场合相同，是多种支付方式结算，不重复登记
            If Nvl(mrsBooks!单据号) <> .TextMatrix(i, .ColIndex("单据号")) Or Nvl(mrsBooks!场合) <> .TextMatrix(i, .ColIndex("场合")) Then
                i = i + 1: .Rows = i + 1
                .RowData(i) = mrsBooks!id
                .TextMatrix(i, .ColIndex("单据号")) = mrsBooks!单据号 & ""
                .TextMatrix(i, .ColIndex("场合")) = mrsBooks!场合 & ""
                .TextMatrix(i, .ColIndex("病历费")) = Format(mrsBooks!实收金额 & "", "0.00")
                .TextMatrix(i, .ColIndex("发生时间")) = Format(mrsBooks!发生时间 & "", "YYYY-MM-DD")
                .TextMatrix(i, .ColIndex("操作员")) = mrsBooks!操作员姓名 & ""
                .TextMatrix(i, .ColIndex("备注")) = mrsBooks!摘要 & ""
                .TextMatrix(i, .ColIndex("记账")) = IIf(Val(mrsBooks!记帐费用 & "") = 1, "√", "")
                '1-实收;2-记账;3-划价
                .Cell(flexcpData, i, .ColIndex("记账")) = IIf(Val(mrsBooks!记帐费用 & "") = 1, 2, IIf(Val(mrsBooks!记录状态 & "") = 1, 1, 3))
            End If
            mrsBooks.MoveNext
        Loop
        If .Rows > 1 Then .Row = 1
    End With
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ClearPatiInfo()
    mPati.病人ID = 0
    mPati.卡号 = ""
    mPati.密码 = ""
    mPati.年龄 = ""
    mPati.性别 = ""
    mPati.姓名 = ""
    mPati.出生日期 = ""
    mPati.民族 = ""
    mPati.门诊号 = 0
    
    txtPatient.Text = ""
    txtNation.Text = ""
    txt性别.Text = ""
    txt出生日期.Text = "____-__-__"
    txt出生时间.Text = "__:__"
    txt年龄.Text = ""
    txt身份证号.Text = ""
    txt门诊号.Text = ""
    txt应退.Text = "0.00"
    txt未退.Text = "0.00"
    vsBooks.Rows = 1
    vsPay.Rows = 1
End Sub

Private Sub txtPatient_GotFocus()
    If Not txtPatient.Enabled Or txtPatient.Locked Then Exit Sub
    zlControl.TxtSelAll txtPatient
    If IsCardType(IDKind, "姓名") Then
        Call zlCommFun.OpenIme(True)
    End If
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim blnCard As Boolean, blnICCard As Boolean
    Dim strCardNo As String
    Dim blnPass As Boolean
    On Error GoTo errH
    If txtPatient.Locked Or txtPatient.Enabled = False Then Exit Sub
    If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub

    If IsCardType(IDKind, "姓名") Then
        '105567:李南春,2017/5/25,卡号加密导致第一个汉字拼音不能触发输入法
        blnPass = txtPatient.PasswordChar <> ""
        blnCard = zlCommFun.InputIsCard(txtPatient, KeyAscii, False)
        txtPatient.IMEMode = 0
        blnPass = txtPatient.PasswordChar = "" And blnPass
        If blnPass Then
            If txtPatient.SelLength = Len(txtPatient.Text) Then
                txtPatient.Text = ""
            End If
            SendKeys Chr(KeyAscii): KeyAscii = 0: Exit Sub
        End If
    ElseIf IsCardType(IDKind, "门诊号") Or IsCardType(IDKind, "住院号") Or IsCardType(IDKind, "手机号") Then
        If KeyAscii <> vbKeyReturn And KeyAscii <> vbKeyBack Then
            If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0: Exit Sub
        End If
    Else
        txtPatient.PasswordChar = IIf(IDKind.ShowPassText, "*", "")
        '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
        txtPatient.IMEMode = 0
    End If
    
    If Not (blnCard And Len(txtPatient.Text) = IDKind.GetCardNoLen - 1 And KeyAscii <> 8 _
        Or KeyAscii = 13 And Trim(txtPatient.Text) <> "") Then
        '不是刷卡和回车,则退出
        Exit Sub
    End If

    If KeyAscii <> 13 Then
        txtPatient.Text = txtPatient.Text & Chr(KeyAscii)
        txtPatient.SelStart = Len(txtPatient.Text)
    End If

    KeyAscii = 0
    strCardNo = Trim(txtPatient.Text)
    If Not GetPatient(txtPatient.Text, blnCard) Then
        Call ClearPatiInfo
        txtPatient.Text = strCardNo: zlControl.TxtSelAll txtPatient

        If InStr(1, "+*-", Left(txtPatient.Text & " ", 1)) > 0 Then
            KeyAscii = 0
            DoEvents
            If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus
            zlControl.TxtSelAll txtPatient
            
            Exit Sub
        End If
        Exit Sub
    End If

    txtPatient.Text = mPati.姓名
    txtPatient.PasswordChar = ""
    '55766:文本框有一bug:如果先为密文显示,后设置成非密文显示后,不能输入五笔
    txtPatient.IMEMode = 0

    Call LoadPatiBooks
    zlCommFun.PressKey vbKeyTab
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtPatient_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Function GetPatient(ByVal strInput As String, Optional ByVal blnCard As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:获取病人信息
    '入参:blnCard=表示是否就诊卡刷卡
    '出参:
    '返回:病人读取成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-03 10:46:29
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim vRect As RECT, rsTemp As ADODB.Recordset
    Dim strSQL As String, strPati As String, strWhere As String, blnHavePass As Boolean
    Dim lng病人ID As Long, blnCancel As Boolean, blnICCard As Boolean
    Dim strPassWord As String, bln存在帐户 As Boolean, strErrMsg As String
    Dim strCardNo As String, lng卡类别ID As Long, blnIsMobileNO As Boolean
    
    txtPatient.ForeColor = &HFF0000
    strErrMsg = ""
    blnIsMobileNO = IDKind.IsMobileNO(strInput)
    If IsCardType(IDKind, "IC卡号") Then blnICCard = (InStr(1, "-+*.", Left(txtPatient.Text, 1)) = 0)
    If blnCard And Not (Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2))) Then    '刷卡或缺省的卡
        
        If IDKind.Cards.按缺省卡查找 And Not IDKind.GetfaultCard Is Nothing Then
            lng卡类别ID = IDKind.GetfaultCard.接口序号
        ElseIf IDKind.GetCurCard.接口序号 > 0 Then
            lng卡类别ID = IDKind.GetCurCard.接口序号
        Else
            lng卡类别ID = -1
        End If
        '短名|全名|读卡标志|卡类别ID|卡号长度|缺省标志(1-当前缺省;0-非缺省)|是否存在帐户(1-存在帐户;0-不存在帐户);…
        If GetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then
            If blnIsMobileNO Then
                '手机号查找
                If GetPatiID("手机号", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
            Else
                GoTo NotFoundPati:
            End If
        End If
        
        If lng病人ID <= 0 Then GoTo NotFoundPati:
        strWhere = strWhere & " And A.病人ID=[1]"
        strCardNo = strInput
        strInput = "-" & lng病人ID
        blnHavePass = True
    ElseIf Left(strInput, 1) = "-" And IsNumeric(Mid(strInput, 2)) Then   '病人ID
        strWhere = strWhere & " And A.病人ID=[1]"
    ElseIf Left(strInput, 1) = "*" And IsNumeric(Mid(strInput, 2)) Then '门诊号
        strWhere = strWhere & " And A.门诊号=[1]"
    ElseIf Left(strInput, 1) = "+" And IsNumeric(Mid(strInput, 2)) Then '住院号
        strWhere = strWhere & " And A.病人ID = (Select Nvl(Max(病人ID),0) As 病人ID From 病案主页 Where 住院号 = [1])"
    ElseIf IsCardType(IDKind, "姓名") And blnIsMobileNO Then
        '手机号查找
        If GetPatiID("手机号", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
        If lng病人ID <= 0 Then GoTo NotFoundPati:
        strWhere = strWhere & " And A.病人ID=[1]"
        strInput = "-" & lng病人ID
    Else
        If mPati.病人ID <> 0 Then
            If mPati.姓名 = strInput Then
                '74309:李南春，2014-7-7，病人姓名显示颜色处理
                Call SetPatiColor(txtPatient, mPati.病人类型, txtPatient.ForeColor)
                GetPatient = True: Exit Function
            End If
        End If
        Select Case IDKind.GetCurCard.名称
            Case "姓名", "姓名或就诊卡"
                '通过姓名模糊查找病人(允许输入病人标识时)
                strPati = _
                " Select A.病人ID as ID,A.病人ID,A.姓名,A.性别,A.年龄," & _
                "        A.门诊号,A.住院号,A.出生日期,A.身份证号,A.手机号" & _
                " From 病人信息 A,部门表 B" & _
                " Where A.停用时间 is NULL And A.当前科室ID=B.ID(+) And Rownum <101 And A.姓名 Like [1]"
                strPati = strPati & " Order by  A.姓名"
                
                vRect = zlControl.GetControlRect(txtPatient.hWnd)
                Set rsTemp = zlDatabase.ShowSQLSelect(Me, strPati, 0, "病人选择", 1, "", "请选择病人", False, False, True, vRect.Left, vRect.Top, txtPatient.Height, blnCancel, False, True, strInput & "%", 101)
                If blnCancel Then GoTo NotFoundPati:
                If rsTemp Is Nothing Then GoTo NotFoundPati:
                If rsTemp.State <> 1 Then GoTo NotFoundPati:
                If rsTemp.RecordCount = 0 Then GoTo NotFoundPati:
                If Val(Nvl(rsTemp!病人ID)) = 0 Then GoTo NotFoundPati:
                
                strInput = "-" & rsTemp!病人ID
                strWhere = strWhere & " And A.病人ID=[1]"
            Case "医保号"
                strInput = UCase(strInput)
                strWhere = strWhere & "  And A.医保号=[2]"
             Case "身份证号", "身份证", "二代身份证"
                strInput = UCase(strInput)
                '问题号:54197
                 If GetPatiID("身份证", strInput, False, lng病人ID, strPassWord, strErrMsg, , , , False) = False Then lng病人ID = 0
                 strInput = "-" & lng病人ID
                strWhere = strWhere & " And A.病人ID=[1]"
            Case "联系人身份证号", "联系人身份证" '问题号:51071
                strInput = UCase(strInput)
                 If GetPatiID("联系人身份证", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                 strInput = "-" & lng病人ID
                strWhere = strWhere & " And A.病人ID=[1]"
            Case "IC卡号", "IC卡"
                strInput = UCase(strInput)
                If GetPatiID("IC卡", strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then lng病人ID = 0
                strInput = "-" & lng病人ID
                strWhere = strWhere & " And A.病人ID=[1]"
            Case "门诊号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strWhere = strWhere & " And A.门诊号=[2]"
            Case "住院号"
                If Not IsNumeric(strInput) Then strInput = "0"
                strWhere = strWhere & " And A.病人ID = (Select Max(病人id) From 病案主页 Where 住院号 = [2])"
            Case Else
                '其他类别的号码
                If Val(IDKind.GetCurCard.接口序号) > 0 Then
                    lng卡类别ID = IDKind.GetCurCard.接口序号
                    bln存在帐户 = IDKind.GetCurCard.是否存在帐户
                    If GetPatiID(lng卡类别ID, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                    If lng病人ID = 0 Then GoTo NotFoundPati:
                    strCardNo = strInput
                    blnHavePass = True
                Else
                    If GetPatiID(IDKind.GetCurCard.名称, strInput, False, lng病人ID, strPassWord, strErrMsg) = False Then GoTo NotFoundPati:
                End If
                If lng病人ID <= 0 Then GoTo NotFoundPati:
                strWhere = strWhere & " And A.病人ID=[1]"
                strInput = "-" & lng病人ID
        End Select
    End If
    On Error GoTo errH
    strSQL = "Select A.病人ID,A.姓名,A.性别,A.民族,A.出生日期,A.年龄,A.身份证号,A.门诊号,A.手机号," & vbNewLine & _
    "         Nvl(Nvl(A.病人类型,B.病人类型),Decode(Nvl(A.险类,B.险类),Null,'普通病人','医保病人')) 病人类型 " & vbNewLine & _
    " From 病人信息 A,病案主页 B " & vbNewLine & _
    " Where A.病人ID=B.病人ID(+) And A.主页ID=B.主页ID(+) And A.停用时间 is NULL " & strWhere
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "病人信息", Val(Mid(strInput, 2)), strInput)
    If rsTemp.EOF Then GoTo NotFoundPati:
    LoadPati (rsTemp!病人ID)
    Call SetPatiColor(txtPatient, mPati.病人类型, txtPatient.ForeColor) '74309:李南春，2014-7-7，病人姓名显示颜色处理
    GetPatient = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    Call ClearPatiInfo
    Exit Function
NotFoundPati:
    If strErrMsg <> "" Then Exit Function
    Call ClearPatiInfo
    If blnCard Then
        MsgBox "不能确定病人信息，请检查是否正确刷卡！", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
    Else
        MsgBox "病人信息未找到,请检查是否输入正确!", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
    End If
End Function

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

Private Sub vsBooks_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim str结算方式 As String
    Dim dblBackMoney As Double
    On Error GoTo Errhand
    With vsBooks
        If NewRow < 1 Then Exit Sub
        If OldRow = NewRow Then Exit Sub
        .Cell(flexcpBackColor, 1, 0, .Rows - 1, .Cols - 1) = C_COLOR_背景
        .Cell(flexcpBackColor, NewRow, 0, NewRow, .Cols - 1) = C_COLOR_蓝色
        
        If .Cell(flexcpData, .RowSel, .ColIndex("记账")) = 1 Then
            dblBackMoney = Val(.TextMatrix(NewRow, .ColIndex("病历费")))
            txt应退.Text = Format(dblBackMoney, "0.00")
            txt未退.Text = Format(dblBackMoney, "0.00")
        Else
            txt应退.Text = "0.00": txt未退.Text = "0.00"
        End If
        
        '定位到支付方式
        vsPay.Clear 1: vsPay.Rows = 1
        mPayInfo.bln医保 = False: mPayInfo.bln全退 = False
        mrsBooks.Filter = " 单据号 = '" & .TextMatrix(NewRow, .ColIndex("单据号")) & "' And 场合 = '" & .TextMatrix(NewRow, .ColIndex("场合")) & "'"
        If mrsBooks.RecordCount = 0 Then
            cmdOK.Enabled = False
        Else
            cmdOK.Enabled = True
            With vsPay
                Do While Not mrsBooks.EOF
                    If Not IsNull(mrsBooks!结算方式) Then
                        .Rows = .Rows + 1
                        .RowData(.Rows - 1) = IIf(Nvl(mrsBooks!性质, 0) = 1, 0, Nvl(mrsBooks!是否退现, 0))
                        .TextMatrix(.Rows - 1, .ColIndex("支付方式")) = mrsBooks!结算方式
                        .Cell(flexcpData, .Rows - 1, .ColIndex("支付方式")) = Val(Nvl(mrsBooks!性质))
                        .TextMatrix(.Rows - 1, .ColIndex("支付金额")) = Format(Nvl(mrsBooks!金额), "0.00")
                        .TextMatrix(.Rows - 1, .ColIndex("卡号")) = Nvl(mrsBooks!卡号)
                        .TextMatrix(.Rows - 1, .ColIndex("交易流水号")) = Nvl(mrsBooks!交易流水号)
                        .TextMatrix(.Rows - 1, .ColIndex("交易说明")) = Nvl(mrsBooks!交易说明)
                        If mrsBooks!是否全退 = 1 Then mPayInfo.bln全退 = True
                        If Val(Nvl(mrsBooks!性质)) = 3 Or Val(Nvl(mrsBooks!性质)) = 4 Then mPayInfo.bln医保 = True
                        .TextMatrix(.Rows - 1, .ColIndex("卡类别ID")) = Nvl(mrsBooks!卡类别ID, Nvl(mrsBooks!结算卡序号, 0))
                        .Cell(flexcpData, .Rows - 1, .ColIndex("卡类别ID")) = IIf(Val(Nvl(mrsBooks!结算卡序号)) > 0, 1, 0)
                        If dblBackMoney > 0 Then
                            .Cell(flexcpChecked, .Rows - 1, .ColIndex("选择")) = 1
                            dblBackMoney = dblBackMoney - Val(Nvl(mrsBooks!金额))
                        End If
                    End If
                    mrsBooks.MoveNext
                Loop
            End With
            mrsBooks.MoveFirst
        End If
    End With
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub Init退费方式()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:加载有效的支付方式
    '编制:刘兴洪
    '日期:2011-07-08 11:41:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim j As Long, strPayType As String, varData As Variant, varTemp As Variant, i As Long
    Dim rsTemp As ADODB.Recordset, blnFind As Boolean
    Dim strSQL As String, objCard As Card, objCards As Cards
    Dim lngKey As Long
    
    Set mobjPayCards = New Cards
    Set objCards = New Cards
    
    On Error GoTo Errhand
    strSQL = "" & _
    "   Select Nvl(A.缺省标志,0) as 缺省,B.编码,B.名称,B.性质,B.应付款" & _
    "   From 结算方式应用 A,结算方式 B" & _
    "   Where A.结算方式=B.名称 And A.应用场合=[1]" & _
    "           And Nvl(B.性质,1) IN(1,2)  " & _
    "   Order by B.编码"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, "就诊卡")
    '短|全名|刷卡标志|卡类别ID(消费卡序号)|长度|是否消费卡|结算方式|是否密文|是否自制卡;…
    If Not gobjSquare Is Nothing Then
    ' zlGetCards(ByVal BytType As Byte)
        '入参:bytType-  0-所有医疗卡;
    '                        1-启用的医疗卡,
    '                        2-所有存在三方账户的三方卡
    '                        3-启用的三方账户的医疗卡
       Set objCards = gobjSquare.objSquareCard.zlGetCards(3)
    End If
    With rsTemp
        .Filter = 0
        If .RecordCount <> 0 Then .MoveFirst
        lngKey = 1
        Do While Not .EOF
            For i = 1 To objCards.count
                If objCards(i).结算方式 = Nvl(rsTemp!名称) Then
                    blnFind = True
                    Exit For
                End If
            Next
            If Not blnFind Then
                If Not (Val(Nvl(rsTemp!性质)) = 3 Or Val(Nvl(rsTemp!性质)) = 4 _
                    Or Val(Nvl(rsTemp!性质)) = 7 Or Val(Nvl(rsTemp!性质)) = 8 _
                    Or Val(Nvl(rsTemp!应付款)) = 1) Then
                    
                    '不加入医保的结算方式或退支票的
                    Set objCard = New Card
                    objCard.短名 = Mid(Nvl(!名称), 1, 1)
                    objCard.接口编码 = Nvl(!编码)
                    objCard.接口程序名 = ""
                    objCard.接口序号 = -1 * lngKey
                    objCard.结算方式 = Nvl(!名称)
                    objCard.名称 = Nvl(!名称)
                    objCard.启用 = True
                    objCard.缺省标志 = Val(Nvl(rsTemp!缺省)) = 1
                    objCard.支付启用 = True
                    objCard.结算性质 = Val(!性质)
                    objCard.是否退现 = True
                     
                    mobjPayCards.Add objCard, "K" & lngKey
                    lngKey = lngKey + 1
              End If
            End If
            .MoveNext
        Loop
    End With
    
    '加三方卡
    For i = 1 To objCards.count
        rsTemp.Filter = "名称='" & objCards(i).结算方式 & "'" '结算方式要设置了"就诊卡"应用场合才能使用
        If Not rsTemp.EOF Then
            mobjPayCards.Add objCards(i), "K" & lngKey
            lngKey = lngKey + 1
        End If
    Next
    
    If mobjPayCards.count = 0 Then
        MsgBox "没有可用的结算方式,请先到结算方式管理中设置。", vbExclamation, gstrSysName
        Exit Sub
    End If
    
    
    Exit Sub
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function IsCheckCancelValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:检查退费时的数据有效性
    '返回:有效,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-12 18:43:19
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strName As String, bln消费卡 As Boolean, lng卡类别ID As Long
    Dim lngRow As Long, objDelObject  As clsCardObject
    Dim dblMoney As Double, strErrMsg As String
   '问题:48249
    Dim strSQL As String, rsBill As Recordset, rsTemp As ADODB.Recordset, lngCardBill As Long
    Dim intStyle As Integer, bln退费 As Boolean
    
    On Error GoTo Errhand
    If mrsBooks Is Nothing Then
        strErrMsg = "没有找到病历费信息，不能退费！"
    ElseIf mrsBooks.EOF Then
        strErrMsg = "没有找到病历费信息，不能退费！"
    End If
    If strErrMsg <> "" Then
        MsgBox strErrMsg, vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    If mPayInfo.bln医保 Then
        MsgBox "由于单据" & Nvl(mrsBooks!单据号) & "使用了医保支付方式，只能通过【门诊挂号管理】进行退费！", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    
    '要检查是否单独收的病历费
    mrsBooks.MoveFirst
    Do While Not mrsBooks.EOF
        If Val(Nvl(mrsBooks!卡类别ID)) > 0 Or Val(Nvl(mrsBooks!结算卡序号)) > 0 Then
            If Nvl(mrsBooks!序号, 0) <> 1 Then
                If mrsBooks!场合 = "挂号" Then
                    MsgBox "使用三方支付的病历费不能单独退费，请到【门诊挂号管理】通过退号功能退费！", vbInformation + vbOKOnly, gstrSysName
                Else
                    MsgBox "使用三方支付的病历费不能单独退费，请到【医疗卡发放管理】通过退卡功能退费！", vbInformation + vbOKOnly, gstrSysName
                End If
                Exit Function
            End If
        End If
        mrsBooks.MoveNext
    Loop
    mrsBooks.MoveFirst
    
    intStyle = Val(zlDatabase.GetPara("已结帐单据操作", 100))
    strSQL = "Select B.NO From 住院费用记录 a,病人结帐记录 b Where a.结帐id=b.id And a.记录性质 In (5,15) And a.记录状态=1 And b.记录状态=1 And a.no=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, App.ProductName, vsBooks.TextMatrix(vsBooks.RowSel, vsBooks.ColIndex("单据号")))
    If rsTemp.EOF Then bln退费 = True
    Select Case intStyle
        Case 0
            bln退费 = True
        Case 1
            If bln退费 = False Then
                If MsgBox("单据" & vsBooks.TextMatrix(vsBooks.RowSel, vsBooks.ColIndex("单据号")) & "已做结账处理，是否继续退费", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    bln退费 = True
                End If
            End If
        Case 2
            If bln退费 = False Then
                MsgBox "单据" & vsBooks.TextMatrix(vsBooks.RowSel, vsBooks.ColIndex("单据号")) & "已做结账处理，必须先结账作废再退费", vbInformation + vbOKOnly, gstrSysName
            End If
    End Select
    If bln退费 = False Then Exit Function
    
    Set mobjDelObjects = New clsCardObjects
    With vsPay
        For lngRow = 1 To .Rows - 1
            If .Cell(flexcpChecked, lngRow, .ColIndex("选择")) = 1 And .TextMatrix(lngRow, .ColIndex("卡类别ID")) > 0 Then
                bln消费卡 = .Cell(flexcpData, lngRow, .ColIndex("卡类别ID")) = 1
                lng卡类别ID = .TextMatrix(lngRow, .ColIndex("卡类别ID"))
                
                If Val(Nvl(mrsBooks!记帐费用)) = 1 Then IsCheckCancelValied = True: Exit Function
                If lng卡类别ID <= 0 Then IsCheckCancelValied = True: Exit Function
            
                '不为零,需要获取相应的支付对象
                Set objDelObject = zlGetClsCardObject(lng卡类别ID, bln消费卡)
            If objDelObject Is Nothing Then
                
                    MsgBox "你未启用选择的退费接口 ,不能在此工作站上进行退费!", vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                End If
                If Not objDelObject.CardPreporty.启用 Then
                    MsgBox "你未启用" & mobjDelObject.CardPreporty.名称 & "接口 ,不能在此工作站上进行退费!", vbInformation + vbOKOnly, gstrSysName
                    Exit Function
                End If
                If objDelObject.CardObject Is Nothing Then
                    If zlCreatePatiCardObject(objDelObject.CardPreporty, mobjDelObject.CardObject) = False Then
                        Exit Function
                    End If
                End If
                If Not objDelObject.InitCompents Then
                    If objDelObject.CardObject.zlInitComponents(Me, mlngModule, glngSys, gstrDBUser, gcnOracle, False, "") = False Then
                          Exit Function
                    End If
                    objDelObject.InitCompents = True
                End If
                
                '4.3.3.2.6   zlReturnCheck:帐户回退交易前的检查
                'zlPaymentCheck帐户扣款交易检查
                '参数名  参数类型    入/出   备注
                'frmMain Object  In  调用的主窗体
                'lngModule   Long    In  模块号
                'lngCardTypeID   Long    In  卡类别ID:医疗卡类别.ID
                'strCardNo   String  IN  卡号
                'strBalanceIDs:格式:收费类型( 1-预交款,2-结帐,3-收费,4-挂号,5-医疗卡收款)|ID1,ID2…IDn||收费类型n|ID1,ID2…IDn
                'dblMoney    Double  IN  退款金额
                'strSwapNo   String  In  交易流水号(退款时检查)
                'strSwapMemo String  In  交易说明(退款时传入)
                '    Boolean 函数返回    True:调用成功,False:调用失败
                '说明:
                '在调用扣款前，由于存在Oracle事务问题，因此，再调用回退交易前，先进行数据的合法性检查,以便控制死锁情况。
            
                '卡类别ID,卡号,是否消费卡(1-是;0-否),交易流水号,交易说明,strNO,结帐ID
                'mcolBillBalance.Add Array(Val(Nvl(rsTmp!卡类别ID)), Trim(Nvl(rsTmp!卡号)), IIf(Val(Nvl(rsTmp!结算卡序号)) <> 0, 1, 0), Trim(Nvl(rsTmp!交易流水号)), Trim(Nvl(rsTmp!交易说明))), strNO
                Dim str卡号 As String, str交易流水号 As String, str交易说明 As String, str结算信息 As String
                Dim strXMLExpend As String, str密码 As String
                str卡号 = .TextMatrix(lngRow, .ColIndex("卡号"))
                str交易流水号 = .TextMatrix(lngRow, .ColIndex("交易流水号"))
                str交易说明 = .TextMatrix(lngRow, .ColIndex("交易说明"))
                str结算信息 = IIf(mrsBooks!场合 = "挂号", 4, 5) & "|" & Nvl(mrsBooks!结帐ID)
                dblMoney = Val(.Cell(flexcpData, lngRow, .ColIndex("支付金额")))
                If objDelObject.CardObject.zlReturncheck(Me, mlngModule, objDelObject.CardPreporty.接口序号, str卡号, str结算信息, dblMoney, str交易流水号, str交易说明, strXMLExpend) = False Then
                    If txtPatient.Enabled And txtPatient.Visible Then txtPatient.SetFocus: Exit Function
                    Exit Function
                End If
                
                '100610:李南春,2016/10/13，预交退款和余额退款是否验证刷卡
                If objDelObject.CardPreporty.是否退款验卡 Then
                '   zlBrushCard(frmMain As Object, _
                    ByVal lngModule As Long, _
                    ByVal lngCardTypeID As Long, _
                    ByVal strPatiName As String, ByVal strSex As String, _
                    ByVal strOld As String, ByRef dbl金额 As Double, _
                    Optional ByRef strCardNo As String, _
                    Optional ByRef strPassWord As String, _
                    Optional ByVal strXmlIn As String = "") As Boolean
                    '       strXmlIn-三方卡调用XML入参,目前格式如下:
                    '       <IN>
                    '           <CZLX>0</CZLX>    //操作类型,0-正常调用刷卡,1-转账调用刷卡,2-退款调用刷卡
                    '       </IN>
                    Err = 0: On Error Resume Next
                    If objDelObject.CardObject.zlBrushCard(Me, mlngModule, objDelObject.CardPreporty.接口序号, _
                     mPati.姓名, mPati.性别, mPati.年龄, dblMoney, _
                     str卡号, str密码, "<IN><CZLX>2</CZLX></IN>") = False Then
                        If Err = 450 Then
                            Err = 0: On Error GoTo Errhand
                            If mobjDelObject.CardObject.zlBrushCard(Me, mlngModule, objDelObject.CardPreporty.接口序号, _
                             mPati.姓名, mPati.性别, mPati.年龄, dblMoney, str卡号, str密码) = False Then Exit Function
                        Else
                            Exit Function
                        End If
                    End If
                End If
    
                mobjDelObjects.Add objDelObject, False, lng卡类别ID, Nothing, bln消费卡, IIf(bln消费卡, "X", "K") & lng卡类别ID
            End If
        Next
    End With
    IsCheckCancelValied = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function IsNoCanc()
    '检查病历费是否已经被退费
    Dim strSQL As String, rsTemp As ADODB.Recordset
    Dim strTable As String, strWhere As String, blnHave As Boolean
    Dim i As Integer, dblBalance As Double
    On Error GoTo Errhand
    If mrsBooks Is Nothing Then Exit Function
    If mrsBooks.EOF Then Exit Function
    
    With vsPay
        For i = 1 To .Rows - 1
            If .Cell(flexcpChecked, i, .ColIndex("选择")) = 1 Then
                dblBalance = dblBalance + CDbl(.Cell(flexcpData, i, .ColIndex("支付金额")))
                blnHave = True
            End If
        Next
        If dblBalance <> CDbl(txt应退.Text) And blnHave Then
            MsgBox "退费金额不一致，请重新选择退费金额。", vbInformation + vbOKOnly, gstrSysName
            Exit Function
        End If
    End With
    If Not blnHave Then
        mrsBooks.MoveFirst
        Do While Not mrsBooks.EOF
            If mrsBooks!是否退现 = 0 Then
                MsgBox "原支付方式不支持退现，请选择退费方式。", vbInformation + vbOKOnly, gstrSysName
                Exit Function
            End If
            mrsBooks.MoveNext
        Loop
        mrsBooks.MoveFirst
    End If
    
    If Nvl(mrsBooks!场合) = "挂号" Then
        strTable = "门诊费用记录"
        strWhere = "记录性质=4 And 记录状态 = 1 And 附加标志=1"
    Else
        strTable = "住院费用记录"
        strWhere = "记录性质=5 And 记录状态 = 1 and 附加标志=8"
    End If
    
    strSQL = "Select 1 From " & strTable & " Where " & strWhere & " And 病人ID=[1] And no=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "病历费检查", mPati.病人ID, Nvl(mrsBooks!单据号))
    If rsTemp.RecordCount = 0 Then
        MsgBox "当前病历费已被其他人员退费!", vbInformation + vbOKOnly, gstrSysName
        Exit Function
    End If
    IsNoCanc = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub vsPay_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim dblMoney As Double, i As Long
    
    With vsPay
        If Row < 1 Or Col <> .ColIndex("选择") Then Exit Sub
        If mPayInfo.bln医保 Then
            .Cell(flexcpChecked, Row, .ColIndex("选择")) = 2
            MsgBox "由于单据" & vsBooks.TextMatrix(vsBooks.RowSel, .ColIndex("单据号")) & "使用了医保支付方式，只能通过【门诊挂号管理】进行退费！", vbInformation + vbOKOnly, gstrSysName
            Exit Sub
        End If
        '取消选择
        If .Cell(flexcpChecked, Row, .ColIndex("选择")) = 2 Then
            .Cell(flexcpChecked, Row, .ColIndex("支付金额")) = 0
            For i = 1 To .Rows - 1
                If .Cell(flexcpChecked, i, .ColIndex("选择")) = 1 Then dblMoney = dblMoney + CDbl(.TextMatrix(i, .ColIndex("支付金额")))
            Next
            txt未退.Text = Format(CDbl(txt应退.Text) - dblMoney, "0.00")
            Exit Sub
        End If
        '如果已经超过了退款金额，不能再选择
        If Val(txt未退.Text) = 0 Then .Cell(flexcpChecked, Row, .ColIndex("选择")) = 2: Exit Sub
        dblMoney = Val(.TextMatrix(Row, .ColIndex("支付金额")))
        '如果选择的结算金额大于退款，取消其他选择
        If dblMoney >= CDbl(txt应退.Text) Then
            .Cell(flexcpChecked, 1, .ColIndex("选择"), .Rows - 1, .ColIndex("选择")) = 2
            .Cell(flexcpData, Row, .ColIndex("支付金额"), .Rows - 1, .ColIndex("支付金额")) = 0
            .Cell(flexcpChecked, Row, .ColIndex("选择")) = 1
            .Cell(flexcpData, Row, .ColIndex("支付金额")) = CDbl(txt应退.Text)
            txt未退.Text = "0.00"
        Else
            '求未退金额
            dblMoney = CDbl(txt未退.Text) - dblMoney
            If dblMoney <= 0 Then
                .Cell(flexcpData, Row, .ColIndex("支付金额")) = Val(txt未退.Text)
                txt未退.Text = "0.00"
            Else
                .Cell(flexcpData, Row, .ColIndex("支付金额")) = Val(txt未退.Text) - dblMoney
                txt未退.Text = Format(dblMoney, "0.00")
            End If
        End If
    End With
End Sub

Private Sub vsPay_EnterCell()
    With vsPay
        If .RowSel < 1 Then Exit Sub
        .Cell(flexcpBackColor, 1, 0, .Rows - 1, .Cols - 1) = C_COLOR_背景
        .Cell(flexcpBackColor, .RowSel, 0, .RowSel, .Cols - 1) = C_COLOR_蓝色
    End With
End Sub

Private Sub vsPay_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> vsPay.ColIndex("选择") Then Cancel = True
End Sub

Private Function SaveData() As Boolean
    Dim strSQL As String, rsTemp As ADODB.Recordset, str结算方式 As String
    Dim blnTrans As Boolean, blnHave As Boolean, blnOraclTrans As Boolean
    Dim i As Integer, strBalance As String, dblDeposit As Double
    On Error GoTo Errhand
    If mrsBooks Is Nothing Then Exit Function
    If mrsBooks.EOF Then Exit Function
    
    If Nvl(mrsBooks!场合) = "挂号" Then
        With vsPay
            For i = 1 To .Rows - 1
                If .Cell(flexcpChecked, i, .ColIndex("选择")) = 1 Then
                    If .Cell(flexcpData, i, .ColIndex("支付方式")) = 7 Or .Cell(flexcpData, i, .ColIndex("支付方式")) = 8 Then
                        strBalance = strBalance & "|" & .TextMatrix(i, .ColIndex("支付方式")) & "," & .Cell(flexcpData, i, .ColIndex("支付金额")) & "," & "1"
                    ElseIf .Cell(flexcpData, i, .ColIndex("支付方式")) = 0 Then
                    '退预交
                    dblDeposit = CDbl(.Cell(flexcpData, i, .ColIndex("支付金额")))
                    Else
                        strBalance = strBalance & "|" & .TextMatrix(i, .ColIndex("支付方式")) & "," & .Cell(flexcpData, i, .ColIndex("支付金额")) & "," & "0"
                    End If
                    blnHave = True
                End If
            Next
        End With
        '没有选支付方式，按退现金处理
        If blnHave = False Then
            strSQL = "Select 名称 from 结算方式 Where 性质 = 1"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            If rsTemp.RecordCount > 0 Then
                strBalance = Nvl(rsTemp!名称) & "," & Val(txt应退.Text) & "," & "0"
            Else
                strBalance = "现金," & Val(txt应退.Text) & "," & "0"
            End If
        End If
        'zl_病人挂号记录_Delete
        strSQL = "zl_病人挂号记录_出诊_DELETE("
        '  单据号_In       门诊费用记录.No%Type,
        strSQL = strSQL & "'" & Nvl(mrsBooks!单据号) & "',"
        '  操作员编号_In   门诊费用记录.操作员编号%Type,
        strSQL = strSQL & "'" & UserInfo.编号 & "',"
        '  操作员姓名_In   门诊费用记录.操作员姓名%Type,
        strSQL = strSQL & "'" & UserInfo.姓名 & "',"
        '  摘要_In         门诊费用记录.摘要%Type := Null, --预约取消时 填写 存放预约取消原因
        strSQL = strSQL & " NULL ,"
        '  删除门诊号_In   Number := 0,
        strSQL = strSQL & "" & 0 & ","
        '  非原样退结算_In Varchar2 := Null,
        strSQL = strSQL & "NULL" & ","
        '  退费类型_In     In Number := 0, --0-全退 1-退挂号费 2-退病历费
        strSQL = strSQL & "" & 2 & ","
        '  退指定结算_In   病人预交记录.结算方式%Type := Null
        strSQL = strSQL & "NULL" & ","
        '  退号重用_In   Number := 1
        strSQL = strSQL & 1 & ",'"
        '  结算方式_In   Varchar2 := Null
        strSQL = strSQL & strBalance & "',"
        '   退预交_In       病人预交记录.冲预交%Type
        strSQL = strSQL & dblDeposit & ")"
    Else
        With vsPay
            For i = 1 To .Rows - 1
                If .Cell(flexcpChecked, i, .ColIndex("选择")) = 1 Then
                    str结算方式 = .TextMatrix(i, .ColIndex("支付方式"))
                    blnHave = True: Exit For
                End If
            Next
        End With
        '没有选支付方式，按退现金处理
        If blnHave = False Then
            strSQL = "Select 名称 from 结算方式 Where 性质 = 1"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            If rsTemp.RecordCount > 0 Then
                str结算方式 = Nvl(rsTemp!名称)
            Else
                str结算方式 = "现金"
            End If
        End If
        strSQL = "zl_医疗卡记录_DELETE('" & Nvl(mrsBooks!单据号) & "','" & UserInfo.编号 & "','" & UserInfo.姓名 & "',2,'" & str结算方式 & "')"
    End If
    gcnOracle.BeginTrans: blnTrans = True
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    If CallBackBalanceInterface(Nvl(mrsBooks!单据号), blnOraclTrans) = False Then
        If blnOraclTrans = False Then gcnOracle.RollbackTrans
        Exit Function
    End If
    If blnOraclTrans = False Then gcnOracle.CommitTrans
    blnTrans = False
    SaveData = True
    Exit Function
Errhand:
    If blnTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CallBackBalanceInterface(ByVal strNO As String, ByRef blnTrancs As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:调用回退接口
    '入参:blnTrancs-是否处理了事务
    '返回:成功,返回true,否则返回False
    '编制:刘兴洪
    '日期:2011-07-13 10:33:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim str结算信息 As String, str卡号 As String, strSwapGlideNO As String, strSwapMemo As String, strSwapExtendInfor As String
    Dim varData As Variant, varTemp As Variant, i As Long, lng结帐ID As Long, cllPro As Collection, cllProAfter As Collection
    Dim bln消费卡 As Boolean, lng卡类别ID As Long, strSQL As String, rsTemp As ADODB.Recordset
    Dim str交易信息 As String, strTemp As String, dblMoney As Double, blnThree As Boolean
    Dim objDelObject  As clsCardObject, lngRow As Long
    On Error GoTo errHandle
    blnTrancs = False
    
    If Val(Nvl(mrsBooks!记帐费用)) = 1 Then CallBackBalanceInterface = True: Exit Function
    Set cllPro = New Collection: Set cllProAfter = New Collection
    With vsPay
        For lngRow = 1 To .Rows - 1
            If .Cell(flexcpChecked, lngRow, .ColIndex("选择")) = 1 And .TextMatrix(lngRow, .ColIndex("卡类别ID")) > 0 Then
                bln消费卡 = .Cell(flexcpData, lngRow, .ColIndex("卡类别ID")) = 1
                lng卡类别ID = .TextMatrix(lngRow, .ColIndex("卡类别ID"))
                
                If lng结帐ID = 0 Then
                    If .TextMatrix(.Row, .Col) = "挂号" Then
                        strSQL = "Select 结帐ID,记帐费用 From 门诊费用记录  Where 记录性质=4 and NO=[1] and 记录状态=2"
                    Else
                        strSQL = "Select 结帐ID,记帐费用 From 住院费用记录  Where 记录性质=5 and NO=[1] and 记录状态=2"
                    End If
                    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strNO)
                    If rsTemp.EOF Then
                        gcnOracle.RollbackTrans: blnTrancs = True
                        MsgBox "未找到病历费的退费信息信息，不能继续", vbInformation + vbOKOnly, gstrSysName
                        Exit Function
                    End If
                    lng结帐ID = Val(Nvl(rsTemp!结帐ID))
                End If
                If .TextMatrix(.Row, .Col) = "挂号" Then
                    strSwapExtendInfor = "4|" & lng结帐ID: strTemp = strSwapExtendInfor
                Else
                    strSwapExtendInfor = "5|" & lng结帐ID: strTemp = strSwapExtendInfor
                End If
                
                '退费接口
                Set objDelObject = mobjDelObjects(IIf(bln消费卡, "X", "K") & lng卡类别ID)
                If Not objDelObject.InitCompents Then
                    If objDelObject.CardObject.zlInitComponents(Me, mlngModule, glngSys, gstrDBUser, gcnOracle, False, "") = False Then
                          Exit Function
                    End If
                    objDelObject.InitCompents = True
                End If
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
                '       strSwapNo-交易流水号(收款时的交易流水号)
                '       strSwapMemo-交易说明(收款时的交易说明)
                '       strSwapExtendInfor-交易的扩展信息
                '           格式为:项目名称1|项目内容2||…||项目名称n|项目内容n 每个项目中不能包含|字符
                str卡号 = .TextMatrix(lngRow, .ColIndex("卡号"))
                strSwapGlideNO = .TextMatrix(lngRow, .ColIndex("交易流水号"))
                strSwapMemo = .TextMatrix(lngRow, .ColIndex("交易说明"))
                str结算信息 = IIf(mrsBooks!场合 = "挂号", 4, 5) & "|" & Nvl(mrsBooks!结帐ID)
                dblMoney = Val(.Cell(flexcpData, lngRow, .ColIndex("支付金额")))
                If objDelObject.CardObject.zlReturnMoney(Me, mlngModule, lng卡类别ID, str卡号, str结算信息, dblMoney, strSwapGlideNO, strSwapMemo, strSwapExtendInfor) = False Then
                    Exit Function
                End If
                
                '更新交易信息
                '    Zl_三方接口更新_Update
                strSQL = "Zl_三方接口更新_Update("
                '  卡类别id_In   病人预交记录.卡类别id%Type,
                strSQL = strSQL & "" & lng卡类别ID & ","
                '  消费卡_In     Number,
                strSQL = strSQL & "" & IIf(bln消费卡, 1, 0) & ","
                '  卡号_In       病人预交记录.卡号%Type,
                strSQL = strSQL & "'" & str卡号 & "',"
                '  结帐ids_In    Varchar2,
                strSQL = strSQL & "'" & lng结帐ID & "',"
                '  交易流水号_In 病人预交记录.交易流水号%Type,
                strSQL = strSQL & "'" & strSwapGlideNO & "',"
                '  交易说明_In   病人预交记录.交易说明%Type
                strSQL = strSQL & "'" & strSwapMemo & "')"
                zlAddArray cllPro, strSQL
                
                If strTemp <> strSwapExtendInfor Then
                    'strSwapExtendInfor:交易扩展信息,格式:项目名称|项目内容||...
                    varData = Split(strSwapExtendInfor, "||")
                    Set cllPro = New Collection
                    For i = 0 To UBound(varData)
                        If Trim(varData(i)) <> "" Then
                            varTemp = Split(varData(i) & "|", "|")
                            If varTemp(0) <> "" Then
                                strTemp = varTemp(0) & "|" & varTemp(1)
                                If zlCommFun.ActualLen(str交易信息 & "||" & strTemp) > 2000 Then
                                    str交易信息 = Mid(str交易信息, 3)
                                    'Zl_三方结算交易_Insert
                                    strSQL = "Zl_三方结算交易_Insert("
                                    '卡类别id_In 病人预交记录.卡类别id%Type,
                                    strSQL = strSQL & "" & lng卡类别ID & ","
                                    '消费卡_In   Number,
                                    strSQL = strSQL & "" & IIf(bln消费卡, 1, 0) & ","
                                    '卡号_In     病人预交记录.卡号%Type,
                                    strSQL = strSQL & "'" & str卡号 & "',"
                                    '结帐ids_In  Varchar2,
                                    strSQL = strSQL & "'" & lng结帐ID & "',"
                                    '交易信息_In Varchar2:交易项目|交易内容||...
                                    strSQL = strSQL & "'" & str交易信息 & "')"
                                    zlAddArray cllProAfter, strSQL
                                    str交易信息 = ""
                                End If
                                str交易信息 = str交易信息 & "||" & strTemp
                            End If
                        End If
                    Next
                    If str交易信息 <> "" Then
                        str交易信息 = Mid(str交易信息, 3)
                        'Zl_三方结算交易_Insert
                        strSQL = "Zl_三方结算交易_Insert("
                        '卡类别id_In 病人预交记录.卡类别id%Type,
                        strSQL = strSQL & "" & lng卡类别ID & ","
                        '消费卡_In   Number,
                        strSQL = strSQL & "" & IIf(bln消费卡, 1, 0) & ","
                        '卡号_In     病人预交记录.卡号%Type,
                        strSQL = strSQL & "'" & str卡号 & "',"
                        '结帐ids_In  Varchar2,
                        strSQL = strSQL & "'" & lng结帐ID & "',"
                        '交易信息_In Varchar2:交易项目|交易内容||...
                        strSQL = strSQL & "'" & str交易信息 & "')"
                        zlAddArray cllProAfter, strSQL
                    End If
                End If
            End If
        Next
    End With
    
    '更新交易信息,先提交,这样避免风险,再更新相关的交易信息
    zlExecuteProcedureArrAy cllPro, Me.Caption, , True

    Err = 0: On Error GoTo ErrOthers:
    zlExecuteProcedureArrAy cllProAfter, Me.Caption

    CallBackBalanceInterface = True: blnTrancs = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans: blnTrancs = True
    Call ErrCenter
    Exit Function
ErrOthers:
    '扩展信息,允许保存一部分,以便查证
    If ErrCenter() = 1 Then
        gcnOracle.RollbackTrans
        Resume
    End If
    CallBackBalanceInterface = True
    gcnOracle.CommitTrans: blnTrancs = True
End Function

'控件名称是否匹配
Private Function IsCardType(ByVal IDKindCtl As IDKindNew, ByVal strCardName As String) As Boolean
    If IDKindCtl Is Nothing Then Exit Function
    If UCase(TypeName(IDKindCtl)) <> "IDKINDNEW" Then Exit Function
    Select Case strCardName
     Case "姓名", "姓名或就诊卡"
          IsCardType = IDKindCtl.GetCurCard.名称 Like "姓名*"
     Case "身份证", "身份证号", "二代身份证"
          IsCardType = IDKindCtl.GetCurCard.名称 Like "*身份证*"
     Case "IC卡号", "IC卡"
          IsCardType = IDKindCtl.GetCurCard.名称 Like "IC卡*"
     Case "医保号"
          IsCardType = IDKindCtl.GetCurCard.名称 = "医保号"
     Case "门诊号"
          IsCardType = IDKindCtl.GetCurCard.名称 = "门诊号"
     Case "住院号"
          IsCardType = IDKindCtl.GetCurCard.名称 = "住院号"
     Case "手机号"
          IsCardType = IDKindCtl.GetCurCard.名称 = "手机号"
     Case Else
            If IDKindCtl.GetCurCard Is Nothing Then Exit Function
            If Not IsNumeric(strCardName) Or Val(strCardName) <= 0 Then
                  IsCardType = strCardName = IDKindCtl.GetCurCard.名称
            Else
                If IDKindCtl.GetCurCard.接口序号 <= 0 Then Exit Function
                IsCardType = IDKindCtl.GetCurCard.接口序号 = Val(strCardName)
            End If
     End Select
End Function
