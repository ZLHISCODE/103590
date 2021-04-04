VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form FrmDrugPaymentCard 
   Caption         =   "药品付款单"
   ClientHeight    =   6975
   ClientLeft      =   600
   ClientTop       =   2550
   ClientWidth     =   11400
   Icon            =   "FrmDrugPaymentCard.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   Picture         =   "FrmDrugPaymentCard.frx":0E42
   ScaleHeight     =   6975
   ScaleWidth      =   11400
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshProvider 
      Height          =   3495
      Left            =   780
      TabIndex        =   46
      Top             =   1800
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   6165
      _Version        =   393216
      FixedCols       =   0
      GridColor       =   32768
      AllowBigSelection=   0   'False
      FocusRect       =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Fra2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      ForeColor       =   &H80000008&
      Height          =   5565
      Left            =   6720
      TabIndex        =   31
      Top             =   -90
      Width           =   4935
      Begin VB.PictureBox picdown 
         BorderStyle     =   0  'None
         Height          =   1155
         Left            =   240
         ScaleHeight     =   1155
         ScaleWidth      =   4455
         TabIndex        =   45
         Top             =   4290
         Width           =   4455
         Begin VB.TextBox Txt付款说明 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   870
            MaxLength       =   50
            TabIndex        =   20
            Top             =   0
            Width           =   3585
         End
         Begin VB.TextBox Txt填制日期 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   2580
            Locked          =   -1  'True
            TabIndex        =   24
            Top             =   420
            Width           =   1875
         End
         Begin VB.TextBox Txt审核日期 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   2580
            Locked          =   -1  'True
            TabIndex        =   28
            Top             =   810
            Width           =   1875
         End
         Begin VB.TextBox Txt审核人 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   630
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   810
            Width           =   1005
         End
         Begin VB.TextBox Txt填制人 
            Appearance      =   0  'Flat
            Height          =   300
            Left            =   630
            Locked          =   -1  'True
            TabIndex        =   22
            Top             =   420
            Width           =   1005
         End
         Begin VB.Label Lbl付款说明 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "付款说明:"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   0
            TabIndex        =   19
            Top             =   60
            Width           =   810
         End
         Begin VB.Label Lbl审核日期 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "审核日期"
            Height          =   180
            Left            =   1770
            TabIndex        =   27
            Top             =   885
            Width           =   720
         End
         Begin VB.Label Lbl审核人 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "审核人"
            Height          =   180
            Left            =   30
            TabIndex        =   25
            Top             =   870
            Width           =   540
         End
         Begin VB.Label Lbl填制日期 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "填制日期"
            Height          =   180
            Left            =   1800
            TabIndex        =   23
            Top             =   480
            Width           =   720
         End
         Begin VB.Label Lbl填制人 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "填制人"
            Height          =   180
            Left            =   30
            TabIndex        =   21
            Top             =   480
            Width           =   540
         End
      End
      Begin VB.PictureBox picup 
         Appearance      =   0  'Flat
         BackColor       =   &H80000004&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1725
         Left            =   240
         ScaleHeight     =   1725
         ScaleWidth      =   4455
         TabIndex        =   37
         Top             =   720
         Width           =   4455
         Begin VB.TextBox TxtNo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H00C00000&
            Height          =   300
            Left            =   2925
            Locked          =   -1  'True
            MaxLength       =   8
            TabIndex        =   38
            Top             =   0
            Width           =   1275
         End
         Begin VB.Label lbl银行帐号 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "银行帐号:"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   180
            TabIndex        =   54
            Top             =   1260
            Width           =   810
         End
         Begin VB.Label txt银行帐号 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   960
            TabIndex        =   53
            Top             =   1260
            Width           =   90
         End
         Begin VB.Label txt附件数 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   3840
            TabIndex        =   51
            Top             =   1035
            Width           =   90
         End
         Begin VB.Label txt税务号 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   960
            TabIndex        =   50
            Top             =   1530
            Width           =   90
         End
         Begin VB.Label txt开户行 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   960
            TabIndex        =   49
            Top             =   990
            Width           =   90
         End
         Begin VB.Label txt电话地址 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   960
            TabIndex        =   48
            Top             =   720
            Width           =   90
         End
         Begin VB.Label txt单位名称 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   960
            TabIndex        =   47
            Top             =   450
            Width           =   90
         End
         Begin VB.Label LblNo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "NO"
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
            Left            =   2595
            TabIndex        =   44
            Top             =   45
            Width           =   240
         End
         Begin VB.Label Lbl单位名称 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "单位名称:"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   180
            TabIndex        =   43
            Top             =   450
            Width           =   810
         End
         Begin VB.Label Lbl电话地址 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "地址电话:"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   180
            TabIndex        =   42
            Top             =   720
            Width           =   810
         End
         Begin VB.Label Lbl开户行 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "开户银行:"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   180
            TabIndex        =   41
            Top             =   990
            Width           =   810
         End
         Begin VB.Label Lbl附件数 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "附件数:"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   3180
            TabIndex        =   40
            Top             =   1035
            Width           =   630
         End
         Begin VB.Label lbl税务号 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "税务登记号:"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   0
            TabIndex        =   39
            Top             =   1530
            Width           =   990
         End
      End
      Begin ZL9BillEdit.BillEdit mshPaymentList 
         Height          =   1665
         Left            =   240
         TabIndex        =   18
         Top             =   2535
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   2937
         Appearance      =   0
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
      Begin VB.Label Lbl标题 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "药品付款通知单"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   1155
         TabIndex        =   35
         Top             =   330
         Width           =   2100
      End
   End
   Begin VB.Frame Fra1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      ForeColor       =   &H80000008&
      Height          =   5565
      Left            =   0
      TabIndex        =   32
      Top             =   -90
      Width           =   6735
      Begin MSComctlLib.TreeView tvwProvider 
         Height          =   3585
         Left            =   1320
         TabIndex        =   36
         Top             =   2160
         Visible         =   0   'False
         Width           =   5505
         _ExtentX        =   9710
         _ExtentY        =   6324
         _Version        =   393217
         HideSelection   =   0   'False
         Indentation     =   441
         LabelEdit       =   1
         LineStyle       =   1
         Style           =   7
         ImageList       =   "imgTree"
         BorderStyle     =   1
         Appearance      =   0
      End
      Begin VB.TextBox Txt供药单位 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   930
         TabIndex        =   1
         Top             =   210
         Width           =   4275
      End
      Begin VB.CommandButton Cmd供应商 
         Caption         =   "…"
         Height          =   300
         Left            =   5220
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   210
         Width           =   255
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid mshPurchaseList 
         DragIcon        =   "FrmDrugPaymentCard.frx":1184
         Height          =   2745
         Left            =   30
         TabIndex        =   12
         Top             =   1425
         Width           =   7125
         _ExtentX        =   12568
         _ExtentY        =   4842
         _Version        =   393216
         BackColor       =   -2147483624
         FixedCols       =   0
         SelectionMode   =   1
         AllowUserResizing=   3
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSComctlLib.ImageList imgTree 
         Left            =   6120
         Top             =   90
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
               Picture         =   "FrmDrugPaymentCard.frx":12CE
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDrugPaymentCard.frx":2FDA
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "FrmDrugPaymentCard.frx":4CE4
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin ZL9BillEdit.BillEdit mshImprest 
         Height          =   885
         Left            =   120
         TabIndex        =   14
         Top             =   4560
         Width           =   4455
         _ExtentX        =   7858
         _ExtentY        =   1561
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
      Begin VB.TextBox Txt发票号 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   930
         TabIndex        =   4
         ToolTipText     =   "输入格式:开始发票号-结束发票号"
         Top             =   600
         Width           =   1995
      End
      Begin VB.TextBox Txt单号 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   315
         Left            =   3480
         TabIndex        =   6
         ToolTipText     =   "输入格式:开始NO-结束NO"
         Top             =   600
         Width           =   1995
      End
      Begin VB.Label Lbl单号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "NO"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3240
         TabIndex        =   5
         Top             =   660
         Width           =   180
      End
      Begin VB.Label Lbl发票号 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "发票号"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   300
         TabIndex        =   3
         Top             =   667
         Width           =   540
      End
      Begin VB.Label Label1 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "预付款发票清单"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   13
         Top             =   4320
         Width           =   1260
      End
      Begin VB.Label Lbl付款金额 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "付款金额:"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3720
         TabIndex        =   10
         Top             =   1110
         Width           =   810
      End
      Begin VB.Label Txt付款金额 
         Height          =   180
         Left            =   4650
         TabIndex        =   11
         Top             =   1110
         Width           =   2010
      End
      Begin VB.Label Lbl供应商 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "供药单位"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   0
         Top             =   270
         Width           =   720
      End
      Begin VB.Label Lbl清单 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "未付款发票清单"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   1110
         Width           =   1260
      End
      Begin VB.Label Lbl累计合计 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "累计应付:"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1740
         TabIndex        =   8
         Top             =   1110
         Width           =   810
      End
      Begin VB.Label Lbl合计 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   2730
         TabIndex        =   9
         Top             =   1110
         Width           =   1065
      End
   End
   Begin VB.Frame Fra4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   0
      TabIndex        =   34
      Top             =   5370
      Width           =   6735
      Begin VB.CommandButton cmdHelp 
         Caption         =   "帮助(&H)"
         Height          =   350
         Left            =   240
         Picture         =   "FrmDrugPaymentCard.frx":5B36
         TabIndex        =   52
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton Cmd按选择付款 
         Caption         =   "按选择付款…(&U)"
         Enabled         =   0   'False
         Height          =   350
         Left            =   5040
         Picture         =   "FrmDrugPaymentCard.frx":5C80
         TabIndex        =   17
         Top             =   240
         Width           =   1575
      End
      Begin VB.CommandButton Cmd全选 
         Caption         =   "全选(&A)"
         Enabled         =   0   'False
         Height          =   350
         Left            =   1800
         Picture         =   "FrmDrugPaymentCard.frx":5DCA
         TabIndex        =   15
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton Cmd清除 
         Caption         =   "全清(&C)"
         Enabled         =   0   'False
         Height          =   350
         Left            =   3210
         Picture         =   "FrmDrugPaymentCard.frx":5F14
         TabIndex        =   16
         Top             =   240
         Width           =   1100
      End
   End
   Begin VB.Frame Fra3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000016&
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   6720
      TabIndex        =   33
      Top             =   5370
      Width           =   4935
      Begin VB.CommandButton Cmd取消核销 
         Cancel          =   -1  'True
         Caption         =   "取消(&C)"
         Height          =   350
         Left            =   2700
         Picture         =   "FrmDrugPaymentCard.frx":605E
         TabIndex        =   30
         Top             =   240
         Width           =   1100
      End
      Begin VB.CommandButton Cmd核销 
         Caption         =   "确定(&O)"
         Enabled         =   0   'False
         Height          =   350
         Left            =   900
         Picture         =   "FrmDrugPaymentCard.frx":61A8
         TabIndex        =   29
         Top             =   240
         Width           =   1100
      End
   End
   Begin VB.Timer LimitTime 
      Enabled         =   0   'False
      Interval        =   8000
      Left            =   6660
      Top             =   0
   End
   Begin MSComctlLib.ImageList imlTbrClr 
      Left            =   705
      Top             =   -1035
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDrugPaymentCard.frx":62F2
            Key             =   "PrintView"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDrugPaymentCard.frx":650E
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDrugPaymentCard.frx":672A
            Key             =   "Add"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDrugPaymentCard.frx":6946
            Key             =   "Modify"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDrugPaymentCard.frx":6B62
            Key             =   "Delete"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDrugPaymentCard.frx":6D7E
            Key             =   "Annul"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDrugPaymentCard.frx":6F9A
            Key             =   "Store"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDrugPaymentCard.frx":71B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDrugPaymentCard.frx":73D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDrugPaymentCard.frx":76EE
            Key             =   "Find"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDrugPaymentCard.frx":790A
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmDrugPaymentCard.frx":7B26
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmDrugPaymentCard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private CurMoney As Currency            '当前付款总金额
Private CurLastMoney As Currency        '上次选择付款金额
Private mintUnit As Integer             '0:药库单位 1:门诊单位 2:住院单位 3:售价单位

Private mblnSave As Boolean
Private mblnSuccess As Boolean
Private mstr单据号 As String
Private mint编辑状态 As Integer         '编辑属性 1:表示新增;2:表示修改,3:表示审核;4:表示取消
Private mint记录状态 As Integer
Private mblnChange As Boolean
Private mintParallelRecord As Integer   '对于新增后单据并行执行的处理： 1、代表正常情况；2、已经删除的记录；3、已经审核的记录
Private mintpurchaseclick As Boolean
Dim mstrPrivs As String                     '权限

Private Const mconintCol标志 As Integer = 0
Private Const mconintcol发票号  As Integer = 1
Private Const mconintcol入库单号 As Integer = 2
Private Const mconintcol药品信息 As Integer = 3
Private Const mconIntCol规格 As Integer = 4
Private Const mconIntCol单位 As Integer = 5
Private Const mconintcol发票金额 As Integer = 6
Private Const mconIntCol数量 As Integer = 7
Private Const mconIntCol采购价 As Integer = 8
Private Const mconintcol批发价 As Integer = 9
Private Const mconintcol批发金额 As Integer = 10
Private Const mconIntCol售价 As Integer = 11
Private Const mconIntCol售价金额 As Integer = 12
Private Const mconIntCol产地 As Integer = 13
Private Const mconIntCol批号 As Integer = 14
Private Const mconintcol入库日期 As Integer = 15

Private Const mconIntColS As Integer = 16

Private Function GetDepend() As Boolean
    Dim rsDepend As New Recordset
    Dim rs结算方式 As New Recordset
    Dim intLop As Integer
    
    On Error GoTo errHandle
    GetDepend = False
    With rsDepend
        If .State = 1 Then .Close
        gstrSQL = "Select ID,上级ID,编码,简码,名称,末级,地址||电话 as 电话地址,开户银行,帐号,税务登记号 From 药品供应商 Where " & _
              " To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01' Start with 上级ID is Null Connect by prior ID=上级ID"
        Call SQLTest(App.Title, "药品付款单", gstrSQL)
        Set rsDepend = zldatabase.OpenSQLRecord(gstrSQL, "GetDepend")
        Call SQLTest
        
        If .EOF Then
            MsgBox "药品供应商的信息不全，请在供药单位管理中进行设置！", vbInformation, gstrSysName
            Exit Function
        End If
        
    End With
        
    
    With rs结算方式
        If .State = 1 Then .Close
        gstrSQL = "Select * From 结算方式应用 Where 应用场合='付药款' Order by 缺省标志 desc"
        Call SQLTest(App.Title, "药品付款单", gstrSQL)
        Set rs结算方式 = zldatabase.OpenSQLRecord(gstrSQL, "GetDepend")
        Call SQLTest
        
        If .EOF Then
            MsgBox "结算方式应用信息不全,请在结算方式管理中进行设置！", vbInformation, gstrSysName
            Exit Function
        End If
        mshPaymentList.Clear
        For intLop = 1 To .RecordCount
            mshPaymentList.AddItem !结算方式
            .MoveNext
        Next
        mshPaymentList.ListIndex = 0
        
        .Close
    End With
    
    With rsDepend
        tvwProvider.Nodes.Clear
        tvwProvider.Nodes.Add , , "R", "所有供应商", 1, 1
        tvwProvider.Nodes("R").Tag = 0
        .MoveFirst
        
        Do While Not .EOF
            If IsNull(!上级ID) Then
                If !末级 = 1 Then
                    tvwProvider.Nodes.Add "R", 4, "K_" & !Id, "[" & !编码 & "]" & !名称, 3, 3
                Else
                    tvwProvider.Nodes.Add "R", 4, "K_" & !Id, "[" & !编码 & "]" & !名称, 2, 2
                End If
            Else
                If !末级 = 1 Then
                    tvwProvider.Nodes.Add "K_" & !上级ID, 4, "K_" & !Id, "[" & !编码 & "]" & !名称, 3, 3
                Else
                    tvwProvider.Nodes.Add "K_" & !上级ID, 4, "K_" & !Id, "[" & !编码 & "]" & !名称, 2, 2
                End If
            End If
            tvwProvider.Nodes("K_" & !Id).Tag = !末级
            .MoveNext
        Loop
        tvwProvider.Nodes("R").Selected = True
        tvwProvider.Nodes("R").Expanded = True
        
    End With
    GetDepend = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Public Sub ShowCard(FrmMain As Form, ByVal str单据号 As String, ByVal int编辑状态 As Integer, _
        Optional int记录状态 As Integer = 1, Optional BlnSuccess As Boolean = False)
    mblnSave = False
    mblnSuccess = False
    mstr单据号 = str单据号
    mint编辑状态 = int编辑状态
    mint记录状态 = int记录状态
    mblnSuccess = BlnSuccess
    mblnChange = False
    mintParallelRecord = 1
    mstrPrivs = GetPrivFunc(glngSys, 1320)
    
    mintUnit = GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\药品付款事务", "药品单位", "0")
    If Not GetDepend Then Exit Sub
    
    If mint编辑状态 = 1 Then
        mstr单据号 = NextNo(31)
        TxtNo = mstr单据号
        
    ElseIf mint编辑状态 = 2 Then
'        mblnEdit = True
    ElseIf mint编辑状态 = 3 Then
        'mblnEdit = False
        Cmd核销.Caption = "审核(&V)"
    ElseIf mint编辑状态 = 4 Then
        'mblnEdit = False
        Cmd核销.Caption = "打印(&P)"
        If InStr(mstrPrivs, "付款通知单打印") = 0 Then
            Cmd核销.Visible = False
        Else
            Cmd核销.Visible = True
        End If
    End If
    
    Me.Show vbModal, FrmMain
    BlnSuccess = mblnSuccess
    str单据号 = mstr单据号
    
End Sub


Private Sub CmdHelp_Click()
    Call ShowHelp(App.ProductName, Me.hWnd, Me.Name)
End Sub

Private Sub Cmd按选择付款_Click()
    Dim Array发票号() As String
    Dim Str发票号 As String
    Dim curPayment As Currency
    Dim intRow As Integer
    Dim strTemp As String
    Dim rs结算方式 As New Recordset
    
    Str发票号 = ""
    
'    With mshImprest
'        curImprest = 0
'        For introw = 1 To .Rows - 1
'            If .TextMatrix(introw, 0) = "√" Then
'                curImprest = curImprest + .TextMatrix(introw, 2)
'            End If
'        Next
'    End With
    
    curPayment = CurLastMoney
    
    With mshPaymentList
        .ClearMsf
        .Cols = 3
        .rows = 3

        .TextMatrix(0, 0) = "付款方式"
        .TextMatrix(0, 1) = "付款金额"
        .TextMatrix(0, 2) = "结算号码"
    
        With rs结算方式
            gstrSQL = "Select * From 结算方式应用 Where 应用场合='付药款' Order by 缺省标志 desc"
            Call OpenRecordset(rs结算方式, "结算方式应用")
            
            If .EOF Then
                MsgBox "结算方式应用信息不全！", vbInformation, gstrSysName
                Exit Sub
            End If
            mshPaymentList.Clear
            For intRow = 1 To .RecordCount
                mshPaymentList.AddItem !结算方式
                .MoveNext
            Next
            mshPaymentList.ListIndex = 0
            .Close
        End With
            
        .TextMatrix(1, 0) = mshPaymentList.CboText
        .TextMatrix(1, 1) = GetFormat(curPayment, 2)
    End With
    
    Cmd按选择付款.Enabled = False
    Cmd核销.Enabled = True
    mshPaymentList.Active = True
    
    '统计入库单ID
    With mshPurchaseList
        For intRow = 1 To .rows - 2
            If .TextMatrix(intRow, mconintCol标志) <> "" Then
                strTemp = "'" & String(8 - Len(.TextMatrix(intRow, GetCol(mshPurchaseList, "发票号"))), "0") & .TextMatrix(intRow, GetCol(mshPurchaseList, "发票号")) & "'"
                
                If Str发票号 = "" Then
                    Str发票号 = strTemp
                Else
                    If InStr(1, Str发票号, strTemp) = 0 Then
                        Str发票号 = Str发票号 & "," & strTemp
                    End If
                End If
            End If
        Next
    End With
    Array发票号 = Split(Str发票号, ",")
    Me.txt附件数 = UBound(Array发票号) + 1
End Sub

Private Sub Cmd供应商_Click()
    tvwProvider.Visible = tvwProvider.Visible Xor True
    If tvwProvider.Visible Then
        tvwProvider.Top = Txt供药单位.Top + Txt供药单位.Height
        tvwProvider.SetFocus
    End If
    Cmd核销.Enabled = False
End Sub

Private Function SaveVerify() As Boolean
    Dim intRow As Integer
    Dim NO_IN As String
    Dim 付款金额_IN As Double
    Dim 单位ID_IN As Long
    Dim 审核人_IN As String
    
    SaveVerify = False
    
    NO_IN = TxtNo
    单位ID_IN = Txt供药单位.Tag
    审核人_IN = UserInfo.用户姓名
    付款金额_IN = 0
    On Error GoTo errHandle:
    
    With mshPaymentList
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, 0) <> "" And IIf(.TextMatrix(intRow, 1) = "", 0, .TextMatrix(intRow, 1)) <> 0 Then
                付款金额_IN = 付款金额_IN + Val(.TextMatrix(intRow, 1))
            End If
        Next
    End With
    'zl_药品付款管理_VERIFY( /*NO_IN*/, /*单位ID_IN*/, /*付款金额_IN*/, /*审核人_IN*/ );
    gstrSQL = "zl_药品付款管理_VERIFY('" & NO_IN & "'," & 单位ID_IN & "," & 付款金额_IN _
        & ",'" & 审核人_IN & "')"
    Call SQLTest(App.Title, Me.Caption, gstrSQL)
    gcnOracle.Execute gstrSQL, , adCmdStoredProc
    Call SQLTest
    
    
    SaveVerify = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Function


Private Function SaveCard() As Boolean
    Dim IntTotalRows As Integer
    Dim intRow As Integer
    Dim intLop As Integer
    Dim Cur余额 As Currency
    Dim curImprest As Currency
    Dim NO_IN As String
    Dim 序号_IN As Integer
    Dim 预付款_IN As Integer
    Dim 单位ID_IN As Long
    Dim 金额_IN As Double
    Dim 结算方式_IN As String
    Dim 结算号码_IN As String
    Dim 填制人_IN As String
    Dim 填制日期_IN As String
    Dim 付款序号_IN As Long
    Dim 摘要_IN As String
    
    SaveCard = False
    With mshPaymentList
        For intRow = 1 To .rows - 1
            Cur余额 = Cur余额 + Val(.TextMatrix(intRow, 1))
        Next
    End With
    
    
    If Cur余额 <> CurLastMoney Then
        MsgBox "付款金额不平,请检查付款金额与入库单发票金额和预付款之差是否相同!", vbInformation, gstrSysName
        mshPaymentList.SetFocus
        Exit Function
    End If
    
    IntTotalRows = IIf(LTrim(RTrim(mshPaymentList.TextMatrix(1, 1))) = "", 0, 1)
    If IntTotalRows < 1 Then Exit Function
    IntTotalRows = IIf(LTrim(RTrim(mshPaymentList.TextMatrix(mshPaymentList.rows - 1, 1))) = "", mshPaymentList.rows - 2, mshPaymentList.rows - 1)
    If IntTotalRows < 1 Then Exit Function
    If CheckData(IntTotalRows) = False Then Exit Function
    
    NO_IN = TxtNo
    预付款_IN = 0
    单位ID_IN = Txt供药单位.Tag
    填制人_IN = UserInfo.用户姓名
    填制日期_IN = Format(zldatabase.Currentdate, "yyyy-mm-dd hh:mm:ss")
    摘要_IN = Txt付款说明
    付款序号_IN = zldatabase.GetNextId("药品付款记录")
    
    On Error GoTo errHandle:
    
    '开始事务
    gcnOracle.BeginTrans
    
    If mint编辑状态 = 2 Then
        gstrSQL = "zl_药品付款管理_delete('" & NO_IN & "')"
            
        Call SQLTest(App.Title, Me.Caption, gstrSQL)
        gcnOracle.Execute gstrSQL, , adCmdStoredProc
        Call SQLTest
    End If
        
    '循环保存每行数据
    With mshPaymentList
        'zl_药品付款管理_INSERT( /*NO_IN*/, /*序号_IN*/, /*预付款_IN*/, /*单位ID_IN*/,
            '/*金额_IN*/, /*结算方式_IN*/, /*结算号码_IN*/, /*填制人_IN*/, /*填制日期_IN*/,
            '/*付款序号_IN*/, /*摘要_IN*/ );
        For intRow = 1 To IntTotalRows
            'Modified by zyb 2002-11-08
            'If Val(.TextMatrix(intRow, 1)) > 0 Then
            If Val(.TextMatrix(intRow, 1)) <> 0 Then
                序号_IN = intRow
                金额_IN = .TextMatrix(intRow, 1)
                结算方式_IN = .TextMatrix(intRow, 0)
                结算号码_IN = .TextMatrix(intRow, 2)
                gstrSQL = "zl_药品付款管理_INSERT('" & NO_IN & "'," & 序号_IN & "," & 预付款_IN & "," & 单位ID_IN _
                    & "," & 金额_IN & ",'" & 结算方式_IN & "','" & 结算号码_IN & "','" & 填制人_IN & "',to_date('" _
                    & 填制日期_IN & "','yyyy-mm-dd HH24:MI:SS')," & 付款序号_IN & ",'" & 摘要_IN & "')"
                Call SQLTest(App.Title, Me.Caption, gstrSQL)
                gcnOracle.Execute gstrSQL, , adCmdStoredProc
                Call SQLTest
            End If
        Next
    End With
                        
    '对应采购清单
    With mshPurchaseList
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, mconintCol标志) <> "" Then
                gstrSQL = "Update 药品应付记录 Set 付款序号=" & 付款序号_IN & " where 收发id=" & .RowData(intRow)
                Call SQLTest(App.Title, Me.Caption, gstrSQL)
                gcnOracle.Execute gstrSQL
                Call SQLTest
                
            End If
        Next
    End With
    '保存预副款
    With mshImprest
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                gstrSQL = "Update 药品付款记录 Set 付款序号=" & 付款序号_IN & " where id=" & .RowData(intRow)
                Call SQLTest(App.Title, Me.Caption, gstrSQL)
                gcnOracle.Execute gstrSQL
                Call SQLTest
                
            End If
        Next
    End With
    '提交事务
    gcnOracle.CommitTrans
    SaveCard = True
    Exit Function
errHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Cmd核销_Click()
    Dim BlnSuccess As Boolean
    
    If mint编辑状态 = 4 Then    '查看
        '打印
        FrmBillPrint.ShowME Me, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1320_1", "zl8_bill_1320_1"), mint记录状态, 0, 1320, "药品付款单", TxtNo.Text
        '退出
        Unload Me
        Exit Sub
    End If
    
    If mint编辑状态 = 3 Then        '审核
        If SaveVerify = True Then
            mblnChange = False
            mblnSave = False
            mblnSuccess = True

            If GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\药品付款事务", "审核打印", "0") = "1" Then
                '打印
                If InStr(mstrPrivs, "付款通知单打印") <> 0 Then
                    ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1320_1", "zl8_bill_1320_1"), Me, "单据编号=" & TxtNo.Text, "记录状态=" & mint记录状态, 2
                End If
            End If
            Unload Me
        End If
        Exit Sub
    End If
            
    BlnSuccess = SaveCard
        
    If BlnSuccess = True Then
        mblnChange = False
        mblnSave = False
        mblnSuccess = True
        If GetSetting("ZLSOFT", "私有模块\" & gstrDbUser & "\" & App.ProductName & "\药品付款事务", "存盘打印", "0") = "1" Then
            '打印
            If InStr(mstrPrivs, "付款通知单打印") <> 0 Then
                ReportOpen gcnOracle, glngSys, IIf(glngSys \ 100 = 1, "zl1_bill_1320_1", "zl8_bill_1320_1"), Me, "单据编号=" & TxtNo.Text, "记录状态=" & mint记录状态, 2
            End If
        End If
        If mint编辑状态 = 2 Then   '修改
            Unload Me
            Exit Sub
        End If
    Else
        Exit Sub
    End If
    
    mstr单据号 = NextNo(31)
    TxtNo = mstr单据号
    mblnSave = False
'    mblnEdit = True
    
    initGrid
    Txt付款金额 = ""
    Lbl合计 = ""
    Txt供药单位 = ""
    Txt供药单位.Tag = 0
    tvwProvider.Tag = 0
    txt单位名称 = ""
    txt电话地址 = ""
    
    txt开户行 = ""
    txt税务号 = ""
    txt银行帐号 = ""
    Txt付款说明 = ""
    Txt发票号 = ""
    Txt单号 = ""
    Cmd核销.Enabled = False
    
End Sub

Private Sub Cmd清除_Click()
    Dim IntChk As Integer
    For IntChk = 1 To mshPurchaseList.rows - 2
        mshPurchaseList.TextMatrix(IntChk, 0) = ""
    Next
    For IntChk = 1 To mshImprest.rows - 2
        If mshImprest.TextMatrix(IntChk, 0) <> "" Then
            mshImprest.TextMatrix(IntChk, 0) = ""
        End If
    Next
    
    BanlanceMoney
End Sub

Private Sub Cmd取消核销_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub Cmd全选_Click()
    Dim IntChk As Integer
    Cmd清除_Click
    For IntChk = 1 To mshPurchaseList.rows - 2
        mshPurchaseList.Row = IntChk
        mshPurchaseList_KeyDown vbKeySpace, 0
    Next
End Sub


Private Sub Form_Activate()
    mblnChange = False
    Select Case mintParallelRecord
        Case 1
            '正常
        Case 2
            '单据已被删除
            MsgBox "该单据已被删除，请检查！", vbOKOnly, gstrSysName
            Unload Me
            Exit Sub
        Case 3
            '修改的单据已被审核
            MsgBox "该单据已被其他人审核，请检查！", vbOKOnly, gstrSysName
            Unload Me
            Exit Sub
    End Select
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then zlcommfun.PressKey (vbKeyTab)
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = 13 Then KeyAscii = 0
End Sub

Private Sub initGrid()
    Dim IntCol As Integer
    
    With mshPurchaseList
        .Clear
        .rows = 2
        .Cols = mconIntColS
        .TextMatrix(0, mconintCol标志) = "标志"
        .TextMatrix(0, mconintcol发票号) = "发票号"
        .TextMatrix(0, mconintcol入库单号) = "入库单号"
        .TextMatrix(0, mconintcol药品信息) = "药品信息"
        .TextMatrix(0, mconIntCol规格) = "规格"
        .TextMatrix(0, mconIntCol单位) = "单位"
        .TextMatrix(0, mconintcol发票金额) = "发票金额"
        .TextMatrix(0, mconIntCol数量) = "数量"
        .TextMatrix(0, mconIntCol采购价) = "采购价"
        .TextMatrix(0, mconintcol批发价) = "批发价"
        .TextMatrix(0, mconintcol批发金额) = "批发金额"
        .TextMatrix(0, mconIntCol售价) = "售价"
        .TextMatrix(0, mconIntCol售价金额) = "售价金额"
        .TextMatrix(0, mconIntCol产地) = "产地"
        .TextMatrix(0, mconIntCol批号) = "批号"
        .TextMatrix(0, mconintcol入库日期) = "入库日期"
        
        .ColAlignment(mconintcol发票号) = flexAlignLeftCenter
        .ColAlignment(mconintcol入库单号) = flexAlignLeftCenter
        .ColAlignment(mconIntCol单位) = flexAlignCenterCenter
        .ColAlignment(mconintcol发票金额) = flexAlignRightCenter
        .ColAlignment(mconIntCol数量) = flexAlignRightCenter
        .ColAlignment(mconIntCol采购价) = flexAlignRightCenter
        .ColAlignment(mconintcol批发价) = flexAlignRightCenter
        .ColAlignment(mconintcol批发金额) = flexAlignRightCenter
        .ColAlignment(mconIntCol售价) = flexAlignRightCenter
        .ColAlignment(mconIntCol售价金额) = flexAlignRightCenter
        
        '列头中间对齐
        For IntCol = 0 To .Cols - 1
            .ColAlignmentFixed(IntCol) = flexAlignCenterCenter
        Next
        
        .ColWidth(mconintCol标志) = 450
        .ColWidth(mconintcol发票号) = 800
        .ColWidth(mconintcol入库单号) = 800
        .ColWidth(mconintcol药品信息) = 2000
        .ColWidth(mconIntCol规格) = 800
        .ColWidth(mconIntCol单位) = 450
        .ColWidth(mconintcol发票金额) = 1000
        .ColWidth(mconIntCol数量) = 1000
        .ColWidth(mconIntCol采购价) = 1000
        .ColWidth(mconintcol批发价) = 1000
        .ColWidth(mconintcol批发金额) = 1000
        .ColWidth(mconIntCol售价) = 1000
        .ColWidth(mconIntCol售价金额) = 1000
        .ColWidth(mconIntCol产地) = 2000
        .ColWidth(mconIntCol批号) = 800
        .ColWidth(mconintcol入库日期) = 1000
    End With
    
    With mshPaymentList
        .ClearMsf
        .Cols = 3
        .rows = 3
        .TextMatrix(0, 0) = "付款方式"
        .TextMatrix(0, 1) = "付款金额"
        .TextMatrix(0, 2) = "结算号码"
        .ColWidth(0) = 800
        .ColWidth(1) = 1200
        .ColWidth(2) = 1600
        
        .ColData(0) = 3
        .ColData(1) = 4
        .ColData(2) = 4
    End With
    
    With mshImprest
        .ClearMsf
        
        .Cols = 4
        .rows = 4
        .Active = True
        
        .TextMatrix(0, 0) = "选择"
        .TextMatrix(0, 1) = "结算方式"
        .TextMatrix(0, 2) = "结算金额"
        .TextMatrix(0, 3) = "结算号码"
        
        .ColWidth(0) = 500
        .ColWidth(1) = 800
        .ColWidth(2) = 1200
        .ColWidth(3) = 1600
        
        .ColData(0) = -1
        .ColData(1) = 5
        .ColData(2) = 5
        .ColData(3) = 5
        .LocateCol = 0
    End With
End Sub

Private Sub initCard()
    initGrid
    On Error GoTo errHandle
    If mint编辑状态 = 1 Then
        Txt发票号.Enabled = True
        Txt单号.Enabled = True
        Txt填制人 = UserInfo.用户姓名
        Txt填制日期 = Format(zldatabase.Currentdate, "yyyy-MM-dd")
        Exit Sub
    Else
        Dim rsPayment As New Recordset
        Dim intRecord As Integer
        Dim intLop As Integer
        
        gstrSQL = "SELECT a.序号, a.金额, a.结算方式, a.结算号码, a.摘要,a.付款序号,a.填制人,a.填制日期,a.审核人,a.审核日期, " _
                & " b.名称, b.id,b.地址 || b.电话 as 电话地址,开户银行,帐号,税务登记号 " _
                & " FROM 药品付款记录 a, 药品供应商 b " _
                & "Where a.单位id = b.ID " _
                & "  AND no = '" & mstr单据号 _
                & "' AND 记录状态 = " & mint记录状态 _
               & " order by a.序号 "
        Call SQLTest(App.Title, Me.Caption, gstrSQL)
        Set rsPayment = zldatabase.OpenSQLRecord(gstrSQL, "initCard")
        Call SQLTest
        
        If Not rsPayment.EOF Then
            intRecord = rsPayment.RecordCount
            rsPayment.MoveFirst
            Txt供药单位.Text = rsPayment!名称
            Txt供药单位.Tag = rsPayment!Id
            Txt付款说明 = IIf(IsNull(rsPayment!摘要), "", rsPayment!摘要)
            txt附件数 = Get附件数(IIf(IsNull(rsPayment!付款序号), 0, rsPayment!付款序号))
            Txt付款说明.Tag = IIf(IsNull(rsPayment!付款序号), 0, rsPayment!付款序号)
            Txt填制人 = rsPayment!填制人
            If mint编辑状态 = 2 Then
                Txt填制人 = UserInfo.用户姓名
            End If
            Txt填制日期 = Format(rsPayment!填制日期, "yyyy-mm-dd hh:mm:ss")
            Txt审核人 = IIf(IsNull(rsPayment!审核人), "", rsPayment!审核人)
            Txt审核日期 = IIf(IsNull(rsPayment!审核日期), "", Format(rsPayment!审核日期, "yyyy-mm-dd hh:mm:ss"))
                        
            tvwProvider.Tag = "1"
            txt单位名称.Caption = rsPayment!名称
            txt电话地址 = IIf(IsNull(rsPayment!电话地址), "", rsPayment!电话地址)
            
            txt开户行 = IIf(IsNull(rsPayment!开户银行), "", rsPayment!开户银行)
            txt税务号 = IIf(IsNull(rsPayment!税务登记号), "", rsPayment!税务登记号)
            txt银行帐号 = IIf(IsNull(rsPayment!帐号), "", rsPayment!帐号)
            
            If (mint编辑状态 = 2 Or mint编辑状态 = 3) And Txt审核人 <> "" Then
                mintParallelRecord = 3
                Exit Sub
            End If
            
            RefreshPurchaseList
            RefreshImprest
            BanlanceMoney
            
            With mshPaymentList
                For intLop = 1 To intRecord
                    .TextMatrix(intLop, 0) = IIf(IsNull(rsPayment!结算方式), "", rsPayment!结算方式)
                    .TextMatrix(intLop, 1) = rsPayment!金额
                    .TextMatrix(intLop, 2) = IIf(IsNull(rsPayment!结算号码), "", rsPayment!结算号码)
                    If intLop = .rows - 1 Then .rows = .rows + 1
                    rsPayment.MoveNext
                Next
            End With
            
            
            Cmd按选择付款.Enabled = False
            mshPaymentList.Active = False
            If mint编辑状态 = 3 Or mint编辑状态 = 4 Then
                Txt供药单位.Enabled = False
                Cmd供应商.Enabled = False
                Cmd全选.Enabled = False
                Cmd清除.Enabled = False
                Cmd核销.Enabled = True
                mshImprest.Active = False
                Txt付款说明.Enabled = False
            Else
                Txt供药单位.Enabled = True
                Cmd供应商.Enabled = True
                Cmd全选.Enabled = True
                Cmd清除.Enabled = True
                Cmd核销.Enabled = False
            End If
            
        Else
            mintParallelRecord = 2
            Exit Sub
        End If
        
        'LockUserCons
    End If
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    
    CurLastMoney = 0
    TxtNo = mstr单据号
    
    Me.Txt供药单位.Tag = 0
    tvwProvider.Tag = "0" '如果为1,则表示已选择;否则为未选择
    
    initCard
    RestoreWinState Me
    mshPurchaseList.Enabled = True
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    
'    If Me.Height < 6135 Then
'        Me.Height = 6135
'    End If
'    If Me.Width < 11760 Then
'        Me.Width = 11760
'    End If
    
    With Fra2
        '.Top = (Me.ScaleHeight - .Height - Fra3.Height) / 2 + 30
        '.Top = (Me.ScaleHeight - .Height - Fra3.Height) + 30
        .Left = Me.ScaleWidth - .Width + 20
        .Height = Me.ScaleHeight - Fra3.Height
    End With
    
    With Lbl标题
        .Left = Fra2.Width / 2 - .Width / 2
    End With
    
    With Fra3
        .Top = Fra2.Top + Fra2.Height - 120
        .Left = Fra2.Left
    End With
    
    With picdown
        .Top = Fra2.Height - .Height - 50
    End With
    
    With mshPaymentList
        .Height = picdown.Top - .Top - 50
    End With
        
    
    With Fra4
        .Top = Me.ScaleHeight - .Height + 20
        .Width = Fra2.Left
    End With
    
    With Fra1
        .Width = Fra2.Left
        .Height = Me.ScaleHeight - Fra4.Height + 220
    End With
    
    With mshImprest
        .Top = Fra1.Height - 1900
        .Height = 1800
        .Left = mshPurchaseList.Left   ' + 50
        .Width = Fra1.Width - 200
    End With
    
    With Label1
        .Left = Lbl清单.Left
        .Top = mshImprest.Top - .Height - 100
    End With
    
    With mshPurchaseList
        .Height = Label1.Top - .Top - 100
        .Width = Fra1.Width - 50
    End With
    
    With Cmd按选择付款
        .Left = Fra4.Width - .Width - 250
    End With
    
    With Txt付款金额
        .Left = Fra1.Width - .Width - 100
    End With
    
    With Lbl付款金额
        .Left = Txt付款金额.Left - .Width - 100
    End With
    
    With Lbl合计
        .Left = Lbl付款金额.Left - .Width - 100
    End With
    
    With Lbl累计合计
        .Left = Lbl合计.Left - .Width   '- 100
    End With
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    SaveWinState Me
End Sub

'取指定列头的列位置
Private Function GetCol(mshFlex As MSHFlexGrid, ByVal ColName As String) As Integer
    Dim i As Integer
    
    GetCol = -1
    With mshFlex
        For i = 0 To .Cols - 1
            If .TextMatrix(0, i) = ColName Then
                GetCol = i
                Exit Function
            End If
        Next
        
    End With
End Function


Private Function BanlanceMoney()
    Dim intRow As Integer
    Dim cur累计合计 As Currency
    Dim curImprest As Currency
    Dim IntCol As Integer
    
    IntCol = GetCol(mshPurchaseList, "发票金额")
    
    cur累计合计 = 0
    For intRow = 1 To mshPurchaseList.rows - 1
        If mshPurchaseList.TextMatrix(intRow, 0) <> "" Then
            cur累计合计 = cur累计合计 + Val(mshPurchaseList.TextMatrix(intRow, IntCol))
        End If
    Next
    
    If cur累计合计 <> 0 Then
        Txt付款金额 = "[￥" & GetFormat(cur累计合计, 2) & "]"
    Else
        Txt付款金额 = ""
        If mint编辑状态 = 1 Then
            With mshPaymentList
                .ClearMsf
                .Cols = 3
                .rows = 3
                .TextMatrix(0, 0) = "付款方式"
                .TextMatrix(0, 1) = "付款金额"
                .TextMatrix(0, 2) = "结算号码"
            End With
        End If
    End If
    
    curImprest = 0
    With mshImprest
        For intRow = 1 To .rows - 1
            If .TextMatrix(intRow, 0) <> "" Then
                curImprest = curImprest + .TextMatrix(intRow, 2)
            End If
        Next
        If curImprest > cur累计合计 And cur累计合计 > 0 Then
            Cmd按选择付款.Enabled = False
            Cmd核销.Enabled = False
            If mintpurchaseclick = False Then
                MsgBox "对不起,当前选择的预付款金额大于了应付款金额，请重新选择预付款！", vbOKOnly, gstrSysName
                mshImprest.SetFocus
            End If
            Exit Function
        End If
    End With
    
    If CurLastMoney <> (cur累计合计 - curImprest) And cur累计合计 <> 0 Then
        Cmd按选择付款.Enabled = True
        Cmd核销.Enabled = False
        mshPaymentList.Active = False
    Else
        Cmd按选择付款.Enabled = False
        Cmd核销.Enabled = False
    End If
    CurLastMoney = cur累计合计 - curImprest
End Function


Private Sub Label2_Click()

End Sub

Private Sub mshImprest_DblClick(Cancel As Boolean)
    If mint编辑状态 > 2 Then Exit Sub
    If mshImprest.TextMatrix(mshImprest.Row, 1) = "" Then
        Cancel = True
        Exit Sub
    End If
    With mshImprest
        If .TextMatrix(.Row, 0) = "" Then
            .TextMatrix(.Row, 0) = "√"
        Else
            .TextMatrix(.Row, 0) = ""
        End If
        Cancel = True
        BanlanceMoney
    End With
    
End Sub

Private Sub mshPaymentList_AfterDeleteRow()
    Dim Cur余额 As Currency
    Dim intLop As Integer
    
    Cur余额 = 0
    
    For intLop = 1 To mshPaymentList.rows - 1
        If intLop <> mshPaymentList.Row Then
            Cur余额 = Cur余额 + Val(mshPaymentList.TextMatrix(intLop, 1))
        End If
    Next
    Cur余额 = CurLastMoney - Cur余额
    
    If Cur余额 <> 0 Then
        mshPaymentList.TextMatrix(mshPaymentList.Row, 1) = Format(Cur余额, "#####0.00;-#####0.00; ;")
        mshPaymentList.TextMatrix(mshPaymentList.Row, 0) = mshPaymentList.CboText
    End If
End Sub

Private Sub mshPaymentList_cboClick(ListIndex As Long)
    With mshPaymentList
        If .Col <> 0 Then Exit Sub
        .TextMatrix(.Row, .Col) = .CboText
    End With
End Sub

Private Sub mshPaymentList_cboKeyDown(KeyCode As Integer, Shift As Integer)
    With mshPaymentList
        .TextMatrix(.Row, .Col) = .CboText
    End With
End Sub

Private Sub mshPaymentList_EnterCell(Row As Long, Col As Long)
    With mshPaymentList
    Select Case Col
        Case 1
            .TxtCheck = True
            .MaxLength = 16
            .TextMask = ".1234567890"
        Case 2
            .TxtCheck = True
            .MaxLength = 10
    End Select
    End With
    
End Sub

Private Sub mshPaymentList_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim intLop As Integer
    
    If mshPaymentList.Col = 2 Then
        If KeyCode <> vbKeyReturn Then
            mshPaymentList.ColData(2) = 4
            mshPaymentList.TxtCheck = False
        Else
            mshPaymentList.ColData(2) = 0
            mshPaymentList.TxtCheck = True
            mshPaymentList.TextLen = 10
        End If
    End If
    If mshPaymentList.Col = 1 _
            And mshPaymentList.Row = mshPaymentList.rows - 1 _
            And KeyCode = vbKeyReturn Then
        Txt付款说明.SetFocus
    End If
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If mshPaymentList.TxtVisible = False Then Exit Sub
'    If Chk预付款.Value = 1 Then Exit Sub
    Dim Cur余额 As Currency
    Dim curImprest As Currency
    
    
    If mshPaymentList.Col = 1 Then
        Cur余额 = 0
        For intLop = 1 To mshPaymentList.rows - 1
            If intLop <> mshPaymentList.Row Then
                Cur余额 = Cur余额 + Val(mshPaymentList.TextMatrix(intLop, 1))
            End If
        Next
        
        Cur余额 = CurLastMoney - Cur余额
        
        
        
        If Val(mshPaymentList.Text) = 0 And Cur余额 > 0 Then
            MsgBox "付款金额不能为空!", vbInformation, gstrSysName
            Cancel = True
            mshPaymentList.TxtSetFocus
            Exit Sub
        End If
        If Not IsNumeric(mshPaymentList.Text) And Trim(mshPaymentList.Text) <> "" Then
            MsgBox "付款金额中含有非法字符!", vbInformation, gstrSysName
            Cancel = True
            mshPaymentList.TxtSetFocus
            Exit Sub
        End If
        If Val(mshPaymentList.Text) < 0 Then
            MsgBox "付款分录金额不能为负数!", vbInformation, gstrSysName
            Cancel = True
            mshPaymentList.TxtSetFocus
            Exit Sub
        End If
        If Val(mshPaymentList.Text) >= 10 ^ 14 - 1 Then
            MsgBox "付款金额必须小于" & (10 ^ 14 - 1), vbInformation + vbOKOnly, gstrSysName
            Cancel = True
            mshPaymentList.TxtSetFocus
            Exit Sub
        End If
        If Trim(mshPaymentList.Text) = "" Then Exit Sub
        Cur余额 = Cur余额 - IIf(Trim(mshPaymentList.Text) = "", 0, mshPaymentList.Text)
        If Cur余额 < 0 Then
            MsgBox "付款金额超出总额!", vbInformation, gstrSysName
            Cancel = True
            mshPaymentList.TxtSetFocus
            Exit Sub
        End If
        If mshPaymentList.Row >= mshPaymentList.rows - 1 And Cur余额 > 0 Then
            mshPaymentList.rows = mshPaymentList.rows + 1
        End If
                
        mshPaymentList.Text = GetFormat(mshPaymentList.Text, 2)
        If Cur余额 > 0 Then
            mshPaymentList.TextMatrix(mshPaymentList.Row + 1, 1) = GetFormat(Cur余额, 2)
            mshPaymentList.TextMatrix(mshPaymentList.Row + 1, 0) = mshPaymentList.CboText
        End If
    End If
End Sub

Private Sub mshPurchaseList_DblClick()
    If mint编辑状态 > 2 Then Exit Sub
    
    If mshPurchaseList.TextMatrix(mshPurchaseList.Row, 1) = "" Then Exit Sub
    
    If mshPurchaseList.TextMatrix(mshPurchaseList.Row, 0) <> "" Then
        mshPurchaseList.TextMatrix(mshPurchaseList.Row, 0) = ""
    Else
        mshPurchaseList.TextMatrix(mshPurchaseList.Row, 0) = "√"
    End If
    mintpurchaseclick = True
    BanlanceMoney
    mintpurchaseclick = False
End Sub

Private Sub mshPurchaseList_DragDrop(Source As Control, x As Single, y As Single)
    If mshPurchaseList.Tag = "" Then Exit Sub
    If mshPurchaseList.MouseCol = 0 Then Exit Sub
    mshPurchaseList.Redraw = False
    mshPurchaseList.ColPosition(Val(mshPurchaseList.Tag)) = mshPurchaseList.MouseCol
    DoSort
    mshPurchaseList.Redraw = True
End Sub

Private Sub mshPurchaseList_GotFocus()
    tvwProvider.Visible = False
End Sub

Private Sub mshPurchaseList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then mshPurchaseList_DblClick
End Sub


Private Sub mshPurchaseList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    mshPurchaseList.Tag = ""
    If mshPurchaseList.MouseRow <> 0 Then Exit Sub
    If mshPurchaseList.MouseCol = 0 Then Exit Sub
    mshPurchaseList.Tag = Str(mshPurchaseList.MouseCol)
    mshPurchaseList.Drag 1
End Sub

Private Sub tvwProvider_DblClick()
    Dim rsProvider As New Recordset
    
    On Error GoTo errHandle
    If tvwProvider.SelectedItem.Children <> 0 Then Exit Sub
    If tvwProvider.SelectedItem.Tag = 0 Then Exit Sub
    
    Txt供药单位 = tvwProvider.SelectedItem
    Txt供药单位.Tag = Mid(tvwProvider.SelectedItem.Key, 3)
    tvwProvider.Tag = "1"
    tvwProvider.Visible = False
    
    With rsProvider
        gstrSQL = "Select 编码,名称,地址||电话 as 电话地址,开户银行,帐号,税务登记号 " _
            & " From 药品供应商  " _
            & "Where id=" & Txt供药单位.Tag
        Call SQLTest(App.Title, Me.Caption, gstrSQL)
        Set rsProvider = zldatabase.OpenSQLRecord(gstrSQL, "tvwProvider_DblClick")
        Call SQLTest
        
        If .EOF Then Exit Sub
        
        txt单位名称 = "[" & !编码 & "]" & !名称
        txt电话地址 = IIf(IsNull(!电话地址), "", !电话地址)
        txt开户行 = IIf(IsNull(!开户银行), "", !开户银行)
        txt银行帐号 = IIf(IsNull(!帐号), "", !帐号)
        txt税务号 = IIf(IsNull(!税务登记号), "", !税务登记号)
    End With

    Call RefreshPurchaseList
    Call RefreshImprest
    Call BanlanceMoney
    
'    Chk预付款.Enabled = True
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


'刷新付款清单
Private Function RefreshPurchaseList()
    Dim Rec累计合计 As New ADODB.Recordset
    Dim rsPayment As New Recordset
    Dim strLevel As String
    Dim strUnitName As String
    Dim str包装系数 As String
    Dim intLop As Integer
    
    Txt单号 = ""
    Txt发票号 = ""
    On Error GoTo errHandle
    With rsPayment
        If .State = 1 Then .Close
        
        If glngSys \ 100 = 8 Then
            strUnitName = Choose(mintUnit + 1, "b.药库单位", "b.售价单位")
            str包装系数 = Choose(mintUnit + 1, "b.药库包装", "1")
        Else
            strUnitName = Choose(mintUnit + 1, "b.药库单位", "b.门诊单位", "b.住院单位", "b.售价单位")
            str包装系数 = Choose(mintUnit + 1, "b.药库包装", "b.门诊包装", "b.住院包装", "1")
        End If
        
        
        If mint编辑状态 = 1 Then
            
            gstrSQL = "SELECT distinct a.审核日期 AS 入库日期, a.no, c.发票号, c.发票金额," _
                & "('[' || b.编码 || ']' || decode(e.名称,null,d.通用名称,e.名称)) AS 药品信息, b.规格, " _
                & "a.成本价*" & str包装系数 & " AS 采购价, b.指导批发价 * a.实际数量 AS 批发金额, " _
                & "b.指导批发价*" & str包装系数 & " as 指导批发价, a.产地, a.批号," & strUnitName & " AS 单位, a.实际数量/" & str包装系数 & "  AS 数量," _
                & "a.零售价*" & str包装系数 & " AS 单价, a.实际数量 * a.零售价 AS 售价金额, c.收发id,c.付款序号 " _
                & " FROM (SELECT * From 药品收发记录 Where 单据 = 1 AND 供药单位id =" & Txt供药单位.Tag _
                        & " AND 审核人 IS NOT NULL) a," _
                    & " 药品目录 b," _
                    & " 药品应付记录 c, " _
                    & " 药品信息 d," _
                    & " 药品别名 e " _
               & " Where c.收发id = a.ID " _
                 & " AND a.药品id = b.药品id " _
                 & " AND b.药品id = e.药品id (+) " _
                 & " AND b.药名id = d.药名id " _
                 & " AND c.发票号 IS NOT NULL " _
                 & " AND c.发票金额 <> 0 " _
                 & " AND c.付款序号 IS NULL " _
                 & " AND c.供药单位id IS NOT NULL " _
                 & " AND c.供药单位id =" & Txt供药单位.Tag _
               & " ORDER BY c.发票号, a.no "
                
        ElseIf mint编辑状态 = 2 Then
            '修改付款单
            gstrSQL = "SELECT distinct a.审核日期 AS 入库日期, a.no, c.发票号, c.发票金额," _
                & "('[' || b.编码 || ']' || decode(e.名称,null,d.通用名称,e.名称)) AS 药品信息, b.规格, " _
                & "a.成本价*" & str包装系数 & " AS 采购价, b.指导批发价 * a.实际数量 AS 批发金额, " _
                & "b.指导批发价*" & str包装系数 & " as 指导批发价, a.产地, a.批号," & strUnitName & " AS 单位, a.实际数量/" & str包装系数 & "  AS 数量," _
                & "a.零售价*" & str包装系数 & " AS 单价, a.实际数量 * a.零售价 AS 售价金额, c.收发id,c.付款序号 " _
                & " FROM (SELECT * From 药品收发记录 Where 单据 = 1 AND 供药单位id =" & Txt供药单位.Tag _
                        & " AND 审核人 IS NOT NULL) a," _
                    & " 药品目录 b," _
                    & " 药品应付记录 c, " _
                    & " 药品信息 d," _
                    & " 药品别名 e " _
               & " Where c.收发id = a.ID " _
                 & " AND a.药品id = b.药品id " _
                 & " AND b.药品id = e.药品id (+) " _
                 & " AND b.药名id = d.药名id " _
                 & " AND c.发票号 IS NOT NULL " _
                 & " AND c.发票金额 <> 0 " _
                 & " AND c.付款序号 IS NULL " _
                 & " AND c.供药单位id =" & Txt供药单位.Tag  '_
               '& " ORDER BY c.发票号, a.no "
               
             gstrSQL = gstrSQL & _
                 " union " _
                & "SELECT distinct a.审核日期 AS 入库日期, a.no, c.发票号, c.发票金额," _
                & "('[' || b.编码 || ']' || decode(e.名称,null,d.通用名称,e.名称)) AS 药品信息, b.规格, " _
                & "a.成本价*" & str包装系数 & " AS 采购价, b.指导批发价 * a.实际数量 AS 批发金额, " _
                & "b.指导批发价*" & str包装系数 & " as 指导批发价, a.产地, a.批号," & strUnitName & " AS 单位, a.实际数量/" & str包装系数 & "  AS 数量," _
                & "a.零售价*" & str包装系数 & " AS 单价, a.实际数量 * a.零售价 AS 售价金额, c.收发id,c.付款序号 " _
                & " FROM (SELECT * From 药品收发记录 Where 单据 = 1 AND 供药单位id =" & Txt供药单位.Tag _
                        & " AND 审核人 IS NOT NULL) a," _
                    & " 药品目录 b," _
                    & " 药品应付记录 c, " _
                    & " 药品信息 d," _
                    & " 药品别名 e " _
               & " Where c.收发id = a.ID " _
                 & " AND a.药品id = b.药品id " _
                 & " AND b.药品id = e.药品id (+) " _
                 & " AND b.药名id = d.药名id " _
                 & " AND c.发票号 IS NOT NULL " _
                 & " AND c.发票金额 <> 0 " _
                 & " AND c.付款序号 =" & Txt付款说明.Tag _
                 & " AND c.供药单位id =" & Txt供药单位.Tag
                    
            '   & " ORDER BY c.发票号, a.no "
        Else
            gstrSQL = "SELECT distinct a.审核日期 AS 入库日期, a.no, c.发票号, c.发票金额," _
                & "('[' || b.编码 || ']' || decode(e.名称,null,d.通用名称,e.名称)) AS 药品信息, b.规格, " _
                & "a.成本价*" & str包装系数 & " AS 采购价, b.指导批发价 * a.实际数量 AS 批发金额, " _
                & "b.指导批发价*" & str包装系数 & " as 指导批发价, a.产地, a.批号," & strUnitName & " AS 单位, a.实际数量/" & str包装系数 & "  AS 数量," _
                & "a.零售价*" & str包装系数 & " AS 单价, a.实际数量 * a.零售价 AS 售价金额, c.收发id,c.付款序号 " _
                & " FROM (SELECT * From 药品收发记录 Where 单据 = 1 AND 供药单位id =" & Txt供药单位.Tag _
                        & " AND 审核人 IS NOT NULL) a," _
                    & " 药品目录 b," _
                    & " 药品应付记录 c, " _
                    & " 药品信息 d," _
                    & " 药品别名 e " _
               & " Where c.收发id = a.ID " _
                 & " AND a.药品id = b.药品id " _
                 & " AND b.药品id = e.药品id (+) " _
                 & " AND b.药名id = d.药名id " _
                 & " AND c.发票号 IS NOT NULL " _
                 & " AND c.发票金额 <> 0 " _
                 & " AND c.付款序号 =" & Txt付款说明.Tag _
                 & " AND c.供药单位id =" & Txt供药单位.Tag _
               & " ORDER BY c.发票号, a.no "
        End If
        Call SQLTest(App.Title, Me.Caption, gstrSQL)
        Set rsPayment = zldatabase.OpenSQLRecord(gstrSQL, "RefreshPurchaseList")
        Call SQLTest
        
        If mint编辑状态 = 1 Or mint编辑状态 = 2 Then
            initGrid
        End If
        
        If .RecordCount > 0 And InStr(1, "12", mint编辑状态) <> 0 Then
            Cmd全选.Enabled = True
            Cmd清除.Enabled = True
        Else
            Cmd全选.Enabled = False
            Cmd清除.Enabled = False
        End If
        If .EOF Then Exit Function
        .MoveFirst
        For intLop = 1 To .RecordCount
            mshPurchaseList.TextMatrix(intLop, mconintCol标志) = IIf(IsNull(!付款序号), "", "√")
            mshPurchaseList.TextMatrix(intLop, mconintcol发票号) = !发票号
            mshPurchaseList.TextMatrix(intLop, mconintcol入库单号) = !No
            mshPurchaseList.TextMatrix(intLop, mconintcol药品信息) = !药品信息
            mshPurchaseList.TextMatrix(intLop, mconIntCol规格) = IIf(IsNull(!规格), "", !规格)
            mshPurchaseList.TextMatrix(intLop, mconIntCol单位) = !单位
            mshPurchaseList.TextMatrix(intLop, mconintcol发票金额) = GetFormat(!发票金额, 2)
            mshPurchaseList.TextMatrix(intLop, mconIntCol数量) = GetFormat(!数量, 3)
            
            mshPurchaseList.TextMatrix(intLop, mconIntCol采购价) = GetFormat(!采购价, 4)
            mshPurchaseList.TextMatrix(intLop, mconintcol批发价) = GetFormat(!指导批发价, 4)
            mshPurchaseList.TextMatrix(intLop, mconintcol批发金额) = GetFormat(!批发金额, 2)
            
            mshPurchaseList.TextMatrix(intLop, mconIntCol售价) = GetFormat(!单价, 4)
            mshPurchaseList.TextMatrix(intLop, mconIntCol售价金额) = GetFormat(!售价金额, 4)
            
            mshPurchaseList.TextMatrix(intLop, mconIntCol产地) = IIf(IsNull(!产地), "", !产地)
            mshPurchaseList.TextMatrix(intLop, mconIntCol批号) = IIf(IsNull(!批号), "", !批号)
            
            mshPurchaseList.TextMatrix(intLop, mconintcol入库日期) = Format(IIf(IsNull(!入库日期), "", !入库日期), "yyyy-MM-dd")
            
            mshPurchaseList.RowData(intLop) = !收发id
            If intLop = mshPurchaseList.rows - 1 Then mshPurchaseList.rows = mshPurchaseList.rows + 1
            .MoveNext
        Next
        
    End With
    
    With Rec累计合计
        If .State = 1 Then .Close
        If mint编辑状态 = 1 Then
            gstrSQL = "Select Sum(发票金额) as 合计 From 药品应付记录  " _
                   & " Where 发票号 is Not Null And 发票金额<>0 " _
                   & "   And 付款序号 is Null " _
                   & "   and 供药单位ID=" & Txt供药单位.Tag
            Call SQLTest(App.Title, Me.Caption, gstrSQL)
            Set Rec累计合计 = zldatabase.OpenSQLRecord(gstrSQL, "cmd产地_Click")
            Call SQLTest
            
        Else
            gstrSQL = "Select sum(合计) as 合计 " _
                   & "  From (" & _
                          " Select Sum(发票金额) as 合计 From 药品应付记录 Where 发票号 is Not Null And 发票金额<>0  And 供药单位ID=" & Txt供药单位.Tag & " And 付款序号 is Null" _
                        & " Union Select Sum(发票金额) as 合计 From 药品应付记录 Where 发票号 is Not Null And 发票金额<>0 And 供药单位ID=" & Txt供药单位.Tag & " And 付款序号=" & Txt付款说明.Tag & ")"
            Call SQLTest(App.Title, Me.Caption, gstrSQL)
            Set Rec累计合计 = zldatabase.OpenSQLRecord(gstrSQL, "cmd产地_Click")
            Call SQLTest
            
        End If
        Lbl合计 = GetFormat(!合计, 2)
    End With
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

'刷新预付款
Private Sub RefreshImprest()
    Dim rsImprest As New Recordset
    Dim intRow As Integer
    Dim intRecord As Integer
    
    On Error GoTo errHandle
    If mint编辑状态 = 1 Then
        gstrSQL = "select id,结算方式,结算号码,金额,付款序号 " _
                & " from 药品付款记录 " _
               & " where 单位id=" & Txt供药单位.Tag _
                & "  and 预付款=1 " _
                & "  and nvl(付款序号,0)=0 " _
                & "  and 审核日期 is not null "
                '& "  and 记录状态=1 "
        gstrSQL = gstrSQL _
               & " union all " _
               & "select id,结算方式,结算号码,金额,付款序号 " _
                & " from 药品付款记录 " _
               & " where 单位id=" & Txt供药单位.Tag _
                & "  and nvl(预付款,0)=0 " _
                & "  and nvl(付款序号,0)=0 " _
                & "  and 审核日期 is not null " _
                & "  and 记录状态=2 "
        
                
    ElseIf mint编辑状态 = 2 Then
        gstrSQL = "select id,结算方式,结算号码,金额,付款序号 " _
                & " from 药品付款记录 " _
               & " where 单位id=" & Txt供药单位.Tag _
                & "  and 预付款=1 " _
                & "  and nvl(付款序号,0)=0 " _
                & "  and 审核日期 is not null "
                '& "  and 记录状态=1 "
        gstrSQL = gstrSQL _
            & " union all " _
               & "select id,结算方式,结算号码,金额,付款序号 " _
                & " from 药品付款记录 " _
               & " where 单位id=" & Txt供药单位.Tag _
                & "  and nvl(预付款,0)=0 " _
                & "  and nvl(付款序号,0)=0 " _
                & "  and 审核日期 is not null " _
                & "  and 记录状态=2 " _
            & " union " _
            & "select id,结算方式,结算号码,金额 ,付款序号" _
                & " from 药品付款记录 " _
               & " where 单位id=" & Txt供药单位.Tag _
                & "  and 预付款=1 " _
                & "  and nvl(付款序号,0)=" & Txt付款说明.Tag _
                & "  and 审核日期 is not null " _
            & "union  select id,结算方式,结算号码,金额 ,付款序号" _
                & " from 药品付款记录 " _
               & " where 单位id=" & Txt供药单位.Tag _
                & "  and nvl(预付款,0)=0 " _
                & "  and nvl(付款序号,0)=" & Txt付款说明.Tag _
                & "  and 审核日期 is not null " _
                & "  and (记录状态=2) "
    Else
        gstrSQL = "select id,结算方式,结算号码,金额,付款序号 " _
                & " from 药品付款记录 " _
               & " where 单位id=" & Txt供药单位.Tag _
                & "  and 预付款=1 " _
                & "  and nvl(付款序号,0)=" & Txt付款说明.Tag _
                & "  and 审核日期 is not null " _
                & "union  select id,结算方式,结算号码,金额 ,付款序号" _
                & " from 药品付款记录 " _
               & " where 单位id=" & Txt供药单位.Tag _
                & "  and nvl(预付款,0)=0 " _
                & "  and nvl(付款序号,0)=" & Txt付款说明.Tag _
                & "  and 审核日期 is not null " _
                & "  and (记录状态=2) "
                
                '& "  and (记录状态=1 or 记录状态=3)  "
                

    End If
    Call SQLTest(App.Title, Me.Caption, gstrSQL)
    Set rsImprest = zldatabase.OpenSQLRecord(gstrSQL, "RefreshImprest")
    Call SQLTest
    
    If rsImprest.EOF Then Exit Sub
    intRecord = rsImprest.RecordCount
    rsImprest.MoveFirst
    With mshImprest
        For intRow = 1 To intRecord
            .TextMatrix(intRow, 0) = IIf(IIf(IsNull(rsImprest!付款序号), 0, rsImprest!付款序号) > 0, "√", "")
            .TextMatrix(intRow, 1) = rsImprest!结算方式
            .TextMatrix(intRow, 2) = rsImprest!金额
            .TextMatrix(intRow, 3) = IIf(IsNull(rsImprest!结算号码), "", rsImprest!结算号码)
            .RowData(intRow) = rsImprest!Id
            If intRow = .rows - 1 Then .rows = .rows + 1
            rsImprest.MoveNext
        Next
        rsImprest.Close
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub tvwProvider_LostFocus()
'    tvwProvider.Visible = False
End Sub

Private Sub TxtNO_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey (vbKeyTab)
End Sub

Private Sub Txt单号_GotFocus()
    With Txt单号
        .SelStart = 0
        .SelLength = 100
    End With
End Sub

Private Sub Txt单号_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then mshPurchaseList.SetFocus
End Sub

Private Sub Txt单号_Validate(Cancel As Boolean)
    Call SelAccord
End Sub

Private Sub Txt发票号_GotFocus()
    With Txt发票号
        .SelStart = 0
        .SelLength = 100
    End With
End Sub

Private Sub Txt发票号_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then TxtNo.SetFocus
End Sub

Private Sub Txt发票号_Validate(Cancel As Boolean)
    Call SelAccord
End Sub

Private Sub txt付款说明_GotFocus()
    tvwProvider.Visible = False
End Sub

Private Sub txt付款说明_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then If Cmd核销.Enabled Then Cmd核销.SetFocus
End Sub

Private Sub Txt供药单位_GotFocus()
    tvwProvider.Visible = False
End Sub

Private Sub txt供药单位_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strInput As String
    Dim rec供应商 As New ADODB.Recordset
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If LTrim(RTrim(Txt供药单位)) = "" Then Exit Sub
    If InStr(1, Txt供药单位, "[") <> 0 Then
        If InStr(2, Txt供药单位, "]") <> 0 Then
            strInput = Mid(Txt供药单位.Text, 2, InStr(2, Txt供药单位, "]") - 2)
        Else
            strInput = Mid(Txt供药单位.Text, 2)
        End If
    Else
        strInput = Txt供药单位.Text
    End If
    
    With rec供应商
        gstrSQL = "Select ID,编码,名称,简码,地址||电话 as 电话地址,开户银行,帐号,税务登记号  From 药品供应商 Where (编码 like '" & UCase(strInput) & "%' Or 名称 like '" & UCase(strInput) & "%' Or 简码 like '" & UCase(strInput) & "%') And To_Char(撤档时间,'yyyy-MM-dd')='3000-01-01' And 末级=1"
        Call OpenRecordset(rec供应商, "药品供应商")
        
        If .EOF Then
            MsgBox "输入值无效！", vbInformation, gstrSysName
            KeyCode = 0
            Txt供药单位 = ""
            tvwProvider.Tag = "0"
            Exit Sub
        End If
        If .RecordCount > 1 Then
            Set mshProvider.Recordset = rec供应商
            SetProviderWidth Txt供药单位.Left, Txt供药单位.Top + Txt供药单位.Height + Fra1.Top
            Exit Sub
        Else
            Txt供药单位 = "[" & !编码 & "]" & !名称
            Txt供药单位.Tag = !Id
            tvwProvider.Tag = "1"
        End If
    End With
    
    txt单位名称 = rec供应商!名称
    txt电话地址 = IIf(IsNull(rec供应商!电话地址), "", rec供应商!电话地址)
    txt开户行 = IIf(IsNull(rec供应商!开户银行), "", rec供应商!开户银行)
    txt银行帐号 = IIf(IsNull(rec供应商!帐号), "", rec供应商!帐号)
    txt税务号 = IIf(IsNull(rec供应商!税务登记号), "", rec供应商!税务登记号)
    Call RefreshPurchaseList
    Call RefreshImprest
    
    Call BanlanceMoney
    
End Sub


Private Function CheckData(ByVal 总行数 As Integer) As Boolean
    Dim IntCheck As Integer
    
    CheckData = False
    With mshPaymentList
        For IntCheck = 1 To 总行数
            If Val(.TextMatrix(IntCheck, 1)) = 0 And LTrim(RTrim(.TextMatrix(IntCheck, 1))) = "" Then
                MsgBox "第" & IntCheck & "行的付款金额不能为零！", vbInformation, gstrSysName
                Exit Function
            End If
            If Not IsNumeric(.TextMatrix(IntCheck, 1)) Then
                MsgBox "第" & IntCheck & "行的付款金额中含有非法字符！", vbInformation, gstrSysName
                Exit Function
            End If
            If Val(.TextMatrix(IntCheck, 1)) > 10 ^ 11 - 1 Then
                MsgBox "第" & IntCheck & "行的付款金额超过最大值！", vbInformation, gstrSysName
                Exit Function
            End If
            If LenB(StrConv(.TextMatrix(IntCheck, 2), vbFromUnicode)) > 10 Then
                MsgBox "第" & IntCheck & "行的结算号码长度超长!(最多10个字符)", vbInformation, gstrSysName
                Exit Function
            End If
        Next
        If LenB(StrConv(Txt付款说明.Text, vbFromUnicode)) > 50 Then
            MsgBox "付款说明的长度超长!(最多为50个字符或25个汉字)", vbInformation, gstrSysName
            Txt付款说明.SetFocus
            Exit Function
        End If
        
        CheckData = True
    End With
End Function

Sub DoSort()
    
    mshPurchaseList.Col = 0
    mshPurchaseList.ColSel = mshPurchaseList.Cols - 1
    mshPurchaseList.Sort = 2 ' 标准降序
    
End Sub

Private Function Get附件数(ByVal LngPay As Long) As Long
    Dim Rec附件数 As New ADODB.Recordset
    If LngPay = 0 Then Get附件数 = 0: Exit Function
    With Rec附件数
        gstrSQL = "Select distinct 发票号 as PayCount From 药品应付记录 Where 付款序号=" & LngPay
        Call OpenRecordset(Rec附件数, "附件数")
        
        If .EOF Then
            Get附件数 = 0
        Else
            Get附件数 = .RecordCount
        End If
    End With
End Function



Private Sub mshProvider_DblClick()
    mshProvider_KeyPress 13
End Sub

Private Sub mshProvider_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Integer
    Dim sngWidth As Single
    
    With mshProvider
        If KeyCode = vbKeyRight Then
            If .ColPos(.Cols - 1) - .ColPos(.LeftCol) > .Width Then
                .LeftCol = .LeftCol + 1
                .Col = .LeftCol
                .ColSel = .Cols - 1
            ElseIf .ColPos(.Cols - 1) - .ColPos(.LeftCol) + .ColWidth(.Cols - 1) > .Width Then
                .LeftCol = .LeftCol + 1
                .Col = .LeftCol
                .ColSel = .Cols - 1
                
            End If
        ElseIf KeyCode = vbKeyLeft Then
            If .LeftCol <> 0 Then
                .LeftCol = .LeftCol - 1
                .Col = .LeftCol
                .ColSel = .Cols - 1
            End If
        ElseIf KeyCode = vbKeyHome Then
            If .LeftCol <> 0 Then
                .LeftCol = 0
                .Col = .LeftCol
                .ColSel = .Cols - 1
            End If
        ElseIf KeyCode = vbKeyEnd Then
            For i = .Cols - 1 To 0 Step -1
                sngWidth = sngWidth + .ColWidth(i)
                If sngWidth > .Width Then
                    .LeftCol = i + 1
                    .Col = .LeftCol
                    .ColSel = .Cols - 1
                    Exit For
                End If
            Next
        ElseIf KeyCode = vbKeyReturn Then
            Call mshProvider_KeyPress(13)
        End If
    End With
End Sub

Private Sub mshProvider_KeyPress(KeyAscii As Integer)
    With mshProvider
        If KeyAscii = 13 Then
            Txt供药单位.Text = "[" & .TextMatrix(.Row, 1) & "]" & .TextMatrix(.Row, 2)
            Txt供药单位.Tag = .TextMatrix(.Row, 0)
            tvwProvider.Tag = "1"
            txt单位名称.Caption = .TextMatrix(.Row, 2)
            txt电话地址 = .TextMatrix(.Row, 4)
            txt银行帐号 = .TextMatrix(.Row, 6)
            txt开户行 = .TextMatrix(.Row, 5)
            txt税务号 = .TextMatrix(.Row, 7)
            
            .Visible = False
            Call RefreshPurchaseList
            Call RefreshImprest
            Call BanlanceMoney
            mshPurchaseList.SetFocus
        End If
    End With
End Sub

Private Sub mshProvider_LostFocus()
    SaveFlexState mshProvider, Me.Caption
    If mshProvider.Visible Then mshProvider.Visible = False
End Sub


'设置供应商选择器的宽度及相关属性
Private Sub SetProviderWidth(ByVal sngLeft As Single, ByVal sngTop As Single)
    
    With mshProvider
        .Visible = True
        .Left = sngLeft
        .Top = sngTop
'        If RestoreFlexState(mshProvider, Me.Caption) = False Then
            'Select ID,名称,编码,简码,地址||电话 as 电话地址,开户银行,帐号,税务登记号
            
            .ColWidth(0) = 0
            .ColWidth(1) = 1000
            .ColWidth(2) = 2500
            .ColWidth(3) = 1000
            
            .ColWidth(4) = 1500
            .ColWidth(5) = 1500
            .ColWidth(6) = 1000
            .ColWidth(7) = 1000
            
'        End If
        
        .SetFocus
        .Row = 1
        .Col = 0
        .ColSel = .Cols - 1
    End With
End Sub

Private Sub SelAccord()
    '根据用户输入的发票号及NO，选择相应的记录
    Dim strInvoice As String, strBill As String, StrTmp As String
    Dim intDo As Integer, lngRow As Long, lngRows As Long, lngLastRow As Long
    Dim intYear As Integer, strYear As String
    Dim arrInvoice, arrBill, blnFind As Boolean
    
    lngLastRow = mshPurchaseList.Row
    strInvoice = Trim(Txt发票号)
    strBill = Trim(Txt单号)
    If strInvoice = "" And strBill = "" Then Exit Sub
    
    '检查输入格式
    arrInvoice = Split(strInvoice, "-")
    arrBill = Split(strBill, "-")
    If UBound(arrInvoice) > 1 Then
        MsgBox "输入格式不对（123或123-300），请重新输入！", vbInformation, gstrSysName
        Txt发票号.SetFocus
        Exit Sub
    End If
    If UBound(arrBill) > 1 Then
        MsgBox "输入格式不对（C0000001或C0000001-C0000020），请重新输入！", vbInformation, gstrSysName
        Txt单号.SetFocus
        Exit Sub
    End If
    
    '--如果不满八位,则按规则产生--
    Txt单号 = ""
    For intDo = 0 To UBound(arrBill)
        StrTmp = UCase(LTrim(arrBill(intDo)))
        If Len(StrTmp) < 8 Then
            intYear = Format(zldatabase.Currentdate, "YYYY") - 1990
            strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
            StrTmp = strYear & String(7 - Len(StrTmp), "0") & StrTmp
        End If
        arrBill(intDo) = StrTmp
        Txt单号 = Txt单号 & IIf(Txt单号 = "", "", "-") & StrTmp
    Next
    
    '循环选择
    Call Cmd清除_Click
    lngRows = mshPurchaseList.rows - 1
    mshPurchaseList.Redraw = False
    For lngRow = 1 To lngRows
        blnFind = False
        If strInvoice <> "" And strBill <> "" Then
            '都不为空
            StrTmp = mshPurchaseList.TextMatrix(lngRow, mconintcol发票号)
            If UBound(arrInvoice) = 1 Then
                If arrInvoice(0) <= StrTmp Then
                    blnFind = (StrTmp <= arrInvoice(1))
                End If
            Else
                blnFind = (StrTmp = arrInvoice(0))
            End If
            If blnFind Then
                blnFind = False
                StrTmp = mshPurchaseList.TextMatrix(lngRow, mconintcol入库单号)
                If UBound(arrBill) = 1 Then
                    If arrBill(0) <= StrTmp Then
                        blnFind = (StrTmp <= arrBill(1))
                    End If
                Else
                    blnFind = (arrBill(0) = StrTmp)
                End If
            End If
        ElseIf strInvoice <> "" Then
            '仅输入发票号
            StrTmp = mshPurchaseList.TextMatrix(lngRow, mconintcol发票号)
            If UBound(arrInvoice) = 1 Then
                If arrInvoice(0) <= StrTmp Then
                    blnFind = (StrTmp <= arrInvoice(1))
                End If
            Else
                blnFind = (StrTmp = arrInvoice(0))
            End If
        Else
            '仅输入单据号
            StrTmp = mshPurchaseList.TextMatrix(lngRow, mconintcol入库单号)
            If UBound(arrBill) = 1 Then
                If arrBill(0) <= StrTmp Then
                    blnFind = (StrTmp <= arrBill(1))
                End If
            Else
                blnFind = (arrBill(0) = StrTmp)
            End If
        End If
        
        '如果找到，执行双击事件
        If blnFind And Trim(mshPurchaseList.TextMatrix(lngRow, mconintCol标志)) = "" Then
            With mshPurchaseList
                .Row = lngRow
                .Col = 1
            End With
            Call mshPurchaseList_DblClick
        End If
    Next
    mshPurchaseList.Row = lngLastRow
    mshPurchaseList.Redraw = True
End Sub
