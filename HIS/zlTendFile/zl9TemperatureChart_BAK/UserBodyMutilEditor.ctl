VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.Unicode.9600.ocx"
Begin VB.UserControl UserBodyMutilEditor 
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   8220
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   12090
   KeyPreview      =   -1  'True
   LockControls    =   -1  'True
   ScaleHeight     =   8220
   ScaleWidth      =   12090
   Begin VB.PictureBox picFilter 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1725
      Left            =   4800
      ScaleHeight     =   1695
      ScaleWidth      =   2115
      TabIndex        =   46
      TabStop         =   0   'False
      Top             =   465
      Visible         =   0   'False
      Width           =   2145
      Begin VB.ListBox lstFilter 
         Appearance      =   0  'Flat
         Height          =   1290
         Left            =   -15
         Style           =   1  'Checkbox
         TabIndex        =   49
         Top             =   -15
         Width           =   2145
      End
      Begin VB.CommandButton cmdFilterCancel 
         Height          =   315
         Left            =   1530
         Picture         =   "UserBodyMutilEditor.ctx":0000
         Style           =   1  'Graphical
         TabIndex        =   48
         ToolTipText     =   "取消"
         Top             =   1320
         Width           =   450
      End
      Begin VB.CommandButton cmdFilterOK 
         Height          =   315
         Left            =   990
         Picture         =   "UserBodyMutilEditor.ctx":058A
         Style           =   1  'Graphical
         TabIndex        =   47
         ToolTipText     =   "确认"
         Top             =   1320
         Width           =   450
      End
   End
   Begin VB.PictureBox picPati 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6855
      Left            =   6900
      ScaleHeight     =   6825
      ScaleWidth      =   5145
      TabIndex        =   27
      TabStop         =   0   'False
      Top             =   435
      Visible         =   0   'False
      Width           =   5175
      Begin XtremeReportControl.ReportControl rptPati 
         Height          =   5370
         Left            =   0
         TabIndex        =   28
         TabStop         =   0   'False
         Top             =   915
         Width           =   5160
         _Version        =   589884
         _ExtentX        =   9102
         _ExtentY        =   9472
         _StockProps     =   0
         BorderStyle     =   1
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.CommandButton cmdFilterUserOk 
         Height          =   315
         Left            =   3990
         Picture         =   "UserBodyMutilEditor.ctx":0B14
         Style           =   1  'Graphical
         TabIndex        =   40
         ToolTipText     =   "确认"
         Top             =   6435
         Width           =   450
      End
      Begin VB.CommandButton cmdFilterUserCancle 
         Height          =   315
         Left            =   4530
         Picture         =   "UserBodyMutilEditor.ctx":109E
         Style           =   1  'Graphical
         TabIndex        =   39
         ToolTipText     =   "取消"
         Top             =   6435
         Width           =   450
      End
      Begin VB.CheckBox chkSwitch 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   240
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   6510
         Width           =   195
      End
      Begin VB.CheckBox chkPati 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "病人本人"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   705
         TabIndex        =   37
         Top             =   6510
         Width           =   1095
      End
      Begin VB.CheckBox chkPati 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "婴儿"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   1905
         TabIndex        =   36
         Top             =   6510
         Width           =   735
      End
      Begin VB.CheckBox chkScope 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "待入科"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   35
         Top             =   135
         Width           =   915
      End
      Begin VB.CheckBox chkScope 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "在院"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   1
         Left            =   1147
         TabIndex        =   34
         Top             =   135
         Width           =   750
      End
      Begin VB.CheckBox chkScope 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "转出"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   3
         Left            =   1980
         TabIndex        =   33
         Top             =   135
         Width           =   735
      End
      Begin VB.CheckBox chkScope 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "出院"
         ForeColor       =   &H80000008&
         Height          =   180
         Index           =   2
         Left            =   150
         TabIndex        =   32
         Top             =   510
         Width           =   795
      End
      Begin VB.TextBox txtChange 
         Alignment       =   2  'Center
         BackColor       =   &H8000000F&
         BorderStyle     =   0  'None
         Height          =   180
         IMEMode         =   3  'DISABLE
         Left            =   3600
         MaxLength       =   3
         TabIndex        =   31
         Text            =   "7"
         Top             =   120
         Width           =   285
      End
      Begin VB.Frame fraChange 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   3570
         TabIndex        =   30
         Top             =   315
         Width           =   300
      End
      Begin VB.CommandButton CmdRef 
         Caption         =   "刷新"
         Height          =   315
         Left            =   4485
         Style           =   1  'Graphical
         TabIndex        =   29
         ToolTipText     =   "取消"
         Top             =   480
         Width           =   555
      End
      Begin MSComCtl2.DTPicker dtpE 
         Height          =   300
         Index           =   0
         Left            =   3045
         TabIndex        =   41
         Top             =   480
         Width           =   1290
         _ExtentX        =   2275
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   127795203
         CurrentDate     =   37068
      End
      Begin MSComCtl2.DTPicker dtpB 
         Height          =   300
         Index           =   0
         Left            =   1380
         TabIndex        =   42
         Top             =   465
         Width           =   1275
         _ExtentX        =   2249
         _ExtentY        =   529
         _Version        =   393216
         CalendarTitleBackColor=   -2147483647
         CalendarTitleForeColor=   -2147483634
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   127795203
         CurrentDate     =   37068
      End
      Begin VB.Label lbl出院时间 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "时间"
         Height          =   180
         Left            =   945
         TabIndex        =   45
         Top             =   510
         Width           =   360
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "至"
         Height          =   180
         Index           =   0
         Left            =   2760
         TabIndex        =   44
         Top             =   525
         Width           =   180
      End
      Begin VB.Label lbl转出 
         AutoSize        =   -1  'True
         Caption         =   "显示最近    天的转出病人"
         Height          =   180
         Left            =   2820
         TabIndex        =   43
         Top             =   105
         Width           =   2160
      End
   End
   Begin VB.PictureBox picNull 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   870
      Left            =   2250
      ScaleHeight     =   840
      ScaleWidth      =   7530
      TabIndex        =   50
      Top             =   3105
      Visible         =   0   'False
      Width           =   7560
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "请点击过滤或添加病人进行数据添加"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   21.75
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   0
         TabIndex        =   51
         Top             =   120
         Width           =   6960
      End
   End
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   375
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   6255
      TabIndex        =   57
      Top             =   5460
      Width           =   6255
   End
   Begin MSComctlLib.ImageList imgRPT 
      Left            =   11400
      Top             =   480
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserBodyMutilEditor.ctx":1628
            Key             =   "woman"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserBodyMutilEditor.ctx":7E8A
            Key             =   "man"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   11400
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserBodyMutilEditor.ctx":E6EC
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox pic过滤条件 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   345
      Left            =   480
      ScaleHeight     =   345
      ScaleWidth      =   11130
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   0
      Width           =   11130
      Begin VB.ComboBox cboPati 
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   7380
         Style           =   2  'Dropdown List
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   0
         Width           =   1185
      End
      Begin VB.CommandButton cmdSift 
         Appearance      =   0  'Flat
         Height          =   260
         Left            =   6560
         Picture         =   "UserBodyMutilEditor.ctx":EA86
         Style           =   1  'Graphical
         TabIndex        =   8
         TabStop         =   0   'False
         ToolTipText     =   "选择项目(F4)"
         Top             =   10
         Width           =   270
      End
      Begin VB.TextBox txtFilter 
         Height          =   300
         Left            =   4810
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   0
         Width           =   2040
      End
      Begin MSComCtl2.DTPicker dtpDate 
         Height          =   300
         Left            =   840
         TabIndex        =   3
         Top             =   0
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   127795203
         CurrentDate     =   40624
      End
      Begin VB.CommandButton cmdAddUser 
         Caption         =   "添加病人(&A)"
         Height          =   315
         Left            =   9720
         TabIndex        =   10
         Top             =   -15
         Width           =   1245
      End
      Begin VB.ComboBox cboUnit 
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   2820
         Style           =   2  'Dropdown List
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   0
         Width           =   1425
      End
      Begin VB.CommandButton cmdFilter 
         Caption         =   "过滤(&F)"
         Height          =   315
         Left            =   8640
         TabIndex        =   9
         Top             =   0
         Width           =   885
      End
      Begin VB.Label lblPati 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "病人"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   6960
         TabIndex        =   25
         Top             =   60
         Width           =   360
      End
      Begin VB.Label lblFilter 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "条件"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   4380
         TabIndex        =   6
         Top             =   60
         Width           =   360
      End
      Begin VB.Label lblDate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "发生时间"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   60
         TabIndex        =   2
         Top             =   60
         Width           =   720
      End
      Begin VB.Label lblUnit 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "科室"
         ForeColor       =   &H00C00000&
         Height          =   180
         Left            =   2400
         TabIndex        =   4
         Top             =   60
         Width           =   360
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4500
      Left            =   0
      ScaleHeight     =   4500
      ScaleWidth      =   4845
      TabIndex        =   1
      Top             =   360
      Width           =   4845
      Begin VB.PictureBox PicLst 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   1410
         Left            =   1080
         ScaleHeight     =   1380
         ScaleWidth      =   1185
         TabIndex        =   19
         Top             =   3150
         Visible         =   0   'False
         Width           =   1215
         Begin VB.ListBox lstSelect 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Height          =   570
            Index           =   0
            ItemData        =   "UserBodyMutilEditor.ctx":EB7C
            Left            =   -10
            List            =   "UserBodyMutilEditor.ctx":EB89
            TabIndex        =   21
            Top             =   825
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtLst 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   300
            Left            =   -10
            MultiLine       =   -1  'True
            TabIndex        =   20
            Top             =   255
            Width           =   1215
         End
         Begin VB.Label lbllst 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000018&
            Caption         =   "录入："
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   30
            TabIndex        =   60
            Top             =   30
            Width           =   540
         End
         Begin VB.Label lbllst 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000018&
            Caption         =   "选择："
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   15
            TabIndex        =   59
            Top             =   615
            Width           =   540
         End
      End
      Begin VB.ComboBox cbo体温标识 
         ForeColor       =   &H00C00000&
         Height          =   300
         Left            =   3600
         Style           =   2  'Dropdown List
         TabIndex        =   26
         TabStop         =   0   'False
         Top             =   3300
         Visible         =   0   'False
         Width           =   1185
      End
      Begin VB.ListBox lstSelect 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1080
         Index           =   1
         ItemData        =   "UserBodyMutilEditor.ctx":EBA2
         Left            =   2280
         List            =   "UserBodyMutilEditor.ctx":EBAF
         Style           =   1  'Checkbox
         TabIndex        =   22
         Top             =   3120
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VB.PictureBox picDouble 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   1080
         ScaleHeight     =   240
         ScaleWidth      =   900
         TabIndex        =   15
         TabStop         =   0   'False
         Top             =   2880
         Visible         =   0   'False
         Width           =   930
         Begin VB.TextBox txtDnInput 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   525
            MaxLength       =   12
            TabIndex        =   17
            Top             =   30
            Width           =   345
         End
         Begin VB.TextBox txtUpInput 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   180
            Left            =   30
            MaxLength       =   12
            TabIndex        =   16
            Top             =   30
            Width           =   375
         End
         Begin VB.Label lblSplit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "/"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   435
            TabIndex        =   18
            Top             =   30
            Width           =   105
         End
      End
      Begin VB.ListBox lstNote 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   930
         ItemData        =   "UserBodyMutilEditor.ctx":EBC8
         Left            =   120
         List            =   "UserBodyMutilEditor.ctx":EBD2
         TabIndex        =   14
         Top             =   3120
         Visible         =   0   'False
         Width           =   975
      End
      Begin VB.PictureBox picInput 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   150
         ScaleHeight     =   225
         ScaleWidth      =   945
         TabIndex        =   11
         Top             =   2880
         Visible         =   0   'False
         Width           =   975
         Begin VB.TextBox txtInput 
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   270
            Left            =   0
            TabIndex        =   12
            Top             =   0
            Width           =   180
         End
         Begin VB.Label lblCheck 
            Alignment       =   2  'Center
            BackColor       =   &H80000018&
            Height          =   135
            Left            =   240
            TabIndex        =   13
            Top             =   0
            Visible         =   0   'False
            Width           =   135
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid VsfData 
         Height          =   2655
         Left            =   0
         TabIndex        =   23
         Top             =   480
         Width           =   4305
         _cx             =   7594
         _cy             =   4683
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
         BackColorSel    =   16764057
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   4
         FixedRows       =   3
         FixedCols       =   1
         RowHeightMin    =   255
         RowHeightMax    =   5000
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"UserBodyMutilEditor.ctx":EBE2
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
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
         AutoSizeMouse   =   0   'False
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
   Begin VB.PictureBox picHistory 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   2115
      Left            =   360
      ScaleHeight     =   2115
      ScaleWidth      =   5625
      TabIndex        =   52
      Top             =   5655
      Width           =   5625
      Begin VB.PictureBox picDay 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   345
         Left            =   105
         ScaleHeight     =   345
         ScaleWidth      =   4470
         TabIndex        =   58
         Top             =   165
         Width           =   4470
         Begin VB.TextBox txt显示天数 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   855
            MaxLength       =   2
            TabIndex        =   54
            TabStop         =   0   'False
            Text            =   "1"
            Top             =   0
            Width           =   645
         End
         Begin VB.Label lblDayInfo 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Height          =   180
            Left            =   1695
            TabIndex        =   55
            Top             =   60
            Width           =   90
         End
         Begin VB.Label lblDay 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "显示天数"
            Height          =   180
            Left            =   60
            TabIndex        =   53
            Top             =   60
            Width           =   720
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfHistory 
         Height          =   1695
         Left            =   45
         TabIndex        =   56
         Top             =   600
         Width           =   4305
         _cx             =   7594
         _cy             =   2990
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
         BackColorSel    =   16764057
         ForeColorSel    =   0
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   4
         FixedRows       =   3
         FixedCols       =   1
         RowHeightMin    =   255
         RowHeightMax    =   5000
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"UserBodyMutilEditor.ctx":EC44
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   0
         MergeCompare    =   0
         AutoResize      =   0   'False
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
         AutoSizeMouse   =   0   'False
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
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
End
Attribute VB_Name = "UserBodyMutilEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit


Private Enum PATI_COLUMN
    c_选择 = 0
    c_图标 = 1
    c_排序 = 2
    c_状态 = 3
    c_床号 = 4
    C_病人ID = 5
    c_主页ID = 6
    c_姓名 = 7
    c_年龄 = 8
    c_住院号 = 9
    c_入院日期 = 10
    c_出院日期 = 11
End Enum

Private Enum Scope
    待入科 = 0
    在院 = 1
    出院 = 2
    转出 = 3
End Enum

Private Const c文件ID As Integer = 1
Private Const c床号 As Integer = 2
Private Const c姓名 As Integer = 3
Private Const c年龄 As Integer = 4
Private Const c病人ID As Integer = 5
Private Const c主页ID As Integer = 6
Private Const c婴儿 As Integer = 7
Private Const c记录ID As Integer = 8
Private Const c护理等级 As Integer = 9
Private Const c体温标识 As Integer = 10
Private Const c出院 As Integer = 11
Private Const c日期 As Integer = 12
Private Const c时间 As Integer = 13
Private Const RootCol As Integer = 14  '固定表头列数

Private mcbrMenuBar部位 As CommandBarControl
Private mcbrToolBar As CommandBar

'---病人基本信息
Private mlng文件ID As Long
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mlngBaby As Long

Private mrsItems As New ADODB.Recordset
Private mrsCell As New ADODB.Recordset
Private mrsPati As New ADODB.Recordset
Private mrsPart As New ADODB.Recordset
Private mrsCopy As New ADODB.Recordset
Private mrsData As New ADODB.Recordset
Private mrsHistory As New ADODB.Recordset

Private mstrSQL As String
Private mfrmParent As Object
Private mlng病区ID As Long
Private mlng科室id As Long '用户选择的科室ID
Private mlng格式ID As Long '体温单格式ID
Private mstrDate As String '用户选择的时间
Private mblnInit As Boolean
Private mstrPrivs As String
Private mintBigSize As Integer '护理文件显示模式
Private mintPreDays As Integer '超期录入天数
Private mlngHours As Long    '数据补录时限
Private mstrScope As String  '参数病人显示范围
Private mintChange As Integer '参数最近转出天数
Private mdtOutEnd As String '参数出院显示终止时间
Private mdtOutBegin As String '参数出院显示开始时间
Private mblnShow As Boolean
Private mblnChage As Boolean
Private mblnNullRow As Boolean
Private mblnClearRow As Boolean
Private mblnRefreshData As Boolean
Private mbln出院 As Boolean
Private mblnSaveData As Boolean
Private mblnDateFouces As Boolean
Private mblnChkClick As Boolean
Private mstrTabHead As String ' 表头信息
Private mstrItemNo As String '项目序号信息
Private mintPatiNo As Integer '病人类型 (所有、病人、婴儿)
Private mint心率应用 As Integer
Private mstrNote As String '体温曲线未记说明信息
Private mintType As Integer
Private mstrModifyTime As String
Private mint数据来源 As Integer, mintModify As Integer
Private mstr科室性质 As String
Private mbln脉搏共用显示 As Boolean

Public Event AfterRowColChange(ByVal strInfo As String, ByVal blnImportant As Boolean)
Public Event UsrHelp()
Public Event UsrExit()

Public Function ShowMe(ByVal frmParent As Form, ByVal lngDeptID As Long, Optional ByVal strPrivs As String, Optional ByVal bytSize As Byte = 0) As Boolean
    '******************************************************************************************************************
    '功能： 显示体温单内容
    '参数： frmParent           上级窗体对象
    '       lngDeptID           要显示护理记录的科室
    '返回： 无
    '******************************************************************************************************************
    Dim lngRow As Long
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo Errhand
    Err = 0

    mblnInit = False
    mlng病区ID = lngDeptID
    mstrPrivs = strPrivs
    mintBigSize = bytSize ' zlDatabase.GetPara("护理文件显示模式", glngSys, 1255, 0)
    Set mfrmParent = frmParent

    mintPreDays = Val(zlDatabase.GetPara("超期录入护理数据天数", glngSys, 1255, "1"))
    mbln脉搏共用显示 = (Val(zlDatabase.GetPara("脉搏短绌以(心率/脉搏)方式录入", glngSys, 1255, "1")) = 1)
    
    If mrsItems.State = 0 Then
        Call InitMenuBar
        Call InitEnv            '初始化环境
        cbsThis.ActiveMenuBar.Visible = False
        cbsThis.RecalcLayout
    End If
    Call GetLocalSetting '从注册表中读取护理界面参数
    Call InitCons
    Call InitVariable
    
    If cboUnit.ListCount = 0 Then
        MsgBox "您不属于当前病区的任何科室，不能使用该功能！", vbInformation, gstrSysName
        Exit Function
    End If
    
    Call ReSetFontSize
    ShowMe = True
    Exit Function
Errhand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ReSetFontSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '功能:设置字体大小
    '入参:bytSize：0-小(缺省)，1-大
    '编制:刘鹏飞
    '日期:2012-06-20 15:15:00
    '问题:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCtrl As Control
    Dim CtlFont As StdFont
    Dim bytFontSize As Byte
    bytFontSize = IIf(mintBigSize = 0, 9, IIf(mintBigSize = 1, 12, mintBigSize))
    
    UserControl.FontSize = bytFontSize
    UserControl.FontName = "宋体"
    For Each objCtrl In UserControl.Controls
        Select Case UCase(TypeName(objCtrl))
        Case UCase("Label")
            If UCase(objCtrl.Name) <> UCase("lblInfo") Then
            objCtrl.FontSize = bytFontSize
            objCtrl.Height = TextHeight("刘") + 20
            End If
        Case UCase("ListBox")
            objCtrl.FontSize = bytFontSize
        Case UCase("VsFlexGrid")
            objCtrl.FontSize = bytFontSize
        Case UCase("ComboBox")
            objCtrl.FontSize = bytFontSize
        Case UCase("OptionButton")
            objCtrl.FontSize = bytFontSize
            objCtrl.Width = TextWidth("刘鹏" & objCtrl.Caption) - TextWidth("刘") / 3
        Case UCase("CheckBox")
            objCtrl.FontSize = bytFontSize
            If UCase(objCtrl.Name) <> UCase("chkSwitch") Then
                objCtrl.Width = TextWidth("刘鹏" & objCtrl.Caption) - TextWidth("刘") / 3
            End If
        Case UCase("DTPicker")
            objCtrl.Font.Size = bytFontSize
            objCtrl.Width = TextWidth("2012-01-01") + 400
            objCtrl.Height = TextHeight("刘") * 1.5
        Case UCase("TextBox")
          objCtrl.FontSize = bytFontSize
        Case UCase("ReportControl")
            Set CtlFont = objCtrl.PaintManager.CaptionFont
            CtlFont.Size = bytFontSize
            Set objCtrl.PaintManager.CaptionFont = CtlFont
            
            Set CtlFont = objCtrl.PaintManager.TextFont
            CtlFont.Size = bytFontSize
            Set objCtrl.PaintManager.TextFont = CtlFont
            objCtrl.Redraw
        Case UCase("DockingPane")
            Set CtlFont = objCtrl.PaintManager.CaptionFont
            If CtlFont Is Nothing Then
                Set CtlFont = UserControl.Font
            End If
            CtlFont.Size = bytFontSize
            Set objCtrl.PaintManager.CaptionFont = CtlFont
        Case UCase("CommandBars")
            Set CtlFont = objCtrl.Options.Font
            If CtlFont Is Nothing Then
                Set CtlFont = UserControl.Font
            End If
            CtlFont.Size = bytFontSize
            Set objCtrl.Options.Font = CtlFont
        Case UCase("TabControl")
            Set CtlFont = objCtrl.PaintManager.Font
            If CtlFont Is Nothing Then
                Set CtlFont = UserControl.Font
            End If
            CtlFont.Size = bytFontSize
            Set objCtrl.PaintManager.Font = CtlFont
        Case UCase("CommandButton")
            objCtrl.FontSize = bytFontSize
            objCtrl.Width = TextWidth(" " & IIf(objCtrl.Caption = "", "  ", objCtrl.Caption) & " ")
        End Select
    Next
    
    '移动控件位置
    lblDate.Left = 60
    dtpDate.Left = lblDate.Left + lblDate.Width + TextWidth("刘") / 2
    lblUnit.Left = dtpDate.Left + dtpDate.Width + TextWidth("刘")
    cboUnit.Left = lblUnit.Left + lblUnit.Width + TextWidth("刘") / 2
    lblFilter.Left = cboUnit.Left + cboUnit.Width + TextWidth("刘")
    txtFilter.Left = lblFilter.Left + lblFilter.Width + TextWidth("刘") / 2
    cmdSift.Height = txtFilter.Height - TextHeight("刘") / 4
    cmdSift.Width = TextWidth("刘") + TextWidth("刘") / 2
    cmdSift.Left = txtFilter.Left + txtFilter.Width - cmdSift.Width
    lblPati.Left = txtFilter.Left + txtFilter.Width + TextWidth("刘")
    cboPati.Left = lblPati.Left + lblPati.Width + TextWidth("刘") / 2
    cmdFilter.Left = IIf(lblPati.Visible = True, cboPati.Left + cboPati.Width + (TextWidth("刘") / 2) + 15, lblPati.Left)
    cmdFilter.Top = cboUnit.Top
    cmdFilter.Height = TextHeight("刘") + 100
    cmdAddUser.Left = cmdFilter.Left + cmdFilter.Width + TextWidth("刘") + 75
    cmdAddUser.Top = cmdFilter.Top
    cmdAddUser.Height = cmdFilter.Height
    pic过滤条件.Width = cmdAddUser.Left + cmdAddUser.Width + TextWidth("刘") / 2
    
    txt显示天数.Left = lblDay.Left + lblDay.Width + TextWidth("刘")
    lblDay.Top = txt显示天数.Top + (txt显示天数.Height - lblDay.Height) \ 2
    lblDayInfo.Top = lblDate.Top
    lblDayInfo.Left = txt显示天数.Left + txt显示天数.Width + TextWidth("刘")
    
    
    '添加病人移动
    chkScope(0).Left = 150
    chkScope(1).Left = chkScope(0).Left + chkScope(0).Width + TextWidth("刘") / 2
    chkScope(1).Left = chkScope(0).Left + chkScope(0).Width + TextWidth("刘") / 2
    chkScope(3).Left = chkScope(1).Left + chkScope(1).Width + TextWidth("刘") / 2
    
    chkScope(2).Left = chkScope(0).Left
    lbl出院时间.Left = chkScope(2).Left + chkScope(2).Width + TextWidth("刘") / 2
    lbl出院时间.Top = chkScope(2).Top
    dtpB(0).Left = lbl出院时间.Left + lbl出院时间.Width + TextWidth("刘") / 2
    dtpB(0).Top = lbl出院时间.Top - (dtpB(0).Height - lbl出院时间.Height) \ 2
    Label2(0).Left = dtpB(0).Left + dtpB(0).Width + TextWidth("刘") / 2
    Label2(0).Top = lbl出院时间.Top
    dtpE(0).Left = Label2(0).Left + Label2(0).Width + TextWidth("刘") / 2
    dtpE(0).Top = dtpB(0).Top
    CmdRef.Left = dtpE(0).Left + dtpE(0).Width + TextWidth("刘") / 2
    CmdRef.Height = TextHeight("刘") + 100
    CmdRef.Top = dtpE(0).Top
    
    lbl转出.Left = Label2(0).Left
    fraChange.Left = lbl转出.Left + TextWidth("显示最近")
    fraChange.Top = lbl转出.Height + lbl转出.Top
    fraChange.Width = TextWidth("转科")
    txtChange.Left = fraChange.Left + (fraChange.Width - txtChange.Width) / 2
    txtChange.Height = TextHeight("刘")
    txtChange.Top = fraChange.Top - txtChange.Height
    
    chkPati(1).Left = chkPati(0).Left + chkPati(0).Width + TextWidth("刘") / 2
    picPati.Width = lbl转出.Left + lbl转出.Width + TextWidth("刘") / 2
    rptPati.Width = picPati.Width - 15
    
    '条件
    picFilter.Width = TextWidth("三天内体温存在超过37.5度的病人") + TextWidth("刘鹏飞")
    If picFilter.Width < 2145 Then picFilter.Width = 2145
    lstFilter.Height = lstFilter.ListCount * TextHeight("刘") + 30
    picFilter.Height = lstFilter.Height + cmdFilterOK.Height + 120
End Sub

Private Sub GetLocalSetting()
'功能：从注册表读取出院病人的时间范围
    Dim i As Integer
    Dim curDate As Date, intDay As Integer

    '病人显示范围
    mintChange = Val(zlDatabase.GetPara("最近转出天数", glngSys, p住院护士站, 7))
    '如果大于30天就取缺省值
    If mintChange > 30 Then mintChange = 7
    
    '出院病人时间范围
    curDate = zlDatabase.Currentdate
    mdtOutEnd = Format(curDate, "yyyy-MM-dd")
    mdtOutBegin = Format(CDate(mdtOutEnd) - 3, "yyyy-MM-dd")
    
    txtChange.Text = mintChange
    dtpE(0).Value = mdtOutEnd
    dtpE(0).Tag = mdtOutEnd
    dtpB(0).Value = mdtOutBegin
    dtpB(0).Tag = mdtOutBegin
    
    For i = 0 To 3
        If i = 在院 Then
            chkScope(i).Value = 1
        Else
            chkScope(i).Value = 0
        End If
    Next i
    
    dtpB(0).Enabled = False
    dtpE(0).Enabled = False
    txtChange.Enabled = False
End Sub

Private Function RefreshHistoryData(ByVal lngRow As Long) As Boolean
'提取病人历史数据信息
    Dim lng文件ID As Long, int护理等级 As Integer, int婴儿 As Integer
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer
    Dim lngItemNO As Long, lngHistoryRow As Long, lngHistoryCol As Long
    Dim strContent As String
    Dim strValues As String, strFileds As String
    Dim strKey As String, strKeys As String
    Dim intCOl As Integer
    Dim strPart As String, lng来源ID As Long, int共用 As Integer, int显示 As Integer, intModify As Integer, strNote As String
    Dim bln脉搏 As Boolean, blnAllow As Boolean, bln心率单独 As Boolean
    Dim lngRecordID As Long
    Dim arrRecordID
    Dim strStart As String, strEnd As String, strTime As String
    
    On Error GoTo Errhand:
    
    If Not mblnInit Then Exit Function
    If lngRow < VsfData.FixedRows - 1 Or Val(VsfData.TextMatrix(lngRow, c文件ID)) <= 0 Then Exit Function
    lng文件ID = Val(VsfData.TextMatrix(lngRow, c文件ID))
    int护理等级 = Val(VsfData.TextMatrix(lngRow, c护理等级))
    int婴儿 = Val(VsfData.TextMatrix(lngRow, c婴儿))
    
    bln心率单独 = True
    bln脉搏 = False
    mrsItems.Filter = 0
    mrsItems.Filter = "项目序号=" & gint脉搏
    If mrsItems.RecordCount > 0 Then bln脉搏 = True
    
    If mrsHistory Is Nothing Then Exit Function
    strFileds = "ID|行号|项目序号|数据|部位|数据来源|来源ID|共用|显示|修改|状态"
    
    '清空历史数据集
    mrsHistory.Filter = 0
    If mrsHistory.RecordCount <> 0 Then mrsHistory.MoveFirst
    Do While True
        If mrsHistory.EOF Then Exit Do
        mrsHistory.Delete
        mrsHistory.Update
        mrsHistory.MoveNext
    Loop
    vsfHistory.Rows = vsfHistory.FixedRows
    lngHistoryRow = vsfHistory.Rows - 1
    vsfHistory.RowHeight(-1) = 300 + mintBigSize * 300 / 3
    vsfHistory.FontSize = 9 + mintBigSize * 9 / 3
    
    lngRecordID = 0
    arrRecordID = Array()
    
    '时间范围
    strStart = Format(DateAdd("d", -1 * Val(txt显示天数.Text), IIf(IsDate(mstrDate), mstrDate, zlDatabase.Currentdate)), "yyyy-MM-dd") & " 00:00:00"
    strEnd = Format(DateAdd("d", mintPreDays, IIf(IsDate(mstrDate), mstrDate, zlDatabase.Currentdate)), "yyyy-MM-dd") & " 23:59:59"
    
    strTime = VsfData.TextMatrix(lngRow, c日期)
    If strTime <> "" Then
        If InStr(1, strTime, ";") <> 0 Then strTime = Split(strTime, ";")(0)
        If IsDate(Format(strTime, "YYYY-MM-DD HH:mm:ss")) Then
            If CDate(strStart) < CDate(strTime) Then strStart = Format(strTime, "YYYY-MM-DD HH:mm:ss")
        End If
        txt显示天数.Text = DateDiff("d", CDate(strStart), IIf(IsDate(mstrDate), mstrDate, zlDatabase.Currentdate))
    End If

    lblDayInfo.Caption = "时间范围：" & strStart & " 到 " & strEnd
    
    '提取选择病人的历史数据
    mstrSQL = "SELECT B.ID,B.发生时间,C.项目序号,C.记录内容,C.数据来源,C.体温部位,C.未记说明,C.来源ID,C.共用,C.显示,DECODE(C.项目序号,-1,1,C.记录标记) 记录标记" & vbNewLine & _
        " FROM 病人护理文件 A,病人护理数据 B,病人护理明细 C,护理记录项目 D,体温记录项目 E" & vbNewLine & _
        " WHERE A.ID=B.文件ID AND B.ID=C.记录ID AND A.ID=[1]" & vbNewLine & _
        " AND Mod(C.记录类型,5)<>5 And C.终止版本 IS NULL  AND B.发生时间 Between [2] And [3] And C.项目序号=D.项目序号 And D.项目序号=E.项目序号 AND nvl(D.护理等级,3) >=[4] And Nvl(D.适用病人,0) In (0,[5])" & vbNewLine & _
        " Order By B.发生时间,DECODE(C.项目序号,-1,1,0),DECODE(C.项目序号,-1,1,C.记录标记),项目序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "同步数据", lng文件ID, CDate(strStart), CDate(strEnd), int护理等级, IIf(int婴儿 = 0, 1, 2))
    
    If rsTemp.RecordCount = 0 Then GoTo NextPos
    rsTemp.MoveFirst
    '检查存在多少记录
    With rsTemp
        Do While Not .EOF
            If lngRecordID <> Nvl(rsTemp!Id, 0) Then
                ReDim Preserve arrRecordID(UBound(arrRecordID) + 1)
                arrRecordID(UBound(arrRecordID)) = rsTemp!Id
                lngHistoryRow = lngHistoryRow + 1
                If lngHistoryRow > vsfHistory.Rows - 1 Then vsfHistory.Rows = vsfHistory.Rows + 1
                vsfHistory.TextMatrix(lngHistoryRow, c日期) = Format(rsTemp!发生时间, "yyyy-MM-dd")
                vsfHistory.TextMatrix(lngHistoryRow, c时间) = Format(rsTemp!发生时间, "HH:mm:ss")
            End If
            lngRecordID = Val(Nvl(rsTemp!Id, 0))
        .MoveNext
        Loop
    End With
    
    '循环赋值
    For lngRecordID = 0 To UBound(arrRecordID)
        lngHistoryRow = lngRecordID + vsfHistory.FixedRows
        rsTemp.Filter = "ID=" & Val(arrRecordID(lngRecordID))
        For lngHistoryCol = RootCol To vsfHistory.Cols - 1
            lngItemNO = Val(vsfHistory.TextMatrix(0, lngHistoryCol))
            
            If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst
    
            strContent = ""
            strPart = ""
            strNote = ""
            lng来源ID = 0
            int共用 = 0
            int显示 = 0
            intModify = 0
            With rsTemp
                Do While Not .EOF
                    If lngItemNO <> 4 Then
                        blnAllow = False
                        bln心率单独 = False
                        intModify = 0
                        
                        If InStr(1, "," & gint体温 & "," & gint脉搏 & "," & gint心率 & ",", "," & Val(Nvl(!项目序号)) & ",") > 0 Then
                            Select Case Val(Nvl(!项目序号))
                                Case gint体温
                                    If gint体温 = lngItemNO Then blnAllow = True
                                Case gint脉搏
                                    If gint脉搏 = lngItemNO Then blnAllow = True
                                Case gint心率
                                    If bln脉搏 = True And mint心率应用 = 2 Then
                                        If gint脉搏 = lngItemNO Then blnAllow = True
                                    Else
                                        If gint心率 = lngItemNO Then blnAllow = True: bln心率单独 = True
                                    End If
                            End Select
                            
                            If blnAllow = True Then
                                If Val(Nvl(!记录标记)) = 0 Then
                                    strContent = Nvl(!记录内容)
                                    strPart = Nvl(!体温部位)
                                    lng来源ID = Val(Nvl(!来源ID))
                                    int共用 = Val(Nvl(!共用))
                                    int显示 = Val(Nvl(!显示))
                                    strNote = Nvl(!未记说明)
                                Else '组装物理降温和脉搏短轴
                                    If bln心率单独 = False Then
                                        If strContent <> "" Then
                                            If InStr(1, strContent, "/") = 0 Then
                                                '脉搏短绌显示格式:心率/脉搏
                                                If mbln脉搏共用显示 And lngItemNO = 2 Then
                                                    strContent = Nvl(!记录内容) & "/" & strContent
                                                Else
                                                    strContent = strContent & "/" & Nvl(!记录内容)
                                                End If
                                            Else
                                                '脉搏短绌显示格式:心率/脉搏
                                                If mbln脉搏共用显示 And lngItemNO = 2 Then
                                                    strContent = Nvl(!记录内容) & "/" & Split(strContent, "/")(0)
                                                Else
                                                    strContent = Split(strContent, "/")(0) & "/" & Nvl(!记录内容)
                                                End If
                                                
                                            End If
                                        Else
                                            strContent = Nvl(!记录内容)
                                        End If
                                        
                                        Exit Do
                                    Else
                                        strPart = Nvl(!体温部位)
                                        lng来源ID = Val(Nvl(!来源ID))
                                        int共用 = Val(Nvl(!共用))
                                        int显示 = Val(Nvl(!显示))
                                        strContent = Nvl(!记录内容)
                                        strNote = Nvl(!未记说明)
                                        Exit Do
                                    End If
                                End If
                            End If
                        Else
                            If Val(Nvl(!项目序号)) = lngItemNO Then
                                strPart = Nvl(!体温部位)
                                lng来源ID = Val(Nvl(!来源ID))
                                int共用 = Val(Nvl(!共用))
                                int显示 = Val(Nvl(!显示))
                                strContent = Nvl(!记录内容)
                                strNote = Nvl(!未记说明)
                                Exit Do
                            End If
                        End If
                    ElseIf InStr(1, ",4,5,", "," & Val(!项目序号) & ",") <> 0 And lngItemNO = 4 Then
                        Select Case Val(!项目序号)
                            Case 4
                                If strContent <> "" Or Nvl(!记录内容) <> "" Then
                                    If InStr(1, strContent, "/") > 0 Then
                                        strContent = Nvl(!记录内容) & "/" & Trim(Split(strContent, "/")(1))
                                    Else
                                        strContent = Nvl(!记录内容) & "/"
                                    End If
                                    strPart = Nvl(!体温部位)
                                    lng来源ID = Val(Nvl(!来源ID))
                                    int共用 = Val(Nvl(!共用))
                                    int显示 = Val(Nvl(!显示))
                                End If
                            Case 5
                                If strContent <> "" Or Nvl(!记录内容) <> "" Then
                                    If InStr(1, strContent, "/") > 0 Then
                                        strContent = Trim(Split(strContent, "/")(0)) & "/" & Nvl(!记录内容)
                                    Else
                                        strContent = "/" & Nvl(!记录内容)
                                    End If
                                End If
                        End Select
                    End If
                    .MoveNext
                Loop
                
                If strContent = "/" Then strContent = ""
                If lngItemNO = 4 Then
                    If InStr(1, strContent, "/") <> 0 Then
                        '问题号:53505,修改人：李涛，血压显示文字
                        If Split(strContent, "/")(0) = "拒测" Or Split(strContent, "/")(0) = "外出" Or Split(strContent, "/")(0) = "请假" Or Split(strContent, "/")(0) = "未测" Then
                            strContent = Split(strContent, "/")(0)
                        Else
                            If Not IsNumeric(Split(strContent, "/")(0)) And Not IsNumeric(Split(strContent, "/")(1)) Then
                                strContent = ""
                            End If
                        End If
                    End If
                End If
                
                '如果是体温曲线项目，并且部位不为空
                mrsItems.Filter = "项目序号=" & lngItemNO
                If mrsItems.RecordCount > 0 Then
                    If Nvl(mrsItems!记录法, 2) = 1 Or lngItemNO = gint呼吸 Then
                        If strNote <> "" And strContent = "" Then
                            strContent = strNote
                            strPart = ""
                        Else
                            If strContent <> "" Then strContent = IIf(strPart = "", "", strPart & ":") & strContent
                        End If
                    Else
                        strPart = ""
                    End If
                Else
                    strPart = ""
                End If
                
                If strContent <> "" Then
                    '将同步的数据装载到记录集中
                    strKey = lngHistoryRow & "," & lngHistoryCol
                    strValues = strKey & "|" & lngHistoryRow & "|" & lngItemNO & "|" & strContent & "|" & strPart & "|1|" & lng来源ID & "|" & int共用 & "|" & int显示 & "|" & intModify & "|0"
                    Call Record_Update(mrsHistory, strFileds, strValues, "ID|" & strKey)
                    vsfHistory.TextMatrix(lngHistoryRow, lngHistoryCol) = strContent
                    vsfHistory.RowData(lngHistoryRow) = Val(arrRecordID(lngRecordID))
                    If lng来源ID <> 0 Then '同步过来的数据
                        vsfHistory.Cell(flexcpForeColor, lngHistoryRow, lngHistoryCol, lngHistoryRow, lngHistoryCol) = 255
                    Else
                        vsfHistory.Cell(flexcpForeColor, lngHistoryRow, lngHistoryCol, lngHistoryRow, lngHistoryCol) = &H80000008
                    End If
'                    If lngItemNo = gint体温 Or (lngItemNo = gint脉搏 And mint心率应用 = 2) Then
'                        vsfHistory.Cell(flexcpForeColor, lngHistoryRow, lngHistoryCol, lngHistoryRow, lngHistoryCol) = RGB(0, 0, 255)
'                    Else
'                        vsfHistory.Cell(flexcpForeColor, lngHistoryRow, lngHistoryCol, lngHistoryRow, lngHistoryCol) = 255 '&H8080FF
'                    End If
                End If
            End With
        Next lngHistoryCol
    Next lngRecordID
    
    vsfHistory.Cell(flexcpAlignment, vsfHistory.FixedRows, c时间, vsfHistory.Rows - 1, vsfHistory.Cols - 1) = flexAlignCenterCenter
    
    '不适用此病人的列隐藏
    For lngHistoryCol = RootCol To vsfHistory.Cols - 1
        vsfHistory.ColHidden(lngHistoryCol) = False
        lngItemNO = Val(vsfHistory.TextMatrix(0, lngHistoryCol))
        mrsItems.Filter = 0
        mrsItems.Filter = "项目序号=" & lngItemNO & " And 护理等级>=" & int护理等级
        If mrsItems.RecordCount = 0 Then
            vsfHistory.ColHidden(lngHistoryCol) = True
        Else
            '检查是否适用于此病人
            If Val(vsfHistory.TextMatrix(2, lngHistoryCol)) = 1 Then
                vsfHistory.ColHidden(lngHistoryCol) = IIf(int婴儿 = 0, False, True)
            ElseIf vsfHistory.TextMatrix(2, lngHistoryCol) = 2 Then
                vsfHistory.ColHidden(lngHistoryCol) = IIf(int婴儿 <> 0, False, True)
            End If
        End If
    Next lngHistoryCol
    vsfHistory.RowHeight(-1) = 300 + mintBigSize * 300 / 3
    vsfHistory.FontSize = 9 + mintBigSize * 9 / 3
    
    mrsItems.Filter = 0
NextPos:
    RefreshHistoryData = True
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub RefreshPatiList()
    '刷新病人清单
    Call LoadPatient
    If mrsPati.RecordCount > 0 Then mrsPati.MoveFirst
    rptPati.Records.DeleteAll
End Sub

Private Function InitMenuBar() As Boolean
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
    '******************************************************************************************************************
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim cbrPop As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim RS As ADODB.Recordset
    Dim objExtendedBar As CommandBar

    On Error GoTo Errhand

    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbsThis.ActiveMenuBar.Title = "菜单栏"
    cbsThis.ActiveMenuBar.Visible = False

    Set cbsThis.Icons = zlCommFun.GetPubIcons
    With cbsThis.Options
        .AlwaysShowFullMenus = False
        .ShowExpandButtonAlways = False
        .UseDisabledIcons = True
        .SetIconSize True, 24, 24
        .LargeIcons = True
    End With

    '------------------------------------------------------------------------------------------------------------------
    '菜单定义
    cbsThis.ActiveMenuBar.Title = "菜单"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)

    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义
    Set cbrToolBar = cbsThis.Add("标准工具", xtpBarTop)
    cbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    cbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    cbrToolBar.ShowTextBelowIcons = False
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Send, "同步"):   cbrControl.ToolTipText = "数据同步"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Clear, "清除"):   cbrControl.ToolTipText = "清除、删除某行数据"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "空行"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "增加空行"
        Set mcbrMenuBar部位 = .Add(xtpControlButtonPopup, conMenu_Edit_Compend, "部位"): mcbrMenuBar部位.ToolTipText = "体温部位"
        
        Set cbrPop = .Add(xtpControlButtonPopup, conMenu_Edit_Append, "特殊处理"): cbrPop.BeginGroup = True
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 0, "正常", -1, False): cbrControl.IconId = 1: cbrControl.Parameter = ""
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 1, "灌肠[E]", -1, False):  cbrControl.IconId = 1: cbrControl.Parameter = "E"
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 2, "灌肠后大便[/E]", -1, False):  cbrControl.IconId = 1: cbrControl.Parameter = "/E"
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 3, "大便失禁[※]", -1, False): cbrControl.IconId = 1: cbrControl.Parameter = "※"
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 4, "人工肛门[☆]", -1, False): cbrControl.IconId = 1: cbrControl.Parameter = "☆"
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 5, "导尿[C]", -1, False):   cbrControl.IconId = 1: cbrControl.Parameter = "C"
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 6, "保留导尿[/C]", -1, False):   cbrControl.IconId = 1: cbrControl.Parameter = "/C"

        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Blankoff, "清空表格"): cbrControl.ToolTipText = "清除表格所有行": cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "取消")
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助")
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
    End With

    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel Then
            cbrControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    Set mcbrToolBar = cbrToolBar
    '------------------------------------------------------------------------------------------------------------------
    '工具栏定义
    Set cbrToolBar = cbsThis.Add("过滤条件", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    cbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    With cbrToolBar.Controls
        Set cbrCustom = .Add(xtpControlCustom, conMenu_View_LocationItem, "")
        cbrCustom.flags = xtpFlagAlignLeft
        cbrCustom.Handle = pic过滤条件.hWnd
        cbrCustom.ToolTipText = "条件"
    End With

    '快键绑定
    With cbsThis.KeyBindings
        
        .Add FALT, Asc("0"), conMenu_Edit_Append * 10
        .Add FALT, Asc("1"), (conMenu_Edit_Append * 10 + 1)
        .Add FALT, Asc("2"), (conMenu_Edit_Append * 10 + 2)
        .Add FALT, Asc("3"), (conMenu_Edit_Append * 10 + 3)
        .Add FALT, Asc("4"), (conMenu_Edit_Append * 10 + 4)
        .Add FALT, Asc("5"), (conMenu_Edit_Append * 10 + 5)
        .Add FALT, Asc("6"), (conMenu_Edit_Append * 10 + 6)
        
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save
        .Add FCONTROL, Asc("R"), conMenu_Edit_Transf_Cancle
        .Add 0, VK_F1, conMenu_Help_Help
    End With

    InitMenuBar = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub AddActiveMenu(ByVal lngItemNO As Long)
    '------------------------------------------------------------
    '根据项目添加菜单(主要用来添加体温曲线项目部位信息)
    Dim varTmp As Variant
    Dim strPart As String
    Dim RS As New ADODB.Recordset
    Dim cbrControl As CommandBarControl
    Dim i As Integer
    
    On Error GoTo Errhand
    
    If Not mcbrMenuBar部位 Is Nothing Then
        If mcbrMenuBar部位.CommandBar.Controls.Count <> 0 Then
            Call mcbrMenuBar部位.CommandBar.Controls.DeleteAll
        End If
    End If
    
    If mrsPart Is Nothing Then Exit Sub
    If lngItemNO = 0 Then Exit Sub
    
    If lngItemNO = gint体温 Then '体温
        mstrSQL = "Select 记录符 From 体温记录项目 Where 项目序号=[1]"
        Set RS = zlDatabase.OpenSQLRecord(mstrSQL, "体温单批量录入", gint体温)
        If RS.BOF = False Then
            varTmp = Split(Nvl(RS("记录符").Value, "·,×,○"), ",")
        Else
            varTmp = Split("·,×,○", ",")
        End If
        
        Set cbrControl = mcbrMenuBar部位.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Compend * 10, "口温" & varTmp(0) & " (&1)", -1, False): cbrControl.Parameter = "口温": cbrControl.IconId = 1
        Set cbrControl = mcbrMenuBar部位.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Compend * 10, "腋温" & varTmp(1) & " (&2)", -1, False): cbrControl.Parameter = "腋温": cbrControl.IconId = 1
        Set cbrControl = mcbrMenuBar部位.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Compend * 10, "肛温" & varTmp(2) & " (&3)", -1, False): cbrControl.Parameter = "肛温": cbrControl.IconId = 1
    ElseIf lngItemNO = gint呼吸 Then '呼吸
        mstrSQL = "Select 记录符 From 体温记录项目 Where 项目序号=[1]"
        Set RS = zlDatabase.OpenSQLRecord(mstrSQL, "体温单批量录入", gint呼吸)
        If RS.BOF = False Then
            varTmp = Nvl(RS("记录符").Value, "⊕")
        Else
            varTmp = "⊕"
        End If
        
        Set cbrControl = mcbrMenuBar部位.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Compend * 10, "自主呼吸" & varTmp & " (&1)", -1, False): cbrControl.Parameter = "自主呼吸": cbrControl.IconId = 1
        Set cbrControl = mcbrMenuBar部位.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Compend * 10, "呼吸机 (&2)", -1, False): cbrControl.Parameter = "呼吸机": cbrControl.IconId = 1
    ElseIf lngItemNO = gint脉搏 Then '脉搏
        mstrSQL = "Select 记录符 From 体温记录项目 Where 项目序号=[1]"
        Set RS = zlDatabase.OpenSQLRecord(mstrSQL, "体温单批量录入", gint脉搏)
        
        If RS.BOF = False Then
            varTmp = Nvl(RS("记录符").Value, "+")
        Else
            varTmp = "+"
        End If
        
        Set cbrControl = mcbrMenuBar部位.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Compend * 10, "      " & varTmp & " (&1)", -1, False): cbrControl.Parameter = "": cbrControl.IconId = 1
        Set cbrControl = mcbrMenuBar部位.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Compend * 10, "起搏器" & " (&2)", -1, False): cbrControl.Parameter = "起搏器": cbrControl.IconId = 1
    Else '其他曲线项目部位信息
        varTmp = ""
        mstrSQL = "Select 记录符 From 体温记录项目 Where 项目序号=[1]"
        Set RS = zlDatabase.OpenSQLRecord(mstrSQL, "体温单批量录入", lngItemNO)
        If RS.BOF = False Then
            varTmp = Nvl(RS("记录符").Value)
        End If
        mrsPart.Filter = 0
        mrsPart.Filter = "项目序号=" & lngItemNO
        If mrsPart.RecordCount > 1 Then
            i = 1
            varTmp = varTmp & String(mrsPart.RecordCount - 1 - UBound(Split(varTmp, ",")), ",")
            Do While Not mrsPart.EOF
                strPart = Nvl(mrsPart!部位)
                If strPart = "" Then strPart = "   "
                varTmp = Split(varTmp, ",")
                Set cbrControl = mcbrMenuBar部位.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Compend * 10, strPart & varTmp(i - 1) & " (&1)", -1, False): cbrControl.Parameter = strPart: cbrControl.IconId = 1
                i = i + 1
            mrsPart.MoveNext
            Loop
        End If
    End If
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitCons()
    '隐藏输入控件
    picFilter.Visible = False
    picPati.Visible = False
    picInput.Visible = False
    picDouble.Visible = False
    lstNote.Visible = False
    lstSelect(0).Visible = False
    lstSelect(1).Visible = False
    cbo体温标识.Visible = False
    PicLst.Visible = False
    txtLst.Visible = False
    mintType = 0
End Sub

Private Sub InitVariable(Optional ByVal blnClearDate As Boolean = True)
    '清除常量
    mstrModifyTime = ""
    mblnChage = False
    mblnSaveData = False
    If blnClearDate = True Then
        mint心率应用 = 0
        mbln出院 = False
    End If
    mstrTabHead = ""
    mstrItemNo = ""
    mint数据来源 = 0
    mintModify = 0
    mintType = 0
    mblnShow = False
    mblnNullRow = False
    mblnClearRow = False
    mblnRefreshData = False
    mblnChkClick = False
    mblnDateFouces = False
End Sub

Private Function InitFilter() As Boolean
'功能：初始化体温单批量录入过滤条件
    Dim strFilter As String, strFilterID As String
    Dim arrFilter() As String, arrFilterID() As String
    Dim arrSel() As String
    Dim strSel As String
    Dim i As Integer
    Dim blnSelAll As Boolean
    
    strSel = zlDatabase.GetPara("体温单过滤条件", glngSys, 1255)
    
    '51286,刘鹏飞,2012-07-11,添加过滤一级及以上护理登记的病人
    If strSel = "" Then
        strSel = "1;1;1;1;1;1;1"
    Else
        arrSel = Split(strSel, ";")
        strSel = strSel & String(IIf(6 - UBound(arrSel) < 0, 0, 6 - UBound(arrSel)), ";")
    End If
    
    arrSel = Split(strSel, ";")
    txtFilter.Tag = ""
    txtFilter.Text = ""
    strFilter = "全部;入院三天内的病人;手术三天内的病人;三天内体温存在超过37.5度的病人;危/重病人;转入三天内的病人;一级及以上护理等级的病人" & IIf(mstr科室性质 = "产科" Or mstr科室性质 = "所有", ";分娩后三天内的病人", "")
    strFilterID = "0;1;2;3;4;5;6" & IIf(mstr科室性质 = "产科" Or mstr科室性质 = "所有", ";7", "")
    arrFilter = Split(strFilter, ";")
    arrFilterID = Split(strFilterID, ";")
    
    blnSelAll = True
    lstFilter.Clear
    For i = 0 To UBound(arrFilter)
        lstFilter.AddItem CStr(arrFilter(i))
        lstFilter.ItemData(lstFilter.NewIndex) = Val(arrFilterID(i))
        
        If i <> 0 Then
            If Val(arrSel(i - 1)) = 1 Then
                txtFilter.Text = txtFilter.Text & ";" & arrFilter(i)
                txtFilter.Tag = txtFilter.Tag & ";" & arrFilterID(i)
            Else
                blnSelAll = False
            End If
        End If
    Next i
    
    If blnSelAll = True Then
        txtFilter.Text = "全部"
        txtFilter.Tag = 0
    Else
        txtFilter.Text = Mid(txtFilter.Text, 2)
        txtFilter.Tag = Mid(txtFilter.Tag, 2)
    End If
    
    '设置条件大小
    picFilter.Width = LenB(StrConv(lstFilter.List(lstFilter.ListCount \ 2), vbFromUnicode)) * 160 + 500
    If picFilter.Width < TextWidth("三天内体温存在超过37.5度的病人") + TextWidth("刘鹏飞") Then
        picFilter.Width = TextWidth("三天内体温存在超过37.5度的病人") + TextWidth("刘鹏飞")
    End If
    If picFilter.Width < 2145 Then picFilter.Width = 2145
    lstFilter.Height = lstFilter.ListCount * TextHeight("刘") + 30
    picFilter.Height = lstFilter.Height + cmdFilterOK.Height + 120
    
    InitFilter = True
    Exit Function
End Function

Private Sub InitEnv()
    Dim curDate As Date
    Dim intDay As Integer
    Dim RS As New ADODB.Recordset
    Dim blnVisible As Boolean
    On Error GoTo Errhand
    
    mlngHours = Val(Mid(Val(zlDatabase.GetPara("数据补录时限", glngSys)), 1, 6))
    txt显示天数.Tag = 1
    
    dtpDate.Value = Format(date, "YYYY-MM-DD")
    
    If mrsPart Is Nothing Then Set mrsPart = New ADODB.Recordset
    If mrsPart.State = 1 Then mrsPart.Close
    
    '提取所有部位信息
    mstrSQL = "SELECT 项目序号,部位,缺省项,固定项 FROM 体温部位"
    Call zlDatabase.OpenRecordset(mrsPart, mstrSQL, "提取部位提取")
    
    '打开现存在的所有护理记录项目
    mstrSQL = " Select B.分组名,B.项目序号,B.项目名称,B.项目类型,B.项目性质,B.项目长度,B.项目小数,B.项目表示,B.项目单位,B.项目值域,B.护理等级,B.适用病人,B.应用方式,nvl(A.记录法,2) 记录法" & _
              " From 护理记录项目 B,体温记录项目 A" & _
              " Where B.项目序号=A.项目序号(+) And B.应用方式<>0 " & _
              " Order by B.项目序号"
    Set mrsItems = zlDatabase.OpenSQLRecord(mstrSQL, "打开现存在的所有护理记录项目")
    
    '提取未记说明信息
    mstrNote = ""
    mstrSQL = "Select 编码,名称 From 常用体温说明"
    Call zlDatabase.OpenRecordset(RS, mstrSQL, "未记说明信息")
    lstNote.Clear
    With RS
        Do While Not .EOF
            lstNote.AddItem Nvl(!名称)
            lstNote.ItemData(lstNote.NewIndex) = Val(!编码)
            mstrNote = mstrNote & "," & Nvl(!名称)
        .MoveNext
        Loop
    End With
    If lstNote.ListCount > 0 Then lstNote.ListIndex = 0
    
    If Left(mstrNote, 1) = "," Then mstrNote = Mid(mstrNote, 2)
    
    '提取体温单清单
    gstrSQL = " Select ID FROM 病历文件列表 WHERE 种类=3 AND 保留=-1 AND NVL(通用,0)>0 "
    Set RS = zlDatabase.OpenSQLRecord(gstrSQL, "提取除体温单外的护理文件清单")
    
    If RS.RecordCount > 0 Then
        mlng格式ID = Val(Nvl(RS!Id))
    Else
        mlng格式ID = 0
        MsgBox "在病人文件列表中没有找到体温单相关的文件,请检查!", vbInformation, gstrSysName
    End If
    
    blnVisible = False
    '提取当前病区下的所有科室
    mstrSQL = " Select distinct B.ID,B.编码||'-'||B.名称 AS 科室,decode(nvl(E.工作性质,''),'产科',1,0) 性质" & _
              " From 病区科室对应 A,部门表 B,部门人员 C,人员表 D,部门性质说明 E" & _
              " Where A.科室ID = b.ID And A.科室ID=C.部门ID And C.人员ID=D.ID And A.病区ID = [1]" & _
              IIf(InStr(1, mstrPrivs, "当前病区") <> 0, "", " And D.ID=[2]") & _
              " And B.ID=E.部门ID(+) And E.工作性质(+)='产科'" & _
              " Order by B.编码||'-'||B.名称"
    Set RS = zlDatabase.OpenSQLRecord(mstrSQL, "提取当前病区下的所有科室", mlng病区ID, glngUserId)
    With cboUnit
        .Clear
        .Tag = ""
        If InStr(1, mstrPrivs, "当前病区") <> 0 Then
            .AddItem "所有科室"
            .ItemData(.NewIndex) = -1
        End If
        Do While Not RS.EOF
            .AddItem zlCommFun.Nvl(RS!科室)
            .ItemData(.NewIndex) = RS!Id
            .Tag = .Tag & "[LPF]" & RS!性质
            If blnVisible = False Then blnVisible = (Val(RS!性质) = 1)
            RS.MoveNext
        Loop
        .Tag = IIf(blnVisible = True, 1, 0) & .Tag
        If Left(.Tag, 5) = "[LPF]" Then .Tag = Mid(.Tag, 6)
        If .ListCount <> 0 Then .ListIndex = 0
    End With
    
    '加载过滤条件信息
    Call InitFilter
    
    '加载病人选择
    With cboPati
        .AddItem "所有": .ItemData(.NewIndex) = 0
        .AddItem "病人本人": .ItemData(.NewIndex) = 1
        .AddItem "婴儿": .ItemData(.NewIndex) = 2
        .ListIndex = 0
    End With
    
    cbo体温标识.Clear
    cbo体温标识.AddItem "2次/日"
    cbo体温标识.AddItem "4次/日"
    cbo体温标识.AddItem "6次/日"
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub


Private Sub LoadPatient()
    Dim strSQL As String
    On Error GoTo Errhand
    '58890:刘鹏飞,2013-02-26,在院病人读取性能优化(关联在院病人表进行查询)
    '入院等入科和转科待入科病人(病人科室所属的病区都可接收)
    'c.科室id + 0,说明：通过H表的索引连接过滤后，记录数量很少，再连接B表则更快
    If chkScope(待入科).Value = 1 Then
        strSQL = _
            "Select /*+ RULE */Distinct" & vbNewLine & _
            " Decode(B.状态,1,0,Decode(c.开始原因,3,1,2)) As 排序, Decode(Nvl(b.病案状态, 0), 0, 999, b.病案状态) As 排序2," & _
            " Decode(B.状态,1,'入院待入住病人',Decode(c.开始原因,3,'转科待入住病人','转病区待入住病人')) As 类型," & _
            " a.病人id, b.主页id, A.门诊号,B.住院号, a.姓名, a.性别, b.年龄," & vbNewLine & _
            " d.名称 As 科室, c.科室id, c.经治医师 As 住院医师,b.责任护士, b.病案状态, lpad(nvl(C.床号,' '),10,' ') as 床号," & _
            " e.名称 As 护理等级, b.费别,b.当前病况, b.入院日期, b.出院日期,B.出院方式, b.病人类型, b.状态, b.险类, a.就诊卡号," & vbNewLine & _
            " -1 As 路径状态,trunc(sysdate)-trunc(b.入院日期)+1 as 住院天数,Z.颜色" & vbNewLine & _
            "From 病人信息 A, 病案主页 B, 病人变动记录 C, 部门表 D, 收费项目目录 E,病人类型 Z,在院病人 R" & vbNewLine & _
            "Where B.病人类型=Z.名称(+) And A.病人ID=R.病人ID And a.病人id = b.病人id And Nvl(b.主页id, 0) <> 0 And b.病人id = c.病人id And b.主页id = c.主页id And c.科室id = d.Id" & vbNewLine & _
            "      And (d.站点='" & gstrNodeNo & "' Or d.站点 is Null)" & vbNewLine & _
            "      And b.护理等级id = e.Id(+) And Nvl(c.附加床位, 0) = 0 And c.终止时间 Is Null" & vbNewLine & _
            "      And ((c.开始原因 in(1,3) And Exists(Select 1 From 病区科室对应 H Where c.科室id = h.科室id And h.病区id = [1])) or (c.开始原因=15 And c.病区id = [1]))" & vbNewLine & _
            "      And ((c.开始原因 = 1 And b.状态 = 1) Or (c.开始原因 in (3,15) And c.开始时间 Is Null And b.状态 = 2)) "
    End If
    '在院病人
    If chkScope(在院).Value = 1 Then
        strSQL = strSQL & IIf(strSQL <> "", " Union All ", "") & _
            "Select /*+ RULE */ Decode(B.状态,3,4,DECODE(B.出院病床, NULL, 3.1,DECODE(B.状态,2,3.2,3))) as 排序," & _
            " Decode(Nvl(B.病案状态,0),0,999,B.病案状态) as 排序2," & _
            " Decode(B.状态,3,'预出院病人',DECODE(B.出院病床, NULL, '家庭病床',DECODE(B.状态,2,'预转科病人', '在院病人'))) as 类型," & _
            " A.病人ID,B.主页ID,A.门诊号,B.住院号,A.姓名,A.性别,B.年龄,C.名称 as 科室,B.出院科室ID 科室ID,B.住院医师,B.责任护士,B.病案状态," & _
            " lpad(nvl(B.出院病床,' '),10,' ') as 床号,E.名称 as 护理等级,B.费别,B.当前病况,B.入院日期,B.出院日期,B.出院方式,B.病人类型," & _
            " B.状态,B.险类,A.就诊卡号,Nvl(B.路径状态,-1) 路径状态,trunc(sysdate)-trunc(b.入院日期)+1 as 住院天数,z.颜色" & _
            " From 病人信息 A,病案主页 B,部门表 C,收费项目目录 E,病人类型 Z,在院病人 R" & _
            " Where B.病人类型=Z.名称(+) And A.病人ID=B.病人ID And A.住院次数=B.主页ID And Nvl(B.主页ID,0)<>0 And Nvl(B.状态,0)<>1" & _
            " And B.出院科室ID=C.ID And B.护理等级ID=E.ID(+) And (c.站点='" & gstrNodeNo & "' Or c.站点 is Null)" & _
            " And Nvl(B.病案状态,0)<>5 And B.封存时间 is NULL And A.病人ID=R.病人ID And R.病区ID=[1]"
    End If
    '出院病人:出院病人可能已有多次住院
    If chkScope(出院).Value = 1 Then
        strSQL = strSQL & IIf(strSQL <> "", " Union All ", "") & _
            "Select /*+ RULE */ Decode(B.出院方式,'死亡',7,6) as 排序," & _
            " Decode(Nvl(B.病案状态,0),0,999,B.病案状态) as 排序2," & _
            " Decode(B.出院方式,'死亡','死亡病人','出院病人') as 类型," & _
            " A.病人ID,B.主页ID,A.门诊号,B.住院号,A.姓名,A.性别,B.年龄,C.名称 as 科室,B.出院科室ID 科室ID,B.住院医师,B.责任护士,B.病案状态," & _
            " lpad(nvl(B.出院病床,' '),10,' ') AS 床号,E.名称 as 护理等级,B.费别,B.当前病况,B.入院日期,B.出院日期,B.出院方式,B.病人类型," & _
            " B.状态,B.险类,A.就诊卡号,Nvl(B.路径状态,-1) 路径状态,trunc(b.出院日期)-trunc(b.入院日期)+1 as 住院天数,z.颜色" & _
            " From 病人信息 A,病案主页 B,部门表 C,收费项目目录 E,病人类型 Z" & _
            " Where B.病人类型=Z.名称(+) And A.病人ID=B.病人ID And Nvl(B.主页ID,0)<>0 And B.状态=0" & _
            " And B.出院科室ID=C.ID And B.护理等级ID=E.ID(+) And B.当前病区ID+0=[1] And (c.站点='" & gstrNodeNo & "' Or c.站点 is Null)" & _
            " And B.出院日期 Between [2] And [3] And Nvl(B.病案状态,0)<>5 And B.封存时间 is NULL"
    End If
    '转出病人:在院,医生和床号显示本科转出前的
    If chkScope(转出).Value = 1 Then
        strSQL = strSQL & IIf(strSQL <> "", " Union All ", "") & _
            "Select /*+ RULE */ Distinct 5 as 排序,Decode(Nvl(B.病案状态,0),0,999,B.病案状态) as 排序2,'转出病人' as 类型," & _
            " A.病人ID,B.主页ID,A.门诊号,B.住院号,A.姓名,A.性别,B.年龄,D.名称 as 科室,C.科室ID,C.经治医师 as 住院医师,B.责任护士,B.病案状态," & _
            " lpad(nvl(C.床号,' '),10,' ') as 床号,E.名称 as 护理等级,B.费别,B.当前病况,B.入院日期,B.出院日期,B.出院方式,B.病人类型," & _
            " B.状态,B.险类,A.就诊卡号,Nvl(B.路径状态,-1) 路径状态,trunc(sysdate)-trunc(b.入院日期)+1 as 住院天数,z.颜色" & _
            " From 病人信息 A,病案主页 B,病人变动记录 C,部门表 D,收费项目目录 E,病人类型 Z" & _
            " Where B.病人类型=Z.名称(+) And A.病人ID=B.病人ID And Nvl(B.主页ID,0)<>0 And B.护理等级ID=E.ID(+)" & _
            " And B.病人ID=C.病人ID And B.主页ID=C.主页ID" & _
            " And B.当前病区ID<>[1] And C.病区ID+0=[1] And C.科室ID=D.ID" & _
            " And Nvl(C.附加床位,0)=0 And C.终止原因 In(3,15) And C.终止时间 Between Sysdate-[4] And Sysdate" & _
            " And Nvl(B.状态,0)<>2 And Nvl(B.病案状态,0)<>5 And B.封存时间 is NULL"
    End If
    '再次过滤出有体温单文件的病人
    
    strSQL = "SELECT A.排序,A.排序2,A.类型,A.病人ID,A.主页ID,A.门诊号,A.住院号,A.姓名,A.性别,A.年龄,A.科室,A.科室ID,A.住院医师,A.责任护士,A.病案状态," & _
            " lpad(nvl(A.床号,' '),10,' ') as 床号,A.护理等级,A.费别,A.当前病况,A.入院日期,A.出院日期,A.出院方式,A.病人类型," & _
            " A.状态,A.险类,A.就诊卡号,A.路径状态,A.住院天数,A.颜色" & _
            " From (" & strSQL & ") A,病人护理文件 B" & _
            " Where A.病人ID=B.病人ID and A.主页ID=B.主页ID And nvl(B.婴儿,0)=0 And B.归档人 is null and B.结束时间 is null and B.格式ID=[5]"
    strSQL = strSQL & " Order by A.排序,A.床号,A.主页ID"
    
    Screen.MousePointer = 11
    On Error GoTo Errhand
    Set mrsPati = zlDatabase.OpenSQLRecord(strSQL, "提取病人列表", mlng病区ID, _
        CDate(Format(mdtOutBegin, "yyyy-MM-dd 00:00:00")), CDate(Format(mdtOutEnd, "yyyy-MM-dd 23:59:59")), _
        mintChange, mlng格式ID)
    Screen.MousePointer = 0
    Exit Sub
Errhand:
    Screen.MousePointer = 0
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cboUnit_Click()
    Dim ArrCode() As String
    Dim blnVisble As Boolean
    
    On Error GoTo Errhand
    
    If cboUnit.ListCount = 0 Then GoTo ErrNext
    If cboUnit.Tag = "" Then cboUnit.Tag = "0"

    ArrCode = Split(cboUnit.Tag, "[LPF]")
    mlng科室id = Val(cboUnit.ItemData(cboUnit.ListIndex))
    '获取科室性质
    Call Get科室性质(mlng科室id)
    
    '重新刷新过滤条件
    Call InitFilter
    
    '只有科室为妇产科才进行婴儿过滤
    blnVisble = (Val(ArrCode(cboUnit.ListIndex)) = 1)
    lblPati.Visible = blnVisble
    cboPati.Visible = blnVisble
    cboPati.Enabled = blnVisble
    cmdFilter.Left = IIf(blnVisble = True, cboPati.Left + cboPati.Width + 75, lblPati.Left)
    cmdAddUser.Left = cmdFilter.Left + cmdFilter.Width + 195
ErrNext:
    Call dtpDate_GotFocus
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cboUnit_KeyPress(KeyAscii As Integer)
    Call zlControl.CboMatchIndex(cboUnit.hWnd, KeyAscii)
End Sub

Private Sub cbo体温标识_Click()
    Call Save体温标识(VsfData.Row)
End Sub

Private Sub Save体温标识(ByVal lngCurRow As Long, Optional ByVal str体温标识 As String = "")
    Dim lngRow As Long
    Dim lng病人ID As Long, lng主页ID As Long, lng婴儿 As Long
    On Error GoTo Errhand
    '保存病人体温标识
    lng病人ID = Val(VsfData.TextMatrix(lngCurRow, c病人ID))
    lng主页ID = Val(VsfData.TextMatrix(lngCurRow, c主页ID))
    lng婴儿 = Val(VsfData.TextMatrix(lngCurRow, c婴儿))
    
    gstrSQL = "ZL_病人体温标识_Update(" & lng病人ID & "," & lng主页ID & "," & lng婴儿 & ",'" & IIf(str体温标识 = "", cbo体温标识.Text, str体温标识) & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, "保存病人体温标识")
    
    For lngRow = VsfData.FixedRows To VsfData.Rows - 1
        If lng病人ID = Val(VsfData.TextMatrix(lngRow, c病人ID)) And lng主页ID = Val(VsfData.TextMatrix(lngRow, c主页ID)) And lng婴儿 = Val(VsfData.TextMatrix(lngRow, c婴儿)) Then
            VsfData.TextMatrix(lngRow, c体温标识) = IIf(str体温标识 = "", cbo体温标识.Text, str体温标识)
        End If
    Next lngRow
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cbo体温标识_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call MoveNextCell
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngStartRow As Long, lngRow As Long, lngCol As Long, lngItemNO As Long, lngRow1 As Long
    Dim strKey As String, strFileds As String, strValues As String
    Dim strPart As String, strValue As String, strPart1 As String, strPart2 As String
    Dim strTime As String, strPatientTime As String, strInfo As String
    Dim arrValue() As Variant, arrCOL() As Variant, arrPart() As Variant, i As Long, intSate As Integer
    Dim arrID() As Variant
    Dim cbrCheck As CommandBarControl
    
    '体温标识
    Dim str体温标识 As String
    
    Select Case Control.Id
        Case conMenu_Edit_Send '数据同步
            If VsfData.Row < VsfData.FixedRows Then Exit Sub
            arrID = Array()
            '将选择列的数据 同步到其它空行列
            lngStartRow = VsfData.Row
            strTime = VsfData.TextMatrix(lngStartRow, c时间)
            str体温标识 = Trim(VsfData.TextMatrix(lngStartRow, c体温标识))
            strFileds = "ID|行号|项目序号|数据|部位|数据来源|状态"
            '提取本列的数据信息
            If mrsCell Is Nothing Then Exit Sub
            mrsCell.Filter = 0
            mrsCopy.Filter = 0
            '对于保存后的数据mrscell记录集可能为空,此处进行复制,用完后删除赋值过的Mrscell数据
            mrsCopy.Filter = "行号=" & lngStartRow
            Do While Not mrsCopy.EOF
                mrsCell.Filter = "ID='" & Nvl(mrsCopy!Id) & "' And 状态=1"
                If mrsCell.RecordCount = 0 Then
                    strValues = Nvl(mrsCopy!Id) & "|" & Val(Nvl(mrsCopy!行号)) & "|" & Val(Nvl(mrsCopy!项目序号)) & "|" & Nvl(mrsCopy!数据) & "|" & _
                        Nvl(mrsCopy!部位) & "|" & Val(Nvl(mrsCopy!数据来源)) & "|" & 0
                    Call Record_Update(mrsCell, strFileds, strValues, "ID|" & Nvl(mrsCopy!Id))
                    ReDim Preserve arrID(UBound(arrID) + 1)
                    arrID(UBound(arrID)) = Nvl(mrsCopy!Id)
                End If
                mrsCopy.MoveNext
            Loop
            
            mrsCell.Filter = 0
            mrsCell.Filter = "行号=" & lngStartRow & " And 状态=1"
            If mrsCell.RecordCount = 0 Then Exit Sub
            arrValue = Array()
            arrCOL = Array()
            arrPart = Array()
            Do While Not mrsCell.EOF
                lngCol = Val(Split(mrsCell!Id, ",")(1))
                lngItemNO = Val(Nvl(mrsCell!项目序号))
                strPart = Trim(Nvl(mrsCell!部位))
                strPart1 = Trim(GetPart(lngItemNO))
                strPart2 = ""
                If strPart <> strPart1 Then
                    strPart2 = strPart
                End If
                strValue = Val(Nvl(mrsCell!项目序号)) & "|" & Nvl(mrsCell!数据) & "|" & Nvl(mrsCell!部位) & "|0|1"
                ReDim Preserve arrValue(UBound(arrValue) + 1)
                arrValue(UBound(arrValue)) = strValue
                ReDim Preserve arrCOL(UBound(arrCOL) + 1)
                arrCOL(UBound(arrCOL)) = lngCol
                ReDim Preserve arrPart(UBound(arrPart) + 1)
                arrPart(UBound(arrPart)) = strPart2
            mrsCell.MoveNext
            Loop
            
            mrsCell.Filter = 0
            '开始复制数据 有数据的列不进行赋值
            For lngRow = VsfData.FixedRows To VsfData.Rows - 1
                If lngRow <> lngStartRow And VsfData.RowHidden(lngRow) = False Then
                    '如果用户已经输入时间 就不在进行时间的同步
                    If Trim(VsfData.TextMatrix(lngRow, c时间)) = "" And strTime <> "" Then
                        '用户没有录入时间 就需要检查同步的时间是否合法(不合法不进行复制 用户需要手工录入)
                        strPatientTime = VsfData.TextMatrix(lngRow, c日期)
                        If CheckDateTime(strTime, strPatientTime, strInfo) = True Then
                            VsfData.TextMatrix(lngRow, c时间) = strTime
                        End If
                    End If
                    
                    For i = 0 To UBound(arrValue)
                        strKey = lngRow & "," & Val(arrCOL(i))
                        mrsCell.Filter = "ID='" & strKey & "' And 状态=1"
                        If mrsCell.RecordCount = 0 Then
                            strValues = strKey & "|" & lngRow & "|" & CStr(arrValue(i))
                            strValue = Split(CStr(arrValue(i)), "|")(1)
                            If Trim(CStr(arrPart(i))) <> "" Then
                                strValue = CStr(arrPart(i)) & ":" & strValue
                            End If
                            Call Record_Update(mrsCell, strFileds, strValues, strKey)
                            VsfData.TextMatrix(lngRow, Val(arrCOL(i))) = strValue
                            mblnChage = True
                        End If
                    Next i
                    '同步体温标识
                    If str体温标识 <> "" And Trim(VsfData.TextMatrix(lngRow, c体温标识)) = "" Then
                        Call Save体温标识(lngRow, str体温标识)
                    End If
                End If
            Next lngRow
            
            '同上面所说用完后删除刚才复制的信息
            mrsCell.Filter = 0
            For i = 0 To UBound(arrID)
                mrsCell.Filter = "ID='" & CStr(arrID(i)) & "'"
                mrsCell.Delete
                mrsCell.Update
            Next i
            
            VsfData.Cell(flexcpAlignment, VsfData.FixedRows, c时间, VsfData.Rows - 1, VsfData.Cols - 1) = flexAlignCenterCenter
            Call InitCons
        Case conMenu_Edit_Clear '清除
           Call Edit_Clear
        Case conMenu_Edit_NewItem '添加空行
            If VsfData.Row < VsfData.FixedRows Then Exit Sub
            lngStartRow = VsfData.Row
            lngRow1 = VsfData.Row + 1
            VsfData.Rows = VsfData.Rows + 1
            VsfData.TextMatrix(VsfData.Rows - 1, c文件ID) = VsfData.TextMatrix(lngStartRow, c文件ID)
            VsfData.TextMatrix(VsfData.Rows - 1, c床号) = VsfData.TextMatrix(lngStartRow, c床号)
            VsfData.TextMatrix(VsfData.Rows - 1, c姓名) = VsfData.TextMatrix(lngStartRow, c姓名)
            VsfData.TextMatrix(VsfData.Rows - 1, c年龄) = VsfData.TextMatrix(lngStartRow, c年龄)
            VsfData.TextMatrix(VsfData.Rows - 1, c病人ID) = VsfData.TextMatrix(lngStartRow, c病人ID)
            VsfData.TextMatrix(VsfData.Rows - 1, c主页ID) = VsfData.TextMatrix(lngStartRow, c主页ID)
            VsfData.TextMatrix(VsfData.Rows - 1, c婴儿) = VsfData.TextMatrix(lngStartRow, c婴儿)
            VsfData.TextMatrix(VsfData.Rows - 1, c护理等级) = VsfData.TextMatrix(lngStartRow, c护理等级)
            VsfData.TextMatrix(VsfData.Rows - 1, c体温标识) = VsfData.TextMatrix(lngStartRow, c体温标识)
            VsfData.TextMatrix(VsfData.Rows - 1, c日期) = VsfData.TextMatrix(lngStartRow, c日期)
            VsfData.TextMatrix(VsfData.Rows - 1, c出院) = VsfData.TextMatrix(lngStartRow, c出院)
            lngStartRow = lngStartRow + 1
            
            For lngRow = VsfData.Rows - 2 To lngStartRow Step -1
                mrsCell.Filter = "行号=" & lngRow
                If mrsCell.RecordCount > 0 Then
                    mrsCell.MoveFirst
                    Do While Not mrsCell.EOF
                        strFileds = "ID|行号"
                        strKey = Nvl(mrsCell!Id)
                        lngCol = Val(Split(Nvl(mrsCell!Id, ","), ",")(1))
                        strValues = lngRow + 1 & "," & lngCol & "|" & lngRow + 1
                        Call Record_Update(mrsCell, strFileds, strValues, "ID|" & strKey)
                    mrsCell.MoveNext
                    Loop
                End If
                '更新mrsCopy数据集
                If Not mrsCopy Is Nothing Then
                    mrsCopy.Filter = "行号=" & lngRow
                    If mrsCopy.RecordCount > 0 Then
                        mrsCopy.MoveFirst
                        Do While Not mrsCopy.EOF
                            strFileds = "ID|行号"
                            strKey = Nvl(mrsCopy!Id)
                            lngCol = Val(Split(Nvl(mrsCopy!Id, ","), ",")(1))
                            strValues = lngRow + 1 & "," & lngCol & "|" & lngRow + 1
                            Call Record_Update(mrsCopy, strFileds, strValues, "ID|" & strKey)
                        mrsCopy.MoveNext
                        Loop
                    End If
                End If
                
                If Not mrsData Is Nothing Then
                    '更新恢复数据的行号
                    mrsData.Filter = "行号=" & lngRow
                    If mrsData.RecordCount > 0 Then
                        mrsData.MoveFirst
                        Do While Not mrsData.EOF
                            strFileds = "行号"
                            strValues = lngRow + 1
                            Call Record_Update(mrsData, strFileds, strValues, "行号|" & lngRow)
                        mrsData.MoveNext
                        Loop
                    End If
                End If
                
                VsfData.RowPosition(lngRow) = lngRow + 1
            Next lngRow
            VsfData.Cell(flexcpAlignment, VsfData.FixedRows, c时间, VsfData.Rows - 1, VsfData.Cols - 1) = flexAlignCenterCenter
            mblnChage = True
            VsfData.RowHeight(-1) = 300 + mintBigSize * 300 / 3
            VsfData.Select lngRow1, c时间
            VsfData.SetFocus
            '设置编辑颜色
            Call SetTabEditColor
    
        Case conMenu_Edit_Save '保存
            If Not SaveDate Then Exit Sub
        Case conMenu_Edit_Transf_Cancle '取消
            If Not EditCancle Then Exit Sub
        Case conMenu_Edit_Blankoff '清空表格所有行(不进行数据处理)
            Call InitCons  '隐藏编辑控件
            Call InitVariable(False)
            '清空现存记录集
            mrsCell.Filter = 0
            Do While Not mrsCell.EOF
               mrsCell.Delete
               mrsCell.Update
               mrsCell.MoveNext
            Loop
            mrsCopy.Filter = 0
            Do While Not mrsCopy.EOF
               mrsCopy.Delete
               mrsCopy.Update
               mrsCopy.MoveNext
            Loop
            mrsData.Filter = 0
            Do While Not mrsData.EOF
               mrsData.Delete
               mrsData.Update
               mrsData.MoveNext
            Loop
            mrsHistory.Filter = 0
            Do While Not mrsHistory.EOF
                mrsHistory.Delete
                mrsHistory.Update
                mrsHistory.MoveNext
            Loop
            
            Call ColligationTab(False)
            Call ColligationHistoryTab
            VsfData.Select VsfData.FixedRows, c时间
            Call AdjustRowFlag(VsfData, VsfData.FixedRows)
        Case conMenu_Edit_Compend * 10  '部位
            If VsfData.Row < VsfData.FixedRows Then Exit Sub
            strPart = Trim(Control.Parameter)
            lngRow = VsfData.Row
            lngCol = VsfData.Col
            lngItemNO = Val(VsfData.TextMatrix(0, lngCol))
            strValue = Trim(VsfData.TextMatrix(lngRow, lngCol))
            If InStr(1, strValue, ":") <> 0 Then
                strValue = Split(strValue, ":")(1)
            End If
            '如果本列vsfdata没有数值,则检查用户是否已经录入数据
            If picInput.Visible = True And mintType = 1 Then
                strValue = txtInput.Text
                strPart2 = GetPart(lngItemNO)
                txtInput.Tag = Trim(strPart)
                '更新部位菜单的选择项
                Call VsfData_AfterRowColChange(lngRow, c时间, lngRow, lngCol)
                For Each cbrCheck In mcbrToolBar.Controls(4).CommandBar.Controls
                    If cbrCheck.Parameter = Control.Parameter Then
                        cbrCheck.Checked = True
                    Else
                        cbrCheck.Checked = False
                    End If
                Next
                Exit Sub
            End If
            
            Call InitCons
            mintType = 0
            strFileds = "ID|行号|项目序号|数据|部位|数据来源|状态"
            If strValue <> "" And IsNumeric(strValue) Then
                strKey = lngRow & "," & lngCol
                mrsCell.Filter = "ID='" & strKey & "'"
                If mrsCell.RecordCount = 0 Then '将mrscopy的值复制过来
                    mrsCopy.Filter = "ID='" & strKey & "'"
                    strValues = Nvl(mrsCopy!Id) & "|" & Val(Nvl(mrsCopy!行号)) & "|" & Val(Nvl(mrsCopy!项目序号)) & "|" & Nvl(mrsCopy!数据) & "|" & _
                        Nvl(mrsCopy!部位) & "|" & Val(Nvl(mrsCopy!数据来源)) & "|" & 0
                    Call Record_Update(mrsCell, strFileds, strValues, "ID|" & Nvl(mrsCopy!Id))
                End If
                strFileds = "部位|状态"
                strValues = strPart
                If mrsCell.RecordCount > 0 Then
                    strPart1 = Trim(Nvl(mrsCell!部位))
                    strPart2 = GetPart(lngItemNO)
                    If strPart1 = "" Then strPart1 = strPart2
                    If strPart1 <> strPart Then
                        Call Record_Update(mrsCell, strFileds, strValues & "|1", "ID|" & strKey)
                        If strPart <> strPart2 And strPart <> "" Then
                            VsfData.TextMatrix(lngRow, lngCol) = strPart & ":" & strValue
                        Else
                            VsfData.TextMatrix(lngRow, lngCol) = strValue
                        End If
                        If picInput.Visible = True Then txtInput.Tag = Trim(strPart)
                        mblnChage = True
                        '更新部位菜单的选择项
                        Call VsfData_AfterRowColChange(lngRow, c时间, lngRow, lngCol)
                    End If
                End If
            End If
        Case conMenu_Edit_Append * 10, conMenu_Edit_Append * 10 + 1, conMenu_Edit_Append * 10 + 2, conMenu_Edit_Append * 10 + 3, conMenu_Edit_Append * 10 + 4, conMenu_Edit_Append * 10 + 5, conMenu_Edit_Append * 10 + 6
            If VsfData.Row < VsfData.FixedRows Then Exit Sub
            lngRow = VsfData.Row
            lngCol = VsfData.Col
            lngItemNO = Val(VsfData.TextMatrix(0, lngCol))
            strValue = Trim(VsfData.TextMatrix(lngRow, lngCol))
            
            If mintType = 1 And picInput.Visible = True Then strValue = txtInput.Text
            strPart = ""
            If InStr(1, "," & gint大便 & "," & gint入液 & ",", "," & lngItemNO & ",") = 0 Then Exit Sub
            Select Case Control.Id
                Case conMenu_Edit_Append * 10 + 1
                    strPart = "E"
                    If InStr(1, UCase(strValue), "/E") > 0 Then
                        strValue = Mid(strValue, 1, InStr(1, UCase(strValue), "/E") - 1)
                    End If
                    If InStr(1, UCase(strValue), "E") > 0 Then
                        strValue = Mid(strValue, 1, InStr(1, UCase(strValue), "E") - 1)
                    End If
                Case conMenu_Edit_Append * 10 + 2
                    strPart = "/E"
                    If InStr(1, UCase(strValue), "/E") > 0 Then
                        strValue = Mid(strValue, 1, InStr(1, UCase(strValue), "/E") - 1)
                    End If
                Case conMenu_Edit_Append * 10 + 3
                    strPart = "※"
                    strValue = ""
                Case conMenu_Edit_Append * 10 + 4
                    strPart = "☆"
                    strValue = ""
                Case conMenu_Edit_Append * 10 + 5
                    strPart = "C"
                    If InStr(1, UCase(strValue), "/C") > 0 Then
                        strValue = Mid(strValue, 1, InStr(1, UCase(strValue), "/C") - 1)
                    End If
                    If InStr(1, UCase(strValue), "C") > 0 Then
                        strValue = Mid(strValue, 1, InStr(1, UCase(strValue), "C") - 1)
                    End If
                Case conMenu_Edit_Append * 10 + 6
                    strPart = "/C"
                    If InStr(1, UCase(strValue), "/C") > 0 Then
                        strValue = Mid(strValue, 1, InStr(1, UCase(strValue), "/C") - 1)
                    End If
                Case conMenu_Edit_Append * 10
                    strPart = ""
                    If lngItemNO = gint大便 Then
                        For i = 0 To 4
                            Select Case i
                                Case 0
                                    strPart1 = "E"
                                Case 1
                                    strPart1 = "/"
                                Case 2
                                    strPart1 = "*"
                                Case 3
                                    strPart1 = "※"
                                Case 4
                                    strPart1 = "☆"
                            End Select
                            strValue = Replace(UCase(strValue), strPart1, "")
                        Next i
                    Else
                        strValue = Replace(UCase(Replace(UCase(strValue), "C", "")), "/", "")
                    End If
            End Select
            If IsNumeric(strValue) Then
                strValue = strValue
            Else
                strValue = ""
            End If
            strValue = strValue & Trim(strPart)
            If Left(strValue, 1) = "/" Then strValue = 1 & strValue
            
            If mintType = 1 And picInput.Visible = True Then
                If Len(txtInput.Text) > txtInput.MaxLength Then
                    RaiseEvent AfterRowColChange("选择的内容超过项目长度,请在护理记录项管理中设置项目长度.", True)
                    Exit Sub
                End If
                txtInput.Text = strValue
                For Each cbrCheck In mcbrToolBar.Controls(5).CommandBar.Controls
                    If cbrCheck.Id = Control.Id Then
                        cbrCheck.Checked = True
                    Else
                        cbrCheck.Checked = False
                    End If
                Next

                Exit Sub
            End If
            
            Call InitCons
            mintType = 0
            '更新数据变动记录集
            strFileds = "ID|行号|项目序号|数据|部位|数据来源|状态"
            strKey = lngRow & "," & lngCol
            mrsCell.Filter = "ID='" & strKey & "'"
            If mrsCell.RecordCount = 0 Then '将mrscopy的值复制过来
                mrsCopy.Filter = "ID='" & strKey & "'"
                If mrsCopy.RecordCount > 0 Then
                    strValues = Nvl(mrsCopy!Id) & "|" & Val(Nvl(mrsCopy!行号)) & "|" & Val(Nvl(mrsCopy!项目序号)) & "|" & Nvl(mrsCopy!数据) & "|" & _
                        Nvl(mrsCopy!部位) & "|" & Val(Nvl(mrsCopy!数据来源)) & "|" & 0
                    Call Record_Update(mrsCell, strFileds, strValues, "ID|" & Nvl(mrsCopy!Id))
                    intSate = 1
                Else
                    '添加新的记录
                    If Trim(strValue) = "" Then Exit Sub
                    strValues = strKey & "|" & lngRow & "|" & lngItemNO & "|" & strValue & "|" & _
                        "" & "|" & 0 & "|" & 1
                    Call Record_Update(mrsCell, strFileds, strValues, "ID|" & strKey)
                    GoTo ErrGO
                End If
            Else
                intSate = IIf(Trim(strValue) = "", 3, 1)
            End If
            mrsCell.Filter = "ID='" & strKey & "'"
            strFileds = "数据|状态"
            strValues = strValue & "|" & intSate
            If mrsCell.RecordCount > 0 Then
                If Trim(Nvl(mrsCell!数据)) <> strValue Then
                    Call Record_Update(mrsCell, strFileds, strValues, "ID|" & strKey)
ErrGO:
                    VsfData.TextMatrix(lngRow, lngCol) = strValue
                    mblnChage = True
                    Call VsfData_AfterRowColChange(lngRow, c时间, lngRow, lngCol)
                End If
            End If
            
        Case conMenu_Help_Help '帮助
            RaiseEvent UsrHelp
        Case conMenu_File_Exit '退出
            RaiseEvent UsrExit
    End Select
End Sub

Private Sub cbsThis_Resize()
    Dim lngScaleLeft As Long, lngScaleTop As Long, lngScaleRight As Long, lngScaleBottom As Long
    Call cbsThis.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
    
    Err = 0: On Error Resume Next
    
    If Not mblnInit Then picSplit.Top = ScaleHeight - 3000
    
    picMain.Move lngScaleLeft, lngScaleTop, lngScaleRight, picSplit.Top - lngScaleTop
    VsfData.Move lngScaleLeft + 100, 100, lngScaleRight - lngScaleLeft - 100 * 2
    VsfData.Height = picMain.Height - VsfData.Top
    
    picHistory.Move lngScaleLeft, picSplit.Top + picSplit.Height, lngScaleRight, lngScaleBottom - picSplit.Top
    picDay.Left = VsfData.Left
    picDay.Top = 60
    picDay.Width = VsfData.Width
    vsfHistory.Left = VsfData.Left
    vsfHistory.Top = picDay.Top + picDay.Height + 60
    vsfHistory.Height = picHistory.Height - vsfHistory.Top - 50
    vsfHistory.Width = VsfData.Width
    picSplit.Left = lngScaleLeft
    picSplit.Width = picMain.Width
    picNull.Move lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom
    With lblInfo
        .Top = (picNull.Height - .Height) / 2
        .Left = (picNull.Width - .Width) / 2
    End With
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.Id
        Case conMenu_Edit_Save, conMenu_Edit_Transf_Cancle
            Control.Enabled = mblnChage
        Case conMenu_Edit_NewItem '添加空行
            Control.Enabled = (mblnNullRow And mblnInit)
        Case conMenu_Edit_Compend '部位
            Control.Enabled = (mcbrMenuBar部位.CommandBar.Controls.Count <> 0)
        Case conMenu_Edit_Clear, conMenu_Edit_Send '清除 数据同步
            Control.Enabled = (mblnClearRow And mblnInit)
        Case conMenu_Edit_Blankoff '清空表格所有行(不进行数据处理)
            Control.Enabled = mblnNullRow
            picNull.Visible = Not mblnNullRow
            If mblnNullRow <> (VsfData.ScrollBars = flexScrollBarBoth) Then
                VsfData.ScrollBars = IIf(mblnNullRow, flexScrollBarBoth, flexScrollBarNone)
            End If
        Case conMenu_Edit_Append * 10 + 0, conMenu_Edit_Append
            Control.Enabled = is大便或入液(1) Or is大便或入液(2)
        Case conMenu_Edit_Append * 10 + 1, conMenu_Edit_Append * 10 + 2, conMenu_Edit_Append * 10 + 3, conMenu_Edit_Append * 10 + 4
            Control.Enabled = is大便或入液(1)
        Case conMenu_Edit_Append * 10 + 5, conMenu_Edit_Append * 10 + 6
            Control.Enabled = is大便或入液(2)
        Case conMenu_View_LocationItem
            'dtpDate.Enabled = Not mblnInit
    End Select
End Sub

Private Sub Edit_Clear()
'---------------------------------------
'功能:清除数据信息
'---------------------------------------
    Dim lngStartRow As Long, lngRow As Long, lngCol As Long, lngItemNO As Long
    Dim lngRow1 As Long
    Dim strKey As String, strFileds As String, strValues As String
    
    '清除已经录入的数据信息
    If VsfData.Row < VsfData.FixedRows Then Exit Sub
    strFileds = "ID|行号|项目序号|数据|部位|数据来源|状态"
    On Error GoTo Errhand
    
    lngRow = VsfData.Row
    lngRow1 = lngRow
    '清除列的数据信息
    For lngCol = c时间 To VsfData.Cols - 1
        VsfData.TextMatrix(lngRow, lngCol) = ""
    Next lngCol
   
    '清除记录集信息
    mrsCell.Filter = "行号=" & lngRow
    mrsCell.Sort = "ID"
    Do While Not mrsCell.EOF
        mrsCell.Delete
        mrsCell.Update
        mblnChage = True
    mrsCell.MoveNext
    Loop
    
    mrsCell.Filter = "行号=" & lngRow
    mrsCopy.Filter = "行号=" & lngRow
    If mrsCopy.RecordCount > 0 Then
        Do While Not mrsCopy.EOF
            strValues = Nvl(mrsCopy!Id) & "|" & lngRow & "|" & Val(Nvl(mrsCopy!项目序号)) & "|"
            If InStr(1, ",0,9", "," & mrsCopy!数据来源 & ",") <> 0 Then
                strValues = strValues & "|" & Nvl(mrsCopy!部位) & "|" & Nvl(mrsCopy!数据来源) & "|1"
                Call Record_Update(mrsCell, strFileds, strValues, "ID|" & Nvl(mrsCopy!Id))
            Else
                strValues = strValues & Nvl(mrsCopy!数据) & "|" & Nvl(mrsCopy!部位) & "|" & Nvl(mrsCopy!数据来源) & "|0"
                Call Record_Update(mrsCell, strFileds, strValues, "ID|" & Nvl(mrsCopy!Id))
            End If
            mrsCopy.MoveNext
        Loop
'        mrsCell.Filter = 0
'        Call OutputRsData(mrsCell, True)
        
        mrsData.Filter = "行号=" & lngRow
        If mrsData.RecordCount > 0 Then
            VsfData.TextMatrix(lngRow, c时间) = mrsData.Fields(c时间).Value
            mrsData!删除 = 1
            mrsData.Update
        End If
        '删除后添加一行空行,隐藏删除行
        VsfData.RowHidden(lngRow) = True
        lngStartRow = lngRow
        VsfData.Rows = VsfData.Rows + 1
        VsfData.TextMatrix(VsfData.Rows - 1, c文件ID) = VsfData.TextMatrix(lngStartRow, c文件ID)
        VsfData.TextMatrix(VsfData.Rows - 1, c床号) = VsfData.TextMatrix(lngStartRow, c床号)
        VsfData.TextMatrix(VsfData.Rows - 1, c姓名) = VsfData.TextMatrix(lngStartRow, c姓名)
        VsfData.TextMatrix(VsfData.Rows - 1, c年龄) = VsfData.TextMatrix(lngStartRow, c年龄)
        VsfData.TextMatrix(VsfData.Rows - 1, c病人ID) = VsfData.TextMatrix(lngStartRow, c病人ID)
        VsfData.TextMatrix(VsfData.Rows - 1, c主页ID) = VsfData.TextMatrix(lngStartRow, c主页ID)
        VsfData.TextMatrix(VsfData.Rows - 1, c婴儿) = VsfData.TextMatrix(lngStartRow, c婴儿)
        VsfData.TextMatrix(VsfData.Rows - 1, c护理等级) = VsfData.TextMatrix(lngStartRow, c护理等级)
        VsfData.TextMatrix(VsfData.Rows - 1, c日期) = VsfData.TextMatrix(lngStartRow, c日期)
        VsfData.TextMatrix(VsfData.Rows - 1, c出院) = VsfData.TextMatrix(lngStartRow, c出院)
        
        lngStartRow = lngStartRow + 1
        For lngRow = VsfData.Rows - 2 To lngStartRow Step -1
            '更新原始记录集
            mrsCell.Filter = "行号=" & lngRow
            If mrsCell.RecordCount > 0 Then
                mrsCell.MoveFirst
                Do While Not mrsCell.EOF
                    strFileds = "ID|行号"
                    strKey = Nvl(mrsCell!Id)
                    lngCol = Val(Split(Nvl(mrsCell!Id, ","), ",")(1))
                    strValues = lngRow + 1 & "," & lngCol & "|" & lngRow + 1
                    Call Record_Update(mrsCell, strFileds, strValues, "ID|" & strKey)
                mrsCell.MoveNext
                Loop
            End If
            '更新mrsCopy数据集
            mrsCopy.Filter = "行号=" & lngRow
            If mrsCopy.RecordCount > 0 Then
                mrsCopy.MoveFirst
                Do While Not mrsCopy.EOF
                    strFileds = "ID|行号"
                    strKey = Nvl(mrsCopy!Id)
                    lngCol = Val(Split(Nvl(mrsCopy!Id, ","), ",")(1))
                    strValues = lngRow + 1 & "," & lngCol & "|" & lngRow + 1
                    Call Record_Update(mrsCopy, strFileds, strValues, "ID|" & strKey)
                mrsCopy.MoveNext
                Loop
            End If

            '更新恢复数据的行号
            mrsData.Filter = "行号=" & lngRow
            If mrsData.RecordCount > 0 Then
                mrsData.MoveFirst
                Do While Not mrsData.EOF
                    strFileds = "行号"
                    strValues = lngRow + 1
                    Call Record_Update(mrsData, strFileds, strValues, "行号|" & lngRow)
                mrsData.MoveNext
                Loop
            End If
            VsfData.RowPosition(lngRow) = lngRow + 1
        Next lngRow
        VsfData.Cell(flexcpAlignment, VsfData.FixedRows, c时间, VsfData.Rows - 1, VsfData.Cols - 1) = flexAlignCenterCenter
        lngRow1 = lngRow1 + 1
    End If
    VsfData.RowHeight(-1) = 300 + mintBigSize * 300 / 3
    VsfData.Select lngRow1, c时间
    VsfData.SetFocus
    
    '设置编辑颜色
    Call SetTabEditColor
    mblnChage = True
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function EditCancle() As Boolean
'---------------------------------------------------
'功能:用户取消操作
'---------------------------------------------------
    '用户取消操作时清空录入数据信息,病人列表信息不变(取出重复项)
    Dim lngRow As Long, lngCol As Long, lngRows As Long, lngCOls As Long
    Dim lng行号 As Long
    Dim rsPati As New ADODB.Recordset
    Dim strFileds As String, strValues As String
    Dim lngID As Long
    
    On Error GoTo Errhand
    
    VsfData.Cell(flexcpText, lngRow, c时间, VsfData.Rows - 1, VsfData.Cols - 1) = ""
        
    Set mrsCell = New ADODB.Recordset
    gstrFields = "ID," & adLongVarChar & ",40|行号," & adDouble & ",18|项目序号," & adDouble & ",18|数据," & adLongVarChar & ",40|" & _
        "部位," & adLongVarChar & ",20|数据来源," & adDouble & ",1|来源ID," & adDouble & ",18|共用," & adDouble & ",1|显示," & adDouble & ",1|" & _
        "修改," & adDouble & ",1|状态," & adDouble & ",1"
    Call Record_Init(mrsCell, gstrFields)
    
    If mblnSaveData = False Then
        Call Record_Init(mrsCopy, gstrFields)
    End If
    
    gstrFields = "ID|行号|项目序号|数据|部位|数据来源|来源ID|共用|显示|修改|状态"
    '重新加载表格信息
    Call ColligationTab(False)
    
    '开始恢复数据
    mrsData.Filter = 0
    If mrsData.RecordCount > 0 Then mrsData.MoveFirst
    lngCOls = VsfData.Cols - 1
    lngRows = mrsData.RecordCount - 1
    
    For lngRow = 0 To lngRows
        If lngRow > VsfData.Rows - VsfData.FixedRows - 1 Then VsfData.Rows = VsfData.Rows + 1
        For lngCol = c文件ID To lngCOls
            If lngCol = c姓名 Then
                VsfData.TextMatrix(VsfData.FixedRows + lngRow, lngCol) = IIf(Val(Nvl(mrsData.Fields(c婴儿).Value)) > 0, Space(4), "") & Nvl(mrsData.Fields(lngCol).Value)
            Else
                VsfData.TextMatrix(VsfData.FixedRows + lngRow, lngCol) = Nvl(mrsData.Fields(lngCol).Value)
            End If
        Next
        If mrsData!删除 = 1 Then VsfData.RowHidden(VsfData.FixedRows + lngRow) = True
        '重新排列行号
        lng行号 = Val(Nvl(mrsData!行号))
        mrsCopy.Filter = "行号=" & lng行号
        Do While Not mrsCopy.EOF
            mrsCopy!行号 = lngRow + VsfData.FixedRows
            mrsCopy.Update
        mrsCopy.MoveNext
        Loop
        mrsData!行号 = lngRow + VsfData.FixedRows
        mrsData.MoveNext
    Next
    
    If VsfData.Rows <= VsfData.FixedRows Then VsfData.Rows = VsfData.Rows + 1
    VsfData.RowHeight(-1) = 300 + mintBigSize * 300 / 3
    VsfData.Cell(flexcpAlignment, VsfData.FixedRows, c时间, VsfData.Rows - 1, VsfData.Cols - 1) = flexAlignCenterCenter
    '设置编辑颜色
    Call SetTabEditColor
    
    VsfData.Select VsfData.FixedRows, c时间
    VsfData.SetFocus
    
    mblnChage = False
    mblnShow = False
    mbln出院 = False
    
    Call InitCons
    
    EditCancle = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckEditData() As Boolean
'-------------------------------------------------
'功能:检查是否已经发生数据编辑
'-------------------------------------------------
    Dim lngCOls As Long, lngCol As Long
    On Error GoTo Errhand
    
    '对于病人列表全是手工添加的病人,如果该时间段内没有录入任何数据这可以进行日期的切换
    If Format(mstrDate, "YYYY-MM-DD") = Format(dtpDate.Value, "YYYY-MM-DD") Then Exit Function
    
    If Not mblnRefreshData Then
        If mblnSaveData = True Then
            '全是手工添加的病人，保存后切换日期就只保留病人信息,数据信息全部清空
            lngCOls = VsfData.Cols - 1
            mrsData.Filter = 0
            If mrsData.RecordCount > 0 Then mrsData.MoveFirst
            Do While Not mrsData.EOF
                For lngCol = c时间 To lngCOls
                    mrsData.Fields(lngCol) = ""
                Next lngCol
                mrsData("删除") = 0
                mrsData.Update
            Loop
        End If
        mblnSaveData = False
        Call EditCancle
        Exit Function
'        If mrsCell Is Nothing Then Exit Function
'        mrsCell.Filter = 0
'        mrsCell.Filter = "状态<>3"
'        If mrsCell.RecordCount = 0 Then
'            VsfData.Cell(flexcpText, VsfData.FixedRows, c时间, VsfData.Rows - 1, VsfData.Cols - 1) = ""
'            mblnChage = False
'            Call InitCons
'            mblnSaveData = False
'            Exit Function
'        Else
'            'MsgBox "对于已经发生数据的日期进行修改时,请先点击取消按钮,在进行日期切换此操作！", vbInformation, gstrSysName
'            Call EditCancle
'            mblnSaveData = False
'            Exit Function
'        End If
    Else '如果病人列表中包含过滤出来的病人则需要在改变日期后重新刷新病人信息
'        If MsgBox("由于被切换的日期已经发生数据,如果继续操作需要手工重新过滤/添加病人,请问是否继续?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
'            VsfData.Rows = VsfData.FixedRows + 1
'            VsfData.Cell(flexcpText, VsfData.FixedRows, 0, VsfData.Rows - 1, VsfData.Cols - 1) = ""
'            Call InitCons
'            Call InitVariable
'            If cmdFilter.Enabled = True Then cmdFilter.SetFocus
'            Exit Function
'        End If
        '直接重新过滤信息
        Call cmdFilter_Click
        Exit Function
    End If
    CheckEditData = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetPart(ByVal lng项目序号 As Long) As String
'功能:提取默认的体温部位
    Dim strPart As String
    mrsPart.Filter = "项目序号=" & lng项目序号 & " and 缺省项=1"
    If mrsPart.RecordCount > 0 Then strPart = Trim(zlCommFun.Nvl(mrsPart("部位")))
    GetPart = strPart
End Function

Private Sub chkPati_Click(Index As Integer)
    Dim i As Integer
    Dim blnTrue As Boolean
    
    For i = 0 To chkPati.Count - 1
        If i <> Index Then
            blnTrue = (chkPati(i).Value <> 0)
        End If
    Next i
    
    If Not blnTrue And chkPati(Index).Value = 0 Then chkPati(IIf(Index = 0, 1, 0)).Value = 1
    blnTrue = (chkPati(IIf(Index = 0, 1, 0)).Value <> 0)
    
    If blnTrue And chkPati(Index).Value <> 0 Then
        mintPatiNo = 0
    ElseIf blnTrue Then
        mintPatiNo = IIf(Index = 0, 2, 1)
    Else
        mintPatiNo = IIf(Index = 0, 1, 2)
    End If
    
    For i = 0 To cboPati.ListCount - 1
        If mintPatiNo = cboPati.ItemData(i) Then Call zlControl.CboSetIndex(cboPati.hWnd, i)
    Next i
End Sub

Private Sub chkScope_Click(Index As Integer)
    Dim blnEnable As Boolean
    blnEnable = (chkScope(Index).Value = 1)
    
    Select Case Index
        Case 出院
            dtpB(0).Enabled = blnEnable
            dtpE(0).Enabled = blnEnable
        Case 转出
            txtChange.Enabled = blnEnable
    End Select
End Sub

Private Sub chkScope_Validate(Index As Integer, Cancel As Boolean)
    Dim i As Integer
    Dim blnAll As Boolean
    
    For i = 0 To 3
        If blnAll = False Then
            blnAll = (chkScope(i).Value = 1)
        End If
    Next i
    
    If blnAll = False Then
        chkScope(Index).Value = 1
        RaiseEvent AfterRowColChange("添加病人时类型至少需要选择一项", True)
        Cancel = True
    End If
End Sub

Private Sub chkSwitch_Click()
    '开始进行病人批量选择
    Dim intValue As Integer
    Dim lngLoop As Long
    Dim objRow As ReportRow
    Dim arrIndex()
    Dim i As Integer
    
    If mblnChkClick = True Then mblnChkClick = False: Exit Sub
    
    intValue = chkSwitch.Value
    
    arrIndex = Array()
    '记录展开的组的索引
    For Each objRow In rptPati.Rows
       If objRow.GroupRow Then
           If objRow.Expanded = True Then
               ReDim Preserve arrIndex(UBound(arrIndex) + 1)
               arrIndex(UBound(arrIndex)) = objRow.Index
           End If
       End If
    Next
    
    '进行批量选择
    For Each objRow In rptPati.Rows
        If objRow.GroupRow And objRow.Childs.Count > 0 Then
            For lngLoop = 0 To objRow.Childs.Count - 1
                If Not (objRow.Childs(lngLoop).Record Is Nothing) Then
                    'If Trim(objRow.Childs(lngLoop).Record.Item(c_出院日期).Value) <> "" Then Exit For
                    objRow.Childs(lngLoop).Record.Item(c_选择).Checked = IIf(intValue = 0, False, True)
                End If
            Next lngLoop
        End If
    Next
    
    rptPati.Populate
    
    '还原展开的组的
    For Each objRow In rptPati.Rows
       If objRow.GroupRow Then
           objRow.Expanded = False
           For i = 0 To UBound(arrIndex)
               If objRow.Index = Val(arrIndex(i)) Then
                   objRow.Expanded = True
                   Exit For
               End If
           Next i
       End If
    Next
End Sub

Private Sub chkSwitch_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    zlCommFun.ShowTipInfo chkSwitch.hWnd, "对病人进行批量全选/反选操作(不包含出院病人)"
End Sub

Private Sub cmdAddUser_Click()
    Dim lngColor As Long
    Dim lngLoop As Long
    Dim objRow As ReportRow
    Dim objRecord As ReportRecord
    Dim objItem As ReportRecordItem
    Dim strPatient As String '病人列表信息
    Dim lngRow As Long, lngID As Long 'VSF选择的病人ID
    Dim lngLeft As Long, lngTop  As Long, lngRight As Long, lngBottom As Long
    Dim ArrCode() As String
    Dim blnVisible As Boolean
    
    If mrsPati.State = 0 Then
        Call RefreshPatiList '刷新病人列表信息
    End If
    
    mrsPati.Filter = ""
    If rptPati.Records.Count = 0 And mrsPati.RecordCount > 0 Then
        '显示病人列表供选择
        With mrsPati
            .MoveFirst
            
            Do While Not .EOF
                Set objRecord = rptPati.Records.Add()
                objRecord.Tag = CStr(!病人ID & "," & !主页ID)
                Set objItem = objRecord.AddItem("")
                objItem.HasCheckbox = True
                objItem.Checked = False
                
                Set objItem = objRecord.AddItem(""): objItem.Icon = IIf(!性别 = "男", 1, 0)
                Set objItem = objRecord.AddItem(CStr(!排序))
                objItem.Caption = CStr(!排序 & !类型)
                Set objItem = objRecord.AddItem(CStr(!排序 & !类型))
                objItem.Caption = CStr(!排序 & !类型)
                
                Set objItem = objRecord.AddItem(LPAD(Nvl(!床号), 10, " "))
                objItem.Caption = Trim(Nvl(!床号, " "))
                objRecord.AddItem Val(!病人ID)
                objRecord.AddItem Val(!主页ID)
                objRecord.AddItem CStr(Nvl(!姓名))
                objRecord.AddItem CStr(Nvl(!年龄))
                Set objItem = objRecord.AddItem(CStr(Nvl(!住院号)))
                objItem.Caption = Nvl(!住院号, " ")
                
                '53881:刘鹏飞,2012-09-19,入出院日期应该加上时分秒，避免检查录入时间发生错误
                Set objItem = objRecord.AddItem(Format(!入院日期, "yyyy-MM-dd HH:mm:ss"))
                objItem.Caption = Format(!入院日期, "yyyy-MM-dd  HH:mm:ss")
                Set objItem = objRecord.AddItem(Format(!出院日期, "yyyy-MM-dd HH:mm:ss"))
                objItem.Caption = Format(!出院日期, "yyyy-MM-dd  HH:mm:ss")
                
                '提取病人类型的颜色
                lngColor = Nvl(!颜色, 0)
                If lngColor <> 0 Then objRecord.Item(c_姓名).ForeColor = lngColor
                
                .MoveNext
            Loop
            
            .MoveFirst
        End With
    End If
    
    If cboUnit.Tag = "" Then cboUnit.Tag = "0"
    ArrCode = Split(cboUnit.Tag, "[LPF]")
    blnVisible = (Val(ArrCode(0)) = 1)
    chkPati(0).Visible = blnVisible
    chkPati(1).Visible = blnVisible
    
    Select Case mintPatiNo
        Case 1
            chkPati(0).Value = 1
            chkPati(1).Value = 0
        Case 2
            chkPati(0).Value = 0
            chkPati(1).Value = 1
        Case Else
            chkPati(0).Value = 1
            chkPati(1).Value = 1
    End Select
    chkPati(0).Enabled = (Not mblnNullRow And blnVisible)
    chkPati(1).Enabled = (Not mblnNullRow And blnVisible)
    '调整坐标
    rptPati.Populate '缺省不选中任何行
    picPati.Left = cmdAddUser.Left + 60
    picPati.Top = picMain.Top
    picPati.Visible = True
    
    With chkSwitch
        .Value = 0
        .Top = rptPati.Top + 100
        .Left = rptPati.Left + (rptPati.Columns(c_选择).Width * Screen.TwipsPerPixelX - .Width) / 2
        .ZOrder 0
    End With
    
    cmdFilterUserCancle.Left = picPati.ScaleWidth - cmdFilterUserCancle.Width - 100
    cmdFilterUserOk.Left = cmdFilterUserCancle.Left - cmdFilterUserCancle.Width - 60
    
    strPatient = ""
    lngRow = 0
    If mblnInit = True Then
        If VsfData.Cols >= RootCol Then
            For lngRow = VsfData.FixedRows To VsfData.Rows - 1
                strPatient = strPatient & "," & VsfData.TextMatrix(lngRow, c病人ID)
                If VsfData.Row = lngRow Then
                    lngID = Val(VsfData.TextMatrix(lngRow, c病人ID))
                End If
            Next lngRow
        End If
    End If
    
    If Left(strPatient, 1) = "," Then strPatient = Mid(strPatient, 2)
    
    '清空所有选择列
    For lngLoop = 0 To rptPati.Rows.Count - 1
         If Not (rptPati.Rows(lngLoop).Record Is Nothing) Then
            rptPati.Rows(lngLoop).Record.Item(c_选择).Checked = False
         End If
    Next
    
    '如果已经进行了刷新 就勾选已经过滤出来的病人
'    For lngLoop = 0 To rptPati.Rows.Count - 1
'         If Not (rptPati.Rows(lngLoop).Record Is Nothing) Then
'             If InStr(1, "," & strPatient & ",", "," & Val(rptPati.Rows(lngLoop).Record.Item(c_病人ID).Value) & ",") <> 0 Then
'                 rptPati.Rows(lngLoop).Record.Item(c_选择).Checked = True
'             Else
'                rptPati.Rows(lngLoop).Record.Item(c_选择).Checked = False
'             End If
'         End If
'     Next
    
    '选中当前病人(先折叠组的话,Rows.Count只有组的个数了,所以先定位,再折叠)
    For lngLoop = 0 To rptPati.Rows.Count - 1
        If Not (rptPati.Rows(lngLoop).Record Is Nothing) Then
            If lngID <> 0 Then
                If Val(rptPati.Rows(lngLoop).Record.Item(C_病人ID).Value) = lngID Then
                    Set rptPati.FocusedRow = rptPati.Rows(lngLoop)
                    Exit For
                End If
            Else
                 Set rptPati.FocusedRow = rptPati.Rows(lngLoop)
                 Exit For
            End If
        End If
    Next
    
    '折叠所有组 (选中病人那一组不折叠)
    For Each objRow In rptPati.Rows
        If objRow.GroupRow And objRow.Index <> rptPati.FocusedRow.ParentRow.Index Then
            objRow.Expanded = False
        End If
    Next
    
    chkSwitch.Enabled = (rptPati.Records.Count > 0)
    
    If rptPati.Records.Count > 0 Then rptPati.FocusedRow.EnsureVisible
    rptPati.SetFocus
End Sub

Private Sub cmdFilter_Click()
'根据用户设置的过滤条件过滤病人信息
    mblnInit = False
    mlng科室id = Val(cboUnit.ItemData(cboUnit.ListIndex))

    mstrDate = Format(dtpDate.Value, "YYYY-MM-DD")
    Call InitCons  '隐藏编辑控件
    Call InitVariable '清除常量信息
    Call zlRefreshDate '刷新数据
    mblnInit = True
    
    '保存数据集
    Call Data_Save
End Sub

Private Function zlRefreshDate(Optional blnFillPage As Boolean = True) As Boolean
'-----------------------------------------------------
'功能:刷新数据
'blnFillPage 是否重新提取病人信息
'-----------------------------------------------------
    Dim ArrCode() As String
    Dim blnVisible As Boolean
    
    '只有科室为妇产科才进行婴儿过滤
    If cboUnit.Tag = "" Then cboUnit.Tag = "0"
    ArrCode = Split(cboUnit.Tag, "[LPF]")
    blnVisible = (Val(ArrCode(cboUnit.ListIndex)) = 1)
    If blnVisible = True Then
        mintPatiNo = cboPati.ItemData(cboPati.ListIndex)
    Else
        mintPatiNo = 1
    End If
    '提取体温曲线数据
    Call InitCurveDate
    '绑定表格列
    Call ColligationTab(blnFillPage)
    '处理时历史数据列
    Call ColligationHistoryTab
End Function

Private Sub InitCurveDate()
'----------------------------------------
'提取日常要编辑的体温数据
'----------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strTmp As String
    Dim i As Integer
    Dim strFind As String
    On Error GoTo Errhand
        
        '初始化数据记录集
        If Not (mrsCell Is Nothing) Then Set mrsCell = Nothing
        If Not (mrsCopy Is Nothing) Then Set mrsCopy = Nothing
        Set mrsCell = New ADODB.Recordset
        Set mrsCopy = New ADODB.Recordset
        
        gstrFields = "ID," & adLongVarChar & ",40|行号," & adDouble & ",18|项目序号," & adDouble & ",18|数据," & adLongVarChar & ",40|" & _
            "部位," & adLongVarChar & ",20|数据来源," & adDouble & ",1|来源ID," & adDouble & ",18|共用," & adDouble & ",1|显示," & adDouble & ",1|" & _
            "修改," & adDouble & ",1|状态," & adDouble & ",1"
        Call Record_Init(mrsCell, gstrFields)
        Call Record_Init(mrsCopy, gstrFields)
        Call Record_Init(mrsHistory, gstrFields)
        
        gstrFields = "ID|行号|项目序号|数据|部位|数据来源|来源ID|共用|显示|修改|状态"
        
        mstrTabHead = "|文件ID|床号|姓名|年龄|病人ID|主页ID|婴儿|记录ID|护理等级|体温标识|出院|日期|时间"
        mstrItemNo = ""
        
        Select Case mintPatiNo
            Case 1
                strFind = " And instr('0,1',B.适用病人)<>0"
            Case 2
                strFind = " And instr('0,2',B.适用病人)<>0"
            Case Else
                strFind = ""
        End Select
        '提取要录入的表格数据信息
        mstrSQL = "SELECT /*+ RULE */ A.项目序号,DECODE(A.项目序号,4,'血压',A.记录名) || DECODE(nvl(A.单位,''),'','', '(' || A.单位 || ')') 项目名称,A.排列序号,B.分组名  FROM 体温记录项目 A,诊治所见项目 C, 护理记录项目 B" & vbNewLine & _
                "WHERE  B.项目ID=C.ID(+) AND A.项目序号=B.项目序号 AND NVL(B.应用方式,0)=1 And A.项目序号<>5 And B.项目性质=1 " & strFind & vbNewLine & _
                "AND (B.适用科室=1 OR (B.适用科室=2 AND EXISTS (SELECT 1 FROM 护理适用科室 D," & vbNewLine & _
                "Table(Cast(f_num2list([1]) As zlTools.t_Numlist)) E WHERE D.项目序号=B.项目序号 AND D.科室ID=E.Column_Value)))" & vbNewLine & _
                "ORDER BY Decode(A.项目序号,1,0,1),A.排列序号"

        If mlng科室id = -1 Then
            For i = 1 To cboUnit.ListCount - 1
                strTmp = strTmp & "," & cboUnit.ItemData(i)
            Next i
        Else
            strTmp = CStr(mlng科室id)
        End If
        
        If Left(strTmp, 1) = "," Then strTmp = Mid(strTmp, 2)
        
        Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "体温单批量录入", strTmp)
        
        '提取体温曲线项目
        rsTemp.Filter = "分组名='1)体温曲线项目'"
        With rsTemp
            Do While Not .EOF
                mstrTabHead = mstrTabHead & "|" & Nvl(!项目名称)
                mstrItemNo = mstrItemNo & "|" & Val(Nvl(!项目序号))
            .MoveNext
            Loop
        End With
        
        If Left(mstrItemNo, 1) = "|" Then mstrItemNo = Mid(mstrItemNo, 2)
        '提取收缩压舒张压
        rsTemp.Filter = "项目序号=4"
        'mrsItems.Filter="项目序号=4"
        If rsTemp.RecordCount > 0 Then '收缩压和舒张压必须同时存在
            mstrTabHead = mstrTabHead & "|" & Nvl(rsTemp!项目名称)    ' "|血压(" & Nvl(mrsItems!项目单位) & ")"
            mstrItemNo = mstrItemNo & "|4"
        End If
        
        '提取剩余体温表格项目
        rsTemp.Filter = "分组名<>'1)体温曲线项目' and 项目序号<>4"
        rsTemp.Sort = "排列序号"
        With rsTemp
            Do While Not .EOF
                mstrTabHead = mstrTabHead & "|" & Nvl(!项目名称)
                mstrItemNo = mstrItemNo & "|" & Val(Nvl(!项目序号))
            .MoveNext
            Loop
        End With
        
        '确定心率是否和脉搏公用
        mrsItems.Filter = "项目序号=" & gint心率
        If mrsItems.RecordCount > 0 Then mint心率应用 = Val(Nvl(mrsItems!应用方式, 0))
        mrsItems.Filter = 0
        
        Set mrsData = CopyNewRs
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function CopyNewRs() As ADODB.Recordset
'功能:初始化项目列记录集
    Dim arrCOL() As String
    Dim i As Integer
    Dim strHead As String
    Dim rsNewRs As New ADODB.Recordset
    strHead = Mid(mstrTabHead, 2)
    arrCOL = Split(strHead, "|")
    
    '记录集格式
    '"行号|文件ID|床号|姓名|病人ID|主页ID|婴儿|日期|出院|时间" + 体温数据项目
    With rsNewRs
        .Fields.Append "行号", adDouble, 18
        For i = 0 To UBound(arrCOL)
            Select Case CStr(arrCOL(i))
                Case "文件ID,病人ID,主页ID,记录ID"
                    .Fields.Append CStr(arrCOL(i)), adDouble, 18, adFldIsNullable
                Case "婴儿,出院,护理等级"
                    .Fields.Append CStr(arrCOL(i)), adDouble, 1, adFldIsNullable
                Case "日期"
                    .Fields.Append CStr(arrCOL(i)), adLongVarChar, 50, adFldIsNullable
                Case Else
                    .Fields.Append CStr(arrCOL(i)), adLongVarChar, 20, adFldIsNullable
            End Select
        Next i
        .Fields.Append "删除", adDouble, 1 '-- 1表示保存后删除 2表示保存后修改了时间 ,0 未保存数据
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    Set CopyNewRs = rsNewRs
End Function

Private Sub ColligationTab(Optional blnFillPage As Boolean = True)
'-------------------------------------------------
'绑定表格列数据
'-------------------------------------------------
    Dim arrCOL() As String, arrNo() As String
    Dim lngCount As Long
    Dim lngRow As Long, lngCol As Long
    
    
    arrCOL = Split(mstrTabHead, "|")
    If mstrItemNo <> "" Then arrNo = Split(mstrItemNo, "|")
    With VsfData
        .Clear
        .Cols = IIf((UBound(arrCOL) + 1) = 0, RootCol, UBound(arrCOL) + 1)
        .FixedRows = 4
        .FixedCols = 1
        .Rows = 5
         
         '隐藏部分裂
        .ColHidden(c文件ID) = True
        .ColHidden(c病人ID) = True
        .ColHidden(c主页ID) = True
        .ColHidden(c婴儿) = True
        .ColHidden(c记录ID) = True
        .ColHidden(c日期) = True
        .ColHidden(c出院) = True
        .ColHidden(c护理等级) = True
        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True
        .ColWidth(0) = 250
        .ColWidth(c姓名) = 1500 + mintBigSize * 1500 / 3
        .ColAlignment(c姓名) = flexAlignLeftCenter
        .ColAlignment(c床号) = flexAlignRightCenter
        .ColAlignment(c年龄) = flexAlignLeftCenter
        .ColWidth(c体温标识) = 1000

        .FrozenCols = c时间
        .SheetBorder = &H40C0&
        
        .RowHeight(-1) = 300 + mintBigSize * 300 / 3
        .FontName = "宋体"
        .Font.Size = 9 + mintBigSize * 9 / 3
        '设置列头
        For lngCount = 0 To UBound(arrCOL)
            .TextMatrix(.FixedRows - 1, lngCount) = arrCOL(lngCount)
            If lngCount >= c时间 Then
                .ColWidth(lngCount) = 1200 + mintBigSize * 1200 / 3
                .ColAlignment(lngCount) = flexAlignCenterCenter
            End If
        Next lngCount
        
        '设置隐藏行
        For lngCol = 0 To .Cols - 1
            If lngCol < RootCol Then
                .TextMatrix(0, lngCol) = ""
                .TextMatrix(1, lngCol) = ""
                .TextMatrix(2, lngCol) = ""
            Else
                mrsItems.Filter = "项目序号=" & Val(arrNo(lngCol - RootCol))
                .TextMatrix(0, lngCol) = mrsItems!项目序号
                .TextMatrix(1, lngCol) = Nvl(mrsItems!项目类型, 0) & "|" & Val(Nvl(mrsItems!项目小数, 0)) & "|" & Nvl(mrsItems!项目值域)
                .TextMatrix(2, lngCol) = Val(Nvl(mrsItems!适用病人, 0))
            End If
        Next lngCol
        
         '固定行格式为居中
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        .Cell(flexcpAlignment, .FixedRows, c时间, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .Cell(flexcpText, .FixedRows, c时间, .Rows - 1, .Cols - 1) = ""
        .Cell(flexcpForeColor, 0, 0, .Rows - 1, .Cols - 1) = &H80000012
        
         If blnFillPage = True Then Call FillPage
    End With
End Sub

Private Sub FillPage()
'-----------------------------------------------------------------------------------------------------------------
'功能:提取符合条件的病人列表信息  入院三天内的病人 + 手术三天内的病人 + 三天内体温存在超过37.5度的病人 + 危/重病人
'-----------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim arrFilter() As String
    Dim strFilter As String, strPatient As String
    Dim strOutTime As String
    Dim i As Integer
    Dim strBegin As String, strEnd As String
    Dim strFind As String
    On Error GoTo Errhand
    
    strBegin = Format(Format(CDate(mstrDate) - 3, "YYYY-MM-DD") & " 00:00:00", "YYYY-MM-DD HH:mm:ss")
    strEnd = Format(Format(CDate(mstrDate), "YYYY-MM-DD") & " 23:59:59", "YYYY-MM-DD HH:mm:ss")
    
    'txtFilter.Tag 表示满足所有条件
    strFilter = txtFilter.Tag
    If Val(txtFilter.Tag) = 0 Then
       strFilter = "1;1;1;1;1;1;1"
    Else
        strFilter = ";;;;;;"
        arrFilter = Split(strFilter, ";")
        For i = 0 To UBound(Split(txtFilter.Tag, ";"))
            arrFilter(Val(Split(txtFilter.Tag, ";")(i)) - 1) = 1
        Next i
        strFilter = Join(arrFilter, ";")
    End If
    
    arrFilter = Split(strFilter, ";")
    
    strPatient = ""
    '58890:刘鹏飞,2013-02-26,在院病人读取性能优化(关联在院病人表进行查询)
    '此处对于出院病人不进行提取
    If Val(arrFilter(0)) = 1 Then '入院三天内的病人
        strPatient = "" & _
            " SELECT 1 AS 性质,B.病人ID, B.主页ID, A.姓名, A.性别,B.年龄, B.住院号, lpad(nvl(B.出院病床,' '),10,' ') AS 床号,0 AS 婴儿" & _
            " FROM 病人信息 A,病案主页 B,在院病人 C" & _
            " Where A.病人ID = B.病人ID And A.住院次数=B.主页ID And NVL(B.主页ID, 0) <> 0 " & _
            " AND Nvl(B.病案状态,0)<>5 AND B.封存时间 is NULL" & _
            " AND B.入院日期 BETWEEN [1] AND [2] And A.病人ID=C.病人ID And C.病区ID=[3]" & _
            IIf(mlng科室id = -1, "", " And C.科室ID=[4]")
    End If
    
    If Val(arrFilter(1)) = 1 Then '手术三天内的病人
        If strPatient <> "" Then strPatient = strPatient & " UNION "
        '提取体温单中自由标注手术的病人
        strPatient = strPatient & _
                " SELECT 1 AS 性质,B.病人ID,B.主页ID, A.姓名, A.性别,B.年龄,B.住院号, lpad(nvl(B.出院病床,' '),10,' ') AS 床号,0 AS 婴儿" & vbNewLine & _
                " FROM 病人信息 A,病案主页 B,在院病人 F, 病人护理文件 C ,病人护理数据 D,病人护理明细 E" & vbNewLine & _
                " WHERE A.病人ID = B.病人ID And A.住院次数=B.主页ID And A.病人ID=F.病人ID AND NVL(B.主页ID, 0) <> 0 AND F.病区ID = [3]" & vbNewLine & _
                " AND NVL(B.病案状态,0)<>5 AND B.封存时间 IS NULL" & vbNewLine & _
                " AND B.病人ID=C.病人ID AND B.主页ID=C.主页ID AND C.格式ID=[5] AND C.ID=D.文件ID AND D.ID=E.记录ID" & vbNewLine & _
                " AND E.记录类型=4 AND E.项目名称<>'分娩' AND E.终止版本 IS NULL" & vbNewLine & _
                " AND D.发生时间 BETWEEN [1] AND [2]" & vbNewLine & _
                IIf(mlng科室id = -1, "", " And F.科室ID=[4]")
                
        If strPatient <> "" Then strPatient = strPatient & " UNION "
        '从医嘱中提取病人手术信息
        strPatient = strPatient & _
                    "SELECT 1 AS 性质,B.病人ID,B.主页ID, A.姓名, A.性别,B.年龄,B.住院号, lpad(nvl(B.出院病床,' '),10,' ') AS 床号,0 AS 婴儿" & vbNewLine & _
                    "FROM  病人信息 A,病案主页 B,在院病人 F,(SELECT D.病人ID,D.主页ID FROM (SELECT DISTINCT A.病人ID,A.主页ID" & vbNewLine & _
                    "           FROM 病人医嘱记录 A, 诊疗项目目录 B" & vbNewLine & _
                    "           WHERE A.诊疗项目ID = B.ID AND A.诊疗类别 = 'F' And A.相关ID is null And A.医嘱状态 in (3,8) AND A.开始执行时间 BETWEEN [1] AND [2]" & vbNewLine & _
                    "           UNION" & vbNewLine & _
                    "           SELECT DISTINCT A.病人ID,A.主页ID FROM 病人新生儿记录 A WHERE A.出生时间 BETWEEN [1] AND [2]) D GROUP BY D.病人ID,D.主页ID) C" & vbNewLine & _
                    "WHERE A.病人ID = B.病人ID And A.住院次数=B.主页ID And A.病人ID=F.病人ID AND NVL(B.主页ID, 0) <> 0 AND F.病区ID = [3]" & vbNewLine & _
                    "AND NVL(B.病案状态,0)<>5 AND B.封存时间 IS NULL" & vbNewLine & _
                    "AND B.病人ID=C.病人ID AND B.主页ID=C.主页ID" & vbNewLine & _
                    IIf(mlng科室id = -1, "", " And F.科室ID=[4]")
    End If
    
    If Val(arrFilter(2)) = 1 Then '三天内体温存在超过37.5度的病人
        If strPatient <> "" Then strPatient = strPatient & " UNION "
        strPatient = strPatient & _
                    " SELECT 1 AS 性质,B.病人ID,B.主页ID, A.姓名, A.性别,B.年龄,B.住院号, lpad(nvl(B.出院病床,' '),10,' ') AS 床号,0 AS 婴儿" & vbNewLine & _
                    " FROM 病人信息 A,病案主页 B,在院病人 F, 病人护理文件 C ,病人护理数据 D,病人护理明细 E" & vbNewLine & _
                    " WHERE A.病人ID = B.病人ID And A.住院次数=B.主页ID And A.病人ID=F.病人ID AND NVL(B.主页ID, 0) <> 0 AND F.病区ID = [3]" & vbNewLine & _
                    " AND NVL(B.病案状态,0)<>5 AND B.封存时间 IS NULL" & vbNewLine & _
                    " AND B.病人ID=C.病人ID AND B.主页ID=C.主页ID AND C.格式ID=[5] AND C.ID=D.文件ID AND D.ID=E.记录ID" & vbNewLine & _
                    " AND E.记录类型=1 AND E.项目序号=1 AND LENGTH( TRANSLATE(E.记录内容,'-.0123456789' || E.记录内容,'-.0123456789')) =LENGTH(E.记录内容)" & vbNewLine & _
                    " AND zl_to_number(E.记录内容)>=37.5 AND E.终止版本 IS NULL" & vbNewLine & _
                    " AND D.发生时间 BETWEEN [1] AND [2]" & vbNewLine & _
                    IIf(mlng科室id = -1, "", " And F.科室ID=[4]")
    End If
    
    If Val(arrFilter(3)) = 1 Then '危/重病人
        If strPatient <> "" Then strPatient = strPatient & " UNION "
        strPatient = strPatient & _
               " SELECT 1 AS 性质,B.病人ID,B.主页ID, A.姓名, A.性别,B.年龄,B.住院号, lpad(nvl(B.出院病床,' '),10,' ') AS 床号,0 AS 婴儿" & _
               " FROM 病人信息 A,病案主页 B,在院病人 F " & _
               " Where A.病人ID = b.病人ID And A.住院次数=B.主页ID And A.病人ID=F.病人ID And NVL(b.主页ID, 0) <> 0 And F.病区ID = [3]" & _
               " AND Nvl(B.病案状态,0)<>5 AND B.封存时间 is NULL" & _
               " AND Instr(',' || '危,重' || ',',','|| B.当前病况 || ',')>0 " & _
               IIf(mlng科室id = -1, "", " And F.科室ID=[4]")
    End If
    
    If Val(arrFilter(4)) = 1 Then '转入三天内的病人
        If strPatient <> "" Then strPatient = strPatient & " UNION "
        strPatient = strPatient & _
            " SELECT 1 AS 性质,B.病人ID, B.主页ID, A.姓名, A.性别,B.年龄, B.住院号, lpad(nvl(B.出院病床,' '),10,' ') AS 床号,0 AS 婴儿" & vbNewLine & _
            " FROM 病人信息 A,病案主页 B,病人变动记录 C,在院病人 F" & vbNewLine & _
            " Where A.病人ID = b.病人ID And A.住院次数=B.主页ID And NVL(b.主页ID, 0) <> 0 And b.病人id = c.病人id And b.主页id = c.主页id " & vbNewLine & _
            " And A.病人ID=F.病人ID And F.病区ID= [3] AND Nvl(B.病案状态,0)<>5 AND B.封存时间 is NULL" & vbNewLine & _
            " AND Nvl(c.附加床位, 0) = 0 And C.终止时间 IS null  And C.开始原因 in (3,15) And C.开始时间  is not null And B.状态=0" & vbNewLine & _
            " AND C.开始时间 BETWEEN [1] AND [2]" & vbNewLine & _
             IIf(mlng科室id = -1, "", " And F.科室ID=[4]")
    End If
    
    '51286,刘鹏飞,2012-07-11,添加"一级及以上护理等级的病人"
    If Val(arrFilter(5)) = 1 Then
        If strPatient <> "" Then strPatient = strPatient & " UNION "
        strPatient = strPatient & _
               " SELECT 1 AS 性质,B.病人ID,B.主页ID, A.姓名, A.性别,B.年龄,B.住院号, lpad(nvl(B.出院病床,' '),10,' ') AS 床号,0 AS 婴儿" & _
               " FROM 病人信息 A,病案主页 B,在院病人 F " & _
               " Where A.病人ID = b.病人ID And A.住院次数=B.主页ID And A.病人ID=F.病人ID And NVL(b.主页ID, 0) <> 0 And F.病区ID = [3]" & _
               " AND Nvl(B.病案状态,0)<>5 AND B.封存时间 is NULL" & _
               " AND zl_PatitTendGrade(B.病人ID,B.主页ID)<=1 " & _
               IIf(mlng科室id = -1, "", " And F.科室ID=[4]")
    End If
    
    If mstr科室性质 = "产科" Or mstr科室性质 = "所有" Then
        If Val(arrFilter(6)) = 1 Then '分娩后三天内的病人
            If strPatient <> "" Then strPatient = strPatient & " UNION "
            strPatient = strPatient & _
                    " SELECT 1 AS 性质,B.病人ID,B.主页ID, A.姓名, A.性别,B.年龄,B.住院号, lpad(nvl(B.出院病床,' '),10,' ') AS 床号,0 AS 婴儿" & vbNewLine & _
                    " FROM 病人信息 A,病案主页 B,在院病人 F, 病人护理文件 C ,病人护理数据 D,病人护理明细 E" & vbNewLine & _
                    " WHERE A.病人ID = B.病人ID And A.住院次数=B.主页ID And A.病人ID=F.病人ID AND NVL(B.主页ID, 0) <> 0 AND F.病区ID=[3]" & vbNewLine & _
                    " AND NVL(B.病案状态,0)<>5 AND B.封存时间 IS NULL" & vbNewLine & _
                    " AND B.病人ID=C.病人ID AND B.主页ID=C.主页ID AND C.格式ID=[5] AND C.ID=D.文件ID AND D.ID=E.记录ID" & vbNewLine & _
                    " AND E.记录类型=4 AND E.项目名称='分娩' AND E.终止版本 IS NULL" & vbNewLine & _
                    " AND D.发生时间 BETWEEN [1] AND [2]" & vbNewLine & _
                    IIf(mlng科室id = -1, "", " And F.科室ID=[4]")
        End If
    End If
    
    If strPatient = "" Then Exit Sub
    
    Select Case mintPatiNo
        Case 1
            '只提取病人本人
            strPatient = strPatient
        Case 2
            '只提取婴儿信息
            strPatient = " Select B.性质,B.病人ID,B.主页ID,NVL(A.婴儿姓名,B.姓名||'之子'||A.序号) AS 姓名,B.性别,Zl_Age_Calc(0,A.出生时间,sysdate) 年龄,B.住院号,lpad(B.床号,10,' ') as 床号,A.序号 AS 婴儿" & _
              " From 病人新生儿记录 A,(" & strPatient & ") B" & _
              " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID"
        Case Else
             '提取病人及新生儿列表
            strPatient = strPatient & _
                  " UNION " & _
                  " Select B.性质,B.病人ID,B.主页ID,NVL(A.婴儿姓名,B.姓名||'之子'||A.序号) AS 姓名,B.性别,Decode(nvl(A.序号,0),0,B.年龄,Zl_Age_Calc(0,A.出生时间,sysdate)) 年龄,B.住院号,lpad(B.床号,10,' ') as 床号,A.序号 AS 婴儿" & _
                  " From 病人新生儿记录 A,(" & strPatient & ") B" & _
                  " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID"
    End Select
   
    mstrSQL = " SELECT  A.性质,A.病人ID,A.主页ID,A.婴儿,A.姓名,A.年龄,lpad(A.床号,10,' ') as 床号,nvl(zl_PatitTendGrade(A.病人ID,A.主页ID),3) 护理等级,C.信息值 AS 体温标识, MAX(B.ID) AS 文件ID,B.开始时间" & _
              " FROM (" & strPatient & ") A,病人护理文件 B,病案主页从表 C" & _
              " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And A.婴儿=B.婴儿 And A.病人ID=C.病人ID(+) And A.主页ID=C.主页ID(+) And C.信息名(+)='体温标识'||DECODE(A.婴儿,0,'',A.婴儿) " & _
              " And B.归档人 is null And B.结束时间 is null And B.格式ID=[5]" & _
              " GROUP BY A.性质,A.病人ID,A.主页ID,A.婴儿,C.信息值,A.姓名 ,A.年龄,A.床号,B.开始时间" & _
              " Order by A.性质,A.床号,A.婴儿"
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "提取病人清单", CDate(strBegin), CDate(strEnd), mlng病区ID, mlng科室id, mlng格式ID)
     
    strOutTime = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")
     
    '填充数据到表格
    With rsTemp
        Do While Not .EOF
            mblnNullRow = True
            mblnRefreshData = True
            If .AbsolutePosition > VsfData.Rows - VsfData.FixedRows Then VsfData.Rows = VsfData.Rows + 1
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, c文件ID) = !文件ID
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, c床号) = Nvl(!床号)
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, c姓名) = IIf(!婴儿 > 0, Space(4), "") & !姓名
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, c年龄) = Nvl(!年龄)
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, c病人ID) = !病人ID
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, c主页ID) = !主页ID
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, c婴儿) = Nvl(!婴儿, 0)
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, c护理等级) = Val(!护理等级)
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, c体温标识) = Nvl(!体温标识)
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, c日期) = Format(!开始时间, "YYYY-MM-DD HH:mm:ss") & ";" & strOutTime
            VsfData.TextMatrix(.AbsolutePosition + VsfData.FixedRows - 1, c出院) = 0
            .MoveNext
        Loop
    End With
    
    If VsfData.Rows <= VsfData.FixedRows Then VsfData.Rows = VsfData.Rows + 1
    VsfData.RowHeight(-1) = 300 + mintBigSize * 300 / 3
    VsfData.Select VsfData.FixedRows, c时间
    '设置编辑颜色
    Call SetTabEditColor
    
    VsfData.Cell(flexcpForeColor, VsfData.FixedRows, c体温标识, VsfData.Rows - 1, c体温标识) = RGB(0, 0, 255)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub ColligationHistoryTab()
'-------------------------------------------------
'绑定表格列数据
'-------------------------------------------------
    Dim arrCOL() As String, arrNo() As String
    Dim lngCount As Long
    Dim lngRow As Long, lngCol As Long
    
    
    arrCOL = Split(mstrTabHead, "|")
    If mstrItemNo <> "" Then arrNo = Split(mstrItemNo, "|")
    With vsfHistory
        .Clear
        .Cols = IIf((UBound(arrCOL) + 1) = 0, RootCol, UBound(arrCOL) + 1)
        .FixedRows = 4
        .FixedCols = 1
        .Rows = 4
         
         '隐藏部分裂
        .ColHidden(c文件ID) = True
        .ColHidden(c床号) = True
        .ColHidden(c姓名) = True
        .ColHidden(c年龄) = True
        .ColHidden(c病人ID) = True
        .ColHidden(c主页ID) = True
        .ColHidden(c婴儿) = True
        .ColHidden(c记录ID) = True
        .ColHidden(c护理等级) = True
        .ColHidden(c体温标识) = True
        .ColHidden(c出院) = True
        
        .RowHidden(0) = True
        .RowHidden(1) = True
        .RowHidden(2) = True
        .ColWidth(0) = 250
        
        .FrozenCols = c时间
        .SheetBorder = &H40C0&
        
        .RowHeight(-1) = 300 + mintBigSize * 300 / 3
        .FontName = "宋体"
        .Font.Size = 9 + mintBigSize * 9 / 3
        '设置列头
        For lngCount = 0 To UBound(arrCOL)
            .TextMatrix(.FixedRows - 1, lngCount) = arrCOL(lngCount)
            If lngCount >= c日期 Then
                .ColWidth(lngCount) = 1200 + mintBigSize * 1200 / 3
                .ColAlignment(lngCount) = flexAlignCenterCenter
            End If
        Next lngCount
        
        '设置隐藏行
        For lngCol = 0 To .Cols - 1
            If lngCol < RootCol Then
                .TextMatrix(0, lngCol) = ""
                .TextMatrix(1, lngCol) = ""
                .TextMatrix(2, lngCol) = ""
            Else
                mrsItems.Filter = "项目序号=" & Val(arrNo(lngCol - RootCol))
                .TextMatrix(0, lngCol) = mrsItems!项目序号
                .TextMatrix(1, lngCol) = Nvl(mrsItems!项目类型, 0) & "|" & Val(Nvl(mrsItems!项目小数, 0)) & "|" & Nvl(mrsItems!项目值域)
                .TextMatrix(2, lngCol) = Val(Nvl(mrsItems!适用病人, 0))
            End If
        Next lngCol
        
         '固定行格式为居中
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        If .Rows > .FixedRows Then
            .Cell(flexcpAlignment, .FixedRows, c日期, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
            .Cell(flexcpText, .FixedRows, c日期, .Rows - 1, .Cols - 1) = ""
        End If
        .Cell(flexcpForeColor, 0, 0, .Rows - 1, .Cols - 1) = &H80000012
    End With
End Sub

Private Sub cmdFilterCancel_Click()
    picFilter.Visible = False
End Sub

Private Sub cmdFilterOK_Click()
    Dim i As Integer
    Dim strValue As String
    Dim arrValue() As String, ArrCode() As String
    
    If lstFilter.SelCount = 0 Then
        MsgBox "请至少选择一种过滤条件。", vbInformation, gstrSysName
        lstFilter.SetFocus
        Exit Sub
    End If
    
    If lstFilter.Selected(0) = True Then
        txtFilter.Text = "全部"
        txtFilter.Tag = 0
    Else
        txtFilter.Text = ""
        txtFilter.Tag = ""
        For i = 1 To lstFilter.ListCount - 1
            If lstFilter.Selected(i) Then
                txtFilter.Text = txtFilter.Text & ";" & lstFilter.List(i)
                txtFilter.Tag = txtFilter.Tag & ";" & lstFilter.ItemData(i)
            End If
        Next
        txtFilter.Text = Mid(txtFilter.Text, 2)
        txtFilter.Tag = Mid(txtFilter.Tag, 2)
    End If
    
    txtFilter.SetFocus
    picFilter.Visible = False
    
    '保存过滤条件信息
    If Val(txtFilter.Tag) = 0 Then
        strValue = "1;1;1;1;1;1;1"
    Else
        strValue = "0;0;0;0;0;0;0"
        arrValue = Split(strValue, ";")
        ArrCode = Split(txtFilter.Tag, ";")
        For i = 0 To UBound(ArrCode)
            arrValue(Val(ArrCode(i)) - 1) = 1
        Next i
        strValue = Join(arrValue, ";")
    End If
    
    Call zlDatabase.SetPara("体温单过滤条件", strValue, glngSys, 1255)
    
    '开始重新加载数据信息
    'Call cmdFilter_Click
End Sub

Private Sub cmdFilterUserCancle_Click()
    picPati.Visible = False
    VsfData.SetFocus
End Sub

Private Sub cmdFilterUserOk_Click()
    '添加病人
    Dim rsTemp As New ADODB.Recordset
    Dim objRow As ReportRow
    Dim lngLoop As Long
    Dim strPatient As String, strSQL As String
    Dim lngRow As Long, lngTempRow As Long
    Dim strCurDate As String, strInTime As String, strOutTime As String
    Dim blnNullRow As Long, blnOut As Boolean
    
    '病人信息变量
    Dim lng病人ID As Long, lng主页ID As Long, str姓名 As String, str性别 As String, str年龄 As String, str住院号 As String, str床号 As String, intBaby As Integer
    
    strPatient = ""
    For Each objRow In rptPati.Rows
        If objRow.GroupRow And objRow.Childs.Count > 0 Then
            For lngLoop = 0 To objRow.Childs.Count - 1
                If Not (objRow.Childs(lngLoop).Record Is Nothing) Then
                    If objRow.Childs(lngLoop).Record.Item(c_选择).Checked = True Then
                        lng病人ID = Val(objRow.Childs(lngLoop).Record.Item(C_病人ID).Value)
                        lng主页ID = Val(objRow.Childs(lngLoop).Record.Item(c_主页ID).Value)
                        str姓名 = objRow.Childs(lngLoop).Record.Item(c_姓名).Value
                        str性别 = IIf(Val(objRow.Childs(lngLoop).Record.Item(c_图标).Icon) = 1, "男", "女")
                        str年龄 = objRow.Childs(lngLoop).Record.Item(c_年龄).Value
                        str住院号 = Val(objRow.Childs(lngLoop).Record.Item(c_住院号).Value)
                        str床号 = objRow.Childs(lngLoop).Record.Item(c_床号).Value
                        strOutTime = objRow.Childs(lngLoop).Record.Item(c_出院日期).Value
                        intBaby = 0
                        
                        strSQL = ""
                        strSQL = "SELECT 1 性质,"
                        strSQL = strSQL & lng病人ID & " 病人ID,"
                        strSQL = strSQL & lng主页ID & " 主页ID,"
                        strSQL = strSQL & "'" & str姓名 & "' 姓名,"
                        strSQL = strSQL & "'" & str性别 & "' 性别,"
                        strSQL = strSQL & "'" & str年龄 & "' 年龄,"
                        strSQL = strSQL & "" & str住院号 & " 住院号,"
                        strSQL = strSQL & "'" & str床号 & "' 床号,"
                        strSQL = strSQL & "" & intBaby & " 婴儿,"
                        strSQL = strSQL & "'" & strOutTime & "' 出院日期"
                        strSQL = strSQL & " FROM dual"
                        
                        strPatient = strPatient & vbCrLf & IIf(strPatient = "", strSQL, " UNION " & vbCrLf & strSQL)
                    End If
                End If
            Next lngLoop
        End If
    Next
    
    '隐藏所有PIC
    Call InitCons
    
    If Trim(strPatient) = "" Then Exit Sub
    On Error GoTo Errhand:
    
    '如果还未添加列此处需要添加列信息
    If Not mblnNullRow Then
        mstrDate = Format(dtpDate.Value, "YYYY-MM-DD")
        Call InitVariable
        Call zlRefreshDate(False)
        mblnInit = True
    End If
    
    blnNullRow = mblnNullRow
    
    strPatient = "SELECT 性质,病人ID,主页ID,姓名,性别,年龄,住院号,lpad(床号,10,' ') as  床号,婴儿,出院日期 FROM (" & strPatient & ")"
    
    Select Case mintPatiNo
        Case 1
            strPatient = strPatient
        Case 2
            strPatient = " Select B.性质,B.病人ID,B.主页ID,NVL(A.婴儿姓名,B.姓名||'之子'||A.序号) AS 姓名,B.性别,Zl_Age_Calc(0,A.出生时间,sysdate) 年龄,B.住院号,lpad(B.床号,10,' ') as 床号,A.序号 AS 婴儿,B.出院日期" & _
                  " From 病人新生儿记录 A,(" & strPatient & ") B" & _
                  " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID"
        Case Else
            '提取病人和新生儿列表
            strPatient = strPatient & _
                  " UNION " & _
                  " Select B.性质,B.病人ID,B.主页ID,NVL(A.婴儿姓名,B.姓名||'之子'||A.序号) AS 姓名,B.性别,Decode(nvl(A.序号,0),0,B.年龄,Zl_Age_Calc(0,A.出生时间,sysdate)) 年龄,B.住院号,lpad(B.床号,10,' ') as 床号,A.序号 AS 婴儿,B.出院日期" & _
                  " From 病人新生儿记录 A,(" & strPatient & ") B" & _
                  " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID"
    End Select

     mstrSQL = " SELECT  A.性质, A.病人ID,A.主页ID,A.婴儿,nvl(zl_PatitTendGrade(A.病人ID,A.主页ID),3) 护理等级,C.信息值 AS 体温标识,A.姓名,A.年龄,lpad(A.床号,10,' ') as 床号,A.出院日期,MAX(B.ID) AS 文件ID,B.开始时间" & _
              " FROM (" & strPatient & ") A,病人护理文件 B,病案主页从表 C" & _
              " Where A.病人ID=B.病人ID And A.主页ID=B.主页ID And A.婴儿=B.婴儿 And A.病人ID=C.病人ID(+) And A.主页ID=C.主页ID(+) And C.信息名(+)='体温标识'||DECODE(A.婴儿,0,'',A.婴儿) " & _
              " And B.归档人 is null And B.结束时间 is null And B.格式ID=[1]" & _
              " GROUP BY A.性质,A.病人ID,A.主页ID,A.婴儿,C.信息值,A.姓名,A.年龄,A.床号,A.出院日期,B.开始时间" & _
              " Order by A.性质,A.床号,A.婴儿"
     Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "提取病人清单", mlng格式ID)
     
     strCurDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")
    '填充数据到表格
    lngRow = 0
    With rsTemp
        Do While Not .EOF
            blnOut = True
            mblnNullRow = True
            
            If blnNullRow = False Then
                If .AbsolutePosition > VsfData.Rows - VsfData.FixedRows Then VsfData.Rows = VsfData.Rows + 1
                lngTempRow = .AbsolutePosition + VsfData.FixedRows - 1
            Else
                If VsfData.Rows > VsfData.FixedRows Then
                    If VsfData.TextMatrix(VsfData.Rows - 1, c文件ID) <> 0 Then
                        VsfData.Rows = VsfData.Rows + 1
                    End If
                Else
                    VsfData.Rows = VsfData.Rows + 1
                End If
                
                lngTempRow = VsfData.Rows - 1
            End If
            strOutTime = Trim(Nvl(!出院日期))
            If strOutTime = "" Then strOutTime = strCurDate: blnOut = False
            
            VsfData.TextMatrix(lngTempRow, c文件ID) = !文件ID
            VsfData.TextMatrix(lngTempRow, c床号) = Nvl(!床号)
            VsfData.TextMatrix(lngTempRow, c姓名) = IIf(!婴儿 > 0, Space(4), "") & !姓名
            VsfData.TextMatrix(lngTempRow, c年龄) = Nvl(!年龄)
            VsfData.TextMatrix(lngTempRow, c病人ID) = !病人ID
            VsfData.TextMatrix(lngTempRow, c主页ID) = !主页ID
            VsfData.TextMatrix(lngTempRow, c婴儿) = Nvl(!婴儿, 0)
            VsfData.TextMatrix(lngTempRow, c护理等级) = Val(!护理等级)
            VsfData.TextMatrix(lngTempRow, c体温标识) = Nvl(!体温标识)
            VsfData.TextMatrix(lngTempRow, c日期) = Format(!开始时间, "YYYY-MM-DD HH:mm:ss") & ";" & strOutTime
            VsfData.TextMatrix(lngTempRow, c出院) = IIf(blnOut = True, 1, 0)
            
            If lngRow = 0 Then lngRow = lngTempRow
            .MoveNext
        Loop
    End With
    
    If VsfData.Rows <= VsfData.FixedRows Then VsfData.Rows = VsfData.Rows + 1
    VsfData.RowHeight(-1) = 300 + mintBigSize * 300 / 3
    
    If lngRow = 0 Then lngRow = VsfData.Rows - 1
    VsfData.Select lngRow, c时间

    '设置编辑颜色
    Call SetTabEditColor
    '保存数据集
    If Not mblnSaveData Then
        Call Data_Save
    End If

    VsfData.Cell(flexcpForeColor, VsfData.FixedRows, c体温标识, VsfData.Rows - 1, c体温标识) = RGB(0, 0, 255)
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetTabEditColor()
'-----------------------------------------------
'功能:判断该病人的护理等级是否能使用某个项目
'-----------------------------------------------
    Dim intRow As Integer, intCOl As Integer
    Dim int护理等级 As Integer, int婴儿 As Integer
    Dim lngItemNO As Long
    Dim blnTrue As Boolean
    
    For intRow = VsfData.FixedRows To VsfData.Rows - 1
        If VsfData.RowHidden(intRow) = False And Val(VsfData.TextMatrix(intRow, c文件ID)) <> 0 Then
            int护理等级 = Val(VsfData.TextMatrix(intRow, c护理等级))
            int婴儿 = Val(VsfData.TextMatrix(intRow, c婴儿))
            For intCOl = RootCol To VsfData.Cols - 1
                blnTrue = False
                lngItemNO = Val(VsfData.TextMatrix(0, intCOl))
                mrsItems.Filter = 0
                mrsItems.Filter = "项目序号=" & lngItemNO & " And 护理等级>=" & int护理等级
                If mrsItems.RecordCount > 0 Then
                    VsfData.Cell(flexcpBackColor, intRow, intCOl, intRow, intCOl) = &H80000005
                Else
                    VsfData.Cell(flexcpBackColor, intRow, intCOl, intRow, intCOl) = &H80000016
                    blnTrue = True
                End If
                '检查是否适用于此病人
                If Not blnTrue Then
                    If Val(VsfData.TextMatrix(2, intCOl)) = 1 Then
                        VsfData.Cell(flexcpBackColor, intRow, intCOl, intRow, intCOl) = IIf(int婴儿 = 0, &H80000005, &H80000016)
                    ElseIf VsfData.TextMatrix(2, intCOl) = 2 Then
                        VsfData.Cell(flexcpBackColor, intRow, intCOl, intRow, intCOl) = IIf(int婴儿 <> 0, &H80000005, &H80000016)
                    End If
                End If
            Next intCOl
        End If
    Next intRow
    VsfData.Cell(flexcpForeColor, VsfData.FixedRows, c体温标识, VsfData.Rows - 1, c体温标识) = RGB(0, 0, 255)
End Sub

Private Sub CmdRef_Click()
    Set mrsPati = New ADODB.Recordset
    If mrsPati.State = adStateOpen Then mrsPart.Close
    Call cmdAddUser_Click
End Sub

Private Sub cmdSift_Click()
    Dim i As Integer
    
    For i = 0 To lstFilter.ListCount - 1
        If Val(txtFilter.Tag) = 0 Then
            lstFilter.Selected(i) = True
        ElseIf InStr(1, ";" & txtFilter.Tag & ";", ";" & lstFilter.ItemData(i) & ";") <> 0 Then
            lstFilter.Selected(i) = True
        Else
            lstFilter.Selected(i) = False
        End If
    Next i
    lstFilter.ListIndex = 0
    With picFilter
        .Top = picMain.Top
        .Left = txtFilter.Left + 60
        .Visible = True
        .ZOrder 0
    End With
End Sub

Private Sub dtpB_Change(Index As Integer)
'时间范围改变时刷新
    If dtpB(Index).Value >= dtpE(Index).Value Then
        RaiseEvent AfterRowColChange("时间范围的开始时间应小于结束时间", True)
        dtpB(Index).Value = dtpB(Index).Tag
        dtpB(Index).SetFocus: Exit Sub
    Else
        dtpB(Index).Tag = dtpB(Index).Value
        If Index = 0 Then mdtOutBegin = dtpB(Index).Value
    End If
End Sub

Private Sub dtpE_Change(Index As Integer)
    If dtpB(Index).Value >= dtpE(Index).Value Then
        RaiseEvent AfterRowColChange("时间范围的开始时间应小于结束时间", True)
        dtpE(Index).Value = dtpE(Index).Tag
        dtpE(Index).SetFocus: Exit Sub
    Else
        dtpE(Index).Tag = dtpE(Index).Value
        If Index = 0 Then mdtOutEnd = dtpE(Index).Value
    End If
End Sub

Private Sub dtpDate_Change()
    Dim blnCancle As Boolean
    Call dtpDate_Validate(blnCancle)
    If blnCancle = True Then
        dtpDate.SetFocus
    End If
End Sub

Private Sub dtpDate_GotFocus()
    If Not mblnDateFouces Then Call InitCons
End Sub

Private Sub dtpDate_Validate(Cancel As Boolean)
    If Not mblnInit Then Exit Sub
    If CheckEditData Then
        Cancel = True
        dtpDate.Value = Format(mstrDate, "YYYY-MM-DD")
        Exit Sub
    End If
    mstrDate = Format(dtpDate.Value, "YYYY-MM-DD")
End Sub

Private Sub lblCheck_DblClick()
    Call picInput_KeyPress(vbKeySpace)
End Sub

Private Sub lstFilter_ItemCheck(Item As Integer)
    Dim i As Integer
    
    If Item = 0 Then
        For i = 1 To lstFilter.ListCount - 1
            lstFilter.Selected(i) = lstFilter.Selected(0)
        Next
    ElseIf Not lstFilter.Selected(Item) Then
        lstFilter.Selected(0) = False
    ElseIf lstFilter.SelCount = lstFilter.ListCount - 1 Then
        lstFilter.Selected(0) = True
    End If
End Sub

Private Sub lstFilter_LostFocus()
    If Not UserControl.ActiveControl Is cmdFilterOK _
        And Not UserControl.ActiveControl Is cmdFilterCancel _
        And Not UserControl.ActiveControl Is lstFilter _
        And Not UserControl.ActiveControl Is picFilter Then picFilter.Visible = False: mblnDateFouces = False
End Sub

Private Sub lstNote_DblClick()
    Call lstNote_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub lstNote_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngRow As Long, lngCol As Long, lngItemNO As Long
    Dim strNote As String
    Dim intCount As Integer, intCOl As Integer, intCols As Integer
    Dim intStartCol As Integer, intEndCol As Integer
    Dim blnAll As Boolean
    
    If KeyCode = vbKeyReturn Then
        If Shift <> 0 Then Exit Sub
        If picInput.Visible = False Then Exit Sub
        
        lngRow = VsfData.Row 'Val(Split(txtInput.Tag, "|")(0))
        lngCol = VsfData.Col 'Val(Split(txtInput.Tag, "|")(1))
        strNote = lstNote.Text
        
        VsfData.TextMatrix(lngRow, lngCol) = strNote
        txtInput.Text = strNote
        mrsItems.Filter = 0
        
        '检查其他的曲线是否无数值
        intStartCol = RootCol
        intCount = 0
        intCols = 0
        For intCOl = intStartCol To VsfData.Cols - 1
            lngItemNO = Val(VsfData.TextMatrix(0, intCOl))
            mrsItems.Filter = "项目序号=" & lngItemNO
            If Trim(Nvl(mrsItems!分组名)) = "1)体温曲线项目" Then
                If Trim(VsfData.TextMatrix(lngRow, intCOl)) = "" Then
                    intCount = intCount + 1
                End If
                intCols = intCols + 1
                intEndCol = intCOl
            End If
        Next intCOl
        
        '循环赋值
        If intCount = intCols - 1 Then
            For intCOl = intStartCol To intEndCol
                VsfData.TextMatrix(lngRow, intCOl) = strNote
            Next intCOl
            blnAll = True
        Else
            intCount = 0
            intCols = 1
            blnAll = False
        End If
        
        If blnAll = True Then
            '定位到第一行体温项目
            VsfData.Col = intStartCol
        Else
            VsfData.Col = lngCol
        End If
        
        For intCOl = 1 To intCols
            picInput.Tag = ""
            mblnDateFouces = True
            Call MoveNextCell(vbKeyReturn)
        Next intCOl
        
    ElseIf KeyCode = vbKeyEscape And Shift = 0 Then
        If picInput.Visible = True Then picInput.SetFocus
    End If
End Sub

Private Sub lstNote_LostFocus()
    Call lstNote_KeyDown(vbKeyEscape, 0)
End Sub



Private Sub picDouble_GotFocus()
    If picDouble.Visible = True Then txtUpInput.SetFocus
End Sub

Private Sub picFilter_GotFocus()
    lstFilter.SetFocus
End Sub

Private Sub picFilter_Resize()
    On Error Resume Next
    lstFilter.Left = -15
    lstFilter.Top = -15
    lstFilter.Width = picFilter.Width
    
    cmdFilterCancel.Left = picFilter.ScaleWidth - cmdFilterCancel.Width - 100
    cmdFilterOK.Left = cmdFilterCancel.Left - cmdFilterOK.Width - 60
    
    cmdFilterOK.Top = lstFilter.Height + (picFilter.ScaleHeight - lstFilter.Height - cmdFilterOK.Height) / 2
    cmdFilterCancel.Top = cmdFilterOK.Top
End Sub

Private Sub picHistory_Resize()
    txt显示天数.Left = lblDay.Left + lblDay.Width + TextWidth("刘")
    lblDay.Top = txt显示天数.Top + (txt显示天数.Height - lblDay.Height) \ 2
    lblDayInfo.Top = lblDate.Top
    lblDayInfo.Left = txt显示天数.Left + txt显示天数.Width + TextWidth("刘")
End Sub

Private Sub picInput_DblClick()
    Call picInput_KeyPress(vbKeySpace)
End Sub

Private Sub picInput_GotFocus()
    If picInput.Visible = True And txtInput.Visible = True Then txtInput.SetFocus
End Sub

Private Sub picInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        picInput.Visible = False
        lstNote.Visible = False
        picInput.Tag = ""
        txtInput.Tag = ""
        txtInput.Text = ""
        lstNote.Tag = ""
        mintType = 0
        mblnShow = False
        VsfData.SetFocus
    ElseIf KeyAscii = vbKeySpace Then
        If lblCheck.Caption = "√" Then
            lblCheck.Caption = ""
        Else
            lblCheck.Caption = "√"
        End If
    ElseIf KeyAscii = vbKeyReturn Then
        If txtInput.Visible = False Then
            mblnDateFouces = True
            Call VsfData_KeyDown(vbKeyReturn, 0)
        End If
    ElseIf KeyAscii = vbKeyLeft Then
        If txtInput.Visible = False Then
            mblnDateFouces = True
            Call MoveNextCell(KeyAscii)
        End If
    End If
End Sub

Private Sub PicLst_GotFocus()
    If PicLst.Visible = False Then Exit Sub
    If Trim(txtLst.Text) = "" Then
        PicLst.Tag = 0
        lstSelect(0).SetFocus
    Else
        PicLst.Tag = 1
        txtLst.SetFocus
    End If
End Sub

Private Sub picSplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    picSplit.Tag = 1
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    If Val(picSplit.Tag) = 0 Then Exit Sub
    
    If picSplit.Top + Y < 4000 Then
        picSplit.Top = 4000
    ElseIf ScaleHeight - (picSplit.Top + Y) < 3000 Then
        picSplit.Top = ScaleHeight - 3000
    Else
        picSplit.Move picSplit.Left, picSplit.Top + Y
    End If
End Sub

Private Sub picSplit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Val(picSplit.Tag) = 1 Then Call cbsThis_Resize

    picSplit.Tag = 0
End Sub

Private Sub rptPati_ItemCheck(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    Dim objRow As ReportRow
    Dim lngLoop As Long
    Dim blnAll As Boolean
    
    If Item.Index = c_选择 Then
        For Each objRow In rptPati.Rows
            If objRow.GroupRow And objRow.Childs.Count > 0 Then
                For lngLoop = 0 To objRow.Childs.Count - 1
                    If Not (objRow.Childs(lngLoop).Record Is Nothing) Then
                        If Trim(objRow.Childs(lngLoop).Record.Item(c_出院日期).Value) <> "" Then Exit For
                        blnAll = True
                        If objRow.Childs(lngLoop).Record.Item(c_选择).Checked = False Then
                            blnAll = False
                            GoTo NextCheck
                        End If
                    End If
                Next lngLoop
            End If
        Next
    End If
NextCheck:
    mblnChkClick = True
    chkSwitch.Value = IIf(blnAll = True, 1, 0)
End Sub

Private Sub rptPati_LostFocus()
    If Not UserControl.ActiveControl Is cmdFilterUserOk _
        And Not UserControl.ActiveControl Is cmdFilterUserCancle _
        And Not UserControl.ActiveControl Is rptPati _
        And Not UserControl.ActiveControl Is picPati Then picPati.Visible = False: mblnDateFouces = False
End Sub

Private Sub rptPati_RowDblClick(ByVal Row As XtremeReportControl.IReportRow, ByVal Item As XtremeReportControl.IReportRecordItem)
    '添加病人信息
    If Not Row.Record Is Nothing Then
        Row.Record.Item(c_选择).Checked = True
        Call cmdFilterUserOk_Click
    End If
End Sub

Private Sub txtChange_KeyPress(KeyAscii As Integer)
    If InStr("1234567890", Chr(KeyAscii)) = 0 And KeyAscii <> vbKeyBack And KeyAscii <> vbKeyReturn Then KeyAscii = 0
    If KeyAscii <> vbKeyReturn Then Exit Sub
    mintChange = Val(txtChange.Text)
End Sub

Private Sub txtChange_GotFocus()
    Call zlControl.TxtSelAll(txtChange)
End Sub

Private Sub txtChange_Validate(Cancel As Boolean)
    If Val(txtChange.Text) > 30 Then
        RaiseEvent AfterRowColChange("转出病人天数不能超过30天", True)
        Cancel = True
    Else
        mintChange = Val(txtChange.Text)
    End If
End Sub

Private Sub txtDnInput_GotFocus()
    Call zlControl.TxtSelAll(txtDnInput)
End Sub

Private Sub txtDnInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Shift = vbCtrlMask Then
            Exit Sub
        Else
            Call VsfData_KeyDown(KeyCode, Shift)
        End If
    End If
    
    If KeyCode = vbKeyLeft Then
        If txtDnInput.SelStart = 0 Then
            txtUpInput.SetFocus
        End If
    End If
End Sub

Private Sub txtDnInput_KeyPress(KeyAscii As Integer)
    Call txtUpInput_KeyPress(KeyAscii)
End Sub

Private Sub txtDnInput_LostFocus()
    mblnDateFouces = False
End Sub





Private Sub txtInput_GotFocus()
    Call zlControl.TxtSelAll(txtInput)
    lstNote.Visible = False
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Shift = vbCtrlMask Then
            Exit Sub
        Else
            Call VsfData_KeyDown(KeyCode, Shift)
        End If
    End If
    
    If KeyCode = vbKeyLeft Then
        Call MoveNextCell(vbKeyLeft)
    End If
    
    If KeyCode = vbKeyDown Then '显示未记说明信息
        If picInput.Visible = False Or txtInput.Visible = False Then Exit Sub
        If VsfData.Col < RootCol Or VsfData.Col > VsfData.Cols - 2 Then Exit Sub
        If InStr(1, ",0,9,", "," & mint数据来源 & ",") = 0 Then Exit Sub
        
        With lstNote
            .Top = picInput.Top + picInput.Height
            .Left = picInput.Left
            .FontName = VsfData.FontName
            .Font.Size = VsfData.Font.Size
            .Width = LenB(StrConv(.List(.ListCount \ 2), vbFromUnicode)) * 160 + 500
            If .Width < picInput.Width Then .Width = picInput.Width
            .Height = .ListCount * 210 + 30
            If .Height + .Top + picMain.Top > ScaleHeight Then
                .Top = ScaleHeight - picMain.Top - .Height
            End If
            .Visible = True
            lstNote.SetFocus
        End With
    End If
    
    '隐藏编辑框
    If KeyCode = vbKeyEscape And Shift = 0 Then
        Call picInput_KeyPress(vbKeyEscape)
    End If
End Sub

Private Sub lstSelect_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Shift = vbShiftMask Then Exit Sub
        mblnDateFouces = True
        Call VsfData_KeyDown(KeyCode, Shift)
    ElseIf KeyCode = vbKeyLeft Then
        mblnDateFouces = True
        Call MoveNextCell(vbKeyLeft)
    ElseIf Index = 0 And Shift = vbShiftMask And KeyCode = vbKeyUp Then
        KeyCode = 0
        txtLst.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        lstSelect(Index).Visible = False
        If Index = 0 Then
            PicLst.Visible = False
            txtLst.Visible = False
        End If
        mblnShow = False
        VsfData.SetFocus
    End If
End Sub

Private Sub lstSelect_DblClick(Index As Integer)
    Call lstSelect_KeyDown(Index, vbKeyReturn, 0)
End Sub

Private Sub lstSelect_GotFocus(Index As Integer)
    Dim i As Integer, j As Integer
    PicLst.Tag = 0
    j = lstSelect(Index).ListCount - 1
    If Index = 0 And j >= 0 Then
        If lstSelect(Index).ListIndex < 0 Then lstSelect(Index).ListIndex = 0
    End If
End Sub

Private Sub txtInput_LostFocus()
    mblnDateFouces = False
End Sub

Private Sub txtLst_GotFocus()
    PicLst.Tag = 1
    Call zlControl.TxtSelAll(txtLst)
    lstSelect(0).ListIndex = -1
End Sub

Private Sub txtLst_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnAllow As Boolean
    
    If KeyCode = vbKeyReturn And Shift = vbShiftMask Then Exit Sub
    If KeyCode = vbKeyReturn Then
        mblnDateFouces = True
        Call VsfData_KeyDown(vbKeyReturn, 0)
    ElseIf KeyCode = vbKeyLeft And txtLst.SelStart = 0 Then
        mblnDateFouces = True
        Call MoveNextCell(KeyCode)
    ElseIf Shift = vbShiftMask And KeyCode = vbKeyDown Then
        KeyCode = 0
        lstSelect(0).SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        lstSelect(0).Visible = False
        txtLst.Visible = False
        PicLst.Visible = False
        mblnShow = False
        VsfData.SetFocus
    End If
End Sub

Private Sub txtUpInput_GotFocus()
    Call zlControl.TxtSelAll(txtUpInput)
End Sub

Private Sub txtUpInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Shift = vbCtrlMask Then
            Exit Sub
        Else
            txtDnInput.SetFocus
        End If
    End If
    
    If KeyCode = vbKeyLeft And txtUpInput.SelStart = 0 Then
        Call MoveNextCell(vbKeyLeft)
    End If
End Sub

Private Sub txtUpInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        picDouble.Visible = False
        picDouble.Tag = ""
        mblnShow = False
        VsfData.SetFocus
    End If
End Sub

Private Sub txtUpInput_LostFocus()
    mblnDateFouces = False
End Sub

Private Sub txt显示天数_GotFocus()
    Call zlControl.TxtSelAll(txt显示天数)
End Sub

Private Sub txt显示天数_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnCancel As Boolean
    If KeyCode = vbKeyReturn Then Call txt显示天数_Validate(blnCancel)
End Sub

Private Sub txt显示天数_KeyPress(KeyAscii As Integer)
    If Not ((KeyAscii >= 48 And KeyAscii <= 57) Or KeyAscii = vbKeyBack) Then
        KeyAscii = 0
    End If
End Sub

Private Sub txt显示天数_Validate(Cancel As Boolean)
    If Val(txt显示天数.Text) = Val(txt显示天数.Tag) Then Exit Sub
    txt显示天数.Tag = txt显示天数.Text
    Call RefreshHistoryData(VsfData.Row)
End Sub

Private Sub UserControl_Initialize()
    '初始化病人选择器
    Dim objCol As ReportColumn
    With rptPati
        Set objCol = .Columns.Add(c_选择, "", 18, False): objCol.AllowDrag = False
        Set objCol = .Columns.Add(c_图标, "", 18, False): objCol.Alignment = xtpAlignmentCenter: objCol.AllowDrag = False
        Set objCol = .Columns.Add(c_排序, "排序", 0, True)
        Set objCol = .Columns.Add(c_状态, "状态", 0, True)
        Set objCol = .Columns.Add(c_床号, "床号", 40, True)
        Set objCol = .Columns.Add(C_病人ID, "病人ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(c_主页ID, "主页ID", 0, False): objCol.Visible = False
        Set objCol = .Columns.Add(c_姓名, "姓名", 60, True)
        Set objCol = .Columns.Add(c_年龄, "年龄", 60, True)
        Set objCol = .Columns.Add(c_住院号, "住院号", 60, True)
        Set objCol = .Columns.Add(c_入院日期, "入院日期", 120, True)
        Set objCol = .Columns.Add(c_出院日期, "出院日期", 120, True)
        For Each objCol In .Columns
            If objCol.Index <> c_选择 Then
                objCol.Editable = False
            Else
                objCol.Sortable = True
                objCol.Editable = True
            End If
            objCol.Groupable = (objCol.Index = c_状态)
        Next
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有病人..."
        End With
        
        .PreviewMode = False
        .AllowColumnSort = True
        .AllowColumnRemove = False
        .MultipleSelection = False
        .ShowItemsInGroups = False
        .SetImageList UserControl.imgRPT
        
        .GroupsOrder.DeleteAll
        .GroupsOrder.Add .Columns.Find(c_排序)
        .GroupsOrder(0).SortAscending = True
        .SortOrder.Add .Columns.Find(c_床号)
    End With
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then Call InitCons
End Sub

Private Sub UserControl_Terminate()
    Dim strValue As String
    Dim i As Integer
    Dim arrValue() As String, ArrCode() As String
    
    mstrNote = ""
    If Not (mrsItems Is Nothing) Then Set mrsItems = Nothing
    If Not (mrsPati Is Nothing) Then Set mrsPati = Nothing
    If Not (mrsCell Is Nothing) Then Set mrsCell = Nothing
    If Not (mrsPart Is Nothing) Then Set mrsPart = Nothing
    If Not (mrsCopy Is Nothing) Then Set mrsCopy = Nothing
    If Not (mrsData Is Nothing) Then Set mrsData = Nothing
    '保存过滤条件信息
'    If Val(txtFilter.Tag) = 0 Then
'        strValue = "1;1;1;1"
'    Else
'        strValue = "0;0;0;0"
'        arrValue = Split(strValue, ";")
'        ArrCode = Split(txtFilter.Tag, ";")
'        For i = 0 To UBound(ArrCode)
'            arrValue(Val(ArrCode(i)) - 1) = 1
'        Next i
'        strValue = Join(arrValue, ";")
'    End If
    
    'Call zlDatabase.SetPara("体温单过滤条件", strValue, glngSys, 1255)
End Sub

Private Sub VsfData_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strInfo As String
    Dim strCols As String
    Dim intMax As Integer
    Dim lngItemNO As Long
    Dim strText As String, strPart As String, strKey As String
    Dim lngCol As Long
    Dim cbrControl As CommandBarControl
    Dim blnCheck As Boolean
    
    If mblnInit = False Then Exit Sub
    If OldRow = NewRow And OldCol = NewCol Then Exit Sub
    
    Call AdjustRowFlag(VsfData, NewRow)
    mblnClearRow = False
    
    If NewRow >= VsfData.FixedRows Then
        For lngCol = c时间 + 1 To VsfData.Cols - 1
            If Trim(VsfData.TextMatrix(NewRow, lngCol)) <> "" Then
                mblnClearRow = True
                Exit For
            End If
        Next lngCol
    End If
        
    If NewCol >= RootCol And NewRow >= VsfData.FixedRows Then
        lngItemNO = Val(VsfData.TextMatrix(0, NewCol))
    Else
        If NewCol <> c体温标识 Then
            Call AddActiveMenu(0)
            GoTo ErrInfo
        End If
    End If
    '显示当前项目的相关信息
    mrsItems.Filter = 0
    mrsItems.Filter = "项目序号=" & lngItemNO
    If mrsItems.RecordCount <> 0 Then
        If Nvl(mrsItems!项目值域) <> "" Then
            If mrsItems!项目类型 = 0 Then
                strInfo = "有效范围:" & Split(mrsItems!项目值域, ";")(0) & "～" & Split(mrsItems!项目值域, ";")(1)
            Else
                strInfo = "有效范围:" & mrsItems!项目值域
            End If
        Else
            strInfo = ""
        End If
        
        If lngItemNO = gint体温 Then
            strInfo = strInfo & Space(4) & "物理降温:38/37"
        ElseIf lngItemNO = gint脉搏 And mint心率应用 = 2 Then
            strInfo = strInfo & Space(4) & "脉搏短轴:100/120"
        ElseIf lngItemNO = 4 Then
            strInfo = strInfo & Space(4) & "收缩压/舒张压:110/90"
        End If
        
        If Trim(Nvl(mrsItems!分组名)) = "1)体温曲线项目" Then
             strInfo = strInfo & Space(4) & "按↓进行未记说明选择"
        End If
        
        '体温曲线项目才有部位信息
        If Trim(Nvl(mrsItems!分组名)) <> "1)体温曲线项目" Then
             lngItemNO = 0
        Else
            If Val(VsfData.TextMatrix(VsfData.Row, c病人ID)) = 0 Or Val(VsfData.TextMatrix(VsfData.Row, c文件ID)) = 0 Then lngItemNO = 0
        End If
        
        Call AddActiveMenu(lngItemNO)
        
        If lngItemNO <> 0 Then
            strText = Trim(VsfData.TextMatrix(NewRow, NewCol))
            If strText = "" Then
                strPart = ""
            Else
                strKey = NewRow & "," & NewCol
                mrsCell.Filter = "ID='" & strKey & "'"
                strPart = ""
                If mrsCell.RecordCount > 0 Then
                    strPart = Trim(Nvl(mrsCell!部位))
                End If
            End If
            
            If strPart = "" Then
                mrsPart.Filter = "项目序号=" & lngItemNO & " and 缺省项=1"
                If mrsPart.RecordCount > 0 Then strPart = Trim(Nvl(mrsPart!部位))
                If lngItemNO = gint呼吸 And strPart = "" Then
                    strPart = "自主呼吸"
                End If
            End If
            
            '根据部位信息选择部位菜单的部位
            For Each cbrControl In mcbrToolBar.Controls(4).CommandBar.Controls
                If Trim(cbrControl.Parameter) = Trim(strPart) Then
                    cbrControl.Checked = True
                Else
                    cbrControl.Checked = False
                End If
            Next
        Else
            '确定大便或入液类型
            lngItemNO = Val(VsfData.TextMatrix(0, NewCol))
            strText = Trim(VsfData.TextMatrix(NewRow, NewCol))
            blnCheck = False
            For Each cbrControl In mcbrToolBar.Controls(5).CommandBar.Controls
                cbrControl.Checked = False
                If lngItemNO = gint大便 Then
                    Select Case cbrControl.Id
                        Case conMenu_Edit_Append * 10 + 1
                            cbrControl.Checked = (InStr(1, UCase(strText), "/E") = 0 And InStr(1, UCase(strText), "E") > 0)
                        Case conMenu_Edit_Append * 10 + 2
                            cbrControl.Checked = (InStr(1, UCase(strText), "/E") > 0)
                        Case conMenu_Edit_Append * 10 + 3
                            cbrControl.Checked = (UCase(strText) = "*" Or UCase(strText) = "※")
                        Case conMenu_Edit_Append * 10 + 4
                            cbrControl.Checked = (UCase(strText) = "☆")
                    End Select
                   
                ElseIf lngItemNO = gint入液 Then
                    Select Case cbrControl.Id
                        Case conMenu_Edit_Append * 10 + 5
                            cbrControl.Checked = (InStr(1, UCase(strText), "/C") = 0 And InStr(1, UCase(strText), "C") > 0)
                        Case conMenu_Edit_Append * 10 + 6
                            cbrControl.Checked = InStr(1, UCase(strText), "/C") > 0
                    End Select
                End If
                If blnCheck = False Then blnCheck = cbrControl.Checked
            Next
            If blnCheck = False Then
                 mcbrToolBar.Controls(5).CommandBar.Controls(1).Checked = True
            End If
        End If
    End If

    mrsItems.Filter = 0
    
ErrInfo:
    RaiseEvent AfterRowColChange(strInfo, False)
    '提取该病人历史数据
    If OldRow <> NewRow Then
        Call RefreshHistoryData(NewRow)
    End If
End Sub

Private Sub VsfData_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    If Not mblnInit Then Exit Sub
    Call InitCons
End Sub

Private Sub VsfData_DblClick()
    Call VsfData_KeyDown(Asc("L"), 0)
End Sub

Private Sub VsfData_EnterCell()
    Dim lngItemNO As Long
    Dim strName As String
    Dim int护理等级 As Integer
    Dim strKey As String, strInfo As String
    Dim rsObj As New ADODB.Recordset
    
    picInput.Visible = False
    lstNote.Visible = False
    picInput.Tag = ""
    lstNote.Tag = ""
    txtInput.Tag = ""
    picDouble.Visible = False
    picDouble.Tag = ""
    lstSelect(0).Visible = False
    lstSelect(1).Visible = False
    PicLst.Visible = False
    PicLst.Tag = ""
    txtLst.Visible = False
    cbo体温标识.Visible = False
    
    mintType = 0
    
    VsfData.SetFocus
    
    If Not mblnInit Then Exit Sub
    If Not mblnShow Then Exit Sub
    If VsfData.Col < RootCol - 1 And VsfData.Col <> c体温标识 Then Exit Sub
    '如果无病人信息也不能编辑
    If Val(VsfData.TextMatrix(VsfData.Row, c病人ID)) = 0 Or Val(VsfData.TextMatrix(VsfData.Row, c文件ID)) = 0 Then Exit Sub
    
    '如果数据已经保存，并且该列存在同步过来的数据。就不允许修改时间
    If VsfData.Col = c时间 Then
        mrsCopy.Filter = 0
        mrsCopy.Filter = "行号=" & VsfData.Row
        Do While Not mrsCopy.EOF
            mint数据来源 = Val(Nvl(mrsCopy!数据来源))
            If InStr(1, ",0,9,", "," & mint数据来源 & ",") = 0 Then
                strInfo = "此行数据已经保存并且包含同步过来的数据,不能修改时间."
                RaiseEvent AfterRowColChange(strInfo, True)
                Exit Sub
            End If
            mrsCopy.MoveNext
        Loop
    End If
    
    mint数据来源 = 0
    mintModify = 0
    strName = VsfData.TextMatrix(VsfData.FixedRows - 1, VsfData.Col)
    lngItemNO = Val(VsfData.TextMatrix(0, VsfData.Col))
    int护理等级 = Val(VsfData.TextMatrix(VsfData.Row, c护理等级))
    
    '检查护理等级和适用病人
    If VsfData.Col >= RootCol Then
        mrsItems.Filter = "项目序号=" & lngItemNO & " And 护理等级>=" & int护理等级
        If mrsItems.RecordCount = 0 Then
            strInfo = "项目[" & strName & "]的护理等级不适用该病人."
            RaiseEvent AfterRowColChange(strInfo, True)
            Exit Sub
        End If
        
        '是否适用病人
        If Val(VsfData.TextMatrix(2, VsfData.Col)) = 1 Then
            If Val(VsfData.TextMatrix(VsfData.Row, c婴儿)) <> 0 Then
                strInfo = "项目[" & strName & "]只适用于病人."
                RaiseEvent AfterRowColChange(strInfo, True)
                Exit Sub
            End If
        ElseIf VsfData.TextMatrix(2, VsfData.Col) = 2 Then
           If Val(VsfData.TextMatrix(VsfData.Row, c婴儿)) = 0 Then
                strInfo = "项目[" & strName & "]只适用于婴儿."
                RaiseEvent AfterRowColChange(strInfo, True)
                Exit Sub
            End If
        End If
    End If
    
    '检查数据是否是同步过来的
    mrsCell.Filter = 0
    strKey = VsfData.Row & "," & VsfData.Col
    mrsCell.Filter = "ID='" & strKey & "'"
    If mrsCell.RecordCount > 0 Then '保存后mrsCell为空
        Set rsObj = mrsCell.Clone
    Else
        Set rsObj = mrsCopy.Clone
    End If
    rsObj.Filter = "ID='" & strKey & "'"
    
    If rsObj.RecordCount > 0 Then
        lngItemNO = Val(Nvl(rsObj!项目序号))
        mint数据来源 = Val(Nvl(rsObj!数据来源))
        mintModify = Val(Nvl(rsObj!修改))
        If InStr(1, ",0,9,", "," & Val(rsObj!数据来源) & ",") = 0 Then
            If Not (lngItemNO = gint体温 Or (lngItemNO = gint脉搏 And mint心率应用 = 2)) Then
                strInfo = "同步过来的[" & strName & "]数据不能进行修改."
                RaiseEvent AfterRowColChange(strInfo, True)
                Exit Sub
            Else
                If mintModify = 1 Then
                    If lngItemNO = gint体温 Then
                        strInfo = "同步过来的[" & strName & "]数据如果包含物理降温不能进行修改."
                    Else
                        strInfo = "同步过来的[" & strName & "]数据如果包含脉搏短轴不能进行修改."
                    End If
                    RaiseEvent AfterRowColChange(strInfo, True)
                    Exit Sub
                End If
            End If
        End If
    End If
    
    Call ShowInput
End Sub

Private Sub VsfData_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
       '跳到下一行或下一列
       Call MoveNextCell
    Else
        If Not (KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or Shift <> 0) Then
            mblnShow = True
            Call VsfData_EnterCell
        End If
    End If
End Sub

Private Sub VsfData_GotFocus()
    picFilter.Visible = False
    picPati.Visible = False
End Sub

Private Sub ShowInput()
'----------------------------------
'功能显示输入框信息
'----------------------------------
    Dim strText As String, strText1 As String, strPart As String
    Dim intCOl As Integer, intRow As Integer
    Dim CellRect As RECT
    Dim lngItemNO As Long
    Dim intType As Integer, intIndex As Integer
    Dim strLen As String
    Dim strTmp As String, strPoint As String
    Dim arrValue() As String, arrValue1() As String
    Dim blnSelect As Boolean
    Dim i As Integer, j As Integer
    
    Call InitCons
    intType = -1
    intCOl = VsfData.Col
    intRow = VsfData.Row
    
    CellRect.Left = VsfData.CellLeft + VsfData.Left
    CellRect.Top = VsfData.CellTop + VsfData.Top
    CellRect.Bottom = VsfData.CellHeight + 20
    CellRect.Right = VsfData.CellWidth + 20
    
    strPart = ""
    If intCOl = c时间 Then
        strText1 = Trim(VsfData.TextMatrix(intRow, intCOl))
        If strText1 = "" Then
            '如果用户已经录入时点信息，则下面的时点以此时点为准
            If Not IsDate(mstrModifyTime) Then
                strText = Format(zlDatabase.Currentdate, "HH:mm")
            Else
                strText = Format(mstrModifyTime, "HH:mm")
            End If
        Else
            strText = Format(strText1, "HH:mm")
        End If
        intType = -1
    ElseIf intCOl = c体温标识 Then
        Call zlControl.CboLocate(cbo体温标识, VsfData.TextMatrix(intRow, intCOl))
        intType = -2
    Else
        strText = Trim(VsfData.TextMatrix(intRow, intCOl))
        If InStr(1, strText, ":") <> 0 Then
            strPart = Trim(Split(strText, ":")(0))
            strText = Trim(Split(strText, ":")(1))
        End If
        strText1 = strText
        lngItemNO = VsfData.TextMatrix(0, intCOl)
        intType = 0
    End If
    
    If intType = 0 Then
        If lngItemNO <> 4 Then
            mintType = 1
            mrsItems.Filter = "项目序号=" & lngItemNO
            If InStr(1, ",2,3,5,", "," & Val(Nvl(mrsItems!项目表示)) & ",") = 0 Then
                strLen = Nvl(mrsItems!项目长度, 0) & ";" & Nvl(mrsItems!项目小数, 0)
                If lngItemNO = gint体温 Or (lngItemNO = gint脉搏 And mint心率应用 = 2) Then
                    strLen = (Val(Split(strLen, ";")(0)) + Val(Split(strLen, ";")(0)) + 1) & ";" & IIf(Val(Split(strLen, ";")(1)) = 0, 0, 1) * 2
                End If
                
                If Val(strLen) <> 0 Then
                    txtInput.MaxLength = Val(Split(strLen, ";")(0)) + IIf(Val(Split(strLen, ";")(1)) = 0, 0, 1)
                Else
                    txtInput.MaxLength = 0
                End If
            Else
                mintType = Val(Nvl(mrsItems!项目表示))
                strText1 = Nvl(mrsItems!项目值域, ";")
            End If
        Else
            mintType = 4
            mrsItems.Filter = "项目序号=4 or 项目序号=5"
            mrsItems.Sort = "项目序号"
            Do While Not mrsItems.EOF
                strTmp = Val(strTmp) + Val(Nvl(mrsItems!项目长度))
                strPoint = Val(strPoint) + Val(Nvl(mrsItems!项目小数))
                strLen = strTmp & ";" & strPoint
                Select Case Val(mrsItems!项目序号)
                    Case 4
                        If Val(strLen) <> 0 Then
                            txtUpInput.MaxLength = Val(Split(strLen, ";")(0)) + IIf(Val(Split(strLen, ";")(1)) = 0, 0, 1)
                        Else
                            txtUpInput.MaxLength = 0
                        End If
                    Case 5
                        If Val(strLen) <> 0 Then
                            txtDnInput.MaxLength = Val(Split(strLen, ";")(0)) + IIf(Val(Split(strLen, ";")(1)) = 0, 0, 1)
                        Else
                            txtDnInput.MaxLength = 0
                        End If
                End Select
            mrsItems.MoveNext
            Loop
        End If
    ElseIf intType = -1 Then
        mintType = 1
        txtInput.MaxLength = 5
    Else
        mintType = -2
    End If
    
    Select Case mintType
        Case -2 '体温标识
            With cbo体温标识
                .Left = CellRect.Left
                .Top = CellRect.Top
                .Width = CellRect.Right
                .FontName = VsfData.FontName
                .Font.Size = VsfData.Font.Size
                .Visible = True
                .ZOrder 0
            End With
        Case 1
            With picInput
                .Left = CellRect.Left
                .Top = CellRect.Top
                .Width = CellRect.Right
                .Height = CellRect.Bottom
                .FontName = VsfData.FontName
                .Font.Size = VsfData.Font.Size
                .Visible = True
                .ZOrder 0
            End With
            
            lblCheck.Visible = False
            
            With txtInput
                .Top = 0
                .Left = 0
                .Width = CellRect.Right
                .Height = CellRect.Bottom
                .FontName = VsfData.FontName
                .Font.Size = VsfData.Font.Size
                .Width = .Width - (180 + IIf(mintBigSize, 180 * 1 / 3, 0)) / 2 '宋体9号时减去90,字体越大扣除的边距越小,以保证文本框分行与实际一致
                .Visible = True
                .Text = strText
                .Tag = strPart  'intRow & "|" & intCOl
                .ZOrder 0
                picInput.Tag = strText1
            End With
            
            picInput.SetFocus
        Case 2, 3 '单选或多选
            Select Case mintType
                Case 2
                    intIndex = 0
                    If Left(strText1, 1) <> ";" Then strText1 = ";" & strText1
                Case 3
                    intIndex = 1
            End Select
            
            strText = Trim(VsfData.TextMatrix(intRow, intCOl))
            arrValue = Split(strText1, ";") '值域
            lstSelect(intIndex).Clear
        
            PicLst.Tag = "1"
            For i = 0 To UBound(arrValue)
                If Left(arrValue(i), 1) = "√" Then arrValue(i) = Mid(arrValue(i), 2): strText1 = arrValue(i)
                lstSelect(intIndex).AddItem arrValue(i), i
                 
                If intIndex = 0 Then
                   ReDim arrValue1(0)
                   arrValue1(0) = strText
                   txtLst.Text = strText
                Else
                   arrValue1 = Split(strText, ",")
                End If
                For j = 0 To UBound(arrValue1)
                    If arrValue1(j) = arrValue(i) Then
                        lstSelect(intIndex).Selected(i) = True
                        blnSelect = True
                    End If
                Next j
            Next i
            If blnSelect = False And strText1 <> "" And IIf(intIndex = 0, Trim(txtLst.Text) = "", True) Then
                For i = 0 To lstSelect(intIndex).ListCount - 1
                    If lstSelect(intIndex).List(i) = strText1 Then
                        lstSelect(intIndex).Selected(i) = True
                    End If
                Next i
            End If
            If lstSelect(intIndex).ListIndex >= 0 Then txtLst.Text = "": PicLst.Tag = "0"
            
            '控件显示
            '51600,刘鹏飞,2012-07-16,单选项目提供可以选择和录入功能
            If intIndex = 0 Then
                mrsItems.Filter = "项目序号=" & lngItemNO
                If mrsItems.RecordCount > 0 Then
                    strLen = Nvl(mrsItems!项目长度, 0) & ";" & Nvl(mrsItems!项目小数, 0)
                End If
                With PicLst
                    .FontName = VsfData.FontName
                    .FontSize = VsfData.FontSize
                    .Left = CellRect.Left
                    .Top = CellRect.Top
                    .Height = 80 + CellRect.Bottom + PicLst.TextHeight("刘") * 2 + lstSelect(intIndex).ListCount * (PicLst.TextHeight("刘") + PicLst.TextHeight("刘") \ 4)
                    If .Height < CellRect.Bottom Then .Height = CellRect.Bottom
                    .Width = LenB(StrConv(lstSelect(intIndex).List(lstSelect(intIndex).ListCount \ 2), vbFromUnicode)) * 100 + 500    '以中间项的长度为依据
                    If .Width < CellRect.Right Then .Width = CellRect.Right
                    If .Height > VsfData.Height Then
                        .Height = VsfData.Height
                    End If
                    If .Top + .Height > VsfData.Height Then
                        .Top = CellRect.Top + CellRect.Bottom - .Height
                    End If
                    If .Top < 0 Then .Top = VsfData.Top
                
                    PicLst.Visible = True
                    PicLst.ZOrder 0
                End With
                
                With lbllst(0)
                    .Left = 20
                    .Top = 20
                    If .Width > PicLst.Width Then
                        PicLst.Width = .Width + PicLst.TextWidth("刘")
                    End If
                    .FontName = VsfData.FontName
                    .FontSize = VsfData.FontSize
                    .Visible = True
                End With
                
                With txtLst
                    .Top = lbllst(0).Top + lbllst(0).Height + 20
                    .Left = -10
                    .Width = PicLst.Width
                    .Height = CellRect.Bottom
                    .FontName = VsfData.FontName
                    .FontSize = VsfData.FontSize
                    .MaxLength = Val(Split(strLen, ";")(0)) + IIf(Val(Split(strLen, ";")(1)) = 0, 0, 1)
                    .Visible = True
                End With
                
                With lbllst(1)
                    .Left = 20
                    .Top = txtLst.Top + txtLst.Height + 20
                    .FontName = VsfData.FontName
                    .FontSize = VsfData.FontSize
                    .Visible = True
                End With
                
                With lstSelect(intIndex)
                    .Top = lbllst(1).Top + lbllst(1).Height + 20
                    .Left = -10
                    .FontName = VsfData.FontName
                    .FontSize = VsfData.FontSize
                    .Width = PicLst.Width
                    .Height = PicLst.Height - .Top
                    .Visible = True
                    .Enabled = True
                    .ZOrder 0
                    .Tag = strText
                End With
                If lstSelect(intIndex).Top + lstSelect(intIndex).Height <> PicLst.Height Then
                    PicLst.Height = lstSelect(intIndex).Top + lstSelect(intIndex).Height
                End If
                PicLst.SetFocus
            Else
                lstSelect(intIndex).Top = CellRect.Top
                lstSelect(intIndex).Left = CellRect.Left
                lstSelect(intIndex).FontName = VsfData.FontName
                lstSelect(intIndex).FontSize = VsfData.FontSize
                lstSelect(intIndex).Height = lstSelect(intIndex).ListCount * (PicLst.TextHeight("刘") + PicLst.TextHeight("刘") \ 4)
                If lstSelect(intIndex).Height < CellRect.Bottom Then lstSelect(intIndex).Height = CellRect.Bottom
                lstSelect(intIndex).Width = LenB(StrConv(lstSelect(intIndex).List(lstSelect(intIndex).ListCount \ 2), vbFromUnicode)) * 100 + 500    '以中间项的长度为依据
                If lstSelect(intIndex).Width < CellRect.Right Then lstSelect(intIndex).Width = CellRect.Right
                If lstSelect(intIndex).Height > VsfData.Height Then
                    lstSelect(intIndex).Height = VsfData.Height
                End If
                If lstSelect(intIndex).Top + lstSelect(intIndex).Height > VsfData.Height Then
                    lstSelect(intIndex).Top = CellRect.Top + CellRect.Bottom - lstSelect(intIndex).Height
                End If
                If lstSelect(intIndex).Top < 0 Then lstSelect(intIndex).Top = VsfData.Top
                
                lstSelect(intIndex).Visible = True
                lstSelect(intIndex).Enabled = True
                lstSelect(intIndex).ZOrder 0
                
                lstSelect(intIndex).Tag = strText
                lstSelect(intIndex).SetFocus
            End If
        Case 4 '血压
            With picDouble
                .Left = CellRect.Left
                .Top = CellRect.Top
                .Width = CellRect.Right
                .Height = CellRect.Bottom
                .FontName = VsfData.FontName
                .Font.Size = VsfData.Font.Size
                .Visible = True
                .Tag = strText1
                .ZOrder 0
            End With
            
            If strText = "" Then strText = "/"
            arrValue = Split(strText, "/")
            
            lblSplit.FontName = VsfData.FontName
            lblSplit.FontSize = VsfData.FontSize
            lblSplit.Left = (picDouble.Width - lblSplit.Width) / 2
            If mintBigSize = 1 Then
                lblSplit.Width = 150
            Else
                lblSplit.Width = 105
            End If
    
            With txtUpInput
                .Text = arrValue(0)
                .FontName = VsfData.FontName
                .Font.Size = VsfData.Font.Size
                .Width = (picDouble.Width - lblSplit.Width) * 0.4
                .ZOrder 0
            End With
            
            With txtDnInput
                .Text = arrValue(1)
                .FontName = VsfData.FontName
                .Font.Size = VsfData.Font.Size
                .Left = lblSplit.Left + lblSplit.Width
                .Width = picDouble.Width - .Left
                .ZOrder 0
            End With
            
            picDouble.SetFocus
        Case 5 '选择
            With picInput
                .Left = CellRect.Left
                .Top = CellRect.Top
                .Width = CellRect.Right
                .Height = CellRect.Bottom
                .FontName = VsfData.FontName
                .Font.Size = VsfData.Font.Size
                .Visible = True
                .ZOrder 0
            End With
            
            txtInput.Visible = False
            
            With lblCheck
                .Top = 0
                .Left = 0
                .Caption = strText
                .Width = CellRect.Right
                .Height = CellRect.Bottom
                .FontName = VsfData.FontName
                .Font.Size = VsfData.Font.Size
                .Width = .Width - (180 + IIf(mintBigSize, 180 * 1 / 3, 0)) / 2 '宋体9号时减去90,字体越大扣除的边距越小,以保证文本框分行与实际一致
                .Visible = True
                .ZOrder 0
                picInput.Tag = strText1
            End With
            
            picInput.SetFocus
    End Select
End Sub

Private Sub MoveNextCell(Optional KeyCode As Integer = vbKeyReturn)
'--------------------------------------------
'功能:检查数据并赋值 、移动到下一行或下一列
'--------------------------------------------
    Dim lngItemNO As Integer, i As Integer, intIndex As Integer
    Dim strText As String, strErrMsg As String, strPatiTime As String, strOldValue As String
    Dim intCOl As Integer, intRow As Integer
    Dim blnValidate As Boolean, blnSave As Boolean
    Dim strFileds As String, strValues As String, strKey As String, strPart As String
    Dim int来源ID As Long, int共用 As Integer, int显示 As Integer, int修改 As Integer
    Dim intState As Integer
    Dim strData As String
   
    'If picInput.Visible = False Then Exit Sub
    
    If mblnInit = False Then Exit Sub
    
    strFileds = "ID|行号|项目序号|数据|部位|数据来源|来源ID|共用|显示|修改|状态"
    intCOl = VsfData.Col
    intRow = VsfData.Row
    blnValidate = False
    blnSave = False
    strOldValue = ""
    If KeyCode = vbKeyReturn And InStr(1, ",0,-2,", "," & mintType & ",") = 0 Then ' (picInput.Visible = True Or picDouble.Visible = True) Then
        mlng文件ID = Val(VsfData.TextMatrix(intRow, c文件ID))
        mlng病人ID = Val(VsfData.TextMatrix(intRow, c病人ID))
        mlng主页ID = Val(VsfData.TextMatrix(intRow, c主页ID))
        mlngBaby = Val(VsfData.TextMatrix(intRow, c婴儿))
        strPatiTime = VsfData.TextMatrix(intRow, c日期)
        mbln出院 = (Val(VsfData.TextMatrix(intRow, c出院)) = 1)
        
        Select Case mintType
            Case 1
                strText = Trim(txtInput.Text)
'                If InStr(txtInput.Text, "/") > 0 And mbln脉搏共用显示 Then
'                    strData = Split(Trim(txtInput.Text), "/")(1) & "/" & Split(Trim(txtInput.Text), "/")(0)
'                    strText = strData
'                End If
                
                strPart = Trim(txtInput.Tag)
                strOldValue = picInput.Tag
            Case 2, 3
                If mintType = 2 Then
                    intIndex = 0
                Else
                    intIndex = 1
                End If
                strText = ""
                strPart = ""
                For i = 0 To lstSelect(intIndex).ListCount - 1
                  If lstSelect(intIndex).Selected(i) = True Then
                      strText = strText & "," & Replace(lstSelect(intIndex).List(i), ",", "")
                  End If
                Next i
                If Left(strText, 1) = "," Then strText = Mid(strText, 2)
                '51600，刘鹏飞，2012-07-16，单选即可录入也可以输入
                If intIndex = 0 And Val(PicLst.Tag) = 1 Then strText = Trim(txtLst.Text)
                strOldValue = lstSelect(intIndex).Tag
            Case 4
                strText = Trim(txtUpInput.Text) & "/" & Trim(txtDnInput.Text)
                strPart = ""
                If strText = "/" Then strText = ""
                strOldValue = picDouble.Tag
            Case 5
                strText = lblCheck.Caption
                strPart = ""
                strOldValue = picInput.Tag
        End Select
        
        '检查时间和数据是否合法
        If intCOl = c时间 Then
            If Not CheckDateTime(strText, strPatiTime, strErrMsg) Then picInput.SetFocus: GoTo ErrInfo
            '此处重新获取列号,因为对于保存的数据修改时间后会删除原有时间数据，派生出一条本时间新的数据(隐藏该列，复制一行新数据)
            intRow = VsfData.Row
            mstrModifyTime = Format(strText, "HH:mm")
            blnValidate = True
        ElseIf intCOl > c时间 Then
            lngItemNO = Val(VsfData.TextMatrix(0, intCOl))
            If Not CheckValid(strText, lngItemNO, strErrMsg) Then
                Select Case mintType
                    Case 1
                        picInput.SetFocus
                    Case 2, 3
                        If mintType = 2 Then
                            intIndex = 0
                        Else
                            intIndex = 1
                        End If
                        lstSelect(intIndex).SetFocus
                    Case 4
                        picDouble.SetFocus
                    Case Else
                        picInput.SetFocus
                End Select
                GoTo ErrInfo
            End If
            blnValidate = True
            If mlng病人ID = 0 Or mlng文件ID = 0 Or mlng主页ID = 0 Then
                blnSave = False
            Else
                blnSave = True
            End If
        End If
        
        If blnValidate = True Then
            mrsCopy.Filter = 0
            VsfData.TextMatrix(intRow, intCOl) = IIf(strPart = "", "", strPart & ":") & strText
            VsfData.Cell(flexcpAlignment, intRow, intCOl, intRow, intCOl) = flexAlignCenterCenter
            '进行数据处理
            If blnSave = True Then
                    strKey = intRow & "," & intCOl
                    '检查修改的数据是否已经保存
                    mrsCopy.Filter = "ID='" & strKey & "'"
                    If mrsCopy.RecordCount > 0 Then
                        int来源ID = Val(Nvl(mrsCopy!来源ID))
                        int共用 = Val(Nvl(mrsCopy!共用))
                        int显示 = Val(Nvl(mrsCopy!显示))
                        int修改 = Val(Nvl(mrsCopy!修改))
                        strPatiTime = Nvl(mrsCopy!部位)
                        intState = 1
                    Else
                        int来源ID = 0: int共用 = 0: int显示 = 0: int修改 = 0: strPatiTime = ""
                        intState = IIf(Trim(strText) = "", 3, 1)
                        mrsCell.Filter = "ID='" & strKey & "'"
                        If mrsCell.RecordCount > 0 Then
                            int来源ID = Val(Nvl(mrsCell!来源ID))
                            int共用 = Val(Nvl(mrsCell!共用))
                            int显示 = Val(Nvl(mrsCell!显示))
                            int修改 = Val(Nvl(mrsCell!修改))
                            strPatiTime = Nvl(mrsCell!部位)
                        End If
                    End If
                    If Trim(strOldValue) <> Trim(strText) Or (Trim(strText) <> "" And strPart <> strPatiTime) Then
                        strValues = strKey & "|" & intRow & "|" & lngItemNO & "|" & strText & "|" & strPart & "|" & mint数据来源 & "|" & _
                            int来源ID & "|" & int共用 & "|" & int显示 & "|" & int修改 & "|" & intState
                        Call Record_Update(mrsCell, strFileds, strValues, "ID|" & strKey)
                    End If
                    mblnChage = True
            End If
        End If
    End If
    
    mintType = 0
    '开始移动行或列
    With VsfData
        If KeyCode = vbKeyReturn Then
          
NextCol2: '跳到下一行
            If .Col < .FixedCols Then
                .Col = .Col + 1: GoTo NextCol2
            End If
            If .Col < .Cols - 1 Then
                .Col = .Col + 1
                If .ColHidden(.Col) = True Then GoTo NextCol2
            Else
NextRow2: '跳到下一列
                If .Row < .Rows - 1 Then
                    intRow = .Row + 1
                    If .RowHidden(intRow) = True Then GoTo NextRow2
                    intCOl = c时间
                    .Select intRow, intCOl
                Else
                    intRow = .Row
                    intCOl = c时间
                    .Select intRow, intCOl
                End If
            End If
            '如果该列或行不可见就自动显示该列
            If .ColIsVisible(.Col) = False Then
                .LeftCol = .Col
            End If
            If .RowIsVisible(.Row) = False Then
                .TopRow = .Row
            End If

            Exit Sub
        End If
        '左键
        If KeyCode = vbKeyLeft Then
PreCol2:
            If .Col > c时间 Then
                .Col = .Col - 1
                If .ColHidden(.Col) = True Then GoTo PreCol2
            Else
PreRow2:
                If .Row > .FixedRows Then
                    intRow = .Row - 1
                    If .RowHidden(intRow) Then GoTo PreRow2
                    intCOl = .Cols - 1
                    .Select intRow, intCOl
                End If
            End If
            '如果该列或行不可见就自动显示该列
            If .ColIsVisible(.Col) = False Then
                .LeftCol = .Col
            End If
            If .RowIsVisible(.Row) = False Then
                .TopRow = .Row
            End If
            Exit Sub
        End If
    End With
    
    Exit Sub
ErrInfo:
    RaiseEvent AfterRowColChange(strErrMsg, True)
End Sub

Private Function SaveDate() As Boolean
'------------------------------------------
'功能:保存数据信息
'------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim blnTrans As Boolean, blnTimeNull As Boolean
    Dim lngRow As Long, lngCol As Long, lngItemCode As Long, lngRecordID As Long
    Dim strKey As String, strPart As String, strValue As String
    Dim strTime As String, strEnd As String, strMarkTime As String, strSQL As String
    Dim arrSQL() As String, arrData() As String
    Dim i As Integer, intRow As Integer
    Dim strValues As String, strNote As String, strSaveRows As String
    Dim blnData As Boolean, blnSave As Boolean
    '病人相关信息
    Dim lng文件ID As Long, lng病人ID As Long, lng主页ID As Long, lng婴儿 As Long
    On Error GoTo Errhand
    
    mrsCell.Filter = 0
    '检查有数据的列是否填写时间
    For lngRow = VsfData.FixedRows To VsfData.Rows - 1
        If Val(VsfData.TextMatrix(lngRow, c文件ID)) <> 0 And VsfData.RowHidden(lngRow) = False Then
            blnTimeNull = IIf(Trim(VsfData.TextMatrix(lngRow, c时间)) = "", True, False)
            If blnTimeNull = True Then
                mrsCell.Filter = "行号=" & lngRow & " And 状态=1"
                If mrsCell.RecordCount > 0 Then
                    mblnShow = True
                    VsfData.Select lngRow, c时间
                    RaiseEvent AfterRowColChange("时间不能为空,请录入时间.", True)
                    Exit Function
                End If
            End If
        End If
    Next lngRow
    
    Screen.MousePointer = 11
          
    ReDim Preserve arrSQL(1 To 1)
    
    strSaveRows = ""
    '首先检查时间是否更新
    mrsData.Filter = 0
    For lngRow = VsfData.FixedRows To VsfData.Rows - 1
        If VsfData.RowHidden(lngRow) = False Then
            mrsData.Filter = "行号=" & lngRow
            If mrsData.RecordCount > 0 Then
                lngRecordID = Val(Nvl(mrsData!记录ID))
                If Val(Nvl(mrsData!删除)) = 2 And lngRecordID > 0 Then '表示时间修改
                    strTime = Format(Format(mstrDate, "YYYY-MM-DD") & " " & VsfData.TextMatrix(lngRow, c时间), "YYYY-MM-DD HH:mm:ss")
                    strMarkTime = strTime
                    strMarkTime = "To_Date('" & strMarkTime & "','yyyy-mm-dd hh24:mi:ss')"
                    strSQL = "ZL_体温单数据_发生时间("
                    'ID_IN       IN 病人护理数据.ID%TYPE,
                    strSQL = strSQL & lngRecordID & ","
                    '发生时间_IN IN 病人护理数据.发生时间%TYPE
                    strSQL = strSQL & strMarkTime & ")"
                    arrSQL(ReDimArray(arrSQL)) = strSQL
                    
                    strSaveRows = strSaveRows & "," & lngRow
                End If
            End If
        End If
    Next lngRow
    
    If Left(strSaveRows, 1) = "," Then strSaveRows = Mid(strSaveRows, 2)
    
    intRow = 0
    blnSave = False
    '数据检查成功后开始提取记录集
    mrsCell.Filter = 0
    mrsCell.Sort = "行号"
    With mrsCell
        Do While Not .EOF
            If Val(Nvl(mrsCell!状态)) = 1 Then
                If intRow <> Val(!行号) Then
ErrRow:
                    If blnSave = True Then
                        If InStr(1, "," & strSaveRows & ",", "," & lngRow & ",") = 0 Then
                            strSaveRows = strSaveRows & "," & lngRow
                        End If
                        intRow = lngRow
                        blnSave = False
                        If .EOF Then Exit Do
                    End If
                End If
                
                strKey = !Id
                lngRow = Val(Split(strKey, ",")(0))
                lngCol = Val(Split(strKey, ",")(1))
                
                strTime = Format(Format(mstrDate, "YYYY-MM-DD") & " " & VsfData.TextMatrix(lngRow, c时间), "YYYY-MM-DD HH:mm:ss")
                strEnd = strTime
                strMarkTime = strTime
                
                lngItemCode = Val(!项目序号)
                strPart = Nvl(!部位)
                strValue = Nvl(!数据)
                strNote = ""
                
            
                
                lng文件ID = Val(VsfData.TextMatrix(lngRow, c文件ID))
                lng病人ID = Val(VsfData.TextMatrix(lngRow, c病人ID))
                lng主页ID = Val(VsfData.TextMatrix(lngRow, c主页ID))
                lng婴儿 = Val(VsfData.TextMatrix(lngRow, c婴儿))
                
                strMarkTime = "To_Date('" & strMarkTime & "','yyyy-mm-dd hh24:mi:ss')"

                mrsItems.Filter = 0
                mrsItems.Filter = "项目序号=" & lngItemCode
                If mrsItems!分组名 = "1)体温曲线项目" Then
                    '--记录内容
                    If strValue = "不升" And lngItemCode = gint体温 Then
                        strNote = ""
                    Else
                        If IsNumeric(strValue) Or InStr(1, strValue, "/") > 0 Then
                             strNote = ""
                        Else
                            strNote = strValue
                            strValue = ""
                        End If
                    End If
                Else
                     strNote = ""
                End If
                
                '问题号:56853,修改人:李涛,脉搏心率显示方式：心率/脉搏
                If lngItemCode = gint脉搏 And mint心率应用 = 2 Then
                    If mbln脉搏共用显示 And strValue <> "" Then
                        strValue = Split(Nvl(!数据), "/")(1) & "/" & Split(Nvl(!数据), "/")(0)
                    End If
                End If
                
                If lngItemCode = 4 And (Nvl(!数据) = "未测" Or Nvl(!数据) = "拒测" Or Nvl(!数据) = "外出" Or Nvl(!数据) = "请假") Then
                    strValue = Nvl(!数据) & "/" & Nvl(!数据)
                End If
                    
                '更新数据信息
                strSQL = "Zl_体温单数据_Update("
                '文件id_In   In 病人护理文件.Id%Type,  --病人护理文件ID
                strSQL = strSQL & Val(lng文件ID) & ","
                '发生时间_In In 病人护理数据.发生时间%Type, --护理数据的发生时间
                strSQL = strSQL & strMarkTime & ","
                '记录类型_In In 病人护理明细.记录类型%Type, --护理项目=1，上标说明=2，入出转标记=3，手术日标记=4,下标说明=6
                strSQL = strSQL & "1,"
                '项目序号_In In 病人护理明细.项目序号%Type, --护理项目的序号，非护理项目固定为0
                strSQL = strSQL & lngItemCode & ","
                '记录内容_In In 病人护理明细.记录内容%Type := Null, --记录内容，如果内容为空，即清除以前的内容  36或36/37
                strSQL = strSQL & "'" & strValue & "',"
                '体温部位_In In 病人护理明细.体温部位%Type := Null, --删除数据时不用填写部位 除活动项目外
                strSQL = strSQL & IIf(strValue <> "", "'" & strPart & "'", "NULL") & ","
                '复试合格_In In Number := 0,
                strSQL = strSQL & "NULL,"
                '未记说明_In In 病人护理明细.未记说明%Type := Null, --未记说明
                strSQL = strSQL & "'" & strNote & "',"
                '他人记录_In In Number := 1,
                strSQL = strSQL & "1,"
                '数据来源_In In 病人护理明细.数据来源%Type := 0,
                strSQL = strSQL & "0,"
                '来源id_In   In 病人护理明细.来源id%Type := Null,
                strSQL = strSQL & IIf(Val(Nvl(!来源ID)) = 0, "NULL", Val(Nvl(!来源ID))) & ","
                '共用_In     In 病人护理明细.共用%Type := 0,
                strSQL = strSQL & Val(Nvl(!共用))
                strSQL = strSQL & ")"

                arrSQL(ReDimArray(arrSQL)) = strSQL
                
                If intRow <> Val(!行号) Then blnSave = True
            
            End If
        .MoveNext
        Loop
        If blnSave = True Then GoTo ErrRow
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    '循环执行SQL保存数据
    gcnOracle.BeginTrans
    blnTrans = True
    
    blnData = False
    
    'Debug.Print "---保存开始:" & Now
    For i = 1 To UBound(arrSQL)
        If arrSQL(i) <> "" Then Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "保存体温数据"): blnData = True: 'Debug.Print CStr(arrSQL(i))
    Next
    gcnOracle.CommitTrans
    
    'Debug.Print "---保存结束:" & Now
    blnTrans = False
    
    Screen.MousePointer = 0
    mblnChage = False
    mblnShow = False
    mblnSaveData = True
    
    Call InitCons
    
    If Left(strSaveRows, 1) = "," Then strSaveRows = Mid(strSaveRows, 2)
    '更新记录ID
    For lngRow = VsfData.FixedRows To VsfData.Rows - 1
        blnTimeNull = IIf(Trim(VsfData.TextMatrix(lngRow, c时间)) = "", True, False)
        If Not blnTimeNull And VsfData.RowHidden(lngRow) = False Then
            If InStr(1, "," & strSaveRows & ",", "," & lngRow & ",") <> 0 Then
                strTime = Format(Format(mstrDate, "YYYY-MM-DD") & " " & VsfData.TextMatrix(lngRow, c时间), "YYYY-MM-DD HH:mm:ss")
                strSQL = " Select A.ID From 病人护理数据 A,病人护理文件 B" & vbNewLine & _
                              " Where A.文件ID=B.ID And B.ID=[1] And A.发生时间=[2]"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取记录ID", Val(VsfData.TextMatrix(lngRow, c文件ID)), CDate(strTime))
                If rsTemp.RecordCount <> 0 Then
                    VsfData.TextMatrix(lngRow, c记录ID) = Val(Nvl(rsTemp!Id))
                End If
            End If
        End If
    Next lngRow
    
    SaveDate = True
    
    If blnData = True Then
        '保存数据集
        Call CopyCellData
        Call Data_Save
    End If
    Exit Function
Errhand:
    Screen.MousePointer = 0
    If blnTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function CopyCellData() As Boolean
'------------------------------------------------
'功能:Copy保存后的数据
'------------------------------------------------
    Dim i As Integer
    
    '删除状态=3的数据或数值为空的数据
    mrsCell.Filter = 0
    mrsCell.Filter = "状态=3 or 数据=''"
    Do While Not mrsCell.EOF
        mrsCell.Delete
        mrsCell.Update
        mrsCell.MoveNext
    Loop
    '修改状态为0
    mrsCell.Filter = 0
    Do While Not mrsCell.EOF
        mrsCell!状态 = 0
        mrsCell.Update
        mrsCell.MoveNext
    Loop
    
    'mrsCopy中有的 mrscell没有赋值给mrscell
    mrsCopy.Filter = 0
    Do While Not mrsCopy.EOF
        mrsCell.Filter = "ID='" & Nvl(mrsCopy!Id) & "'"
        If mrsCell.RecordCount = 0 Then
            mrsCell.AddNew
            For i = 0 To mrsCopy.Fields.Count - 1
                '目前MrsCell记录集只包含 adLongVarChar 和 adDouble 两种类型
                If mrsCopy.Fields(i).Type = adLongVarChar Then
                    mrsCell.Fields(mrsCopy.Fields(i).Name).Value = Nvl(mrsCopy.Fields(i).Value)
                Else
                    mrsCell.Fields(mrsCopy.Fields(i).Name).Value = Val(Nvl(mrsCopy.Fields(i).Value))
                End If
            Next i
            mrsCell.Update
        End If
    mrsCopy.MoveNext
    Loop
    
    '删除记录集信息
    mrsCopy.Filter = 0
    Do While Not mrsCopy.EOF
        mrsCopy.Delete
        mrsCopy.Update
        mrsCopy.MoveNext
    Loop
    
    '开始拷贝数据
    mrsCell.Filter = 0
    mrsCell.Sort = "行号,ID"
    Do While Not mrsCell.EOF
        mrsCopy.AddNew
        For i = 0 To mrsCell.Fields.Count - 1
            '目前MrsCell记录集只包含 adLongVarChar 和 adDouble 两种类型
            If mrsCell.Fields(i).Type = adLongVarChar Then
                mrsCopy.Fields(mrsCell.Fields(i).Name).Value = Nvl(mrsCell.Fields(i).Value)
            Else
                mrsCopy.Fields(mrsCell.Fields(i).Name).Value = Val(Nvl(mrsCell.Fields(i).Value))
            End If
        Next i
        mrsCopy.Update
    mrsCell.MoveNext
    Loop

    '删除记录集信息
    mrsCell.Filter = 0
    Do While Not mrsCell.EOF
        mrsCell.Delete
        mrsCell.Update
        mrsCell.MoveNext
    Loop
End Function

Private Function Data_Save() As Boolean
'-------------------------------------------------------
'功能:保存数据保存的后的列信息,或刷新后的列信息,一边帅新
'------------------------------------------------------
    Dim lngRows As Long, lngStartRow As Long, lngCol As Long, lngCOls As Long
    On Error GoTo Errhand
    
    If mrsData Is Nothing Then Exit Function
    '清除内存集
    mrsData.Filter = 0
    Do While Not mrsData.EOF
        mrsData.Delete
        mrsData.Update
        mrsData.MoveNext
    Loop
    
    lngRows = VsfData.Rows - 1
    lngCOls = VsfData.Cols - 1
    
    '开始复制行数据
    For lngStartRow = VsfData.FixedRows To lngRows
        mrsData.AddNew
        mrsData("行号") = lngStartRow
        For lngCol = c文件ID To lngCOls
            mrsData.Fields(lngCol).Value = Trim(VsfData.TextMatrix(lngStartRow, lngCol))
        Next lngCol
        mrsData("删除") = IIf(VsfData.RowHidden(lngStartRow), 1, 0)
        mrsData.Update
    Next lngStartRow
    
    If mblnNullRow Then Call RefreshHistoryData(VsfData.Row)
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckDateTime(strReturn As String, ByVal strPatientTime As String, strInfo As String) As Boolean
'-----------------------------------------------------------------------------
'功能:检查录入的时点是否合法
'strPatientTime 体温单开始时间;病人出院时间
'-----------------------------------------------------------------------------
    Dim strText As String, strTime As String
    
    strText = Trim(strReturn)
    
    If Trim(strText) = "" Then
        strInfo = "时间不能为空！"
        Exit Function
    End If
    If Len(strText) <= 2 Then
        strText = String(2 - Len(strText), "0") & strText
        strText = strText & ":00"
    End If
    If Val(Mid(strText, 1, 2)) < 0 Or Val(Mid(strText, 1, 2)) > 23 Then
        strInfo = "录入的时间无效，小时应该在0-23之间！"
        Exit Function
    End If
    If Mid(strText, 3, 1) <> ":" Then
        strInfo = "录入的时间格式错误[04:00]！"
        Exit Function
    End If
    If Len(strText) < 5 Then strText = strText & String(5 - Len(strText), "0")
    If Not (Val(Mid(strText, 4, 2)) >= 0 And Val(Mid(strText, 4, 2)) <= 59) Then
        strInfo = "录入的时间无效，分钟应该在0-59之间！"
        Exit Function
    End If
    If Len(strText) > 5 Then
        strInfo = "录入的时间格式错误[04:00]！"
        Exit Function
    End If
    
    If Trim(strText) <> Trim(picInput.Tag) Then
        strTime = Format(Format(mstrDate, "YYYY-MM-DD") & " " & strText, "YYYY-MM-DD HH:mm:ss")
        '检查录入数据的时间是否超过体温单开始时间和数据补录时间
        If Not CheckTime(strTime, strPatientTime, strInfo) Then Exit Function
        '数据检测成功后检测改时间是否存在同步过来的数据信息
        If Not CheckPaseDate(strTime) Then Exit Function
    End If
    
    strReturn = strText
    CheckDateTime = True
End Function

Private Function CheckPaseDate(ByVal strTime As String) As Boolean
'------------------------------------------------------------------
'检查该店是否存在数据信息
'------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim i As Integer
    Dim lngItemNO As Long, lngRow As Long, lngCol As Long
    Dim strContent As String
    Dim strValues As String, strFileds As String
    Dim arrValues() As Variant, arrKeys() As Variant, arrID() As String
    Dim strKey As String, strKeys As String
    Dim intCOl As Integer
    Dim blnUpdate As Boolean
    Dim strPart As String, lng来源ID As Long, int共用 As Integer, int显示 As Integer, intModify As Integer, strNote As String
    Dim bln脉搏 As Boolean, blnAllow As Boolean, bln心率单独 As Boolean
    Dim intState As Integer
    
    On Error GoTo Errhand:
    
    arrValues = Array()
    arrKeys = Array()
    
    bln心率单独 = True
    bln脉搏 = False
    mrsItems.Filter = 0
    mrsItems.Filter = "项目序号=" & gint脉搏
    If mrsItems.RecordCount > 0 Then bln脉搏 = True
    
    If mrsCell Is Nothing Then Exit Function
    strFileds = "ID|行号|项目序号|数据|部位|数据来源|来源ID|共用|显示|修改|状态"
    lngRow = VsfData.Row
    
    VsfData.Cell(flexcpForeColor, lngRow, c时间, lngRow, VsfData.Cols - 1) = &H80000012
    
    blnUpdate = False
    '修改时间是 检查是否是保存的数据
    mrsCopy.Filter = 0
    mrsData.Filter = 0
    mrsCopy.Filter = "行号=" & lngRow
    If mrsCopy.RecordCount > 0 Then
        mrsData.Filter = "行号=" & lngRow
        If Format(strTime, "HH:mm") <> Format(mrsData.Fields(c时间).Value, "HH:mm") Then
            '修改mrsdata记录集删除=2 表示时间发生修改
            intState = Val(Nvl(mrsData!删除))
            If intState = 1 Then
                intState = 1
            Else
                intState = 2
            End If
            mrsData!删除 = intState
            mrsData.Update
            mblnChage = True
'            '提取这一行数据信息
'            For lngCol = RootCol To VsfData.Cols - 1
'                blnUpdate = False
'                strKey = lngRow & "," & lngCol
'
'                mrsCell.Filter = 0
'                mrsCell.Filter = "ID='" & strKey & "'"
'                If Not mrsCell.EOF Then
'                    strValues = Nvl(mrsCell!项目序号) & "|" & VsfData.TextMatrix(lngRow, lngCol) & "|" & Nvl(mrsCell!部位) & "|" & _
'                        Nvl(mrsCell!数据来源) & "|" & Nvl(mrsCell!来源ID) & "|" & Nvl(mrsCell!共用) & "|" & Nvl(mrsCell!显示) & "|" & _
'                        Nvl(mrsCell!修改) & "|" & 1
'                    blnUpdate = True
'                Else
'                    mrsCopy.Filter = "ID='" & strKey & "'"
'                    If Not mrsCopy.EOF Then
'                        strValues = Nvl(mrsCopy!项目序号) & "|" & VsfData.TextMatrix(lngRow, lngCol) & "|" & Nvl(mrsCopy!部位) & "|" & _
'                            Nvl(mrsCopy!数据来源) & "|" & Nvl(mrsCopy!来源ID) & "|" & Nvl(mrsCopy!共用) & "|" & Nvl(mrsCopy!显示) & "|" & _
'                            Nvl(mrsCopy!修改) & "|" & 1
'                         blnUpdate = True
'                    End If
'                End If
'
'                If blnUpdate = True Then
'                    ReDim Preserve arrValues(UBound(arrValues) + 1)
'                    arrValues(UBound(arrValues)) = strValues
'                    ReDim Preserve arrKeys(UBound(arrKeys) + 1)
'                    arrKeys(UBound(arrKeys)) = lngRow + 1 & "," & lngCol
'                End If
'            Next lngCol
'            '时间发生改变 为记录集mrscell打上删除标记并复制一行新的数据
'            Call Edit_Clear
'            VsfData.Row = lngRow + 1
        End If
    End If
    'lngRow = VsfData.Row
    
    '开始进行恢复操作(以保存的数据修改时间)
'    For i = 0 To UBound(arrValues)
'        lngCol = Val(Split(CStr(arrKeys(i)), ",")(1))
'        strValues = CStr(arrKeys(i)) & "|" & lngRow & "|" & CStr(arrValues(i))
'        Call Record_Update(mrsCell, strFileds, strValues, "ID|" & CStr(arrKeys(i)))
'        VsfData.TextMatrix(lngRow, lngCol) = Split(CStr(arrValues(i)), "|")(1)
'    Next i

    mrsCell.Filter = 0
    strKeys = ""
    '更换本列同步过来的数据来源
    mrsCell.Filter = "行号=" & lngRow
    With mrsCell
        Do While Not .EOF
            If InStr(1, ",0,9,", "," & Val(Nvl(mrsCell!数据来源)) & ",") = 0 Then
                strKey = Nvl(mrsCell!Id, ",")
                intCOl = Val(Split(strKey, ",")(1))
                strKeys = strKeys & "|" & strKey
            Else
                If mblnChage = False Then mblnChage = True
            End If
        .MoveNext
        Loop
    End With
    
    mrsCell.Filter = 0
    '清空数据来源记录集
    If Left(strKeys, 1) = "|" Then strKeys = Mid(strKeys, 2)
    If strKeys <> "" Then
        arrID = Split(strKeys, "|")
        For i = 0 To UBound(arrID)
            mrsCell.Filter = "ID='" & CStr(arrID(i)) & "'"
            mrsCell!数据来源 = 0
            mrsCell!状态 = 1
            mrsCell.Update
            blnUpdate = True
        Next i
    End If
    mrsCell.Filter = 0
    strKey = ""
    
    '检测该点是否存在同步过来的数据
    mstrSQL = "SELECT C.项目序号,C.记录内容,C.数据来源,C.体温部位,C.未记说明,C.来源ID,C.共用,C.显示,DECODE(C.项目序号,-1,1,C.记录标记) 记录标记" & vbNewLine & _
        " FROM 病人护理文件 A,病人护理数据 B,病人护理明细 C,护理记录项目 D" & vbNewLine & _
        " WHERE A.ID=B.文件ID AND B.ID=C.记录ID AND A.ID=[1] AND A.病人ID=[2] AND A.主页ID=[3]" & vbNewLine & _
        " AND nvl(C.来源ID,0)>0 And Mod(C.记录类型,5)<>5 AND C.终止版本 IS NULL  AND B.发生时间=[4] And C.项目序号=D.项目序号 AND nvl(D.护理等级,3) >=[5] And Nvl(D.适用病人,0) In (0,[6])" & vbNewLine & _
        " Order By B.发生时间,DECODE(C.项目序号,-1,1,0),DECODE(C.项目序号,-1,1,C.记录标记)"
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "同步数据", mlng文件ID, mlng病人ID, mlng主页ID, CDate(strTime), Val(VsfData.TextMatrix(lngRow, c护理等级)), IIf(Val(VsfData.TextMatrix(lngRow, c婴儿)) = 0, 1, 2))
    
    If rsTemp.RecordCount = 0 Then GoTo NextPos
    
    For i = RootCol To VsfData.Cols - 1
        lngItemNO = Val(VsfData.TextMatrix(0, i))
        
        If rsTemp.RecordCount > 0 Then rsTemp.MoveFirst

        strContent = ""
        strPart = ""
        strNote = ""
        lng来源ID = 0
        int共用 = 0
        int显示 = 0
        intModify = 0
        With rsTemp
            Do While Not .EOF
                If lngItemNO <> 4 Then
                    blnAllow = False
                    bln心率单独 = False
                    intModify = 0
                    
                    If InStr(1, "," & gint体温 & "," & gint脉搏 & "," & gint心率 & ",", "," & Val(Nvl(!项目序号)) & ",") > 0 Then
                        Select Case Val(Nvl(!项目序号))
                            Case gint体温
                                If gint体温 = lngItemNO Then blnAllow = True
                            Case gint脉搏
                                If gint脉搏 = lngItemNO Then blnAllow = True
                            Case gint心率
                                If bln脉搏 = True And mint心率应用 = 2 Then
                                    If gint脉搏 = lngItemNO Then blnAllow = True
                                Else
                                    If gint心率 = lngItemNO Then blnAllow = True: bln心率单独 = True
                                End If
                        End Select
                        
                        If blnAllow = True Then
                            If Val(Nvl(!记录标记)) = 0 And InStr(1, ",0,9,", "," & Val(Nvl(!数据来源)) & ",") = 0 Then
                                    
                                strContent = Nvl(!记录内容)
                                strPart = Nvl(!体温部位)
                                lng来源ID = Val(Nvl(!来源ID))
                                int共用 = Val(Nvl(!共用))
                                int显示 = Val(Nvl(!显示))
                                strNote = Nvl(!未记说明)
                            Else '组装物理降温和脉搏短轴
                                If bln心率单独 = False Then
                                    If strContent <> "" Then
                                        If InStr(1, strContent, "/") = 0 Then
                                            strContent = strContent & "/" & Nvl(!记录内容)
                                        Else
                                            strContent = Split(strContent, "/")(0) & "/" & Nvl(!记录内容)
                                        End If
                                    Else
                                        strContent = Nvl(!记录内容)
                                    End If
                                    
                                    If InStr(1, ",0,9,", "," & Val(Nvl(!数据来源)) & ",") = 0 Then
                                        intModify = 1
                                    End If
                                    
                                    Exit Do
                                Else
                                    If InStr(1, ",0,9,", "," & Val(Nvl(!数据来源)) & ",") = 0 Then
                                        strPart = Nvl(!体温部位)
                                        lng来源ID = Val(Nvl(!来源ID))
                                        int共用 = Val(Nvl(!共用))
                                        int显示 = Val(Nvl(!显示))
                                        intModify = 1
                                        strContent = Nvl(!记录内容)
                                        strNote = Nvl(!未记说明)
                                        Exit Do
                                    End If
                                End If
                            End If
                        End If
                    Else
                        If Val(Nvl(!项目序号)) = lngItemNO And InStr(1, ",0,9,", "," & Val(Nvl(!数据来源)) & ",") = 0 Then
                            strPart = Nvl(!体温部位)
                            lng来源ID = Val(Nvl(!来源ID))
                            int共用 = Val(Nvl(!共用))
                            int显示 = Val(Nvl(!显示))
                            strContent = Nvl(!记录内容)
                            strNote = Nvl(!未记说明)
                            intModify = 1
                            Exit Do
                        End If
                    End If
                ElseIf InStr(1, ",4,5,", "," & Val(!项目序号) & ",") <> 0 And lngItemNO = 4 Then
                    Select Case Val(!项目序号)
                        Case 4
                            If strContent <> "" Or Nvl(!记录内容) <> "" Then
                                If InStr(1, strContent, "/") > 0 Then
                                    strContent = Nvl(!记录内容) & "/" & Trim(Split(strContent, "/")(1))
                                Else
                                    strContent = Nvl(!记录内容) & "/"
                                End If
                                strNote = Nvl(!未记说明)
                                strPart = Nvl(!体温部位)
                                lng来源ID = Val(Nvl(!来源ID))
                                int共用 = Val(Nvl(!共用))
                                int显示 = Val(Nvl(!显示))
                                intModify = 1 '不能进行修改
                            End If
                        Case 5
                            If strContent <> "" Or Nvl(!记录内容) <> "" Then
                                If InStr(1, strContent, "/") > 0 Then
                                    strContent = Trim(Split(strContent, "/")(0)) & "/" & Nvl(!记录内容)
                                Else
                                    strContent = "/" & Nvl(!记录内容)
                                End If
                            End If
                    End Select
                End If
                .MoveNext
            Loop
            
            If strContent = "/" Then strContent = ""
            If lngItemNO = 4 Then
                If InStr(1, strContent, "/") <> 0 Then
                    If Not IsNumeric(Split(strContent, "/")(0)) And Not IsNumeric(Split(strContent, "/")(1)) Then
                        strContent = ""
                    End If
                End If
            End If
            
            '如果是体温曲线项目，并且部位不为空
            mrsItems.Filter = "项目序号=" & lngItemNO
            If mrsItems.RecordCount > 0 Then
                If Nvl(mrsItems!记录法, 2) = 1 Then
                    If strNote <> "" And strContent = "" Then
                        strContent = strNote
                    Else
                        If strContent <> "" Then strContent = IIf(strPart = "", "", strPart & ":") & strContent
                    End If
                End If
            End If

            If strContent <> "" Then
                '将同步的数据装载到记录集中
                strKey = lngRow & "," & i
                strValues = strKey & "|" & lngRow & "|" & lngItemNO & "|" & strContent & "|" & strPart & "|1|" & lng来源ID & "|" & int共用 & "|" & int显示 & "|" & intModify & "|0"
                Call Record_Update(mrsCell, strFileds, strValues, "ID|" & strKey)
                VsfData.TextMatrix(lngRow, i) = strContent
                If lngItemNO = gint体温 Or (lngItemNO = gint脉搏 And mint心率应用 = 2) Then
                    VsfData.Cell(flexcpForeColor, lngRow, i, lngRow, i) = RGB(0, 0, 255)
                Else
                    VsfData.Cell(flexcpForeColor, lngRow, i, lngRow, i) = 255 '&H8080FF
                End If
            End If
        End With
    Next i
    VsfData.Cell(flexcpAlignment, VsfData.FixedRows, c时间, VsfData.Rows - 1, VsfData.Cols - 1) = flexAlignCenterCenter
    mrsItems.Filter = 0
NextPos:
    CheckPaseDate = True
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckTime(ByVal strTime As String, ByVal strPatientTime As String, strInfo As String) As Boolean
'-------------------------------------------------------------
'功能:检查数据补录和超期录入
'strPatientTime 体温单开始日期;病人出院日期
'-------------------------------------------------------------
    Dim strInTime As String, strOutTime As String, strCurrDate As String
    
    On Error GoTo Errhand
    
    strInTime = Split(strPatientTime, ";")(0)
    strOutTime = Split(strPatientTime, ";")(1)
    
    If mbln出院 = False Then
        strOutTime = DateAdd("d", mintPreDays, CDate(strOutTime))
    End If
    
    If Format(strTime, "YYYY-MM-DD HH:mm") > Format(strOutTime, "YYYY-MM-DD HH:mm") Then
        If mbln出院 = False Then
            strInfo = "记录数据时间已超出参数[超期录入天数：" & mintPreDays & "天]所指定的范围!"
        Else
            strInfo = "记录数据时间不能大于[病人出院时间：" & Format(strOutTime, "YYYY-MM-DD HH:mm:ss") & "]!"
        End If
        Exit Function
    End If
    
    If Format(strTime, "YYYY-MM-DD HH:mm") < Format(strInTime, "YYYY-MM-DD HH:mm") Then
        strInfo = strInfo & "记录数据时间不能小于[体温单开始时间：" & Format(strInTime, "YYYY-MM-DD HH:mm:ss") & "]!"
        Exit Function
    End If
    
    strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
    If Not IsAllowInput(mlng病人ID, mlng主页ID, strTime, strCurrDate) Then
        strInfo = "记录数据时间[" & strTime & "]有误![超过数据补录的有效时限:" & mlngHours & "小时]"
        Exit Function
    End If
    
    CheckTime = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function IsAllowInput(ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal strTime As String, ByVal strCurTime As String) As Boolean
    '取出指定病人在指定时间之后关键点的时间
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo Errhand
    
    IsAllowInput = True
    gstrSQL = "" & _
              " SELECT DECODE(终止原因,1,'出院',3,'转科',10,'预出院',15,'转病区',DECODE(开始原因,10,'出院','未定义')) AS 类型,终止时间 AS 时间" & _
              " From 病人变动记录" & _
              " WHERE (终止原因 IN (1,3,10,15) OR 开始原因=10) And 病人ID=[1] And 主页ID=[2] And [3] <= 终止时间" & _
              " ORDER BY 终止时间"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取出指定病人在指定时间之后关键点的时间", lng病人ID, lng主页ID, CDate(strTime))
    If rsTemp.RecordCount = 0 Then Exit Function
    
    '只取第一条符合的记录
    strTime = Format(DateAdd("H", mlngHours, rsTemp!时间), "yyyy-MM-dd HH:mm")
    strCurTime = Format(strCurTime, "yyyy-MM-dd HH:mm")
    
    If strTime < strCurTime Then IsAllowInput = False
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckValid(strReturn As String, ByVal lngItemNO As Long, strInfo As String) As Boolean
    Dim dblMin As Double, dblMax As Double
    Dim strText As String, strText1 As String, strName As String, strGroupName As String, strFormat As String, strFormat1 As String
    Dim arrValue() As String
    Dim strValue As String
    Dim i As Integer
    Dim blnCheck As Boolean
    Dim blnAllow As Boolean
    
    strText = Trim(strReturn)
    mrsItems.Filter = 0
    mrsItems.Filter = "项目序号=" & lngItemNO
    If mrsItems.RecordCount = 0 Then Exit Function
    
    strName = mrsItems!项目名称
    strGroupName = mrsItems!分组名
    
    blnAllow = True
    If strName = "身高" Or strName = "体重" Then
        blnAllow = IsNumeric(strInfo)
    Else
        blnAllow = IIf(InStr(1, "," & gint大便 & "," & gint入液 & ",", "," & lngItemNO & ",") > 0, False, True)
    End If
    
    If strText <> "" Then
        If mrsItems!项目类型 = 0 And InStr(1, "0,4", Nvl(mrsItems!项目表示, 0)) <> 0 Then
            If lngItemNO = 4 Then
                If InStr(1, strText, "/") = 0 Then
                    strInfo = "[血压]数据格式错误：收缩压/舒张压！"
                    Exit Function
                End If
                
                '--问题号53505,修改人：李涛，血压显示文字
                If Trim(Split(strText, "/")(0)) = "拒测" Or Trim(Split(strText, "/")(0)) = "未测" Or Trim(Split(strText, "/")(0)) = "请假" Or Trim(Split(strText, "/")(0)) = "外出" Then
                     strReturn = Trim(Split(strText, "/")(0))
                     CheckValid = True
                     Exit Function
                Else
                    If Trim(Split(strText, "/")(0)) = "" Or Trim(Split(strText, "/")(1)) = "" Then
                        strInfo = "[血压]数据格式录入错误：收缩压/舒张压！"
                        Exit Function
                    End If
                End If
            ElseIf lngItemNO = gint脉搏 And mint心率应用 <> 2 Then
                If InStr(1, strText, "/") <> 0 Then
                    strInfo = "[" & strName & "]数据格式录入错误,请检查心率是否和脉搏共用！"
                    Exit Function
                End If
            ElseIf lngItemNO <> gint体温 And Not (lngItemNO = gint脉搏 And mint心率应用 = 2) And blnAllow = True Then
                If InStr(1, strText, "/") <> 0 Then
                    strInfo = "[" & strName & "]数据格式录入错误,请检查！"
                    Exit Function
                End If
            End If
            
            arrValue = Split(strText, "/")
            
            For i = 0 To UBound(arrValue)
                strText = arrValue(i)
                blnCheck = False
                
                If strGroupName = "1)体温曲线项目" Then
                    If Not IsNumeric(strText) Then
                        If InStr(1, "," & mstrNote & "," & IIf(lngItemNO = gint体温, ",不升,", ""), "," & strText & ",") <> 0 Then
                            blnCheck = True
                        Else
                            strInfo = "[" & strName & "]数据格式录入错误,请检查！"
                            Exit Function
                        End If
                    End If
                Else
                    If Not IsNumeric(strText) And blnAllow = True Then
                        strInfo = "[" & strName & "]数据格式录入错误,请检查！"
                        Exit Function
                    End If
                End If
                
                If blnCheck = True Then
                    If UBound(arrValue) > 0 Then
                        strInfo = "[" & strName & "]数据格式录入错误,请检查！"
                        Exit Function
                    End If
                End If
                
                If Nvl(mrsItems!项目小数, 0) <> 0 And blnAllow = True Then  '等于零是通过控件的MaxLength来控制的
                    If InStr(1, strText, ".") <> 0 Then strText1 = Mid(strText, 1, InStr(1, strText, ".") - 1)
                    If Len(strText1) > mrsItems!项目长度 Then
                        mrsItems.Filter = 0
                        strInfo = "[" & strName & "]录入的数据超过了合法精度！"
                        Exit Function
                    End If
        
                    If InStr(1, strText, ".") <> 0 Then
                        strText1 = Mid(strText, InStr(1, strText, ".") + 1)
                        If Len(strText1) > mrsItems!项目小数 Then
                            mrsItems.Filter = 0
                            strInfo = "[" & strName & "]录入的小数部分超过了合法精度！"
                            Exit Function
                        End If
                    End If
                End If
                If Not IsNull(mrsItems!项目值域) And Not blnCheck And blnAllow = True And Nvl(mrsItems!项目表示, 0) = 0 Then
                    dblMin = Split(mrsItems!项目值域, ";")(0)
                    dblMax = Split(mrsItems!项目值域, ";")(1)
                    If Not (Val(strText) >= dblMin And Val(strText) <= dblMax) Then
                        mrsItems.Filter = 0
                        strInfo = "[" & strName & "]录入的数据不在" & Format(dblMin, "#0.00") & "～" & Format(dblMax, "#0.00") & "的有效范围！"
                        Exit Function
                    End If
                End If
                
                If blnCheck = True Then
                    strFormat = strText
                Else
                    strFormat = strFormat & "/" & IIf(blnAllow = True, Val(strText), strText)
                End If
                
                If i = UBound(arrValue) Then
                    If Left(strFormat, 1) = "/" Then strFormat = Mid(strFormat, 2)
                End If
            Next i
        Else '文本类型
            If LenB(StrConv(strText, vbFromUnicode)) > mrsItems!项目长度 Then
                strInfo = "[" & strName & "]录入的数据超过了最大长度：" & mrsItems!项目长度 & "！"
                mrsItems.Filter = 0
                Exit Function
            End If
            strFormat = strText
        End If
    Else
    
    End If
    
    strFormat1 = strFormat
    
    '对于数据来源<>0,9的 体温,脉搏数据 进行编辑(无物理降温和脉搏短轴可以录入物理降温,脉搏短轴)
    If InStr(1, ",0,9,", "," & mint数据来源 & ",") = 0 Then
        If lngItemNO = gint体温 Or (lngItemNO = gint脉搏 And mint心率应用 = 2) Then
            strValue = picInput.Tag
            If InStr(1, strFormat1, "/") <> 0 Then
                strFormat1 = Split(strFormat1, "/")(0)
            End If
            If InStr(1, strValue, "/") = 0 Then
                If Trim(strFormat1) <> Trim(strValue) Then
                    If lngItemNO = 1 Then
                        strInfo = "同步过来的[" & strName & "]数据只允许修改物理降温部分."
                    Else
                        strInfo = "同步过来的[" & strName & "]数据只允许修改脉搏短轴部分."
                    End If
                    
                    txtInput.Text = strValue
                    Exit Function
                End If
            Else
                If mintModify = 1 Then
                    If strFormat <> strValue Then
                        If lngItemNO = 1 Then
                            strInfo = "同步过来的[" & strName & "]数据如果包括物理降温,不允许修改."
                        Else
                            strInfo = "同步过来的[" & strName & "]数据如果包括脉搏短轴,不允许修改."
                        End If
                        txtInput.Text = strValue
                        Exit Function
                    End If
                Else
                    If strFormat1 <> Split(strValue, "/")(0) Then
                        If lngItemNO = 1 Then
                            strInfo = "同步过来的[" & strName & "]数据只允许修改物理降温部分."
                        Else
                            strInfo = "同步过来的[" & strName & "]数据只允许修改脉搏短轴部分."
                        End If
                        txtInput.Text = strValue
                        Exit Function
                    End If
                End If
            End If
        End If
    End If
    
    mrsItems.Filter = 0
    strReturn = strFormat
    CheckValid = True
End Function

Private Sub AdjustRowFlag(ByRef objVsf As Object, ByVal intRow As Integer)
    '-----------------------------------------------------------------------------------------
    '功能:
    '参数:
    '-----------------------------------------------------------------------------------------
    If objVsf.FixedCols = 0 Then Exit Sub
    
    If Not (objVsf.Cell(flexcpPicture, intRow, 0) Is Nothing) Then Exit Sub
    Set objVsf.Cell(flexcpPicture, 1, 0, objVsf.Rows - 1, 0) = Nothing
    Set objVsf.Cell(flexcpPicture, intRow, 0) = ils16.ListImages(1).Picture
    
End Sub

Private Sub vsfHistory_DblClick()
    Call vsfHistory_KeyDown(vbKeyReturn, 0)
End Sub

Private Sub vsfHistory_KeyDown(KeyCode As Integer, Shift As Integer)
    '功能：复制历史数据的值
    Dim lngItemNO As Integer
    Dim strText As String, strInfo As String, strOldValue As String, strName As String
    Dim intCOl As Integer, intRow As Integer
    Dim blnValidate As Boolean, blnSave As Boolean
    Dim strFileds As String, strValues As String, strKey As String, strPart As String
    Dim int来源ID As Long, int共用 As Integer, int显示 As Integer, int修改 As Integer
    Dim intState As Integer
    Dim rsObj As New ADODB.Recordset
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    If vsfHistory.Col < RootCol Or vsfHistory.Row < vsfHistory.FixedRows Then Exit Sub
    If Val(vsfHistory.RowData(vsfHistory.Row)) = 0 Then Exit Sub
    If vsfHistory.ColHidden(vsfHistory.Col) Then Exit Sub
    If mrsHistory Is Nothing Then Exit Sub
    
    strFileds = "ID|行号|项目序号|数据|部位|数据来源|来源ID|共用|显示|修改|状态"
    VsfData.Col = vsfHistory.Col
    intCOl = VsfData.Col
    intRow = VsfData.Row
    blnValidate = True
    blnSave = True
    strOldValue = ""
    
    mrsHistory.Filter = 0
    strKey = vsfHistory.Row & "," & vsfHistory.Col
    mrsHistory.Filter = "ID='" & strKey & "'"
    If mrsHistory.RecordCount > 0 Then '过滤到说明存在数据
        strText = Nvl(mrsHistory!数据)
        strPart = Nvl(mrsHistory!部位)
    Else
        strText = vsfHistory.TextMatrix(vsfHistory.Row, vsfHistory.Col)
        strPart = ""
    End If
    strOldValue = VsfData.TextMatrix(intRow, intCOl)
    
    '检查要替换的数据是否是同步过来的
    strName = VsfData.TextMatrix(VsfData.FixedRows - 1, intCOl)
    mrsCell.Filter = 0
    strKey = intRow & "," & intCOl
    mrsCell.Filter = "ID='" & strKey & "'"
    If mrsCell.RecordCount > 0 Then '保存后mrsCell为空
        Set rsObj = mrsCell.Clone
    Else
        Set rsObj = mrsCopy.Clone
    End If
    rsObj.Filter = "ID='" & strKey & "'"
    If rsObj.RecordCount > 0 Then
        lngItemNO = Val(Nvl(rsObj!项目序号))
        mint数据来源 = Val(Nvl(rsObj!数据来源))
        mintModify = Val(Nvl(rsObj!修改))
        If InStr(1, ",0,9,", "," & Val(rsObj!数据来源) & ",") = 0 Then
            If Not (lngItemNO = gint体温 Or (lngItemNO = gint脉搏 And mint心率应用 = 2)) Then
                strInfo = "同步过来的[" & strName & "]数据不能进行修改."
                RaiseEvent AfterRowColChange(strInfo, True)
                Exit Sub
            Else
                If mintModify = 1 Then
                    If lngItemNO = gint体温 Then
                        strInfo = "同步过来的[" & strName & "]数据如果包含物理降温不能进行修改."
                    Else
                        strInfo = "同步过来的[" & strName & "]数据如果包含脉搏短轴不能进行修改."
                    End If
                    RaiseEvent AfterRowColChange(strInfo, True)
                    Exit Sub
                Else
                    If InStr(1, strOldValue, "/") <> 0 Then
                        strInfo = Split(strOldValue, "/")(0)
                    Else
                        strInfo = strOldValue
                    End If
                    strInfo = Mid(strInfo, InStr(1, strInfo, ":") + 1)
                    If strInfo <> "" Then
                        If InStr(1, strText, "/") = 0 Then
                            strInfo = "同步过来的[" & strName & "]数据不能进行修改."
                            RaiseEvent AfterRowColChange(strInfo, True)
                            Exit Sub
                        Else
                            strText = strInfo & "/" & Split(strText, "/")(1)
                        End If
                    End If
                End If
            End If
        End If
    End If
    
    lngItemNO = Val(vsfHistory.TextMatrix(0, intCOl))
    If blnValidate = True Then
        mrsCopy.Filter = 0
        VsfData.TextMatrix(intRow, intCOl) = IIf(strPart = "", "", strPart & ":") & strText
        VsfData.Cell(flexcpAlignment, intRow, intCOl, intRow, intCOl) = flexAlignCenterCenter
        '进行数据处理
        If blnSave = True Then
            If Trim(strOldValue) <> Trim(strText) Then
                strKey = intRow & "," & intCOl
                '检查修改的数据是否已经保存
                mrsCopy.Filter = "ID='" & strKey & "'"
                If mrsCopy.RecordCount > 0 Then
                    int来源ID = Val(Nvl(mrsCopy!来源ID))
                    int共用 = Val(Nvl(mrsCopy!共用))
                    int显示 = Val(Nvl(mrsCopy!显示))
                    int修改 = Val(Nvl(mrsCopy!修改))
                    intState = 1
                Else
                    int来源ID = 0: int共用 = 0: int显示 = 0: int修改 = 0
                    intState = IIf(Trim(strText) = "", 3, 1)
                End If
                strValues = strKey & "|" & intRow & "|" & lngItemNO & "|" & strText & "|" & strPart & "|" & mint数据来源 & "|" & _
                    int来源ID & "|" & int共用 & "|" & int显示 & "|" & int修改 & "|" & intState
                Call Record_Update(mrsCell, strFileds, strValues, "ID|" & strKey)
                mblnChage = True
            End If
        End If
    End If
    
    If mblnShow And mintType <> 0 Then Call ShowInput
    mblnShow = False: mintType = 0
    
    mblnClearRow = False
    For intCOl = c时间 + 1 To VsfData.Cols - 1
        If Trim(VsfData.TextMatrix(intRow, intCOl)) <> "" Then
            mblnClearRow = True
            Exit For
        End If
    Next intCOl
End Sub

Private Function is大便或入液(ByVal intType As Integer) As Boolean
'检查是否是大便项目或入夜项目  大便项目序号=10 入夜=9
'intType=1 为大便项目 否则为入液项目
    Dim lngItemNO As Long
    Dim strKey As String
    Dim rsObj As New ADODB.Recordset
    
    On Error GoTo Errhand
    
    If VsfData.Col < RootCol Or VsfData.Row < VsfData.FixedRows Then Exit Function
    If mblnInit = False Or mblnNullRow = False Then Exit Function
    
    '提取项目序号
    lngItemNO = Val(VsfData.TextMatrix(0, VsfData.Col))
    If intType = 1 Then
        If lngItemNO <> 10 Then Exit Function
    Else
        If lngItemNO <> 9 Then Exit Function
    End If
    
    mrsItems.Filter = "项目序号=" & lngItemNO
    If InStr(1, ",2,3,5,", "," & Val(Nvl(mrsItems!项目表示)) & ",") > 0 Then Exit Function
    
    '检查是否是同步的数据
    mrsCell.Filter = 0
    strKey = VsfData.Row & "," & VsfData.Col
    mrsCell.Filter = "ID='" & strKey & "'"
    If mrsCell.RecordCount > 0 Then '保存后mrsCell为空
        Set rsObj = mrsCell.Clone
    Else
        Set rsObj = mrsCopy.Clone
    End If
    rsObj.Filter = "ID='" & strKey & "'"
    
    If rsObj.RecordCount > 0 Then
        If InStr(1, ",0,9,", "," & Val(rsObj!数据来源) & ",") = 0 Then Exit Function
    End If
    
    is大便或入液 = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function Get科室性质(ByVal lng科室ID As Long)
    Dim rsTemp As ADODB.Recordset
    gstrSQL = "select 工作性质 from 部门性质说明 where 部门ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "部门性质", lng科室ID)
    If Not rsTemp.BOF Then
        mstr科室性质 = rsTemp!工作性质
    Else
        mstr科室性质 = "所有"
    End If
End Function


