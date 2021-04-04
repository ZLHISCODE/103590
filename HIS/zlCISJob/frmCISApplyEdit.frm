VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCISApplyEdit 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "新增访问申请"
   ClientHeight    =   8490
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7320
   Icon            =   "frmCISApplyEdit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8490
   ScaleWidth      =   7320
   StartUpPosition =   1  '所有者中心
   Begin MSComctlLib.ImageList img16 
      Left            =   2040
      Top             =   7920
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISApplyEdit.frx":6852
            Key             =   "girl"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISApplyEdit.frx":D0B4
            Key             =   "boy"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISApplyEdit.frx":13916
            Key             =   "访问时限"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISApplyEdit.frx":13EB0
            Key             =   "访问内容"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISApplyEdit.frx":1444A
            Key             =   "访问医生"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISApplyEdit.frx":149E4
            Key             =   "访问病人"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISApplyEdit.frx":14F7E
            Key             =   "AllCheck"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISApplyEdit.frx":150D8
            Key             =   "unCheck"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picParent 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7575
      Left            =   0
      ScaleHeight     =   7575
      ScaleWidth      =   7305
      TabIndex        =   52
      Top             =   360
      Width           =   7300
      Begin VB.Frame fraPatiType 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "设置申请访问的指定病人"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   7335
         Left            =   120
         TabIndex        =   53
         Top             =   120
         Width           =   7095
         Begin VB.PictureBox picPati 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   6975
            Left            =   120
            ScaleHeight     =   6975
            ScaleWidth      =   6900
            TabIndex        =   54
            Top             =   240
            Width           =   6900
            Begin VB.TextBox txtFind 
               Appearance      =   0  'Flat
               Height          =   270
               Left            =   4680
               TabIndex        =   2
               Top             =   120
               Width           =   1815
            End
            Begin VB.PictureBox picTmp 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   1
               Left            =   3495
               ScaleHeight     =   240
               ScaleWidth      =   1140
               TabIndex        =   56
               Top             =   120
               Width           =   1170
               Begin VB.ComboBox cboFind 
                  Appearance      =   0  'Flat
                  Height          =   300
                  Left            =   -30
                  Style           =   2  'Dropdown List
                  TabIndex        =   1
                  Top             =   -30
                  Width           =   1215
               End
            End
            Begin VB.PictureBox picTmp 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               ForeColor       =   &H80000008&
               Height          =   270
               Index           =   2
               Left            =   1065
               ScaleHeight     =   240
               ScaleWidth      =   1905
               TabIndex        =   55
               Top             =   130
               Width           =   1935
               Begin VB.ComboBox cboDept 
                  Height          =   300
                  Left            =   -30
                  Style           =   2  'Dropdown List
                  TabIndex        =   0
                  Top             =   -30
                  Width           =   1960
               End
            End
            Begin VB.CommandButton cmdAdd 
               Height          =   315
               Left            =   6550
               Picture         =   "frmCISApplyEdit.frx":15232
               Style           =   1  'Graphical
               TabIndex        =   4
               Top             =   480
               Width           =   330
            End
            Begin VB.CommandButton cmdDel 
               Height          =   315
               Left            =   6550
               Picture         =   "frmCISApplyEdit.frx":1BA84
               Style           =   1  'Graphical
               TabIndex        =   5
               Top             =   960
               Width           =   330
            End
            Begin VSFlex8Ctl.VSFlexGrid vsPati 
               Height          =   6435
               Left            =   0
               TabIndex        =   3
               Top             =   480
               Width           =   6525
               _cx             =   11509
               _cy             =   11351
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
               MouseIcon       =   "frmCISApplyEdit.frx":222D6
               BackColor       =   -2147483643
               ForeColor       =   -2147483640
               BackColorFixed  =   -2147483633
               ForeColorFixed  =   -2147483630
               BackColorSel    =   16444122
               ForeColorSel    =   -2147483640
               BackColorBkg    =   -2147483643
               BackColorAlternate=   -2147483643
               GridColor       =   16777215
               GridColorFixed  =   16777215
               TreeColor       =   -2147483632
               FloodColor      =   192
               SheetBorder     =   0
               FocusRect       =   0
               HighLight       =   1
               AllowSelection  =   0   'False
               AllowBigSelection=   0   'False
               AllowUserResizing=   1
               SelectionMode   =   1
               GridLines       =   1
               GridLinesFixed  =   1
               GridLineWidth   =   1
               Rows            =   2
               Cols            =   7
               FixedRows       =   1
               FixedCols       =   0
               RowHeightMin    =   250
               RowHeightMax    =   2000
               ColWidthMin     =   0
               ColWidthMax     =   10000
               ExtendLastCol   =   0   'False
               FormatString    =   $"frmCISApplyEdit.frx":22BB0
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
               AllowUserFreezing=   1
               BackColorFrozen =   0
               ForeColorFrozen =   0
               WallPaperAlignment=   9
               AccessibleName  =   ""
               AccessibleDescription=   ""
               AccessibleValue =   ""
               AccessibleRole  =   24
               Begin VB.PictureBox picTmp 
                  Appearance      =   0  'Flat
                  AutoRedraw      =   -1  'True
                  BackColor       =   &H80000005&
                  BorderStyle     =   0  'None
                  ForeColor       =   &H80000008&
                  Height          =   240
                  Index           =   0
                  Left            =   1920
                  ScaleHeight     =   240
                  ScaleWidth      =   480
                  TabIndex        =   57
                  Top             =   1680
                  Visible         =   0   'False
                  Width           =   480
               End
            End
            Begin VB.Label lblDept 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "↓病人科室"
               ForeColor       =   &H80000008&
               Height          =   180
               Left            =   80
               TabIndex        =   58
               Top             =   160
               Width           =   900
            End
            Begin VB.Image imgSentence 
               Appearance      =   0  'Flat
               Height          =   360
               Left            =   2970
               Picture         =   "frmCISApplyEdit.frx":22C4B
               ToolTipText     =   "显示当前选择科室最近的病人"
               Top             =   90
               Width           =   360
            End
         End
      End
   End
   Begin VB.PictureBox picAppInfo 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7575
      Left            =   0
      ScaleHeight     =   7575
      ScaleWidth      =   7305
      TabIndex        =   40
      Top             =   360
      Width           =   7300
      Begin VB.Frame fraInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "详细内容"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   6375
         Left            =   240
         TabIndex        =   41
         Top             =   960
         Width           =   6855
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "体检报告"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   7
            Left            =   1440
            TabIndex        =   11
            Top             =   1200
            Width           =   1095
         End
         Begin VB.Frame fraFile 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1000
            Index           =   3
            Left            =   360
            TabIndex        =   47
            Top             =   3120
            Width           =   6135
            Begin VB.CheckBox chkHlInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "所有的护理记录"
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   0
               Left            =   0
               TabIndex        =   18
               Top             =   0
               Width           =   1575
            End
            Begin VB.CheckBox chkHlInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "体温单"
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   1
               Left            =   0
               TabIndex        =   19
               Top             =   350
               Width           =   855
            End
            Begin VB.CheckBox chkHlInfo 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "指定的护理记录"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   2
               Left            =   0
               TabIndex        =   20
               Top             =   720
               Width           =   1575
            End
            Begin VB.TextBox txtHlInfo 
               Appearance      =   0  'Flat
               Height          =   270
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   21
               Top             =   720
               Width           =   4095
            End
            Begin VB.Image imgHlInfo 
               Appearance      =   0  'Flat
               Height          =   360
               Left            =   5760
               Picture         =   "frmCISApplyEdit.frx":23335
               ToolTipText     =   "选择本科室最近的病人"
               Top             =   670
               Width           =   360
            End
         End
         Begin VB.Frame fraFile 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   700
            Index           =   2
            Left            =   360
            TabIndex        =   46
            Top             =   5640
            Width           =   6255
            Begin VB.OptionButton optJybg 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "所有检验报告"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   0
               TabIndex        =   27
               Top             =   0
               Width           =   1455
            End
            Begin VB.OptionButton optJybg 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "指定的检验报告"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   0
               TabIndex        =   28
               Top             =   360
               Width           =   1575
            End
            Begin VB.TextBox txtJybgTpye 
               Appearance      =   0  'Flat
               Height          =   270
               Left            =   1680
               Locked          =   -1  'True
               TabIndex        =   29
               Top             =   360
               Width           =   4095
            End
            Begin VB.Image imgJybgTpye 
               Appearance      =   0  'Flat
               Height          =   360
               Left            =   5760
               Picture         =   "frmCISApplyEdit.frx":23A1F
               ToolTipText     =   "选择本科室最近的病人"
               Top             =   310
               Width           =   360
            End
         End
         Begin VB.Frame fraFile 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   855
            Index           =   1
            Left            =   240
            TabIndex        =   45
            Top             =   4440
            Width           =   6255
            Begin VB.OptionButton optJcbg 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "所有检查报告"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   23
               Top             =   120
               Width           =   1455
            End
            Begin VB.OptionButton optJcbg 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "指定的检查报告"
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   1
               Left            =   120
               TabIndex        =   24
               Top             =   480
               Width           =   1575
            End
            Begin VB.TextBox txtJcbgTpye 
               Appearance      =   0  'Flat
               Height          =   270
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   25
               Top             =   525
               Width           =   4095
            End
            Begin VB.Image imgJcbgTpye 
               Appearance      =   0  'Flat
               Height          =   360
               Left            =   5880
               Picture         =   "frmCISApplyEdit.frx":24109
               ToolTipText     =   "选择本科室最近的病人"
               Top             =   480
               Width           =   360
            End
         End
         Begin VB.Frame fraFile 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1215
            Index           =   0
            Left            =   240
            TabIndex        =   44
            Top             =   1560
            Width           =   6375
            Begin VB.OptionButton optDzbl 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "所有电子病历"
               ForeColor       =   &H00000000&
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   12
               Top             =   120
               Width           =   1455
            End
            Begin VB.OptionButton optDzbl 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "指定类型的病历"
               ForeColor       =   &H80000008&
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   13
               Top             =   480
               Width           =   1575
            End
            Begin VB.OptionButton optDzbl 
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               Caption         =   "指定的病历文件"
               ForeColor       =   &H80000008&
               Height          =   375
               Index           =   2
               Left            =   120
               TabIndex        =   15
               Top             =   840
               Width           =   1575
            End
            Begin VB.TextBox txtDzblTpye 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   270
               Index           =   0
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   14
               Top             =   480
               Width           =   4095
            End
            Begin VB.TextBox txtDzblTpye 
               Appearance      =   0  'Flat
               Height          =   270
               Index           =   1
               Left            =   1800
               Locked          =   -1  'True
               TabIndex        =   16
               Top             =   885
               Width           =   4095
            End
            Begin VB.Image imgDzblTpye 
               Appearance      =   0  'Flat
               Height          =   360
               Index           =   0
               Left            =   5880
               Picture         =   "frmCISApplyEdit.frx":247F3
               ToolTipText     =   "选择本科室最近的病人"
               Top             =   435
               Width           =   360
            End
            Begin VB.Image imgDzblTpye 
               Appearance      =   0  'Flat
               Height          =   360
               Index           =   1
               Left            =   5880
               Picture         =   "frmCISApplyEdit.frx":24EDD
               ToolTipText     =   "选择本科室最近的病人"
               Top             =   840
               Width           =   360
            End
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "临床路径"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   6
            Left            =   1440
            TabIndex        =   9
            Top             =   720
            Width           =   1095
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "检验报告"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   5
            Left            =   120
            TabIndex        =   26
            Top             =   5280
            Width           =   1095
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "检查报告"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   4
            Left            =   120
            TabIndex        =   22
            Top             =   4080
            Width           =   1095
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "病案首页"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   7
            Top             =   240
            Width           =   1095
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "医嘱清单"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   120
            TabIndex        =   8
            Top             =   720
            Width           =   1095
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "电子病历"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   10
            Top             =   1200
            Width           =   1095
         End
         Begin VB.CheckBox chkInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "护理记录"
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   3
            Left            =   120
            TabIndex        =   17
            Top             =   2760
            Width           =   1095
         End
      End
      Begin VB.CheckBox chkAllInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "所有内容"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   360
         TabIndex        =   6
         Top             =   240
         Value           =   1  'Checked
         Width           =   1335
      End
      Begin VB.Frame fraTmp 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "所有内容"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   6375
         Left            =   240
         TabIndex        =   42
         Top             =   960
         Width           =   6735
         Begin VB.PictureBox picTmp 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1410
            Index           =   5
            Left            =   1680
            Picture         =   "frmCISApplyEdit.frx":255C7
            ScaleHeight     =   1410
            ScaleWidth      =   3435
            TabIndex        =   43
            Top             =   2280
            Width           =   3435
         End
      End
      Begin VB.Line lineTmp 
         BorderColor     =   &H80000000&
         BorderWidth     =   2
         Index           =   0
         X1              =   360
         X2              =   6840
         Y1              =   720
         Y2              =   720
      End
   End
   Begin VB.PictureBox picTime 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7575
      Left            =   0
      ScaleHeight     =   7575
      ScaleWidth      =   7305
      TabIndex        =   48
      Top             =   360
      Width           =   7300
      Begin VB.Frame fraReault 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "访问原因"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   1095
         Left            =   120
         TabIndex        =   59
         Top             =   6360
         Width           =   7095
         Begin VB.TextBox txtReault 
            Appearance      =   0  'Flat
            Height          =   735
            Left            =   120
            MaxLength       =   100
            MultiLine       =   -1  'True
            TabIndex        =   35
            Top             =   240
            Width           =   6855
         End
      End
      Begin VB.Frame fraTime 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "设置访问时限"
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   6135
         Left            =   120
         TabIndex        =   49
         Top             =   120
         Width           =   7095
         Begin VB.OptionButton optTimeTpye 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "已归档的病历(门诊已诊或历史住院病历)"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   1320
            TabIndex        =   34
            Top             =   2520
            Width           =   4000
         End
         Begin VB.OptionButton optTimeTpye 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "未归档的病历(门诊就诊或当前在院)"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   33
            Top             =   2160
            Width           =   4000
         End
         Begin VB.OptionButton optTimeTpye 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "不限制"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   1320
            TabIndex        =   32
            Top             =   1800
            Value           =   -1  'True
            Width           =   855
         End
         Begin MSComCtl2.DTPicker dtpTime 
            Height          =   300
            Index           =   0
            Left            =   1635
            TabIndex        =   30
            Top             =   480
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   214892547
            CurrentDate     =   40976
         End
         Begin MSComCtl2.DTPicker dtpTime 
            Height          =   300
            Index           =   1
            Left            =   3960
            TabIndex        =   31
            Top             =   480
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   214892547
            CurrentDate     =   40976
         End
         Begin VB.Line Line 
            BorderColor     =   &H80000000&
            BorderWidth     =   3
            X1              =   3600
            X2              =   3880
            Y1              =   600
            Y2              =   600
         End
         Begin VB.Label lblTmp 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "访问内容的时间限制："
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   840
            TabIndex        =   51
            Top             =   1320
            Width           =   1800
         End
         Begin VB.Label lbltime 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "访问时段"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   840
            TabIndex        =   50
            Top             =   503
            Width           =   735
         End
         Begin VB.Line lineTmp 
            BorderColor     =   &H80000000&
            BorderWidth     =   2
            Index           =   1
            X1              =   240
            X2              =   6720
            Y1              =   960
            Y2              =   960
         End
      End
   End
   Begin VB.CommandButton cmdQuit 
      Cancel          =   -1  'True
      Caption         =   "取消(&Q)"
      Height          =   375
      Left            =   5880
      TabIndex        =   38
      Top             =   8040
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Appearance      =   0  'Flat
      Caption         =   "确定(&O)"
      Height          =   375
      Left            =   4320
      TabIndex        =   37
      Top             =   8040
      Width           =   1215
   End
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   7980
      Left            =   0
      TabIndex        =   36
      Top             =   0
      Width           =   7330
      _Version        =   589884
      _ExtentX        =   12929
      _ExtentY        =   14076
      _StockProps     =   64
   End
   Begin VB.Image imtmp 
      Height          =   360
      Left            =   120
      Picture         =   "frmCISApplyEdit.frx":273DF
      Stretch         =   -1  'True
      Top             =   8040
      Width           =   360
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin VB.Label lblTmp 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      Caption         =   "新增访问申请"
      ForeColor       =   &H8000000A&
      Height          =   180
      Index           =   0
      Left            =   480
      TabIndex        =   39
      Top             =   8160
      Width           =   1080
   End
End
Attribute VB_Name = "frmCISApplyEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mintType As Integer '=0时为新增申请,=1时为修改申请
Private mrsPati As ADODB.Recordset
Private mblnOk As Boolean
Private mlngApplyID As Long

Private mstrNewEMR As String


Private Enum colList
    col_病人Id = 0
    col_姓名 = 1
    col_性别 = 2
    col_年龄 = 3
    COL_标识号 = 4
    col_科室 = 5
    COL_当前状态 = 6
End Enum

Private Enum FileIndex
    File_首页 = 0
    File_医嘱 = 1
    File_病历 = 2
    File_护理 = 3
    File_检查 = 4
    File_检验 = 5
    File_路径 = 6
    File_体检 = 7
End Enum

Private Enum CmdIndex
    Cmd_所有科室 = 1
    Cmd_门诊科室 = 2
    Cmd_住院科室 = 3
End Enum

Public Function ShowEdit(frmParent As Object, ByVal intType As String, ByRef lngApplyID As Long, Optional ByRef rsPati As ADODB.Recordset) As Boolean
'功能：访问申请内容编辑器
'rsPati 申请界面记录集
    On Error Resume Next
    mintType = intType
    Set mrsPati = rsPati
    mblnOk = False
    mlngApplyID = lngApplyID
    
    If mlngApplyID = 0 And mintType = 1 Then Exit Function
    Me.Show 1, frmParent
    lngApplyID = mlngApplyID
    ShowEdit = mblnOk
    On Error GoTo 0
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    lblDept.Tag = Control.Parameter
    lblDept.Caption = decode(lblDept.Tag, "", "↓病人科室", "门诊", "↓门诊科室", "住院", "↓住院科室")
    Call LoadDept
    
    '执行结果下拉菜单初始化
    cboFind.Clear
    cboFind.AddItem "姓名"
    cboFind.AddItem "身份证号"
    cboFind.AddItem "病人ID"
    If lblDept.Tag = "" Or lblDept.Tag = "门诊" Then
        cboFind.AddItem "门诊号"
    End If
    
    If lblDept.Tag = "" Or lblDept.Tag = "住院" Then
        cboFind.AddItem "住院号"
    End If
    
    cboFind.ListIndex = 0
    
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Control.Checked = Control.Parameter = lblDept.Tag
End Sub

Private Sub chkAllInfo_Click()
    Dim i As Long
    If chkAllInfo.Tag = "1" Then Exit Sub
    fraInfo.Visible = Not (chkAllInfo.Value = 1)
    fraTmp.Visible = chkAllInfo.Value = 1
    
    If chkAllInfo.Value = 0 Then
        For i = 0 To 7
            chkInfo(i).Value = 1
        Next
        optDzbl(0).Value = True
        chkHlInfo(0).Value = 1
        optJcbg(0).Value = True
        optJybg(0).Value = True
        Call SetFileCtl
    End If
End Sub

Private Sub chkHlInfo_Click(Index As Integer)
    If chkHlInfo(0).Tag = "1" Then Exit Sub
    chkHlInfo(0).Tag = "1"
    If Index = 0 Then
        If chkHlInfo(0).Value = 1 Then chkHlInfo(1).Value = 0: chkHlInfo(2).Value = 0
    Else
        If chkHlInfo(1).Value = 1 Or chkHlInfo(2).Value Then chkHlInfo(0).Value = 0
    End If
    chkHlInfo(0).Tag = ""
    Call SetFileCtl
End Sub

Private Sub chkInfo_Click(Index As Integer)
    Call SetFileCtl
End Sub


Private Sub SetFileCtl()
    fraInfo.Visible = Not (chkAllInfo.Value = 1)
    fraTmp.Visible = chkAllInfo.Value = 1

    'File_病历
    fraFile(0).Enabled = chkInfo(File_病历).Value = 1
    optDzbl(0).ForeColor = IIf(chkInfo(File_病历).Value = 1, &H0, &H80000000)
    optDzbl(1).ForeColor = IIf(chkInfo(File_病历).Value = 1, &H0, &H80000000)
    optDzbl(2).ForeColor = IIf(chkInfo(File_病历).Value = 1, &H0, &H80000000)
    txtDzblTpye(0).ForeColor = IIf(chkInfo(File_病历).Value = 1, &H0, &H80000000)
    txtDzblTpye(1).ForeColor = IIf(chkInfo(File_病历).Value = 1, &H0, &H80000000)
    
    txtDzblTpye(0).BackColor = IIf(optDzbl(1).Value = True And chkInfo(File_病历).Value = 1, &HFFFFFF, &H80000004)
    txtDzblTpye(1).BackColor = IIf(optDzbl(2).Value = True And chkInfo(File_病历).Value = 1, &HFFFFFF, &H80000004)
    
     'File_护理
    fraFile(3).Enabled = chkInfo(File_护理).Value = 1
    chkHlInfo(0).ForeColor = IIf(chkInfo(File_护理).Value = 1, &H0, &H80000000)
    chkHlInfo(1).ForeColor = IIf(chkInfo(File_护理).Value = 1, &H0, &H80000000)
    chkHlInfo(2).ForeColor = IIf(chkInfo(File_护理).Value = 1, &H0, &H80000000)
    txtHlInfo.ForeColor = IIf(chkInfo(File_护理).Value = 1, &H0, &H80000000)
    
    txtHlInfo.BackColor = IIf(chkHlInfo(2).Value = 1 And chkInfo(File_护理).Value = 1, &HFFFFFF, &H80000004)
    
    'File_检查
    fraFile(1).Enabled = chkInfo(File_检查).Value = 1
    optJcbg(0).ForeColor = IIf(chkInfo(File_检查).Value = 1, &H0, &H80000000)
    optJcbg(1).ForeColor = IIf(chkInfo(File_检查).Value = 1, &H0, &H80000000)
    txtJcbgTpye.ForeColor = IIf(chkInfo(File_检查).Value = 1, &H0, &H80000000)
    
    txtJcbgTpye.BackColor = IIf(optJcbg(1).Value = True And chkInfo(File_检查).Value = 1, &HFFFFFF, &H80000004)
    
    'File_检验
    fraFile(2).Enabled = chkInfo(File_检验).Value = 1
    optJybg(0).ForeColor = IIf(chkInfo(File_检验).Value = 1, &H0, &H80000000)
    optJybg(1).ForeColor = IIf(chkInfo(File_检验).Value = 1, &H0, &H80000000)
    txtJybgTpye.ForeColor = IIf(chkInfo(File_检验).Value = 1, &H0, &H80000000)
    
    txtJybgTpye.BackColor = IIf(optJybg(1).Value = True And chkInfo(File_检验).Value = 1, &HFFFFFF, &H80000004)
    
    '初始化
    If optDzbl(0).Value = False And optDzbl(1).Value = False And optDzbl(2).Value = False Then optDzbl(0).Value = True
    If chkHlInfo(0).Value = 0 And chkHlInfo(1).Value = 0 And chkHlInfo(2).Value = 0 Then chkHlInfo(0).Value = 1
    If optJcbg(0).Value = False And optJcbg(1).Value = False Then optJcbg(0).Value = True
    If optJybg(0).Value = False And optJybg(1).Value = False Then optJybg(0).Value = True
End Sub


Private Sub cmdAdd_Click()
    If Val(vsPati.TextMatrix(vsPati.Rows - 1, col_病人Id)) <> 0 Or vsPati.Rows < 2 Then
        vsPati.Rows = vsPati.Rows + 1
    End If
    vsPati.Row = vsPati.Rows - 1
    vsPati.SetFocus
End Sub


Private Sub cmdDel_Click()
    If vsPati.Row < 1 Then Exit Sub
    If Val(vsPati.TextMatrix(vsPati.Row, col_病人Id)) <> 0 Then
        vsPati.Tag = Replace(vsPati.Tag, Val(vsPati.TextMatrix(vsPati.Row, col_病人Id)), "")
    End If
    vsPati.RemoveItem vsPati.Row
    If vsPati.Rows < 2 Then
        Call cmdAdd_Click
    End If
    vsPati.SetFocus
End Sub

Private Sub cmdOK_Click()
    Dim str病人ids As String
    Dim str病人姓名 As String
    Dim strXML As String
    Dim lngID As Long
    Dim arrSQL As Variant
    Dim strSQL As String
    Dim i As Long
    Dim curDate As Date
    Dim blnTran As Boolean
    Dim lngTmp As Long
    
    
    On Error GoTo errH
    '获取访问病人
    For i = 1 To vsPati.Rows - 1
        If Val(vsPati.TextMatrix(i, col_病人Id)) <> 0 Then
            str病人ids = str病人ids & "," & Val(vsPati.TextMatrix(i, col_病人Id))
            str病人姓名 = str病人姓名 & "," & Val(vsPati.TextMatrix(i, col_姓名))
        End If
    Next
    str病人ids = Mid(str病人ids, 2)
    str病人姓名 = Mid(str病人姓名, 2)
    
    If str病人ids = "" Then
        Me.tbcSub.Item(0).Selected = True
        MsgBox "当前尚未录入需要申请访问病历的病人信息,请重新录入。", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If chkAllInfo.Value = 0 Then
        For i = 0 To 7
            If chkInfo(i).Value = 1 Then
                lngTmp = lngTmp + 1
            End If
        Next
        If lngTmp = 0 Then
            Me.tbcSub.Item(1).Selected = True
            MsgBox "当前尚未录入需要申请访问病历的权限内容,请重新录入。", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    '检查访问内容
    For i = 0 To 1
        If txtDzblTpye(i).BackColor = &HFFFFFF And txtDzblTpye(i).Text = "" And chkAllInfo.Value = 0 Then
            Me.tbcSub.Item(1).Selected = True
            MsgBox "当前尚未录入病历文件" & IIf(i = 0, "种类", "") & ",请重新录入!!!", vbInformation, gstrSysName
            txtDzblTpye(i).SetFocus
            Exit Sub
        End If
    Next
    If txtHlInfo.BackColor = &HFFFFFF And txtHlInfo.Text = "" And chkAllInfo.Value = 0 Then
        Me.tbcSub.Item(1).Selected = True
        MsgBox "当前尚未录入护理记录文件,请重新录入。", vbInformation, gstrSysName
        txtHlInfo.SetFocus
        Exit Sub
    End If
    If txtJcbgTpye.BackColor = &HFFFFFF And txtJcbgTpye.Text = "" And chkAllInfo.Value = 0 Then
        Me.tbcSub.Item(1).Selected = True
        MsgBox "当前尚未录入检查报告类型,请重新录入。", vbInformation, gstrSysName
        txtJcbgTpye.SetFocus
        Exit Sub
    End If
    If txtJybgTpye.BackColor = &HFFFFFF And txtJybgTpye.Text = "" And chkAllInfo.Value = 0 Then
        Me.tbcSub.Item(1).Selected = True
        MsgBox "当前尚未录入检验报告类型,请重新录入。", vbInformation, gstrSysName
        txtJybgTpye.SetFocus
        Exit Sub
    End If
    
    '检查访问原因
    If txtReault.Text = "" Then
        Me.tbcSub.Item(2).Selected = True
        MsgBox "当前尚未录入访问原因,请重新录入。", vbInformation, gstrSysName
        txtReault.SetFocus
        Exit Sub
    End If
    
    If ZLCommFun.ActualLen(txtReault.Text) > txtReault.MaxLength Then
        Me.tbcSub.Item(2).Selected = True
        MsgBox "访问原因内容过多，最多允许 " & txtReault.MaxLength \ 2 & " 个汉字或 " & txtReault.MaxLength & " 个字符。", vbInformation, gstrSysName
        txtReault.SetFocus: Exit Sub
    End If
    
    
    '检查访问时间
    If dtpTime(0).Value >= dtpTime(1).Value Then
        Me.tbcSub.Item(2).Selected = True
        MsgBox "当前访问起始时间必须小于终止时间,请重新录入。", vbInformation, gstrSysName
        txtReault.SetFocus
        Exit Sub
    End If
    
    strXML = GetInfoXml
    
    '保存数据
    lngID = mlngApplyID
    If lngID = 0 Then lngID = zlDatabase.GetNextId("电子病历访问申请")
    curDate = zlDatabase.Currentdate
    strSQL = "Zl_电子病历访问申请_Update(" & mintType & "," & lngID & ",'" & strXML & "',To_Date('" & Format(dtpTime(0).Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                            "To_Date('" & Format(dtpTime(1).Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                            IIf(optTimeTpye(0).Value, 0, IIf(optTimeTpye(1).Value, 1, 2)) & ",'" & Replace(txtReault.Text, "'", "") & _
                            "','" & UserInfo.姓名 & "',To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
    arrSQL = Array()
    For i = 1 To vsPati.Rows - 1
        If Val(vsPati.TextMatrix(i, col_病人Id)) <> 0 Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "Zl_电子病历申请访问病人_Insert(" & lngID & "," & Val(vsPati.TextMatrix(i, col_病人Id)) & ")"
        End If
    Next
    
    Screen.MousePointer = 11
    gcnOracle.BeginTrans: blnTran = True
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    For i = 0 To UBound(arrSQL)
        Debug.Print CStr(arrSQL(i))
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    gcnOracle.CommitTrans: blnTran = False
    mlngApplyID = lngID
    
    '保存xml数据
'    Call Sys.SaveLobV2("电子病历访问申请", "访问申请", "ID =[1]", "", mlngApplyID)
    
    mblnOk = True
    On Error GoTo 0
    Screen.MousePointer = 0
    Unload Me
    Exit Sub
errH:
    If blnTran Then gcnOracle.RollbackTrans: blnTran = False
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
End Sub


Private Function ReadXmlSet() As Boolean
    '获取申请内容的Xml并解析
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim strErr As String
    Dim strValue As String

    
    On Error GoTo errH
    If mlngApplyID = 0 Then Exit Function
    
    strXML = Sys.ReadXML("电子病历访问申请", "访问内容", "ID=[1]", strErr, mlngApplyID)
    If Err.Number = 0 And strErr <> "" Then
        MsgBox strErr, vbInformation, gstrSysName
        Exit Function
    End If
    
    If objXML.OpenXMLDocument(strXML) = False Then Exit Function
    
    '所有内容
    strValue = "": Call objXML.GetSingleNodeValue("all_files", strValue, xsNumber)
    chkAllInfo.Tag = "1"
    chkAllInfo.Value = Val(strValue)
    chkAllInfo.Tag = ""
    If Val(strValue) = 0 Then
        '病案首页、医嘱、临床路径、病人体检
        strValue = "": Call objXML.GetSingleNodeValue("medical_record", strValue, xsNumber): If Val(strValue) = 1 Then chkInfo(File_首页).Value = 1
        strValue = "": Call objXML.GetSingleNodeValue("advice", strValue, xsNumber): If Val(strValue) = 1 Then chkInfo(File_医嘱).Value = 1
        strValue = "": Call objXML.GetSingleNodeValue("cispath", strValue, xsNumber): If Val(strValue) = 1 Then chkInfo(File_路径).Value = 1
        strValue = "": Call objXML.GetSingleNodeValue("patipeis", strValue, xsNumber): If Val(strValue) = 1 Then chkInfo(File_体检).Value = 1
        
        '护理记录
        strValue = "": Call objXML.GetSingleNodeValue("nursing_record", strValue, xsNumber)
        If Val(strValue) = 1 Then
            chkInfo(File_护理).Value = 1
            strValue = "": Call objXML.GetSingleNodeValue("nursing_info/nursing_all", strValue, xsNumber): If Val(strValue) = 1 Then chkHlInfo(0).Value = 1
            strValue = "": Call objXML.GetSingleNodeValue("nursing_info/thermometer", strValue, xsNumber): If Val(strValue) = 1 Then chkHlInfo(1).Value = 1
            strValue = "": Call objXML.GetSingleNodeValue("nursing_info/record_file", strValue, xsNumber):
            If Val(strValue) = 1 Then
                chkHlInfo(2).Value = 1
                If GetXmlString(objXML, "nursing_info/file_name", strValue) Then
                    txtHlInfo.Text = strValue
                End If
            End If
        End If
        
        '检查报告
        strValue = "": Call objXML.GetSingleNodeValue("pacs_report", strValue, xsNumber)
        If Val(strValue) = 1 Then
            chkInfo(File_检查).Value = 1
            strValue = "": Call objXML.GetSingleNodeValue("pacs_info/pacs_type", strValue, xsNumber)
            'pacs_type =0所有检查报告 =1指定类型的检查报告
            optJcbg(Val(strValue)).Value = True
            If Val(strValue) = 1 Then

                If GetXmlString(objXML, "pacs_info/pacs_report_type/type_name", strValue) Then
                    txtJcbgTpye.Text = strValue
                End If
            End If
        End If
        
        '检验报告
        strValue = "": Call objXML.GetSingleNodeValue("lis_report", strValue, xsNumber)
        If Val(strValue) = 1 Then
            chkInfo(File_检验).Value = 1
            strValue = "": Call objXML.GetSingleNodeValue("lis_info/lis_type", strValue, xsNumber)
            'lis_type =0 所有检验报告 =1指定类型的检验报告
            optJybg(Val(strValue)).Value = True
            
            If Val(strValue) = 1 Then
                If GetXmlString(objXML, "lis_info/lis_report_type/type_name", strValue) Then
                    txtJybgTpye.Text = strValue
                End If
            End If
        End If
        
        '电子病历
        strValue = "": Call objXML.GetSingleNodeValue("emr", strValue, xsNumber)
        If Val(strValue) = 1 Then
            chkInfo(File_病历).Value = 1
            strValue = "": Call objXML.GetSingleNodeValue("emr_info/emr_type", strValue, xsNumber)
            'emr_type =0 所有电子病历  =1指定类型的电子病历  =1指定种类的电子病历
            optDzbl(Val(strValue)) = True
            
            If Val(strValue) = 1 Then
                If GetXmlString(objXML, "emr_info/standard_class/class_name", strValue) Then
                    txtDzblTpye(0).Text = strValue
                End If
            ElseIf Val(strValue) = 2 Then
                If GetXmlString(objXML, "emr_info/antetype_class/class_name", strValue) Then
                    txtDzblTpye(1).Text = strValue
                End If
            End If
        End If
    End If
    ReadXmlSet = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetXmlString(objXML As Object, ByVal strNode As String, ByRef strValue As String) As Boolean
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    strValue = ""
    If objXML.GetMultiNodeRecord(strNode, rsTmp) Then
        If Not rsTmp Is Nothing Then
            Do While Not rsTmp.EOF
                strValue = strValue & "," & rsTmp!node_value
                rsTmp.MoveNext
            Loop
            strValue = Mid(strValue, 2)
        End If
    End If
    GetXmlString = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetInfoXml() As String
    Dim objXML As New zl9ComLib.clsXML
    Dim i As Long
    
    On Error GoTo errH
    With objXML
        .ClearXmlText
        .AppendNode "app_info"                          '父节点[病人信息]
        .appendData "all_files", chkAllInfo.Value       '<所有内容>类型：N
        If chkAllInfo.Value <> 1 Then
            .appendData "medical_record", chkInfo(File_首页).Value  '<病案首页>类型：N
            .appendData "advice", chkInfo(File_医嘱).Value          '<医嘱清单>类型：N
            .appendData "emr", chkInfo(File_病历).Value             '<电子病历>类型：N
                If chkInfo(File_病历).Value = 1 Then
                    .AppendNode "emr_info"  '父节点[电子病历详细]
                        .appendData "emr_type", IIf(optDzbl(0).Value = True, 0, IIf(optDzbl(1).Value = True, 1, 2)) '<电子病历类型>类型：N
                        If optDzbl(1).Value And txtDzblTpye(0).Text <> "" Then
                            .AppendNode "standard_class"  '父节点[按标准分类]
                            For i = 0 To UBound(Split(txtDzblTpye(0).Text, ","))
                                .appendData "class_name", Split(txtDzblTpye(0).Text, ",")(i)
                            Next
                            .AppendNode "standard_class", True
                        ElseIf optDzbl(2).Value And txtDzblTpye(1).Text <> "" Then
                            .AppendNode "antetype_class"  '父节点[按病历原型]
                            For i = 0 To UBound(Split(txtDzblTpye(1).Text, ","))
                                .appendData "class_name", Split(txtDzblTpye(1).Text, ",")(i)
                            Next
                            .AppendNode "antetype_class", True
                        End If
                    .AppendNode "emr_info", True
                End If
            .appendData "nursing_record", chkInfo(File_护理).Value      '<护理记录>类型：N
                If chkInfo(File_护理).Value = 1 Then
                    .AppendNode "nursing_info"  '父节点[护理记录详细]
                        .appendData "nursing_all", chkHlInfo(0).Value  '<所有护理记录>类型：N
                        .appendData "thermometer", chkHlInfo(1).Value  '<是否允许访问体温单>类型：N
                        .appendData "record_file", chkHlInfo(2).Value   '<是否指定护理记录>类型：N
                        If chkHlInfo(2).Value = 1 And txtHlInfo.Text <> "" Then
                            For i = 0 To UBound(Split(txtHlInfo.Text, ","))
                                .appendData "file_name", Split(txtHlInfo.Text, ",")(i)
                            Next
                        End If
                    .AppendNode "nursing_info", True
                End If
            .appendData "pacs_report", chkInfo(File_检查).Value         '<检查报告>类型：N
                If chkInfo(File_检查).Value = 1 Then
                    .AppendNode "pacs_info"  '父节点[检查报告详细]
                        .appendData "pacs_type", IIf(optJcbg(0).Value = True, 0, 1) '<检查报告类型>类型：N
                        If optJcbg(1).Value And txtJcbgTpye.Text <> "" Then
                            .AppendNode "pacs_report_type"  '父节点[按标准分类]
                            For i = 0 To UBound(Split(txtJcbgTpye.Text, ","))
                                .appendData "type_name", Split(txtJcbgTpye.Text, ",")(i)
                            Next
                            .AppendNode "pacs_report_type", True
                        End If
                    .AppendNode "pacs_info", True
                End If
            .appendData "lis_report", chkInfo(File_检验).Value          '<检验报告>类型：N
                If chkInfo(File_检验).Value = 1 Then
                    .AppendNode "lis_info"  '父节点[检验报告详细]
                        .appendData "lis_type", IIf(optJybg(0).Value = True, 0, 1) '<检验报告类型>类型：N
                        If optJybg(1).Value And txtJybgTpye.Text <> "" Then
                            .AppendNode "lis_report_type"  '父节点[按标准分类]
                            For i = 0 To UBound(Split(txtJybgTpye.Text, ","))
                                .appendData "type_name", Split(txtJybgTpye.Text, ",")(i)
                            Next
                            .AppendNode "lis_report_type", True
                        End If
                    .AppendNode "lis_info", True
                End If
            .appendData "cispath", chkInfo(File_路径).Value             '<临床路径>类型：N
            .appendData "patipeis", chkInfo(File_体检).Value            '<病人体检>类型：N
         End If
        .AppendNode "app_info", True
        GetInfoXml = .XmlText
    End With
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim curDate As Date

    On Error GoTo errH
    Me.Caption = IIf(mintType = 0, "新增访问申请", "修改访问申请")
    lblTmp(0).Caption = Me.Caption
    'tabControl
    '-----------------------------------------------------
    With Me.tbcSub
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        '绑定子窗体时会Form_Load，且自动选中第一个加入的卡片
        '如果设置当前卡片隐藏,则不会自动切换选择,但显示内容未变
        '任意指定索引号无效，最终变为0-N，只是可能改变加入顺序。
        .InsertItem(0, "被访病人", picParent.hwnd, 0).Tag = "被访病人"
        .InsertItem(1, "访问内容", picAppInfo.hwnd, 0).Tag = "访问内容"
        .InsertItem(2, "访问时限和原因", picTime.hwnd, 0).Tag = "访问时限"
        
        .Item(2).Selected = True
        .Item(1).Selected = True
        .Item(0).Selected = True
    End With
    
    Call LoadDept
    
    '初始化病人表格
    Call InitPatiTable
    
     
    '执行结果下拉菜单初始化
    cboFind.Clear
    cboFind.AddItem "姓名"
    cboFind.AddItem "身份证号"
    cboFind.AddItem "门诊号"
    cboFind.AddItem "住院号"
    cboFind.AddItem "病人ID"
    cboFind.ListIndex = 0

    
    If mintType = 1 Then
        Call LoadPati
        Call ReadXmlSet
        Call SetFileCtl
        Call LoadOther
    Else
        '主界面勾选新增
        If Not mrsPati Is Nothing Then
            If mrsPati.State <> 0 Then
                If Not mrsPati.EOF Then
                    Call LoadPati
                End If
            End If
        End If
        chkAllInfo.Value = 1
        curDate = zlDatabase.Currentdate
        dtpTime(0).Value = Format(curDate, "yyyy-MM-dd hh:mm")
        dtpTime(1).Value = Format(curDate + 7, "yyyy-MM-dd hh:mm")
        optTimeTpye(0).Value = True
        SetFileCtl
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadOther()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    On Error GoTo errH
    strSQL = "Select a.访问开始时间, a.访问结束时间, a.内容时限, a.申请原因 From 电子病历访问申请 A Where a.Id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngApplyID)
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            dtpTime(0).Value = Format(rsTmp!访问开始时间 & "", "yyyy-MM-dd hh:mm")
            dtpTime(1).Value = Format(rsTmp!访问结束时间 & "", "yyyy-MM-dd hh:mm")
            optTimeTpye(Val(rsTmp!内容时限 & "")).Value = True
            txtReault.Text = rsTmp!申请原因 & ""
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadPati()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim lngRow As Long
    
    On Error GoTo errH
    With vsPati
        If mrsPati Is Nothing Then
            strSQL = "Select d.Id, d.类型, d.姓名, d.性别, d.年龄, g.名称 As 科室, d.标识号, d.当前状态" & vbNewLine & _
                    "From (Select Row_Number() Over(Partition By ID Order By 操作时间 Desc) As Top, c.*" & vbNewLine & _
                    "       From (Select '门诊' As 类型, a.病人id As ID, a.姓名, a.性别, a.年龄, a.执行部门id As 科室, a.执行时间 As 操作时间, a.门诊号 As 标识号," & vbNewLine & _
                    "                     Decode(a.执行状态, 1, '在' || To_Char(a.执行时间, 'yyyy-mm-dd') || '门诊就诊离院', '门诊正在就诊') As 当前状态" & vbNewLine & _
                    "              From 病人挂号记录 A, 电子病历申请访问病人 G" & vbNewLine & _
                    "              Where g.病人id = a.病人id And g.申请id = [1] And 记录状态 = 1" & vbNewLine & _
                    "              Union All" & vbNewLine & _
                    "              Select '住院' As 类型, b.病人id As ID, b.姓名, b.性别, b.年龄, b.出院科室id As 科室, b.入院日期 As 操作时间, b.住院号 As 标识号," & vbNewLine & _
                    "                     Decode(b.出院日期, Null, '在院', '第' || b.主页id || '次住院离院') As 当前状态" & vbNewLine & _
                    "              From 病案主页 B, 电子病历申请访问病人 H" & vbNewLine & _
                    "              Where h.病人id = b.病人id And h.申请id = [1]) C) D, 部门表 G" & vbNewLine & _
                    "Where g.Id = d.科室 And d.Top = 1" & vbNewLine & _
                    "Order By d.操作时间 Desc"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngApplyID)
        Else
            Set rsTmp = mrsPati
            If Not rsTmp Is Nothing Then rsTmp.MoveFirst
        End If
        If Not rsTmp Is Nothing Then
            Do While Not rsTmp.EOF
                If InStr(.Tag, "," & rsTmp!ID & ",") <= 0 Then
                    If Val(.TextMatrix(.Rows - 1, col_病人Id)) <> 0 Then
                        .Rows = .Rows + 1
                    End If
                    lngRow = .Rows - 1
                    
                    .TextMatrix(lngRow, col_病人Id) = rsTmp!ID & ""
                    .TextMatrix(lngRow, col_姓名) = rsTmp!姓名 & ""
                    Set .Cell(flexcpPicture, lngRow, col_姓名) = img16.ListImages(IIf(rsTmp!性别 & "" = "女", "girl", "boy")).Picture
                    .TextMatrix(lngRow, col_性别) = rsTmp!性别 & ""
                    .TextMatrix(lngRow, col_年龄) = rsTmp!年龄 & ""
                    .TextMatrix(lngRow, COL_标识号) = rsTmp!标识号 & ""
                    .TextMatrix(lngRow, col_科室) = rsTmp!科室 & ""
                    .TextMatrix(lngRow, COL_当前状态) = rsTmp!当前状态 & ""
                    .Tag = .Tag & "," & rsTmp!ID & ","
                End If
                rsTmp.MoveNext
            Loop
        End If
        .WordWrap = True
        '自动调整行高
        .AutoSize col_姓名, COL_当前状态
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub




Private Sub Form_Unload(Cancel As Integer)
    Set mrsPati = Nothing
End Sub

Private Sub imgDzblTpye_Click(Index As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim vPoint As POINTAPI
    Dim blnCancel As Boolean
    Dim strTmp As String

    If txtDzblTpye(Index).BackColor = &H80000004 Then Exit Sub
    vPoint = zlcontrol.GetCoordPos(imgDzblTpye(Index).Container.hwnd, imgDzblTpye(Index).Left, imgDzblTpye(Index).Top)
    blnCancel = True
    On Error GoTo errH
    If Index = 0 Then
        strSQL = "select 1 as ID, '门诊病历' as 类型名称," & IIf(InStr("," & txtDzblTpye(0).Text & ",", ",门诊病历,"), 1, 0) & " as 已勾选check from dual" & vbNewLine & _
                "union all" & vbNewLine & _
                "select 2 as ID, '住院病历' as 类型名称," & IIf(InStr("," & txtDzblTpye(0).Text & ",", ",住院病历,"), 1, 0) & " as 已勾选check from dual" & vbNewLine & _
                "union all" & vbNewLine & _
                "select 4 as ID, '护理病历' as 类型名称," & IIf(InStr("," & txtDzblTpye(0).Text & ",", ",护理病历,"), 1, 0) & " as 已勾选check from dual" & vbNewLine & _
                "union all" & vbNewLine & _
                "select 5 as ID, '疾病证明报告' as 类型名称," & IIf(InStr("," & txtDzblTpye(0).Text & ",", ",疾病证明报告,"), 1, 0) & " as 已勾选check from dual" & vbNewLine & _
                "union all" & vbNewLine & _
                "select 6 as ID, '知情文件' as 类型名称," & IIf(InStr("," & txtDzblTpye(0).Text & ",", ",知情文件,"), 1, 0) & " as 已勾选check from dual"
        Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSQL, 0, "选择病历文件种类", True, "", "", True, True, True, vPoint.X, vPoint.Y, imgDzblTpye(0).Height, blnCancel, True, True)
        If Not blnCancel Then
            If Not rsTmp Is Nothing Then
                Do While Not rsTmp.EOF
                    strTmp = strTmp & "," & rsTmp!类型名称
                    rsTmp.MoveNext
                Loop
                txtDzblTpye(0).Text = Mid(strTmp, 2)
            Else
                MsgBox "未查找到可以选择的病历文件种类!", vbInformation, Me.Caption
                Exit Sub
            End If
        End If
    Else

        '新病历
        
        If mstrNewEMR = "" Then
            Set rsTmp = Nothing
            On Error Resume Next
            If Not gobjEmr Is Nothing Then
                strSQL = "Select Title as 名称 From Antetype_List Where nvl(disable,0)=0 Order By Code"
                Call gobjEmr.OpenSQLRecordset(strSQL, "", rsTmp)
            End If
            Err.Clear: On Error GoTo 0
            If Not rsTmp Is Nothing Then
                Do While Not rsTmp.EOF
                    mstrNewEMR = mstrNewEMR & "," & rsTmp!名称
                    rsTmp.MoveNext
                Loop
            End If
            mstrNewEMR = Mid(mstrNewEMR, 2)
        End If
            
        strSQL = ""
        If mstrNewEMR <> "" Then
            strSQL = "Select Rownum + 100000 As ID, '新版病历' As 病历种类, b.C2 As 名称, Decode(d.C2, Null, 0, 1) As 已勾选check" & vbNewLine & _
                     "From Table(Cast(f_Str2list2([2]) As Zltools.t_Strlist2)) B," & vbNewLine & _
                     "     (Select Replace(C2, '【新版病历】', '') As C2" & vbNewLine & _
                     "       From Table(Cast(f_Str2list2([1]) As Zltools.t_Strlist2)) C" & vbNewLine & _
                     "       Where Instr(C2, '【新版病历】') > 0) D" & vbNewLine & _
                     "Where b.C2 = d.C2(+) union all "
        End If

        strSQL = strSQL & " Select a.ID,Decode(a.种类, 1, '门诊病历', 2, '住院病历', 4, '护理病历', 5, '疾病证明', 6, '知情文件') As 病历种类, a.名称," & vbNewLine & _
                "       Decode(b.C2, Null, 0, 1) As 已勾选check" & vbNewLine & _
                "From 病历文件列表 A, Table(Cast(f_Str2list2([1]) As Zltools.t_Strlist2)) B" & vbNewLine & _
                "Where a.种类 In (1, 2, 4, 5, 6) And a.名称 = b.C2(+)" & vbNewLine & _
                "Order By 种类, 编号"

        Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSQL, 0, "选择病历文件", True, "", "", True, True, True, vPoint.X, vPoint.Y, imgDzblTpye(1).Height, blnCancel, True, True, txtDzblTpye(1).Text, mstrNewEMR)
        If Not blnCancel Then
            If Not rsTmp Is Nothing Then
                Do While Not rsTmp.EOF
                    strTmp = strTmp & "," & IIf(rsTmp!病历种类 & "" = "新版病历", "【新版病历】", "") & rsTmp!名称
                    rsTmp.MoveNext
                Loop
                txtDzblTpye(1).Text = Mid(strTmp, 2)
            Else
                MsgBox "未查找到可以选择的病历文件!", vbInformation, Me.Caption
                Exit Sub
            End If
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub imgHlInfo_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim vPoint As POINTAPI
    Dim blnCancel As Boolean
    Dim strTmp As String
    
    If txtHlInfo.BackColor = &H80000004 Then Exit Sub
    vPoint = zlcontrol.GetCoordPos(imgHlInfo.Container.hwnd, imgHlInfo.Left, imgHlInfo.Top)
    blnCancel = True
    On Error GoTo errH
    
    strSQL = "Select a.ID,'护理记录' As 病历种类, a.名称," & vbNewLine & _
            "       Decode(b.C2, Null, 0, 1) As 已勾选check" & vbNewLine & _
            "From 病历文件列表 A, Table(Cast(f_Str2list2([1]) As Zltools.t_Strlist2)) B" & vbNewLine & _
            "Where a.种类 =3 AND A.保留<>-1 And a.名称 = b.C2(+)" & vbNewLine & _
            "Order By 种类, 编号"

    Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSQL, 0, "选择护理记录文件", True, "", "", True, True, True, vPoint.X, vPoint.Y, imgDzblTpye(1).Height, blnCancel, True, True, txtHlInfo.Text)
    If Not blnCancel Then
        If Not rsTmp Is Nothing Then
            Do While Not rsTmp.EOF
                strTmp = strTmp & "," & rsTmp!名称
                rsTmp.MoveNext
            Loop
            txtHlInfo.Text = Mid(strTmp, 2)
        Else
            MsgBox "未查找到可以选择的护理记录文件!", vbInformation, Me.Caption
            Exit Sub
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub imgJcbgTpye_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim vPoint As POINTAPI
    Dim blnCancel As Boolean
    Dim strTmp As String
    
    If txtJcbgTpye.BackColor = &H80000004 Then Exit Sub
    vPoint = zlcontrol.GetCoordPos(imgJcbgTpye.Container.hwnd, imgJcbgTpye.Left, imgJcbgTpye.Top)
    blnCancel = True
    On Error GoTo errH
    
    strSQL = "Select a.编码 As ID, a.名称, Decode(b.C2, Null, 0, 1) As 已勾选check" & vbNewLine & _
            "From 诊疗检查类型 A, Table(Cast(f_Str2list2([1]) As Zltools.t_Strlist2)) B" & vbNewLine & _
            "Where a.名称 = b.C2(+)" & vbNewLine & _
            "Order By 编码"
    Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSQL, 0, "选择检查报告类型", True, "", "", True, True, True, vPoint.X, vPoint.Y, imgJcbgTpye.Height, blnCancel, True, True, txtJcbgTpye.Text)
    If Not blnCancel Then
        If Not rsTmp Is Nothing Then
            Do While Not rsTmp.EOF
                strTmp = strTmp & "," & rsTmp!名称
                rsTmp.MoveNext
            Loop
            txtJcbgTpye.Text = Mid(strTmp, 2)
        Else
            MsgBox "未查找到可以选择的检查报告类型!", vbInformation, Me.Caption
            Exit Sub
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub imgJybgTpye_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim vPoint As POINTAPI
    Dim blnCancel As Boolean
    Dim strTmp As String
    
    If txtJybgTpye.BackColor = &H80000004 Then Exit Sub
    vPoint = zlcontrol.GetCoordPos(imgJybgTpye.Container.hwnd, imgJybgTpye.Left, imgJybgTpye.Top)
    blnCancel = True
    On Error GoTo errH
    
    strSQL = "Select a.编码 As ID, a.名称, Decode(b.C2, Null, 0, 1) As 已勾选check" & vbNewLine & _
            "From 诊疗检验类型 A, Table(Cast(f_Str2list2([1]) As Zltools.t_Strlist2)) B" & vbNewLine & _
            "Where a.名称 = b.C2(+)" & vbNewLine & _
            "Order By 编码"
    Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSQL, 0, "选择检验报告类型", True, "", "", True, True, True, vPoint.X, vPoint.Y, imgJybgTpye.Height, blnCancel, True, True, txtJybgTpye.Text)
    If Not blnCancel Then
        If Not rsTmp Is Nothing Then
            Do While Not rsTmp.EOF
                strTmp = strTmp & "," & rsTmp!名称
                rsTmp.MoveNext
            Loop
            txtJybgTpye.Text = Mid(strTmp, 2)
        Else
            MsgBox "未查找到可以选择的检验报告类型!", vbInformation, Me.Caption
            Exit Sub
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub imgSentence_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim vPoint As POINTAPI
    Dim blnCancel As Boolean
    Dim lngRow As Long
    
    vPoint = zlcontrol.GetCoordPos(cboDept.Container.hwnd, cboDept.Left, cboDept.Top)
    blnCancel = True
    On Error GoTo errH
    
    If lblDept.Tag = "" Then
        strSQL = "Select d.Id,d.类型, d.姓名, d.性别, d.年龄, g.名称 As 科室, d.标识号,d.当前状态" & vbNewLine & _
                    "From (Select Row_Number() Over(Partition By ID Order By 操作时间 Desc) As Top, c.*" & vbNewLine & _
                    "       From (Select '门诊' As 类型, a.病人id As ID, a.姓名, a.性别, a.年龄, a.执行部门id As 科室, a.执行时间 As 操作时间, a.门诊号 As 标识号,decode(A.执行状态,1,'在'||to_char(A.执行时间,'yyyy-mm-dd') || '门诊就诊离院','门诊正在就诊') as 当前状态" & vbNewLine & _
                    "              From 病人挂号记录 A" & vbNewLine & _
                    "              Where 记录状态=1 And a.执行时间 Between Sysdate - 7 And Sysdate" & vbNewLine & _
                    "              Union All" & vbNewLine & _
                    "              Select '住院' As 类型, b.病人id As ID, b.姓名, b.性别, b.年龄, b.出院科室id As 科室, b.入院日期 As 操作时间, b.住院号 As 标识号,decode(B.出院日期,null,'在院','第'||b.主页id||'次住院离院') as 当前状态" & vbNewLine & _
                    "              From 病案主页 B" & vbNewLine & _
                    "              Where  b.入院日期 Between Sysdate - 7 And Sysdate) C) D, 部门表 G" & vbNewLine & _
                    "Where g.Id = d.科室 And d.Top = 1" & IIf(cboDept.Text = "所有部门", "", " And D.科室=[1]") & vbNewLine & _
                    "Order By d.操作时间 Desc"
    ElseIf lblDept.Tag = "门诊" Then
        strSQL = "Select d.Id,d.类型, d.姓名, d.性别, d.年龄, g.名称 As 科室, d.标识号,d.当前状态" & vbNewLine & _
                    "From (Select Row_Number() Over(Partition By ID Order By 操作时间 Desc) As Top, c.*" & vbNewLine & _
                    "       From (Select '门诊' As 类型, a.病人id As ID, a.姓名, a.性别, a.年龄, a.执行部门id As 科室, a.执行时间 As 操作时间, a.门诊号 As 标识号,decode(A.执行状态,1,'在'||to_char(A.执行时间,'yyyy-mm-dd') || '门诊就诊离院','门诊正在就诊') as 当前状态" & vbNewLine & _
                    "              From 病人挂号记录 A" & vbNewLine & _
                    "              Where 记录状态=1 And a.执行时间 Between Sysdate - 7 And Sysdate) C) D, 部门表 G" & vbNewLine & _
                    "Where g.Id = d.科室 And d.Top = 1" & IIf(cboDept.Text = "所有部门", "", " And D.科室=[1]") & vbNewLine & _
                    "Order By d.操作时间 Desc"

    ElseIf lblDept.Tag = "住院" Then
        strSQL = "Select d.Id,d.类型, d.姓名, d.性别, d.年龄, g.名称 As 科室, d.标识号,d.当前状态" & vbNewLine & _
                    "From (Select Row_Number() Over(Partition By ID Order By 操作时间 Desc) As Top, c.*" & vbNewLine & _
                    "       From (Select '住院' As 类型, b.病人id As ID, b.姓名, b.性别, b.年龄, b.出院科室id As 科室, b.入院日期 As 操作时间, b.住院号 As 标识号,decode(B.出院日期,null,'在院','第'||b.主页id||'次住院离院') as 当前状态" & vbNewLine & _
                    "              From 病案主页 B" & vbNewLine & _
                    "              Where  b.入院日期 Between Sysdate - 7 And Sysdate) C) D, 部门表 G" & vbNewLine & _
                    "Where g.Id = d.科室 And d.Top = 1" & IIf(cboDept.Text = "所有部门", "", " And D.科室=[1]") & vbNewLine & _
                    "Order By d.操作时间 Desc"
    End If
    
    Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSQL, 0, "选择最近7天的病人", True, "", "", True, True, True, vPoint.X, vPoint.Y, cboDept.Height, blnCancel, True, True, cboDept.ItemData(cboDept.ListIndex))
    With vsPati
        If Not blnCancel Then
            If Not rsTmp Is Nothing Then
                Do While Not rsTmp.EOF
                    If InStr(.Tag, "," & rsTmp!ID & ",") <= 0 Then
                        If Val(.TextMatrix(.Rows - 1, col_病人Id)) <> 0 Then
                            .Rows = .Rows + 1
                        End If
                        lngRow = .Rows - 1
                        
                        .TextMatrix(lngRow, col_病人Id) = rsTmp!ID & ""
                        .TextMatrix(lngRow, col_姓名) = rsTmp!姓名 & ""
                        Set .Cell(flexcpPicture, lngRow, col_姓名) = img16.ListImages(IIf(rsTmp!性别 & "" = "女", "girl", "boy")).Picture
                        .TextMatrix(lngRow, col_性别) = rsTmp!性别 & ""
                        .TextMatrix(lngRow, col_年龄) = rsTmp!年龄 & ""
                        .TextMatrix(lngRow, COL_标识号) = rsTmp!标识号 & ""
                        .TextMatrix(lngRow, col_科室) = rsTmp!科室 & ""
                        .TextMatrix(lngRow, COL_当前状态) = rsTmp!当前状态 & ""
                        .Tag = .Tag & "," & rsTmp!ID & ","
                    End If
                    rsTmp.MoveNext
                Loop
                .WordWrap = True
                '自动调整行高
                .AutoSize col_姓名, COL_当前状态
            Else
                MsgBox "未查找到本科室近期的" & lblDept.Tag & "病人!", vbInformation, Me.Caption
                Exit Sub
            End If
        End If
        
    End With
    Exit Sub
errH:
    MsgBox "未查找到本科室近期的" & lblDept.Tag & "病人!", vbInformation, gstrSysName
    blnCancel = True
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub lblDept_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetLBLFace(lblDept, True)
End Sub

Private Sub optDzbl_Click(Index As Integer)
    Call SetFileCtl
End Sub

Private Sub optJcbg_Click(Index As Integer)
    Call SetFileCtl
End Sub

Private Sub optJybg_Click(Index As Integer)
    Call SetFileCtl
End Sub

Private Sub picPati_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetLBLFace(lblDept, False)
End Sub

Private Sub txtDzblTpye_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    Call ZLCommFun.ShowTipInfo(txtDzblTpye(Index).hwnd, Replace(txtDzblTpye(Index).Text, ",", "、" & vbCrLf), True, True)
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim vPoint As POINTAPI
    Dim blnCancel As Boolean
    Dim lngRow As Long
    Dim colPati As Collection, str病人ids As String, i As Long
    
    
    
    If KeyAscii = vbKeyReturn Then
        If Len(txtFind.Text) < 1 Then Exit Sub
        vPoint = zlcontrol.GetCoordPos(cboDept.Container.hwnd, cboDept.Left, cboDept.Top)
        blnCancel = True
        On Error GoTo errH

        If cboFind.Text = "门诊号" Then
            strSQL = "Select d.Id,d.类型, d.姓名, d.性别, d.年龄, g.名称 As 科室, d.标识号,d.当前状态" & vbNewLine & _
                        "From (Select Row_Number() Over(Partition By ID Order By 操作时间 Desc) As Top, c.*" & vbNewLine & _
                        "       From (Select '门诊' As 类型, a.病人id As ID, a.姓名, a.性别, a.年龄, a.执行部门id As 科室, a.执行时间 As 操作时间, a.门诊号 As 标识号,decode(A.执行状态,1,'在'||to_char(A.执行时间,'yyyy-mm-dd') || '门诊就诊离院','门诊正在就诊') as 当前状态" & vbNewLine & _
                        "              From 病人挂号记录 A" & vbNewLine & _
                        "              Where A.记录状态=1 And A.门诊号=[2]) C) D, 部门表 G" & vbNewLine & _
                        "Where g.Id = d.科室 And d.Top = 1" & IIf(cboDept.Text = "所有部门", "", " And D.科室=[1]") & vbNewLine & _
                        "Order By d.操作时间 Desc"
        ElseIf cboFind.Text = "住院号" Then
            strSQL = "Select d.Id,d.类型, d.姓名, d.性别, d.年龄, g.名称 As 科室, d.标识号,d.当前状态" & vbNewLine & _
                        "From (Select Row_Number() Over(Partition By ID Order By 操作时间 Desc) As Top, c.*" & vbNewLine & _
                        "       From (Select '住院' As 类型, b.病人id As ID, b.姓名, b.性别, b.年龄, b.出院科室id As 科室, b.入院日期 As 操作时间, b.住院号 As 标识号,decode(B.出院日期,null,'在院','第'||b.主页id||'次住院离院') as 当前状态" & vbNewLine & _
                        "              From 病案主页 B" & vbNewLine & _
                        "              Where B.住院号=[2]) C) D, 部门表 G" & vbNewLine & _
                        "Where g.Id = d.科室 And d.Top = 1" & IIf(cboDept.Text = "所有部门", "", " And D.科室=[1]") & vbNewLine & _
                        "Order By d.操作时间 Desc"
        Else
            If cboFind.Text = "身份证号" Then
                Set colPati = PatiSvrGetpatiinfo(1, 0, 1240, 0, 2, txtFind.Text)
            End If
        
            If Not colPati Is Nothing Then
                If colPati.Count > 0 Then
                    For i = 1 To colPati.Count
                        If InStr("," & str病人ids & ",", "," & Val(GetColVal(colPati(i), "_pati_id")) & ",") = 0 Then
                           str病人ids = str病人ids & "," & Val(GetColVal(colPati(i), "_pati_id"))
                        End If
                    Next
                End If
            End If
            If str病人ids <> "" Then str病人ids = Mid(str病人ids, 2)
        
            
            If lblDept.Tag = "" Then
                strSQL = "Select d.Id,d.类型, d.姓名, d.性别, d.年龄, g.名称 As 科室, d.标识号,d.当前状态" & vbNewLine & _
                            "From (Select Row_Number() Over(Partition By ID Order By 操作时间 Desc) As Top, c.*" & vbNewLine & _
                            "       From (Select '门诊' As 类型, a.病人id As ID, a.姓名, a.性别, a.年龄, a.执行部门id As 科室, a.执行时间 As 操作时间, a.门诊号 As 标识号,decode(A.执行状态,1,'在'||to_char(A.执行时间,'yyyy-mm-dd') || '门诊就诊离院','门诊正在就诊') as 当前状态" & vbNewLine & _
                            "              From 病人挂号记录 A" & vbNewLine & _
                            "              Where A.记录状态=1 And " & decode(cboFind.Text, "身份证号", " A.病人ID in (Select Column_Value As 病人id From Table(Cast(f_Str2list([3]) As Zltools.t_Strlist))) ", "病人ID", "A.病人ID =[2]", "姓名", "A.姓名 like [2]") & vbNewLine & _
                            "              Union All" & vbNewLine & _
                            "              Select '住院' As 类型, b.病人id As ID, b.姓名, b.性别, b.年龄, b.出院科室id As 科室, b.入院日期 As 操作时间, b.住院号 As 标识号,decode(B.出院日期,null,'在院','第'||b.主页id||'次住院离院') as 当前状态" & vbNewLine & _
                            "              From 病案主页 B" & vbNewLine & _
                            "              Where " & decode(cboFind.Text, "身份证号", " b.病人ID in (Select Column_Value As 病人id From Table(Cast(f_Str2list([3]) As Zltools.t_Strlist))) ", "病人ID", "B.病人ID =[2]", "姓名", "B.姓名 like [2]") & ") C) D, 部门表 G" & vbNewLine & _
                            "Where g.Id = d.科室 And d.Top = 1" & IIf(cboDept.Text = "所有部门", "", " And D.科室=[1]") & vbNewLine & _
                            "Order By d.操作时间 Desc"
            ElseIf lblDept.Tag = "住院" Then
                strSQL = "Select d.Id,d.类型, d.姓名, d.性别, d.年龄, g.名称 As 科室, d.标识号,d.当前状态" & vbNewLine & _
                            "From (Select Row_Number() Over(Partition By ID Order By 操作时间 Desc) As Top, c.*" & vbNewLine & _
                            "       From (Select '住院' As 类型, b.病人id As ID, b.姓名, b.性别, b.年龄, b.出院科室id As 科室, b.入院日期 As 操作时间, b.住院号 As 标识号,decode(B.出院日期,null,'在院','第'||b.主页id||'次住院离院') as 当前状态" & vbNewLine & _
                            "              From 病案主页 B" & vbNewLine & _
                            "              Where " & decode(cboFind.Text, "身份证号", " b.病人ID in (Select Column_Value As 病人id From Table(Cast(f_Str2list([3]) As Zltools.t_Strlist))) ", "病人ID", "B.病人ID =[2]", "姓名", "B.姓名 like [2]") & ") C) D, 部门表 G" & vbNewLine & _
                            "Where g.Id = d.科室 And d.Top = 1" & IIf(cboDept.Text = "所有部门", "", " And D.科室=[1]") & vbNewLine & _
                            "Order By d.操作时间 Desc"

            ElseIf lblDept.Tag = "门诊" Then
                strSQL = "Select d.Id,d.类型, d.姓名, d.性别, d.年龄, g.名称 As 科室, d.标识号,d.当前状态" & vbNewLine & _
                            "From (Select Row_Number() Over(Partition By ID Order By 操作时间 Desc) As Top, c.*" & vbNewLine & _
                            "       From (Select '门诊' As 类型, a.病人id As ID, a.姓名, a.性别, a.年龄, a.执行部门id As 科室, a.执行时间 As 操作时间, a.门诊号 As 标识号,decode(A.执行状态,1,'在'||to_char(A.执行时间,'yyyy-mm-dd') || '门诊就诊离院','门诊正在就诊') as 当前状态" & vbNewLine & _
                            "              From 病人挂号记录 A" & vbNewLine & _
                            "              Where A.记录状态=1 And " & decode(cboFind.Text, "身份证号", " A.病人ID in (Select Column_Value As 病人id From Table(Cast(f_Str2list([3]) As Zltools.t_Strlist))) ", "病人ID", "A.病人ID =[2]", "姓名", "A.姓名 like [2]") & ") C) D, 部门表 G" & vbNewLine & _
                            "Where g.Id = d.科室 And d.Top = 1" & IIf(cboDept.Text = "所有部门", "", " And D.科室=[1]") & vbNewLine & _
                            "Order By d.操作时间 Desc"
            End If
        End If
        Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, strSQL, 0, "查找病人", True, "", "", True, True, True, vPoint.X, vPoint.Y, cboDept.Height, blnCancel, True, True, cboDept.ItemData(cboDept.ListIndex), IIf(InStr(",门诊号,住院号,病人ID,", cboFind.Text) > 0, Val(txtFind.Text), IIf(cboFind.Text = "姓名", txtFind.Text & "%", txtFind.Text)), str病人ids)
        With vsPati
            If Not blnCancel Then
                If Not rsTmp Is Nothing Then
                    Do While Not rsTmp.EOF
                        If InStr(.Tag, "," & rsTmp!ID & ",") <= 0 Then
                            If Val(.TextMatrix(.Rows - 1, col_病人Id)) <> 0 Then
                                .Rows = .Rows + 1
                            End If
                            lngRow = .Rows - 1
                            
                            .TextMatrix(lngRow, col_病人Id) = rsTmp!ID & ""
                            .TextMatrix(lngRow, col_姓名) = rsTmp!姓名 & ""
                            Set .Cell(flexcpPicture, lngRow, col_姓名) = img16.ListImages(IIf(rsTmp!性别 & "" = "女", "girl", "boy")).Picture
                            .TextMatrix(lngRow, col_性别) = rsTmp!性别 & ""
                            .TextMatrix(lngRow, col_年龄) = rsTmp!年龄 & ""
                            .TextMatrix(lngRow, COL_标识号) = rsTmp!标识号 & ""
                            .TextMatrix(lngRow, col_科室) = rsTmp!科室 & ""
                            .TextMatrix(lngRow, COL_当前状态) = rsTmp!当前状态 & ""
                            .Tag = .Tag & "," & rsTmp!ID & ","
                        End If
                        rsTmp.MoveNext
                    Loop
                    .WordWrap = True
                    '自动调整行高
                    .AutoSize col_姓名, COL_当前状态
                Else
                    MsgBox "在当前范围未查找到" & lblDept.Tag & "病人!", vbInformation, Me.Caption
                    Exit Sub
                End If
            End If
        End With
    Else
        Select Case cboFind.Text
            Case "住院号", "门诊号", "病人ID"
                If InStr("0123456789" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 And InStr(",3,22,24,", "," & KeyAscii & ",") = 0 Then KeyAscii = 0
            Case "身份证号"
                If InStr("1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 And InStr(",3,22,24,", "," & KeyAscii & ",") = 0 Then KeyAscii = 0
            Case "姓名"
        End Select
    End If
    Exit Sub
errH:
    MsgBox "在当前范围未查找到" & lblDept.Tag & "病人!", vbInformation, gstrSysName
    blnCancel = True
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadDept()
'加载查询科室
    Dim rsTmp As Recordset
    Dim strSQL As String
    Dim i As Long
    Dim strFiter As String
    
    strSQL = "Select B.ID,B.编码,B.名称 From " & _
            " 部门表 B, 部门性质说明 C" & vbNewLine & _
            " Where B.Id = C.部门id " & _
            "  And C.工作性质 = '临床' " & decode(lblDept.Tag, "", " And C.服务对象 <> 0 ", "门诊", " And C.服务对象 in (1,3) ", "住院", " And C.服务对象 in (2,3) ") & "  And (B.撤档时间 = To_Date('3000-01-01', 'YYYY-MM-DD') Or B.撤档时间 Is Null) Order By B.编码"

    On Error GoTo errH
    cboDept.Clear
    '所有部门
    cboDept.AddItem "所有部门"
    cboDept.ItemData(cboDept.NewIndex) = -1
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    
    For i = 1 To rsTmp.RecordCount
        cboDept.AddItem rsTmp!编码 & "-" & rsTmp!名称
        cboDept.ItemData(cboDept.NewIndex) = rsTmp!ID & ""
        rsTmp.MoveNext
    Next
    If cboDept.ListIndex = -1 And cboDept.ListCount > 0 Then
        Call Cbo.SetIndex(cboDept.hwnd, 0)
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub txtFind_GotFocus()
    If txtFind.Text <> "" Then
        Call zlcontrol.TxtSelAll(txtFind)
    End If
End Sub

Private Sub txtHlInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ZLCommFun.ShowTipInfo(txtHlInfo.hwnd, Replace(txtHlInfo.Text, ",", "、" & vbCrLf), True, True)
End Sub

Private Sub txtJcbgTpye_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ZLCommFun.ShowTipInfo(txtJcbgTpye.hwnd, Replace(txtJcbgTpye.Text, ",", "、" & vbCrLf), True, True)
End Sub

Private Sub txtJybgTpye_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call ZLCommFun.ShowTipInfo(txtJybgTpye.hwnd, Replace(txtJybgTpye.Text, ",", "、" & vbCrLf), True, True)
End Sub


Private Sub InitPatiTable()
'功能：初始化病人清单格式
    Dim arrHead As Variant, strHead As String, i As Long, lngWidth As Long

    strHead = "病人ID;姓名,1300,1;性别,700,4;年龄,700,4;标识号,950,1;科室,1000,1;当前状态,1700,1"
    arrHead = Split(strHead, ";")
    With vsPati
        .Clear
        .FixedRows = 1
        .FixedCols = 0
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        .SelectionMode = flexSelectionByRow
        .FocusRect = flexFocusLight
        .HighLight = flexHighlightWithFocus
        .BackColorSel = &HFAEADA

        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            .FixedAlignment(.FixedCols + i) = 4
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                lngWidth = Val(Split(arrHead(i), ",")(1))
                .ColWidth(.FixedCols + i) = lngWidth
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                '为了支持zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0    '为了支持zl9PrintMode
            End If
            .colData(.FixedCols + i) = .ColWidth(.FixedCols + i)    '记录原始列宽用于列选择器
        Next
        .Editable = flexEDNone
    End With
End Sub


Private Sub txtReault_GotFocus()
    Call zlcontrol.TxtSelAll(txtReault)
End Sub

Private Sub txtReault_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub lblDept_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim objPopup As CommandBar
    Dim objControl As CommandBarControl
    Dim vRect As RECT, strSQL As String
    Dim str单位 As String
    Dim rsTmp As Recordset
    
    On Error GoTo errH
    
    Set objPopup = cbsMain.Add("Popup", xtpBarPopup)
    With objPopup.Controls
        Set objControl = .Add(xtpControlButton, Cmd_所有科室, "所有科室")
        objControl.Parameter = ""
        Set objControl = .Add(xtpControlButton, Cmd_住院科室, "住院科室")
        objControl.Parameter = "住院"
        Set objControl = .Add(xtpControlButton, Cmd_门诊科室, "门诊科室")
        objControl.Parameter = "门诊"
    End With
    GetWindowRect picPati.hwnd, vRect
    objPopup.ShowPopup , vRect.Left * Screen.TwipsPerPixelX + lblDept.Left + lblDept.Width, vRect.Top * Screen.TwipsPerPixelY + lblDept.Top
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub



Private Sub SetLBLFace(ByRef objCtl As Object, ByVal blnOver As Boolean)
    If blnOver Then
        If objCtl.BorderStyle = 0 Then
            objCtl.BorderStyle = 1
            objCtl.BackStyle = 1
        End If
    Else
        If objCtl.BorderStyle = 1 Then
            objCtl.BorderStyle = 0
            objCtl.BackStyle = 0
        End If
    End If
End Sub
