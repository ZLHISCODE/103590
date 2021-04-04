VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmCaseTendBodySetData 
   Caption         =   "体温数据编辑"
   ClientHeight    =   5580
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   8910
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCaseTendBodySetData.frx":0000
   LockControls    =   -1  'True
   ScaleHeight     =   5580
   ScaleWidth      =   8910
   StartUpPosition =   2  '屏幕中心
   Begin VB.PictureBox picSplit 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   45
      Left            =   1845
      MousePointer    =   7  'Size N S
      ScaleHeight     =   45
      ScaleWidth      =   6255
      TabIndex        =   50
      Top             =   4725
      Width           =   6255
   End
   Begin VB.Frame fraOper 
      Caption         =   "设置手术/分娩"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   150
      TabIndex        =   46
      Top             =   4005
      Width           =   5415
      Begin zl9TemperatureChartGD.VsfGrid vsfOper 
         Height          =   570
         Left            =   165
         TabIndex        =   49
         Top             =   285
         Width           =   3975
         _ExtentX        =   7011
         _ExtentY        =   1005
      End
   End
   Begin VB.Timer tmr1 
      Interval        =   60
      Left            =   7680
      Top             =   1440
   End
   Begin VB.PictureBox picStb 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      FillColor       =   &H80000008&
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   1440
      ScaleHeight     =   360
      ScaleWidth      =   2415
      TabIndex        =   28
      TabStop         =   0   'False
      Top             =   5280
      Width           =   2415
      Begin VB.Label lblStb 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Height          =   195
         Left            =   0
         TabIndex        =   29
         Top             =   0
         Width           =   45
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   5400
      Top             =   360
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   4680
      ScaleHeight     =   2655
      ScaleWidth      =   3855
      TabIndex        =   15
      TabStop         =   0   'False
      Top             =   1800
      Width           =   3855
      Begin VB.Frame FraTable 
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   0
         TabIndex        =   16
         Top             =   120
         Width           =   3735
         Begin VB.PictureBox picEdit 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            ForeColor       =   &H80000008&
            Height          =   285
            Left            =   2175
            ScaleHeight     =   255
            ScaleWidth      =   1305
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   540
            Visible         =   0   'False
            Width           =   1335
            Begin VB.CommandButton cmdColor 
               Caption         =   "‥"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Left            =   1080
               TabIndex        =   22
               Top             =   30
               Width           =   285
            End
            Begin VB.TextBox txtEdit 
               BackColor       =   &H80000018&
               BorderStyle     =   0  'None
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Left            =   480
               TabIndex        =   21
               Top             =   0
               Width           =   800
            End
            Begin VB.PictureBox picHour 
               Appearance      =   0  'Flat
               BackColor       =   &H80000018&
               ForeColor       =   &H80000008&
               Height          =   285
               Left            =   -10
               ScaleHeight     =   255
               ScaleWidth      =   465
               TabIndex        =   53
               TabStop         =   0   'False
               Top             =   -10
               Visible         =   0   'False
               Width           =   495
               Begin VB.TextBox txtHour 
                  Alignment       =   2  'Center
                  BackColor       =   &H80000018&
                  BorderStyle     =   0  'None
                  BeginProperty Font 
                     Name            =   "宋体"
                     Size            =   9
                     Charset         =   0
                     Weight          =   400
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  Height          =   195
                  Left            =   0
                  MaxLength       =   2
                  TabIndex        =   19
                  Top             =   15
                  Width           =   315
               End
               Begin VB.Label lblHour 
                  Appearance      =   0  'Flat
                  AutoSize        =   -1  'True
                  BackColor       =   &H80000005&
                  BackStyle       =   0  'Transparent
                  Caption         =   "h"
                  BeginProperty Font 
                     Name            =   "宋体"
                     Size            =   9
                     Charset         =   0
                     Weight          =   700
                     Underline       =   0   'False
                     Italic          =   0   'False
                     Strikethrough   =   0   'False
                  EndProperty
                  ForeColor       =   &H80000008&
                  Height          =   180
                  Left            =   345
                  TabIndex        =   20
                  Top             =   45
                  Width           =   105
               End
            End
            Begin VB.Label lblCheck 
               Alignment       =   2  'Center
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
               ForeColor       =   &H80000008&
               Height          =   255
               Left            =   0
               TabIndex        =   54
               Top             =   0
               Visible         =   0   'False
               Width           =   255
            End
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
            ItemData        =   "frmCaseTendBodySetData.frx":08CA
            Left            =   840
            List            =   "frmCaseTendBodySetData.frx":08D7
            Style           =   1  'Checkbox
            TabIndex        =   43
            Top             =   1080
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.PictureBox PicLst 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            ForeColor       =   &H80000008&
            Height          =   1455
            Left            =   0
            ScaleHeight     =   1425
            ScaleWidth      =   1185
            TabIndex        =   40
            Top             =   675
            Visible         =   0   'False
            Width           =   1215
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
               Height          =   570
               Index           =   0
               ItemData        =   "frmCaseTendBodySetData.frx":08F0
               Left            =   -15
               List            =   "frmCaseTendBodySetData.frx":08FD
               TabIndex        =   42
               Top             =   870
               Visible         =   0   'False
               Width           =   1215
            End
            Begin VB.TextBox txtLst 
               Appearance      =   0  'Flat
               BackColor       =   &H00FFFFFF&
               Height          =   300
               Left            =   -10
               MultiLine       =   -1  'True
               TabIndex        =   41
               Top             =   270
               Width           =   1215
            End
            Begin VB.Label lbllst 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000018&
               Caption         =   "录入："
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   2
               Left            =   30
               TabIndex        =   52
               Top             =   30
               Width           =   540
            End
            Begin VB.Label lbllst 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000018&
               Caption         =   "选择："
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   180
               Index           =   3
               Left            =   15
               TabIndex        =   51
               Top             =   615
               Width           =   540
            End
         End
         Begin VB.PictureBox picColor 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1215
            Left            =   2160
            ScaleHeight     =   1215
            ScaleWidth      =   1575
            TabIndex        =   30
            TabStop         =   0   'False
            Top             =   840
            Visible         =   0   'False
            Width           =   1575
            Begin zl9TemperatureChartGD.ColorPicker usrColor 
               Height          =   2190
               Left            =   120
               TabIndex        =   31
               Top             =   -450
               Width           =   2190
               _ExtentX        =   3863
               _ExtentY        =   3863
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfTab 
            Height          =   1005
            Left            =   0
            TabIndex        =   17
            Top             =   120
            Width           =   2895
            _cx             =   5106
            _cy             =   1773
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
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483632
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483634
            FocusRect       =   2
            HighLight       =   0
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   270
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
            WallPaperAlignment=   8
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VB.Label lbllst 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   75
            Index           =   1
            Left            =   1440
            TabIndex        =   45
            Top             =   1560
            Visible         =   0   'False
            Width           =   285
         End
         Begin VB.Label lbllst 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   195
            Index           =   0
            Left            =   120
            TabIndex        =   44
            Top             =   1560
            Visible         =   0   'False
            Width           =   45
         End
      End
   End
   Begin VB.PictureBox picCurve 
      Appearance      =   0  'Flat
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   240
      ScaleHeight     =   2895
      ScaleWidth      =   7815
      TabIndex        =   14
      TabStop         =   0   'False
      Top             =   1080
      Width           =   7815
      Begin VB.Frame FraTime 
         Height          =   615
         Left            =   0
         TabIndex        =   1
         Top             =   0
         Width           =   7605
         Begin VB.PictureBox picToolBar 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   350
            Left            =   4920
            ScaleHeight     =   345
            ScaleWidth      =   2775
            TabIndex        =   32
            TabStop         =   0   'False
            Top             =   200
            Width           =   2775
            Begin VB.OptionButton OptTime 
               Caption         =   "24"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   5
               Left            =   2280
               Style           =   1  'Graphical
               TabIndex        =   39
               Top             =   0
               Width           =   350
            End
            Begin VB.OptionButton OptTime 
               Caption         =   "20"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   4
               Left            =   1920
               Style           =   1  'Graphical
               TabIndex        =   38
               Top             =   0
               Width           =   350
            End
            Begin VB.OptionButton OptTime 
               Caption         =   "16"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   3
               Left            =   1560
               Style           =   1  'Graphical
               TabIndex        =   37
               Top             =   0
               Width           =   350
            End
            Begin VB.OptionButton OptTime 
               Caption         =   "12"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   2
               Left            =   1200
               Style           =   1  'Graphical
               TabIndex        =   36
               Top             =   0
               Width           =   350
            End
            Begin VB.OptionButton OptTime 
               Caption         =   "8"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   1
               Left            =   840
               Style           =   1  'Graphical
               TabIndex        =   35
               Top             =   0
               Width           =   350
            End
            Begin VB.OptionButton OptTime 
               Caption         =   "4"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   300
               Index           =   0
               Left            =   480
               Style           =   1  'Graphical
               TabIndex        =   34
               Top             =   0
               Width           =   350
            End
            Begin VB.Label lblPtime 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "时点:"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   180
               Left            =   0
               TabIndex        =   33
               Top             =   45
               Width           =   450
            End
         End
         Begin VB.PictureBox picPre 
            Appearance      =   0  'Flat
            BackColor       =   &H80000004&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   375
            Left            =   3400
            ScaleHeight     =   375
            ScaleWidth      =   1500
            TabIndex        =   4
            TabStop         =   0   'False
            Top             =   150
            Width           =   1500
            Begin VB.PictureBox picBut 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000004&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   360
               Index           =   3
               Left            =   1080
               Picture         =   "frmCaseTendBodySetData.frx":0916
               ScaleHeight     =   360
               ScaleWidth      =   360
               TabIndex        =   8
               TabStop         =   0   'False
               Top             =   10
               Width           =   360
            End
            Begin VB.PictureBox picBut 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000004&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   360
               Index           =   2
               Left            =   720
               Picture         =   "frmCaseTendBodySetData.frx":0B20
               ScaleHeight     =   360
               ScaleWidth      =   360
               TabIndex        =   7
               TabStop         =   0   'False
               Top             =   0
               Width           =   360
            End
            Begin VB.PictureBox picBut 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000004&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   360
               Index           =   1
               Left            =   360
               Picture         =   "frmCaseTendBodySetData.frx":0D2A
               ScaleHeight     =   360
               ScaleWidth      =   360
               TabIndex        =   6
               TabStop         =   0   'False
               Top             =   0
               Width           =   360
            End
            Begin VB.PictureBox picBut 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000004&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   360
               Index           =   0
               Left            =   0
               Picture         =   "frmCaseTendBodySetData.frx":0F34
               ScaleHeight     =   360
               ScaleWidth      =   360
               TabIndex        =   5
               TabStop         =   0   'False
               Top             =   0
               Width           =   360
            End
            Begin VB.PictureBox picBut1 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000004&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   360
               Index           =   3
               Left            =   1080
               Picture         =   "frmCaseTendBodySetData.frx":113E
               ScaleHeight     =   360
               ScaleWidth      =   360
               TabIndex        =   27
               TabStop         =   0   'False
               Top             =   10
               Visible         =   0   'False
               Width           =   360
            End
            Begin VB.PictureBox picBut1 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000004&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   360
               Index           =   2
               Left            =   720
               Picture         =   "frmCaseTendBodySetData.frx":1348
               ScaleHeight     =   360
               ScaleWidth      =   360
               TabIndex        =   26
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   360
            End
            Begin VB.PictureBox picBut1 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000004&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   360
               Index           =   1
               Left            =   360
               Picture         =   "frmCaseTendBodySetData.frx":1552
               ScaleHeight     =   360
               ScaleWidth      =   360
               TabIndex        =   25
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   360
            End
            Begin VB.PictureBox picBut1 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000004&
               BorderStyle     =   0  'None
               ForeColor       =   &H80000008&
               Height          =   360
               Index           =   0
               Left            =   0
               Picture         =   "frmCaseTendBodySetData.frx":175C
               ScaleHeight     =   360
               ScaleWidth      =   360
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   0
               Visible         =   0   'False
               Width           =   360
            End
         End
         Begin MSComCtl2.DTPicker dkpTime 
            Height          =   300
            Left            =   1440
            TabIndex        =   3
            Top             =   210
            Width           =   825
            _ExtentX        =   1455
            _ExtentY        =   529
            _Version        =   393216
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            CustomFormat    =   "HH:mm"
            Format          =   129826819
            UpDown          =   -1  'True
            CurrentDate     =   40568
         End
         Begin VB.Label lblTime 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "00:00～05:59"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   240
            TabIndex        =   2
            Top             =   250
            Width           =   1080
         End
      End
      Begin VB.Frame FraData 
         BeginProperty Font 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1965
         Left            =   0
         TabIndex        =   9
         Top             =   620
         Width           =   5700
         Begin VB.PictureBox picValue 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1455
            Left            =   480
            ScaleHeight     =   1455
            ScaleWidth      =   1575
            TabIndex        =   47
            TabStop         =   0   'False
            Top             =   480
            Visible         =   0   'False
            Width           =   1575
            Begin zl9TemperatureChartGD.ColorPicker usrValue 
               Height          =   2190
               Left            =   120
               TabIndex        =   48
               Top             =   -360
               Width           =   2190
               _ExtentX        =   3863
               _ExtentY        =   3863
            End
         End
         Begin VB.PictureBox pic未记 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   930
            Left            =   2160
            ScaleHeight     =   930
            ScaleWidth      =   1215
            TabIndex        =   11
            TabStop         =   0   'False
            Top             =   840
            Visible         =   0   'False
            Width           =   1215
            Begin VB.ListBox lst未记 
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
               ItemData        =   "frmCaseTendBodySetData.frx":1966
               Left            =   0
               List            =   "frmCaseTendBodySetData.frx":1970
               TabIndex        =   12
               Top             =   0
               Visible         =   0   'False
               Width           =   1215
            End
         End
         Begin zl9TemperatureChartGD.VsfGrid vsfCurve 
            Height          =   1215
            Left            =   120
            TabIndex        =   10
            Top             =   240
            Width           =   3975
            _ExtentX        =   7011
            _ExtentY        =   2143
         End
      End
   End
   Begin XtremeSuiteControls.TabControl tbcThis 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   5295
      _Version        =   589884
      _ExtentX        =   9340
      _ExtentY        =   4895
      _StockProps     =   64
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   13
      Top             =   5220
      Width           =   8910
      _ExtentX        =   15716
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCaseTendBodySetData.frx":1980
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "欢迎使用中联有限公司软件"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12806
            Key             =   "ZLNOTE"
            Object.ToolTipText     =   "消息提示信息"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   2
            MinWidth        =   2
            Text            =   "数据类型"
            TextSave        =   "数据类型"
            Key             =   "ZLDataType"
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
      Left            =   6120
      Top             =   360
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
            Picture         =   "frmCaseTendBodySetData.frx":2214
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComCtl2.DTPicker dkpDate 
      Height          =   300
      Left            =   2280
      TabIndex        =   23
      Top             =   240
      Width           =   1335
      _ExtentX        =   2355
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "yyyy-MM-dd"
      Format          =   129826819
      UpDown          =   -1  'True
      CurrentDate     =   40619
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   90
      Top             =   60
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmCaseTendBodySetData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum TYPE_Curve
    COL_Null = 0
    COL_分组名 = 1
    COL_字符串 = 2
    COL_项目序号 = 3
    COL_项目名称 = 4
    col_数据 = 5
    col_颜色 = 6
    col_复查 = 7
    COL_部位 = 8
    Col_未记说明 = 9
End Enum

Private Enum TYPE_Tab
    COL_tab分组名 = 0
    COL_tab字符串 = 1
    COL_tab项目序号 = 2
    COL_TabNull = 3
    COL_tab项目名称 = 4
End Enum

Private Enum TYPE_Oper
    Col_OperNull = 0
    Col_OperTime = 1
    Col_OperType = 2
End Enum

Private Enum Enum_No
     Item体温 = 1
     Item脉搏 = 2
     Item心率 = -1
     Item收缩压 = 4
     Item舒张压 = 5
End Enum

Private Type Type_Item
    类型 As String
    值域 As String
    项目类型 As Integer
    项目小数 As Double
    记录频次 As Integer
    项目表示 As Integer
    项目性质 As Integer
    项目长度 As Long
    部位 As String
    项目号 As Long
    项目名 As String
    入院首测 As Integer
End Type

Private Type type_Patient
    lng病人ID As Long
    lng主页ID As Long
    lng文件ID As Long
    lng婴儿 As Long
    lng科室ID As Long
    lng护理等级 As Long
    lng病区ID As Long
End Type
Private mT_Patient As type_Patient

'手术上下标列
Private Type Type_OptRow
    手术 As Integer
    上标 As Integer
    下标 As Integer
End Type

Private mOptRow As Type_OptRow

'工具栏:
Private mcbrToolBar As CommandBar

Private mblnStart As Boolean
Private mblnFileBack As Boolean
Private mblnScroll As Boolean
Private mblnEdit As Boolean
Private mblnAllRefresh As Boolean
Private marrTime() As String
Private Const mFontSize As Integer = 9 '定义字体初始大小为9号字体
Private mintPreDays As Integer '超期录入时限
Private mintBigSize As Integer '是否放大
Private mlngHours As Long '数据补录时限
Private mbln汇总当天 As Boolean
Private mbln录入小时 As Boolean  '全天汇总显示录入时间
Private mstrActiveItem As String
Private mint心率应用 As Integer
Private mblnEdit心率 As Boolean
Private mstrBegin As String '某段时间点的开始和结束时间 00:00-05:59
Private mstrEnd As String
Private mstrBTime As String  '体温单的开始时间和结束时间
Private mstrETime As String
Private mstrOverDate As String '病人实际出院时间(即体温单实际终止时间)
Private mstrDate As String '体温单当前页的第一天时间
Private mblnChage As Boolean
Private mblnCurveChange As Boolean
Private mblnOK As Boolean
Private mblnMove As Boolean
Private mstrSQL As String
Private mblnInit As Boolean
Private mstr未记说明 As String
Private mArrdkpTime() As Variant
Private mArrModfy() As Integer
Private mArrValue() As String
Private marrDate() As Integer
Private mstrPart As String
Private mbln出院 As Boolean
Private mblnResize As Boolean
Private mbln脉搏共用显示 As Boolean

'记录集
Private mrsPart As New ADODB.Recordset '体温部位
Private mrsCurve As New ADODB.Recordset '体温数据
Private mrsNote As New ADODB.Recordset '上下标
Private mrsOper As New ADODB.Recordset '手术
Private mrsRecodeID As New ADODB.Recordset '记录体温曲线项目的记录ID和时间

Public Function ShowEditor(ByVal frmParent As Object, ByVal strParam As String, ByVal strTime As String, ByVal strDayTime As String, _
    ByVal int心率应用 As Integer, Optional blnMove As Boolean = False, Optional ByVal bytSize As Byte = 0) As Boolean
'----------------------------------------------------------------------------------------------------------------------------------------------------------
'功能:调用体温单编辑窗体
'参数:frmParent 父窗体,strParam 格式:病人ID;主页Id;文件ID;婴儿;科室ID;护理护理等级  strTime 某段时间的时间范围 例如:2011-01-25 00:00:00;2011-01-25 05:59:59

'     strDayTime 一周开始时间; int心率应用=2 表示脉搏和心率公用 blnMove 历史数据是否转移
'     bytSize 0-9号字体 1-12号字体
'----------------------------------------------------------------------------------------------------------------------------------------------------------
    Dim arrParam() As String
    Dim blnShowing As Boolean
    
    If strParam = "" Then Exit Function
    arrParam = Split(strParam, ";")
    If UBound(arrParam) < 3 Then Exit Function
        
    mblnStart = True
    mblnChage = False
    mblnCurveChange = False
    mblnMove = False
    mblnInit = False
    mblnEdit = False
    mblnOK = False
    mblnResize = False
    mblnAllRefresh = False
    mbln汇总当天 = False
    mstrOverDate = ""
    
    mT_Patient.lng科室ID = 0
    mT_Patient.lng护理等级 = 3
    
    mT_Patient.lng病人ID = Val(arrParam(0))
    mT_Patient.lng主页ID = Val(arrParam(1))
    mT_Patient.lng文件ID = Val(arrParam(2))
    mT_Patient.lng婴儿 = Val(arrParam(3))
    
    If UBound(arrParam) > 3 Then mT_Patient.lng科室ID = arrParam(4)
    If UBound(arrParam) > 4 Then mT_Patient.lng护理等级 = arrParam(5)
    
    If mT_Patient.lng病人ID = 0 And mT_Patient.lng主页ID = 0 And mT_Patient.lng科室ID = 0 Then
        MsgBox "文件ID,病人ID,主页ID不能为空,请检查!", vbInformation, gstrSysName
        Exit Function
    End If
    
    mstrBegin = Format(Split(strTime, ";")(0), "YYYY-MM-DD HH:mm:ss")
    mstrEnd = Format(Split(strTime, ";")(1), "YYYY-MM-DD HH:mm:ss")
    mstrDate = strDayTime
    
    If Not ChekPatientOut(mT_Patient.lng文件ID, mT_Patient.lng病人ID, mT_Patient.lng主页ID, mT_Patient.lng婴儿) Then Exit Function
    mintBigSize = bytSize
    Me.Font.Size = IIf(mintBigSize = 0, 9, 12)
    mint心率应用 = int心率应用
    mblnEdit心率 = True
    mblnMove = blnMove
    
    If Not OpenPatientInfo Then Exit Function
    
    '检查文件是否归档
    mblnFileBack = CheckFileBack(mT_Patient.lng文件ID, mblnMove)
    '初始化工具栏
    Call InitCommandBars
    '提取数据
    Call GetTableRowName
    Call zlRefreshData
    mblnInit = True
    mblnResize = True
    Me.Show 1
    
    ShowEditor = mblnOK
End Function

Public Function ChekPatientOut(ByVal lng文件ID As Long, ByVal lng病人ID As Long, ByVal lng主页ID As Long, ByVal intBaby As Long) As Boolean
'-----------------------------------------------------------------------------------------------
'功能:提取体温单开始时间和结束时间 并检查病人是否出院
'-----------------------------------------------------------------------------------------------
    Dim strSQL As String, strNewSql As String
    Dim strBeginDate As String, strEndDate As String
    Dim rsTemp As New ADODB.Recordset
    Dim strMaxDate As String, strCurrDate As String
    Dim intDay As Integer
    
    mbln出院 = False
    On Error GoTo Errhand
    
    'mintBigSize =  zldatabase.GetPara("护理文件显示模式", glngSys, 1255, 0)
    mintPreDays = Val(zlDatabase.GetPara("超期录入护理数据天数", glngSys, 1255, "1"))
    gintHourBegin = zlDatabase.GetPara("体温开始时间", glngSys, 1255, 4)
    mlngHours = Val(Mid(Val(zlDatabase.GetPara("数据补录时限", glngSys)), 1, 6))
    mbln汇总当天 = (Val(zlDatabase.GetPara("汇总波动显示当天数据", glngSys, 1255, 0)) = 1)
    '51282,刘鹏飞,2012-08-03,全天汇总显示录入时间(DYEY要求手工录入汇总时间H
    mbln录入小时 = (Val(zlDatabase.GetPara("全天汇总显示录入时间", glngSys, 1255, 0)) = 1)
    mbln脉搏共用显示 = (Val(zlDatabase.GetPara("脉搏短绌以(心率/脉搏)方式录入", glngSys, 1255, 0)) = 1)
    If mintPreDays < 0 Then mintPreDays = 0
        
    '提取婴儿医嘱信息(转科，出院),存在医嘱以医嘱信息为准，否则以母亲出院日期为准
    strNewSql = "   (SELECT /*+ RULE */  病人ID,主页ID,婴儿时间,DECODE(nvl(婴儿,0),0, DECODE(NVL(出院日期,''),'',0,1), DECODE(NVL(婴儿时间,''),'',0,1))记录" & vbNewLine & _
                "       FROM (SELECT A.病人ID,A.主页ID,B.开始执行时间 婴儿时间, A.出院日期,B.婴儿" & vbNewLine & _
                "           FROM 病案主页 A," & vbNewLine & _
                "               (SELECT B.病人ID, B.主页ID, B.婴儿, 开始执行时间" & vbNewLine & _
                "                FROM 病人医嘱记录 B, 诊疗项目目录 C" & vbNewLine & _
                "                WHERE B.诊疗项目ID + 0 = C.ID AND B.医嘱状态 = 8 AND nvl(B.婴儿,0)<>0  AND C.类别 = 'Z'" & vbNewLine & _
                "                AND EXISTS (SELECT 1 FROM TABLE(CAST(F_STR2LIST('3,5,11') AS ZLTOOLS.T_STRLIST))" & vbNewLine & _
                "                               WHERE C.操作类型 = COLUMN_VALUE) And  B.病人ID = [2] AND B.主页ID = [3] AND B.婴儿(+) = [4]) B" & vbNewLine & _
                "           WHERE A.病人ID = [2] AND A.主页ID = [3] AND A.病人ID = B.病人ID(+) AND A.主页ID = B.主页ID(+)" & vbNewLine & _
                "           ORDER BY B.开始执行时间 DESC)" & vbNewLine & _
                "       WHERE ROWNUM < 2)  E"
    
    strSQL = _
       "Select Decode(b.出生时间,Null,a.开始,b.出生时间) As 开始,decode(E.记录,0,Decode(Sign(NVL(E.婴儿时间,a.终止) - d.发生时间), 1,NVL(E.婴儿时间,a.终止) ,d.发生时间),NVL(E.婴儿时间,a.终止)) 终止,E.记录" & vbNewLine & _
        "       From" & vbNewLine & _
        "       (Select 病人ID,主页id,Min(开始时间) as 开始,Max(Nvl(终止时间,sysdate)) as 终止" & vbNewLine & _
        "       From 病人变动记录" & vbNewLine & _
        "       Where 开始时间 is Not Null And 病人ID=[2] And 主页ID=[3] Group By 病人ID,主页id) a," & vbNewLine & _
        "       (Select 病人ID,主页id,出生时间 From 病人新生儿记录 Where 病人ID =[2] And 主页ID =[3] And 序号=[4]) b," & vbNewLine & _
        "       (SELECT NVL(发生时间,SYSDATE) 发生时间 FROM (select max(发生时间) 发生时间 from 病人护理文件 A,病人护理数据 B" & vbNewLine & _
        "       where A.ID=B.文件ID and A.ID=[1] and A.病人ID=[2] and A.主页ID=[3] and A.婴儿=[4])) d," & vbNewLine & _
        strNewSql & vbNewLine & _
        "       Where a.病人ID=E.病人ID And A.主页ID=E.主页ID And a.病人id=b.病人id(+) And a.主页id=b.主页id(+)"
        
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlPrint", lng文件ID, lng病人ID, lng主页ID, intBaby)
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        strBeginDate = Format(rsTemp!开始, "YYYY-MM-DD HH:MM:SS")
        strEndDate = Format(rsTemp!终止, "YYYY-MM-DD HH:MM:SS")
        mbln出院 = Not (Val(rsTemp!记录) = 0)
    Else
        MsgBox "无此病人本次住院信息,请检查!", vbInformation, gstrSysName
        Exit Function '无数病人变动信息退出
    End If
    
    '提取用户设置的体温单开始时间(婴儿以出生时间为准)
    If intBaby = 0 Then
        strSQL = "select 开始时间 from 病人护理文件 where ID=[1] and 病人ID=[2] and 主页id=[3] and nvl(婴儿,0)=[4]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "提取体温单开始时间", lng文件ID, lng病人ID, lng主页ID, intBaby)
        If rsTemp.RecordCount <> 0 Then
            strBeginDate = Format(rsTemp!开始时间, "YYYY-MM-DD HH:mm:ss")
        End If
    End If
    
    strCurrDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")

    mstrBTime = strBeginDate
    mstrOverDate = strEndDate
    mstrETime = strEndDate
    If CDate(mstrETime) < CDate(Format(strCurrDate, "YYYY-MM-DD HH:mm:ss")) And Not mbln出院 Then mstrETime = CDate(Format(strCurrDate, "YYYY-MM-DD HH:mm:ss"))
    If mstrBTime > mstrETime Then mstrBTime = mstrETime
    If mstrDate < mstrBTime Then mstrDate = mstrBTime
    
    '病人出院已出院时间为终止时间
    If mbln出院 = True Then
        '出院时间和入院时间如果在同一列，则将出院时间后移一列（内蒙需求:出院也要录入体温）
        mstrETime = Format(RetrunEndTime(CDate(mstrBTime), CDate(mstrETime), gintHourBegin), "YYYY-MM-DD HH:mm:ss")
        strMaxDate = Format(mstrETime, "YYYY-MM-DD")
    Else
        intDay = mintPreDays - DateDiff("D", CDate(strCurrDate), CDate(mstrETime))
        If intDay < 0 Then intDay = 0
        strMaxDate = Format(DateAdd("d", intDay, CDate(mstrETime)), "yyyy-MM-dd")
        If CDate(mstrETime) < CDate(Format(DateAdd("d", mintPreDays, CDate(strCurrDate)), "yyyy-MM-dd HH:mm:ss")) Then
            mstrETime = Format(DateAdd("d", mintPreDays, CDate(strCurrDate)), "yyyy-MM-dd HH:mm:ss")
        End If
    End If
    
    mstrETime = Format(strMaxDate & " " & Format(mstrETime, "HH:mm:ss"), "yyyy-MM-DD HH:mm:ss")
    
    dkpDate.Value = Format(mstrBegin, "YYYY-MM-DD")
    dkpDate.MaxDate = Format(strMaxDate, "YYYY-MM-DD")
    dkpDate.MinDate = Format(mstrBTime, "YYYY-MM-DD")
    
    If CDate(mstrBegin) < CDate(mstrBTime) Then
        mstrBegin = Format(mstrBTime, "YYYY-MM-DD HH:mm:ss")
    End If
    
    If CDate(mstrEnd) > CDate(mstrETime) Then
        mstrEnd = Format(mstrETime, "YYYY-MM-DD HH:mm:ss")
    End If
    
    ChekPatientOut = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function OpenPatientInfo() As Boolean
    Dim rsTmp As New ADODB.Recordset
    On Error GoTo Errhand
    '提取科室信息
    mstrSQL = "Select 出院科室ID from 病案主页 Where 病人id=[1] And 主页id=[2] "
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, mT_Patient.lng病人ID, mT_Patient.lng主页ID)
    If rsTmp.BOF = False Then
        mT_Patient.lng科室ID = Val(zlCommFun.Nvl(rsTmp("出院科室ID").Value))
    End If
    
    '提取护理等级
    mstrSQL = "Select zl_PatitTendGrade([1],[2]) As 护理等级 From dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, mT_Patient.lng病人ID, mT_Patient.lng主页ID)
    If rsTmp.BOF = False Then mT_Patient.lng护理等级 = zlCommFun.Nvl(rsTmp("护理等级"), 3)
    
    OpenPatientInfo = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub InitCommandBars()
'--------------------------------------------------------------------------------
'功能:初始化工具栏
'--------------------------------------------------------------------------------
    Dim cbrControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim cbrLable As CommandBarControl
    Dim cbrPop As CommandBarControl
    Dim CtlFont As StdFont
    
    On Error GoTo Errhand
    
     '初始设置
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbsMain.ActiveMenuBar.Title = "菜单栏"
    cbsMain.ActiveMenuBar.Visible = False
    
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize False, 24, 24
        .SetIconSize True, 16, 16
        .UseSharedImageList = False 'ImageList方式时,因同一App中共享,在AddImageList之前设置为False
        Set CtlFont = .Font
        If CtlFont Is Nothing Then
            Set CtlFont = Me.Font
        End If
        CtlFont.Size = IIf(mintBigSize = 0, 9, 12)
        Set .Font = CtlFont
    End With

  '------------------------------------------------------------------------------------------------------------------
    '工具栏定义
    Set mcbrToolBar = cbsMain.Add("标准", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    mcbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    
    With mcbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "保存")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "取消")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Curve_Show, "曲线显示"): cbrControl.ToolTipText = "设置曲线数据显示"
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "添加项目"): cbrControl.ToolTipText = "添加活动项目": cbrControl.BeginGroup = True

        Set cbrPop = .Add(xtpControlButtonPopup, conMenu_Edit_Append, "特殊处理")
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 0, "正常", -1, False): cbrControl.IconId = 1: cbrControl.Parameter = ""
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 1, "灌肠[E]", -1, False):  cbrControl.IconId = 1: cbrControl.Parameter = "E"
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 2, "灌肠后大便[/E]", -1, False):  cbrControl.IconId = 1: cbrControl.Parameter = "/E"
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 3, "大便失禁[※]", -1, False): cbrControl.IconId = 1: cbrControl.Parameter = "※"
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 4, "人工肛门[☆]", -1, False): cbrControl.IconId = 1: cbrControl.Parameter = "☆"
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 5, "导尿[C]", -1, False):   cbrControl.IconId = 1: cbrControl.Parameter = "C"
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 6, "保留导尿[/C]", -1, False):   cbrControl.IconId = 1: cbrControl.Parameter = "/C"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "帮助")
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "退出")
        
    End With
    
    '定位工具栏
    '------------------------------------------------------------------------------------------------------------------
    For Each cbrControl In mcbrToolBar.Controls
        If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel Then
            cbrControl.Style = xtpButtonIconAndCaption
        End If
    Next
    
    With dkpDate
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Width = .Width + .Width * mintBigSize / 3
        .Height = 300 + 300 * mintBigSize / 3
    End With
    
    '超期补录
    '------------------------------------------------------------------------------------------------------------------
    Set cbrLable = mcbrToolBar.Controls.Add(xtpControlLabel, conMenu_View_Option, "时间")
    cbrLable.flags = xtpFlagRightAlign
    
    Set cbrCustom = mcbrToolBar.Controls.Add(xtpControlCustom, conMenu_View_Location, "")
    dkpDate.Visible = True
    cbrCustom.Handle = dkpDate.hWnd
    cbrCustom.flags = xtpFlagRightAlign
    
    '快键绑定
    With cbsMain.KeyBindings
        .Add FALT, Asc("0"), conMenu_Edit_Append * 10
        .Add FALT, Asc("1"), (conMenu_Edit_Append * 10 + 1)
        .Add FALT, Asc("2"), (conMenu_Edit_Append * 10 + 2)
        .Add FALT, Asc("3"), (conMenu_Edit_Append * 10 + 3)
        .Add FALT, Asc("4"), (conMenu_Edit_Append * 10 + 4)
        .Add FALT, Asc("5"), (conMenu_Edit_Append * 10 + 5)
        .Add FALT, Asc("6"), (conMenu_Edit_Append * 10 + 6)
        
        .Add FCONTROL, Asc("D"), conMenu_Edit_Curve_Show '设置显示
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem '添加活动项目
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save '保存
        .Add FCONTROL, Asc("R"), conMenu_Edit_Reuse '取消
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    Call InitDateTimeRange(marrTime, gintHourBegin)
     
    '加载表格控件
    Call InitTabControl
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitTabControl()
'--------------------------------------------------------------------------------
'功能:初始化TabControl
'--------------------------------------------------------------------------------
    On Error GoTo Errhand
    Dim tabItem As TabControlItem
    Dim CtlFont As StdFont
    
    With tbcThis
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .ShowIcons = True
            .OneNoteColors = True
            .Position = xtpTabPositionTop
            .ClientFrame = xtpTabFrameSingleLine
            .DisableLunaColors = False
            .Layout = xtpTabLayoutAutoSize
            Set CtlFont = .Font
            If CtlFont Is Nothing Then
                Set CtlFont = Me.Font
            End If
            CtlFont.Size = IIf(mintBigSize = 0, 9, 12)
            Set .Font = CtlFont
        End With
        
        Set tabItem = .InsertItem(1, "体温曲线", picCurve.hWnd, 0)
        tabItem.Tag = "曲线"
        Set tabItem = .InsertItem(2, "体温表格", picTab.hWnd, 0)
        tabItem.Tag = "表格"
        If gintEditorCurveState = 0 Then
            .Item(0).Selected = True
        Else
            .Item(1).Selected = True
        End If
    End With
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub SetColSelect(Optional blnInit As Boolean = False)
'-------------------------------------
'功能:设置表格选择列
'------------------------------------
    Dim intRow As Integer, intCOl As Integer
    Dim intOldRow As Integer, intOldCol As Integer
    
    If mblnInit = False Then Exit Sub
    
    If tbcThis.Selected.Tag = "曲线" Then
        vsfCurve.SetFocus
        If blnInit = True Then
            intOldRow = vsfCurve.Row
            intOldCol = vsfCurve.Col
            intRow = vsfCurve.Row
            intCOl = col_数据
            If intRow = vsfCurve.Row And intCOl = vsfCurve.Col Then
                vsfCurve.Col = COL_部位
            End If
            vsfCurve.Col = col_数据
        Else
            intOldRow = vsfCurve.Row
            intOldCol = vsfCurve.Col
            intRow = vsfCurve.Row
            intCOl = vsfCurve.Col
            If intRow = vsfCurve.Row And intCOl = vsfCurve.Col Then
                If intCOl < vsfCurve.Cols - 1 Then
                    vsfCurve.Col = intCOl + 1
                Else
                    If intRow < vsfCurve.Rows - 1 Then
                        vsfCurve.Row = intRow + 1
                    Else
                        If intRow - 1 > 0 Then
                            vsfCurve.Row = intRow - 1
                        End If
                    End If
                End If
            End If
            vsfCurve.Col = intCOl
        End If
        Call vsfCurve_AfterRowColChange(intOldRow, intOldCol, intRow, intCOl)
    ElseIf tbcThis.Selected.Tag = "表格" Then
        vsfTab.SetFocus
        If blnInit = True Then
            intOldRow = vsfTab.Row
            intOldCol = vsfTab.Col
            intRow = vsfTab.FixedRows
            intCOl = vsfTab.FixedCols
            If intRow = vsfTab.Row And intCOl = vsfTab.Col Then
                Call vsfTab_BeforeRowColChange(intRow, intCOl, intRow, intCOl, False)
            End If
            vsfTab.Select vsfTab.FixedRows, vsfTab.FixedCols
        Else
            intOldRow = vsfTab.Row
            intOldCol = vsfTab.Col
            intRow = vsfTab.Row
            intCOl = vsfTab.Col
            vsfTab.Select vsfTab.Row, vsfTab.Col
        End If
        Call vsfTab_AfterRowColChange(intOldRow, intOldCol, intRow, intCOl)
    End If
End Sub

Private Sub InitTabCurve(ByVal strTabName As String)
'-------------------------------------------------------
'功能:初始化体温曲线项目
'参数:所有表头的信息
'-------------------------------------------------------
    Dim varTabName() As String, varCode() As String
    Dim intRow As Integer, intCOl As Integer
    
    If strTabName = "" Then Exit Sub
    varTabName = Split(strTabName, ";")
    
    With vsfCurve
        .Rows = UBound(varTabName) + 2
        .Cols = 0
        
        .NewColumn "", 255, 4
        .NewColumn "分组名", 1500 + 1500 * mintBigSize / 3, 1
        .NewColumn "字符串", 0, 1
        .NewColumn "项目序号", 0, 1
        .NewColumn "项目名称", 1200 + 1200 * mintBigSize / 3, 1
        .NewColumn "数据", 2300 + 2300 * mintBigSize / 3, 1, , 4
        .NewColumn "数据", 300 + 300 * mintBigSize / 3, 0
        .NewColumn "复试合格", 900 + 900 * mintBigSize / 3, 4
        .NewColumn "部位", 1000 + 1000 * mintBigSize / 3, 4
        .NewColumn "未记说明", 1080 + 1080 * mintBigSize / 3, 4, "...", 1
        .Body.RowHeight(0) = 300 + 300 * mintBigSize / 3
        .FixedCols = 5
        .FixedRows = 1
        
        .Body.Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Body.ColHidden(COL_字符串) = True
        .Body.ColHidden(COL_项目序号) = True
        .Body.WordWrap = True
        .Body.MergeCells = flexMergeRestrictColumns
        .Body.MergeCol(COL_分组名) = True
        .Body.MergeRow(0) = True
        
        For intRow = .FixedRows To .Rows - 1
            varCode = Split(varTabName(intRow - 1), "'")
            If UBound(varCode) > 2 Then
                .TextMatrix(intRow, COL_分组名) = varCode(0)
                .TextMatrix(intRow, COL_字符串) = varCode(1)
                .TextMatrix(intRow, COL_项目序号) = varCode(2)
                .TextMatrix(intRow, COL_项目名称) = varCode(3)
                If varCode(0) = "2)上下标说明" Then
                    Select Case Val(varCode(2))
                        Case 2
                            mOptRow.上标 = intRow
                        Case 4
                            mOptRow.手术 = intRow
                        Case 6
                            mOptRow.下标 = intRow
                    End Select
                End If
            End If
            .Body.RowHeight(intRow) = 300 + 300 * mintBigSize / 3
            .RowData(intRow) = 0
        Next intRow

        .Cell(flexcpAlignment, 0, .FixedCols, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .Cell(flexcpText, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = ""
    End With
End Sub

Private Sub InitTabOper()
'-------------------------------------------------------
'功能:初始化手术分娩录入表格
'-------------------------------------------------------
    Dim intRow As Integer, intCOl As Integer
        
    With vsfOper
        .Rows = 2
        .Cols = 0
        
        .NewColumn "", 255, 4
        .NewColumn "时间", 1000 + 1000 * mintBigSize / 3, 4, , 4
        .NewColumn "数据", 2000 + 2000 * mintBigSize / 3, 4, "手术|分娩|手术分娩", 1
        .NewColumn "", 255, 4
        .ExtendLastCol = True
        .Body.RowHeightMin = 300 + 300 * mintBigSize / 3
        .FixedCols = 1
        .FixedRows = 1
        
        .Body.Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Body.WordWrap = False
        .Body.AllowUserResizing = flexResizeNone

        .Cell(flexcpAlignment, 0, .FixedCols, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .Cell(flexcpText, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = ""
    End With
End Sub

Private Sub InitTabTable(ByVal strTabName As String)
'-------------------------------------------------------
'功能:初始化体温表格项目
'参数:所有表头的信息(不包含汇总项目)
'-------------------------------------------------------
    Dim varTabName() As String, varCode() As String
    Dim intRow As Integer, intCOl As Integer
    
    If strTabName = "" Then Exit Sub
    varTabName = Split(strTabName, ";")
    
    With vsfTab
        .Rows = UBound(varTabName) + 2
        .Cols = 11
        
        .FixedCols = 5
        .FixedRows = 1
        
        .ColWidth(3) = 255
        .ColAlignment(3) = 4
        
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .ColHidden(COL_tab分组名) = True
        .ColHidden(COL_tab字符串) = True
        .ColHidden(COL_tab项目序号) = True
        .WordWrap = True
        .ScrollBars = flexScrollBarBoth
        .SelectionMode = flexSelectionByRow
        
        '初始表头
        For intCOl = .FixedCols - 1 To .Cols - 1
            If intCOl = .FixedCols - 1 Then
                .TextMatrix(0, intCOl) = "名称/频次"
            Else
                .TextMatrix(0, intCOl) = intCOl - .FixedCols + 1
                .ColWidth(intCOl) = 1200 + 1200 * mintBigSize / 3
            End If
        Next intCOl
        
        For intRow = 1 To .Rows - 1
            varCode = Split(varTabName(intRow - 1), "'")
            .TextMatrix(intRow, COL_tab分组名) = varCode(0)
            .TextMatrix(intRow, COL_tab字符串) = varCode(1)
            .TextMatrix(intRow, COL_tab项目序号) = varCode(2)
            .TextMatrix(intRow, COL_TabNull) = ""
            .TextMatrix(intRow, COL_tab项目名称) = varCode(3)
        Next intRow
        
        .ColWidth(COL_tab项目名称) = 1200 + 1200 * mintBigSize / 3
        .RowHeight(-1) = 300 + 300 * mintBigSize / 3
                
        .Cell(flexcpAlignment, 0, .FixedCols, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .Cell(flexcpText, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = ""
    End With
End Sub

Private Sub AdjustRowFlag(ByRef objVsf As Object, ByVal intRow As Integer)
    '-----------------------------------------------------------------------------------------
    '功能:
    '参数:
    '-----------------------------------------------------------------------------------------
    If objVsf.FixedCols = 0 Then Exit Sub
    
    If Not (objVsf.Cell(flexcpPicture, intRow, COL_TabNull) Is Nothing) Then Exit Sub
    Set objVsf.Cell(flexcpPicture, 1, COL_TabNull, objVsf.Rows - 1, COL_TabNull) = Nothing
    Set objVsf.Cell(flexcpPicture, intRow, COL_TabNull) = ils16.ListImages(1).Picture
    
End Sub

Private Function InitTime() As String
'--------------------------------------------------------
'功能:提取一天的时间段信息
'--------------------------------------------------------
    Dim i As Integer
    Dim strName As String
    
    Call InitDateTimeRange(marrTime, gintHourBegin)
    For i = 0 To UBound(marrTime) - 1
        strName = strName & ";" & Format(Split(marrTime(i), ",")(0), "HH:mm") & "～" & Format(Split(marrTime(i), ",")(1), "HH:mm")
    Next i
    
    If Left(strName, 1) = ";" Then strName = Mid(strName, 2)
    
    strName = "项目\时间范围" & ";" & strName
    InitTime = strName
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRow As Long, lngCol As Long, lngItemNO As Long, intType As Integer, i As Integer
    Dim strValue As String, strPart As String, strPart1 As String, strName As String
    Dim strTime As String, strErrMsg As String
    Dim strTmp As String, arrStr
    Dim cbrCheck As CommandBarControl
    
    Select Case Control.Id
    
        Case conMenu_Edit_Save '保存
        
            If picEdit.Visible = True Then
                Call vsfTab_EnterCell
            End If
            If Not ChangeCurveTime Then Exit Sub
            If Not SaveData Then Exit Sub
            Call GetTableRowName
            Call zlRefreshData
            Call SetColSelect
            
        Case conMenu_Edit_Reuse '取消
            Call GetTableRowName
            Call zlRefreshData
            mblnChage = False
            mblnCurveChange = False
            Call txtEdit_KeyPress(vbKeyEscape)
            Call SetColSelect
            
        Case conMenu_Edit_NewItem '添加活动项目
            Call txtEdit_KeyPress(vbKeyEscape)
            mblnScroll = True
            If frmCaseTendBodyActiveItem.ShowMe(vsfTab, Me) Then
                vsfTab.Refresh
            End If
        Case conMenu_Edit_Curve_Show '设置显示
            If mblnChage Then
                If MsgBox("数据已经发生改变,请问是否需要保存?", vbInformation + vbDefaultButton1 + vbYesNo, gstrSysName) = vbYes Then
                    If Not ChangeCurveTime Then Exit Sub
                    If Not SaveData Then Exit Sub
                End If
            End If
            '调用显示窗体
            Call gobjTendEditor.BodyEditCur(1, Format(mstrBegin, "YYYY-MM-DD"))
            
        Case conMenu_Edit_Append * 10, conMenu_Edit_Append * 10 + 1, conMenu_Edit_Append * 10 + 2, conMenu_Edit_Append * 10 + 3, conMenu_Edit_Append * 10 + 4, conMenu_Edit_Append * 10 + 5, conMenu_Edit_Append * 10 + 6
            If vsfTab.Row < vsfTab.FixedRows Or vsfTab.Col < vsfTab.FixedCols Then Exit Sub
            lngRow = vsfTab.Row
            lngCol = vsfTab.Col
            lngItemNO = Val(vsfTab.TextMatrix(lngRow, COL_tab项目序号))
            strName = vsfTab.TextMatrix(lngRow, COL_tab项目名称)
            strValue = Trim(vsfTab.TextMatrix(lngRow, lngCol))
            strTmp = vsfTab.TextMatrix(lngRow, COL_tab字符串)
            strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
            arrStr = Split(strTmp, ",")
            
            intType = 0
            If picEdit.Visible = True And txtEdit.Visible = True Then intType = 1
            If intType = 1 Then strValue = txtEdit.Text
            
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
            
            If intType = 1 Then
                txtEdit.Text = strValue
                For Each cbrCheck In mcbrToolBar.Controls(5).CommandBar.Controls
                    If cbrCheck.Id = Control.Id Then
                        cbrCheck.Checked = True
                    Else
                        cbrCheck.Checked = False
                    End If
                Next

                Exit Sub
            End If
            
            '非编辑状态下
            If IsWaveItem(lngItemNO) And InStr(1, Trim(vsfTab.TextMatrix(lngRow, lngCol)), "-") <> 0 Then
                strErrMsg = "对于数值已经形成波动范围的波动项目不能进行修改、删除操作"
                lblStb.Caption = strErrMsg: lblStb.ForeColor = 255
                Exit Sub
            End If
            Call txtEdit_KeyPress(vbKeyEscape)
            strPart = CStr(arrStr(7))
            mrsCurve.Filter = "项目序号=" & lngItemNO & " and 项目名称='" & strName & "' and 列号=" & lngCol - vsfTab.FixedCols + 1
            If mrsCurve.RecordCount > 0 Then
                If mrsCurve!状态 <> 1 And mrsCurve!状态 <> 3 Then '原有的数据 修改、删除后的状态始终为2
                    mrsCurve!状态 = 2
                    mrsCurve!数值 = strValue
                Else '对于新增数据的处理
                    If Trim(vsfTab.TextMatrix(lngRow, lngCol)) = "" Then
                        mrsCurve!状态 = 3
                        mrsCurve!数值 = strValue
                    Else
                        mrsCurve!状态 = 1
                        mrsCurve!数值 = strValue
                    End If
                End If
                mrsCurve.Update
            Else '不存在记录就新增数据
                If Trim(strValue) <> "" Then
                    strTime = GetAnimalItemTime(lngRow, lngCol, strErrMsg)
                    If strErrMsg <> "" Then
                        lblStb.Caption = strErrMsg: lblStb.ForeColor = 255
                        Exit Sub
                    End If
                    gstrFields = "序号|分组名|数值|部位|标记|时间|项目序号|项目名称|复查|未记说明|数据来源|修改|显示|来源ID|共用|状态|列号|记录类型"
                    gstrValues = GetMaxID & "|2)体温表格项目|" & strValue & "|" & strPart & "|" & _
                        0 & "|" & strTime & "|" & lngItemNO & "|" & strName & "|0||0|0|0|0|0|1|" & lngCol - vsfTab.FixedCols + 1 & "|1"
                    Call Record_Add(mrsCurve, gstrFields, gstrValues)
                End If
            End If
            vsfTab.TextMatrix(lngRow, lngCol) = strValue
            Call vsfTab_AfterRowColChange(0, 0, lngRow, lngCol)
            mblnChage = True
            
        Case conMenu_Help_Help '帮助
            Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_File_Exit '退出
            Unload Me
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
'    Me.Height = 6600 + 6600 * mintBigSize / 3
'    Me.Width = 9600 + 9600 * mintBigSize / 3
    On Error Resume Next
    fraOper.Height = 735 + 735 * mintBigSize / 3
    Bottom = stbThis.Height + fraOper.Height
    
    With picStb
        .Top = stbThis.Top + 50
        .Left = stbThis.Panels(2).Left + 50
        .Height = stbThis.Height - 50
        .Width = stbThis.Panels(2).Width - 50
    End With
    
    With lblStb
        .Font.Size = 9 + 9 * mintBigSize / 3
        .Height = TextHeight("刘")
        .Top = (picStb.Height - .Height) \ 2
        .Left = 10
    End With
End Sub

Private Sub cbsMain_Resize()
    On Error Resume Next
    
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long  '客户区域的大小

    Call cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    If mblnResize = True Then picSplit.Top = ScaleHeight * 0.7: mblnResize = False
    
    With tbcThis
        .Top = lngTop
        .Left = lngLeft
        .Width = lngRight - lngLeft
        .Height = picSplit.Top - lngTop
    End With
    
    picSplit.Width = tbcThis.Width
    picSplit.Left = lngLeft
    
    With fraOper
        .Top = picSplit.Top + picSplit.Height + 50
        .Left = lngLeft
        .Width = lngRight - lngLeft
        .Height = lngBottom - picSplit.Top + 650
        .Font.Size = 9 + 9 * mintBigSize / 3
    End With
    
    With vsfOper
        .Top = 270
        .Left = 120
        .Width = lngRight - lngLeft - (.Left * 2)
        .Height = fraOper.Height - .Top - 120
        .Body.Font.Size = 9 + 9 * mintBigSize / 3
    End With
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim frmMain As Form
    Dim blnEnable As Boolean

    Select Case Control.Id
        Case conMenu_Edit_Save, conMenu_Edit_Reuse
             Control.Enabled = IIf(mblnChage = True, True, False)
        Case conMenu_Edit_NewItem
            If tbcThis.Selected.Tag = "表格" Then
                Control.Enabled = Not mblnFileBack
            Else
                Control.Enabled = False
            End If
        Case conMenu_Edit_Append * 10 + 0, conMenu_Edit_Append
            Control.Enabled = (is大便或入液(1) Or is大便或入液(2)) And Not mblnFileBack And tbcThis.Selected.Tag = "表格"
        Case conMenu_Edit_Append * 10 + 1, conMenu_Edit_Append * 10 + 2, conMenu_Edit_Append * 10 + 3, conMenu_Edit_Append * 10 + 4
            Control.Enabled = is大便或入液(1) And Not mblnFileBack And tbcThis.Selected.Tag = "表格"
        Case conMenu_Edit_Append * 10 + 5, conMenu_Edit_Append * 10 + 6
            Control.Enabled = is大便或入液(2) And Not mblnFileBack And tbcThis.Selected.Tag = "表格"
        Case conMenu_View_Location
'            Control.Enabled = IIf(mintPreDays > 0, True, False)
'            If Control.Enabled = True Then Control.Enabled = Not mblnFileBack
        Case conMenu_Edit_Curve_Show
            blnEnable = True
            For Each frmMain In Forms
                If frmMain.Name = "frmCaseTendBodySetShowData" Then
                    blnEnable = False
                End If
            Next
            Control.Enabled = blnEnable
    End Select
End Sub

Private Sub cmdColor_Click()
    Call txtEdit_KeyDown(vbKeyDown, vbShiftMask)
End Sub

Private Function dkpDateChageDate(ByVal strValue As String) As Boolean
'------------------------------------------------------------------------------
'补录时间合法时，发生变化就刷新数据
'------------------------------------------------------------------------------
    Dim strErrMsg As String
    Dim strDate As String, strTime As String
    Dim i As Integer
    Dim strCurrDate As String
    Dim intBound As Integer
    Dim strBegin As String, strEnd As String
    Dim intCOl As Integer
    Dim strCurDate As String
    Dim intDay As Integer
    Dim strBTime As String
    
    lblStb.Tag = lblStb.Caption
    
    If Format(strValue, "YYYY-MM-DD") > Format(mstrETime, "YYYY-MM-DD") Then
        If mbln出院 = False Then
            strErrMsg = "录入的日期已超出参数[超期录入天数：" & mintPreDays & "天]所指定的范围！"
        Else
            strErrMsg = "录入的日期不能大于[病人出院时间：" & Format(mstrETime, "YYYY-MM-DD HH:mm:ss") & "]!"
        End If
        GoTo ErrInfo
    End If
    
    If Format(strValue, "YYYY-MM-DD") < Format(mstrBTime, "YYYY-MM-DD") Then
        strErrMsg = "录入的日期不能小于[体温单开始时间：" & Format(mstrBTime, "YYYY-MM-DD HH:mm:ss") & "]！"
        GoTo ErrInfo
    End If
    
    strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
    strCurDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    If Format(strValue, "YYYY-MM-DD") = mstrETime Then
        strDate = Format(Format(mstrETime, "YYYY-MM-DD") & " 00:00:00", "YYYY-MM-DD HH:mm:ss")
        strTime = Format(mstrETime, "YYYY-MM-DD HH:mm:ss")
    ElseIf Format(strValue, "YYYY-MM-DD") = mstrBTime Then
        strDate = Format(mstrBTime, "YYYY-MM-DD HH:mm:ss")
        strTime = strDate
    Else
        strDate = Format(Format(strValue, "YYYY-MM-DD") & " 00:00:00", "YYYY-MM-DD HH:mm:ss")
        strTime = Format(Format(strValue, "YYYY-MM-DD") & " 23:59:00", "YYYY-MM-DD HH:mm:ss")
    End If
    
    If Not IsAllowInput(mT_Patient.lng病人ID, mT_Patient.lng主页ID, strTime, strCurrDate) Then
        strErrMsg = "录入的时间[" & strValue & "]有误！[超过数据补录的有效时限:" & mlngHours & "小时]"
        GoTo ErrInfo
    End If
    
    mblnAllRefresh = True
    
    If UBound(marrTime) = -1 Then Call InitDateTimeRange(marrTime, gintHourBegin)
    intDay = DateDiff("D", CDate(mstrBTime), CDate(strValue)) \ 7
    intDay = (intDay) * 7
    strBTime = Format(DateAdd("d", intDay, CDate(mstrBTime)), "yyyy-MM-dd") & " 00:00:00"
    
    If Format(strValue, "YYYY-MM-DD") = Format(strCurDate, "YYYY-MM-DD") Then
        If Format(strCurDate, "YYYY-MM-DD HH:mm:ss") < Format(strBTime, "YYYY-MM-DD HH:mm:ss") Then
             strDate = Format(strBTime, "YYYY-MM-DD HH:mm:ss")
        ElseIf Format(strCurDate, "YYYY-MM-DD HH:mm:ss") > Format(mstrETime, "YYYY-MM-DD HH:mm:ss") Then
             strDate = Format(mstrETime, "YYYY-MM-DD HH:mm:ss")
        End If
        intCOl = GetCurveColumn(strCurDate, strBTime, gintHourBegin)
        strDate = GetCurveDate(intCOl, strBTime, gintHourBegin)
        strDate = GetCenterTime(Split(strDate, ";")(0), Split(strDate, ";")(1))
    Else
         If Format(strValue, "YYYY-MM-DD") = Format(mstrETime, "YYYY-MM-DD") Then
            intCOl = GetCurveColumn(mstrETime, strBTime, gintHourBegin)
            strDate = GetCurveDate(intCOl, mstrBTime, gintHourBegin)
            strDate = GetCenterTime(Split(strDate, ";")(0), Split(strDate, ";")(1))
         ElseIf Format(strValue, "YYYY-MM-DD") > Format(strCurDate, "YYYY-MM-DD") And Format(strValue, "YYYY-MM-DD") < Format(mstrETime, "YYYY-MM-DD") Then
            strDate = GetCenterTime(Format(strValue, "YYYY-MM-DD 21:00:00"), Format(strValue, "YYYY-MM-DD 23:59:59"))
         End If
    End If

    For i = 0 To UBound(marrTime)
        If Format(strDate, "HH:mm:ss") >= Format(Split(marrTime(i), ",")(0), "HH:mm:ss") And Format(strDate, "HH:mm:ss") <= Format(Split(marrTime(i), ",")(1), "HH:mm:ss") Then
            Exit For
        End If
    Next i
    
    If i > UBound(marrTime) Then i = 0
    
    strBegin = Format(Format(strValue, "YYYY-MM-DD") & " " & Format(Split(marrTime(i), ",")(0), "HH:mm:ss"), "YYYY-MM-DD HH:mm:ss")
    strEnd = Format(Format(strValue, "YYYY-MM-DD") & " " & Format(Split(marrTime(i), ",")(1), "HH:mm:ss"), "YYYY-MM-DD HH:mm:ss")
    
    Call GetCenterTime(CDate(strBegin), CDate(strEnd), intBound)
    For i = 0 To OptTime.Count - 1
        OptTime(i).Caption = gintHourBegin + i * 4
        OptTime(i).Tag = marrTime(i)
        
        If intBound > UBound(marrTime) Then intBound = 0
        If intBound = i Then
            OptTime(i).Value = 1
        End If
    Next i
    
    '如果上面触发了 OptTime_Click 事件 Format(mstrBegin, "YYYY-MM-DD") 和 必定相等
    If Format(mstrBegin, "YYYY-MM-DD") <> Format(dkpDate, "YYYY-MM-DD") Then
        Call OptTime_Click(intBound)
    End If
    
    Call txtEdit_KeyPress(vbKeyEscape)
    
    mblnAllRefresh = False
    dkpDateChageDate = True
    Exit Function
ErrInfo:
    If strErrMsg <> "" Then
        lblStb.Caption = strErrMsg
        lblStb.ForeColor = 255
    End If
    mblnAllRefresh = False
End Function

Private Function CheckDateTime(ByVal lngRow As Long, ByVal strName As String, ByVal strTime As String) As Boolean
'------------------------------------------------------------------
'功能:补录数据时检查数据设置范围
'------------------------------------------------------------------
    Dim strErrMsg As String
    Dim strDate As String
    Dim strCurrDate As String
    Dim strInfo As String
    
    If lngRow <> 0 Then
        strInfo = "第" & lngRow & "行"
    ElseIf strName <> "" Then
        strInfo = strInfo & "[" & strName & "]"
    Else
        strInfo = ""
    End If
    
    If Format(strTime, "YYYY-MM-DD HH:mm") > Format(mstrETime, "YYYY-MM-DD HH:mm") Then
        If mbln出院 = False Then
            strErrMsg = strInfo & "记录数据时间已超出参数[超期录入天数：" & mintPreDays & "天]所指定的范围!"
        Else
            strErrMsg = strInfo & "记录数据时间不能大于[病人出院时间：" & Format(mstrETime, "YYYY-MM-DD HH:mm:ss") & "]!"
        End If
        GoTo ErrInfo
    End If
    
    If Format(strTime, "YYYY-MM-DD HH:mm") < Format(mstrBTime, "YYYY-MM-DD HH:mm") Then
        strErrMsg = strInfo & "记录数据时间不能小于[体温单开始时间：" & Format(mstrBTime, "YYYY-MM-DD HH:mm:ss") & "]!"
        GoTo ErrInfo
    End If
    
    strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
    If Not IsAllowInput(mT_Patient.lng病人ID, mT_Patient.lng主页ID, strTime, strCurrDate) Then
        strErrMsg = strInfo & "记录数据时间[" & strTime & "]有误![超过数据补录的有效时限:" & mlngHours & "小时]"
        GoTo ErrInfo
    End If
    
    CheckDateTime = True
    Exit Function
ErrInfo:
    If strErrMsg <> "" Then
        lblStb.Caption = strErrMsg
        lblStb.ForeColor = 255
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

Private Sub dkpDate_Change()
    Dim strDate As String
    If Not dkpDateChageDate(Format(dkpDate.Value, "YYYY-MM-DD")) Then Exit Sub
    If dkpDate.Enabled = True Then dkpDate.SetFocus
End Sub

Private Sub dkpDate_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call tbcThis_SelectedChanged(tbcThis.Selected)
    End If
End Sub

Private Sub dkpDate_Validate(Cancel As Boolean)
    If Not dkpDateChageDate(Format(dkpDate.Value, "YYYY-MM-DD")) Then
        If Not mblnFileBack Then dkpDate.SetFocus
        Cancel = True
    End If
End Sub

Private Sub dkpTime_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        vsfCurve.SetFocus
    End If
End Sub

Private Sub dkpTime_Change()
    Call ChangeCurveTime
End Sub

Private Sub dkpTime_Validate(Cancel As Boolean)
    If Not ChangeCurveTime Then
        dkpTime.SetFocus
        Cancel = True
    End If
End Sub

Private Sub Form_Activate()
    If Not mblnStart Then Exit Sub
    mblnStart = False
    Call SetColSelect(True)
End Sub

Private Sub Form_Load()
    Call RestoreWinState(Me, App.ProductName)
    If mblnFileBack = True Then lblStb.Caption = "病人体温数据已经归档,不允许进行数据修改.": lblStb.ForeColor = 255
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChage = True Then
        If MsgBox("病人体温数据已经发生改变,请问是否需要保存？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Cancel = True
            Exit Sub
        End If
    End If
    
    mstrPart = ""
    mblnChage = False
    mblnMove = False
    mblnInit = False
    mblnEdit = False
    mbln出院 = False
    mblnAllRefresh = False
    mblnCurveChange = False
    If Not (mrsCurve Is Nothing) Then Set mrsCurve = Nothing
    If Not (mrsPart Is Nothing) Then Set mrsPart = Nothing
    If Not (mrsNote Is Nothing) Then Set mrsNote = Nothing
    If Not (mrsOper Is Nothing) Then Set mrsOper = Nothing
    If Not (mrsRecodeID Is Nothing) Then Set mrsRecodeID = Nothing
    If Not (mcbrToolBar Is Nothing) Then Set mcbrToolBar = Nothing
    '保存窗体
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub FraTime_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intIndex As Integer
    For intIndex = 0 To picBut.Count - 1
        picBut(intIndex).BorderStyle = 0
        picBut(intIndex).BackColor = &H80000004
    Next intIndex
End Sub

Private Sub lblCheck_DblClick()
    Call picEdit_KeyPress(vbKeySpace)
End Sub

Private Sub lstSelect_GotFocus(Index As Integer)
    Dim i As Integer, j As Integer
    PicLst.Tag = 0
    j = lstSelect(Index).ListCount - 1
    If Index = 0 And j >= 0 Then
        If lstSelect(Index).ListIndex < 0 Then lstSelect(Index).ListIndex = 0
    End If
End Sub

Private Sub picHour_GotFocus()
    If picHour.Visible = True Then txtHour.SetFocus
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

Private Sub txtHour_GotFocus()
    Call zlControl.TxtSelAll(txtHour)
End Sub

Private Sub txtHour_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intCOl As Integer, intRow As Integer
    Dim blnAllow As Boolean
    Dim strData As String
    Dim lngColor As Long
    
    If picHour.Visible = False Then Exit Sub
    If KeyCode = vbKeyReturn And Not (Shift = vbShiftMask) Then
        '检查数据合法性
        blnAllow = True
        If picEdit.Visible = True And txtEdit.Tag <> "" Then
            intRow = Split(txtEdit.Tag, "|")(0)
            intCOl = Split(txtEdit.Tag, "|")(1)
            
            If txtEdit.Visible = True Then
                strData = IIf(picHour.Visible = True, "(" & txtHour.Text & "h)", "") & Trim(txtEdit.Text)
                lngColor = txtEdit.ForeColor
            Else
                strData = Trim(lblCheck.Caption)
                lngColor = 0
            End If
            
            If strData & "/#$&/" & lngColor <> picEdit.Tag Then blnAllow = WriteIntoVfgTab(strData, False, False)
        End If
        If blnAllow = True Then
            '移动到下一列
            If txtEdit.Enabled = True Then
                txtEdit.SetFocus
            Else
                Call vsfTab_KeyDown(vbKeyReturn, Shift)
            End If
        Else
            txtHour.SetFocus
        End If
    ElseIf KeyCode = vbKeyLeft And txtHour.SelStart = 0 Then
        Call vsfTab_KeyDown(vbKeyLeft, 0)
    End If
End Sub

Private Sub txtHour_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Call txtEdit_KeyPress(vbKeyEscape)
    Else
        If InStr(1, "0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub txtHour_Validate(Cancel As Boolean)
    Dim strText As String
    strText = txtHour.Text
    If strText = "" Then Exit Sub
    If Not (Val(strText) >= 0 And strText <= 24) Then
        lblStb.Caption = "汇总小时只能在0到24之间，请重新录入！": lblStb.ForeColor = 255
        Cancel = True
    Else
        txtHour.Text = Val(strText)
    End If
End Sub

Private Sub txtLst_GotFocus()
    PicLst.Tag = 1
    Call zlControl.TxtSelAll(txtLst)
    lstSelect(0).ListIndex = -1
End Sub

Private Sub txtLst_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim blnAllow As Boolean
    
    blnAllow = True
    If KeyCode = vbKeyReturn And Shift = vbShiftMask Then Exit Sub
    If KeyCode = vbKeyReturn Then
        If Trim(txtLst.Text) <> lstSelect(0).Tag Then blnAllow = WriteIntoVfgTab(txtLst.Text)
        If blnAllow = True Then Call vsfTab_KeyDown(vbKeyReturn, Shift)
    ElseIf KeyCode = vbKeyLeft And txtLst.SelStart = 0 Then
        Call vsfTab_KeyDown(vbKeyLeft, 0)
    ElseIf Shift = vbShiftMask And KeyCode = vbKeyDown Then
        KeyCode = 0
        lstSelect(0).SetFocus
    ElseIf KeyCode = vbKeyEscape Then
         Call txtEdit_KeyPress(vbKeyEscape)
    End If
End Sub

Private Sub lstSelect_DblClick(Index As Integer)
    Call lstSelect_KeyDown(Index, vbKeyReturn, 0)
End Sub

Private Sub lstSelect_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    Dim blnAllow As Boolean
    Dim strData As String
    Dim i As Integer
    
    strData = ""
    blnAllow = True
    
    If KeyCode = vbKeyReturn Then
        If Shift = vbShiftMask Then Exit Sub
        For i = 0 To lstSelect(Index).ListCount - 1
          If lstSelect(Index).Selected(i) = True Then
              strData = strData & "," & Replace(lstSelect(Index).List(i), ",", "")
          End If
        Next i
        
        If Left(strData, 1) = "," Then strData = Mid(strData, 2)
        If strData <> lstSelect(Index).Tag Then blnAllow = WriteIntoVfgTab(strData)
        If blnAllow = True Then Call vsfTab_KeyDown(vbKeyReturn, Shift)
        
    ElseIf KeyCode = vbKeyLeft Then
        Call vsfTab_KeyDown(vbKeyLeft, 0)
    ElseIf KeyCode = vbKeyEscape Then
        Call txtEdit_KeyPress(vbKeyEscape)
    ElseIf Index = 0 And Shift = vbShiftMask And KeyCode = vbKeyUp Then
        KeyCode = 0
        txtLst.SetFocus
    End If
End Sub

Private Sub lst未记_DblClick()
    Dim intType As Integer
    Dim blnAllow As Boolean
    Dim intCount As Integer
    Dim str未记说明 As String
    Dim intRows As Integer, intRow As Integer
    
    If InStr(1, pic未记.Tag, "|") <> 0 Then
        vsfCurve.Row = Split(pic未记.Tag, "|")(0)
        vsfCurve.Col = Split(pic未记.Tag, "|")(1)
    End If
    
    vsfCurve.TextMatrix(vsfCurve.Row, Col_未记说明) = lst未记.Text
    str未记说明 = lst未记.Text
    vsfCurve.TextMatrix(vsfCurve.Row, col_数据) = Space(vsfCurve.Row) & Space(vsfCurve.Row)
    vsfCurve.TextMatrix(vsfCurve.Row, col_颜色) = Space(vsfCurve.Row) & IIf(vsfCurve.TextMatrix(vsfCurve.Row, COL_分组名) = "2)上下标说明", " ", Space(vsfCurve.Row))
    vsfCurve.TextMatrix(vsfCurve.Row, COL_部位) = ""
    vsfCurve.TextMatrix(vsfCurve.Row, col_复查) = ""
    pic未记.Visible = False
    lst未记.Visible = False: lst未记.Enabled = False
    
    blnAllow = True
    intCount = 0
    intRows = 0
    If Trim(vsfCurve.TextMatrix(vsfCurve.Row, COL_分组名)) = "1)体温曲线项目" Then
        intType = 1
        '如果其它曲线的未记数据为空,直接更新
        For intRow = vsfCurve.FixedRows To vsfCurve.Rows - 1
            If Trim(vsfCurve.TextMatrix(intRow, COL_分组名)) = "1)体温曲线项目" Then
                If vsfCurve.TextMatrix(intRow, Col_未记说明) = "" And Trim(vsfCurve.TextMatrix(intRow, col_数据)) = "" Then
                    intCount = intCount + 1
                End If
                intRows = intRows + 1
            End If
        Next
        '剩下的项目的数据与标记都为空则更新
        If intCount = intRows - 1 Then
            For intRow = vsfCurve.FixedRows To vsfCurve.Rows - 1
                If Trim(vsfCurve.TextMatrix(intRow, COL_分组名)) = "1)体温曲线项目" And vsfCurve.TextMatrix(intRow, Col_未记说明) = "" Then
                    vsfCurve.TextMatrix(intRow, Col_未记说明) = str未记说明
                    vsfCurve.TextMatrix(vsfCurve.Row, col_数据) = Space(vsfCurve.Row) & Space(vsfCurve.Row)
                    vsfCurve.TextMatrix(vsfCurve.Row, col_颜色) = Space(vsfCurve.Row) & IIf(vsfCurve.TextMatrix(vsfCurve.Row, COL_分组名) = "2)上下标说明", " ", Space(vsfCurve.Row))
                    vsfCurve.TextMatrix(vsfCurve.Row, COL_部位) = ""
                    vsfCurve.TextMatrix(vsfCurve.Row, col_复查) = ""
                End If
            Next
        Else
            intCount = 0
        End If
    ElseIf Trim(vsfCurve.TextMatrix(vsfCurve.Row, COL_分组名)) = "2)上下标说明" Then
        If Val(vsfCurve.TextMatrix(vsfCurve.Row, COL_项目序号)) = 4 Then
            'intType = 2
            blnAllow = False
        Else
            blnAllow = False
        End If
    End If
    
    vsfCurve.Cell(flexcpAlignment, vsfCurve.FixedRows, Col_未记说明, vsfCurve.Rows - 1, Col_未记说明) = flexAlignCenterCenter
    
    If blnAllow = True Then
        If intCount = 0 Then
            Call UpdateCurveDate(vsfCurve.Row, vsfCurve.Col, intType)
        ElseIf intCount = intRows - 1 Then
            For intRow = vsfCurve.FixedRows To vsfCurve.Rows - 1
                If Trim(vsfCurve.TextMatrix(intRow, COL_分组名)) = "1)体温曲线项目" Then
                    Call UpdateCurveDate(intRow, Col_未记说明, intType)
                End If
            Next
        End If
        Call vsfCurve.SetFocus
    End If
End Sub

Private Sub lst未记_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        lst未记.Visible = False: lst未记.Enabled = False
        pic未记.Visible = False
    ElseIf KeyCode = vbKeyReturn Then
        Call lst未记_DblClick
    End If
End Sub

Private Sub lst未记_LostFocus()
    lst未记.Visible = False: lst未记.Enabled = False
    pic未记.Visible = False
End Sub

Private Sub OptTime_Click(Index As Integer)
    Dim strBegin As String, strEnd As String
    Dim blnTab As Boolean
    
    If Not mblnInit Then Exit Sub
    
    If mblnCurveChange = True Or (mblnAllRefresh = True And mblnChage = True) Then
        If MsgBox("数据已经发生改变,请问是否进行保存?", vbYesNo + vbQuestion + vbDefaultButton1, gstrSysName) = vbYes Then
            If Not ChangeCurveTime Then Exit Sub
            If Not SaveData Then Exit Sub
            blnTab = True
        Else
            mblnCurveChange = False
            If mblnAllRefresh = True Then
                mblnChage = False
            End If
            blnTab = mblnAllRefresh
        End If
    Else
        blnTab = mblnAllRefresh
    End If
    
    If OptTime(Index).Tag = "" Then Exit Sub
    strBegin = Split(OptTime(Index).Tag, ",")(0)
    strEnd = Split(OptTime(Index).Tag, ",")(1)
    strBegin = Format(Format(dkpDate.Value, " YYYY-MM-DD") & " " & strBegin, "YYYY-MM-DD HH:mm:ss")
    strEnd = Format(Format(dkpDate.Value, " YYYY-MM-DD") & " " & strEnd, "YYYY-MM-DD HH:mm:ss")
    
    If CDate(strBegin) < CDate(mstrBTime) Then
        strBegin = Format(mstrBTime, "YYYY-MM-DD HH:mm:ss")
    End If
    
    If CDate(strEnd) > CDate(mstrETime) Then
        strEnd = Format(mstrETime, "YYYY-MM-DD HH:mm:ss")
    End If
    
    mstrBegin = strBegin
    mstrEnd = strEnd
    
    Call zlRefreshData(True, blnTab)
    
    If mblnStart = False Then
        Call SetColSelect(True)
    End If
End Sub

Public Function SetDate(ByVal strTime As String) As String
'---------------------------------------------------------
' 检查日期
'---------------------------------------------------------
    Dim strVTime As String
    If Not IsDate(strTime) Then Exit Function
    strVTime = Format(strTime, "YYYY-MM-DD HH:mm:ss")
    If CDate(strTime) < CDate(mstrBTime) Then
        strVTime = Format(mstrBTime, "YYYY-MM-DD HH:mm:ss")
    End If
    
    If CDate(strTime) > CDate(mstrETime) Then
        strVTime = Format(mstrETime, "YYYY-MM-DD HH:mm:ss")
    End If
    SetDate = strVTime
End Function

Private Sub picBut_Click(Index As Integer)
    Dim intIndex As Integer, intRow As Integer, intCOl As Integer
    Dim strTime As String
    Dim strOpt As String '手术信息
    Dim lngColor As Long, lngNO As Long
    Dim blnAllow As Boolean
    
    If Not ChangeCurveTime Then Exit Sub
    On Error GoTo Errhand
    Select Case Index
        Case 0 '第一条
            dkpTime.Tag = 0
        Case 1 '上一条
            dkpTime.Tag = Val(dkpTime.Tag) - 1
            If Val(dkpTime.Tag) < 0 Then dkpTime.Tag = 0
        Case 2 '下一条
            dkpTime.Tag = Val(dkpTime.Tag) + 1
            If Val(dkpTime.Tag) > UBound(mArrdkpTime) Then dkpTime.Tag = UBound(mArrdkpTime)
        Case 3 '最后一条
            dkpTime.Tag = UBound(mArrdkpTime)
    End Select
    
    If UBound(mArrdkpTime) = 0 Then
        For intIndex = 0 To picBut.Count - 1
            picBut(intIndex).Visible = False
            picBut(intIndex).Enabled = False
            picBut1(intIndex).Visible = True
            picBut1(intIndex).Enabled = False
        Next intIndex
    Else
        If Val(dkpTime.Tag) = LBound(mArrdkpTime) Then '第一条
            For intIndex = 0 To picBut.Count - 1
                If intIndex < 2 Then
                    picBut(intIndex).Visible = False
                    picBut(intIndex).Enabled = False
                    picBut1(intIndex).Visible = True
                    picBut1(intIndex).Enabled = False
                Else
                    picBut(intIndex).Visible = True
                    picBut(intIndex).Enabled = True
                    picBut1(intIndex).Visible = False
                    picBut1(intIndex).Enabled = False
                End If
            Next intIndex
        ElseIf Val(dkpTime.Tag) = UBound(mArrdkpTime) Then '最后一条
            For intIndex = 0 To picBut.Count - 1
                If intIndex < 2 Then
                    picBut(intIndex).Visible = True
                    picBut(intIndex).Enabled = True
                    picBut1(intIndex).Visible = False
                    picBut1(intIndex).Enabled = False
                Else
                    picBut(intIndex).Visible = False
                    picBut(intIndex).Enabled = False
                    picBut1(intIndex).Visible = True
                    picBut1(intIndex).Enabled = False
                End If
            Next intIndex
        Else '中间某条
            For intIndex = 0 To picBut.Count - 1
                picBut(intIndex).Visible = True
                picBut(intIndex).Enabled = True
                picBut1(intIndex).Visible = False
                picBut1(intIndex).Enabled = False
            Next intIndex
        End If
    End If
    
   '刷新数据
    strTime = Format(mArrdkpTime(Val(dkpTime.Tag)), "YYYY-MM-DD HH:mm:ss")
    dkpTime.Value = Format(strTime, "HH:mm")
    
    '清空所有体温数据信息
    For intRow = vsfCurve.FixedRows To vsfCurve.Rows - 1
        If vsfCurve.TextMatrix(intRow, COL_分组名) <> "" And Val(vsfCurve.TextMatrix(intRow, COL_项目序号)) <> 0 Then
            For intCOl = vsfCurve.FixedCols To vsfCurve.Cols - 1
                vsfCurve.TextMatrix(intRow, intCOl) = ""
            Next intCOl
        End If
    Next intRow
    
    
    blnAllow = False
    ReDim Preserve mArrModfy(vsfCurve.FixedRows To vsfCurve.Rows - 1)
    ReDim Preserve mArrValue(vsfCurve.FixedRows To vsfCurve.Rows - 1)
    ReDim Preserve marrDate(vsfCurve.FixedRows To vsfCurve.Rows - 1)
    '体温数据
    vsfCurve.Cell(flexcpText, vsfCurve.FixedRows, vsfCurve.FixedCols, vsfCurve.Rows - 1, vsfCurve.Cols - 1) = ""
    vsfCurve.Cell(flexcpForeColor, vsfCurve.FixedRows, vsfCurve.FixedCols, vsfCurve.Rows - 1, vsfCurve.Cols - 1) = &H80000012
    
    For intRow = vsfCurve.FixedRows To vsfCurve.Rows - 1
        marrDate(intRow) = 0
        mArrModfy(intRow) = 0
        mArrValue(intRow) = ""

        vsfCurve.Body.MergeRow(intRow) = True
        vsfCurve.TextMatrix(intRow, col_数据) = Space(intRow) & IIf(vsfCurve.TextMatrix(intRow, COL_分组名) = "2)上下标说明", "", "") & Space(intRow)
        vsfCurve.TextMatrix(intRow, col_颜色) = Space(intRow) & IIf(vsfCurve.TextMatrix(intRow, COL_分组名) = "2)上下标说明", " ", Space(intRow))
        If vsfCurve.TextMatrix(intRow, COL_分组名) = "2)上下标说明" Then
             vsfCurve.Cell(flexcpBackColor, intRow, col_颜色, intRow, col_颜色) = RGB(0, 0, 255)
        End If
    Next intRow
    
    mrsCurve.Filter = "时间='" & strTime & "' and 状态<>3"
    With mrsCurve
        Do While Not .EOF
            For intRow = vsfCurve.FixedRows To vsfCurve.Rows - 1
                lngNO = Val(vsfCurve.TextMatrix(intRow, COL_项目序号))
                If !分组名 = vsfCurve.TextMatrix(intRow, COL_分组名) And !项目序号 = lngNO Then
                    vsfCurve.TextMatrix(intRow, col_数据) = Space(intRow) & zlCommFun.Nvl(!数值) & Space(intRow)
                    vsfCurve.TextMatrix(intRow, col_颜色) = vsfCurve.TextMatrix(intRow, col_数据)
                    
                    If Not IsNumeric(zlCommFun.Nvl(!数值)) And zlCommFun.Nvl(!数值) <> "不升" And InStr(1, zlCommFun.Nvl(!数值), "/") = 0 Then
                        vsfCurve.TextMatrix(intRow, COL_部位) = ""
                        vsfCurve.TextMatrix(intRow, Col_未记说明) = zlCommFun.Nvl(!未记说明)
                    Else
                        vsfCurve.TextMatrix(intRow, COL_部位) = zlCommFun.Nvl(!部位)
                        vsfCurve.TextMatrix(intRow, Col_未记说明) = ""
                    End If
                    If lngNO = 1 And (IsNumeric(zlCommFun.Nvl(!数值)) Or zlCommFun.Nvl(!数值) <> "不升") Then
                        vsfCurve.TextMatrix(intRow, col_复查) = IIf(Val(zlCommFun.Nvl(!复查)) = 1, "√", "")
                    End If
                    lngColor = 255
                    If Val(zlCommFun.Nvl(!数据来源)) <> 0 Then
                        If zlCommFun.Nvl(!数值) = "不升" And lngNO = 1 Then
                            lngColor = 255
                        ElseIf lngNO = 1 Or (lngNO = 2 And mint心率应用 = 2 And mblnEdit心率 = True) Then
                            If InStr(1, zlCommFun.Nvl(!数值), "/") = 0 Then
                                lngColor = RGB(0, 0, 255)
                            Else
                                If Val(!修改) = 0 Then
                                    lngColor = RGB(0, 0, 255)
                                Else
                                    lngColor = 255
                                End If
                            End If
                        End If
                        vsfCurve.Cell(flexcpForeColor, intRow, col_数据, intRow, col_数据) = lngColor
                    Else
                        vsfCurve.Cell(flexcpForeColor, intRow, col_数据, intRow, col_数据) = &H80000012
                    End If
                    marrDate(intRow) = Val(CStr(zlCommFun.Nvl(!数据来源)))
                    If InStr(1, ",0,9,", Val(zlCommFun.Nvl(!数据来源))) = 0 Then
                        blnAllow = True
                    End If
                    mArrModfy(intRow) = Val(!修改)
                    mArrValue(intRow) = Val(!数值)
                End If
            Next intRow
        .MoveNext
        Loop
    End With
    
    If blnAllow = True Or mblnFileBack = True Then
        dkpTime.Enabled = False
    Else
        dkpTime.Enabled = True
    End If
    
    '上下标(手术始终保持不变)
    mrsNote.Filter = 0
    With mrsNote
        Do While Not .EOF
            Select Case Val(!记录类型)
                Case 2
                    intRow = mOptRow.上标
                Case 6
                    intRow = mOptRow.下标
            End Select
            vsfCurve.TextMatrix(intRow, col_数据) = Space(intRow) & zlCommFun.Nvl(!内容) & Space(intRow)
            vsfCurve.Cell(flexcpBackColor, intRow, col_颜色, intRow, col_颜色) = IIf(IsNumeric(Nvl(!未记说明)) = False, 16711680, Val(Nvl(!未记说明)))
            vsfCurve.TextMatrix(intRow, COL_部位) = ""
            vsfCurve.TextMatrix(intRow, col_复查) = ""
            vsfCurve.TextMatrix(intRow, Col_未记说明) = ""
        .MoveNext
        Loop
    End With
    
    If mblnStart = False Then
        Call SetColSelect
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub picBut_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim intIndex As Integer
    Dim strInfo As String
    
    For intIndex = 0 To picBut.Count - 1
        If Index = intIndex Then
            picBut(intIndex).BorderStyle = 0
            picBut(intIndex).BackColor = &HFFC0C0
        Else
            picBut(intIndex).BorderStyle = 0
            picBut(intIndex).BackColor = &H80000004
        End If
    Next intIndex
    
    Select Case Index
        Case 0
            strInfo = "第一条"
        Case 1
            strInfo = "上一条"
        Case 2
            strInfo = "下一条"
        Case 3
            strInfo = "最后一条"
    End Select
    
    picBut(Index).ToolTipText = strInfo
End Sub

Private Sub picBut1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strInfo As String
    Select Case Index
        Case 0
            strInfo = "第一条"
        Case 1
            strInfo = "上一条"
        Case 2
            strInfo = "下一条"
        Case 3
            strInfo = "最后一条"
    End Select
    
    picBut1(Index).ToolTipText = strInfo
End Sub

Private Sub picColor_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then picColor.Visible = False
End Sub

Private Sub picColor_Resize()
    With usrColor
        .Top = -450
        .Left = 0
        .Width = picColor.Width
        .Height = picColor.Height
    End With
End Sub

Private Sub picCurve_Resize()
    
    With lblTime
        .Left = 50
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
    End With
    
    With dkpTime
        .Top = 210
        .Left = lblTime.Left + lblTime.Width + 30
        .Height = 300 + 300 * mintBigSize / 3
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
    End With
    
    With FraTime
        .Top = 0
        .Left = 0
        .Width = picCurve.Width
        .Height = dkpTime.Top + 100 + dkpTime.Height
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
    End With
    
    With picPre
        .Top = 150 + 150 * mintBigSize / 3
        .Left = dkpTime.Left + dkpTime.Width + 100
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
    End With
    
    With picToolBar
        .Top = 210
         .Width = 2775 + 2775 * mintBigSize / 3
        .Height = 350 + 350 * mintBigSize / 3
        .Left = FraTime.Width - .Width - 50
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
    End With
    
    With FraData
        .Left = 0
        .Width = picCurve.Width
        .Top = FraTime.Height
        .Height = picCurve.Height - .Top
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
    End With
    
    With vsfCurve
        .Top = 0
        .Left = 0
        .Width = FraData.Width
        .Height = FraData.Height
    End With
    
    With pic未记
        .Width = 1080 + 1080 * mintBigSize / 3
        .Height = 1100 + 1100 * mintBigSize / 3
        .Visible = False
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
    End With
    
    With lst未记
        .Top = 0
        .Left = 0
        .Width = pic未记.Width
        .Height = pic未记.Height
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
    End With
    
    With PicValue
        .Width = 2190
        .Height = 2190 - 450
        .Visible = False
    End With
    
    Call picPre_Resize
End Sub

Private Function GetTableRowName() As String
    Dim rsTemp As New ADODB.Recordset
    Dim strTmp As String
    Dim strTmpCurve As String, strTmpTable As String '曲线和表格项目变量
    Dim strCollectItem As String '表格汇总项目
    Dim arrActive() As String
    Dim str值域 As String
    Dim strSQL As String
    Dim i As Integer, intBound As Integer
    Dim strEndTime As String
    Dim Titem As Type_Item
    Dim strDate As String
    Dim intCOl As Integer
    Dim strCurDate As String
    
    On Error GoTo Errhand
    
    Call InitRecordSet
    
    '检查脉搏心率共用时心率是否使用与此病人
    mstrSQL = "select C.应用方式 From 护理记录项目 C where C.项目序号=[1] And C.护理等级>=[2] And Nvl(C.适用病人,0) In (0,[3]) " & _
            " And (c.适用科室=1 Or (c.适用科室=2 And Exists (Select 1 From 护理适用科室 D Where D.项目序号=c.项目序号 And D.科室id=[4])))"
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "提取信心率", -1, mT_Patient.lng护理等级, IIf(mT_Patient.lng婴儿 = 0, 1, 2), mT_Patient.lng科室ID)
    mblnEdit心率 = IIf(rsTemp.RecordCount = 0, False, True)
    If rsTemp.RecordCount > 0 Then mint心率应用 = Val(zlCommFun.Nvl(rsTemp!应用方式, 0))
    
    '格式组成为 类型'值域,项目类型,项目小数,记录频次,项目表示,项目性质,项目长度,部位,入院首测'项目号'项目名
    strTmp = "2)上下标说明',,,,,,,,'2'上标;2)上下标说明',,,,,,,,'6'下标"
    
    '提取所有体温曲线项目
    mstrSQL = " Select A.排列序号,A.记录名 项目名,A.项目序号 as 项目号,A.记录法,A.入院首测," & _
            " C.项目值域,C.项目类型,C.项目长度,C.项目小数,nvl(A.记录频次,2) 记录频次,C.分组名,C.项目表示,C.项目单位 " & _
            " From 体温记录项目 A,诊治所见项目 B,护理记录项目 C " & _
            " Where c.项目ID=B.ID(+) And A.项目序号=C.项目序号 And (A.记录法=1 OR (A.记录法=2 And A.项目序号=3)) And 项目性质=1 And Nvl(C.应用方式,0)=1 AND C.护理等级>=[1] And Nvl(C.适用病人,0) In (0,[3]) " & _
            " And (c.适用科室=1 Or (c.适用科室=2 And Exists (Select 1 From 护理适用科室 D Where D.项目序号=c.项目序号 And D.科室id=[2])))" & _
            " Order by Decode(A.项目序号,1,0,1),A.排列序号"
            
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, mT_Patient.lng护理等级, mT_Patient.lng科室ID, IIf(mT_Patient.lng婴儿 = 0, 1, 2))
    
    With rsTemp
        Do While Not .EOF
            str值域 = Replace(zlCommFun.Nvl(!项目值域), ":", "")
            If zlCommFun.Nvl(!项目类型) = 0 Then
                If InStr(1, str值域, ";") <> 0 Then str值域 = Split(str值域, ";")(0) & "～" & Split(str值域, ";")(1)
            End If
            str值域 = Replace(Replace(Replace(str值域, ";", ":"), "'", ""), ",", "")
            
            Titem.值域 = str值域
            Titem.项目类型 = Val(zlCommFun.Nvl(!项目类型, 0))
            Titem.项目小数 = Val(zlCommFun.Nvl(!项目小数, 0))
            Titem.记录频次 = Val(zlCommFun.Nvl(!记录频次, 2))
            Titem.项目表示 = Val(zlCommFun.Nvl(!项目表示, 0))
            Titem.项目性质 = 1
            Titem.项目长度 = zlCommFun.Nvl(!项目长度, 3)
            Titem.部位 = ""
            Titem.项目号 = Val(zlCommFun.Nvl(!项目号))
            Titem.项目名 = Replace(Replace(zlCommFun.Nvl(!项目名) & IIf(zlCommFun.Nvl(!项目单位, "") = "", "", "(" & !项目单位 & ")"), ";", ":"), "'", "")
            Titem.入院首测 = Val(zlCommFun.Nvl(!入院首测, 0))
            
            If Titem.项目表示 = 4 Or IsWaveItem(Titem.项目号) Then
                If Titem.记录频次 > 2 Then Titem.记录频次 = 2
            End If
            '记录法=1或记录法=2的呼吸项都为曲线项目
            Titem.类型 = "1)体温曲线项目"
            strTmpCurve = strTmpCurve & ";" & Titem.类型 & "'" & Titem.值域 & "," & Titem.项目类型 & "," & _
                Titem.项目小数 & "," & Titem.记录频次 & "," & Titem.项目表示 & ",1," & Titem.项目长度 & ",," & Titem.入院首测 & "'" & _
                Titem.项目号 & "'" & Titem.项目名
        .MoveNext
        Loop
    End With
    
    mstrActiveItem = ""
    
    strEndTime = DateAdd("d", 6, CDate(Format(Format(mstrDate, "YYYY-MM-DD") & " 23:59:59", "YYYY-MM-DD HH:mm:ss")))
    If strEndTime > mstrETime Then strEndTime = mstrETime
    '提取固定表格项目和有数值的活动项目信息
    Set rsTemp = GetAppendGridItem(mT_Patient.lng文件ID, mT_Patient.lng病人ID, mT_Patient.lng主页ID, mT_Patient.lng护理等级, mT_Patient.lng婴儿, _
        CDate(mstrDate), CDate(strEndTime), IIf(mT_Patient.lng婴儿 = 0, 1, 2), mT_Patient.lng科室ID, mblnMove)
    With rsTemp
        Do While Not .EOF
           str值域 = Replace(zlCommFun.Nvl(!项目值域), ":", "")
            If zlCommFun.Nvl(!项目类型) = 0 Then
                If InStr(1, str值域, ";") <> 0 Then str值域 = Split(str值域, ";")(0) & "～" & Split(str值域, ";")(1)
            End If
            str值域 = Replace(Replace(Replace(str值域, ";", ":"), "'", ""), ",", "")
            
            Titem.值域 = str值域
            Titem.类型 = "2)体温表格项目"
            Titem.项目类型 = Val(zlCommFun.Nvl(!项目类型))
            Titem.项目小数 = Val(zlCommFun.Nvl(!项目小数, 0))
            Titem.记录频次 = Val(zlCommFun.Nvl(!记录频次, 2))
            Titem.项目表示 = Val(zlCommFun.Nvl(!项目表示, 0))
            Titem.项目性质 = Val(zlCommFun.Nvl(!项目性质, 1))
            Titem.项目长度 = zlCommFun.Nvl(!项目长度, 3)
            Titem.部位 = Replace(Replace(Replace(zlCommFun.Nvl(!体温部位), ";", ""), "'", ""), ",", "")
            Titem.项目号 = Val(zlCommFun.Nvl(!项目序号))
            Titem.项目名 = Replace(Replace(IIf(Titem.项目号 = 4, "血压", zlCommFun.Nvl(!记录名)) & IIf(zlCommFun.Nvl(!单位, "") = "", "", "(" & !单位 & ")"), ";", ":"), "'", "")
            Titem.入院首测 = Val(zlCommFun.Nvl(!入院首测, 0))
            If Titem.项目表示 = 4 Or IsWaveItem(Titem.项目号) Then
                If Titem.记录频次 > 2 Then Titem.记录频次 = 2
            End If
            If Titem.项目号 <> gint呼吸 And Titem.项目号 <> 5 Then
                strTmpTable = strTmpTable & ";" & Titem.类型 & "'" & Titem.值域 & "," & Titem.项目类型 & "," & _
                    Titem.项目小数 & "," & Titem.记录频次 & "," & Titem.项目表示 & "," & Titem.项目性质 & "," & Titem.项目长度 & "," & _
                    Titem.部位 & "," & Titem.入院首测 & "'" & Titem.项目号 & "'" & Titem.项目名
                '记录已经存在的活动项目信息
                If Titem.项目性质 = 2 Then
                    mstrActiveItem = mstrActiveItem & ";" & Titem.类型 & "'" & Titem.值域 & "," & Titem.项目类型 & "," & _
                        Titem.项目小数 & "," & Titem.记录频次 & "," & Titem.项目表示 & "," & Titem.项目性质 & "," & Titem.项目长度 & "," & _
                        Titem.部位 & "," & Titem.入院首测 & "'" & Titem.项目号 & "'" & Titem.项目名
                End If
            End If
        .MoveNext
        Loop
    End With
    
    If Left(mstrActiveItem, 1) = ";" Then mstrActiveItem = Mid(mstrActiveItem, 2)
        
    If strTmp <> "" Then strTmpCurve = strTmpCurve & ";" & strTmp
    If Left(strTmpCurve, 1) = ";" Then strTmpCurve = Mid(strTmpCurve, 2)
    If Left(strTmpTable, 1) = ";" Then strTmpTable = Mid(strTmpTable, 2)
    
    '加载体温曲线数据包括手术上下标
    Call InitTabCurve(strTmpCurve)
    
    '加载体温表格数据(不包含汇总项目)
    Call InitTabTable(strTmpTable)
    
    '加载手术录入表格
    Call InitTabOper
    
    mstr未记说明 = ""
    '提取未记说明信息
    mstrSQL = "Select 编码,名称 From 常用体温说明"
    Call zlDatabase.OpenRecordset(rsTemp, mstrSQL, Me.Caption)
    With rsTemp
        Do While Not .EOF
            mstr未记说明 = mstr未记说明 & "," & zlCommFun.Nvl(!名称)
        .MoveNext
        Loop
    End With
    
    If Left(mstr未记说明, 1) = "," Then mstr未记说明 = Mid(mstr未记说明, 2)
    
    '根据选择时间定位在当前时间编辑状态
    strCurDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")
    If Format(mstrBegin, "YYYY-MM-DD") = Format(strCurDate, "YYYY-MM-DD") Then
        If Format(strCurDate, "YYYY-MM-DD HH:mm:ss") < Format(mstrBegin, "YYYY-MM-DD HH:mm:ss") Then
             strCurDate = Format(mstrBegin, "YYYY-MM-DD HH:mm:ss")
        ElseIf Format(strCurDate, "YYYY-MM-DD HH:mm:ss") > Format(strEndTime, "YYYY-MM-DD HH:mm:ss") Then
             strCurDate = Format(mstrEnd, "YYYY-MM-DD HH:mm:ss")
        End If
        intCOl = GetCurveColumn(strCurDate, mstrBegin, gintHourBegin)
        strDate = GetCurveDate(intCOl, mstrBegin, gintHourBegin)
        mstrBegin = Split(strDate, ";")(0)
        mstrEnd = Split(strDate, ";")(1)
    Else
         If Format(mstrBegin, "YYYY-MM-DD") = Format(strEndTime, "YYYY-MM-DD") Then
            intCOl = GetCurveColumn(mstrEnd, mstrBegin, gintHourBegin)
            strDate = GetCurveDate(intCOl, mstrBegin, gintHourBegin)
            mstrBegin = Split(strDate, ";")(0)
            mstrEnd = Split(strDate, ";")(1)
         ElseIf Format(mstrBegin, "YYYY-MM-DD") > Format(strCurDate, "YYYY-MM-DD") And Format(mstrBegin, "YYYY-MM-DD") < Format(strEndTime, "YYYY-MM-DD") Then
            mstrBegin = Format(mstrBegin, "YYYY-MM-DD 21:00:00")
            mstrEnd = Format(mstrBegin, "YYYY-MM-DD 23:59:59")
         End If
    End If
    
    Call GetCenterTime(CDate(mstrBegin), CDate(mstrEnd), intBound)
    For i = 0 To OptTime.Count - 1
        OptTime(i).Caption = gintHourBegin + i * 4
        OptTime(i).Tag = marrTime(i)
        
        If intBound > UBound(marrTime) Then intBound = 0
        If intBound = i Then
            OptTime(i).Value = 1
        End If
    Next i

    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function zlRefreshData(Optional ByVal blnCurve As Boolean = True, Optional ByVal blnTab As Boolean = True) As Boolean
'-----------------------------------------------------------------------------------------------------------------
'功能:提取一段时间内的所有体温数据
'参数 blnCurve是否刷新体温数据 blnTab 是否刷新表格数据
'-----------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim rsDownTab As New ADODB.Recordset
    Dim dtBegin As Date, dtEnd As Date
    Dim lng项目序号 As Long, int项目性质 As Integer, str项目名称 As String, int记录频次 As Integer, int项目表示 As Integer, int入院首测 As Integer
    Dim intRow As Integer, intCOl As Integer, intNum As Integer, strName As String
    Dim strParam As String, strFidlds As String, strPart As String, strTmp As String
    Dim blnAllow As Boolean, blnAdd As Boolean
    Dim strTime As String
    Dim rsCurve As New ADODB.Recordset '临时记录集
    Dim intModify As Integer, int数据来源 As Integer
    Dim lngColor As Long
    Dim i As Integer, int标记 As Integer
    Dim strOperTime As String, strOper As String
    Dim strItems As String, strItemName As String

    On Error GoTo Errhand
    
    If blnCurve = False And blnTab = False Then Exit Function
    
    lblTime.Caption = Format(mstrBegin, "HH:mm") & "～" & Format(mstrEnd, "HH:mm")
    dkpTime.MaxDate = Format(mstrEnd, "HH:mm")
    dkpTime.MinDate = Format(mstrBegin, "HH:mm")
    mArrdkpTime = Array()
        
    '初始化记录集
    gstrFields = "记录ID," & adDouble & ",18|时间," & adLongVarChar & ",20"
    Call Record_Init(mrsRecodeID, gstrFields)
    
    '修改 表示对于同步过来的数据，如果体温没有物理降温,脉搏和心率无短轴 这可以进行物理降温和短轴数据的修改  0 可以修改 1不能修改
    gstrFields = "序号," & adDouble & ",18|分组名," & adLongVarChar & ",40|数值," & adLongVarChar & ",400|部位," & adLongVarChar & ",200|" & _
         "标记," & adDouble & ",1|时间," & adLongVarChar & ",20|原始时间," & adLongVarChar & ",20|项目序号," & adDouble & ",18|项目名称," & adLongVarChar & ",40|" & _
         "复查," & adDouble & ",1|未记说明," & adLongVarChar & ",20|数据来源," & adDouble & ",1|修改," & adDouble & ",1|显示," & adDouble & ",1|" & _
         "来源ID," & adDouble & ",18|共用," & adDouble & ",1|状态," & adDouble & ",1|列号," & adDouble & ",1|记录类型," & adDouble & ",1"
    Call Record_Init(rsCurve, gstrFields)
         
    If blnCurve = True And blnTab = False Then '只刷新体温数据
        If Not mrsCurve Is Nothing And mrsCurve.State = 1 Then
            mrsCurve.Filter = 0
            mrsCurve.Filter = "分组名='2)体温表格项目'"
            Do While Not mrsCurve.EOF
                rsCurve.AddNew
                For i = 0 To mrsCurve.Fields.Count - 1
                    rsCurve.Fields(mrsCurve.Fields(i).Name).Value = mrsCurve.Fields(i).Value
                Next i
                rsCurve.Update
            mrsCurve.MoveNext
            Loop
        End If
    ElseIf blnCurve = False And blnTab = True Then '只刷新表格
        If Not mrsCurve Is Nothing And mrsCurve.State = 1 Then
            mrsCurve.Filter = 0
            mrsCurve.Filter = "分组名='1)体温曲线项目'"
            Do While Not mrsCurve.EOF
                rsCurve.AddNew
                For i = 0 To mrsCurve.Fields.Count - 1
                    rsCurve.Fields(mrsCurve.Fields(i).Name).Value = mrsCurve.Fields(i).Value
                Next i
                rsCurve.Update
            mrsCurve.MoveNext
            Loop
        End If
    End If
         
    Call Record_Init(mrsCurve, gstrFields)
    
    gstrFields = "序号|分组名|数值|部位|标记|时间|原始时间|项目序号|项目名称|复查|未记说明|数据来源|修改|显示|来源ID|共用|状态|列号|记录类型"
    
    '刷新体温曲线数据以及手术上下标
    If blnCurve = True Then
        '1------------------------------------------------------------
        '提取某时间段的所有体温曲线数据
        mstrSQL = _
        " SELECT C.ID 序号,C.记录ID,A.发生时间 As 时间,'1)体温曲线项目' 分组名,C.显示,c.记录内容 As 数值,c.体温部位,c.复试合格,D.记录名,D.项目序号,DECODE(D.项目序号,-1,1,C.记录标记) 记录标记,C.未记说明,C.数据来源,C.来源ID,C.共用" & vbNewLine & _
        "                    FROM 病人护理文件 B,病人护理数据 A,病人护理明细 C,体温记录项目 D,护理记录项目 E" & vbNewLine & _
        "                    Where B.ID=A.文件ID" & vbNewLine & _
        "                        AND A.ID = C.记录ID" & vbNewLine & _
        "                        AND B.ID=[1]" & vbNewLine & _
        "                        AND Nvl(B.婴儿,0)=[4]" & vbNewLine & _
        "                        AND B.病人id=[2]" & vbNewLine & _
        "                        AND B.主页id=[3]" & vbNewLine & _
        "                        AND D.项目序号=C.项目序号" & vbNewLine & _
        "                        AND C.记录类型=1" & vbNewLine & _
        "                        AND E.项目序号=D.项目序号" & vbNewLine & _
        "                        AND E.护理等级>=[7]" & vbNewLine & _
        "                        AND (nvl(D.记录法,1)=1 or (NVL(D.记录法,1)=2 And D.项目序号=3))" & _
        "                        And A.发生时间 BETWEEN [5] And [6] And C.终止版本 Is Null" & vbNewLine & _
        "                        AND (nvl(E.应用方式,0)=1 OR ( -1=[10] and nvl(E.应用方式,0)=2))" & vbNewLine & _
        "                        AND nvl(E.适用病人,0) in (0,[8]) AND (E.适用科室=1 or ( E.适用科室=2 AND Exists (select 1 from 护理适用科室 D where D.项目序号=E.项目序号 and D.科室ID=[9])))" & vbNewLine & _
        "                    Order By A.发生时间,DECODE(D.项目序号,-1,1,0),DECODE(D.项目序号,-1,1,C.记录标记)"
    
        If mblnMove Then
            mstrSQL = Replace(mstrSQL, "病人护理文件", "H病人护理文件")
            mstrSQL = Replace(mstrSQL, "病人护理数据", "H病人护理数据")
            mstrSQL = Replace(mstrSQL, "病人护理明细", "H病人护理明细")
        End If
        
        Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, mT_Patient.lng文件ID, mT_Patient.lng病人ID, mT_Patient.lng主页ID, mT_Patient.lng婴儿, _
             CDate(Format(mstrBegin, "YYYY-MM-DD HH:mm:ss")), CDate(Format(mstrEnd, "YYYY-MM-DD HH:mm:ss")), mT_Patient.lng护理等级, IIf(mT_Patient.lng婴儿 = 0, 1, 2), mT_Patient.lng科室ID, IIf(mint心率应用 = 2, -1, 0))
        With rsTmp
            
            Do While Not .EOF
                
                '添加记录集
                Call Record_Update(mrsRecodeID, "记录ID|时间", Val(Nvl(!记录ID)) & "|" & Format(zlCommFun.Nvl(!时间), "YYYY-MM-DD HH:mm:ss"), "记录ID|" & Val(Nvl(!记录ID)))
                
                intModify = 0
                If strTime = "" Then strTime = Format(zlCommFun.Nvl(!时间), "YYYY-MM-DD HH:mm:ss")
                lng项目序号 = zlCommFun.Nvl(!项目序号)
                Select Case lng项目序号
                    Case gint心率
                        int标记 = 1
                    Case Else
                        int标记 = Val(Nvl(!记录标记))
                End Select
                intModify = IIf(InStr(1, ",0,9,", "," & Val(zlCommFun.Nvl(!数据来源)) & ",") = 0, 1, 0)
                blnAdd = True
                '心率和脉搏公用时，检查脉搏对应的时间是否存在心率
                If mint心率应用 = 2 And lng项目序号 = -1 Then
                    mrsCurve.Filter = "项目序号=2 and 时间='" & Format(zlCommFun.Nvl(!时间), "YYYY-MM-DD HH:mm:ss") & "'"
                    If mrsCurve.RecordCount > 0 Then
                        strParam = "序号|" & mrsCurve("序号")
                        strFidlds = "数值|标记|修改"
                        
                        If InStr(1, ",0,9,", "," & Val(zlCommFun.Nvl(mrsCurve!数据来源)) & ",") = 0 And InStr(1, ",0,9,", "," & Val(zlCommFun.Nvl(!数据来源)) & ",") = 0 Then
                            intModify = 1
                        Else
                            intModify = 0
                        End If
                        
                        '脉搏短轴时心率未未记说明只显示脉搏，脉搏为未记说明时就显示未记说明
                        If UBound(Split(mrsCurve("数值"), "/")) <> -1 Then
                            If IsNumeric(zlCommFun.Nvl(!数值)) Then
                                If mbln脉搏共用显示 Then
                                    gstrValues = zlCommFun.Nvl(!数值) & "/" & Split(mrsCurve("数值"), "/")(0) & "|" & int标记 & "|" & intModify
                                Else
                                    gstrValues = Split(mrsCurve("数值"), "/")(0) & "/" & zlCommFun.Nvl(!数值) & "|" & int标记 & "|" & intModify
                                End If
                            Else
                                gstrValues = Split(mrsCurve("数值"), "/")(0) & "|" & int标记 & "|0"
                            End If
                        Else
                            gstrValues = mrsCurve("数值") & "|1|0"
                        End If
                        
                        Call Record_Update(mrsCurve, strFidlds, gstrValues, strParam)
                        blnAdd = False
                    Else
                        lng项目序号 = 2
                    End If
                End If
                
                '处理物理降温
                If lng项目序号 = 1 And int标记 = 1 Then
                    mrsCurve.Filter = "项目序号=1 and 时间='" & Format(zlCommFun.Nvl(!时间), "YYYY-MM-DD HH:mm:ss") & "' and 标记<>1"
                    If mrsCurve.RecordCount > 0 Then
                        If InStr(1, ",0,9,", "," & Val(zlCommFun.Nvl(mrsCurve!数据来源)) & ",") = 0 And InStr(1, ",0,9,", "," & Val(zlCommFun.Nvl(!数据来源)) & ",") = 0 Then
                            intModify = 1
                        Else
                            intModify = 0
                        End If
                        
                        strParam = "序号|" & mrsCurve("序号")
                        strFidlds = "数值|标记|修改"
                        gstrValues = Split(mrsCurve("数值"), "/")(0) & "/" & zlCommFun.Nvl(!数值) & "|" & int标记 & "|" & intModify
                        Call Record_Update(mrsCurve, strFidlds, gstrValues, strParam)
                    End If
                    blnAdd = False
                End If
                
                If blnAdd Then
                    '进行曲线显示处理
                    strPart = GetPart(lng项目序号)
                    int数据来源 = Val(zlCommFun.Nvl(!数据来源, 0))
                    If Trim(Replace(zlCommFun.Nvl(!数值), "/", "")) = "" Then
                        int数据来源 = 0
                    End If
                    gstrValues = zlCommFun.Nvl(!序号) & "|" & zlCommFun.Nvl(!分组名) & "|" & Trim(Replace(zlCommFun.Nvl(!数值), "/", "")) & "|" & _
                        zlCommFun.Nvl(!体温部位, strPart) & "|" & int标记 & "|" & Format(zlCommFun.Nvl(!时间), "YYYY-MM-DD HH:mm:ss") & "|" & _
                        Format(zlCommFun.Nvl(!时间), "YYYY-MM-DD HH:mm:ss") & "|" & lng项目序号 & "|" & zlCommFun.Nvl(!记录名) & "|" & Val(zlCommFun.Nvl(!复试合格, 0)) & "|" & _
                        zlCommFun.Nvl(!未记说明) & "|" & int数据来源 & "|" & intModify & "|" & Val(zlCommFun.Nvl(!显示, 0)) & "|" & Val(zlCommFun.Nvl(!来源ID, 0)) & "|" & Val(zlCommFun.Nvl(!共用, 0)) & "|0|0|1"
                    Call Record_Add(mrsCurve, gstrFields, gstrValues)
                End If
            .MoveNext
            Loop
        End With

        strTmp = strTime
        If strTmp <> "" Then
            ReDim Preserve mArrdkpTime(UBound(mArrdkpTime) + 1)
            mArrdkpTime(UBound(mArrdkpTime)) = strTmp
        End If
        
        blnAllow = False
        '显示体温数据
        mrsCurve.Filter = 0
        mrsCurve.Sort = "时间"
        
        ReDim Preserve mArrModfy(vsfCurve.FixedRows To vsfCurve.Rows - 1)
        ReDim Preserve mArrValue(vsfCurve.FixedRows To vsfCurve.Rows - 1)
        ReDim Preserve marrDate(vsfCurve.FixedRows To vsfCurve.Rows - 1)
        
        vsfCurve.Cell(flexcpText, vsfCurve.FixedRows, vsfCurve.FixedCols, vsfCurve.Rows - 1, vsfCurve.Cols - 1) = ""
        vsfCurve.Cell(flexcpForeColor, vsfCurve.FixedRows, vsfCurve.FixedCols, vsfCurve.Rows - 1, vsfCurve.Cols - 1) = &H80000012
        vsfCurve.Cell(flexcpBackColor, vsfCurve.FixedRows, vsfCurve.FixedCols, vsfCurve.Rows - 1, vsfCurve.Cols - 1) = &H80000005
        
        For intRow = vsfCurve.FixedRows To vsfCurve.Rows - 1
            marrDate(intRow) = 0
            mArrModfy(intRow) = 0
            mArrValue(intRow) = ""

            vsfCurve.Body.MergeRow(intRow) = True
            vsfCurve.TextMatrix(intRow, col_数据) = Space(intRow) & IIf(vsfCurve.TextMatrix(intRow, COL_分组名) = "2)上下标说明", "", "") & Space(intRow)
            vsfCurve.TextMatrix(intRow, col_颜色) = Space(intRow) & IIf(vsfCurve.TextMatrix(intRow, COL_分组名) = "2)上下标说明", " ", Space(intRow))
            If vsfCurve.TextMatrix(intRow, COL_分组名) = "2)上下标说明" Then
                 vsfCurve.Cell(flexcpBackColor, intRow, col_颜色, intRow, col_颜色) = RGB(0, 0, 255)
            End If
        Next intRow
        
        With mrsCurve
            Do While Not .EOF
                If Format(strTime, "YYYY-MM-DD HH:mm:ss") = Format(!时间, "YYYY-MM-DD HH:mm:ss") Then
                    For intRow = vsfCurve.FixedRows To vsfCurve.Rows - 1
                        lng项目序号 = Val(vsfCurve.TextMatrix(intRow, COL_项目序号))
                        If !分组名 = vsfCurve.TextMatrix(intRow, COL_分组名) And !项目序号 = lng项目序号 Then
                            vsfCurve.TextMatrix(intRow, col_数据) = Space(intRow) & zlCommFun.Nvl(!数值) & Space(intRow)
                            vsfCurve.TextMatrix(intRow, col_颜色) = vsfCurve.TextMatrix(intRow, col_数据)
                            If Not IsNumeric(zlCommFun.Nvl(!数值)) And zlCommFun.Nvl(!数值) <> "不升" And InStr(1, zlCommFun.Nvl(!数值), "/") = 0 Then
                                vsfCurve.TextMatrix(intRow, COL_部位) = ""
                                vsfCurve.TextMatrix(intRow, Col_未记说明) = zlCommFun.Nvl(!未记说明)
                            Else
                                vsfCurve.TextMatrix(intRow, COL_部位) = zlCommFun.Nvl(!部位)
                                vsfCurve.TextMatrix(intRow, Col_未记说明) = ""
                            End If
                            If lng项目序号 = 1 And (IsNumeric(zlCommFun.Nvl(!数值)) Or zlCommFun.Nvl(!数值) <> "不升") Then
                                vsfCurve.TextMatrix(intRow, col_复查) = IIf(Val(zlCommFun.Nvl(!复查)) = 1, "√", "")
                            End If
                            lngColor = 255
                            If InStr(1, ",0,9,", Val(zlCommFun.Nvl(!数据来源))) = 0 Then
                                If zlCommFun.Nvl(!数值) = "不升" And lng项目序号 = 1 Then
                                    lngColor = 255
                                ElseIf lng项目序号 = 1 Or (lng项目序号 = 2 And mint心率应用 = 2 And mblnEdit心率 = True) Then
                                    If InStr(1, zlCommFun.Nvl(!数值), "/") = 0 Then
                                        lngColor = RGB(0, 0, 255)
                                    Else
                                        If Val(!修改) = 0 Then
                                            lngColor = RGB(0, 0, 255)
                                        Else
                                            lngColor = 255
                                        End If
                                    End If
                                End If
                                vsfCurve.Cell(flexcpForeColor, intRow, col_数据, intRow, col_数据) = lngColor
                            Else
                                vsfCurve.Cell(flexcpForeColor, intRow, col_数据, intRow, col_数据) = &H80000012
                            End If
                            marrDate(intRow) = Val(CStr(zlCommFun.Nvl(!数据来源)))
                            If InStr(1, ",0,9,", Val(zlCommFun.Nvl(!数据来源))) = 0 Then
                                blnAllow = True
                            End If
                            mArrModfy(intRow) = Val(!修改)
                            mArrValue(intRow) = Val(!数值)
                            If mbln脉搏共用显示 And InStr(!数值, "/") > 0 Then
                                mArrValue(intRow) = Split(!数值, "/")(1)
                            End If
                            
                        End If
                    Next intRow
                End If
                
                '组织时间字符串,用来判断本段时间内有多少个时间点有数据
                If CDate(Format(strTmp, "YYYY-MM-DD HH:mm:ss")) <> CDate(Format(!时间, "YYYY-MM-DD HH:mm:ss")) Then
                    strTmp = Format(!时间, "YYYY-MM-DD HH:mm:ss")
                    ReDim Preserve mArrdkpTime(UBound(mArrdkpTime) + 1)
                    mArrdkpTime(UBound(mArrdkpTime)) = strTmp
                End If
            .MoveNext
            Loop
        End With
        
        
        If UBound(mArrdkpTime) = -1 Then
            ReDim Preserve mArrdkpTime(UBound(mArrdkpTime) + 1)
            mArrdkpTime(UBound(mArrdkpTime)) = GetCenterTime(CDate(Format(mstrBegin, "YYYY-MM-DD HH:mm:ss")), CDate(Format(mstrEnd, "YYYY-MM-DD HH:mm:ss")))
        End If
        
        dkpTime.Value = Format(CStr(mArrdkpTime(LBound(mArrdkpTime))), "HH:mm")
        dkpTime.Tag = 0
        If UBound(mArrdkpTime) = 0 Then
            For intRow = 0 To picBut.Count - 1
                picBut(intRow).Enabled = False
                picBut(intRow).Visible = False
                picBut1(intRow).Visible = True
                picBut1(intRow).Enabled = False
            Next intRow
        Else
            picBut(0).Visible = False
            picBut(0).Enabled = False
            picBut(1).Visible = False
            picBut(1).Enabled = False
            picBut1(0).Visible = True
            picBut1(0).Enabled = False
            picBut1(1).Visible = True
            picBut1(1).Enabled = False
            picBut(2).Enabled = True
            picBut(2).Visible = True
            picBut(3).Enabled = True
            picBut(3).Visible = True
            picBut1(2).Enabled = False
            picBut1(2).Visible = False
            picBut1(3).Enabled = False
            picBut1(3).Visible = False
        End If
        
        '存在同步过来的数据 时间不允许修改
        If blnAllow = True Or mblnFileBack = True Then
            dkpTime.Enabled = False
        Else
            dkpTime.Enabled = True
        End If
        
        '2----------------------------------------------------------------------------
        '提取手术及上下标说明信息
        
        gstrFields = "序号," & adDouble & ",18|项目序号," & adDouble & ",18|时间," & adLongVarChar & ",20|原始时间," & adLongVarChar & ",20|记录类型," & adDouble & ",1|内容," & _
            adLongVarChar & ",100|项目名称," & adLongVarChar & ",20|未记说明," & adLongVarChar & ",20|记录组号," & adDouble & ",1|数据来源," & adDouble & ",1|显示," & adDouble & ",1|" & _
             "来源ID," & adDouble & ",18|共用," & adDouble & ",1|状态," & adDouble & ",1"
        Call Record_Init(mrsNote, gstrFields)
        gstrFields = "序号|项目序号|时间|原始时间|记录类型|内容|项目名称|未记说明|记录组号|数据来源|显示|来源ID|共用|状态"
        
        mstrSQL = "" & _
             " Select C.ID 序号, B.发生时间 AS 时间,C.记录类型,C.项目序号,C.未记说明,C.记录内容,C.记录组号,C.项目名称,C.数据来源,C.显示,C.来源ID,C.共用" & _
             " FROM 病人护理文件 A, 病人护理数据 B, 病人护理明细 C" & _
             " Where A.ID=B.文件ID and  B.ID = C.记录ID AND A.ID=[1]  AND Nvl(A.婴儿, 0)=[4] AND a.病人id=[2] AND a.主页id=[3] And c.终止版本 Is Null" & _
             " AND c.记录类型 in (2,6)  AND B.发生时间 BETWEEN [5]  And [6]"
             
        If mblnMove Then
            mstrSQL = Replace(mstrSQL, "病人护理文件", "H病人护理文件")
            mstrSQL = Replace(mstrSQL, "病人护理数据", "H病人护理数据")
            mstrSQL = Replace(mstrSQL, "病人护理明细", "H病人护理明细")
        End If
        
        Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, "读取手术、上下标等信息", mT_Patient.lng文件ID, mT_Patient.lng病人ID, mT_Patient.lng主页ID, _
            mT_Patient.lng婴儿, Int(CDate(mstrBegin)), CDate(mstrEnd))
        With rsTmp
            Do While Not .EOF
                gstrValues = zlCommFun.Nvl(!序号) & "|" & zlCommFun.Nvl(!项目序号, 0) & "|" & Format(zlCommFun.Nvl(!时间), "YYYY-MM-DD HH:mm:ss") & "|" & Format(zlCommFun.Nvl(!时间), "YYYY-MM-DD HH:mm:ss") & "|" & zlCommFun.Nvl(!记录类型) & "|" & _
                    zlCommFun.Nvl(!记录内容) & "|" & zlCommFun.Nvl(!项目名称) & "|" & Nvl(!未记说明) & "|" & zlCommFun.Nvl(!记录组号, 0) & "|" & Val(zlCommFun.Nvl(!数据来源, 0)) & "|" & _
                    Val(zlCommFun.Nvl(!显示, 0)) & "|" & Val(zlCommFun.Nvl(!来源ID, 0)) & "|" & Val(zlCommFun.Nvl(!共用, 0)) & "|0"
                Call Record_Add(mrsNote, gstrFields, gstrValues)
            .MoveNext
            Loop
        End With
        
        '添加上下标信息
        mrsNote.Filter = 0
        With mrsNote
            Do While Not .EOF
                    If CDate(Format(!时间, "YYYY-MM-DD HH:mm:ss")) >= CDate(Format(mstrBegin, "YYYY-MM-DD HH:mm:ss")) _
                        And CDate(Format(!时间, "YYYY-MM-DD HH:mm:ss")) <= CDate(Format(mstrEnd, "YYYY-MM-DD HH:mm:ss")) Then
                        Select Case Val(!记录类型)
                            Case 2
                                intRow = mOptRow.上标
                            Case 6
                                intRow = mOptRow.下标
                        End Select
                        vsfCurve.TextMatrix(intRow, col_数据) = Space(intRow) & zlCommFun.Nvl(!内容) & Space(intRow)
                        vsfCurve.Cell(flexcpBackColor, intRow, col_颜色, intRow, col_颜色) = IIf(IsNumeric(Nvl(!未记说明)) = False, 16711680, Val(Nvl(!未记说明)))
                        vsfCurve.TextMatrix(intRow, COL_部位) = ""
                        vsfCurve.TextMatrix(intRow, col_复查) = ""
                        vsfCurve.TextMatrix(intRow, Col_未记说明) = ""
                    End If
            .MoveNext
            Loop
        End With
    End If
    
    '刷新表格数据和手术信息
    If blnTab = True Then
        gstrFields = "序号," & adDouble & ",18|项目序号," & adDouble & ",18|时间," & adLongVarChar & ",20|原始时间," & adLongVarChar & ",20|记录类型," & adDouble & ",1|内容," & _
            adLongVarChar & ",100|项目名称," & adLongVarChar & ",20|未记说明," & adLongVarChar & ",20|记录组号," & adDouble & ",1|数据来源," & adDouble & ",1|显示," & adDouble & ",1|" & _
             "来源ID," & adDouble & ",18|共用," & adDouble & ",1|状态," & adDouble & ",1"
        Call Record_Init(mrsOper, gstrFields)
        gstrFields = "序号|项目序号|时间|原始时间|记录类型|内容|项目名称|未记说明|记录组号|数据来源|显示|来源ID|共用|状态"
        
        '提取手术信息
        mstrSQL = "" & _
             " Select C.ID 序号, B.发生时间 AS 时间,C.记录类型,C.项目序号,C.未记说明,C.记录内容,C.记录组号,C.项目名称,C.数据来源,C.显示,C.来源ID,C.共用" & _
             " FROM 病人护理文件 A, 病人护理数据 B, 病人护理明细 C" & _
             " Where A.ID=B.文件ID and  B.ID = C.记录ID AND A.ID=[1]  AND Nvl(A.婴儿, 0)=[4] AND a.病人id=[2] AND a.主页id=[3] And c.终止版本 Is Null" & _
             " AND c.记录类型=4  AND B.发生时间 BETWEEN [5]  And [6]"
             
        If mblnMove Then
            mstrSQL = Replace(mstrSQL, "病人护理文件", "H病人护理文件")
            mstrSQL = Replace(mstrSQL, "病人护理数据", "H病人护理数据")
            mstrSQL = Replace(mstrSQL, "病人护理明细", "H病人护理明细")
        End If
        
        strTime = CDate(Format(mstrBegin, "YYYY-MM-DD") & " 23:59:59")
        If CDate(strTime) > CDate(mstrETime) Then strTime = Format(mstrETime, "YYYY-MM-DD HH:mm:ss")
        
        Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, "读取手术、上下标等信息", mT_Patient.lng文件ID, mT_Patient.lng病人ID, mT_Patient.lng主页ID, _
            mT_Patient.lng婴儿, Int(CDate(Format(mstrBegin, "YYYY-MM-DD"))), CDate(strTime))
        With rsTmp
            Do While Not .EOF
                gstrValues = zlCommFun.Nvl(!序号) & "|" & zlCommFun.Nvl(!项目序号, 0) & "|" & Format(zlCommFun.Nvl(!时间), "YYYY-MM-DD HH:mm:ss") & "|" & Format(zlCommFun.Nvl(!时间), "YYYY-MM-DD HH:mm:ss") & "|" & zlCommFun.Nvl(!记录类型) & "|" & _
                    zlCommFun.Nvl(!记录内容) & "|" & zlCommFun.Nvl(!项目名称) & "|" & Nvl(!未记说明) & "|" & zlCommFun.Nvl(!记录组号, 0) & "|" & Val(zlCommFun.Nvl(!数据来源, 0)) & "|" & _
                    Val(zlCommFun.Nvl(!显示, 0)) & "|" & Val(zlCommFun.Nvl(!来源ID, 0)) & "|" & Val(zlCommFun.Nvl(!共用, 0)) & "|0"
                Call Record_Add(mrsOper, gstrFields, gstrValues)
            .MoveNext
            Loop
        End With
        
        '添加手术信息
        mrsOper.Filter = 0
        mrsOper.Sort = "时间"
        With mrsOper
            vsfOper.Rows = vsfOper.FixedRows
            Do While Not .EOF
                vsfOper.Rows = vsfOper.Rows + 1
                vsfOper.TextMatrix(vsfOper.Rows - 1, Col_OperTime) = Format(!时间, "HH:mm")
                vsfOper.TextMatrix(vsfOper.Rows - 1, Col_OperType) = Nvl(!项目名称, "手术")
                If InStr(1, ",0,9,", "," & Val(zlCommFun.Nvl(!数据来源)) & ",") = 0 Then
                    vsfOper.Cell(flexcpForeColor, vsfOper.Rows - 1, Col_OperTime, vsfOper.Rows - 1, Col_OperType) = 255
                Else
                    vsfOper.Cell(flexcpForeColor, vsfOper.Rows - 1, Col_OperTime, vsfOper.Rows - 1, Col_OperType) = &H80000012
                End If
                vsfOper.RowData(vsfOper.Rows - 1) = Val(!序号)
            .MoveNext
            Loop
            vsfOper.Rows = vsfOper.Rows + 1
        End With
        
        strItems = ""
        '3------------------------------------------------------------------------------------------------------------
        '组织项目信息
        For intRow = vsfTab.FixedRows To vsfTab.Rows - 1
            lng项目序号 = Val(vsfTab.TextMatrix(intRow, COL_tab项目序号))
            If lng项目序号 <> 4 Then
                i = InStr(1, vsfTab.TextMatrix(intRow, COL_tab项目名称), "(")
                If i > 0 Then
                    strItemName = Trim(Left(vsfTab.TextMatrix(intRow, COL_tab项目名称), i - 1))
                Else
                    strItemName = Trim(vsfTab.TextMatrix(intRow, COL_tab项目名称))
                End If
                If InStr(1, "," & strItems & ",", ",'" & strItemName & "',") = 0 Then
                    strItems = strItems & ",'" & strItemName & "'"
                End If
            End If
        Next intRow
        
        If Left(strItems, 1) = "," Then strItems = Mid(strItems, 2)
        strItems = strItems & ",'收缩压','舒张压'"
        If Left(strItems, 1) = "," Then strItems = Mid(strItems, 2)
        
        '提取一天内(可能含有第二天数据)所有的表格数据信息
        mstrSQL = "SELECT C.Id,a.发生时间 As 时间,C.记录类型,C.显示,C.记录内容 As 结果,C.体温部位,C.未记说明,nvl(C.数据来源,0) 数据来源," & vbNewLine & _
            "  DECODE(E.项目性质,2,C.体温部位 || D.记录名,D.记录名) 项目名称,D.项目序号,C.来源ID,C.共用,E.项目性质 " & _
            "  FROM 病人护理文件 B, 病人护理数据 A,病人护理明细 C,体温记录项目 D,护理记录项目 E " & _
            "  Where B.ID = A.文件ID" & vbNewLine & _
            "  AND A.ID = C.记录ID" & vbNewLine & _
            "  AND B.ID = [1]" & vbNewLine & _
            "  AND Nvl(B.婴儿, 0) = [7]" & vbNewLine & _
            "  AND B.病人id = [2]" & vbNewLine & _
            "  AND B.主页id = [3]" & vbNewLine & _
            "  AND INSTR([6], DECODE(E.项目性质, 2,C.体温部位 || D.记录名, D.记录名)) > 0" & vbNewLine & _
            "  AND D.项目序号 = C.项目序号" & vbNewLine & _
            "  AND Mod(c.记录类型,10) = 1" & vbNewLine & _
            "  AND E.项目序号 = D.项目序号" & vbNewLine & _
            "  AND E.护理等级 >= [8]" & vbNewLine & _
            "  AND A.发生时间 BETWEEN [4] And [5]" & vbNewLine & _
            "  And C.终止版本 Is Null" & vbNewLine & _
            "  AND D.记录法 = 2 And D.项目序号<>3" & vbNewLine & _
            "  UNION ALL "
        '提取非体温表格的汇总项目（体温表格汇总项目子项可能存在非体温项目）
        mstrSQL = mstrSQL & vbNewLine & _
            "  SELECT C.ID,a.发生时间 As 时间,C.记录类型,C.显示,C.记录内容 As 结果,C.体温部位,C.未记说明,nvl(C.数据来源,0) 数据来源," & _
            "   D.项目名称,D.项目序号,C.来源ID,C.共用,D.项目性质" & _
            "   FROM 病人护理文件 B, 病人护理数据 A,病人护理明细 C,(SELECT A.项目序号,A.项目名称, 1 项目性质,B.父序号 FROM 护理记录项目 A,护理汇总项目 B" & vbNewLine & _
            "       WHERE A.项目序号=B.序号 AND NOT EXISTS (SELECT C.项目序号 FROM 体温记录项目 C,护理汇总项目 E WHERE C.项目序号=E.序号 AND C.项目序号=A.项目序号)" & vbNewLine & _
            "       AND NVL(A.应用方式,0)=1 AND NVL(A.护理等级,0)>=[8] AND NVL(A.适用病人,0) IN (0,[9])" & vbNewLine & _
            "       AND (A.适用科室=1 OR (A.适用科室=2 AND EXISTS (SELECT 1 FROM 护理适用科室 D WHERE D.项目序号=A.项目序号 AND D.科室ID=[10])))) D" & _
            "   Where B.ID=A.文件ID And A.ID = C.记录ID   AND B.ID=[1]  AND Nvl(B.婴儿,0)=[7] " & _
            "   AND B.病人id=[2]  AND B.主页id=[3]  AND D.项目序号=C.项目序号  AND C.记录类型=1" & _
            "   AND A.发生时间 BETWEEN [4] And [5] And C.终止版本 Is Null"
            
        mstrSQL = _
            "   Select ID,时间,记录类型,显示,结果,体温部位,未记说明,数据来源,项目名称,项目序号,来源ID,共用,项目性质 From (" & mstrSQL & ")" & _
            "   Order By  Decode(项目名称,'收缩压',0,1)," & strItems & ",时间"
        If mblnMove Then
            mstrSQL = Replace(mstrSQL, "病人护理文件", "H病人护理文件")
            mstrSQL = Replace(mstrSQL, "病人护理数据", "H病人护理数据")
            mstrSQL = Replace(mstrSQL, "病人护理明细", "H病人护理明细")
        End If
        
        strTime = CDate(Format(mstrBegin, "YYYY-MM-DD") & " 23:59:59")
        
        dtBegin = Int(CDate(mstrBegin) - 1)
        dtEnd = CDate(CDate(Format(strTime, "YYYY-MM-DD HH:mm:ss")) + 1)
        If CDate(Format(dtBegin, "YYYY-MM-DD HH:mm:ss")) < CDate(Format(mstrBTime, "YYYY-MM-DD HH:mm:ss")) Then _
            dtBegin = CDate(Format(mstrBTime, "YYYY-MM-DD HH:mm:ss"))
        If CDate(Format(dtEnd, "YYYY-MM-DD HH:mm:ss")) > CDate(Format(mstrETime, "YYYY-MM-DD HH:mm:ss")) Then _
            dtEnd = CDate(Format(mstrETime, "YYYY-MM-DD HH:mm:ss"))
        
        Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, _
                                            mT_Patient.lng文件ID, _
                                            mT_Patient.lng病人ID, _
                                            mT_Patient.lng主页ID, _
                                            CDate(dtBegin), _
                                            CDate(dtEnd), _
                                            strItems, mT_Patient.lng婴儿, mT_Patient.lng护理等级, IIf(mT_Patient.lng婴儿 = 0, 1, 2), mT_Patient.lng科室ID)
        
        gstrFields = "序号|分组名|数值|部位|标记|时间|原始时间|项目序号|项目名称|复查|未记说明|数据来源|修改|显示|来源ID|共用|状态|列号|记录类型"
        gbln出院 = mbln出院
        For intRow = vsfTab.FixedRows To vsfTab.Rows - 1
            If vsfTab.TextMatrix(intRow, COL_tab分组名) = "2)体温表格项目" Then
                int项目性质 = Val(Split(vsfTab.TextMatrix(intRow, COL_tab字符串), ",")(5))
                int记录频次 = Val(Split(vsfTab.TextMatrix(intRow, COL_tab字符串), ",")(3))
                int项目表示 = Val(Split(vsfTab.TextMatrix(intRow, COL_tab字符串), ",")(4))
                int入院首测 = Val(Split(vsfTab.TextMatrix(intRow, COL_tab字符串), ",")(8))
                lng项目序号 = Val(vsfTab.TextMatrix(intRow, COL_tab项目序号))
                str项目名称 = Split(vsfTab.TextMatrix(intRow, COL_tab项目名称), "(")(0)
             
                intNum = 0
                strName = ""
                
                Set rsDownTab = ReturnItemRecord(rsTmp, Int(CDate(mstrBegin)), CDate(mstrBTime), lng项目序号 & ";" & str项目名称 & ";" & _
                                int记录频次 & ";" & int项目表示 & ";" & int项目性质 & ";" & int入院首测, mbln汇总当天, mbln录入小时, True)
                If rsDownTab.RecordCount > 0 Then rsDownTab.MoveFirst
                rsDownTab.Sort = "时间,项目序号,序号"
    
                With rsDownTab
                    Do While Not .EOF
                        blnAdd = False
                        intModify = IIf(InStr(1, ",0,9,", "," & Val(zlCommFun.Nvl(!数据来源)) & ",") = 0, 1, 0)
                        If zlCommFun.Nvl(!序号) <> intNum Or zlCommFun.Nvl(!项目名称) <> strName Then
                            intNum = zlCommFun.Nvl(!序号)
                            strName = zlCommFun.Nvl(!项目名称)
                            '收缩压/舒张压
                            If lng项目序号 = 4 And str项目名称 = "血压" Then
                                Select Case zlCommFun.Nvl(!项目名称)
                                    Case "收缩压"
                                        strParam = ""
                                        strParam = zlCommFun.Nvl(!记录内容)
                                    Case "舒张压"
                                        If InStr(strParam, "/") > 0 Then
                                            strParam = strParam & zlCommFun.Nvl(!记录内容)
                                        Else
                                            strParam = strParam & "/" & zlCommFun.Nvl(!记录内容)
                                        End If
                                        '--问题号:53505,修改人：李涛,血压显示文字
                                        If zlCommFun.Nvl(!记录内容) = "外出" Or zlCommFun.Nvl(!记录内容) = "未测" Or zlCommFun.Nvl(!记录内容) = "拒测" Or zlCommFun.Nvl(!记录内容) = "请假" Then
                                            strParam = zlCommFun.Nvl(!记录内容)
                                        End If
                                        If strParam = "/" Then strParam = ""
                                        blnAdd = True
                                End Select
                            Else
                                strParam = zlCommFun.Nvl(!记录内容)
                                blnAdd = True
                            End If
        
                            If blnAdd = True Then
                                '提取数据时是根据时间段和显示顺序排序的。如果一个时间段有多条数据,只提取前一条
                                mrsCurve.Filter = "分组名='2)体温表格项目' and 项目序号=" & lng项目序号 & " and 项目名称='" & str项目名称 & "' and 列号=" & Val(zlCommFun.Nvl(!序号, 0))
                                '51282,刘鹏飞,2012-08-03,全天汇总显示录入时间(DYEY要求手工录入汇总时间H)
                                If mrsCurve.RecordCount = 0 Then
                                    If Val(Nvl(!汇总小时)) = -1 Then
                                        If InStr(1, strParam, ")") > 0 Then
                                            strName = Replace(Replace(Split(strParam, ")")(0), "(", ""), "h", "")
                                            strParam = Split(strParam, ")")(1)
                                        Else
                                            strName = ""
                                        End If
                                        gstrValues = Val(Split(!汇总小时, ";")(2)) & "|2)体温表格项目|" & strName & "|" & _
                                            zlCommFun.Nvl(!体温部位) & "|0|" & Format(zlCommFun.Nvl(Split(!汇总小时, ";")(1)), "YYYY-MM-DD HH:mm:ss") & "|" & _
                                            Format(zlCommFun.Nvl(Split(!汇总小时, ";")(1)), "YYYY-MM-DD HH:mm:ss") & "|" & lng项目序号 & "|" & str项目名称 & "|0|" & _
                                            Null & "|0|0|1| 0|0|0|" & zlCommFun.Nvl(!序号, 0) & "|11"
                                            Call Record_Add(mrsCurve, gstrFields, gstrValues)
                                    End If
                                    gstrValues = zlCommFun.Nvl(!Id) & "|2)体温表格项目|" & strParam & "|" & _
                                        zlCommFun.Nvl(!体温部位) & "|0|" & Format(zlCommFun.Nvl(!时间), "YYYY-MM-DD HH:mm:ss") & "|" & _
                                        Format(zlCommFun.Nvl(!时间), "YYYY-MM-DD HH:mm:ss") & "|" & lng项目序号 & "|" & str项目名称 & "|0|" & _
                                        zlCommFun.Nvl(!未记说明) & "|" & Val(zlCommFun.Nvl(!数据来源, 0)) & "|" & intModify & "|" & Val(zlCommFun.Nvl(!显示, 0)) & "|" & _
                                        Val(zlCommFun.Nvl(!来源ID, 0)) & "|" & Val(zlCommFun.Nvl(!共用, 0)) & "|0|" & zlCommFun.Nvl(!序号, 0) & "|1"
                                    Call Record_Add(mrsCurve, gstrFields, gstrValues)
                                End If
                                strName = "": strParam = ""
                            End If
                        End If
                    .MoveNext
                    Loop
                End With
            End If
        Next intRow
        
        '展示体温表格数据
        mrsCurve.Filter = 0
        mrsCurve.Filter = "分组名='2)体温表格项目'"
        mrsCurve.Sort = "项目序号,列号,记录类型"
        
        vsfTab.Cell(flexcpText, vsfTab.FixedRows, vsfTab.FixedCols, vsfTab.Rows - 1, vsfTab.Cols - 1) = ""
        strTime = ""
        With mrsCurve
            Do While Not .EOF
                For intRow = vsfTab.FixedRows To vsfTab.Rows - 1
                    blnAllow = False
                    If vsfTab.TextMatrix(intRow, COL_tab项目序号) = !项目序号 And vsfTab.TextMatrix(intRow, COL_tab分组名) = !分组名 Then
                        If Val(Split(vsfTab.TextMatrix(intRow, COL_tab字符串), ",")(5)) = 2 Then
                            If Split(Trim(vsfTab.TextMatrix(intRow, COL_tab项目名称)), "(")(0) <> !项目名称 Then
                                blnAllow = False
                            Else
                                blnAllow = True
                            End If
                        Else
                            blnAllow = True
                        End If
                        If blnAllow = True Then
                            If Val(Nvl(!记录类型)) = 11 Then
                                vsfTab.TextMatrix(intRow, vsfTab.FixedCols + Val(!列号) - 1) = "(" & !数值 & "h)" & vsfTab.TextMatrix(intRow, vsfTab.FixedCols + Val(!列号) - 1)
                            Else
                                vsfTab.TextMatrix(intRow, vsfTab.FixedCols + Val(!列号) - 1) = !数值
                                If Val(zlCommFun.Nvl(!数据来源)) <> 0 Then
                                    vsfTab.Cell(flexcpForeColor, intRow, vsfTab.FixedCols + Val(!列号) - 1, intRow, vsfTab.FixedCols + Val(!列号) - 1) = 255
                                Else
                                    vsfTab.Cell(flexcpForeColor, intRow, vsfTab.FixedCols + Val(!列号) - 1, intRow, vsfTab.FixedCols + Val(!列号) - 1) = &H80000012
                                End If
                                If Val(Split(vsfTab.TextMatrix(intRow, COL_tab字符串), ",")(5)) = 2 And Val(Split(vsfTab.TextMatrix(intRow, COL_tab字符串), ",")(1)) = 1 And _
                                    Val(Split(vsfTab.TextMatrix(intRow, COL_tab字符串), ",")(4)) = 0 Then
                                     vsfTab.Cell(flexcpForeColor, intRow, vsfTab.FixedCols + Val(!列号) - 1, intRow, vsfTab.FixedCols + Val(!列号) - 1) = Val(zlCommFun.Nvl(!未记说明))
                                End If
                            End If
                        End If
                    End If
                Next intRow
            .MoveNext
            Loop
        End With
        
        vsfTab.Cell(flexcpAlignment, vsfTab.FixedRows, vsfTab.FixedCols, vsfTab.Rows - 1, vsfTab.Cols - 1) = flexAlignCenterCenter
    End If
    
    '把未刷新的记录复给原始记录集
    If blnCurve = False Or blnTab = False Then '只刷新体温数据
        If Not rsCurve Is Nothing And rsCurve.State = 1 Then
            rsCurve.Filter = 0
            Do While Not rsCurve.EOF
                mrsCurve.AddNew
                For i = 0 To rsCurve.Fields.Count - 1
                    mrsCurve.Fields(rsCurve.Fields(i).Name).Value = rsCurve.Fields(i).Value
                Next i
                mrsCurve.Update
            rsCurve.MoveNext
            Loop
        End If
    End If
    
    zlRefreshData = True
    
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
    If mrsPart.RecordCount > 0 Then strPart = zlCommFun.Nvl(mrsPart("部位"))
    GetPart = strPart
End Function

Private Function GetCenterTime(ByVal dBegin As Date, ByVal dEnd As Date, Optional intBound As Integer = 0) As String
'------------------------------------------------------------------------------------
'功能:获取某段时间的中点时间,如果当前时间在本段范围并且小与中间时间内则以当前时间为准
'------------------------------------------------------------------------------------
    Dim dblvalue As Double
    Dim strTime As String, strCurDate As String
    Dim i As Integer
    
    strCurDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")
    
    dblvalue = DateDiff("s", dBegin, dEnd)
    strTime = Format(DateAdd("s", Fix(dblvalue / 2), dBegin), "YYYY-MM-DD HH:mm:ss")
    If strTime < mstrBTime Then
        strTime = mstrBTime
    End If
    If strTime > mstrETime Then
        strTime = mstrETime
    End If
    
    For i = 0 To UBound(marrTime)
        If Format(strTime, "HH:mm:ss") >= Format(Split(marrTime(i), ",")(0), "HH:mm:ss") And Format(strTime, "HH:mm:ss") <= Format(Split(marrTime(i), ",")(1), "HH:mm:ss") Then
            Exit For
        End If
    Next i
    If i <= UBound(marrTime) Then
        If gintHourBegin + i * 4 = 24 Then
            strTime = Format(Format(dBegin, "YYYY-MM-DD") & " " & "23:59:59", "YYYY-MM-DD HH:mm:ss")
        Else
            strTime = Format(Format(dBegin, "YYYY-MM-DD") & " " & gintHourBegin + i * 4 & ":00:00", "YYYY-MM-DD HH:mm:ss")
        End If
    End If
    intBound = i
    
    If CDate(strCurDate) >= dBegin And CDate(strCurDate) <= dEnd And CDate(strCurDate) < CDate(strTime) Then
        strTime = strCurDate
    End If
    
    If CDate(strTime) < CDate(mstrBTime) Then
        strTime = Format(mstrBTime, "YYYY-MM-DD HH:mm:ss")
    End If
    If CDate(strTime) > CDate(mstrETime) Then
        strTime = Format(mstrETime, "YYYY-MM-DD HH:mm:ss")
    End If
    GetCenterTime = strTime
End Function

Private Sub picPre_Resize()
    Dim intIndex As Integer
    For intIndex = 0 To picBut.LBound - 1
        picBut1(intIndex).Top = picBut(intIndex).Top
        picBut1(intIndex).Left = picBut(intIndex).Left
        picBut1(intIndex).Width = picBut(intIndex).Width
        picBut1(intIndex).Height = picBut(intIndex).Height
        picBut1(intIndex).Visible = False
    Next intIndex
End Sub

Private Sub picSplit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    picSplit.Tag = 1
    If picSplit.Visible = True Then picSplit.SetFocus
End Sub

Private Sub picSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    If Val(picSplit.Tag) = 0 Then Exit Sub
    
    If picSplit.Top + Y < 4000 Then
        picSplit.Top = 4000
    ElseIf Me.ScaleHeight - (picSplit.Top + Y) < Me.ScaleHeight * 0.3 Then
        picSplit.Top = Me.ScaleHeight * 0.7
    Else
        picSplit.Move picSplit.Left, picSplit.Top + Y
    End If
End Sub

Private Sub picSplit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Val(picSplit.Tag) = 1 Then Call cbsMain_Resize

    picSplit.Tag = 0
End Sub

Private Sub picTab_Resize()
    With FraTable
        .Top = 0
        .Left = 0
        .Width = picTab.Width
        .Height = picTab.Height
    End With
       
    With vsfTab
        .Top = 100
        .Left = 0
        .Width = FraTable.Width
        .Height = FraTable.Height - .Top
    End With
    
    picEdit.Visible = False
    txtEdit.Font.Size = mFontSize + mFontSize * mintBigSize / 3
    lblCheck.Font.Size = txtEdit.Font.Size
    
    With picColor
        .Width = 2190
        .Height = 2190 - 450
        .Visible = False
    End With
    
    With lstSelect(0)
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Visible = False
    End With
    
    With lstSelect(1)
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Visible = False
    End With
    
    With PicLst
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Visible = False
    End With
    
    With txtLst
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Visible = False
    End With
    
    With picHour
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Visible = False
    End With
    
    With lblHour
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Visible = False
    End With
    
    With txtHour
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Visible = False
    End With
End Sub

Private Sub picToolBar_Resize()
    Dim i As Integer
    lblPtime.Left = 0
    lblPtime.Font.Size = mFontSize + mFontSize * mintBigSize / 3
    lblPtime.Top = 45 + 45 * mintBigSize / 3
    
    For i = 0 To 5
        OptTime(i).Font.Size = mFontSize + mFontSize * mintBigSize / 3
        OptTime(i).Height = 300 + 300 * mintBigSize / 3
        OptTime(i).Width = 350 + 350 * mintBigSize / 3
        OptTime(i).Left = i * OptTime(i).Width + lblPtime.Left + lblPtime.Width + 10
    Next i
End Sub

Private Sub picValue_Resize()
    With usrValue
        .Top = -450
        .Left = 0
        .Width = PicValue.Width
        .Height = PicValue.Height
    End With
End Sub

Private Sub stbThis_PanelClick(ByVal Panel As MSComctlLib.Panel)
    Dim strTmp As String
    If Panel.Key = "ZLDataType" Then
        strTmp = "同步数据不能修改-255||同步数据可以修改-" & RGB(0, 0, 255) & "||完全修改-0"
        'frmDataType.ShowPatiType Me, strTmp
        Call zlDatabase.ShowPatiColorTip(Me)
    End If
End Sub

Private Sub tbcThis_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

    If Not mblnInit Then Exit Sub
    
    If Item.Tag = "表格" Then
        If picEdit.Visible = False Then
            Call SetColSelect(True)
        Else
            Call SetColSelect
            txtEdit.SetFocus
        End If
    ElseIf Item.Tag = "曲线" Then
        If mblnStart = False Then
            Call SetColSelect
        Else
            Call SetColSelect(True)
            mblnStart = False
        End If
    End If
    
End Sub

Private Sub tmr1_Timer()
    Dim i As Integer
    Dim strDay As String
    
    '刷新时点按钮显示状态
    
    If mstrBegin = "" Then Exit Sub
    strDay = Format(mstrBegin, "YYYY-MM-DD")
    
    If Format(mstrBegin, "YYYY-MM-DD HH:mm:ss") < Format(mstrBTime, "YYYY-MM-DD HH:mm:ss") Then mstrBegin = mstrBTime
    If Format(mstrEnd, "YYYY-MM-DD HH:mm:ss") > Format(mstrETime, "YYYY-MM-DD HH:mm:ss") Then mstrEnd = mstrETime

    If Format(mstrBegin, "YYYY-MM-DD") = Format(mstrBTime, "YYYY-MM-DD") Or Format(mstrEnd, "YYYY-MM-DD") = Format(mstrETime, "YYYY-MM-DD") Then
        For i = 0 To OptTime.Count - 1
            If OptTime(i).Tag <> "" Then
                If Format(strDay & " " & Split(OptTime(i).Tag, ",")(0), "YYYY-MM-DD HH:mm:ss") > Format(mstrETime, "YYYY-MM-DD HH:mm:ss") Or _
                    Format(strDay & " " & Split(OptTime(i).Tag, ",")(1), "YYYY-MM-DD HH:mm:ss") < Format(mstrBTime, "YYYY-MM-DD HH:mm:ss") Then
                    OptTime(i).Enabled = False
                Else
                    OptTime(i).Enabled = True
                End If
            End If
        Next i
    Else
        For i = 0 To OptTime.Count - 1
            OptTime(i).Enabled = True
        Next i
    End If
End Sub

Private Sub txtEdit_GotFocus()
    Call zlControl.TxtSelAll(txtEdit)
End Sub

Private Sub txtEdit_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim intCOl As Integer, intRow As Integer
    Dim blnAllow As Boolean
    Dim strData As String
    Dim lngColor As Long
    
    If KeyCode = vbKeyDown Then
        If picEdit.Visible = False Then Exit Sub
        '对于类型为文字类型的活动项目使用快捷键可以调用字体颜色设置
        If cmdColor.Visible = True And Shift = vbShiftMask And Val(Split(vsfTab.TextMatrix(vsfTab.Row, COL_tab字符串), ",")(1)) = 1 _
            And Val(Split(vsfTab.TextMatrix(vsfTab.Row, COL_tab字符串), ",")(5)) = 2 And Val(Split(vsfTab.TextMatrix(vsfTab.Row, COL_tab字符串), ",")(4)) = 0 Then
            With picColor
                .Top = picEdit.Top + picEdit.Height
                If .Top + .Height > vsfTab.Top + vsfTab.Height Then
                    .Top = picEdit.Top - .Height
                End If
                If .Top < vsfTab.Top Then .Top = vsfTab.Top
                .Left = picEdit.Left
                .Visible = True
                .ZOrder 0
            End With
            With usrColor
                .Left = 0
                .Top = -450
                .Visible = True
                .ZOrder 0
            End With
            picColor.SetFocus
            usrColor.Color = Val(cmdColor.Tag)
        End If
    ElseIf KeyCode = vbKeyReturn Then
        If Shift = vbShiftMask Then Exit Sub
        '检查数据合法性
        blnAllow = True
        If picEdit.Visible = True And txtEdit.Tag <> "" Then
            intRow = Split(txtEdit.Tag, "|")(0)
            intCOl = Split(txtEdit.Tag, "|")(1)
            
            If txtEdit.Visible = True Then
                strData = IIf(picHour.Visible = True, "(" & txtHour.Text & "h)", "") & Trim(txtEdit.Text)
                lngColor = txtEdit.ForeColor
            Else
                strData = Trim(lblCheck.Caption)
                lngColor = 0
            End If
            
            If strData & "/#$&/" & lngColor <> picEdit.Tag Then blnAllow = WriteIntoVfgTab(strData)
        End If
        If blnAllow = True Then
            '移动到下一列
            Call vsfTab_KeyDown(vbKeyReturn, Shift)
        Else
            Call vsfTab_EnterCell
        End If
    ElseIf KeyCode = vbKeyLeft And txtEdit.SelStart = 0 Then
        If picHour.Visible = False Then
            Call vsfTab_KeyDown(vbKeyLeft, 0)
        Else
            txtHour.SetFocus
        End If
    End If
End Sub

Private Sub picEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        Call txtEdit_KeyPress(KeyAscii)
    ElseIf KeyAscii = vbKeySpace Then
        If lblCheck.Caption = "√" Then
            lblCheck.Caption = ""
        Else
            lblCheck.Caption = "√"
        End If
    ElseIf KeyAscii = vbKeyReturn Then
        Call txtEdit_KeyDown(KeyAscii, 0)
    ElseIf KeyAscii = vbKeyLeft Then
        If txtEdit.Visible = False Then
            Call vsfTab_KeyDown(vbKeyLeft, 0)
        End If
    End If
End Sub

Private Sub txtEdit_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyEscape Then
        With picEdit
            .Visible = False
            .Enabled = False
        End With
        With txtEdit
            .Visible = False
            .Enabled = False
            .Tag = ""
            .Text = ""
        End With
        With cmdColor
            .Visible = False
            .Enabled = False
            .Tag = ""
        End With
        With lstSelect(0)
            .Visible = False
            .Enabled = False
            .Tag = ""
        End With
        With lstSelect(1)
            .Visible = False
            .Enabled = False
            .Tag = ""
        End With
        
        With PicLst
            .Visible = False
            .Tag = ""
        End With
        
        With picHour
            .Visible = False
            .Enabled = False
        End With
        
        With txtHour
            .Visible = False
            .Enabled = False
            .Text = ""
        End With
        
        With lblCheck
            .Visible = False
            .Enabled = False
        End With
        mblnEdit = False
        
        If mblnAllRefresh = False And mblnStart = False Then
            Call SetColSelect
        End If
    End If
End Sub

Private Sub usrColor_LostFocus()
    picColor.Visible = False
End Sub

Private Sub usrColor_pOK()
    Dim intRow As Integer, intCOl As Integer
    Dim strTmp As String, lng项目序号 As Long, str项目名称 As String
    
    If Val(cmdColor.Tag) = usrColor.Color Then picColor.Visible = False:  GoTo GetSetFocus
    cmdColor.Tag = usrColor.Color
    txtEdit.ForeColor = cmdColor.Tag
    picColor.Visible = False
    
    If txtEdit.Tag <> "" Then
        intRow = Val(Split(txtEdit.Tag, "|")(0))
        intCOl = Val(Split(txtEdit.Tag, "|")(1))
    Else
        intRow = vsfTab.Row
        intCOl = vsfTab.Col
    End If
    
    lng项目序号 = Val(vsfTab.TextMatrix(intRow, COL_tab项目序号))
    str项目名称 = Split(vsfTab.TextMatrix(intRow, COL_tab项目名称), "(")(0)

    mrsCurve.Filter = "项目序号=" & lng项目序号 & " and 项目名称='" & str项目名称 & "' and 列号=" & intCOl - vsfTab.FixedCols + 1
    If mrsCurve.RecordCount > 0 Then
        mrsCurve!未记说明 = cmdColor.Tag
        If mrsCurve!状态 <> 1 And mrsCurve!状态 <> 3 Then '原有的数据 修改、删除后的状态始终为2
            mrsCurve!状态 = 2
            mrsCurve!数值 = vsfTab.TextMatrix(intRow, intCOl)
        Else '对于新增数据的处理
            If Trim(vsfTab.TextMatrix(intRow, intCOl)) = "" Then
                mrsCurve!状态 = 3
                mrsCurve!数值 = vsfTab.TextMatrix(intRow, intCOl)
            Else
                mrsCurve!状态 = 1
                mrsCurve!数值 = vsfTab.TextMatrix(intRow, intCOl)
            End If
        End If
        mrsCurve.Update
    End If
    mblnChage = True
    
GetSetFocus:
    If txtEdit.Visible = True Then txtEdit.SetFocus
End Sub

Private Sub usrValue_LostFocus()
    PicValue.Visible = False
End Sub

Private Sub usrValue_pOK()
    If Val(vsfCurve.Cell(flexcpBackColor, usrValue.Tag, col_颜色, usrValue.Tag, col_颜色)) = usrValue.Color Then PicValue.Visible = False: GoTo ErrNext
    vsfCurve.Cell(flexcpBackColor, usrValue.Tag, col_颜色, usrValue.Tag, col_颜色) = usrValue.Color
    If Trim(vsfCurve.TextMatrix(usrValue.Tag, col_数据)) = "" Then GoTo ErrNext
    If Not UpdateCurveDate(usrValue.Tag, col_数据, 2) Then vsfCurve.Cell(flexcpBackColor, usrValue.Tag, col_颜色, usrValue.Tag, col_颜色) = Val(PicValue.Tag)
ErrNext:
    PicValue.Visible = False
    If Val(usrValue.Tag) <= vsfCurve.Rows - 1 Then
        vsfCurve.Body.Select Val(usrValue.Tag), col_数据
    End If
    vsfCurve.SetFocus
End Sub

Private Sub vsfCurve_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strTmp As String
    Dim lng项目序号 As Long
    Dim strDate As String
    Dim lngRect As Long
    On Error Resume Next
    vsfCurve.ComboList(COL_部位) = ""
    vsfCurve.EditMode(COL_部位) = 0
    vsfCurve.EditMode(Col_未记说明) = 0
    lngRect = vsfCurve.Body.FocusRect

    lng项目序号 = Val(vsfCurve.TextMatrix(NewRow, COL_项目序号))
    strDate = Trim(vsfCurve.TextMatrix(NewRow, col_数据))
    Select Case Trim(vsfCurve.TextMatrix(NewRow, COL_分组名))
    
    Case "1)体温曲线项目"
        vsfCurve.EditMode(Col_未记说明) = 1
        If Not mrsPart Is Nothing Then
            mrsPart.Filter = "项目序号=" & lng项目序号
            mrsPart.Sort = "缺省项 DESC"
            With mrsPart
                Do While Not .EOF
                    strTmp = IIf(strTmp = "", zlCommFun.Nvl(!部位), strTmp & "|" & zlCommFun.Nvl(!部位))
                .MoveNext
                Loop
            End With
            If strTmp <> "" Then
                If lng项目序号 = 2 And InStr(1, strTmp, "|") = 0 Then
                    strTmp = " |起搏器"
                End If
                vsfCurve.ComboList(COL_部位) = strTmp
                vsfCurve.EditMode(COL_部位) = 1
            End If
        End If
        
        If NewCol = col_数据 Or NewCol = Col_未记说明 Then
            '数据来源
            If InStr(1, ",0,9,", "," & Val(marrDate(NewRow)) & ",") = 0 Then
                If NewCol = col_数据 Then
                    If lng项目序号 = 1 And strDate = "不升" Then GoTo NotEdit
                    If lng项目序号 = 1 Or (lng项目序号 = 2 And mint心率应用 = 2 And mblnEdit心率 = True) Then
                        If InStr(1, strDate, "/") = 0 Then
                            GoTo GoNext
                        Else
                            If mArrModfy(NewRow) = 0 Then GoTo GoNext
                        End If
                    End If
                End If
            End If
            '数据来源
            If InStr(1, ",0,9,", "," & Val(marrDate(NewRow)) & ",") = 0 Then
NotEdit:
                vsfCurve.EditMode(NewCol) = 0
            Else
GoNext:
                vsfCurve.EditMode(NewCol) = 1
            End If
        End If
        
    Case "2)上下标说明"
        vsfCurve.EditMode(Col_未记说明) = 0
        vsfCurve.EditMode(col_数据) = 1
    End Select
        
    strTmp = ""
    
    If Trim(Split(vsfCurve.TextMatrix(NewRow, COL_字符串), ",")(0)) <> "" Then
        strTmp = "数据范围：" & Trim(Split(vsfCurve.TextMatrix(NewRow, COL_字符串), ",")(0)) & " "
    End If
    
    If Trim(vsfCurve.TextMatrix(NewRow, COL_分组名)) = "1)体温曲线项目" Then
        Select Case lng项目序号
            Case 1 '体温
                strTmp = strTmp & Space(4) & "物理降温表示法38/37"
            Case 2
                If mint心率应用 = 2 And mblnEdit心率 Then strTmp = strTmp & Space(4) & "脉搏短拙表示法100/130"
        End Select
    ElseIf Trim(vsfCurve.TextMatrix(NewRow, COL_分组名)) = "2)上下标说明" Then
'        If lng项目序号 = 4 Then
'            strTmp = "手术行:数据列输入手术时间(如:04:00),部位/手术列选择类型."
'        End If
        strTmp = "在数据列按SHIFT+↓或双击颜色栏进行颜色设置"
    End If
    
    'stbThis.Panels(2).Text = strTmp
    lblStb.Caption = strTmp
    lblStb.ForeColor = &H80000012

End Sub

Private Sub vsfCurve_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngWidth As Long
    If Col = col_颜色 Then
        lngWidth = vsfCurve.Body.ColWidth(Col)
        vsfCurve.Body.ColWidth(col_颜色) = 300
        vsfCurve.Body.ColWidth(col_数据) = vsfCurve.Body.ColWidth(col_数据) + lngWidth - 300
        If vsfCurve.Body.ColWidth(col_数据) < 500 Then vsfCurve.Body.ColWidth(col_数据) = 500
        Call vsfCurve_KeyDown(vbKeyDown, vbShiftMask)
    End If
End Sub

Private Sub vsfCurve_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    Dim blnAllow As Boolean
    Dim intType As Integer
    
    vsfCurve.Tag = vsfCurve.TextMatrix(Row, Col)
    
    Select Case Col
        Case COL_部位
            vsfCurve.TextMatrix(Row, Col) = ""
            If Trim(vsfCurve.TextMatrix(Row, COL_分组名)) = "2)上下标说明" Then
                intType = 2
            ElseIf Trim(vsfCurve.TextMatrix(Row, COL_分组名)) = "1)体温曲线项目" Then
                intType = 1
            End If
            blnAllow = True
        Case Col_未记说明
            If Trim(vsfCurve.TextMatrix(Row, COL_分组名)) = "1)体温曲线项目" And vsfCurve.TextMatrix(Row, Col) <> "" Then
                vsfCurve.TextMatrix(Row, Col) = ""
                vsfCurve.TextMatrix(Row, col_数据) = Space(Row) & IIf(vsfCurve.TextMatrix(Row, COL_分组名) = "2)上下标说明", "", "") & Space(Row)
                vsfCurve.TextMatrix(Row, col_颜色) = Space(Row) & IIf(vsfCurve.TextMatrix(Row, COL_分组名) = "2)上下标说明", " ", "") & Space(Row)
                vsfCurve.TextMatrix(Row, COL_部位) = ""
                vsfCurve.TextMatrix(Row, col_复查) = ""
                blnAllow = True
                intType = 1
            End If
        Case col_数据
            If vsfCurve.TextMatrix(Row, Col) <> "" Then
                If Trim(vsfCurve.TextMatrix(Row, COL_分组名)) = "2)上下标说明" Then
                    intType = 2
                ElseIf Trim(vsfCurve.TextMatrix(Row, COL_分组名)) = "1)体温曲线项目" Then
                    intType = 1
                    If InStr(1, ",0,9,", "," & Val(marrDate(Row)) & ",") = 0 Then
                        Cancel = True
                        lblStb.Caption = "由护理记录单或其它地方同步过来的数据不能删除."
                        lblStb.ForeColor = 255
                        vsfCurve.Body.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
                        Exit Sub
                    End If
                End If
                
                vsfCurve.TextMatrix(Row, col_数据) = Space(Row) & IIf(vsfCurve.TextMatrix(Row, COL_分组名) = "2)上下标说明", "", "") & Space(Row)
                vsfCurve.TextMatrix(Row, col_颜色) = Space(Row) & IIf(vsfCurve.TextMatrix(Row, COL_分组名) = "2)上下标说明", " ", "") & Space(Row)
                vsfCurve.TextMatrix(Row, COL_部位) = ""
                vsfCurve.TextMatrix(Row, col_复查) = ""
                vsfCurve.TextMatrix(Row, Col_未记说明) = ""
                
                blnAllow = True
            End If
    End Select
    
    If blnAllow = True Then Call UpdateCurveDate(Row, Col, intType)
    Cancel = True
End Sub

Private Sub vsfCurve_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub vsfCurve_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
'选择未记说明
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim i As Integer
    Dim blnSelect As Boolean
    
    If Trim(vsfCurve.TextMatrix(Row, COL_分组名)) <> "1)体温曲线项目" Then Exit Sub
    
    On Error GoTo Errhand
    Select Case Col
        Case Col_未记说明
            pic未记.Tag = Row & "|" & Col
            
            strSQL = "Select 编码,名称 From 常用体温说明"
            Call zlDatabase.OpenRecordset(rsTemp, strSQL, Me.Caption)
            lst未记.Clear
            If rsTemp.RecordCount > 0 Then
                i = 0
                With rsTemp
                    Do While Not .EOF
                        lst未记.AddItem zlCommFun.Nvl(!名称)
                        If zlCommFun.Nvl(!名称) = vsfCurve.TextMatrix(vsfCurve.Row, vsfCurve.Col) Then
                            lst未记.Selected(i) = True
                            blnSelect = True
                        End If
                        i = i + 1
                    .MoveNext
                    Loop
                End With
            End If
            
            If blnSelect = False And lst未记.ListCount <> 0 Then lst未记.Selected(0) = True
            
            If lst未记.ListCount > 0 Then
                pic未记.Left = vsfCurve.CellLeft + vsfCurve.Left + 15
                pic未记.Top = vsfCurve.CellTop + vsfCurve.Top + vsfCurve.CellHeight
                If lst未记.Height < vsfCurve.CellHeight + 20 Then lst未记.Height = vsfCurve.CellHeight + 20
                lst未记.Width = vsfCurve.CellWidth + 20
                pic未记.Height = lst未记.Height
                pic未记.Width = lst未记.Width
                
                If pic未记.Top + pic未记.Height > vsfCurve.Top + vsfCurve.Height Then
                    pic未记.Top = vsfCurve.CellTop + vsfTab.Top - pic未记.Height
                End If
                pic未记.Visible = True
                lst未记.Visible = True: lst未记.Enabled = True
                lst未记.SetFocus
            End If
    End Select
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsfCurve_ChangeEdit()
    Select Case vsfCurve.Col
    Case col_数据
        If vsfCurve.TextMatrix(vsfCurve.Row, COL_项目序号) <> 0 Then
            vsfCurve.TextMatrix(vsfCurve.Row, col_数据) = IIf(vsfCurve.EditText = "", " ", vsfCurve.EditText)
            If vsfCurve.TextMatrix(vsfCurve.Row, COL_分组名) <> "2)上下标说明" Then
                vsfCurve.TextMatrix(vsfCurve.Row, col_颜色) = vsfCurve.TextMatrix(vsfCurve.Row, col_数据)
            End If
            If vsfCurve.EditText <> "" Then vsfCurve.TextMatrix(vsfCurve.Row, Col_未记说明) = ""
        End If
    End Select
End Sub

Private Sub vsfCurve_ComboCloseUp(Row As Long, Col As Long, FinishEdit As Boolean)
    Dim intType As Integer
    Dim blnAllow As Boolean
        
    blnAllow = True
    If Trim(vsfCurve.TextMatrix(Row, COL_分组名)) = "1)体温曲线项目" Then
        intType = 1
    ElseIf Trim(vsfCurve.TextMatrix(Row, COL_分组名)) = "2)上下标说明" Then
        If Val(vsfCurve.TextMatrix(Row, COL_项目序号)) = 4 And vsfCurve.EditText <> "" Then
'            intType = 2
'
'            If Trim(vsfCurve.TextMatrix(Row, col_数据)) = "" Then
'                vsfCurve.TextMatrix(Row, col_数据) = Format(GetCenterTime(CDate(mstrBegin), CDate(mstrEnd)), "HH:mm")
'            End If
            blnAllow = False
        Else
            blnAllow = False
        End If
    End If
    If blnAllow = True Then Call UpdateCurveDate(Row, Col, intType, True)
End Sub

Private Sub vsfCurve_KeyDown(KeyCode As Integer, Shift As Integer)
    PicValue.Visible = False
    PicValue.Tag = ""
    With vsfCurve
        If .Col > .FixedCols - 1 And .Row > .FixedRows - 1 Then
            If KeyCode = vbKeyDown And Shift = vbShiftMask Then
                If .Col = Col_未记说明 Then
                    Call vsfCurve_CellButtonClick(.Row, .Col)
                ElseIf (.Col = col_数据 Or .Col = col_颜色) And .TextMatrix(.Row, COL_分组名) = "2)上下标说明" Then
                    vsfCurve.Tag = .TextMatrix(.Row, col_数据)
                    PicValue.Top = .CellTop + .CellHeight + .Top
                    If PicValue.Top + PicValue.Height > .Top + .Height Then
                        PicValue.Top = .CellTop - PicValue.Height
                    End If
                    If PicValue.Top < .Top Then PicValue.Top = .Top
                    PicValue.Left = IIf(.Col = col_颜色, .CellLeft, .CellLeft + .CellWidth) + .Left
                    PicValue.Visible = True
                    PicValue.ZOrder 0
         
                    usrValue.Left = 0
                    usrValue.Top = -450
                    usrValue.Visible = True
                    usrValue.ZOrder 0
                    PicValue.SetFocus
                    usrValue.Color = Val(.Cell(flexcpBackColor, .Row, col_颜色, .Row, col_颜色))
                    PicValue.Tag = Val(usrValue.Color)
                    usrValue.Tag = .Row
                End If
            End If
        End If
    End With
End Sub

Private Sub vsfCurve_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    If KeyCode = vbKeyReturn Then
        If Col = Col_未记说明 Then
            If InStr(1, "," & mstr未记说明 & ",", "," & vsfCurve.EditText & ",") = 0 Then
                vsfCurve.TextMatrix(Row, Col) = ""
                vsfCurve.Cell(flexcpData, Row, Col) = ""
            Else
                vsfCurve.TextMatrix(Row, Col) = vsfCurve.EditText
                vsfCurve.Cell(flexcpData, Row, Col) = vsfCurve.EditText
                vsfCurve.TextMatrix(Row, col_数据) = Space(Row) & IIf(vsfCurve.TextMatrix(Row, COL_分组名) = "2)上下标说明", "", "") & Space(Row)
                vsfCurve.TextMatrix(Row, col_颜色) = Space(Row) & IIf(vsfCurve.TextMatrix(Row, COL_分组名) = "2)上下标说明", " ", "") & Space(Row)
                vsfCurve.TextMatrix(Row, COL_部位) = ""
                vsfCurve.TextMatrix(Row, col_复查) = ""
            End If
        End If
    End If
    If KeyCode = vbKeyDown And Shift = vbShiftMask And Col = col_数据 Then
        Call vsfCurve_KeyDown(KeyCode, Shift)
        Cancel = True
    End If
End Sub

Private Sub vsfCurve_KeyPress(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer, Cancel As Boolean)
    If KeyAscii = 32 Then '
        If Col = col_复查 Then
            If Val(vsfCurve.TextMatrix(Row, col_数据)) <> 0 And Val(vsfCurve.TextMatrix(Row, COL_项目序号)) = 1 Then
                If vsfCurve.TextMatrix(Row, Col) = "" Then
                    vsfCurve.TextMatrix(Row, Col) = "√"
                    vsfCurve.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
                Else
                    vsfCurve.TextMatrix(Row, Col) = ""
                End If
                Call UpdateCurveDate(Row, Col, 1)
            End If
        End If
        If Col = col_颜色 And vsfCurve.TextMatrix(Row, COL_分组名) = "2)上下标说明" Then
            Call vsfCurve_KeyDown(vbKeyDown, vbShiftMask)
        End If
    End If
End Sub

Private Sub vsfCurve_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim lngNO As Long
    Dim strDate As String
    
    On Error Resume Next
    lngNO = Val(vsfCurve.TextMatrix(Row, COL_项目序号))
    strDate = vsfCurve.TextMatrix(Row, COL_tab项目名称)
    
    If KeyAscii <> vbKeyReturn Then
        If lngNO <> 0 Then
            If vsfCurve.TextMatrix(Row, COL_分组名) = "1)体温曲线项目" Then
                If Col <> Col_未记说明 Then
                    If lngNO = 2 And mint心率应用 = 2 And mblnEdit心率 = True Then
                        If FilterKeyAscii(KeyAscii, 99, "0123456789./") = 0 Then KeyAscii = 0
                    ElseIf lngNO = 1 Then
                        '体温不进行检查
                    Else
                        If FilterKeyAscii(KeyAscii, 99, "0123456789.") = 0 Then KeyAscii = 0
                    End If
                Else
                    If FilterKeyAscii(KeyAscii, 99, "'") > 0 Then KeyAscii = 0
                End If
            ElseIf vsfCurve.TextMatrix(Row, COL_分组名) = "2)上下标说明" And lngNO = 4 Then
'                If Col = col_数据 Then
'                    If FilterKeyAscii(KeyAscii, 99, "0123456789:") = 0 Then KeyAscii = 0
'                End If
            End If
        End If
    End If
End Sub

Private Sub vsfCurve_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lng项目序号 As Long, strDate As String
    Dim strName As String
    Dim intRow As Integer
    Dim strData As String
    
    lng项目序号 = Val(vsfCurve.TextMatrix(Row, COL_项目序号))
    strName = vsfCurve.TextMatrix(Row, COL_项目名称)
    vsfCurve.Body.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignLeftCenter
        
    Select Case Col
        Case col_数据
            vsfCurve.TextMatrix(Row, Col) = IIf(RTrim(LTrim(vsfCurve.TextMatrix(Row, Col))) = "", " ", RTrim(LTrim(vsfCurve.TextMatrix(Row, Col))))
            If Row <> mOptRow.上标 And Row <> mOptRow.下标 Then
                vsfCurve.TextMatrix(Row, col_颜色) = vsfCurve.TextMatrix(Row, Col)
            Else
                vsfCurve.TextMatrix(Row, Col) = RTrim(LTrim(vsfCurve.TextMatrix(Row, Col)))
            End If
            vsfCurve.Body.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000018
            strDate = RTrim(LTrim(vsfCurve.TextMatrix(Row, Col)))
    End Select
    
    vsfCurve.Tag = RTrim(LTrim(vsfCurve.TextMatrix(Row, Col)))
     
    If Col = col_数据 Or Col = Col_未记说明 Then
        '数据来源
        If InStr(1, ",0,9,", "," & Val(marrDate(Row)) & ",") = 0 Then
            If Col = col_数据 Then
                If lng项目序号 = 1 And strDate = "不升" Then GoTo NotEdit
                If lng项目序号 = 1 Or (lng项目序号 = 2 And mint心率应用 = 2 And mblnEdit心率 = True) Then
                    If InStr(1, strDate, "/") = 0 Then
                        GoTo GoNext
                    Else
                        If mArrModfy(Row) = 0 Then GoTo GoNext
                    End If
                End If
            End If
NotEdit:
            Cancel = True
            vsfCurve.Body.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
            If lng项目序号 = 1 Then
                lblStb.Caption = "同步过来的[" & strName & "]数据只允许修改物理降温部分."
            ElseIf lng项目序号 = 2 And mint心率应用 = 2 And mblnEdit心率 = True Then
                lblStb.Caption = "同步过来的[" & strName & "]数据只允许修改脉搏短轴部分."
            Else
                lblStb.Caption = "由护理记录单或其它地方同步过来的数据不能修改"
            End If
            lblStb.ForeColor = 255
            vsfCurve.Body.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
        End If
    ElseIf col_复查 = Col Then
        vsfCurve.Body.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
    End If
GoNext:
    If mblnFileBack = True Then
        Cancel = True
        vsfCurve.Body.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
        lblStb.Caption = "病人体温数据已经归档,不允许进行数据修改."
        lblStb.ForeColor = 255
        vsfCurve.Body.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
    End If
    
    If Not CheckDateTime(Row, strName, Format(dkpDate.Value & " " & dkpTime.Value, "YYYY-MM-DD HH:mm:ss")) Then
        Cancel = True
        vsfCurve.Body.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
        vsfCurve.Body.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
    End If
End Sub

Private Sub vsfCurve_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strSpace As String
    vsfCurve.Body.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
    vsfCurve.Body.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
    If Col = col_数据 Then
        vsfCurve.TextMatrix(Row, col_数据) = Space(Row) & vsfCurve.TextMatrix(Row, col_数据) & Space(Row)
        vsfCurve.TextMatrix(Row, col_颜色) = IIf(vsfCurve.TextMatrix(Row, COL_分组名) = "2)上下标说明", Space(Row + 1), vsfCurve.TextMatrix(Row, col_数据))
    End If
End Sub

Private Sub vsfCurve_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim str值域 As String
    Dim lngNO As Long, int小数 As Integer, intType As Integer
    Dim strInfo As String, strText As String, strName As String, strMsg As String, strDate As String
    Dim arrValue() As String
    Dim lngCount As Long, i As Long, strValue As String
    Dim blnOk As Boolean
    
    '检查数据合法性
    If Col = col_数据 Then
        strValue = vsfCurve.Tag
        Select Case vsfCurve.TextMatrix(Row, COL_分组名)
            Case "1)体温曲线项目"
                str值域 = Split(vsfCurve.TextMatrix(Row, COL_字符串), ",")(0)
                lngNO = Val(vsfCurve.TextMatrix(Row, COL_项目序号))
                strName = vsfCurve.TextMatrix(Row, COL_项目名称)
                int小数 = Val(Split(vsfCurve.TextMatrix(Row, COL_字符串), ",")(2))
                intType = 1
                GoTo CheckPoint
            Case "2)上下标说明"
                If InStr(1, ",2,6,", "," & Val(vsfCurve.TextMatrix(Row, COL_项目序号)) & ",") <> 0 Then
                    PicValue.Tag = vsfCurve.Cell(flexcpBackColor, Row, col_颜色, Row, col_颜色)
                    intType = 2: GoTo CheckTag
                End If
        End Select
    End If
    
    Exit Sub
CheckPoint:
    strDate = vsfCurve.EditText
    If Trim(vsfCurve.EditText) <> "" And str值域 <> "" Then
        strInfo = vsfCurve.EditText
        
        '脉搏短轴是如果有/则要求必须输入心率
        If lngNO = 2 And mint心率应用 = 2 And mblnEdit心率 = True Then
            If InStr(1, strInfo, "/") > 0 Then
                If Split(Trim(strInfo), "/")(1) = "" Or Split(Trim(strInfo), "/")(0) = "" Then
                    strMsg = strName & "数据录入错误" & Space(4) & "脉搏短轴:脉搏/心率"
                    GoTo ErrInfo
                Else
                    If Not IsNumeric(Split(Trim(strInfo), "/")(0)) Or Not IsNumeric(Split(Trim(strInfo), "/")(1)) Then
                        strMsg = strName & "数据录入错误" & Space(4) & "有效范围:" & str值域
                        GoTo ErrInfo
                    End If
                End If
            End If
        End If
        
        If lngNO <> 1 And Not (lngNO = 2 And mint心率应用 = 2 And mblnEdit心率 = True) Then
            If InStr(1, strInfo, "/") Then
                strMsg = strName & "数据录入错误" & Space(4) & "有效范围:" & str值域
                GoTo ErrInfo
            End If
        End If

        If UBound(Split(strInfo, "/")) > 1 Then
            strMsg = strName & "数据录入错误，请检查！"
            GoTo ErrInfo
        End If
        
        '检查数据在有效范围内是否有效
        arrValue = Split(strInfo, "/")
        lngCount = UBound(arrValue)
        For i = 0 To lngCount
            blnOk = False
            strText = arrValue(i)
            If i = 0 Then
                '体温曲线项目需要过滤掉未记说明
                If InStr(1, strText, ";") <> 0 And UBound(arrValue) = 0 Then strText = Split(strText, ";")(1)
                If InStr(1, IIf(lngNO = 1, ",不升,", ""), "," & strText & ",") = 0 Then
                    blnOk = False
                Else
                    blnOk = True
                End If
            End If
            
            If Not blnOk Then
                If Not IsNumeric(strText) Then
                    strMsg = strName & "数据录入错误" & Space(4) & "有效范围:" & str值域
                    GoTo ErrInfo
                End If
            End If
            
            If Not blnOk And strText <> "" Then strText = Format(Val(strText), "#0" & IIf(int小数 > 0, ".", "") & String(int小数, "0"))
            If IsNumeric(Split(str值域, "～")(0)) And IsNumeric(strText) Then
                If Not (Val(strText) >= Split(str值域, "～")(0) And Val(strText) <= Split(str值域, "～")(1)) Then
                    strMsg = strName & "超出有效范围(" & str值域 & "),请检查!"
                    GoTo ErrInfo
                End If
            End If
        Next i
    End If
    
    '对于数据来源<>0,9的 体温,脉搏数据 进行编辑(无物理降温和脉搏短轴可以录入物理降温,脉搏短轴)
    If InStr(1, ",0,9,", "," & Val(marrDate(Row)) & ",") = 0 Then
        If Col = col_数据 Then
            If lngNO = 1 Or (lngNO = 2 And mint心率应用 = 2 And mblnEdit心率 = True) Then
                '--问题号:56853,修改人：李涛，问题描述：检查脉搏短轴录入方式是否正确，心率/脉搏
                If (InStr(strValue, "/") > 0 Or InStr(strValue, "/") = 0) And mbln脉搏共用显示 Then
                    If InStr(1, strDate, "/") <> 0 Then
                        strDate = Split(strDate, "/")(1)
                    Else
                        strDate = Split(strDate, "/")(0)
                    End If
                     If strDate <> mArrValue(Row) Then
                        If lngNO = 1 Then
                            strMsg = "同步过来的[" & strName & "]数据只允许修改物理降温部分."
                        Else
                            strMsg = "同步过来的[" & strName & "]数据只允许修改脉搏短轴部分."
                        End If
                        
                        vsfCurve.TextMatrix(Row, col_数据) = Space(Row) & Trim(CStr(mArrValue(Row))) & Space(Row)
                        vsfCurve.TextMatrix(Row, col_颜色) = vsfCurve.TextMatrix(Row, col_数据)
                        GoTo ErrInfo
                    End If
                Else
                    strValue = CStr(mArrValue(Row))
                    If InStr(1, strDate, "/") <> 0 Then
                        strDate = Split(strDate, "/")(0)
                    End If
                
                    If InStr(1, mArrValue(Row), "/") = 0 Then
                        If strDate <> mArrValue(Row) Then
                            If lngNO = 1 Then
                                strMsg = "同步过来的[" & strName & "]数据只允许修改物理降温部分."
                            Else
                                strMsg = "同步过来的[" & strName & "]数据只允许修改脉搏短轴部分."
                            End If
                            
                            vsfCurve.TextMatrix(Row, col_数据) = Space(Row) & Trim(CStr(mArrValue(Row))) & Space(Row)
                            vsfCurve.TextMatrix(Row, col_颜色) = vsfCurve.TextMatrix(Row, col_数据)
                            GoTo ErrInfo
                        End If
                    Else
                        If mArrModfy(Row) <> 0 Then
                            If strDate <> mArrValue(Row) Then
                                If lngNO = 1 Then
                                    strMsg = "同步过来的[" & strName & "]数据如果包括物理降温,不允许修改."
                                Else
                                    strMsg = "同步过来的[" & strName & "]数据如果包括脉搏短轴,不允许修改."
                                End If
                                vsfCurve.TextMatrix(Row, col_数据) = Space(Row) & CStr(mArrValue(Row)) & Space(Row)
                                vsfCurve.TextMatrix(Row, col_颜色) = vsfCurve.TextMatrix(Row, col_数据)
                                GoTo ErrInfo
                            End If
                        Else
                            If strDate <> Split(mArrValue(Row), "/")(0) Then
                                If lngNO = 1 Then
                                    strMsg = "同步过来的[" & strName & "]数据只允许修改物理降温部分."
                                Else
                                    strMsg = "同步过来的[" & strName & "]数据只允许修改脉搏短轴部分."
                                End If
                                vsfCurve.TextMatrix(Row, col_数据) = Space(Row) & CStr(mArrValue(Row)) & Space(Row)
                                vsfCurve.TextMatrix(Row, col_颜色) = vsfCurve.TextMatrix(Row, col_数据)
                                GoTo ErrInfo
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

    '显示缺省部位
    If vsfCurve.TextMatrix(Row, COL_部位) = "" And Trim(vsfCurve.TextMatrix(Row, col_数据)) <> "" Then
        mrsPart.Filter = "项目序号=" & lngNO & " and 缺省项=1"
        If mrsPart.RecordCount > 0 Then
            vsfCurve.TextMatrix(Row, COL_部位) = CStr(zlCommFun.Nvl(mrsPart!部位))
        End If
    End If
    
    GoTo ErrSaveData
    Exit Sub
CheckTag:
    GoTo ErrSaveData
    Exit Sub
ErrInfo:    '错误信息输出
    'stbThis.Panels(2).Text = StrMsg
    lblStb.Caption = strMsg
    lblStb.ForeColor = 255
    vsfCurve.TextMatrix(Row, col_数据) = Space(Row) & strValue & Space(Row)
    vsfCurve.TextMatrix(Row, col_颜色) = vsfCurve.TextMatrix(Row, col_数据)
    Cancel = True
    Exit Sub
ErrSaveData:
     Call UpdateCurveDate(Row, Col, intType)
End Sub

Private Function UpdateCurveDate(ByVal intRow As Integer, ByVal intCOl As Integer, ByVal intType As Integer, _
    Optional blnComList As Boolean = False, Optional blnOper As Boolean = False) As Boolean
'------------------------------------------------------------------------
'功能:进行体温项目.手术.上下标的数据保存
'------------------------------------------------------------------------
    Dim lngNO As Long, strName As String, strTime As String, lngID As Long
    Dim strValue As String, int标记 As Integer, str未记 As String
    Dim str部位 As String
    Dim strData As String
    On Error GoTo Errhand:
    
    If Not blnOper Then
        lngNO = Val(vsfCurve.TextMatrix(intRow, COL_项目序号))
        If UBound(Split(vsfCurve.TextMatrix(intRow, COL_项目名称), "(")) = -1 Then
            strName = vsfCurve.TextMatrix(intRow, COL_项目名称)
        Else
            strName = Split(vsfCurve.TextMatrix(intRow, COL_项目名称), "(")(0)
        End If
        
        If blnComList = True Then
            str部位 = vsfCurve.EditText
            If str部位 = "" Then str部位 = vsfCurve.TextMatrix(intRow, COL_部位)
        Else
            str部位 = vsfCurve.TextMatrix(intRow, COL_部位)
        End If
    Else
        lngNO = 4
        If blnComList = True Then
            strName = vsfOper.EditText
            strTime = Format(vsfOper.TextMatrix(intRow, Col_OperTime), "HH:mm")
        Else
            strName = vsfOper.TextMatrix(intRow, Col_OperType)
            strTime = Format(vsfOper.EditText, "HH:mm")
        End If
        str部位 = ""
        
    End If
    If intType = 1 Then '体温数据处理
        strValue = Trim(vsfCurve.TextMatrix(intRow, col_数据))
       If lngNO = 2 And mint心率应用 = 2 And mblnEdit心率 = True Then
            '反转脉搏和心率数据
            If mbln脉搏共用显示 And InStr(strValue, "/") > 0 Then
                strData = Split(Trim(vsfCurve.TextMatrix(intRow, col_数据)), "/")(1) & "/" & Split(Trim(vsfCurve.TextMatrix(intRow, col_数据)), "/")(0)
                strValue = strData
            End If
        End If
        
        str未记 = Trim(vsfCurve.TextMatrix(intRow, Col_未记说明))
        If strValue <> "" Then str未记 = ""
        '进行数据更新处理
        mrsCurve.Filter = "项目序号=" & lngNO & " and 时间='" & Format(mArrdkpTime(dkpTime.Tag), "YYYY-MM-DD HH:mm:ss") & "'"
        
        If mrsCurve.RecordCount <> 0 Then
            If Val(mrsCurve!状态) <> 1 And Val(mrsCurve!状态) <> 3 Then
                mrsCurve!状态 = 2
                mrsCurve!数值 = strValue
                mrsCurve!部位 = str部位
                mrsCurve!复查 = IIf(vsfCurve.TextMatrix(intRow, col_复查) = "√", 1, 0)
                mrsCurve!修改 = 0
                mArrModfy(intRow) = 0
                mrsCurve!未记说明 = str未记
                
            Else
                If strValue = "" And str未记 = "" Then
                    mrsCurve!状态 = 3
                Else
                    mrsCurve!状态 = 1
                End If

                mrsCurve!数值 = strValue
                mrsCurve!部位 = str部位
                mrsCurve!复查 = IIf(vsfCurve.TextMatrix(intRow, col_复查) = "√", 1, 0)
                mrsCurve!未记说明 = str未记
            End If
            mrsCurve.Update
        Else '新增数据
            If strValue <> "" Or str未记 <> "" Then
                gstrFields = "序号|分组名|数值|部位|标记|时间|项目序号|项目名称|复查|未记说明|数据来源|修改|显示|来源ID|共用|状态|列号|记录类型"
                gstrValues = GetMaxID & "|1)体温曲线项目|" & strValue & "|" & str部位 & "|" & _
                    int标记 & "|" & Format(mArrdkpTime(dkpTime.Tag), "YYYY-MM-DD HH:mm:ss") & "|" & lngNO & "|" & strName & "|" & _
                    Val(vsfCurve.TextMatrix(intRow, col_复查)) & "|" & str未记 & "|0|0|0|0|0|1|0|1"
                Call Record_Add(mrsCurve, gstrFields, gstrValues)
            End If
        End If
        
    ElseIf intType = 2 Then '手术上下标处理
    
        If Not blnOper Then
            strValue = LTrim(RTrim(vsfCurve.TextMatrix(intRow, col_数据)))
            mrsNote.Filter = "记录类型=" & lngNO
            If mrsNote.RecordCount <> 0 Then
                If Val(mrsNote!状态) <> 1 And Val(mrsNote!状态) <> 3 Then 'his提取的数据
                    mrsNote!状态 = 2
                    mrsNote!内容 = LTrim(RTrim(vsfCurve.TextMatrix(intRow, col_数据)))
                    mrsNote!未记说明 = IIf(mrsNote!内容 = "", "", vsfCurve.Cell(flexcpBackColor, intRow, col_颜色, intRow, col_颜色))
                Else
                    If strValue = "" Then
                        mrsNote!状态 = 3
                        mrsNote!内容 = strValue
                        mrsNote!未记说明 = ""
                    Else
                        mrsNote!状态 = 1
                        mrsNote!内容 = strValue
                        mrsNote!未记说明 = IIf(mrsNote!内容 = "", "", vsfCurve.Cell(flexcpBackColor, intRow, col_颜色, intRow, col_颜色))
                    End If
                End If
                mrsNote.Update
            Else
                If lngNO = 2 Then
                    strName = "上标说明"
                ElseIf lngNO = 6 Then
                    strName = "下标说明"
                End If
                strTime = GetCenterTime(CDate(mstrBegin), CDate(mstrEnd))
                
                If strValue <> "" Then
                    lngID = GetMaxID
                    gstrFields = "序号|项目序号|时间|原始时间|记录类型|内容|项目名称|未记说明|记录组号|数据来源|显示|来源ID|共用|状态"
                    gstrValues = lngID & "|" & 0 & "|" & strTime & "|" & strTime & "|" & lngNO & "|" & strValue & "|" & strName & "|" & IIf(lngNO = 4, "", vsfCurve.Cell(flexcpBackColor, intRow, col_颜色, intRow, col_颜色)) & "|0|0|0|0|0|1"
                    Call Record_Add(mrsNote, gstrFields, gstrValues)
                End If
            End If
        Else
            mrsOper.Filter = "记录类型=" & lngNO & " And 序号=" & Val(vsfOper.RowData(intRow))
            If mrsOper.RecordCount <> 0 Then
                If Val(mrsOper!状态) <> 1 And Val(mrsOper!状态) <> 3 Then 'his提取的数据
                    mrsOper!状态 = 2
                    If Trim(strTime) = "" Or strName = "" Then
                       mrsOper!项目名称 = ""
                       mrsOper!内容 = ""
                    ElseIf Trim(strTime) <> "" And strName <> "" Then
                        mrsOper!项目名称 = strName
                        mrsOper!内容 = strName
                    End If
                    If Trim(strTime) <> "" Then mrsOper!时间 = SetDate(Format(Format(mstrBegin, "YYYY-MM-DD") & " " & Trim(strTime) & ":00", "YYYY-MM-DD HH:mm:ss"))
                Else
                    If Trim(strTime) = "" Or strName = "" Then
                        mrsOper!状态 = 3
                        mrsOper!项目名称 = ""
                        mrsOper!内容 = ""
                    Else
                        mrsOper!状态 = 1
                        mrsOper!项目名称 = strName
                        mrsOper!内容 = strName
                    End If
                    If Trim(strTime) <> "" Then mrsOper!时间 = SetDate(Format(Format(mstrBegin, "YYYY-MM-DD") & " " & Trim(strTime) & ":00", "YYYY-MM-DD HH:mm:ss"))
                End If
                mrsOper.Update
            Else
                
                If Trim(strTime) = "" Or strName = "" Then
                    strValue = ""
                Else
                    strValue = 1
                    strTime = SetDate(Format(Format(mstrBegin, "YYYY-MM-DD") & " " & strTime & ":00", "YYYY-MM-DD HH:mm:ss"))
                End If
                
                If strValue <> "" Then
                    strValue = strName
                    lngID = GetMaxID
                    gstrFields = "序号|项目序号|时间|原始时间|记录类型|内容|项目名称|未记说明|记录组号|数据来源|显示|来源ID|共用|状态"
                    gstrValues = lngID & "|" & 0 & "|" & strTime & "|" & strTime & "|" & lngNO & "|" & strValue & "|" & strName & "|" & IIf(lngNO = 4, "", vsfCurve.Cell(flexcpBackColor, intRow, col_颜色, intRow, col_颜色)) & "|0|0|0|0|0|1"
                    vsfOper.RowData(intRow) = lngID
                    Call Record_Add(mrsOper, gstrFields, gstrValues)
                End If
            End If
        End If
        
    End If
    
    If intCOl = col_数据 And Trim(vsfCurve.Tag) <> Trim(vsfCurve.TextMatrix(intRow, col_数据)) Then
        mblnChage = True
        mblnCurveChange = True
    ElseIf intCOl = COL_部位 And Trim(vsfCurve.Tag) <> str部位 Then
        mblnChage = True
        mblnCurveChange = True
    ElseIf intType = 1 And intCOl = Col_未记说明 And Trim(vsfCurve.Tag) <> Trim(vsfCurve.TextMatrix(intRow, Col_未记说明)) Then
        mblnChage = True
        mblnCurveChange = True
    ElseIf intType = 2 And intCOl = col_数据 And PicValue.Visible = True And PicValue.Tag <> vsfCurve.Cell(flexcpBackColor, intRow, col_颜色) Then
        mblnChage = True
        mblnCurveChange = True
    ElseIf lngNO = 1 And intCOl = col_复查 Then
        mblnChage = True
        mblnCurveChange = True
    ElseIf lngNO = 4 Then '手术信息
        If strName <> vsfOper.TextMatrix(intRow, Col_OperType) Or Format(strTime, "HH:mm") <> Format(vsfOper.TextMatrix(intRow, Col_OperTime), "HH:mm:ss") Then
            mblnChage = True
        End If
    End If
    
    UpdateCurveDate = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub vsfOper_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    vsfOper.Body.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
    vsfOper.Body.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
End Sub

Private Sub vsfOper_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '检查是否是同步过来的数据
    Dim lngID As Long, intState As Integer
    lngID = Val(vsfOper.RowData(Row))
    If lngID > 0 Then
        mrsOper.Filter = "记录类型=4 And 序号=" & lngID
        intState = mrsOper!状态
        If InStr(1, ",0,9,", "," & Val(Nvl(mrsOper!数据来源, 0)) & ",") = 0 Then
            Cancel = True
            lblStb.Caption = "同步过来的数据,不允许进行数据删除."
            lblStb.ForeColor = 255
            vsfOper.Body.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
        End If
        
        '完成数据的删除操作
        If intState = 0 Or intState = 2 Then '表示是原有数据
            mrsOper!内容 = ""
            mrsOper!项目名称 = ""
            mrsOper!状态 = 2
        Else '表示新增数据
            mrsOper.Delete
        End If
        mrsOper.Update
        mblnChage = True
    End If
End Sub

Private Sub vsfOper_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Dim intRow As Integer
    '如果上一列没有录入时间和手术信息 不能进行下一行
    If Row >= vsfOper.FixedRows And Col >= vsfOper.FixedCols Then
        If vsfOper.TextMatrix(Row, Col_OperTime) = "" Or (vsfOper.TextMatrix(Row, Col_OperType) = "" And vsfOper.EditText = "") Then Cancel = True
    End If
End Sub

Private Sub vsfOper_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    With vsfOper
        If .EditMode(NewCol) = 1 Then
            .Body.FocusRect = flexFocusSolid
        Else
            .Body.FocusRect = flexFocusLight
        End If
    End With
End Sub

Private Sub vsfOper_ComboCloseUp(Row As Long, Col As Long, FinishEdit As Boolean)
    If Trim(vsfOper.TextMatrix(Row, Col_OperTime)) <> "" Then
        Call UpdateCurveDate(Row, Col, 2, True, True)
    End If
End Sub

Private Sub vsfOper_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        If FilterKeyAscii(KeyAscii, 99, "0123456789:") = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub vsfOper_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mblnFileBack = True Then
        Cancel = True
        vsfOper.Body.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
        lblStb.Caption = "病人体温数据已经归档,不允许进行数据修改."
        lblStb.ForeColor = 255
        vsfOper.Body.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
    End If
    
    '检查是否是同步过来的数据
    If Val(vsfOper.RowData(Row)) > 0 Then
        mrsOper.Filter = "记录类型=4 And 序号=" & Val(vsfOper.RowData(Row))
        If InStr(1, ",0,9,", "," & Val(Nvl(mrsOper!数据来源, 0)) & ",") = 0 Then
            Cancel = True
            lblStb.Caption = "同步过来的数据,不允许进行数据修改."
            lblStb.ForeColor = 255
            vsfOper.Body.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
        End If
    End If
End Sub

Private Sub vsfOper_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '进行数据合法性检查
    Dim strText As String
    Dim strInfo As String, strDate As String
    Dim rsTemp As New ADODB.Recordset
    
    If Row < vsfOper.FixedRows Then Exit Sub
    With vsfOper
        strText = .EditText
        If Col = Col_OperTime Then
            If Trim(strText) = "" Then
                .TextMatrix(Row, Col_OperType) = ""
                GoTo ErrEnd
            End If
            Select Case Len(strText)
            Case 3, 4
                strText = String(4 - Len(strText), "0") & strText
                strText = Mid(strText, 1, 2) & ":" & Mid(strText, 3)
            Case Is < 3
                strText = String(2 - Len(strText), "0") & strText
                strText = Format(Now, "HH") & ":" & strText
            End Select
            
            '合法性检查
            If Mid(strText, 3, 1) <> ":" Then
                strInfo = "录入的时点格式非法！[小时:分钟]"
                GoTo ErrInfo
            End If
            If Mid(strText, 1, 2) < 0 Or Mid(strText, 1, 2) > 23 Then
                strInfo = "录入的时点格式非法！[小时应在0至23之间]"
                GoTo ErrInfo
            End If
            If Mid(strText, 4, 2) < 0 Or Mid(strText, 4, 2) > 59 Then
                strInfo = "录入的时点格式非法！[分钟应在0至59之间]"
                GoTo ErrInfo
            End If
            .EditText = Format(strText, "HH:mm")
            
            '检查录入的时间是否已经存在了手术信息
            strDate = Format(dkpDate.Value & " " & strText, "YYYY-MM-DD HH:mm:ss")
            gstrSQL = "select 1 from 病人护理文件 A,病人护理数据 B,病人护理明细 C" & _
                " Where A.ID=B.文件ID And B.ID=C.记录ID And A.ID=[1] And B.发生时间=[2] And C.记录类型=4"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "是否存在手术", mT_Patient.lng文件ID, CDate(strDate))
            If rsTemp.RecordCount > 0 Then
                strInfo = "改时间已经存在手术信息，请检查！ 时间[" & strDate & "]"
                GoTo ErrInfo
            End If
            
            If Not CheckDateTime(Row, "时间", Format(dkpDate.Value & " " & strText, "YYYY-MM-DD HH:mm:ss")) Then Cancel = True
        End If
    End With
ErrEnd:
    '验证通过进行数据保存操作
    Call UpdateCurveDate(Row, Col, 2, IIf(Col = Col_OperType, True, False), True)
    Exit Sub
ErrInfo:
    lblStb.Caption = strInfo
    lblStb.ForeColor = 255
    Cancel = True
End Sub

Private Sub vsfTab_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strInfo As String
    Dim lngNO As Long, strName As String, strTmp As String, str值域 As String
    Dim arrStr() As String
    Dim cbrControl As CommandBarControl
    Dim blnCheck As Boolean
    Dim strText As String
    
    Call AdjustRowFlag(vsfTab, NewRow)
    
    If mblnInit = False Then Exit Sub
    If NewRow < vsfTab.FixedRows Or NewCol < vsfTab.FixedCols Then Exit Sub
    
    With vsfTab
        lngNO = Val(.TextMatrix(NewRow, COL_tab项目序号))
        strTmp = .TextMatrix(NewRow, COL_tab项目名称)
        If strTmp = "" Then strTmp = "("
        strName = Split(strTmp, "(")(0)
        strTmp = .TextMatrix(NewRow, COL_tab字符串)
        strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
        arrStr = Split(strTmp, ",")
        str值域 = arrStr(0)
        
        If str值域 = "" Then
            strInfo = ""
        Else
            strInfo = strName & "有效范围:" & str值域
        End If
        
        If lngNO = 4 And strName = "血压" Then '血压
            strInfo = strInfo & Space(4) & "录入规则:收缩压/舒张压或(外出、未测、拒测、请假)"
        End If
        
        If Val(arrStr(4)) = 4 Then strInfo = strInfo & Space(4) & "汇总项目" & Space(4) & "录入规则:今天录入" & IIf(mbln汇总当天 = True, "今天", "昨天") & "的数据。"
    End With
    
    lblStb.Caption = strInfo
    lblStb.ForeColor = &H80000012
    
    '检查数据是否允许修改
    mrsCurve.Filter = "项目序号=" & lngNO & " and 项目名称='" & strName & "'" & _
        "   and 列号=" & NewCol - vsfTab.FixedCols + 1
    If mrsCurve.RecordCount > 0 Then
        If InStr(1, ",0,9,", "," & Val(mrsCurve!数据来源) & ",") = 0 Then
            lblStb.Caption = "对于来源于护理记录单或PDA的数据不能进行修改、删除操作"
            lblStb.ForeColor = 255
            Exit Sub
        End If
    End If
    
    If NewCol < vsfTab.FixedCols + Val(arrStr(3)) Then
        '确定大便或入液类型
        strText = Trim(vsfTab.TextMatrix(NewRow, NewCol))
        blnCheck = False
        For Each cbrControl In mcbrToolBar.Controls(5).CommandBar.Controls
            cbrControl.Checked = False
            If lngNO = gint大便 Then
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
               
            ElseIf lngNO = gint入液 Then
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
End Sub

Private Sub vsfTab_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mblnScroll = True
    Call vsfTab_EnterCell
    mblnScroll = False
End Sub

Private Sub vsfTab_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    
    With vsfTab
        If NewRow >= .FixedRows And NewCol >= .FixedRows Then
            If NewCol < .FixedCols + (Split(.TextMatrix(NewRow, COL_tab字符串), ",")(3)) Then
                mrsCurve.Filter = "项目序号=" & Val(.TextMatrix(NewRow, COL_tab项目序号)) & " and 项目名称='" & Split(.TextMatrix(NewRow, COL_tab项目名称), "(")(0) & "'" & _
                    "   and 列号=" & NewCol - .FixedCols + 1
                If mrsCurve.RecordCount > 0 Then
                    If InStr(1, ",0,9,", "," & Val(mrsCurve!数据来源) & ",") = 0 Then
                        .FocusRect = flexFocusHeavy
                    Else
                        .FocusRect = flexFocusSolid
                    End If
                Else
                    .FocusRect = flexFocusSolid
                End If
            Else
                .FocusRect = flexFocusHeavy
            End If
        Else
            .FocusRect = flexFocusNone
        End If
    End With
    
End Sub

Private Sub vsfTab_DblClick()
    With vsfTab
        If .Row > .FixedRows - 1 And .Col > .FixedCols - 1 And .FocusRect = flexFocusSolid Then
            mblnEdit = True
            Call vsfTab_EnterCell
        End If
    End With
End Sub

Private Sub vsfTab_EnterCell()
    Dim intRow As Integer, intCOl As Integer
    Dim strData As String
    Dim blnAllow As Boolean
    Dim blnEdit As Boolean
    Dim strInfo As String, strValue As String, strValue1 As String
    Dim blnSelect As Boolean
    Dim arrValue() As String, arrValue1() As String
    Dim intType As Integer, int频次 As Integer
    Dim i As Integer, j As Integer
    Dim strTime As String, strTmp As String
    Dim arrStr() As String
    Dim intNum As Integer, intLen As Integer
    Dim lngItemNO As Long
    Dim lngColor As Long
    
    If Not mblnInit Then Exit Sub
    blnAllow = True
    blnEdit = True
    blnSelect = False
    '检查数据合法性
    '--51282,刘鹏飞,2012-08-03,全天汇总显示录入时间(DYEY要求手工录入汇总时间H)
    If picEdit.Visible = True And txtEdit.Tag <> "" Then
        intRow = Split(txtEdit.Tag, "|")(0)
        intCOl = Split(txtEdit.Tag, "|")(1)
        
        If txtEdit.Visible = True Then
            strData = IIf(picHour.Visible = True, "(" & txtHour.Text & "h)", "") & Trim(txtEdit.Text)
            lngColor = txtEdit.ForeColor
        Else
            strData = Trim(lblCheck.Caption)
            lngColor = 0
        End If
        
        If strData & "/#$&/" & lngColor <> picEdit.Tag Then blnAllow = WriteIntoVfgTab(strData, False, True, strInfo)
        If cmdColor.Visible = True Then vsfTab.Cell(flexcpForeColor, intRow, intCOl, intRow, intCOl) = Val(cmdColor.Tag)
    End If
     
    '数据不合法
    If blnAllow = False Then
        If vsfTab.Row <> intRow Then vsfTab.Row = intRow
        If vsfTab.Col <> intCOl Then vsfTab.Col = intCOl
        GoTo ErrFouce
        Exit Sub
    End If
    
    If vsfTab.Row < vsfTab.FixedRows And vsfTab.Col < vsfTab.FixedCols Then Exit Sub
    If Not vsfTab.RowIsVisible(vsfTab.Row) Then Exit Sub
    If Not mblnScroll And vsfTab.Visible Then vsfTab.SetFocus
    
    '隐藏所有编辑控件
    pic未记.Visible = False
    picEdit.Visible = False
    picEdit.Tag = ""
    txtEdit.Tag = "": txtEdit.Visible = False: txtEdit.Enabled = False
    picHour.Visible = False: picHour.Enabled = False
    txtHour.Tag = "": txtHour.Visible = False: txtHour.Enabled = False
    lblCheck.Visible = False: lblCheck.Enabled = False
    cmdColor.Visible = False
    cmdColor.Enabled = False
    cmdColor.Tag = 0
    picColor.Visible = False
    PicLst.Visible = False
    PicLst.Tag = ""
    txtLst.Visible = False: txtLst.Text = ""
    lstSelect(0).Visible = False
    lstSelect(0).Enabled = False
    lstSelect(0).Tag = ""
    lstSelect(1).Visible = False
    lstSelect(1).Enabled = False
    lstSelect(1).Tag = ""
    
    If mblnFileBack = True Then
        strInfo = "病人体温数据已经归档,不允许进行数据修改."
        mblnEdit = False
        GoTo ErrInfo
    End If
        
    If mblnEdit = False Then Exit Sub
    
    With vsfTab
        If .Row > .FixedRows - 1 And .Col > .FixedCols - 1 And vsfTab.Col < .FixedCols + Val(Split(vsfTab.TextMatrix(vsfTab.Row, COL_tab字符串), ",")(3)) Then
            
            intType = Val(Split(.TextMatrix(vsfTab.Row, COL_tab字符串), ",")(4))
            int频次 = Val(Split(.TextMatrix(vsfTab.Row, COL_tab字符串), ",")(3))
            '检查录入的项目时间是否超出用户设置的时间范围或补录时间范围
            Call GetAnimalItemTime(.Row, .Col, strInfo)
            If strInfo <> "" Then
                mblnEdit = False
                GoTo ErrInfo
            End If
            '检查波动项目
            If IsWaveItem(Val(.TextMatrix(.Row, COL_tab项目序号))) And InStr(1, Trim(.TextMatrix(.Row, .Col)), "-") <> 0 Then
                strInfo = "对于数值已经形成波动范围的波动项目不能进行修改、删除操作"
                GoTo ErrInfo
            End If
            
            '检查数据来源是否来自护理记录单或PDA
            mrsCurve.Filter = "项目序号=" & Val(.TextMatrix(.Row, COL_tab项目序号)) & " and 项目名称='" & Split(.TextMatrix(.Row, COL_tab项目名称), "(")(0) & "'" & _
                "   and 列号=" & .Col - .FixedCols + 1
            If mrsCurve.RecordCount > 0 Then
                If InStr(1, ",0,9,", "," & Val(mrsCurve!数据来源) & ",") = 0 Then
                    blnEdit = False
                End If
                cmdColor.Tag = Val(mrsCurve!未记说明)
            End If
            '--51282,刘鹏飞,2012-08-03,全天汇总显示录入时间(DYEY要求手工录入汇总时间H)
            If blnEdit = False And Not (intType = 4 And int频次 = 1 And mbln录入小时 = True) Then
                strInfo = "对于来源于护理记录单或PDA的数据不能进行修改、删除操作"
                GoTo ErrInfo
            End If
                  
            If Not (intType = 2 Or intType = 3) Then
                picEdit.Width = .CellWidth + 10
                picEdit.Height = .CellHeight - 5
                picEdit.Top = .CellTop + .Top + 20
                picEdit.Left = .CellLeft + .Left + 15
                picEdit.Enabled = True
                picEdit.Visible = True
                picEdit.ZOrder 0
                
                txtEdit.Top = 0
                txtEdit.Left = 0
                txtEdit.Height = picEdit.Height
            End If
            '对于项目类型是文字类型的活动项目允许设置其字体颜色
             If Val(Split(.TextMatrix(vsfTab.Row, COL_tab字符串), ",")(1)) = 1 And intType = 0 And Val(Split(.TextMatrix(vsfTab.Row, COL_tab字符串), ",")(5)) = 2 Then
                cmdColor.Top = 0
                cmdColor.Height = picEdit.Height
                cmdColor.Width = 300
                cmdColor.Left = picEdit.Width - cmdColor.Width
                txtEdit.Width = cmdColor.Left
                cmdColor.Enabled = True
                cmdColor.Visible = True
                GoTo ShowText
            ElseIf intType = 4 And int频次 = 1 And mbln录入小时 = True Then
                txtHour.Top = 10
                txtHour.Left = 10
                txtHour.Width = picHour.TextWidth("111")
                txtHour.Height = txtEdit.Height
                txtHour.MaxLength = 2
                txtHour.Visible = True
                txtHour.Enabled = True
                
                lblHour.Left = txtHour.Left + txtHour.Width
                lblHour.Top = 10 ' txtHour.Top + (txtHour.Height - lblHour.Height) \ 2
                lblHour.Visible = True
                lblHour.Enabled = True
                
                picHour.Top = -10
                picHour.Left = -10
                picHour.Width = lblHour.Left + lblHour.Width + picHour.TextWidth("1") \ 2
                picHour.Height = picEdit.Height + 20
                picHour.Visible = True
                picHour.Enabled = True
                picHour.ZOrder 0
                
                txtEdit.Top = 10
                txtEdit.Left = picHour.Left + picHour.Width + 10
                txtEdit.Width = picEdit.Width - picHour.Width + 10
                
                strTmp = vsfTab.TextMatrix(vsfTab.Row, COL_tab字符串)
                lngItemNO = Val(vsfTab.TextMatrix(vsfTab.Row, COL_tab项目序号))
                strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
                arrStr = Split(strTmp, ",")
                intNum = Val(arrStr(2))
                intLen = Val(arrStr(6))
                
                If intLen <> 0 Then
                    If lngItemNO <> 4 Then
                        txtEdit.MaxLength = intLen + IIf(intNum = 0, 0, 1)
                    Else
                        txtEdit.MaxLength = intLen * 2 + 1 + IIf(intNum = 0, 0, 1) * 2
                    End If
                Else
                    txtEdit.MaxLength = 0
                End If
                
                If InStr(1, .TextMatrix(vsfTab.Row, vsfTab.Col), ")") > 0 Then
                    txtHour.Text = Replace(Replace(Split(.TextMatrix(vsfTab.Row, vsfTab.Col), ")")(0), "(", ""), "h", "")
                    txtEdit.Text = Split(.TextMatrix(vsfTab.Row, vsfTab.Col), ")")(1)
                Else
                    txtEdit.Text = .TextMatrix(vsfTab.Row, vsfTab.Col)
                End If
                picEdit.Tag = .TextMatrix(vsfTab.Row, vsfTab.Col) & "/#$&/" & .Cell(flexcpForeColor, vsfTab.Row, vsfTab.Col)
                txtEdit.Tag = vsfTab.Row & "|" & vsfTab.Col
                txtEdit.ForeColor = cmdColor.Tag
                txtEdit.Visible = True
                txtEdit.Enabled = blnEdit
                txtEdit.ZOrder 0
                picHour.SetFocus
            ElseIf intType = 2 Or intType = 3 Then '单选
                
                '51600,刘鹏飞,2012-07-16,单选项目提供可以选择和录入功能
                strValue = Split(.TextMatrix(vsfTab.Row, COL_tab字符串), ",")(0)
                Select Case intType
                    Case 2
                        If Left(strValue, 1) <> ":" Then strValue = ":" & strValue
                        intType = 0
                    Case 3
                        intType = 1
                End Select
                
                arrValue = Split(strValue, ":")
                lstSelect(intType).Clear
                PicLst.Tag = "1"
                For i = 0 To UBound(arrValue)
                    If Left(arrValue(i), 1) = "√" Then arrValue(i) = Mid(arrValue(i), 2): strValue1 = arrValue(i)
                    lstSelect(intType).AddItem arrValue(i), i
                     
                     If intType = 0 Then
                        ReDim arrValue1(0)
                        arrValue1(0) = .TextMatrix(.Row, .Col)
                        txtLst.Text = .TextMatrix(.Row, .Col)
                     Else
                        arrValue1 = Split(.TextMatrix(.Row, .Col), ",")
                     End If
                     For j = 0 To UBound(arrValue1)
                        If arrValue1(j) = arrValue(i) Then
                            lstSelect(intType).Selected(i) = True
                            blnSelect = True
                        End If
                    Next j
                Next i
                
                If blnSelect = False And strValue1 <> "" And IIf(intType = 0, Trim(txtLst.Text) = "", True) Then
                    For i = 0 To lstSelect(intType).ListCount - 1
                        If lstSelect(intType).List(i) = strValue1 Then
                            lstSelect(intType).Selected(i) = True
                        End If
                    Next i
                End If
                
                If lstSelect(intType).ListIndex >= 0 Then txtLst.Text = "": PicLst.Tag = 0
                
                '控件显示
                '51600,刘鹏飞,2012-07-16,单选项目提供可以选择和录入功能
                If intType = 0 Then
                    PicLst.FontName = .FontName
                    PicLst.FontSize = .FontSize
                    PicLst.Left = .CellLeft + .Left + 15
                    PicLst.Top = .CellTop + vsfTab.Top
                    PicLst.Height = 80 + (.CellHeight - 5) + PicLst.TextHeight("刘") * 2 + lstSelect(intType).ListCount * (PicLst.TextHeight("刘") + PicLst.TextHeight("刘") \ 4)
                    If PicLst.Height < .CellHeight + 20 Then PicLst.Height = .CellHeight + 20
                    PicLst.Width = LenB(StrConv(lstSelect(intType).List(lstSelect(intType).ListCount \ 2), vbFromUnicode)) * 120 + 500    '以中间项的长度为依据
                    If PicLst.Width < .CellWidth + 20 Then PicLst.Width = .CellWidth + 20
                    If PicLst.Height > vsfTab.Height Then PicLst.Height = vsfTab.Height
                    If PicLst.Top + PicLst.Height > vsfTab.Height Then PicLst.Top = .CellTop + .Top + .CellHeight + 20 - PicLst.Height
                    If PicLst.Top < 0 Then PicLst.Top = vsfTab.Top
                    PicLst.Visible = True
                    PicLst.ZOrder 0
                    
                    lbllst(2).Left = 20
                    lbllst(2).Top = 20
                    If lbllst(2).Width > PicLst.Width Then
                        PicLst.Width = lbllst(2).Width + PicLst.TextWidth("刘")
                    End If
                    lbllst(2).FontName = .FontName
                    lbllst(2).FontSize = .FontSize
                    lbllst(2).Visible = True
            
                    txtLst.Top = lbllst(2).Top + lbllst(2).Height + 20
                    txtLst.Left = -10
                    txtLst.Width = PicLst.Width
                    txtLst.Height = .CellHeight - 5
                    txtLst.FontName = .FontName
                    txtLst.FontSize = .FontSize
                    strTmp = vsfTab.TextMatrix(vsfTab.Row, COL_tab字符串)
                    strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
                    arrStr = Split(strTmp, ",")
                    intNum = Val(arrStr(2))
                    intLen = Val(arrStr(6))
                    txtLst.MaxLength = intLen + IIf(intNum = 0, 0, 1)
                    txtLst.Visible = True
                    
                    lbllst(3).Left = 20
                    lbllst(3).Top = txtLst.Top + txtLst.Height + 20
                    lbllst(3).FontName = .FontName
                    lbllst(3).FontSize = .FontSize
                    lbllst(3).Visible = True
                    
                    lstSelect(intType).Top = lbllst(3).Top + lbllst(3).Height + 20
                    lstSelect(intType).Left = -10
                    lstSelect(intType).FontName = .FontName
                    lstSelect(intType).FontSize = .FontSize
                    lstSelect(intType).Width = PicLst.Width
                    lstSelect(intType).Height = PicLst.Height - lstSelect(intType).Top
                    lstSelect(intType).Visible = True
                    lstSelect(intType).Enabled = True
                    lstSelect(intType).ZOrder 0
                    lstSelect(intType).Tag = .TextMatrix(.Row, .Col)
                    lbllst(intType).Tag = .Row & "|" & .Col
                    
                    If lstSelect(intType).Top + lstSelect(intType).Height <> PicLst.Height Then
                        PicLst.Height = lstSelect(intType).Top + lstSelect(intType).Height
                    End If
                    PicLst.SetFocus
                Else
                    lstSelect(intType).Top = .CellTop + vsfTab.Top
                    lstSelect(intType).Left = .CellLeft + .Left + 15
                    lstSelect(intType).FontName = .FontName
                    lstSelect(intType).FontSize = .FontSize
                    lstSelect(intType).Height = lstSelect(intType).ListCount * (PicLst.TextHeight("刘") + PicLst.TextHeight("刘") \ 4)
                    If lstSelect(intType).Height < .CellHeight + 20 Then lstSelect(intType).Height = .CellHeight + 20
                    lstSelect(intType).Width = LenB(StrConv(lstSelect(intType).List(lstSelect(intType).ListCount \ 2), vbFromUnicode)) * 120 + 500    '以中间项的长度为依据
                    If lstSelect(intType).Width < .CellWidth + 20 Then lstSelect(intType).Width = .CellWidth + 20
                    If lstSelect(intType).Height > vsfTab.Height Then
                        lstSelect(intType).Height = vsfTab.Height
                    End If
                    If lstSelect(intType).Top + lstSelect(intType).Height > vsfTab.Height Then
                        lstSelect(intType).Top = .CellTop + .Top + .CellHeight + 20 - lstSelect(intType).Height
                    End If
                    If lstSelect(intType).Top < 0 Then lstSelect(intType).Top = vsfTab.Top
                    
                    lstSelect(intType).Visible = True
                    lstSelect(intType).Enabled = True
                    lstSelect(intType).ZOrder 0
                    
                    lstSelect(intType).Tag = .TextMatrix(.Row, .Col)
                    lbllst(intType).Tag = .Row & "|" & .Col
                    lstSelect(intType).SetFocus
                End If
            ElseIf intType = 5 Then '选择
                lblCheck.Width = picEdit.Width
                lblCheck.Height = picEdit.Height
                lblCheck.Caption = .TextMatrix(vsfTab.Row, vsfTab.Col)
                picEdit.Tag = .TextMatrix(vsfTab.Row, vsfTab.Col) & "/#$&/" & .Cell(flexcpForeColor, vsfTab.Row, vsfTab.Col)
                txtEdit.Tag = vsfTab.Row & "|" & vsfTab.Col
                lblCheck.Visible = True
                lblCheck.Enabled = True
                lblCheck.ZOrder 0
                picEdit.SetFocus
            Else
                txtEdit.Width = picEdit.Width
ShowText:
                strTmp = vsfTab.TextMatrix(vsfTab.Row, COL_tab字符串)
                lngItemNO = Val(vsfTab.TextMatrix(vsfTab.Row, COL_tab项目序号))
                strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
                arrStr = Split(strTmp, ",")
                intNum = Val(arrStr(2))
                intLen = Val(arrStr(6))
                
                If intLen <> 0 Then
                    If lngItemNO <> 4 Then
                        txtEdit.MaxLength = intLen + IIf(intNum = 0, 0, 1)
                    Else
                        txtEdit.MaxLength = intLen * 2 + 1 + IIf(intNum = 0, 0, 1) * 2
                    End If
                Else
                    txtEdit.MaxLength = 0
                End If
                
                txtEdit.Text = .TextMatrix(vsfTab.Row, vsfTab.Col)
                picEdit.Tag = .TextMatrix(vsfTab.Row, vsfTab.Col) & "/#$&/" & .Cell(flexcpForeColor, vsfTab.Row, vsfTab.Col)
                txtEdit.Tag = vsfTab.Row & "|" & vsfTab.Col
                txtEdit.ForeColor = cmdColor.Tag
                txtEdit.Visible = True
                txtEdit.Enabled = True
                txtEdit.ZOrder 0
                picEdit.SetFocus
            End If
            
        End If
    End With
ErrFouce:
    If picEdit.Visible = True And txtEdit.Enabled = True Then txtEdit.SetFocus: Call zlControl.TxtSelAll(txtEdit)
ErrInfo:
    If strInfo <> "" Then
        lblStb.Caption = strInfo
        lblStb.ForeColor = 255
    End If
End Sub

Private Sub vsfTab_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    vsfTab.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignLeftCenter
    If mblnFileBack = True Then
        Cancel = True
        vsfTab.Body.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
        lblStb.Caption = "病人体温数据已经归档,不允许进行数据修改."
        lblStb.ForeColor = 255
    End If
End Sub

Private Sub vsfTab_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    vsfTab.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
End Sub

Private Sub vsfTab_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim intCols As Integer
    Dim intType As Integer, int频次 As Integer
    Dim blnTrue As Boolean
    Dim blnEdit As Boolean
    Dim strText As String
    
    If vsfTab.Row < vsfTab.FixedRows And vsfTab.Col < vsfTab.FixedCols Then Exit Sub
    
    '屏蔽掉某些功能键
    If KeyCode = vbKeyEscape Or KeyCode = vbKeyBack Or Shift <> 0 Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Then Exit Sub
    
    If KeyCode = vbKeyLeft And (picEdit.Visible = False And lstSelect(0).Visible = False And lstSelect(1).Visible = False) Then Exit Sub
    
    intCols = Val(Split(vsfTab.TextMatrix(vsfTab.Row, COL_tab字符串), ",")(3)) + vsfTab.FixedCols
    
    With vsfTab
        If KeyCode = vbKeyReturn Then
NextCol2: '跳到下一行
            If .Col < vsfTab.FixedCols Then
                .Col = .Col + 1: GoTo NextCol2
            End If
            If .Col < intCols - 1 Then
                .Col = .Col + 1
                If .ColHidden(.Col) = True Then GoTo NextCol2
            Else
NextRow2: '跳到下一列
                If .Row < .Rows - 1 Then
                    .Col = vsfTab.FixedCols: .Row = .Row + 1
                    If .RowHidden(.Row) = True Then GoTo NextRow2
                Else
                    Call txtEdit_KeyPress(vbKeyEscape)
                    .Row = .FixedRows
                    .Col = .FixedCols
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
            If .Col > vsfTab.FixedCols Then
                .Col = .Col - 1
                If .ColHidden(.Col) = True Then GoTo PreCol2
            Else
PreRow2:
                If .Row > vsfTab.FixedRows Then
                    .Row = .Row - 1
                    If .RowHidden(.Row) Then GoTo PreRow2
                    .Col = Val(Split(vsfTab.TextMatrix(vsfTab.Row, COL_tab字符串), ",")(3)) + vsfTab.FixedCols
                    GoTo PreCol2
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
        
        '删除信息
        If KeyCode = vbKeyDelete Then
            If Shift = 0 And .Col > .FixedCols - 1 And .Col < intCols Then
                blnEdit = True
                If .TextMatrix(.Row, .Col) <> "" Then
                    '检查项目是否是波动项目
                    If IsWaveItem(Val(.TextMatrix(.Row, COL_tab项目序号))) And InStr(1, Trim(.TextMatrix(.Row, .Col)), "-") <> 0 Then
                        lblStb.Caption = "对于数值已经形成波动范围的波动项目不能进行修改、删除操作"
                        lblStb.ForeColor = 255
                        GoTo ErrExit
                    End If
                    '检查数据来源是否来自护理记录单或PDA
                    mrsCurve.Filter = "项目序号=" & Val(.TextMatrix(.Row, COL_tab项目序号)) & " and 项目名称='" & Split(.TextMatrix(.Row, COL_tab项目名称), "(")(0) & "'" & _
                        "   and 列号=" & .Col - .FixedCols + 1
                    If mrsCurve.RecordCount > 0 Then
                        If InStr(1, ",0,9,", "," & Val(mrsCurve!数据来源) & ",") = 0 Then
                            blnEdit = False
                        End If
                    End If
                    int频次 = Val(Split(vsfTab.TextMatrix(vsfTab.Row, COL_tab字符串), ",")(3))
                    intType = Val(Split(vsfTab.TextMatrix(vsfTab.Row, COL_tab字符串), ",")(4))
                    If blnEdit = False And Not (intType = 4 And int频次 = 1 And mbln录入小时 = True) Then
                        lblStb.Caption = "对于来源于护理记录单或PDA的数据不能进行修改、删除操作"
                        lblStb.ForeColor = 255
                        GoTo ErrExit
                    End If
                    picTab.Tag = .Row & "|" & .Col
                    FraTable.Tag = .TextMatrix(.Row, .Col)
                    strText = ""
                    If blnEdit = False Then '表明是全天汇总项目，并且mbln录入小时=true
                        If InStr(1, .TextMatrix(.Row, .Col), ")") > 0 Then
                            strText = Split(.TextMatrix(.Row, .Col), ")")(1)
                        Else
                            GoTo ErrExit
                        End If
                    End If
                    blnTrue = WriteIntoVfgTab(strText, True)
                End If
            End If
ErrExit:
            mblnEdit = False
            Exit Sub
        End If
        mblnEdit = True
        Call vsfTab_EnterCell
    End With
End Sub

Private Function SaveData() As Boolean
'--------------------------------------------------------
'功能:进行数据修改保存
'--------------------------------------------------------
    Dim strSQL As String, arrSQL() As String, arrSQLTime() As String
    Dim strTime As String, strEnd As String, strMarkTime As String, strOldTime As String
    Dim lngItemCode As Long, strValue As String, str未记 As String, strTmp As String
    Dim arrTmp() As String
    Dim intModify As Integer
    Dim blnEdit As Boolean
    Dim blnSave As Boolean
    Dim strName As String, strInfo As String
    Dim lngRow As Long, lng记录ID As Long, lngOldID As Long
    Dim i As Integer, int项目首次 As Integer
    Dim blnTran As Boolean
    Dim int检查科室 As Integer
    
    On Error GoTo Errhand
    
    mrsCurve.Filter = 0
    mrsCurve.Sort = "时间,项目序号"
    Screen.MousePointer = 11
    
    ReDim Preserve arrSQL(1 To 1)
    ReDim Preserve arrSQLTime(1 To 1)
    
    mrsRecodeID.Filter = 0
    '体温数据保存
    With mrsCurve
        Do While Not .EOF
            lngItemCode = Val(!项目序号)
            strValue = Nvl(!数值)
            '问题号:53505,修改人：李涛,血压可以录入文字:外出，未测等.
            If lngItemCode = 4 And zlCommFun.Nvl(!项目名称) = "血压" And Nvl(!数值) <> "" Then
                strValue = Nvl(!数值) & "/" & Nvl(!数值)
            End If
            intModify = Val(zlCommFun.Nvl(!修改))
            blnEdit = False
            If intModify = 1 And InStr(1, ",0,9,", Val(zlCommFun.Nvl(!数据来源))) = 0 Then
                blnEdit = False
            Else
                blnEdit = True
            End If
            blnSave = False
            If Val(!状态) <> 3 And Val(!状态) <> 0 Then
               '体温曲线项目处理
                If !分组名 = "1)体温曲线项目" Then
                    strTime = !时间
                    strOldTime = Trim(zlCommFun.Nvl(!原始时间))
                    If strTime = "" Then
                        '时间为空就提取本段时间的中点时间
                        strTime = mstrBegin
                        strEnd = mstrEnd
                        strMarkTime = Format(GetCenterTime(CDate(mstrBegin), CDate(mstrEnd)), "YYYY-MM-DD HH:mm:ss")
                    Else
                        strEnd = strTime
                        strMarkTime = Format(strTime, "YYYY-MM-DD HH:mm:ss")
                    End If
                    strTime = Format(strTime, "YYYY-MM-DD HH:mm:ss")
                    strEnd = Format(strEnd, "YYYY-MM-DD HH:mm:ss")
                    strOldTime = Format(strOldTime, "YYYY-MM-DD HH:mm:ss")
                    int检查科室 = IIf(ISCheckDept(strMarkTime) = True, 1, 0)
                    strMarkTime = "To_Date('" & strMarkTime & "','yyyy-mm-dd hh24:mi:ss')"
                    
                    '调用修改病人护理数据发生时间
                    If strOldTime <> strTime And strOldTime <> "" Then
                        mrsRecodeID.Filter = "时间='" & strOldTime & "'"
                        If mrsRecodeID.RecordCount > 0 Then
                            lng记录ID = Val(mrsRecodeID!记录ID)
                            
                            '相同记录修改过后不再次进行修改
                            If lng记录ID <> lngOldID Then
                                strSQL = "ZL_体温单数据_发生时间("
                                'ID_IN       IN 病人护理数据.ID%TYPE,
                                strSQL = strSQL & lng记录ID & ","
                                '发生时间_IN IN 病人护理数据.发生时间%TYPE
                                strSQL = strSQL & strMarkTime & ")"
                                
                                arrSQLTime(ReDimArray(arrSQLTime)) = strSQL
                            End If
                        End If
                    End If
                    
                    lngOldID = lng记录ID
                    
                    If strValue = "不升" And lngItemCode = Item体温 Then
                        str未记 = ""
                    Else
                        str未记 = !未记说明
                    End If
                    
                    '状态=4只是对时间进行了修改(上面已经处理)
                    If Val(!状态) <> 4 Then
                        '更新数据信息
                        strSQL = "Zl_体温单数据_Update("
                        '文件id_In   In 病人护理文件.Id%Type,  --病人护理文件ID
                        strSQL = strSQL & Val(mT_Patient.lng文件ID) & ","
                        '发生时间_In In 病人护理数据.发生时间%Type, --护理数据的发生时间
                        strSQL = strSQL & strMarkTime & ","
                        '记录类型_In In 病人护理明细.记录类型%Type, --护理项目=1，上标说明=2，入出转标记=3，手术日标记=4,下标说明=6
                        strSQL = strSQL & "1,"
                        '项目序号_In In 病人护理明细.项目序号%Type, --护理项目的序号，非护理项目固定为0
                        strSQL = strSQL & lngItemCode & ","
                        '记录内容_In In 病人护理明细.记录内容%Type := Null, --记录内容，如果内容为空，即清除以前的内容  36或36/37
                        strSQL = strSQL & "'" & strValue & "',"
                        '体温部位_In In 病人护理明细.体温部位%Type := Null, --删除数据时不用填写部位 除活动项目外
                        strSQL = strSQL & IIf(strValue <> "", "'" & Nvl(!部位) & "'", "NULL") & ","
                        '复试合格_In In Number := 0,
                        strSQL = strSQL & IIf(lngItemCode = Item体温 And strValue <> "", Val(!复查), "0") & ","
                        '未记说明_In In 病人护理明细.未记说明%Type := Null, --未记说明
                        strSQL = strSQL & "'" & str未记 & "',"
                        '他人记录_In In Number := 1,
                        strSQL = strSQL & "1,"
                        '数据来源_In In 病人护理明细.数据来源%Type := 0,
                        strSQL = strSQL & "0,"
                        '来源id_In   In 病人护理明细.来源id%Type := Null,
                        strSQL = strSQL & IIf(Val(!来源ID) = 0, "NULL", !来源ID) & ","
                        '共用_In     In 病人护理明细.共用%Type := 0,
                        strSQL = strSQL & Val(!共用)
                        '  项目首次_In In Number := 0,--汇总项目使用，保存数据前是否先删除一段时间内的数据信息。 1 删除
                        '  开始时间_In In 病人护理数据.发生时间%Type := Null, --本记录有效跨度的开始时间
                        '  结束时间_In In 病人护理数据.发生时间%Type := Null, --本记录有效跨度的终止时间，单独记录为每分钟，体温表为4小时,时间跨度内的相同项目记录要删除
                        '  操作员_IN  IN 病人护理数据.保存人%TYPE := NULL,
                        '  检查科室_IN IN Number :=1
                        strSQL = strSQL & ",0,NULL,NULL,NULL,"
                        strSQL = strSQL & int检查科室 & ")"
                        
                        arrSQL(ReDimArray(arrSQL)) = strSQL
                    End If
                '体温表格项目处理
                ElseIf !分组名 = "2)体温表格项目" Then
                    int项目首次 = 0
                    strName = zlCommFun.Nvl(!项目名称)
                    strTmp = GetItemInfo(lngItemCode, strName, lngRow)
                    strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
                    arrTmp = Split(strTmp, ",")
                    
                    strTime = Format(!时间, "YYYY-MM-DD HH:mm:ss")
                    strEnd = strTime
                    strMarkTime = strTime

                    '对于可以录入的汇总项目,需要根据汇总时段删除本时段内的所有数据
                    If Val(arrTmp(4)) = 4 Then
                        strTmp = GetAnimalItemTime(lngRow, !列号 + vsfCurve.FixedCols - 1, strInfo, 1)
                        If strInfo <> "" Then Exit Function
                        strTime = Split(strTmp, ";")(0)
                        strEnd = Split(strTmp, ";")(1)
                        If CDate(strMarkTime) < CDate(mstrBTime) Then strMarkTime = Format(mstrBTime, "YYYY-MM-DD HH:mm:ss")
                        If CDate(strMarkTime) > CDate(mstrETime) Then strMarkTime = Format(mstrETime, "YYYY-MM-DD HH:mm:ss")
                        int项目首次 = 1
                    End If
                    int检查科室 = IIf(ISCheckDept(strMarkTime) = True, 1, 0)
                    strMarkTime = "To_Date('" & strMarkTime & "','yyyy-mm-dd hh24:mi:ss')"
                    
                    '更新数据信息
                    strSQL = "Zl_体温单数据_Update("
                    '文件id_In   In 病人护理文件.Id%Type,  --病人护理文件ID
                    strSQL = strSQL & Val(mT_Patient.lng文件ID) & ","
                    '发生时间_In In 病人护理数据.发生时间%Type, --护理数据的发生时间
                    strSQL = strSQL & strMarkTime & ","
                    '记录类型_In In 病人护理明细.记录类型%Type, --护理项目=1，上标说明=2，入出转标记=3，手术日标记=4,下标说明=6
                    strSQL = strSQL & Val(Nvl(!记录类型, 1)) & ","
                    '项目序号_In In 病人护理明细.项目序号%Type, --护理项目的序号，非护理项目固定为0
                    strSQL = strSQL & lngItemCode & ","
                    '记录内容_In In 病人护理明细.记录内容%Type := Null, --记录内容，如果内容为空，即清除以前的内容  36或36/37
                    strSQL = strSQL & "'" & strValue & "',"
                    '体温部位_In In 病人护理明细.体温部位%Type := Null, --删除数据时不用填写部位 除活动项目外
                    strSQL = strSQL & IIf(Val(arrTmp(5)) = 2, "'" & Nvl(!部位) & "'", "NULL") & ","
                    '复试合格_In In Number := 0,
                    strSQL = strSQL & IIf(lngItemCode = Item体温 And strValue <> "", Val(!复查), "0") & ","
                    '未记说明_In In 病人护理明细.未记说明%Type := Null, --未记说明
                    If Val(arrTmp(1)) = 1 And Val(arrTmp(5)) = 2 Then
                        strSQL = strSQL & "'" & IIf(strValue = "", "", Val(!未记说明)) & "',"
                    Else
                        strSQL = strSQL & "NUll,"
                    End If
                    '他人记录_In In Number := 1,
                    strSQL = strSQL & "1,"
                    '数据来源_In In 病人护理明细.数据来源%Type := 0,
                    strSQL = strSQL & Val(!数据来源) & ","
                    '来源id_In   In 病人护理明细.来源id%Type := Null,
                    strSQL = strSQL & IIf(Val(!来源ID) = 0, "NULL", !来源ID) & ","
                    '共用_In     In 病人护理明细.共用%Type := 0,
                    strSQL = strSQL & Val(!共用) & ","
                    '项目首次_In In Number := 0,--汇总项目使用，保存数据前是否先删除一段时间内的数据信息。 1 删除
                    strSQL = strSQL & int项目首次 & ","
                    '开始时间_In In 病人护理数据.发生时间%Type := Null,
                    strSQL = strSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'),"
                    '结束时间_In In 病人护理数据.发生时间%Type := Null --本记录有效跨度的终止时间，单独记录为每分钟，体温表为4小时,时间跨度内的相同项目记录要删除
                    strSQL = strSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss')"
                    '  操作员_IN  IN 病人护理数据.保存人%TYPE := NULL,
                    '  检查科室_IN IN Number :=1
                    strSQL = strSQL & ",NULL," & int检查科室 & ")"
                    
                    arrSQL(ReDimArray(arrSQL)) = strSQL
                End If
            End If
        .MoveNext
        Loop
    End With
    
    
    '保存手术及上下标说明信息
    mrsOper.Filter = 0
    '先删除掉修改的手术信息,一天可以设置多次手术，如果手术时间和体温数据时间相同，更新手术时间的话，也会刀子和体温数据时间发生变化
    mrsOper.Sort = "时间"
    With mrsOper
        Do While Not .EOF
            If Val(!状态) <> 3 And Val(!状态) <> 0 Then
                lngItemCode = 4
                If Val(!状态) = 2 Then
                    strTime = Format(!原始时间, "YYYY-MM-DD HH:mm:ss")
                    strEnd = strTime
                    strMarkTime = strTime
                    int检查科室 = IIf(ISCheckDept(strMarkTime) = True, 1, 0)
                    strMarkTime = "To_Date('" & strMarkTime & "','yyyy-mm-dd hh24:mi:ss')"
                    
                    '更新数据信息
                    strSQL = "Zl_体温单数据_Update("
                    '文件id_In   In 病人护理文件.Id%Type,  --病人护理文件ID
                    strSQL = strSQL & Val(mT_Patient.lng文件ID) & ","
                    '发生时间_In In 病人护理数据.发生时间%Type, --护理数据的发生时间
                    strSQL = strSQL & strMarkTime & ","
                    '记录类型_In In 病人护理明细.记录类型%Type, --护理项目=1，上标说明=2，入出转标记=3，手术日标记=4,下标说明=6
                    strSQL = strSQL & lngItemCode & ","
                    '项目序号_In In 病人护理明细.项目序号%Type, --护理项目的序号，非护理项目固定为0
                    strSQL = strSQL & 0 & ","
                    '记录内容_In In 病人护理明细.记录内容%Type := Null, --记录内容，如果内容为空，即清除以前的内容  36或36/37
                    strSQL = strSQL & "NULL" & ","
                    '体温部位_In In 病人护理明细.体温部位%Type := Null, --删除数据时不用填写部位 除活动项目外
                    strSQL = strSQL & "NULL,"
                    '复试合格_In In Number := 0,
                    strSQL = strSQL & "NULL,"
                    '未记说明_In In 病人护理明细.未记说明%Type := Null, --未记说明
                    strSQL = strSQL & "NULL" & ","
                    '他人记录_In In Number := 1,
                    strSQL = strSQL & "1,"
                    '数据来源_In In 病人护理明细.数据来源%Type := 0,
                    strSQL = strSQL & Val(!数据来源) & ","
                    '来源id_In   In 病人护理明细.来源id%Type := Null,
                    strSQL = strSQL & IIf(Val(!来源ID) = 0, "NULL", !来源ID) & ","
                    '共用_In     In 病人护理明细.共用%Type := 0,
                    strSQL = strSQL & Val(!共用) & ","
                    '项目首次_In In Number := 0,--汇总项目使用，保存数据前是否先删除一段时间内的数据信息。 1 删除
                    strSQL = strSQL & 0 & ","
                    '开始时间_In In 病人护理数据.发生时间%Type := Null,
                    strSQL = strSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'),"
                    '结束时间_In In 病人护理数据.发生时间%Type := Null --本记录有效跨度的终止时间，单独记录为每分钟，体温表为4小时,时间跨度内的相同项目记录要删除
                    strSQL = strSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss')"
                    '  操作员_IN  IN 病人护理数据.保存人%TYPE := NULL,
                    '  检查科室_IN IN Number :=1
                    strSQL = strSQL & ",NULL," & int检查科室 & ")"
                    arrSQL(ReDimArray(arrSQL)) = strSQL
                End If
                
                strTime = Format(!时间, "YYYY-MM-DD HH:mm:ss")
                strEnd = strTime
                strMarkTime = strTime
                int检查科室 = IIf(ISCheckDept(strMarkTime) = True, 1, 0)
                strMarkTime = "To_Date('" & strMarkTime & "','yyyy-mm-dd hh24:mi:ss')"
                strValue = Trim(zlCommFun.Nvl(!内容))
                If strValue <> "" Then
                    '更新数据信息
                    strSQL = "Zl_体温单数据_Update("
                    '文件id_In   In 病人护理文件.Id%Type,  --病人护理文件ID
                    strSQL = strSQL & Val(mT_Patient.lng文件ID) & ","
                    '发生时间_In In 病人护理数据.发生时间%Type, --护理数据的发生时间
                    strSQL = strSQL & strMarkTime & ","
                    '记录类型_In In 病人护理明细.记录类型%Type, --护理项目=1，上标说明=2，入出转标记=3，手术日标记=4,下标说明=6
                    strSQL = strSQL & lngItemCode & ","
                    '项目序号_In In 病人护理明细.项目序号%Type, --护理项目的序号，非护理项目固定为0
                    strSQL = strSQL & 0 & ","
                    '记录内容_In In 病人护理明细.记录内容%Type := Null, --记录内容，如果内容为空，即清除以前的内容  36或36/37
                    strSQL = strSQL & "'" & strValue & "',"
                    '体温部位_In In 病人护理明细.体温部位%Type := Null, --删除数据时不用填写部位 除活动项目外
                    strSQL = strSQL & "NULL,"
                    '复试合格_In In Number := 0,
                    strSQL = strSQL & "NULL,"
                    '未记说明_In In 病人护理明细.未记说明%Type := Null, --未记说明
                    strSQL = strSQL & IIf(lngItemCode <> 4, "'" & Nvl(!未记说明) & "'", "NULL") & ","
                    '他人记录_In In Number := 1,
                    strSQL = strSQL & "1,"
                    '数据来源_In In 病人护理明细.数据来源%Type := 0,
                    strSQL = strSQL & Val(!数据来源) & ","
                    '来源id_In   In 病人护理明细.来源id%Type := Null,
                    strSQL = strSQL & IIf(Val(!来源ID) = 0, "NULL", !来源ID) & ","
                    '共用_In     In 病人护理明细.共用%Type := 0,
                    strSQL = strSQL & Val(!共用) & ","
                    '项目首次_In In Number := 0,--汇总项目使用，保存数据前是否先删除一段时间内的数据信息。 1 删除
                    strSQL = strSQL & int项目首次 & ","
                    '开始时间_In In 病人护理数据.发生时间%Type := Null,
                    strSQL = strSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'),"
                    '结束时间_In In 病人护理数据.发生时间%Type := Null --本记录有效跨度的终止时间，单独记录为每分钟，体温表为4小时,时间跨度内的相同项目记录要删除
                    strSQL = strSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss')"
                    '  操作员_IN  IN 病人护理数据.保存人%TYPE := NULL,
                    '  检查科室_IN IN Number :=1
                    strSQL = strSQL & ",NULL," & int检查科室 & ")"
                    arrSQL(ReDimArray(arrSQL)) = strSQL
                End If
            End If
        .MoveNext
        Loop
    End With
    
    '下标信息
    mrsNote.Filter = 0
    mrsNote.Sort = "时间"
    With mrsNote
        Do While Not .EOF
        lngItemCode = Val(!记录类型)
        
        If Val(!状态) <> 3 And Val(!状态) <> 0 Then
            strTime = Format(mstrBegin, "YYYY-MM-DD HH:mm:ss")
            strEnd = Format(mstrEnd, "YYYY-MM-DD HH:mm:ss")
            strMarkTime = Format(!时间, "YYYY-MM-DD HH:mm:ss")
            strValue = zlCommFun.Nvl(!内容)
            int项目首次 = 1
            int检查科室 = IIf(ISCheckDept(strMarkTime) = True, 1, 0)
            strMarkTime = "To_Date('" & strMarkTime & "','yyyy-mm-dd hh24:mi:ss')"
            
             '更新数据信息
            strSQL = "Zl_体温单数据_Update("
            '文件id_In   In 病人护理文件.Id%Type,  --病人护理文件ID
            strSQL = strSQL & Val(mT_Patient.lng文件ID) & ","
            '发生时间_In In 病人护理数据.发生时间%Type, --护理数据的发生时间
            strSQL = strSQL & strMarkTime & ","
            '记录类型_In In 病人护理明细.记录类型%Type, --护理项目=1，上标说明=2，入出转标记=3，手术日标记=4,下标说明=6
            strSQL = strSQL & lngItemCode & ","
            '项目序号_In In 病人护理明细.项目序号%Type, --护理项目的序号，非护理项目固定为0
            strSQL = strSQL & 0 & ","
            '记录内容_In In 病人护理明细.记录内容%Type := Null, --记录内容，如果内容为空，即清除以前的内容  36或36/37
            strSQL = strSQL & "'" & strValue & "',"
            '体温部位_In In 病人护理明细.体温部位%Type := Null, --删除数据时不用填写部位 除活动项目外
            strSQL = strSQL & "NULL,"
            '复试合格_In In Number := 0,
            strSQL = strSQL & "NULL,"
            '未记说明_In In 病人护理明细.未记说明%Type := Null, --未记说明
            strSQL = strSQL & IIf(lngItemCode <> 4, "'" & Nvl(!未记说明) & "'", "NULL") & ","
            '他人记录_In In Number := 1,
            strSQL = strSQL & "1,"
            '数据来源_In In 病人护理明细.数据来源%Type := 0,
            strSQL = strSQL & Val(!数据来源) & ","
            '来源id_In   In 病人护理明细.来源id%Type := Null,
            strSQL = strSQL & IIf(Val(!来源ID) = 0, "NULL", !来源ID) & ","
            '共用_In     In 病人护理明细.共用%Type := 0,
            strSQL = strSQL & Val(!共用) & ","
            '项目首次_In In Number := 0,--汇总项目使用，保存数据前是否先删除一段时间内的数据信息。 1 删除
            strSQL = strSQL & int项目首次 & ","
            '开始时间_In In 病人护理数据.发生时间%Type := Null,
            strSQL = strSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'),"
            '结束时间_In In 病人护理数据.发生时间%Type := Null --本记录有效跨度的终止时间，单独记录为每分钟，体温表为4小时,时间跨度内的相同项目记录要删除
            strSQL = strSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss')"
            '  操作员_IN  IN 病人护理数据.保存人%TYPE := NULL,
            '  检查科室_IN IN Number :=1
            strSQL = strSQL & ",NULL," & int检查科室 & ")"
            arrSQL(ReDimArray(arrSQL)) = strSQL
        End If
        .MoveNext
        Loop
    End With
    
     '------------------------------------------------------------------------------------------------------------------
    '循环执行SQL保存数据
    'Debug.Print "--保存数据开始:" & Now
     
    gcnOracle.BeginTrans
    blnTran = True
    '先执行时间变化
    For i = 1 To UBound(arrSQLTime)
        If arrSQLTime(i) <> "" Then Call zlDatabase.ExecuteProcedure(CStr(arrSQLTime(i)), "保存体温数据"):  'Debug.Print CStr(arrSQLTime(i))
    Next
    '在执行数据变化
    For i = 1 To UBound(arrSQL)
        If arrSQL(i) <> "" Then Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "保存体温数据"):  'Debug.Print CStr(arrSQL(i))
    Next
    gcnOracle.CommitTrans
    
    'Debug.Print "--保存数据结束:" & Now
     
    blnTran = False
    
    mblnChage = False
    mblnEdit = False
    mblnCurveChange = False
    mblnOK = True
    Call txtEdit_KeyPress(vbKeyEscape)
    
    SaveData = True
    Screen.MousePointer = 0
    
    Exit Function
Errhand:
    If blnTran = True Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
    Screen.MousePointer = 0
    Call SaveErrLog
End Function

Private Function ISCheckDept(ByVal str发生时间 As String) As Boolean
'功能：是否在Zl_体温单数据_Update中进行科室检查
    'mstrOverDate<=mstrETime 并且病人已经出院，肯定是病人出院时间和入院时间在一列（程序处理后的结果）
    If mbln出院 = True And Format(mstrOverDate, "YYYY-MM-DD HH:mm:ss") < Format(mstrETime, "YYYY-MM-DD HH:mm:ss") Then
        If Format(str发生时间, "YYYY-MM-DD HH:mm:ss") > Format(mstrOverDate, "YYYY-MM-DD HH:mm:ss") And Format(str发生时间, "YYYY-MM-DD HH:mm:ss") <= Format(mstrETime, "YYYY-MM-DD HH:mm:ss") Then
            ISCheckDept = False
        Else
            ISCheckDept = True
        End If
    Else
        ISCheckDept = True
    End If
End Function

Private Function GetItemInfo(ByVal lngItemNO As Long, ByVal strName As String, ByRef lngRow As Long) As String
'---------------------------------------------------------------
'功能:获取项目信息
'---------------------------------------------------------------
    Dim intRow As Integer
    Dim strValue As String
    
    For intRow = vsfTab.FixedRows To vsfTab.Rows - 1
        If Val(vsfTab.TextMatrix(intRow, COL_tab项目序号)) = lngItemNO And Split(vsfTab.TextMatrix(intRow, COL_tab项目名称), "(")(0) = strName Then
            Exit For
        End If
    Next intRow
    
    If intRow >= vsfTab.Rows Then
        For intRow = vsfTab.FixedRows To vsfTab.Rows - 1
            If Val(vsfTab.TextMatrix(intRow, COL_tab项目序号)) = lngItemNO Then
                Exit For
            End If
        Next intRow
    End If
    
    If intRow < vsfTab.Rows Then
        strValue = vsfTab.TextMatrix(intRow, COL_tab字符串)
    End If
    lngRow = intRow
    GetItemInfo = strValue
End Function

Private Function WriteIntoVfgTab(ByVal strText As String, Optional blnDelete As Boolean = False, Optional ByVal blnVisible As Boolean = True, Optional strErrMsg As String = "") As Boolean
'-------------------------------------------------------------------------
'功能:用户编辑的数据写入vsfTab
'参数:strtext 编辑的文本信息   blndelete 是否在VsfTab按Delete 键删除信息
'-------------------------------------------------------------------------
    Dim intRow As Integer, intCOl As Integer
    Dim lng项目序号 As Long, str项目名称 As String, strTmp As String, strPart As String
    Dim arrStr() As String
    Dim str值域 As String, intType As Integer, intNum As Integer, lngLen As Long, int频次 As Integer, int性质 As Integer, int表示 As Integer
    Dim lngColor As String
    Dim blnAllow As Boolean, blnTrue As Boolean
    Dim strValue As String, strHour As String, strHourOld As String
    Dim intIndex As Integer, int记录类型 As Integer
    Dim strTime As String
    
    '--数据修改信息
    Dim int状态 As Integer
    On Error GoTo Errhand
    
    If Not blnDelete Then
        If picEdit.Visible = True And txtEdit.Tag <> "" Then
            intRow = Split(txtEdit.Tag, "|")(0)
            intCOl = Split(txtEdit.Tag, "|")(1)
            If txtEdit.Visible = True Or lblCheck.Visible = True Then
                strTmp = vsfTab.TextMatrix(intRow, COL_tab字符串)
                lng项目序号 = Val(vsfTab.TextMatrix(intRow, COL_tab项目序号))
                str项目名称 = Split(vsfTab.TextMatrix(intRow, COL_tab项目名称), "(")(0)
                strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
                arrStr = Split(strTmp, ",")
                str值域 = arrStr(0)
                intType = Val(arrStr(1))
                intNum = Val(arrStr(2))
                int频次 = Val(arrStr(3))
                int表示 = Val(arrStr(4))
                int性质 = Val(arrStr(5))
                lngLen = Val(arrStr(6))
                strPart = arrStr(7)
                
                If intType = 1 Then str值域 = ""
                '全天汇总项目，并且参数“全天汇总显示录入时间”勾选
                If int表示 = 4 And int频次 = 1 And mbln录入小时 = True Then
                    If InStr(1, strText, ")") > 0 Then
                        strHour = Replace(Replace(Split(strText, ")")(0), "(", ""), "h", "")
                        If strHour <> "" Then
                            If Not (Val(strHour) >= 0 And strHour <= 24) Then
                                lblStb.Caption = "汇总小时只能在0到24之间，请重新录入！": lblStb.ForeColor = 255
                                Exit Function
                            End If
                            strHour = "(" & strHour & "h)"
                        End If
                        strText = Split(strText, ")")(1)
                        If Trim(strText) = "" Then strHour = ""
                    End If
                End If
                If txtEdit.Enabled = True Then
                    blnAllow = CheckValidata(intRow, intCOl, lng项目序号, intType, intNum, str值域, int表示, lngLen, strText, strErrMsg)
                End If
            End If
            strValue = Split(IIf(Trim(picEdit.Tag) = "", "/#$&/", Trim(picEdit.Tag)), "/#$&/")(0)
        ElseIf lstSelect(0).Visible = True Or lstSelect(1).Visible = True Then
            If lstSelect(0).Visible = True Then strValue = lstSelect(0).Tag: intIndex = 0
            If lstSelect(1).Visible = True Then strValue = lstSelect(1).Tag: intIndex = 1
            intRow = Split(lbllst(intIndex).Tag, "|")(0)
            intCOl = Split(lbllst(intIndex).Tag, "|")(1)
            lng项目序号 = Val(vsfTab.TextMatrix(intRow, COL_tab项目序号))
            str项目名称 = Split(vsfTab.TextMatrix(intRow, COL_tab项目名称), "(")(0)
            strTmp = vsfTab.TextMatrix(intRow, COL_tab字符串)
            strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
            arrStr = Split(strTmp, ",")
            intType = Val(arrStr(1))
            int性质 = Val(arrStr(5))
            strPart = arrStr(7)
            
            blnAllow = True
        End If
    Else
        blnAllow = True
        If InStr(1, picTab.Tag, "|") = 0 Then Exit Function
        intRow = Split(picTab.Tag, "|")(0)
        intCOl = Split(picTab.Tag, "|")(1)
        lng项目序号 = Val(vsfTab.TextMatrix(intRow, COL_tab项目序号))
        str项目名称 = Split(vsfTab.TextMatrix(intRow, COL_tab项目名称), "(")(0)
        strTmp = vsfTab.TextMatrix(intRow, COL_tab字符串)
        strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
        arrStr = Split(strTmp, ",")
        intType = Val(arrStr(1))
        int频次 = Val(arrStr(3))
        int表示 = Val(arrStr(4))
        int性质 = Val(arrStr(5))
        strPart = arrStr(7)
        strHour = ""
        strValue = FraTable.Tag
    End If
    
    If blnAllow = True Then
        lngColor = 0
        vsfTab.TextMatrix(intRow, intCOl) = strHour & strText
        If cmdColor.Visible = True Then lngColor = cmdColor.Tag
        vsfTab.Cell(flexcpForeColor, intRow, intCOl, intRow, intCOl) = lngColor
        mblnEdit = True
    Else
        If strErrMsg <> "" Then GoTo ErrInfo
        Exit Function
    End If
    
    mrsCurve.Filter = 0
    int记录类型 = 1
    blnTrue = False
    '更新数据修改标志
    If blnAllow = True Then
        '--51282,刘鹏飞,2012-08-03,全天汇总显示录入时间(DYEY要求手工录入汇总时间H)
        strHour = Replace(Replace(strHour, "(", ""), "h)", "")
        If int表示 = 4 And int频次 = 1 And mbln录入小时 = True Then
            If InStr(1, strValue, ")") > 0 Then
                strHourOld = Replace(Replace(Split(strValue, ")")(0), "(", ""), "h", "")
                strValue = Split(strValue, ")")(1)
            End If
            '更新用户录入的汇总小时
            If Val(strHour) <> Val(strHourOld) Then
                blnTrue = True
                int记录类型 = 11
                GoTo ErrHour
            End If
        End If
ErrBegin:
        If strValue <> strText Then
ErrHour:
            mrsCurve.Filter = "项目序号=" & lng项目序号 & " and 项目名称='" & str项目名称 & "' And 记录类型=" & int记录类型 & " And 列号=" & intCOl - vsfTab.FixedCols + 1
            'Call OutputRsData(mrsCurve, True)
            If mrsCurve.RecordCount > 0 Then
                mrsCurve!未记说明 = lngColor
                If mrsCurve!状态 <> 1 And mrsCurve!状态 <> 3 Then '原有的数据 修改、删除后的状态始终为2
                    mrsCurve!状态 = 2
                    mrsCurve!数值 = IIf(blnTrue = True, strHour, strText)
                Else '对于新增数据的处理
                    If Trim(IIf(blnTrue = True, strHour, strText)) = "" Then
                        mrsCurve!状态 = 3
                        mrsCurve!数值 = ""
                    Else
                        mrsCurve!状态 = 1
                        mrsCurve!数值 = IIf(blnTrue = True, strHour, strText)
                    End If
                End If
                mrsCurve.Update
            Else '不存在记录就新增数据
                If Trim(strText) <> "" Then
                    strTime = GetAnimalItemTime(intRow, intCOl, strErrMsg)
                    If strErrMsg <> "" Then GoTo ErrInfo

                    gstrFields = "序号|分组名|数值|部位|标记|时间|项目序号|项目名称|复查|未记说明|数据来源|修改|显示|来源ID|共用|状态|列号|记录类型"
                    gstrValues = GetMaxID & "|2)体温表格项目|" & IIf(blnTrue = True, strHour, strText) & "|" & strPart & "|" & _
                        0 & "|" & strTime & "|" & lng项目序号 & "|" & str项目名称 & "|0|" & lngColor & "|0|0|0|0|0|1|" & intCOl - vsfTab.FixedCols + 1 & "|" & int记录类型
                    Call Record_Add(mrsCurve, gstrFields, gstrValues)
                End If
            End If
            mblnChage = True
            If blnTrue = True Then
                blnTrue = False
                int记录类型 = 1
                GoTo ErrBegin
            End If
        End If
    End If
    If blnAllow = True And blnVisible = True Then Call txtEdit_KeyPress(vbKeyEscape): mblnEdit = True
    
    WriteIntoVfgTab = True
        
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Exit Function
ErrInfo:
    If strErrMsg <> "" Then
        lblStb.Caption = strErrMsg: lblStb.ForeColor = 255
        vsfTab.TextMatrix(intRow, intCOl) = strValue
    End If
End Function

Private Function GetAnimalItemTime(ByVal intRow As Integer, ByVal intCOl As Integer, Optional strInfo As String = "", Optional IntMode As Integer = 0) As String
'--------------------------------------------------------------------------------
'功能:获取体温表格项目某频次的时间
'arrTime 返回信息 包括 开始时间 中点时间 结束时间
'IntMode 0 返回中间点时间 1,返回开始时间和结束时间 2 返回开始时间;中间点时间;结束时间
'---------------------------------------------------------------------------------
    Dim strTmp As String, lng项目序号 As Long, str项目名称 As String, int频次 As Integer, _
        int项目表示 As String, intType As Integer, intNO As Integer
    Dim arrStr() As String
    Dim strTime As String
    Dim rsTmp As New ADODB.Recordset
    Dim strBegin As String, strEnd As String
    Dim intHour As Integer
    Dim lngRow As Long
    Dim strCurrDate As String, strDate As String
    Dim strReturn As String
    Dim bln波动 As Boolean
    
    On Error GoTo Errhand
    
    strDate = mstrBegin
    strInfo = ""
    lngRow = intRow - vsfTab.FixedRows + 1
    strTmp = vsfTab.TextMatrix(intRow, COL_tab字符串)
    lng项目序号 = Val(vsfTab.TextMatrix(intRow, COL_tab项目序号))
    str项目名称 = vsfTab.TextMatrix(intRow, COL_tab项目名称)
    strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
    arrStr = Split(strTmp, ",")
    int频次 = Val(arrStr(3))
    int项目表示 = Val(arrStr(4))
    
    bln波动 = IsWaveItem(lng项目序号)
    
    '汇总/波动项目类型=2
    If int项目表示 = 4 Or bln波动 Then
        intType = 2
        If int频次 = 0 Then
            int频次 = 2
        ElseIf int频次 > 2 Then
            int频次 = 2
        End If
        
        '由参数确定汇总/波动项目今天录入昨天的数据还是当天的数据
        If Not mbln汇总当天 Then strDate = CDate(mstrBegin) - 1
    Else
        intType = 1
    End If
    
    '获取当前记录的频次
    intNO = intCOl - vsfTab.FixedCols + 1
    
    '根据类型，频次和序号 不可能找不到信息
    mrsTabTime.Filter = "类型=" & intType & " and 频次=" & int频次 & " and 序号=" & intNO
    If mrsTabTime.RecordCount = 0 Then
        strInfo = "请在护理项目管理中设置[" & IIf(intType = 2, "汇总项目", "体温表格项目") & "]时段信息!"
        Exit Function
    End If
    
    With mrsTabTime
        .MoveFirst
        intHour = CInt(24 / int频次)
        strBegin = Format(IIf(IsDate(Trim(Nvl(!开始))) = False, (Val(Nvl(!序号)) - 1) * intHour & ":00:00", !开始), "HH:mm:ss")
        strEnd = Format(IIf(IsDate(Trim(Nvl(!结束))) = False, Val(Nvl(!序号)) * intHour - 1 & ":59:59", !结束), "HH:mm:ss")
        If intNO = int频次 Then
            If strBegin >= strEnd Then
                strBegin = Format(strDate, "YYYY-MM-DD") & " " & strBegin
                strEnd = Format(DateAdd("d", 1, CDate(strDate)), "YYYY-MM-DD") & " " & strEnd
            Else
                strBegin = Format(strDate, "YYYY-MM-DD") & " " & strBegin
                strEnd = Format(strDate, "YYYY-MM-DD") & " " & strEnd
            End If
        Else
            If strBegin >= strEnd Then
                strBegin = Format(strDate, "YYYY-MM-DD") & " " & strBegin
                strEnd = strBegin
            Else
                strBegin = Format(strDate, "YYYY-MM-DD") & " " & strBegin
                strEnd = Format(strDate, "YYYY-MM-DD") & " " & strEnd
            End If
        End If
    End With
    
    '获取系统当前时间
    strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
    '提取中点时间
    intHour = DateDiff("H", CDate(strBegin), CDate(strEnd) + 0.00001) / 2
    strTime = DateAdd("H", intHour, CDate(strBegin)) '中点时间
    
    '汇总项目特殊处理
'    If int项目表示 = 4 Or bln波动 = True Then
'        '体温单开始当天不存在汇总数据录入
'        If Format(mstrBegin, "YYYY-MM-DD") = Format(mstrBTime, "YYYY-MM-DD") Then
'            strInfo = "汇总/波动项目[" & str项目名称 & "]在体温单开始当天不允许录入数据[体温单开始时间：" & Format(mstrBTime, "YYYY-MM-DD HH:mm:ss") & "]！"
'            GoTo ExitFunction
'        End If
'        GoTo ErrNext
'    End If
    
    '对于录入当天的数据 以当前时间为准(在当前时间符合项目的时间范围时)
    If CDate(strCurrDate) >= CDate(strBegin) And CDate(strCurrDate) <= CDate(strEnd) Then
        strTime = strCurrDate
    End If
    
    If CDate(strTime) < CDate(mstrBTime) Then
        strTime = Format(mstrBTime, "YYYY-MM-DD HH:mm:ss")
        If CDate(strTime) > CDate(strEnd) Then
           strInfo = "第" & lngRow & "列[" & str项目名称 & "]的结束时间：" & Format(strEnd, "YYYY-MM-DD HH:mm:ss") & "，不能小于[体温单开始时间：" & Format(mstrBTime, "YYYY-MM-DD HH:mm:ss") & "]！"
           GoTo ExitFunction
        End If
    End If
    
    If CDate(strTime) > CDate(mstrETime) Then
        strTime = Format(mstrETime, "YYYY-MM-DD HH:mm:ss")
        If CDate(strTime) < CDate(strBegin) Then
            If mbln出院 = False Then
                strInfo = "第" & lngRow & "列[" & str项目名称 & "]的开始时间：" & Format(strBegin, "YYYY-MM-DD HH:mm:ss") & "，已超出参数[超期录入天数：" & mintPreDays & "天]所指定的范围！"
            Else
                strInfo = "第" & lngRow & "列[" & str项目名称 & "]的开始时间：" & Format(strBegin, "YYYY-MM-DD HH:mm:ss") & "，不能大于[病人出院时间：" & Format(mstrETime, "YYYY-MM-DD HH:mm:ss") & "]!"
            End If
            GoTo ExitFunction
        End If
    End If
    
ErrNext:
    '检查病人转科后的补录时限
    If Not IsAllowInput(mT_Patient.lng病人ID, mT_Patient.lng主页ID, strEnd, strCurrDate) Then
        strInfo = "记录数据时间[" & strTime & "]有误！[超过数据补录的有效时限:" & mlngHours & "小时]"
        GoTo ExitFunction
    End If
    
    Select Case IntMode
        Case 0
            strReturn = Format(CDate(strTime), "YYYY-MM-DD HH:mm:ss")
        Case 1
           strReturn = Format(CDate(strBegin), "YYYY-MM-DD HH:mm:ss") & ";" & Format(CDate(strEnd), "YYYY-MM-DD HH:mm:ss")
        Case 2
        strReturn = Format(CDate(strBegin), "YYYY-MM-DD HH:mm:ss") & ";" & Format(CDate(strTime), "YYYY-MM-DD HH:mm:ss") & ";" & Format(CDate(strEnd), "YYYY-MM-DD HH:mm:ss")
    End Select
    
    GetAnimalItemTime = strReturn
ExitFunction:
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function InitRecordSet() As Boolean
'----------------------------------------------------------------
'功能:初始化记录集 包括部位信息，汇总项目时段，记录频次时段
'----------------------------------------------------------------
    On Error GoTo Errhand
    '提取所有部位信息
    mstrSQL = "Select 项目序号,部位,缺省项 From 体温部位"
    Call zlDatabase.OpenRecordset(mrsPart, mstrSQL, Me.Caption)
    
    '提取共用记录集信息
    Call InitPublicData
    
    InitRecordSet = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetMaxID() As Long
'----------------------------------------------------
'功能:获取记录mrsCurve中的最大序号
'----------------------------------------------------
    mrsCurve.Filter = 0
    mrsCurve.Sort = "序号 Desc"
    If mrsCurve.RecordCount = 0 Then
        GetMaxID = 1
    Else
        GetMaxID = Val(mrsCurve!序号) + 1
    End If
End Function

Private Function CheckValidata(ByVal intRow As Integer, ByVal intCOl As Integer, ByVal lngNO As Long, ByVal intType As Integer, ByVal int小数 As Integer, ByVal str值域 As String, _
    ByVal int表示 As Integer, ByVal lngLen As Long, strInfo As String, Optional strErrMsg As String = "") As Boolean
'----------------------------------------------------------------------------------------------------------------
'功能:检查数据合法性(对于体温曲线项目和表格项目的检查)
'参数:introw：哪一行 intCol： 那一列  lngNo:项目序号 intype： 项目类型 0数字类型 1 文字类型 str值域：项目值域
'   lngLen：项目长度  strInfo：要校验的文本值
'----------------------------------------------------------------------------------------------------------------
    Dim strName As String, strMsg As String
    Dim lngRow As Long
    Dim arrValue() As String
    Dim lngCount As Long, i As Integer, blnOk As Boolean, strText As String
    Dim blnAllow As Boolean
    
    strName = Split(vsfTab.TextMatrix(intRow, COL_tab项目名称), "(")(0)
    lngRow = intRow - vsfTab.FixedRows + 1
    
    If strInfo = "" Then
        CheckValidata = True
        Exit Function
    End If
    
    blnAllow = True
    
    If strName = "体重" Or strName = "身高" Then
        If IsNumeric(strInfo) Then
            blnAllow = True
        Else
            blnAllow = False
        End If
    End If
    
    '大便次数和排出量不检查
    If blnAllow = True Then blnAllow = IIf(InStr(1, "," & gint大便 & "," & gint入液 & ",", "," & lngNO & ",") > 0, False, True)
    
    If Not (intType = 0 And InStr(1, "0,4", int表示) <> 0) Then
        If LenB(StrConv(strInfo, vbFromUnicode)) > lngLen Then
            strMsg = "第" & lngRow & "行[" & strName & "]的值超长(最大长度:" & lngLen & "),请检查!"
            GoTo ErrInfo
        End If
    Else
        If intType = 0 Then
            If int表示 = 4 Or str值域 = "" Then
                str值域 = "0～" & IIf(lngLen - int小数 > 0, String(lngLen - int小数, "9"), "0") & IIf(int小数 > 0, "." & String(int小数, "9"), "")
            End If
            If lngNO <> 4 And lngNO <> 5 And blnAllow = True Then
                If Not IsNumeric(strInfo) Then
                    strMsg = strName & "数据录入错误" & Space(4) & "有效范围:" & str值域
                    GoTo ErrInfo
                End If
            End If
            
            If lngNO = 4 And strName = "血压" Then
                '--问题号:53505,修改人：李涛,血压可以录入文字说明：外出，未测等
                If strInfo = "外出" Or strInfo = "未测" Or strInfo = "拒测" Or strInfo = "请假" Then
                    CheckValidata = True
                    Exit Function
                Else
                    If InStr(1, strInfo, "/") = 0 Then
                        strMsg = "第" & lngRow & "行[血压]数据的格式错误：收缩压/舒张压！"
                        GoTo ErrInfo
                    End If
                    If Trim(Split(strInfo, "/")(0)) = "" Or Trim(Split(strInfo, "/")(1)) = "" Then
                        strMsg = "第" & lngRow & "行[血压]数据录入错误：收缩压/舒张压！"
                        GoTo ErrInfo
                    End If
                End If
            End If
            
            If UBound(Split(strInfo, "/")) > 1 And blnAllow = True Then
                strMsg = "第" & lngRow & "行[" & strName & "]数据录入错误，请检查！"
                GoTo ErrInfo
            End If
            
            '检查数据在有效范围内是否有效
            arrValue = Split(strInfo, "/")
            lngCount = UBound(arrValue)
            For i = 0 To lngCount
                blnOk = False
                strText = arrValue(i)
                If Not blnOk Then
                    If Not IsNumeric(strText) And blnAllow = True Then
                        strMsg = "第" & lngRow & "行[" & strName & "]数据录入错误" & Space(4) & "有效范围:" & str值域
                        GoTo ErrInfo
                    End If
                End If
                
                If Not blnOk And strText <> "" And blnAllow = True Then strText = Format(Val(strText), "#0" & IIf(int小数 > 0, ".", "") & String(int小数, "0"))
                
                If int表示 <> 4 And blnAllow = True Then
                    If Len(Replace(strText, ".", "")) > lngLen Then
                        strMsg = "第" & lngRow & "行[" & strName & "]的值超长(最大长度:" & lngLen & "),请检查!"
                        GoTo ErrInfo
                    End If
                End If
                
                If IsNumeric(Split(str值域, "～")(0)) And IsNumeric(strText) Then
                    If blnAllow = True Then   '大便次数不进行有效范围检查
                        If Not (Val(strText) >= Split(str值域, "～")(0) And Val(strText) <= Split(str值域, "～")(1)) Then
                            strMsg = strName & "超出有效范围(" & str值域 & "),请检查!"
                            GoTo ErrInfo
                        End If
                    End If
                End If
                
                arrValue(i) = strText
            Next i
            strInfo = Join(arrValue, "/")
        End If
    End If
    
    CheckValidata = True
    Exit Function
    
    CheckValidata = True
    Exit Function
ErrInfo:
    strErrMsg = strMsg
End Function

Private Function ChangeCurveTime() As Boolean
'-----------------------------------------------------------
'功能:检查用户修改体温曲线时点时间是否合法
'-----------------------------------------------------------
    Dim strBegin As String, strEnd As String, strTime As String
    strEnd = Format(mstrEnd, "HH:mm")
    strBegin = Format(mstrBegin, "HH:mm")
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo Errhand
    
    If Format(dkpTime.Value, "HH:mm") = Format(mArrdkpTime(dkpTime.Tag), "HH:mm") Then ChangeCurveTime = True: Exit Function
    
    If Format(dkpTime.Value, "HH:mm") < strBegin And Format(dkpTime.Value, "HH:mm") > strEnd Then
        lblStb.Caption = "体温曲线时间只能在 " & strBegin & "" & strEnd & " 之间"
        lblStb.ForeColor = 255
        dkpTime.Value = Format(mArrdkpTime(dkpTime.Tag), "HH:mm")
        If dkpTime.Enabled = True Then dkpTime.SetFocus
        Exit Function
    End If
    
    If dkpTime.Value = Format(mstrBegin, "HH:mm") Then
        strTime = Format(mstrBegin, "HH:mm:ss")
    ElseIf dkpTime.Value = Format(mstrEnd, "HH:mm") Then
        strTime = Format(mstrEnd, "HH:mm:ss")
    Else
        strTime = Format(dkpTime.Value, "HH:mm:ss")
    End If
    strTime = Format(Format(mstrBegin, "YYYY-MM-DD") & " " & strTime, "YYYY-MM-DD HH:mm:ss")
    
    '检查修改的时间是否已经存在数据
    mstrSQL = "select 1 From 病人护理文件 a,病人护理数据 b" & vbNewLine & _
        " where A.ID=B.文件ID and A.ID=[1] and A.病人ID=[2] and A.主页ID=[3] And nvl(A.婴儿,0)=[4]" & vbNewLine & _
        " and B.发生时间=[5]"
        
    If mblnMove Then
        mstrSQL = Replace(mstrSQL, "病人护理文件", "H病人护理文件")
        mstrSQL = Replace(mstrSQL, "病人护理数据", "H病人护理数据")
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "检查时间", mT_Patient.lng文件ID, mT_Patient.lng病人ID, mT_Patient.lng主页ID, mT_Patient.lng婴儿, CDate(strTime))
    
    If rsTemp.RecordCount > 0 Then
        lblStb.Caption = "改时点已经存在数据,请重新设置时间."
        lblStb.ForeColor = 255
        dkpTime.Value = Format(mArrdkpTime(dkpTime.Tag), "HH:mm")
        If dkpTime.Enabled = True Then dkpTime.SetFocus
        Exit Function
    End If
    
    '检查是否超出设置时间
    If Not CheckDateTime(0, "", strTime) Then
        dkpTime.Value = Format(mArrdkpTime(dkpTime.Tag), "HH:mm")
        If dkpTime.Enabled = True Then dkpTime.SetFocus
        Exit Function
    End If
    
    '修改本时间段内的所有体温曲线数据时点
    mrsCurve.Filter = 0
    mrsCurve.Filter = "分组名='1)体温曲线项目' And 时间='" & Format(mArrdkpTime(dkpTime.Tag), "YYYY-MM-DD HH:mm:ss") & "'"
    If mrsCurve.RecordCount > 0 Then mblnChage = True: mblnCurveChange = True
    
    '状态 1新增 ,2 修改 ,3新增后删除(未保存),4 只是修改时间
    With mrsCurve
        Do While Not .EOF
            !时间 = strTime
             If Val(!状态) <> 1 And Val(!状态) <> 3 Then
                If Val(!状态) = 2 Then
                    mrsCurve!状态 = 2
                Else
                    mrsCurve!状态 = 4
                End If
            Else
                If mrsCurve!数值 = "" And mrsCurve!未记说明 = "" Then
                    mrsCurve!状态 = 3
                Else
                    mrsCurve!状态 = 1
                End If
            End If
            .Update
        .MoveNext
        Loop
    End With
   
    '跟新时间数组的值
    mArrdkpTime(dkpTime.Tag) = Format(strTime, "YYYY-MM-DD HH:mm:ss")
    
    ChangeCurveTime = True
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function is大便或入液(ByVal intType As Integer) As Boolean
'检查是否是大便项目或入夜项目  大便项目序号=10 入夜=9
'intType=1 为大便项目 否则为入液项目
    Dim lngItemNO As Long
    Dim strKey As String
    Dim rsObj As New ADODB.Recordset
    Dim strTmp As String, strName As String, arrStr
    On Error GoTo Errhand
    
    If vsfTab.Col < vsfTab.FixedCols Or vsfTab.Row < vsfTab.FixedRows Then Exit Function
    If mblnInit = False Then Exit Function
    
    '提取项目序号
    lngItemNO = Val(vsfTab.TextMatrix(vsfTab.Row, COL_tab项目序号))
    If intType = 1 Then
        If lngItemNO <> 10 Then Exit Function
    Else
        If lngItemNO <> 9 Then Exit Function
    End If
    strName = vsfTab.TextMatrix(vsfTab.Row, COL_tab项目名称)
    strTmp = vsfTab.TextMatrix(vsfTab.Row, COL_tab字符串)
    strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
    arrStr = Split(strTmp, ",")
    
    '检查记录频次和项目表示
    If vsfTab.Col > vsfTab.FixedCols + Val(arrStr(3)) - 1 Then Exit Function
    If InStr(1, ",2,3,5,", "," & Val(arrStr(4)) & ",") > 0 Then Exit Function
    
    '检查是否是同步的数据
    mrsCurve.Filter = "项目序号=" & lngItemNO & " and 项目名称='" & strName & "'" & _
        "   and 列号=" & vsfTab.Col - vsfTab.FixedCols + 1
    If mrsCurve.RecordCount > 0 Then
        If InStr(1, ",0,9,", "," & Val(mrsCurve!数据来源) & ",") = 0 Then
            Exit Function
        End If
    End If
    
    is大便或入液 = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


