VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.Unicode.9600.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmCaseTendBodySetData 
   Caption         =   "�������ݱ༭"
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
   StartUpPosition =   2  '��Ļ����
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
      Caption         =   "��������/����"
      BeginProperty Font 
         Name            =   "����"
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
         Name            =   "����"
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
         Name            =   "����"
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
               Caption         =   "�E"
               BeginProperty Font 
                  Name            =   "����"
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
                  Name            =   "����"
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
                     Name            =   "����"
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
                     Name            =   "����"
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
                  Name            =   "����"
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
               Name            =   "����"
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
                  Name            =   "����"
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
               Caption         =   "¼�룺"
               BeginProperty Font 
                  Name            =   "����"
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
               Caption         =   "ѡ��"
               BeginProperty Font 
                  Name            =   "����"
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
               Name            =   "����"
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
         Name            =   "����"
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
               Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
                  Name            =   "����"
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
               Caption         =   "ʱ��:"
               BeginProperty Font 
                  Name            =   "����"
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
               Name            =   "����"
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
            Caption         =   "00:00��05:59"
            BeginProperty Font 
               Name            =   "����"
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
            Name            =   "����"
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
         Begin VB.PictureBox picδ�� 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0FFFF&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "����"
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
            Begin VB.ListBox lstδ�� 
               Appearance      =   0  'Flat
               BackColor       =   &H80000018&
               BeginProperty Font 
                  Name            =   "����"
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
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12806
            Key             =   "ZLNOTE"
            Object.ToolTipText     =   "��Ϣ��ʾ��Ϣ"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   2
            MinWidth        =   2
            Text            =   "��������"
            TextSave        =   "��������"
            Key             =   "ZLDataType"
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "����"
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
         Name            =   "����"
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
    COL_������ = 1
    COL_�ַ��� = 2
    COL_��Ŀ��� = 3
    COL_��Ŀ���� = 4
    col_���� = 5
    col_��ɫ = 6
    col_���� = 7
    COL_��λ = 8
    Col_δ��˵�� = 9
End Enum

Private Enum TYPE_Tab
    COL_tab������ = 0
    COL_tab�ַ��� = 1
    COL_tab��Ŀ��� = 2
    COL_TabNull = 3
    COL_tab��Ŀ���� = 4
End Enum

Private Enum TYPE_Oper
    Col_OperNull = 0
    Col_OperTime = 1
    Col_OperType = 2
End Enum

Private Enum Enum_No
     Item���� = 1
     Item���� = 2
     Item���� = -1
     Item����ѹ = 4
     Item����ѹ = 5
End Enum

Private Type Type_Item
    ���� As String
    ֵ�� As String
    ��Ŀ���� As Integer
    ��ĿС�� As Double
    ��¼Ƶ�� As Integer
    ��Ŀ��ʾ As Integer
    ��Ŀ���� As Integer
    ��Ŀ���� As Long
    ��λ As String
    ��Ŀ�� As Long
    ��Ŀ�� As String
    ��Ժ�ײ� As Integer
End Type

Private Type type_Patient
    lng����ID As Long
    lng��ҳID As Long
    lng�ļ�ID As Long
    lngӤ�� As Long
    lng����ID As Long
    lng����ȼ� As Long
    lng����ID As Long
End Type
Private mT_Patient As type_Patient

'�������±���
Private Type Type_OptRow
    ���� As Integer
    �ϱ� As Integer
    �±� As Integer
End Type

Private mOptRow As Type_OptRow

'������:
Private mcbrToolBar As CommandBar

Private mblnStart As Boolean
Private mblnFileBack As Boolean
Private mblnScroll As Boolean
Private mblnEdit As Boolean
Private mblnAllRefresh As Boolean
Private marrTime() As String
Private Const mFontSize As Integer = 9 '���������ʼ��СΪ9������
Private mintPreDays As Integer '����¼��ʱ��
Private mintBigSize As Integer '�Ƿ�Ŵ�
Private mlngHours As Long '���ݲ�¼ʱ��
Private mbln���ܵ��� As Boolean
Private mbln¼��Сʱ As Boolean  'ȫ�������ʾ¼��ʱ��
Private mstrActiveItem As String
Private mint����Ӧ�� As Integer
Private mblnEdit���� As Boolean
Private mstrBegin As String 'ĳ��ʱ���Ŀ�ʼ�ͽ���ʱ�� 00:00-05:59
Private mstrEnd As String
Private mstrBTime As String  '���µ��Ŀ�ʼʱ��ͽ���ʱ��
Private mstrETime As String
Private mstrOverDate As String '����ʵ�ʳ�Ժʱ��(�����µ�ʵ����ֹʱ��)
Private mstrDate As String '���µ���ǰҳ�ĵ�һ��ʱ��
Private mblnChage As Boolean
Private mblnCurveChange As Boolean
Private mblnOK As Boolean
Private mblnMove As Boolean
Private mstrSQL As String
Private mblnInit As Boolean
Private mstrδ��˵�� As String
Private mArrdkpTime() As Variant
Private mArrModfy() As Integer
Private mArrValue() As String
Private marrDate() As Integer
Private mstrPart As String
Private mbln��Ժ As Boolean
Private mblnResize As Boolean
Private mbln����������ʾ As Boolean

'��¼��
Private mrsPart As New ADODB.Recordset '���²�λ
Private mrsCurve As New ADODB.Recordset '��������
Private mrsNote As New ADODB.Recordset '���±�
Private mrsOper As New ADODB.Recordset '����
Private mrsRecodeID As New ADODB.Recordset '��¼����������Ŀ�ļ�¼ID��ʱ��

Public Function ShowEditor(ByVal frmParent As Object, ByVal strParam As String, ByVal strTime As String, ByVal strDayTime As String, _
    ByVal int����Ӧ�� As Integer, Optional blnMove As Boolean = False, Optional ByVal bytSize As Byte = 0) As Boolean
'----------------------------------------------------------------------------------------------------------------------------------------------------------
'����:�������µ��༭����
'����:frmParent ������,strParam ��ʽ:����ID;��ҳId;�ļ�ID;Ӥ��;����ID;������ȼ�  strTime ĳ��ʱ���ʱ�䷶Χ ����:2011-01-25 00:00:00;2011-01-25 05:59:59

'     strDayTime һ�ܿ�ʼʱ��; int����Ӧ��=2 ��ʾ���������ʹ��� blnMove ��ʷ�����Ƿ�ת��
'     bytSize 0-9������ 1-12������
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
    mbln���ܵ��� = False
    mstrOverDate = ""
    
    mT_Patient.lng����ID = 0
    mT_Patient.lng����ȼ� = 3
    
    mT_Patient.lng����ID = Val(arrParam(0))
    mT_Patient.lng��ҳID = Val(arrParam(1))
    mT_Patient.lng�ļ�ID = Val(arrParam(2))
    mT_Patient.lngӤ�� = Val(arrParam(3))
    
    If UBound(arrParam) > 3 Then mT_Patient.lng����ID = arrParam(4)
    If UBound(arrParam) > 4 Then mT_Patient.lng����ȼ� = arrParam(5)
    
    If mT_Patient.lng����ID = 0 And mT_Patient.lng��ҳID = 0 And mT_Patient.lng����ID = 0 Then
        MsgBox "�ļ�ID,����ID,��ҳID����Ϊ��,����!", vbInformation, gstrSysName
        Exit Function
    End If
    
    mstrBegin = Format(Split(strTime, ";")(0), "YYYY-MM-DD HH:mm:ss")
    mstrEnd = Format(Split(strTime, ";")(1), "YYYY-MM-DD HH:mm:ss")
    mstrDate = strDayTime
    
    If Not ChekPatientOut(mT_Patient.lng�ļ�ID, mT_Patient.lng����ID, mT_Patient.lng��ҳID, mT_Patient.lngӤ��) Then Exit Function
    mintBigSize = bytSize
    Me.Font.Size = IIf(mintBigSize = 0, 9, 12)
    mint����Ӧ�� = int����Ӧ��
    mblnEdit���� = True
    mblnMove = blnMove
    
    If Not OpenPatientInfo Then Exit Function
    
    '����ļ��Ƿ�鵵
    mblnFileBack = CheckFileBack(mT_Patient.lng�ļ�ID, mblnMove)
    '��ʼ��������
    Call InitCommandBars
    '��ȡ����
    Call GetTableRowName
    Call zlRefreshData
    mblnInit = True
    mblnResize = True
    Me.Show 1
    
    ShowEditor = mblnOK
End Function

Public Function ChekPatientOut(ByVal lng�ļ�ID As Long, ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal intBaby As Long) As Boolean
'-----------------------------------------------------------------------------------------------
'����:��ȡ���µ���ʼʱ��ͽ���ʱ�� ����鲡���Ƿ��Ժ
'-----------------------------------------------------------------------------------------------
    Dim strSQL As String, strNewSql As String
    Dim strBeginDate As String, strEndDate As String
    Dim rsTemp As New ADODB.Recordset
    Dim strMaxDate As String, strCurrDate As String
    Dim intDay As Integer
    
    mbln��Ժ = False
    On Error GoTo Errhand
    
    'mintBigSize =  zldatabase.GetPara("�����ļ���ʾģʽ", glngSys, 1255, 0)
    mintPreDays = Val(zlDatabase.GetPara("����¼�뻤����������", glngSys, 1255, "1"))
    gintHourBegin = zlDatabase.GetPara("���¿�ʼʱ��", glngSys, 1255, 4)
    mlngHours = Val(Mid(Val(zlDatabase.GetPara("���ݲ�¼ʱ��", glngSys)), 1, 6))
    mbln���ܵ��� = (Val(zlDatabase.GetPara("���ܲ�����ʾ��������", glngSys, 1255, 0)) = 1)
    '51282,������,2012-08-03,ȫ�������ʾ¼��ʱ��(DYEYҪ���ֹ�¼�����ʱ��H
    mbln¼��Сʱ = (Val(zlDatabase.GetPara("ȫ�������ʾ¼��ʱ��", glngSys, 1255, 0)) = 1)
    mbln����������ʾ = (Val(zlDatabase.GetPara("���������(����/����)��ʽ¼��", glngSys, 1255, 0)) = 1)
    If mintPreDays < 0 Then mintPreDays = 0
        
    '��ȡӤ��ҽ����Ϣ(ת�ƣ���Ժ),����ҽ����ҽ����ϢΪ׼��������ĸ�׳�Ժ����Ϊ׼
    strNewSql = "   (SELECT /*+ RULE */  ����ID,��ҳID,Ӥ��ʱ��,DECODE(nvl(Ӥ��,0),0, DECODE(NVL(��Ժ����,''),'',0,1), DECODE(NVL(Ӥ��ʱ��,''),'',0,1))��¼" & vbNewLine & _
                "       FROM (SELECT A.����ID,A.��ҳID,B.��ʼִ��ʱ�� Ӥ��ʱ��, A.��Ժ����,B.Ӥ��" & vbNewLine & _
                "           FROM ������ҳ A," & vbNewLine & _
                "               (SELECT B.����ID, B.��ҳID, B.Ӥ��, ��ʼִ��ʱ��" & vbNewLine & _
                "                FROM ����ҽ����¼ B, ������ĿĿ¼ C" & vbNewLine & _
                "                WHERE B.������ĿID + 0 = C.ID AND B.ҽ��״̬ = 8 AND nvl(B.Ӥ��,0)<>0  AND C.��� = 'Z'" & vbNewLine & _
                "                AND EXISTS (SELECT 1 FROM TABLE(CAST(F_STR2LIST('3,5,11') AS ZLTOOLS.T_STRLIST))" & vbNewLine & _
                "                               WHERE C.�������� = COLUMN_VALUE) And  B.����ID = [2] AND B.��ҳID = [3] AND B.Ӥ��(+) = [4]) B" & vbNewLine & _
                "           WHERE A.����ID = [2] AND A.��ҳID = [3] AND A.����ID = B.����ID(+) AND A.��ҳID = B.��ҳID(+)" & vbNewLine & _
                "           ORDER BY B.��ʼִ��ʱ�� DESC)" & vbNewLine & _
                "       WHERE ROWNUM < 2)  E"
    
    strSQL = _
       "Select Decode(b.����ʱ��,Null,a.��ʼ,b.����ʱ��) As ��ʼ,decode(E.��¼,0,Decode(Sign(NVL(E.Ӥ��ʱ��,a.��ֹ) - d.����ʱ��), 1,NVL(E.Ӥ��ʱ��,a.��ֹ) ,d.����ʱ��),NVL(E.Ӥ��ʱ��,a.��ֹ)) ��ֹ,E.��¼" & vbNewLine & _
        "       From" & vbNewLine & _
        "       (Select ����ID,��ҳid,Min(��ʼʱ��) as ��ʼ,Max(Nvl(��ֹʱ��,sysdate)) as ��ֹ" & vbNewLine & _
        "       From ���˱䶯��¼" & vbNewLine & _
        "       Where ��ʼʱ�� is Not Null And ����ID=[2] And ��ҳID=[3] Group By ����ID,��ҳid) a," & vbNewLine & _
        "       (Select ����ID,��ҳid,����ʱ�� From ������������¼ Where ����ID =[2] And ��ҳID =[3] And ���=[4]) b," & vbNewLine & _
        "       (SELECT NVL(����ʱ��,SYSDATE) ����ʱ�� FROM (select max(����ʱ��) ����ʱ�� from ���˻����ļ� A,���˻������� B" & vbNewLine & _
        "       where A.ID=B.�ļ�ID and A.ID=[1] and A.����ID=[2] and A.��ҳID=[3] and A.Ӥ��=[4])) d," & vbNewLine & _
        strNewSql & vbNewLine & _
        "       Where a.����ID=E.����ID And A.��ҳID=E.��ҳID And a.����id=b.����id(+) And a.��ҳid=b.��ҳid(+)"
        
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "mdlPrint", lng�ļ�ID, lng����ID, lng��ҳID, intBaby)
    If rsTemp.RecordCount > 0 Then
        rsTemp.MoveFirst
        strBeginDate = Format(rsTemp!��ʼ, "YYYY-MM-DD HH:MM:SS")
        strEndDate = Format(rsTemp!��ֹ, "YYYY-MM-DD HH:MM:SS")
        mbln��Ժ = Not (Val(rsTemp!��¼) = 0)
    Else
        MsgBox "�޴˲��˱���סԺ��Ϣ,����!", vbInformation, gstrSysName
        Exit Function '�������˱䶯��Ϣ�˳�
    End If
    
    '��ȡ�û����õ����µ���ʼʱ��(Ӥ���Գ���ʱ��Ϊ׼)
    If intBaby = 0 Then
        strSQL = "select ��ʼʱ�� from ���˻����ļ� where ID=[1] and ����ID=[2] and ��ҳid=[3] and nvl(Ӥ��,0)=[4]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���µ���ʼʱ��", lng�ļ�ID, lng����ID, lng��ҳID, intBaby)
        If rsTemp.RecordCount <> 0 Then
            strBeginDate = Format(rsTemp!��ʼʱ��, "YYYY-MM-DD HH:mm:ss")
        End If
    End If
    
    strCurrDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm:ss")

    mstrBTime = strBeginDate
    mstrOverDate = strEndDate
    mstrETime = strEndDate
    If CDate(mstrETime) < CDate(Format(strCurrDate, "YYYY-MM-DD HH:mm:ss")) And Not mbln��Ժ Then mstrETime = CDate(Format(strCurrDate, "YYYY-MM-DD HH:mm:ss"))
    If mstrBTime > mstrETime Then mstrBTime = mstrETime
    If mstrDate < mstrBTime Then mstrDate = mstrBTime
    
    '���˳�Ժ�ѳ�Ժʱ��Ϊ��ֹʱ��
    If mbln��Ժ = True Then
        '��Ժʱ�����Ժʱ�������ͬһ�У��򽫳�Ժʱ�����һ�У���������:��ԺҲҪ¼�����£�
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
    '��ȡ������Ϣ
    mstrSQL = "Select ��Ժ����ID from ������ҳ Where ����id=[1] And ��ҳid=[2] "
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, mT_Patient.lng����ID, mT_Patient.lng��ҳID)
    If rsTmp.BOF = False Then
        mT_Patient.lng����ID = Val(zlCommFun.Nvl(rsTmp("��Ժ����ID").Value))
    End If
    
    '��ȡ����ȼ�
    mstrSQL = "Select zl_PatitTendGrade([1],[2]) As ����ȼ� From dual"
    Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, mT_Patient.lng����ID, mT_Patient.lng��ҳID)
    If rsTmp.BOF = False Then mT_Patient.lng����ȼ� = zlCommFun.Nvl(rsTmp("����ȼ�"), 3)
    
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
'����:��ʼ��������
'--------------------------------------------------------------------------------
    Dim cbrControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim cbrLable As CommandBarControl
    Dim cbrPop As CommandBarControl
    Dim CtlFont As StdFont
    
    On Error GoTo Errhand
    
     '��ʼ����
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbsMain.ActiveMenuBar.Title = "�˵���"
    cbsMain.ActiveMenuBar.Visible = False
    
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = False
        .SetIconSize False, 24, 24
        .SetIconSize True, 16, 16
        .UseSharedImageList = False 'ImageList��ʽʱ,��ͬһApp�й���,��AddImageList֮ǰ����ΪFalse
        Set CtlFont = .Font
        If CtlFont Is Nothing Then
            Set CtlFont = Me.Font
        End If
        CtlFont.Size = IIf(mintBigSize = 0, 9, 12)
        Set .Font = CtlFont
    End With

  '------------------------------------------------------------------------------------------------------------------
    '����������
    Set mcbrToolBar = cbsMain.Add("��׼", xtpBarTop)
    mcbrToolBar.ShowTextBelowIcons = False
    mcbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
    mcbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
    
    With mcbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Save, "����")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Reuse, "ȡ��")
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Curve_Show, "������ʾ"): cbrControl.ToolTipText = "��������������ʾ"
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_NewItem, "�����Ŀ"): cbrControl.ToolTipText = "��ӻ��Ŀ": cbrControl.BeginGroup = True

        Set cbrPop = .Add(xtpControlButtonPopup, conMenu_Edit_Append, "���⴦��")
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 0, "����", -1, False): cbrControl.IconId = 1: cbrControl.Parameter = ""
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 1, "�೦[E]", -1, False):  cbrControl.IconId = 1: cbrControl.Parameter = "E"
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 2, "�೦����[/E]", -1, False):  cbrControl.IconId = 1: cbrControl.Parameter = "/E"
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 3, "���ʧ��[��]", -1, False): cbrControl.IconId = 1: cbrControl.Parameter = "��"
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 4, "�˹�����[��]", -1, False): cbrControl.IconId = 1: cbrControl.Parameter = "��"
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 5, "����[C]", -1, False):   cbrControl.IconId = 1: cbrControl.Parameter = "C"
        Set cbrControl = cbrPop.CommandBar.Controls.Add(xtpControlButton, conMenu_Edit_Append * 10 + 6, "��������[/C]", -1, False):   cbrControl.IconId = 1: cbrControl.Parameter = "/C"
        
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����")
        cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
        
    End With
    
    '��λ������
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
    
    '���ڲ�¼
    '------------------------------------------------------------------------------------------------------------------
    Set cbrLable = mcbrToolBar.Controls.Add(xtpControlLabel, conMenu_View_Option, "ʱ��")
    cbrLable.flags = xtpFlagRightAlign
    
    Set cbrCustom = mcbrToolBar.Controls.Add(xtpControlCustom, conMenu_View_Location, "")
    dkpDate.Visible = True
    cbrCustom.Handle = dkpDate.hWnd
    cbrCustom.flags = xtpFlagRightAlign
    
    '�����
    With cbsMain.KeyBindings
        .Add FALT, Asc("0"), conMenu_Edit_Append * 10
        .Add FALT, Asc("1"), (conMenu_Edit_Append * 10 + 1)
        .Add FALT, Asc("2"), (conMenu_Edit_Append * 10 + 2)
        .Add FALT, Asc("3"), (conMenu_Edit_Append * 10 + 3)
        .Add FALT, Asc("4"), (conMenu_Edit_Append * 10 + 4)
        .Add FALT, Asc("5"), (conMenu_Edit_Append * 10 + 5)
        .Add FALT, Asc("6"), (conMenu_Edit_Append * 10 + 6)
        
        .Add FCONTROL, Asc("D"), conMenu_Edit_Curve_Show '������ʾ
        .Add FCONTROL, Asc("A"), conMenu_Edit_NewItem '��ӻ��Ŀ
        .Add FCONTROL, Asc("S"), conMenu_Edit_Save '����
        .Add FCONTROL, Asc("R"), conMenu_Edit_Reuse 'ȡ��
        .Add 0, VK_F1, conMenu_Help_Help
    End With
    
    Call InitDateTimeRange(marrTime, gintHourBegin)
     
    '���ر��ؼ�
    Call InitTabControl
    
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitTabControl()
'--------------------------------------------------------------------------------
'����:��ʼ��TabControl
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
        
        Set tabItem = .InsertItem(1, "��������", picCurve.hWnd, 0)
        tabItem.Tag = "����"
        Set tabItem = .InsertItem(2, "���±��", picTab.hWnd, 0)
        tabItem.Tag = "���"
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
'����:���ñ��ѡ����
'------------------------------------
    Dim intRow As Integer, intCOl As Integer
    Dim intOldRow As Integer, intOldCol As Integer
    
    If mblnInit = False Then Exit Sub
    
    If tbcThis.Selected.Tag = "����" Then
        vsfCurve.SetFocus
        If blnInit = True Then
            intOldRow = vsfCurve.Row
            intOldCol = vsfCurve.Col
            intRow = vsfCurve.Row
            intCOl = col_����
            If intRow = vsfCurve.Row And intCOl = vsfCurve.Col Then
                vsfCurve.Col = COL_��λ
            End If
            vsfCurve.Col = col_����
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
    ElseIf tbcThis.Selected.Tag = "���" Then
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
'����:��ʼ������������Ŀ
'����:���б�ͷ����Ϣ
'-------------------------------------------------------
    Dim varTabName() As String, varCode() As String
    Dim intRow As Integer, intCOl As Integer
    
    If strTabName = "" Then Exit Sub
    varTabName = Split(strTabName, ";")
    
    With vsfCurve
        .Rows = UBound(varTabName) + 2
        .Cols = 0
        
        .NewColumn "", 255, 4
        .NewColumn "������", 1500 + 1500 * mintBigSize / 3, 1
        .NewColumn "�ַ���", 0, 1
        .NewColumn "��Ŀ���", 0, 1
        .NewColumn "��Ŀ����", 1200 + 1200 * mintBigSize / 3, 1
        .NewColumn "����", 2300 + 2300 * mintBigSize / 3, 1, , 4
        .NewColumn "����", 300 + 300 * mintBigSize / 3, 0
        .NewColumn "���Ժϸ�", 900 + 900 * mintBigSize / 3, 4
        .NewColumn "��λ", 1000 + 1000 * mintBigSize / 3, 4
        .NewColumn "δ��˵��", 1080 + 1080 * mintBigSize / 3, 4, "...", 1
        .Body.RowHeight(0) = 300 + 300 * mintBigSize / 3
        .FixedCols = 5
        .FixedRows = 1
        
        .Body.Font.Size = mFontSize + mFontSize * mintBigSize / 3
        .Body.ColHidden(COL_�ַ���) = True
        .Body.ColHidden(COL_��Ŀ���) = True
        .Body.WordWrap = True
        .Body.MergeCells = flexMergeRestrictColumns
        .Body.MergeCol(COL_������) = True
        .Body.MergeRow(0) = True
        
        For intRow = .FixedRows To .Rows - 1
            varCode = Split(varTabName(intRow - 1), "'")
            If UBound(varCode) > 2 Then
                .TextMatrix(intRow, COL_������) = varCode(0)
                .TextMatrix(intRow, COL_�ַ���) = varCode(1)
                .TextMatrix(intRow, COL_��Ŀ���) = varCode(2)
                .TextMatrix(intRow, COL_��Ŀ����) = varCode(3)
                If varCode(0) = "2)���±�˵��" Then
                    Select Case Val(varCode(2))
                        Case 2
                            mOptRow.�ϱ� = intRow
                        Case 4
                            mOptRow.���� = intRow
                        Case 6
                            mOptRow.�±� = intRow
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
'����:��ʼ����������¼����
'-------------------------------------------------------
    Dim intRow As Integer, intCOl As Integer
        
    With vsfOper
        .Rows = 2
        .Cols = 0
        
        .NewColumn "", 255, 4
        .NewColumn "ʱ��", 1000 + 1000 * mintBigSize / 3, 4, , 4
        .NewColumn "����", 2000 + 2000 * mintBigSize / 3, 4, "����|����|��������", 1
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
'����:��ʼ�����±����Ŀ
'����:���б�ͷ����Ϣ(������������Ŀ)
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
        .ColHidden(COL_tab������) = True
        .ColHidden(COL_tab�ַ���) = True
        .ColHidden(COL_tab��Ŀ���) = True
        .WordWrap = True
        .ScrollBars = flexScrollBarBoth
        .SelectionMode = flexSelectionByRow
        
        '��ʼ��ͷ
        For intCOl = .FixedCols - 1 To .Cols - 1
            If intCOl = .FixedCols - 1 Then
                .TextMatrix(0, intCOl) = "����/Ƶ��"
            Else
                .TextMatrix(0, intCOl) = intCOl - .FixedCols + 1
                .ColWidth(intCOl) = 1200 + 1200 * mintBigSize / 3
            End If
        Next intCOl
        
        For intRow = 1 To .Rows - 1
            varCode = Split(varTabName(intRow - 1), "'")
            .TextMatrix(intRow, COL_tab������) = varCode(0)
            .TextMatrix(intRow, COL_tab�ַ���) = varCode(1)
            .TextMatrix(intRow, COL_tab��Ŀ���) = varCode(2)
            .TextMatrix(intRow, COL_TabNull) = ""
            .TextMatrix(intRow, COL_tab��Ŀ����) = varCode(3)
        Next intRow
        
        .ColWidth(COL_tab��Ŀ����) = 1200 + 1200 * mintBigSize / 3
        .RowHeight(-1) = 300 + 300 * mintBigSize / 3
                
        .Cell(flexcpAlignment, 0, .FixedCols, .Rows - 1, .Cols - 1) = flexAlignCenterCenter
        .Cell(flexcpText, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = ""
    End With
End Sub

Private Sub AdjustRowFlag(ByRef objVsf As Object, ByVal intRow As Integer)
    '-----------------------------------------------------------------------------------------
    '����:
    '����:
    '-----------------------------------------------------------------------------------------
    If objVsf.FixedCols = 0 Then Exit Sub
    
    If Not (objVsf.Cell(flexcpPicture, intRow, COL_TabNull) Is Nothing) Then Exit Sub
    Set objVsf.Cell(flexcpPicture, 1, COL_TabNull, objVsf.Rows - 1, COL_TabNull) = Nothing
    Set objVsf.Cell(flexcpPicture, intRow, COL_TabNull) = ils16.ListImages(1).Picture
    
End Sub

Private Function InitTime() As String
'--------------------------------------------------------
'����:��ȡһ���ʱ�����Ϣ
'--------------------------------------------------------
    Dim i As Integer
    Dim strName As String
    
    Call InitDateTimeRange(marrTime, gintHourBegin)
    For i = 0 To UBound(marrTime) - 1
        strName = strName & ";" & Format(Split(marrTime(i), ",")(0), "HH:mm") & "��" & Format(Split(marrTime(i), ",")(1), "HH:mm")
    Next i
    
    If Left(strName, 1) = ";" Then strName = Mid(strName, 2)
    
    strName = "��Ŀ\ʱ�䷶Χ" & ";" & strName
    InitTime = strName
End Function

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngRow As Long, lngCol As Long, lngItemNO As Long, intType As Integer, i As Integer
    Dim strValue As String, strPart As String, strPart1 As String, strName As String
    Dim strTime As String, strErrMsg As String
    Dim strTmp As String, arrStr
    Dim cbrCheck As CommandBarControl
    
    Select Case Control.Id
    
        Case conMenu_Edit_Save '����
        
            If picEdit.Visible = True Then
                Call vsfTab_EnterCell
            End If
            If Not ChangeCurveTime Then Exit Sub
            If Not SaveData Then Exit Sub
            Call GetTableRowName
            Call zlRefreshData
            Call SetColSelect
            
        Case conMenu_Edit_Reuse 'ȡ��
            Call GetTableRowName
            Call zlRefreshData
            mblnChage = False
            mblnCurveChange = False
            Call txtEdit_KeyPress(vbKeyEscape)
            Call SetColSelect
            
        Case conMenu_Edit_NewItem '��ӻ��Ŀ
            Call txtEdit_KeyPress(vbKeyEscape)
            mblnScroll = True
            If frmCaseTendBodyActiveItem.ShowMe(vsfTab, Me) Then
                vsfTab.Refresh
            End If
        Case conMenu_Edit_Curve_Show '������ʾ
            If mblnChage Then
                If MsgBox("�����Ѿ������ı�,�����Ƿ���Ҫ����?", vbInformation + vbDefaultButton1 + vbYesNo, gstrSysName) = vbYes Then
                    If Not ChangeCurveTime Then Exit Sub
                    If Not SaveData Then Exit Sub
                End If
            End If
            '������ʾ����
            Call gobjTendEditor.BodyEditCur(1, Format(mstrBegin, "YYYY-MM-DD"))
            
        Case conMenu_Edit_Append * 10, conMenu_Edit_Append * 10 + 1, conMenu_Edit_Append * 10 + 2, conMenu_Edit_Append * 10 + 3, conMenu_Edit_Append * 10 + 4, conMenu_Edit_Append * 10 + 5, conMenu_Edit_Append * 10 + 6
            If vsfTab.Row < vsfTab.FixedRows Or vsfTab.Col < vsfTab.FixedCols Then Exit Sub
            lngRow = vsfTab.Row
            lngCol = vsfTab.Col
            lngItemNO = Val(vsfTab.TextMatrix(lngRow, COL_tab��Ŀ���))
            strName = vsfTab.TextMatrix(lngRow, COL_tab��Ŀ����)
            strValue = Trim(vsfTab.TextMatrix(lngRow, lngCol))
            strTmp = vsfTab.TextMatrix(lngRow, COL_tab�ַ���)
            strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
            arrStr = Split(strTmp, ",")
            
            intType = 0
            If picEdit.Visible = True And txtEdit.Visible = True Then intType = 1
            If intType = 1 Then strValue = txtEdit.Text
            
            strPart = ""
            If InStr(1, "," & gint��� & "," & gint��Һ & ",", "," & lngItemNO & ",") = 0 Then Exit Sub
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
                    strPart = "��"
                    strValue = ""
                Case conMenu_Edit_Append * 10 + 4
                    strPart = "��"
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
                    If lngItemNO = gint��� Then
                        For i = 0 To 4
                            Select Case i
                                Case 0
                                    strPart1 = "E"
                                Case 1
                                    strPart1 = "/"
                                Case 2
                                    strPart1 = "*"
                                Case 3
                                    strPart1 = "��"
                                Case 4
                                    strPart1 = "��"
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
            
            '�Ǳ༭״̬��
            If IsWaveItem(lngItemNO) And InStr(1, Trim(vsfTab.TextMatrix(lngRow, lngCol)), "-") <> 0 Then
                strErrMsg = "������ֵ�Ѿ��γɲ�����Χ�Ĳ�����Ŀ���ܽ����޸ġ�ɾ������"
                lblStb.Caption = strErrMsg: lblStb.ForeColor = 255
                Exit Sub
            End If
            Call txtEdit_KeyPress(vbKeyEscape)
            strPart = CStr(arrStr(7))
            mrsCurve.Filter = "��Ŀ���=" & lngItemNO & " and ��Ŀ����='" & strName & "' and �к�=" & lngCol - vsfTab.FixedCols + 1
            If mrsCurve.RecordCount > 0 Then
                If mrsCurve!״̬ <> 1 And mrsCurve!״̬ <> 3 Then 'ԭ�е����� �޸ġ�ɾ�����״̬ʼ��Ϊ2
                    mrsCurve!״̬ = 2
                    mrsCurve!��ֵ = strValue
                Else '�����������ݵĴ���
                    If Trim(vsfTab.TextMatrix(lngRow, lngCol)) = "" Then
                        mrsCurve!״̬ = 3
                        mrsCurve!��ֵ = strValue
                    Else
                        mrsCurve!״̬ = 1
                        mrsCurve!��ֵ = strValue
                    End If
                End If
                mrsCurve.Update
            Else '�����ڼ�¼����������
                If Trim(strValue) <> "" Then
                    strTime = GetAnimalItemTime(lngRow, lngCol, strErrMsg)
                    If strErrMsg <> "" Then
                        lblStb.Caption = strErrMsg: lblStb.ForeColor = 255
                        Exit Sub
                    End If
                    gstrFields = "���|������|��ֵ|��λ|���|ʱ��|��Ŀ���|��Ŀ����|����|δ��˵��|������Դ|�޸�|��ʾ|��ԴID|����|״̬|�к�|��¼����"
                    gstrValues = GetMaxID & "|2)���±����Ŀ|" & strValue & "|" & strPart & "|" & _
                        0 & "|" & strTime & "|" & lngItemNO & "|" & strName & "|0||0|0|0|0|0|1|" & lngCol - vsfTab.FixedCols + 1 & "|1"
                    Call Record_Add(mrsCurve, gstrFields, gstrValues)
                End If
            End If
            vsfTab.TextMatrix(lngRow, lngCol) = strValue
            Call vsfTab_AfterRowColChange(0, 0, lngRow, lngCol)
            mblnChage = True
            
        Case conMenu_Help_Help '����
            Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_File_Exit '�˳�
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
        .Height = TextHeight("��")
        .Top = (picStb.Height - .Height) \ 2
        .Left = 10
    End With
End Sub

Private Sub cbsMain_Resize()
    On Error Resume Next
    
    Dim lngLeft As Long
    Dim lngTop As Long
    Dim lngRight As Long
    Dim lngBottom As Long  '�ͻ�����Ĵ�С

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
            If tbcThis.Selected.Tag = "���" Then
                Control.Enabled = Not mblnFileBack
            Else
                Control.Enabled = False
            End If
        Case conMenu_Edit_Append * 10 + 0, conMenu_Edit_Append
            Control.Enabled = (is������Һ(1) Or is������Һ(2)) And Not mblnFileBack And tbcThis.Selected.Tag = "���"
        Case conMenu_Edit_Append * 10 + 1, conMenu_Edit_Append * 10 + 2, conMenu_Edit_Append * 10 + 3, conMenu_Edit_Append * 10 + 4
            Control.Enabled = is������Һ(1) And Not mblnFileBack And tbcThis.Selected.Tag = "���"
        Case conMenu_Edit_Append * 10 + 5, conMenu_Edit_Append * 10 + 6
            Control.Enabled = is������Һ(2) And Not mblnFileBack And tbcThis.Selected.Tag = "���"
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
'��¼ʱ��Ϸ�ʱ�������仯��ˢ������
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
        If mbln��Ժ = False Then
            strErrMsg = "¼��������ѳ�������[����¼��������" & mintPreDays & "��]��ָ���ķ�Χ��"
        Else
            strErrMsg = "¼������ڲ��ܴ���[���˳�Ժʱ�䣺" & Format(mstrETime, "YYYY-MM-DD HH:mm:ss") & "]!"
        End If
        GoTo ErrInfo
    End If
    
    If Format(strValue, "YYYY-MM-DD") < Format(mstrBTime, "YYYY-MM-DD") Then
        strErrMsg = "¼������ڲ���С��[���µ���ʼʱ�䣺" & Format(mstrBTime, "YYYY-MM-DD HH:mm:ss") & "]��"
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
    
    If Not IsAllowInput(mT_Patient.lng����ID, mT_Patient.lng��ҳID, strTime, strCurrDate) Then
        strErrMsg = "¼���ʱ��[" & strValue & "]����[�������ݲ�¼����Чʱ��:" & mlngHours & "Сʱ]"
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
    
    '������津���� OptTime_Click �¼� Format(mstrBegin, "YYYY-MM-DD") �� �ض����
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
'����:��¼����ʱ����������÷�Χ
'------------------------------------------------------------------
    Dim strErrMsg As String
    Dim strDate As String
    Dim strCurrDate As String
    Dim strInfo As String
    
    If lngRow <> 0 Then
        strInfo = "��" & lngRow & "��"
    ElseIf strName <> "" Then
        strInfo = strInfo & "[" & strName & "]"
    Else
        strInfo = ""
    End If
    
    If Format(strTime, "YYYY-MM-DD HH:mm") > Format(mstrETime, "YYYY-MM-DD HH:mm") Then
        If mbln��Ժ = False Then
            strErrMsg = strInfo & "��¼����ʱ���ѳ�������[����¼��������" & mintPreDays & "��]��ָ���ķ�Χ!"
        Else
            strErrMsg = strInfo & "��¼����ʱ�䲻�ܴ���[���˳�Ժʱ�䣺" & Format(mstrETime, "YYYY-MM-DD HH:mm:ss") & "]!"
        End If
        GoTo ErrInfo
    End If
    
    If Format(strTime, "YYYY-MM-DD HH:mm") < Format(mstrBTime, "YYYY-MM-DD HH:mm") Then
        strErrMsg = strInfo & "��¼����ʱ�䲻��С��[���µ���ʼʱ�䣺" & Format(mstrBTime, "YYYY-MM-DD HH:mm:ss") & "]!"
        GoTo ErrInfo
    End If
    
    strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
    If Not IsAllowInput(mT_Patient.lng����ID, mT_Patient.lng��ҳID, strTime, strCurrDate) Then
        strErrMsg = strInfo & "��¼����ʱ��[" & strTime & "]����![�������ݲ�¼����Чʱ��:" & mlngHours & "Сʱ]"
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

Public Function IsAllowInput(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal strTime As String, ByVal strCurTime As String) As Boolean
    'ȡ��ָ��������ָ��ʱ��֮��ؼ����ʱ��
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo Errhand
    
    IsAllowInput = True
    gstrSQL = "" & _
              " SELECT DECODE(��ֹԭ��,1,'��Ժ',3,'ת��',10,'Ԥ��Ժ',15,'ת����',DECODE(��ʼԭ��,10,'��Ժ','δ����')) AS ����,��ֹʱ�� AS ʱ��" & _
              " From ���˱䶯��¼" & _
              " WHERE (��ֹԭ�� IN (1,3,10,15) OR ��ʼԭ��=10) And ����ID=[1] And ��ҳID=[2] And [3] <= ��ֹʱ��" & _
              " ORDER BY ��ֹʱ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��ָ��������ָ��ʱ��֮��ؼ����ʱ��", lng����ID, lng��ҳID, CDate(strTime))
    If rsTemp.RecordCount = 0 Then Exit Function
    
    'ֻȡ��һ�����ϵļ�¼
    strTime = Format(DateAdd("H", mlngHours, rsTemp!ʱ��), "yyyy-MM-dd HH:mm")
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
    If mblnFileBack = True Then lblStb.Caption = "�������������Ѿ��鵵,��������������޸�.": lblStb.ForeColor = 255
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChage = True Then
        If MsgBox("�������������Ѿ������ı�,�����Ƿ���Ҫ���棿", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Cancel = True
            Exit Sub
        End If
    End If
    
    mstrPart = ""
    mblnChage = False
    mblnMove = False
    mblnInit = False
    mblnEdit = False
    mbln��Ժ = False
    mblnAllRefresh = False
    mblnCurveChange = False
    If Not (mrsCurve Is Nothing) Then Set mrsCurve = Nothing
    If Not (mrsPart Is Nothing) Then Set mrsPart = Nothing
    If Not (mrsNote Is Nothing) Then Set mrsNote = Nothing
    If Not (mrsOper Is Nothing) Then Set mrsOper = Nothing
    If Not (mrsRecodeID Is Nothing) Then Set mrsRecodeID = Nothing
    If Not (mcbrToolBar Is Nothing) Then Set mcbrToolBar = Nothing
    '���洰��
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
        '������ݺϷ���
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
            '�ƶ�����һ��
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
        lblStb.Caption = "����Сʱֻ����0��24֮�䣬������¼�룡": lblStb.ForeColor = 255
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

Private Sub lstδ��_DblClick()
    Dim intType As Integer
    Dim blnAllow As Boolean
    Dim intCount As Integer
    Dim strδ��˵�� As String
    Dim intRows As Integer, intRow As Integer
    
    If InStr(1, picδ��.Tag, "|") <> 0 Then
        vsfCurve.Row = Split(picδ��.Tag, "|")(0)
        vsfCurve.Col = Split(picδ��.Tag, "|")(1)
    End If
    
    vsfCurve.TextMatrix(vsfCurve.Row, Col_δ��˵��) = lstδ��.Text
    strδ��˵�� = lstδ��.Text
    vsfCurve.TextMatrix(vsfCurve.Row, col_����) = Space(vsfCurve.Row) & Space(vsfCurve.Row)
    vsfCurve.TextMatrix(vsfCurve.Row, col_��ɫ) = Space(vsfCurve.Row) & IIf(vsfCurve.TextMatrix(vsfCurve.Row, COL_������) = "2)���±�˵��", " ", Space(vsfCurve.Row))
    vsfCurve.TextMatrix(vsfCurve.Row, COL_��λ) = ""
    vsfCurve.TextMatrix(vsfCurve.Row, col_����) = ""
    picδ��.Visible = False
    lstδ��.Visible = False: lstδ��.Enabled = False
    
    blnAllow = True
    intCount = 0
    intRows = 0
    If Trim(vsfCurve.TextMatrix(vsfCurve.Row, COL_������)) = "1)����������Ŀ" Then
        intType = 1
        '����������ߵ�δ������Ϊ��,ֱ�Ӹ���
        For intRow = vsfCurve.FixedRows To vsfCurve.Rows - 1
            If Trim(vsfCurve.TextMatrix(intRow, COL_������)) = "1)����������Ŀ" Then
                If vsfCurve.TextMatrix(intRow, Col_δ��˵��) = "" And Trim(vsfCurve.TextMatrix(intRow, col_����)) = "" Then
                    intCount = intCount + 1
                End If
                intRows = intRows + 1
            End If
        Next
        'ʣ�µ���Ŀ���������Ƕ�Ϊ�������
        If intCount = intRows - 1 Then
            For intRow = vsfCurve.FixedRows To vsfCurve.Rows - 1
                If Trim(vsfCurve.TextMatrix(intRow, COL_������)) = "1)����������Ŀ" And vsfCurve.TextMatrix(intRow, Col_δ��˵��) = "" Then
                    vsfCurve.TextMatrix(intRow, Col_δ��˵��) = strδ��˵��
                    vsfCurve.TextMatrix(vsfCurve.Row, col_����) = Space(vsfCurve.Row) & Space(vsfCurve.Row)
                    vsfCurve.TextMatrix(vsfCurve.Row, col_��ɫ) = Space(vsfCurve.Row) & IIf(vsfCurve.TextMatrix(vsfCurve.Row, COL_������) = "2)���±�˵��", " ", Space(vsfCurve.Row))
                    vsfCurve.TextMatrix(vsfCurve.Row, COL_��λ) = ""
                    vsfCurve.TextMatrix(vsfCurve.Row, col_����) = ""
                End If
            Next
        Else
            intCount = 0
        End If
    ElseIf Trim(vsfCurve.TextMatrix(vsfCurve.Row, COL_������)) = "2)���±�˵��" Then
        If Val(vsfCurve.TextMatrix(vsfCurve.Row, COL_��Ŀ���)) = 4 Then
            'intType = 2
            blnAllow = False
        Else
            blnAllow = False
        End If
    End If
    
    vsfCurve.Cell(flexcpAlignment, vsfCurve.FixedRows, Col_δ��˵��, vsfCurve.Rows - 1, Col_δ��˵��) = flexAlignCenterCenter
    
    If blnAllow = True Then
        If intCount = 0 Then
            Call UpdateCurveDate(vsfCurve.Row, vsfCurve.Col, intType)
        ElseIf intCount = intRows - 1 Then
            For intRow = vsfCurve.FixedRows To vsfCurve.Rows - 1
                If Trim(vsfCurve.TextMatrix(intRow, COL_������)) = "1)����������Ŀ" Then
                    Call UpdateCurveDate(intRow, Col_δ��˵��, intType)
                End If
            Next
        End If
        Call vsfCurve.SetFocus
    End If
End Sub

Private Sub lstδ��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        lstδ��.Visible = False: lstδ��.Enabled = False
        picδ��.Visible = False
    ElseIf KeyCode = vbKeyReturn Then
        Call lstδ��_DblClick
    End If
End Sub

Private Sub lstδ��_LostFocus()
    lstδ��.Visible = False: lstδ��.Enabled = False
    picδ��.Visible = False
End Sub

Private Sub OptTime_Click(Index As Integer)
    Dim strBegin As String, strEnd As String
    Dim blnTab As Boolean
    
    If Not mblnInit Then Exit Sub
    
    If mblnCurveChange = True Or (mblnAllRefresh = True And mblnChage = True) Then
        If MsgBox("�����Ѿ������ı�,�����Ƿ���б���?", vbYesNo + vbQuestion + vbDefaultButton1, gstrSysName) = vbYes Then
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
' �������
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
    Dim strOpt As String '������Ϣ
    Dim lngColor As Long, lngNO As Long
    Dim blnAllow As Boolean
    
    If Not ChangeCurveTime Then Exit Sub
    On Error GoTo Errhand
    Select Case Index
        Case 0 '��һ��
            dkpTime.Tag = 0
        Case 1 '��һ��
            dkpTime.Tag = Val(dkpTime.Tag) - 1
            If Val(dkpTime.Tag) < 0 Then dkpTime.Tag = 0
        Case 2 '��һ��
            dkpTime.Tag = Val(dkpTime.Tag) + 1
            If Val(dkpTime.Tag) > UBound(mArrdkpTime) Then dkpTime.Tag = UBound(mArrdkpTime)
        Case 3 '���һ��
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
        If Val(dkpTime.Tag) = LBound(mArrdkpTime) Then '��һ��
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
        ElseIf Val(dkpTime.Tag) = UBound(mArrdkpTime) Then '���һ��
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
        Else '�м�ĳ��
            For intIndex = 0 To picBut.Count - 1
                picBut(intIndex).Visible = True
                picBut(intIndex).Enabled = True
                picBut1(intIndex).Visible = False
                picBut1(intIndex).Enabled = False
            Next intIndex
        End If
    End If
    
   'ˢ������
    strTime = Format(mArrdkpTime(Val(dkpTime.Tag)), "YYYY-MM-DD HH:mm:ss")
    dkpTime.Value = Format(strTime, "HH:mm")
    
    '�����������������Ϣ
    For intRow = vsfCurve.FixedRows To vsfCurve.Rows - 1
        If vsfCurve.TextMatrix(intRow, COL_������) <> "" And Val(vsfCurve.TextMatrix(intRow, COL_��Ŀ���)) <> 0 Then
            For intCOl = vsfCurve.FixedCols To vsfCurve.Cols - 1
                vsfCurve.TextMatrix(intRow, intCOl) = ""
            Next intCOl
        End If
    Next intRow
    
    
    blnAllow = False
    ReDim Preserve mArrModfy(vsfCurve.FixedRows To vsfCurve.Rows - 1)
    ReDim Preserve mArrValue(vsfCurve.FixedRows To vsfCurve.Rows - 1)
    ReDim Preserve marrDate(vsfCurve.FixedRows To vsfCurve.Rows - 1)
    '��������
    vsfCurve.Cell(flexcpText, vsfCurve.FixedRows, vsfCurve.FixedCols, vsfCurve.Rows - 1, vsfCurve.Cols - 1) = ""
    vsfCurve.Cell(flexcpForeColor, vsfCurve.FixedRows, vsfCurve.FixedCols, vsfCurve.Rows - 1, vsfCurve.Cols - 1) = &H80000012
    
    For intRow = vsfCurve.FixedRows To vsfCurve.Rows - 1
        marrDate(intRow) = 0
        mArrModfy(intRow) = 0
        mArrValue(intRow) = ""

        vsfCurve.Body.MergeRow(intRow) = True
        vsfCurve.TextMatrix(intRow, col_����) = Space(intRow) & IIf(vsfCurve.TextMatrix(intRow, COL_������) = "2)���±�˵��", "", "") & Space(intRow)
        vsfCurve.TextMatrix(intRow, col_��ɫ) = Space(intRow) & IIf(vsfCurve.TextMatrix(intRow, COL_������) = "2)���±�˵��", " ", Space(intRow))
        If vsfCurve.TextMatrix(intRow, COL_������) = "2)���±�˵��" Then
             vsfCurve.Cell(flexcpBackColor, intRow, col_��ɫ, intRow, col_��ɫ) = RGB(0, 0, 255)
        End If
    Next intRow
    
    mrsCurve.Filter = "ʱ��='" & strTime & "' and ״̬<>3"
    With mrsCurve
        Do While Not .EOF
            For intRow = vsfCurve.FixedRows To vsfCurve.Rows - 1
                lngNO = Val(vsfCurve.TextMatrix(intRow, COL_��Ŀ���))
                If !������ = vsfCurve.TextMatrix(intRow, COL_������) And !��Ŀ��� = lngNO Then
                    vsfCurve.TextMatrix(intRow, col_����) = Space(intRow) & zlCommFun.Nvl(!��ֵ) & Space(intRow)
                    vsfCurve.TextMatrix(intRow, col_��ɫ) = vsfCurve.TextMatrix(intRow, col_����)
                    
                    If Not IsNumeric(zlCommFun.Nvl(!��ֵ)) And zlCommFun.Nvl(!��ֵ) <> "����" And InStr(1, zlCommFun.Nvl(!��ֵ), "/") = 0 Then
                        vsfCurve.TextMatrix(intRow, COL_��λ) = ""
                        vsfCurve.TextMatrix(intRow, Col_δ��˵��) = zlCommFun.Nvl(!δ��˵��)
                    Else
                        vsfCurve.TextMatrix(intRow, COL_��λ) = zlCommFun.Nvl(!��λ)
                        vsfCurve.TextMatrix(intRow, Col_δ��˵��) = ""
                    End If
                    If lngNO = 1 And (IsNumeric(zlCommFun.Nvl(!��ֵ)) Or zlCommFun.Nvl(!��ֵ) <> "����") Then
                        vsfCurve.TextMatrix(intRow, col_����) = IIf(Val(zlCommFun.Nvl(!����)) = 1, "��", "")
                    End If
                    lngColor = 255
                    If Val(zlCommFun.Nvl(!������Դ)) <> 0 Then
                        If zlCommFun.Nvl(!��ֵ) = "����" And lngNO = 1 Then
                            lngColor = 255
                        ElseIf lngNO = 1 Or (lngNO = 2 And mint����Ӧ�� = 2 And mblnEdit���� = True) Then
                            If InStr(1, zlCommFun.Nvl(!��ֵ), "/") = 0 Then
                                lngColor = RGB(0, 0, 255)
                            Else
                                If Val(!�޸�) = 0 Then
                                    lngColor = RGB(0, 0, 255)
                                Else
                                    lngColor = 255
                                End If
                            End If
                        End If
                        vsfCurve.Cell(flexcpForeColor, intRow, col_����, intRow, col_����) = lngColor
                    Else
                        vsfCurve.Cell(flexcpForeColor, intRow, col_����, intRow, col_����) = &H80000012
                    End If
                    marrDate(intRow) = Val(CStr(zlCommFun.Nvl(!������Դ)))
                    If InStr(1, ",0,9,", Val(zlCommFun.Nvl(!������Դ))) = 0 Then
                        blnAllow = True
                    End If
                    mArrModfy(intRow) = Val(!�޸�)
                    mArrValue(intRow) = Val(!��ֵ)
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
    
    '���±�(����ʼ�ձ��ֲ���)
    mrsNote.Filter = 0
    With mrsNote
        Do While Not .EOF
            Select Case Val(!��¼����)
                Case 2
                    intRow = mOptRow.�ϱ�
                Case 6
                    intRow = mOptRow.�±�
            End Select
            vsfCurve.TextMatrix(intRow, col_����) = Space(intRow) & zlCommFun.Nvl(!����) & Space(intRow)
            vsfCurve.Cell(flexcpBackColor, intRow, col_��ɫ, intRow, col_��ɫ) = IIf(IsNumeric(Nvl(!δ��˵��)) = False, 16711680, Val(Nvl(!δ��˵��)))
            vsfCurve.TextMatrix(intRow, COL_��λ) = ""
            vsfCurve.TextMatrix(intRow, col_����) = ""
            vsfCurve.TextMatrix(intRow, Col_δ��˵��) = ""
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
            strInfo = "��һ��"
        Case 1
            strInfo = "��һ��"
        Case 2
            strInfo = "��һ��"
        Case 3
            strInfo = "���һ��"
    End Select
    
    picBut(Index).ToolTipText = strInfo
End Sub

Private Sub picBut1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strInfo As String
    Select Case Index
        Case 0
            strInfo = "��һ��"
        Case 1
            strInfo = "��һ��"
        Case 2
            strInfo = "��һ��"
        Case 3
            strInfo = "���һ��"
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
    
    With picδ��
        .Width = 1080 + 1080 * mintBigSize / 3
        .Height = 1100 + 1100 * mintBigSize / 3
        .Visible = False
        .Font.Size = mFontSize + mFontSize * mintBigSize / 3
    End With
    
    With lstδ��
        .Top = 0
        .Left = 0
        .Width = picδ��.Width
        .Height = picδ��.Height
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
    Dim strTmpCurve As String, strTmpTable As String '���ߺͱ����Ŀ����
    Dim strCollectItem As String '��������Ŀ
    Dim arrActive() As String
    Dim strֵ�� As String
    Dim strSQL As String
    Dim i As Integer, intBound As Integer
    Dim strEndTime As String
    Dim Titem As Type_Item
    Dim strDate As String
    Dim intCOl As Integer
    Dim strCurDate As String
    
    On Error GoTo Errhand
    
    Call InitRecordSet
    
    '����������ʹ���ʱ�����Ƿ�ʹ����˲���
    mstrSQL = "select C.Ӧ�÷�ʽ From �����¼��Ŀ C where C.��Ŀ���=[1] And C.����ȼ�>=[2] And Nvl(C.���ò���,0) In (0,[3]) " & _
            " And (c.���ÿ���=1 Or (c.���ÿ���=2 And Exists (Select 1 From �������ÿ��� D Where D.��Ŀ���=c.��Ŀ��� And D.����id=[4])))"
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "��ȡ������", -1, mT_Patient.lng����ȼ�, IIf(mT_Patient.lngӤ�� = 0, 1, 2), mT_Patient.lng����ID)
    mblnEdit���� = IIf(rsTemp.RecordCount = 0, False, True)
    If rsTemp.RecordCount > 0 Then mint����Ӧ�� = Val(zlCommFun.Nvl(rsTemp!Ӧ�÷�ʽ, 0))
    
    '��ʽ���Ϊ ����'ֵ��,��Ŀ����,��ĿС��,��¼Ƶ��,��Ŀ��ʾ,��Ŀ����,��Ŀ����,��λ,��Ժ�ײ�'��Ŀ��'��Ŀ��
    strTmp = "2)���±�˵��',,,,,,,,'2'�ϱ�;2)���±�˵��',,,,,,,,'6'�±�"
    
    '��ȡ��������������Ŀ
    mstrSQL = " Select A.�������,A.��¼�� ��Ŀ��,A.��Ŀ��� as ��Ŀ��,A.��¼��,A.��Ժ�ײ�," & _
            " C.��Ŀֵ��,C.��Ŀ����,C.��Ŀ����,C.��ĿС��,nvl(A.��¼Ƶ��,2) ��¼Ƶ��,C.������,C.��Ŀ��ʾ,C.��Ŀ��λ " & _
            " From ���¼�¼��Ŀ A,����������Ŀ B,�����¼��Ŀ C " & _
            " Where c.��ĿID=B.ID(+) And A.��Ŀ���=C.��Ŀ��� And (A.��¼��=1 OR (A.��¼��=2 And A.��Ŀ���=3)) And ��Ŀ����=1 And Nvl(C.Ӧ�÷�ʽ,0)=1 AND C.����ȼ�>=[1] And Nvl(C.���ò���,0) In (0,[3]) " & _
            " And (c.���ÿ���=1 Or (c.���ÿ���=2 And Exists (Select 1 From �������ÿ��� D Where D.��Ŀ���=c.��Ŀ��� And D.����id=[2])))" & _
            " Order by Decode(A.��Ŀ���,1,0,1),A.�������"
            
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, mT_Patient.lng����ȼ�, mT_Patient.lng����ID, IIf(mT_Patient.lngӤ�� = 0, 1, 2))
    
    With rsTemp
        Do While Not .EOF
            strֵ�� = Replace(zlCommFun.Nvl(!��Ŀֵ��), ":", "")
            If zlCommFun.Nvl(!��Ŀ����) = 0 Then
                If InStr(1, strֵ��, ";") <> 0 Then strֵ�� = Split(strֵ��, ";")(0) & "��" & Split(strֵ��, ";")(1)
            End If
            strֵ�� = Replace(Replace(Replace(strֵ��, ";", ":"), "'", ""), ",", "")
            
            Titem.ֵ�� = strֵ��
            Titem.��Ŀ���� = Val(zlCommFun.Nvl(!��Ŀ����, 0))
            Titem.��ĿС�� = Val(zlCommFun.Nvl(!��ĿС��, 0))
            Titem.��¼Ƶ�� = Val(zlCommFun.Nvl(!��¼Ƶ��, 2))
            Titem.��Ŀ��ʾ = Val(zlCommFun.Nvl(!��Ŀ��ʾ, 0))
            Titem.��Ŀ���� = 1
            Titem.��Ŀ���� = zlCommFun.Nvl(!��Ŀ����, 3)
            Titem.��λ = ""
            Titem.��Ŀ�� = Val(zlCommFun.Nvl(!��Ŀ��))
            Titem.��Ŀ�� = Replace(Replace(zlCommFun.Nvl(!��Ŀ��) & IIf(zlCommFun.Nvl(!��Ŀ��λ, "") = "", "", "(" & !��Ŀ��λ & ")"), ";", ":"), "'", "")
            Titem.��Ժ�ײ� = Val(zlCommFun.Nvl(!��Ժ�ײ�, 0))
            
            If Titem.��Ŀ��ʾ = 4 Or IsWaveItem(Titem.��Ŀ��) Then
                If Titem.��¼Ƶ�� > 2 Then Titem.��¼Ƶ�� = 2
            End If
            '��¼��=1���¼��=2�ĺ����Ϊ������Ŀ
            Titem.���� = "1)����������Ŀ"
            strTmpCurve = strTmpCurve & ";" & Titem.���� & "'" & Titem.ֵ�� & "," & Titem.��Ŀ���� & "," & _
                Titem.��ĿС�� & "," & Titem.��¼Ƶ�� & "," & Titem.��Ŀ��ʾ & ",1," & Titem.��Ŀ���� & ",," & Titem.��Ժ�ײ� & "'" & _
                Titem.��Ŀ�� & "'" & Titem.��Ŀ��
        .MoveNext
        Loop
    End With
    
    mstrActiveItem = ""
    
    strEndTime = DateAdd("d", 6, CDate(Format(Format(mstrDate, "YYYY-MM-DD") & " 23:59:59", "YYYY-MM-DD HH:mm:ss")))
    If strEndTime > mstrETime Then strEndTime = mstrETime
    '��ȡ�̶������Ŀ������ֵ�Ļ��Ŀ��Ϣ
    Set rsTemp = GetAppendGridItem(mT_Patient.lng�ļ�ID, mT_Patient.lng����ID, mT_Patient.lng��ҳID, mT_Patient.lng����ȼ�, mT_Patient.lngӤ��, _
        CDate(mstrDate), CDate(strEndTime), IIf(mT_Patient.lngӤ�� = 0, 1, 2), mT_Patient.lng����ID, mblnMove)
    With rsTemp
        Do While Not .EOF
           strֵ�� = Replace(zlCommFun.Nvl(!��Ŀֵ��), ":", "")
            If zlCommFun.Nvl(!��Ŀ����) = 0 Then
                If InStr(1, strֵ��, ";") <> 0 Then strֵ�� = Split(strֵ��, ";")(0) & "��" & Split(strֵ��, ";")(1)
            End If
            strֵ�� = Replace(Replace(Replace(strֵ��, ";", ":"), "'", ""), ",", "")
            
            Titem.ֵ�� = strֵ��
            Titem.���� = "2)���±����Ŀ"
            Titem.��Ŀ���� = Val(zlCommFun.Nvl(!��Ŀ����))
            Titem.��ĿС�� = Val(zlCommFun.Nvl(!��ĿС��, 0))
            Titem.��¼Ƶ�� = Val(zlCommFun.Nvl(!��¼Ƶ��, 2))
            Titem.��Ŀ��ʾ = Val(zlCommFun.Nvl(!��Ŀ��ʾ, 0))
            Titem.��Ŀ���� = Val(zlCommFun.Nvl(!��Ŀ����, 1))
            Titem.��Ŀ���� = zlCommFun.Nvl(!��Ŀ����, 3)
            Titem.��λ = Replace(Replace(Replace(zlCommFun.Nvl(!���²�λ), ";", ""), "'", ""), ",", "")
            Titem.��Ŀ�� = Val(zlCommFun.Nvl(!��Ŀ���))
            Titem.��Ŀ�� = Replace(Replace(IIf(Titem.��Ŀ�� = 4, "Ѫѹ", zlCommFun.Nvl(!��¼��)) & IIf(zlCommFun.Nvl(!��λ, "") = "", "", "(" & !��λ & ")"), ";", ":"), "'", "")
            Titem.��Ժ�ײ� = Val(zlCommFun.Nvl(!��Ժ�ײ�, 0))
            If Titem.��Ŀ��ʾ = 4 Or IsWaveItem(Titem.��Ŀ��) Then
                If Titem.��¼Ƶ�� > 2 Then Titem.��¼Ƶ�� = 2
            End If
            If Titem.��Ŀ�� <> gint���� And Titem.��Ŀ�� <> 5 Then
                strTmpTable = strTmpTable & ";" & Titem.���� & "'" & Titem.ֵ�� & "," & Titem.��Ŀ���� & "," & _
                    Titem.��ĿС�� & "," & Titem.��¼Ƶ�� & "," & Titem.��Ŀ��ʾ & "," & Titem.��Ŀ���� & "," & Titem.��Ŀ���� & "," & _
                    Titem.��λ & "," & Titem.��Ժ�ײ� & "'" & Titem.��Ŀ�� & "'" & Titem.��Ŀ��
                '��¼�Ѿ����ڵĻ��Ŀ��Ϣ
                If Titem.��Ŀ���� = 2 Then
                    mstrActiveItem = mstrActiveItem & ";" & Titem.���� & "'" & Titem.ֵ�� & "," & Titem.��Ŀ���� & "," & _
                        Titem.��ĿС�� & "," & Titem.��¼Ƶ�� & "," & Titem.��Ŀ��ʾ & "," & Titem.��Ŀ���� & "," & Titem.��Ŀ���� & "," & _
                        Titem.��λ & "," & Titem.��Ժ�ײ� & "'" & Titem.��Ŀ�� & "'" & Titem.��Ŀ��
                End If
            End If
        .MoveNext
        Loop
    End With
    
    If Left(mstrActiveItem, 1) = ";" Then mstrActiveItem = Mid(mstrActiveItem, 2)
        
    If strTmp <> "" Then strTmpCurve = strTmpCurve & ";" & strTmp
    If Left(strTmpCurve, 1) = ";" Then strTmpCurve = Mid(strTmpCurve, 2)
    If Left(strTmpTable, 1) = ";" Then strTmpTable = Mid(strTmpTable, 2)
    
    '���������������ݰ����������±�
    Call InitTabCurve(strTmpCurve)
    
    '�������±������(������������Ŀ)
    Call InitTabTable(strTmpTable)
    
    '��������¼����
    Call InitTabOper
    
    mstrδ��˵�� = ""
    '��ȡδ��˵����Ϣ
    mstrSQL = "Select ����,���� From ��������˵��"
    Call zlDatabase.OpenRecordset(rsTemp, mstrSQL, Me.Caption)
    With rsTemp
        Do While Not .EOF
            mstrδ��˵�� = mstrδ��˵�� & "," & zlCommFun.Nvl(!����)
        .MoveNext
        Loop
    End With
    
    If Left(mstrδ��˵��, 1) = "," Then mstrδ��˵�� = Mid(mstrδ��˵��, 2)
    
    '����ѡ��ʱ�䶨λ�ڵ�ǰʱ��༭״̬
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
'����:��ȡһ��ʱ���ڵ�������������
'���� blnCurve�Ƿ�ˢ���������� blnTab �Ƿ�ˢ�±������
'-----------------------------------------------------------------------------------------------------------------
    Dim rsTmp As New ADODB.Recordset
    Dim rsDownTab As New ADODB.Recordset
    Dim dtBegin As Date, dtEnd As Date
    Dim lng��Ŀ��� As Long, int��Ŀ���� As Integer, str��Ŀ���� As String, int��¼Ƶ�� As Integer, int��Ŀ��ʾ As Integer, int��Ժ�ײ� As Integer
    Dim intRow As Integer, intCOl As Integer, intNum As Integer, strName As String
    Dim strParam As String, strFidlds As String, strPart As String, strTmp As String
    Dim blnAllow As Boolean, blnAdd As Boolean
    Dim strTime As String
    Dim rsCurve As New ADODB.Recordset '��ʱ��¼��
    Dim intModify As Integer, int������Դ As Integer
    Dim lngColor As Long
    Dim i As Integer, int��� As Integer
    Dim strOperTime As String, strOper As String
    Dim strItems As String, strItemName As String

    On Error GoTo Errhand
    
    If blnCurve = False And blnTab = False Then Exit Function
    
    lblTime.Caption = Format(mstrBegin, "HH:mm") & "��" & Format(mstrEnd, "HH:mm")
    dkpTime.MaxDate = Format(mstrEnd, "HH:mm")
    dkpTime.MinDate = Format(mstrBegin, "HH:mm")
    mArrdkpTime = Array()
        
    '��ʼ����¼��
    gstrFields = "��¼ID," & adDouble & ",18|ʱ��," & adLongVarChar & ",20"
    Call Record_Init(mrsRecodeID, gstrFields)
    
    '�޸� ��ʾ����ͬ�����������ݣ��������û��������,�����������޶��� ����Խ��������ºͶ������ݵ��޸�  0 �����޸� 1�����޸�
    gstrFields = "���," & adDouble & ",18|������," & adLongVarChar & ",40|��ֵ," & adLongVarChar & ",400|��λ," & adLongVarChar & ",200|" & _
         "���," & adDouble & ",1|ʱ��," & adLongVarChar & ",20|ԭʼʱ��," & adLongVarChar & ",20|��Ŀ���," & adDouble & ",18|��Ŀ����," & adLongVarChar & ",40|" & _
         "����," & adDouble & ",1|δ��˵��," & adLongVarChar & ",20|������Դ," & adDouble & ",1|�޸�," & adDouble & ",1|��ʾ," & adDouble & ",1|" & _
         "��ԴID," & adDouble & ",18|����," & adDouble & ",1|״̬," & adDouble & ",1|�к�," & adDouble & ",1|��¼����," & adDouble & ",1"
    Call Record_Init(rsCurve, gstrFields)
         
    If blnCurve = True And blnTab = False Then 'ֻˢ����������
        If Not mrsCurve Is Nothing And mrsCurve.State = 1 Then
            mrsCurve.Filter = 0
            mrsCurve.Filter = "������='2)���±����Ŀ'"
            Do While Not mrsCurve.EOF
                rsCurve.AddNew
                For i = 0 To mrsCurve.Fields.Count - 1
                    rsCurve.Fields(mrsCurve.Fields(i).Name).Value = mrsCurve.Fields(i).Value
                Next i
                rsCurve.Update
            mrsCurve.MoveNext
            Loop
        End If
    ElseIf blnCurve = False And blnTab = True Then 'ֻˢ�±��
        If Not mrsCurve Is Nothing And mrsCurve.State = 1 Then
            mrsCurve.Filter = 0
            mrsCurve.Filter = "������='1)����������Ŀ'"
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
    
    gstrFields = "���|������|��ֵ|��λ|���|ʱ��|ԭʼʱ��|��Ŀ���|��Ŀ����|����|δ��˵��|������Դ|�޸�|��ʾ|��ԴID|����|״̬|�к�|��¼����"
    
    'ˢ���������������Լ��������±�
    If blnCurve = True Then
        '1------------------------------------------------------------
        '��ȡĳʱ��ε�����������������
        mstrSQL = _
        " SELECT C.ID ���,C.��¼ID,A.����ʱ�� As ʱ��,'1)����������Ŀ' ������,C.��ʾ,c.��¼���� As ��ֵ,c.���²�λ,c.���Ժϸ�,D.��¼��,D.��Ŀ���,DECODE(D.��Ŀ���,-1,1,C.��¼���) ��¼���,C.δ��˵��,C.������Դ,C.��ԴID,C.����" & vbNewLine & _
        "                    FROM ���˻����ļ� B,���˻������� A,���˻�����ϸ C,���¼�¼��Ŀ D,�����¼��Ŀ E" & vbNewLine & _
        "                    Where B.ID=A.�ļ�ID" & vbNewLine & _
        "                        AND A.ID = C.��¼ID" & vbNewLine & _
        "                        AND B.ID=[1]" & vbNewLine & _
        "                        AND Nvl(B.Ӥ��,0)=[4]" & vbNewLine & _
        "                        AND B.����id=[2]" & vbNewLine & _
        "                        AND B.��ҳid=[3]" & vbNewLine & _
        "                        AND D.��Ŀ���=C.��Ŀ���" & vbNewLine & _
        "                        AND C.��¼����=1" & vbNewLine & _
        "                        AND E.��Ŀ���=D.��Ŀ���" & vbNewLine & _
        "                        AND E.����ȼ�>=[7]" & vbNewLine & _
        "                        AND (nvl(D.��¼��,1)=1 or (NVL(D.��¼��,1)=2 And D.��Ŀ���=3))" & _
        "                        And A.����ʱ�� BETWEEN [5] And [6] And C.��ֹ�汾 Is Null" & vbNewLine & _
        "                        AND (nvl(E.Ӧ�÷�ʽ,0)=1 OR ( -1=[10] and nvl(E.Ӧ�÷�ʽ,0)=2))" & vbNewLine & _
        "                        AND nvl(E.���ò���,0) in (0,[8]) AND (E.���ÿ���=1 or ( E.���ÿ���=2 AND Exists (select 1 from �������ÿ��� D where D.��Ŀ���=E.��Ŀ��� and D.����ID=[9])))" & vbNewLine & _
        "                    Order By A.����ʱ��,DECODE(D.��Ŀ���,-1,1,0),DECODE(D.��Ŀ���,-1,1,C.��¼���)"
    
        If mblnMove Then
            mstrSQL = Replace(mstrSQL, "���˻����ļ�", "H���˻����ļ�")
            mstrSQL = Replace(mstrSQL, "���˻�������", "H���˻�������")
            mstrSQL = Replace(mstrSQL, "���˻�����ϸ", "H���˻�����ϸ")
        End If
        
        Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, mT_Patient.lng�ļ�ID, mT_Patient.lng����ID, mT_Patient.lng��ҳID, mT_Patient.lngӤ��, _
             CDate(Format(mstrBegin, "YYYY-MM-DD HH:mm:ss")), CDate(Format(mstrEnd, "YYYY-MM-DD HH:mm:ss")), mT_Patient.lng����ȼ�, IIf(mT_Patient.lngӤ�� = 0, 1, 2), mT_Patient.lng����ID, IIf(mint����Ӧ�� = 2, -1, 0))
        With rsTmp
            
            Do While Not .EOF
                
                '��Ӽ�¼��
                Call Record_Update(mrsRecodeID, "��¼ID|ʱ��", Val(Nvl(!��¼ID)) & "|" & Format(zlCommFun.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss"), "��¼ID|" & Val(Nvl(!��¼ID)))
                
                intModify = 0
                If strTime = "" Then strTime = Format(zlCommFun.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss")
                lng��Ŀ��� = zlCommFun.Nvl(!��Ŀ���)
                Select Case lng��Ŀ���
                    Case gint����
                        int��� = 1
                    Case Else
                        int��� = Val(Nvl(!��¼���))
                End Select
                intModify = IIf(InStr(1, ",0,9,", "," & Val(zlCommFun.Nvl(!������Դ)) & ",") = 0, 1, 0)
                blnAdd = True
                '���ʺ���������ʱ�����������Ӧ��ʱ���Ƿ��������
                If mint����Ӧ�� = 2 And lng��Ŀ��� = -1 Then
                    mrsCurve.Filter = "��Ŀ���=2 and ʱ��='" & Format(zlCommFun.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss") & "'"
                    If mrsCurve.RecordCount > 0 Then
                        strParam = "���|" & mrsCurve("���")
                        strFidlds = "��ֵ|���|�޸�"
                        
                        If InStr(1, ",0,9,", "," & Val(zlCommFun.Nvl(mrsCurve!������Դ)) & ",") = 0 And InStr(1, ",0,9,", "," & Val(zlCommFun.Nvl(!������Դ)) & ",") = 0 Then
                            intModify = 1
                        Else
                            intModify = 0
                        End If
                        
                        '��������ʱ����δδ��˵��ֻ��ʾ����������Ϊδ��˵��ʱ����ʾδ��˵��
                        If UBound(Split(mrsCurve("��ֵ"), "/")) <> -1 Then
                            If IsNumeric(zlCommFun.Nvl(!��ֵ)) Then
                                If mbln����������ʾ Then
                                    gstrValues = zlCommFun.Nvl(!��ֵ) & "/" & Split(mrsCurve("��ֵ"), "/")(0) & "|" & int��� & "|" & intModify
                                Else
                                    gstrValues = Split(mrsCurve("��ֵ"), "/")(0) & "/" & zlCommFun.Nvl(!��ֵ) & "|" & int��� & "|" & intModify
                                End If
                            Else
                                gstrValues = Split(mrsCurve("��ֵ"), "/")(0) & "|" & int��� & "|0"
                            End If
                        Else
                            gstrValues = mrsCurve("��ֵ") & "|1|0"
                        End If
                        
                        Call Record_Update(mrsCurve, strFidlds, gstrValues, strParam)
                        blnAdd = False
                    Else
                        lng��Ŀ��� = 2
                    End If
                End If
                
                '����������
                If lng��Ŀ��� = 1 And int��� = 1 Then
                    mrsCurve.Filter = "��Ŀ���=1 and ʱ��='" & Format(zlCommFun.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss") & "' and ���<>1"
                    If mrsCurve.RecordCount > 0 Then
                        If InStr(1, ",0,9,", "," & Val(zlCommFun.Nvl(mrsCurve!������Դ)) & ",") = 0 And InStr(1, ",0,9,", "," & Val(zlCommFun.Nvl(!������Դ)) & ",") = 0 Then
                            intModify = 1
                        Else
                            intModify = 0
                        End If
                        
                        strParam = "���|" & mrsCurve("���")
                        strFidlds = "��ֵ|���|�޸�"
                        gstrValues = Split(mrsCurve("��ֵ"), "/")(0) & "/" & zlCommFun.Nvl(!��ֵ) & "|" & int��� & "|" & intModify
                        Call Record_Update(mrsCurve, strFidlds, gstrValues, strParam)
                    End If
                    blnAdd = False
                End If
                
                If blnAdd Then
                    '����������ʾ����
                    strPart = GetPart(lng��Ŀ���)
                    int������Դ = Val(zlCommFun.Nvl(!������Դ, 0))
                    If Trim(Replace(zlCommFun.Nvl(!��ֵ), "/", "")) = "" Then
                        int������Դ = 0
                    End If
                    gstrValues = zlCommFun.Nvl(!���) & "|" & zlCommFun.Nvl(!������) & "|" & Trim(Replace(zlCommFun.Nvl(!��ֵ), "/", "")) & "|" & _
                        zlCommFun.Nvl(!���²�λ, strPart) & "|" & int��� & "|" & Format(zlCommFun.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss") & "|" & _
                        Format(zlCommFun.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss") & "|" & lng��Ŀ��� & "|" & zlCommFun.Nvl(!��¼��) & "|" & Val(zlCommFun.Nvl(!���Ժϸ�, 0)) & "|" & _
                        zlCommFun.Nvl(!δ��˵��) & "|" & int������Դ & "|" & intModify & "|" & Val(zlCommFun.Nvl(!��ʾ, 0)) & "|" & Val(zlCommFun.Nvl(!��ԴID, 0)) & "|" & Val(zlCommFun.Nvl(!����, 0)) & "|0|0|1"
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
        '��ʾ��������
        mrsCurve.Filter = 0
        mrsCurve.Sort = "ʱ��"
        
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
            vsfCurve.TextMatrix(intRow, col_����) = Space(intRow) & IIf(vsfCurve.TextMatrix(intRow, COL_������) = "2)���±�˵��", "", "") & Space(intRow)
            vsfCurve.TextMatrix(intRow, col_��ɫ) = Space(intRow) & IIf(vsfCurve.TextMatrix(intRow, COL_������) = "2)���±�˵��", " ", Space(intRow))
            If vsfCurve.TextMatrix(intRow, COL_������) = "2)���±�˵��" Then
                 vsfCurve.Cell(flexcpBackColor, intRow, col_��ɫ, intRow, col_��ɫ) = RGB(0, 0, 255)
            End If
        Next intRow
        
        With mrsCurve
            Do While Not .EOF
                If Format(strTime, "YYYY-MM-DD HH:mm:ss") = Format(!ʱ��, "YYYY-MM-DD HH:mm:ss") Then
                    For intRow = vsfCurve.FixedRows To vsfCurve.Rows - 1
                        lng��Ŀ��� = Val(vsfCurve.TextMatrix(intRow, COL_��Ŀ���))
                        If !������ = vsfCurve.TextMatrix(intRow, COL_������) And !��Ŀ��� = lng��Ŀ��� Then
                            vsfCurve.TextMatrix(intRow, col_����) = Space(intRow) & zlCommFun.Nvl(!��ֵ) & Space(intRow)
                            vsfCurve.TextMatrix(intRow, col_��ɫ) = vsfCurve.TextMatrix(intRow, col_����)
                            If Not IsNumeric(zlCommFun.Nvl(!��ֵ)) And zlCommFun.Nvl(!��ֵ) <> "����" And InStr(1, zlCommFun.Nvl(!��ֵ), "/") = 0 Then
                                vsfCurve.TextMatrix(intRow, COL_��λ) = ""
                                vsfCurve.TextMatrix(intRow, Col_δ��˵��) = zlCommFun.Nvl(!δ��˵��)
                            Else
                                vsfCurve.TextMatrix(intRow, COL_��λ) = zlCommFun.Nvl(!��λ)
                                vsfCurve.TextMatrix(intRow, Col_δ��˵��) = ""
                            End If
                            If lng��Ŀ��� = 1 And (IsNumeric(zlCommFun.Nvl(!��ֵ)) Or zlCommFun.Nvl(!��ֵ) <> "����") Then
                                vsfCurve.TextMatrix(intRow, col_����) = IIf(Val(zlCommFun.Nvl(!����)) = 1, "��", "")
                            End If
                            lngColor = 255
                            If InStr(1, ",0,9,", Val(zlCommFun.Nvl(!������Դ))) = 0 Then
                                If zlCommFun.Nvl(!��ֵ) = "����" And lng��Ŀ��� = 1 Then
                                    lngColor = 255
                                ElseIf lng��Ŀ��� = 1 Or (lng��Ŀ��� = 2 And mint����Ӧ�� = 2 And mblnEdit���� = True) Then
                                    If InStr(1, zlCommFun.Nvl(!��ֵ), "/") = 0 Then
                                        lngColor = RGB(0, 0, 255)
                                    Else
                                        If Val(!�޸�) = 0 Then
                                            lngColor = RGB(0, 0, 255)
                                        Else
                                            lngColor = 255
                                        End If
                                    End If
                                End If
                                vsfCurve.Cell(flexcpForeColor, intRow, col_����, intRow, col_����) = lngColor
                            Else
                                vsfCurve.Cell(flexcpForeColor, intRow, col_����, intRow, col_����) = &H80000012
                            End If
                            marrDate(intRow) = Val(CStr(zlCommFun.Nvl(!������Դ)))
                            If InStr(1, ",0,9,", Val(zlCommFun.Nvl(!������Դ))) = 0 Then
                                blnAllow = True
                            End If
                            mArrModfy(intRow) = Val(!�޸�)
                            mArrValue(intRow) = Val(!��ֵ)
                            If mbln����������ʾ And InStr(!��ֵ, "/") > 0 Then
                                mArrValue(intRow) = Split(!��ֵ, "/")(1)
                            End If
                            
                        End If
                    Next intRow
                End If
                
                '��֯ʱ���ַ���,�����жϱ���ʱ�����ж��ٸ�ʱ���������
                If CDate(Format(strTmp, "YYYY-MM-DD HH:mm:ss")) <> CDate(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")) Then
                    strTmp = Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")
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
        
        '����ͬ������������ ʱ�䲻�����޸�
        If blnAllow = True Or mblnFileBack = True Then
            dkpTime.Enabled = False
        Else
            dkpTime.Enabled = True
        End If
        
        '2----------------------------------------------------------------------------
        '��ȡ���������±�˵����Ϣ
        
        gstrFields = "���," & adDouble & ",18|��Ŀ���," & adDouble & ",18|ʱ��," & adLongVarChar & ",20|ԭʼʱ��," & adLongVarChar & ",20|��¼����," & adDouble & ",1|����," & _
            adLongVarChar & ",100|��Ŀ����," & adLongVarChar & ",20|δ��˵��," & adLongVarChar & ",20|��¼���," & adDouble & ",1|������Դ," & adDouble & ",1|��ʾ," & adDouble & ",1|" & _
             "��ԴID," & adDouble & ",18|����," & adDouble & ",1|״̬," & adDouble & ",1"
        Call Record_Init(mrsNote, gstrFields)
        gstrFields = "���|��Ŀ���|ʱ��|ԭʼʱ��|��¼����|����|��Ŀ����|δ��˵��|��¼���|������Դ|��ʾ|��ԴID|����|״̬"
        
        mstrSQL = "" & _
             " Select C.ID ���, B.����ʱ�� AS ʱ��,C.��¼����,C.��Ŀ���,C.δ��˵��,C.��¼����,C.��¼���,C.��Ŀ����,C.������Դ,C.��ʾ,C.��ԴID,C.����" & _
             " FROM ���˻����ļ� A, ���˻������� B, ���˻�����ϸ C" & _
             " Where A.ID=B.�ļ�ID and  B.ID = C.��¼ID AND A.ID=[1]  AND Nvl(A.Ӥ��, 0)=[4] AND a.����id=[2] AND a.��ҳid=[3] And c.��ֹ�汾 Is Null" & _
             " AND c.��¼���� in (2,6)  AND B.����ʱ�� BETWEEN [5]  And [6]"
             
        If mblnMove Then
            mstrSQL = Replace(mstrSQL, "���˻����ļ�", "H���˻����ļ�")
            mstrSQL = Replace(mstrSQL, "���˻�������", "H���˻�������")
            mstrSQL = Replace(mstrSQL, "���˻�����ϸ", "H���˻�����ϸ")
        End If
        
        Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, "��ȡ���������±����Ϣ", mT_Patient.lng�ļ�ID, mT_Patient.lng����ID, mT_Patient.lng��ҳID, _
            mT_Patient.lngӤ��, Int(CDate(mstrBegin)), CDate(mstrEnd))
        With rsTmp
            Do While Not .EOF
                gstrValues = zlCommFun.Nvl(!���) & "|" & zlCommFun.Nvl(!��Ŀ���, 0) & "|" & Format(zlCommFun.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss") & "|" & Format(zlCommFun.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss") & "|" & zlCommFun.Nvl(!��¼����) & "|" & _
                    zlCommFun.Nvl(!��¼����) & "|" & zlCommFun.Nvl(!��Ŀ����) & "|" & Nvl(!δ��˵��) & "|" & zlCommFun.Nvl(!��¼���, 0) & "|" & Val(zlCommFun.Nvl(!������Դ, 0)) & "|" & _
                    Val(zlCommFun.Nvl(!��ʾ, 0)) & "|" & Val(zlCommFun.Nvl(!��ԴID, 0)) & "|" & Val(zlCommFun.Nvl(!����, 0)) & "|0"
                Call Record_Add(mrsNote, gstrFields, gstrValues)
            .MoveNext
            Loop
        End With
        
        '������±���Ϣ
        mrsNote.Filter = 0
        With mrsNote
            Do While Not .EOF
                    If CDate(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")) >= CDate(Format(mstrBegin, "YYYY-MM-DD HH:mm:ss")) _
                        And CDate(Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")) <= CDate(Format(mstrEnd, "YYYY-MM-DD HH:mm:ss")) Then
                        Select Case Val(!��¼����)
                            Case 2
                                intRow = mOptRow.�ϱ�
                            Case 6
                                intRow = mOptRow.�±�
                        End Select
                        vsfCurve.TextMatrix(intRow, col_����) = Space(intRow) & zlCommFun.Nvl(!����) & Space(intRow)
                        vsfCurve.Cell(flexcpBackColor, intRow, col_��ɫ, intRow, col_��ɫ) = IIf(IsNumeric(Nvl(!δ��˵��)) = False, 16711680, Val(Nvl(!δ��˵��)))
                        vsfCurve.TextMatrix(intRow, COL_��λ) = ""
                        vsfCurve.TextMatrix(intRow, col_����) = ""
                        vsfCurve.TextMatrix(intRow, Col_δ��˵��) = ""
                    End If
            .MoveNext
            Loop
        End With
    End If
    
    'ˢ�±�����ݺ�������Ϣ
    If blnTab = True Then
        gstrFields = "���," & adDouble & ",18|��Ŀ���," & adDouble & ",18|ʱ��," & adLongVarChar & ",20|ԭʼʱ��," & adLongVarChar & ",20|��¼����," & adDouble & ",1|����," & _
            adLongVarChar & ",100|��Ŀ����," & adLongVarChar & ",20|δ��˵��," & adLongVarChar & ",20|��¼���," & adDouble & ",1|������Դ," & adDouble & ",1|��ʾ," & adDouble & ",1|" & _
             "��ԴID," & adDouble & ",18|����," & adDouble & ",1|״̬," & adDouble & ",1"
        Call Record_Init(mrsOper, gstrFields)
        gstrFields = "���|��Ŀ���|ʱ��|ԭʼʱ��|��¼����|����|��Ŀ����|δ��˵��|��¼���|������Դ|��ʾ|��ԴID|����|״̬"
        
        '��ȡ������Ϣ
        mstrSQL = "" & _
             " Select C.ID ���, B.����ʱ�� AS ʱ��,C.��¼����,C.��Ŀ���,C.δ��˵��,C.��¼����,C.��¼���,C.��Ŀ����,C.������Դ,C.��ʾ,C.��ԴID,C.����" & _
             " FROM ���˻����ļ� A, ���˻������� B, ���˻�����ϸ C" & _
             " Where A.ID=B.�ļ�ID and  B.ID = C.��¼ID AND A.ID=[1]  AND Nvl(A.Ӥ��, 0)=[4] AND a.����id=[2] AND a.��ҳid=[3] And c.��ֹ�汾 Is Null" & _
             " AND c.��¼����=4  AND B.����ʱ�� BETWEEN [5]  And [6]"
             
        If mblnMove Then
            mstrSQL = Replace(mstrSQL, "���˻����ļ�", "H���˻����ļ�")
            mstrSQL = Replace(mstrSQL, "���˻�������", "H���˻�������")
            mstrSQL = Replace(mstrSQL, "���˻�����ϸ", "H���˻�����ϸ")
        End If
        
        strTime = CDate(Format(mstrBegin, "YYYY-MM-DD") & " 23:59:59")
        If CDate(strTime) > CDate(mstrETime) Then strTime = Format(mstrETime, "YYYY-MM-DD HH:mm:ss")
        
        Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, "��ȡ���������±����Ϣ", mT_Patient.lng�ļ�ID, mT_Patient.lng����ID, mT_Patient.lng��ҳID, _
            mT_Patient.lngӤ��, Int(CDate(Format(mstrBegin, "YYYY-MM-DD"))), CDate(strTime))
        With rsTmp
            Do While Not .EOF
                gstrValues = zlCommFun.Nvl(!���) & "|" & zlCommFun.Nvl(!��Ŀ���, 0) & "|" & Format(zlCommFun.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss") & "|" & Format(zlCommFun.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss") & "|" & zlCommFun.Nvl(!��¼����) & "|" & _
                    zlCommFun.Nvl(!��¼����) & "|" & zlCommFun.Nvl(!��Ŀ����) & "|" & Nvl(!δ��˵��) & "|" & zlCommFun.Nvl(!��¼���, 0) & "|" & Val(zlCommFun.Nvl(!������Դ, 0)) & "|" & _
                    Val(zlCommFun.Nvl(!��ʾ, 0)) & "|" & Val(zlCommFun.Nvl(!��ԴID, 0)) & "|" & Val(zlCommFun.Nvl(!����, 0)) & "|0"
                Call Record_Add(mrsOper, gstrFields, gstrValues)
            .MoveNext
            Loop
        End With
        
        '���������Ϣ
        mrsOper.Filter = 0
        mrsOper.Sort = "ʱ��"
        With mrsOper
            vsfOper.Rows = vsfOper.FixedRows
            Do While Not .EOF
                vsfOper.Rows = vsfOper.Rows + 1
                vsfOper.TextMatrix(vsfOper.Rows - 1, Col_OperTime) = Format(!ʱ��, "HH:mm")
                vsfOper.TextMatrix(vsfOper.Rows - 1, Col_OperType) = Nvl(!��Ŀ����, "����")
                If InStr(1, ",0,9,", "," & Val(zlCommFun.Nvl(!������Դ)) & ",") = 0 Then
                    vsfOper.Cell(flexcpForeColor, vsfOper.Rows - 1, Col_OperTime, vsfOper.Rows - 1, Col_OperType) = 255
                Else
                    vsfOper.Cell(flexcpForeColor, vsfOper.Rows - 1, Col_OperTime, vsfOper.Rows - 1, Col_OperType) = &H80000012
                End If
                vsfOper.RowData(vsfOper.Rows - 1) = Val(!���)
            .MoveNext
            Loop
            vsfOper.Rows = vsfOper.Rows + 1
        End With
        
        strItems = ""
        '3------------------------------------------------------------------------------------------------------------
        '��֯��Ŀ��Ϣ
        For intRow = vsfTab.FixedRows To vsfTab.Rows - 1
            lng��Ŀ��� = Val(vsfTab.TextMatrix(intRow, COL_tab��Ŀ���))
            If lng��Ŀ��� <> 4 Then
                i = InStr(1, vsfTab.TextMatrix(intRow, COL_tab��Ŀ����), "(")
                If i > 0 Then
                    strItemName = Trim(Left(vsfTab.TextMatrix(intRow, COL_tab��Ŀ����), i - 1))
                Else
                    strItemName = Trim(vsfTab.TextMatrix(intRow, COL_tab��Ŀ����))
                End If
                If InStr(1, "," & strItems & ",", ",'" & strItemName & "',") = 0 Then
                    strItems = strItems & ",'" & strItemName & "'"
                End If
            End If
        Next intRow
        
        If Left(strItems, 1) = "," Then strItems = Mid(strItems, 2)
        strItems = strItems & ",'����ѹ','����ѹ'"
        If Left(strItems, 1) = "," Then strItems = Mid(strItems, 2)
        
        '��ȡһ����(���ܺ��еڶ�������)���еı��������Ϣ
        mstrSQL = "SELECT C.Id,a.����ʱ�� As ʱ��,C.��¼����,C.��ʾ,C.��¼���� As ���,C.���²�λ,C.δ��˵��,nvl(C.������Դ,0) ������Դ," & vbNewLine & _
            "  DECODE(E.��Ŀ����,2,C.���²�λ || D.��¼��,D.��¼��) ��Ŀ����,D.��Ŀ���,C.��ԴID,C.����,E.��Ŀ���� " & _
            "  FROM ���˻����ļ� B, ���˻������� A,���˻�����ϸ C,���¼�¼��Ŀ D,�����¼��Ŀ E " & _
            "  Where B.ID = A.�ļ�ID" & vbNewLine & _
            "  AND A.ID = C.��¼ID" & vbNewLine & _
            "  AND B.ID = [1]" & vbNewLine & _
            "  AND Nvl(B.Ӥ��, 0) = [7]" & vbNewLine & _
            "  AND B.����id = [2]" & vbNewLine & _
            "  AND B.��ҳid = [3]" & vbNewLine & _
            "  AND INSTR([6], DECODE(E.��Ŀ����, 2,C.���²�λ || D.��¼��, D.��¼��)) > 0" & vbNewLine & _
            "  AND D.��Ŀ��� = C.��Ŀ���" & vbNewLine & _
            "  AND Mod(c.��¼����,10) = 1" & vbNewLine & _
            "  AND E.��Ŀ��� = D.��Ŀ���" & vbNewLine & _
            "  AND E.����ȼ� >= [8]" & vbNewLine & _
            "  AND A.����ʱ�� BETWEEN [4] And [5]" & vbNewLine & _
            "  And C.��ֹ�汾 Is Null" & vbNewLine & _
            "  AND D.��¼�� = 2 And D.��Ŀ���<>3" & vbNewLine & _
            "  UNION ALL "
        '��ȡ�����±��Ļ�����Ŀ�����±�������Ŀ������ܴ��ڷ�������Ŀ��
        mstrSQL = mstrSQL & vbNewLine & _
            "  SELECT C.ID,a.����ʱ�� As ʱ��,C.��¼����,C.��ʾ,C.��¼���� As ���,C.���²�λ,C.δ��˵��,nvl(C.������Դ,0) ������Դ," & _
            "   D.��Ŀ����,D.��Ŀ���,C.��ԴID,C.����,D.��Ŀ����" & _
            "   FROM ���˻����ļ� B, ���˻������� A,���˻�����ϸ C,(SELECT A.��Ŀ���,A.��Ŀ����, 1 ��Ŀ����,B.����� FROM �����¼��Ŀ A,���������Ŀ B" & vbNewLine & _
            "       WHERE A.��Ŀ���=B.��� AND NOT EXISTS (SELECT C.��Ŀ��� FROM ���¼�¼��Ŀ C,���������Ŀ E WHERE C.��Ŀ���=E.��� AND C.��Ŀ���=A.��Ŀ���)" & vbNewLine & _
            "       AND NVL(A.Ӧ�÷�ʽ,0)=1 AND NVL(A.����ȼ�,0)>=[8] AND NVL(A.���ò���,0) IN (0,[9])" & vbNewLine & _
            "       AND (A.���ÿ���=1 OR (A.���ÿ���=2 AND EXISTS (SELECT 1 FROM �������ÿ��� D WHERE D.��Ŀ���=A.��Ŀ��� AND D.����ID=[10])))) D" & _
            "   Where B.ID=A.�ļ�ID And A.ID = C.��¼ID   AND B.ID=[1]  AND Nvl(B.Ӥ��,0)=[7] " & _
            "   AND B.����id=[2]  AND B.��ҳid=[3]  AND D.��Ŀ���=C.��Ŀ���  AND C.��¼����=1" & _
            "   AND A.����ʱ�� BETWEEN [4] And [5] And C.��ֹ�汾 Is Null"
            
        mstrSQL = _
            "   Select ID,ʱ��,��¼����,��ʾ,���,���²�λ,δ��˵��,������Դ,��Ŀ����,��Ŀ���,��ԴID,����,��Ŀ���� From (" & mstrSQL & ")" & _
            "   Order By  Decode(��Ŀ����,'����ѹ',0,1)," & strItems & ",ʱ��"
        If mblnMove Then
            mstrSQL = Replace(mstrSQL, "���˻����ļ�", "H���˻����ļ�")
            mstrSQL = Replace(mstrSQL, "���˻�������", "H���˻�������")
            mstrSQL = Replace(mstrSQL, "���˻�����ϸ", "H���˻�����ϸ")
        End If
        
        strTime = CDate(Format(mstrBegin, "YYYY-MM-DD") & " 23:59:59")
        
        dtBegin = Int(CDate(mstrBegin) - 1)
        dtEnd = CDate(CDate(Format(strTime, "YYYY-MM-DD HH:mm:ss")) + 1)
        If CDate(Format(dtBegin, "YYYY-MM-DD HH:mm:ss")) < CDate(Format(mstrBTime, "YYYY-MM-DD HH:mm:ss")) Then _
            dtBegin = CDate(Format(mstrBTime, "YYYY-MM-DD HH:mm:ss"))
        If CDate(Format(dtEnd, "YYYY-MM-DD HH:mm:ss")) > CDate(Format(mstrETime, "YYYY-MM-DD HH:mm:ss")) Then _
            dtEnd = CDate(Format(mstrETime, "YYYY-MM-DD HH:mm:ss"))
        
        Set rsTmp = zlDatabase.OpenSQLRecord(mstrSQL, Me.Caption, _
                                            mT_Patient.lng�ļ�ID, _
                                            mT_Patient.lng����ID, _
                                            mT_Patient.lng��ҳID, _
                                            CDate(dtBegin), _
                                            CDate(dtEnd), _
                                            strItems, mT_Patient.lngӤ��, mT_Patient.lng����ȼ�, IIf(mT_Patient.lngӤ�� = 0, 1, 2), mT_Patient.lng����ID)
        
        gstrFields = "���|������|��ֵ|��λ|���|ʱ��|ԭʼʱ��|��Ŀ���|��Ŀ����|����|δ��˵��|������Դ|�޸�|��ʾ|��ԴID|����|״̬|�к�|��¼����"
        gbln��Ժ = mbln��Ժ
        For intRow = vsfTab.FixedRows To vsfTab.Rows - 1
            If vsfTab.TextMatrix(intRow, COL_tab������) = "2)���±����Ŀ" Then
                int��Ŀ���� = Val(Split(vsfTab.TextMatrix(intRow, COL_tab�ַ���), ",")(5))
                int��¼Ƶ�� = Val(Split(vsfTab.TextMatrix(intRow, COL_tab�ַ���), ",")(3))
                int��Ŀ��ʾ = Val(Split(vsfTab.TextMatrix(intRow, COL_tab�ַ���), ",")(4))
                int��Ժ�ײ� = Val(Split(vsfTab.TextMatrix(intRow, COL_tab�ַ���), ",")(8))
                lng��Ŀ��� = Val(vsfTab.TextMatrix(intRow, COL_tab��Ŀ���))
                str��Ŀ���� = Split(vsfTab.TextMatrix(intRow, COL_tab��Ŀ����), "(")(0)
             
                intNum = 0
                strName = ""
                
                Set rsDownTab = ReturnItemRecord(rsTmp, Int(CDate(mstrBegin)), CDate(mstrBTime), lng��Ŀ��� & ";" & str��Ŀ���� & ";" & _
                                int��¼Ƶ�� & ";" & int��Ŀ��ʾ & ";" & int��Ŀ���� & ";" & int��Ժ�ײ�, mbln���ܵ���, mbln¼��Сʱ, True)
                If rsDownTab.RecordCount > 0 Then rsDownTab.MoveFirst
                rsDownTab.Sort = "ʱ��,��Ŀ���,���"
    
                With rsDownTab
                    Do While Not .EOF
                        blnAdd = False
                        intModify = IIf(InStr(1, ",0,9,", "," & Val(zlCommFun.Nvl(!������Դ)) & ",") = 0, 1, 0)
                        If zlCommFun.Nvl(!���) <> intNum Or zlCommFun.Nvl(!��Ŀ����) <> strName Then
                            intNum = zlCommFun.Nvl(!���)
                            strName = zlCommFun.Nvl(!��Ŀ����)
                            '����ѹ/����ѹ
                            If lng��Ŀ��� = 4 And str��Ŀ���� = "Ѫѹ" Then
                                Select Case zlCommFun.Nvl(!��Ŀ����)
                                    Case "����ѹ"
                                        strParam = ""
                                        strParam = zlCommFun.Nvl(!��¼����)
                                    Case "����ѹ"
                                        If InStr(strParam, "/") > 0 Then
                                            strParam = strParam & zlCommFun.Nvl(!��¼����)
                                        Else
                                            strParam = strParam & "/" & zlCommFun.Nvl(!��¼����)
                                        End If
                                        '--�����:53505,�޸��ˣ�����,Ѫѹ��ʾ����
                                        If zlCommFun.Nvl(!��¼����) = "���" Or zlCommFun.Nvl(!��¼����) = "δ��" Or zlCommFun.Nvl(!��¼����) = "�ܲ�" Or zlCommFun.Nvl(!��¼����) = "���" Then
                                            strParam = zlCommFun.Nvl(!��¼����)
                                        End If
                                        If strParam = "/" Then strParam = ""
                                        blnAdd = True
                                End Select
                            Else
                                strParam = zlCommFun.Nvl(!��¼����)
                                blnAdd = True
                            End If
        
                            If blnAdd = True Then
                                '��ȡ����ʱ�Ǹ���ʱ��κ���ʾ˳������ġ����һ��ʱ����ж�������,ֻ��ȡǰһ��
                                mrsCurve.Filter = "������='2)���±����Ŀ' and ��Ŀ���=" & lng��Ŀ��� & " and ��Ŀ����='" & str��Ŀ���� & "' and �к�=" & Val(zlCommFun.Nvl(!���, 0))
                                '51282,������,2012-08-03,ȫ�������ʾ¼��ʱ��(DYEYҪ���ֹ�¼�����ʱ��H)
                                If mrsCurve.RecordCount = 0 Then
                                    If Val(Nvl(!����Сʱ)) = -1 Then
                                        If InStr(1, strParam, ")") > 0 Then
                                            strName = Replace(Replace(Split(strParam, ")")(0), "(", ""), "h", "")
                                            strParam = Split(strParam, ")")(1)
                                        Else
                                            strName = ""
                                        End If
                                        gstrValues = Val(Split(!����Сʱ, ";")(2)) & "|2)���±����Ŀ|" & strName & "|" & _
                                            zlCommFun.Nvl(!���²�λ) & "|0|" & Format(zlCommFun.Nvl(Split(!����Сʱ, ";")(1)), "YYYY-MM-DD HH:mm:ss") & "|" & _
                                            Format(zlCommFun.Nvl(Split(!����Сʱ, ";")(1)), "YYYY-MM-DD HH:mm:ss") & "|" & lng��Ŀ��� & "|" & str��Ŀ���� & "|0|" & _
                                            Null & "|0|0|1| 0|0|0|" & zlCommFun.Nvl(!���, 0) & "|11"
                                            Call Record_Add(mrsCurve, gstrFields, gstrValues)
                                    End If
                                    gstrValues = zlCommFun.Nvl(!Id) & "|2)���±����Ŀ|" & strParam & "|" & _
                                        zlCommFun.Nvl(!���²�λ) & "|0|" & Format(zlCommFun.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss") & "|" & _
                                        Format(zlCommFun.Nvl(!ʱ��), "YYYY-MM-DD HH:mm:ss") & "|" & lng��Ŀ��� & "|" & str��Ŀ���� & "|0|" & _
                                        zlCommFun.Nvl(!δ��˵��) & "|" & Val(zlCommFun.Nvl(!������Դ, 0)) & "|" & intModify & "|" & Val(zlCommFun.Nvl(!��ʾ, 0)) & "|" & _
                                        Val(zlCommFun.Nvl(!��ԴID, 0)) & "|" & Val(zlCommFun.Nvl(!����, 0)) & "|0|" & zlCommFun.Nvl(!���, 0) & "|1"
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
        
        'չʾ���±������
        mrsCurve.Filter = 0
        mrsCurve.Filter = "������='2)���±����Ŀ'"
        mrsCurve.Sort = "��Ŀ���,�к�,��¼����"
        
        vsfTab.Cell(flexcpText, vsfTab.FixedRows, vsfTab.FixedCols, vsfTab.Rows - 1, vsfTab.Cols - 1) = ""
        strTime = ""
        With mrsCurve
            Do While Not .EOF
                For intRow = vsfTab.FixedRows To vsfTab.Rows - 1
                    blnAllow = False
                    If vsfTab.TextMatrix(intRow, COL_tab��Ŀ���) = !��Ŀ��� And vsfTab.TextMatrix(intRow, COL_tab������) = !������ Then
                        If Val(Split(vsfTab.TextMatrix(intRow, COL_tab�ַ���), ",")(5)) = 2 Then
                            If Split(Trim(vsfTab.TextMatrix(intRow, COL_tab��Ŀ����)), "(")(0) <> !��Ŀ���� Then
                                blnAllow = False
                            Else
                                blnAllow = True
                            End If
                        Else
                            blnAllow = True
                        End If
                        If blnAllow = True Then
                            If Val(Nvl(!��¼����)) = 11 Then
                                vsfTab.TextMatrix(intRow, vsfTab.FixedCols + Val(!�к�) - 1) = "(" & !��ֵ & "h)" & vsfTab.TextMatrix(intRow, vsfTab.FixedCols + Val(!�к�) - 1)
                            Else
                                vsfTab.TextMatrix(intRow, vsfTab.FixedCols + Val(!�к�) - 1) = !��ֵ
                                If Val(zlCommFun.Nvl(!������Դ)) <> 0 Then
                                    vsfTab.Cell(flexcpForeColor, intRow, vsfTab.FixedCols + Val(!�к�) - 1, intRow, vsfTab.FixedCols + Val(!�к�) - 1) = 255
                                Else
                                    vsfTab.Cell(flexcpForeColor, intRow, vsfTab.FixedCols + Val(!�к�) - 1, intRow, vsfTab.FixedCols + Val(!�к�) - 1) = &H80000012
                                End If
                                If Val(Split(vsfTab.TextMatrix(intRow, COL_tab�ַ���), ",")(5)) = 2 And Val(Split(vsfTab.TextMatrix(intRow, COL_tab�ַ���), ",")(1)) = 1 And _
                                    Val(Split(vsfTab.TextMatrix(intRow, COL_tab�ַ���), ",")(4)) = 0 Then
                                     vsfTab.Cell(flexcpForeColor, intRow, vsfTab.FixedCols + Val(!�к�) - 1, intRow, vsfTab.FixedCols + Val(!�к�) - 1) = Val(zlCommFun.Nvl(!δ��˵��))
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
    
    '��δˢ�µļ�¼����ԭʼ��¼��
    If blnCurve = False Or blnTab = False Then 'ֻˢ����������
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

Private Function GetPart(ByVal lng��Ŀ��� As Long) As String
'����:��ȡĬ�ϵ����²�λ
    Dim strPart As String
    mrsPart.Filter = "��Ŀ���=" & lng��Ŀ��� & " and ȱʡ��=1"
    If mrsPart.RecordCount > 0 Then strPart = zlCommFun.Nvl(mrsPart("��λ"))
    GetPart = strPart
End Function

Private Function GetCenterTime(ByVal dBegin As Date, ByVal dEnd As Date, Optional intBound As Integer = 0) As String
'------------------------------------------------------------------------------------
'����:��ȡĳ��ʱ����е�ʱ��,�����ǰʱ���ڱ��η�Χ����С���м�ʱ�������Ե�ǰʱ��Ϊ׼
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
        strTmp = "ͬ�����ݲ����޸�-255||ͬ�����ݿ����޸�-" & RGB(0, 0, 255) & "||��ȫ�޸�-0"
        'frmDataType.ShowPatiType Me, strTmp
        Call zlDatabase.ShowPatiColorTip(Me)
    End If
End Sub

Private Sub tbcThis_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)

    If Not mblnInit Then Exit Sub
    
    If Item.Tag = "���" Then
        If picEdit.Visible = False Then
            Call SetColSelect(True)
        Else
            Call SetColSelect
            txtEdit.SetFocus
        End If
    ElseIf Item.Tag = "����" Then
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
    
    'ˢ��ʱ�㰴ť��ʾ״̬
    
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
        '��������Ϊ�������͵Ļ��Ŀʹ�ÿ�ݼ����Ե���������ɫ����
        If cmdColor.Visible = True And Shift = vbShiftMask And Val(Split(vsfTab.TextMatrix(vsfTab.Row, COL_tab�ַ���), ",")(1)) = 1 _
            And Val(Split(vsfTab.TextMatrix(vsfTab.Row, COL_tab�ַ���), ",")(5)) = 2 And Val(Split(vsfTab.TextMatrix(vsfTab.Row, COL_tab�ַ���), ",")(4)) = 0 Then
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
        '������ݺϷ���
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
            '�ƶ�����һ��
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
        If lblCheck.Caption = "��" Then
            lblCheck.Caption = ""
        Else
            lblCheck.Caption = "��"
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
    Dim strTmp As String, lng��Ŀ��� As Long, str��Ŀ���� As String
    
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
    
    lng��Ŀ��� = Val(vsfTab.TextMatrix(intRow, COL_tab��Ŀ���))
    str��Ŀ���� = Split(vsfTab.TextMatrix(intRow, COL_tab��Ŀ����), "(")(0)

    mrsCurve.Filter = "��Ŀ���=" & lng��Ŀ��� & " and ��Ŀ����='" & str��Ŀ���� & "' and �к�=" & intCOl - vsfTab.FixedCols + 1
    If mrsCurve.RecordCount > 0 Then
        mrsCurve!δ��˵�� = cmdColor.Tag
        If mrsCurve!״̬ <> 1 And mrsCurve!״̬ <> 3 Then 'ԭ�е����� �޸ġ�ɾ�����״̬ʼ��Ϊ2
            mrsCurve!״̬ = 2
            mrsCurve!��ֵ = vsfTab.TextMatrix(intRow, intCOl)
        Else '�����������ݵĴ���
            If Trim(vsfTab.TextMatrix(intRow, intCOl)) = "" Then
                mrsCurve!״̬ = 3
                mrsCurve!��ֵ = vsfTab.TextMatrix(intRow, intCOl)
            Else
                mrsCurve!״̬ = 1
                mrsCurve!��ֵ = vsfTab.TextMatrix(intRow, intCOl)
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
    If Val(vsfCurve.Cell(flexcpBackColor, usrValue.Tag, col_��ɫ, usrValue.Tag, col_��ɫ)) = usrValue.Color Then PicValue.Visible = False: GoTo ErrNext
    vsfCurve.Cell(flexcpBackColor, usrValue.Tag, col_��ɫ, usrValue.Tag, col_��ɫ) = usrValue.Color
    If Trim(vsfCurve.TextMatrix(usrValue.Tag, col_����)) = "" Then GoTo ErrNext
    If Not UpdateCurveDate(usrValue.Tag, col_����, 2) Then vsfCurve.Cell(flexcpBackColor, usrValue.Tag, col_��ɫ, usrValue.Tag, col_��ɫ) = Val(PicValue.Tag)
ErrNext:
    PicValue.Visible = False
    If Val(usrValue.Tag) <= vsfCurve.Rows - 1 Then
        vsfCurve.Body.Select Val(usrValue.Tag), col_����
    End If
    vsfCurve.SetFocus
End Sub

Private Sub vsfCurve_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strTmp As String
    Dim lng��Ŀ��� As Long
    Dim strDate As String
    Dim lngRect As Long
    On Error Resume Next
    vsfCurve.ComboList(COL_��λ) = ""
    vsfCurve.EditMode(COL_��λ) = 0
    vsfCurve.EditMode(Col_δ��˵��) = 0
    lngRect = vsfCurve.Body.FocusRect

    lng��Ŀ��� = Val(vsfCurve.TextMatrix(NewRow, COL_��Ŀ���))
    strDate = Trim(vsfCurve.TextMatrix(NewRow, col_����))
    Select Case Trim(vsfCurve.TextMatrix(NewRow, COL_������))
    
    Case "1)����������Ŀ"
        vsfCurve.EditMode(Col_δ��˵��) = 1
        If Not mrsPart Is Nothing Then
            mrsPart.Filter = "��Ŀ���=" & lng��Ŀ���
            mrsPart.Sort = "ȱʡ�� DESC"
            With mrsPart
                Do While Not .EOF
                    strTmp = IIf(strTmp = "", zlCommFun.Nvl(!��λ), strTmp & "|" & zlCommFun.Nvl(!��λ))
                .MoveNext
                Loop
            End With
            If strTmp <> "" Then
                If lng��Ŀ��� = 2 And InStr(1, strTmp, "|") = 0 Then
                    strTmp = " |����"
                End If
                vsfCurve.ComboList(COL_��λ) = strTmp
                vsfCurve.EditMode(COL_��λ) = 1
            End If
        End If
        
        If NewCol = col_���� Or NewCol = Col_δ��˵�� Then
            '������Դ
            If InStr(1, ",0,9,", "," & Val(marrDate(NewRow)) & ",") = 0 Then
                If NewCol = col_���� Then
                    If lng��Ŀ��� = 1 And strDate = "����" Then GoTo NotEdit
                    If lng��Ŀ��� = 1 Or (lng��Ŀ��� = 2 And mint����Ӧ�� = 2 And mblnEdit���� = True) Then
                        If InStr(1, strDate, "/") = 0 Then
                            GoTo GoNext
                        Else
                            If mArrModfy(NewRow) = 0 Then GoTo GoNext
                        End If
                    End If
                End If
            End If
            '������Դ
            If InStr(1, ",0,9,", "," & Val(marrDate(NewRow)) & ",") = 0 Then
NotEdit:
                vsfCurve.EditMode(NewCol) = 0
            Else
GoNext:
                vsfCurve.EditMode(NewCol) = 1
            End If
        End If
        
    Case "2)���±�˵��"
        vsfCurve.EditMode(Col_δ��˵��) = 0
        vsfCurve.EditMode(col_����) = 1
    End Select
        
    strTmp = ""
    
    If Trim(Split(vsfCurve.TextMatrix(NewRow, COL_�ַ���), ",")(0)) <> "" Then
        strTmp = "���ݷ�Χ��" & Trim(Split(vsfCurve.TextMatrix(NewRow, COL_�ַ���), ",")(0)) & " "
    End If
    
    If Trim(vsfCurve.TextMatrix(NewRow, COL_������)) = "1)����������Ŀ" Then
        Select Case lng��Ŀ���
            Case 1 '����
                strTmp = strTmp & Space(4) & "�����±�ʾ��38/37"
            Case 2
                If mint����Ӧ�� = 2 And mblnEdit���� Then strTmp = strTmp & Space(4) & "������׾��ʾ��100/130"
        End Select
    ElseIf Trim(vsfCurve.TextMatrix(NewRow, COL_������)) = "2)���±�˵��" Then
'        If lng��Ŀ��� = 4 Then
'            strTmp = "������:��������������ʱ��(��:04:00),��λ/������ѡ������."
'        End If
        strTmp = "�������а�SHIFT+����˫����ɫ��������ɫ����"
    End If
    
    'stbThis.Panels(2).Text = strTmp
    lblStb.Caption = strTmp
    lblStb.ForeColor = &H80000012

End Sub

Private Sub vsfCurve_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    Dim lngWidth As Long
    If Col = col_��ɫ Then
        lngWidth = vsfCurve.Body.ColWidth(Col)
        vsfCurve.Body.ColWidth(col_��ɫ) = 300
        vsfCurve.Body.ColWidth(col_����) = vsfCurve.Body.ColWidth(col_����) + lngWidth - 300
        If vsfCurve.Body.ColWidth(col_����) < 500 Then vsfCurve.Body.ColWidth(col_����) = 500
        Call vsfCurve_KeyDown(vbKeyDown, vbShiftMask)
    End If
End Sub

Private Sub vsfCurve_BeforeDeleteRow(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    Dim blnAllow As Boolean
    Dim intType As Integer
    
    vsfCurve.Tag = vsfCurve.TextMatrix(Row, Col)
    
    Select Case Col
        Case COL_��λ
            vsfCurve.TextMatrix(Row, Col) = ""
            If Trim(vsfCurve.TextMatrix(Row, COL_������)) = "2)���±�˵��" Then
                intType = 2
            ElseIf Trim(vsfCurve.TextMatrix(Row, COL_������)) = "1)����������Ŀ" Then
                intType = 1
            End If
            blnAllow = True
        Case Col_δ��˵��
            If Trim(vsfCurve.TextMatrix(Row, COL_������)) = "1)����������Ŀ" And vsfCurve.TextMatrix(Row, Col) <> "" Then
                vsfCurve.TextMatrix(Row, Col) = ""
                vsfCurve.TextMatrix(Row, col_����) = Space(Row) & IIf(vsfCurve.TextMatrix(Row, COL_������) = "2)���±�˵��", "", "") & Space(Row)
                vsfCurve.TextMatrix(Row, col_��ɫ) = Space(Row) & IIf(vsfCurve.TextMatrix(Row, COL_������) = "2)���±�˵��", " ", "") & Space(Row)
                vsfCurve.TextMatrix(Row, COL_��λ) = ""
                vsfCurve.TextMatrix(Row, col_����) = ""
                blnAllow = True
                intType = 1
            End If
        Case col_����
            If vsfCurve.TextMatrix(Row, Col) <> "" Then
                If Trim(vsfCurve.TextMatrix(Row, COL_������)) = "2)���±�˵��" Then
                    intType = 2
                ElseIf Trim(vsfCurve.TextMatrix(Row, COL_������)) = "1)����������Ŀ" Then
                    intType = 1
                    If InStr(1, ",0,9,", "," & Val(marrDate(Row)) & ",") = 0 Then
                        Cancel = True
                        lblStb.Caption = "�ɻ����¼���������ط�ͬ�����������ݲ���ɾ��."
                        lblStb.ForeColor = 255
                        vsfCurve.Body.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
                        Exit Sub
                    End If
                End If
                
                vsfCurve.TextMatrix(Row, col_����) = Space(Row) & IIf(vsfCurve.TextMatrix(Row, COL_������) = "2)���±�˵��", "", "") & Space(Row)
                vsfCurve.TextMatrix(Row, col_��ɫ) = Space(Row) & IIf(vsfCurve.TextMatrix(Row, COL_������) = "2)���±�˵��", " ", "") & Space(Row)
                vsfCurve.TextMatrix(Row, COL_��λ) = ""
                vsfCurve.TextMatrix(Row, col_����) = ""
                vsfCurve.TextMatrix(Row, Col_δ��˵��) = ""
                
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
'ѡ��δ��˵��
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    Dim i As Integer
    Dim blnSelect As Boolean
    
    If Trim(vsfCurve.TextMatrix(Row, COL_������)) <> "1)����������Ŀ" Then Exit Sub
    
    On Error GoTo Errhand
    Select Case Col
        Case Col_δ��˵��
            picδ��.Tag = Row & "|" & Col
            
            strSQL = "Select ����,���� From ��������˵��"
            Call zlDatabase.OpenRecordset(rsTemp, strSQL, Me.Caption)
            lstδ��.Clear
            If rsTemp.RecordCount > 0 Then
                i = 0
                With rsTemp
                    Do While Not .EOF
                        lstδ��.AddItem zlCommFun.Nvl(!����)
                        If zlCommFun.Nvl(!����) = vsfCurve.TextMatrix(vsfCurve.Row, vsfCurve.Col) Then
                            lstδ��.Selected(i) = True
                            blnSelect = True
                        End If
                        i = i + 1
                    .MoveNext
                    Loop
                End With
            End If
            
            If blnSelect = False And lstδ��.ListCount <> 0 Then lstδ��.Selected(0) = True
            
            If lstδ��.ListCount > 0 Then
                picδ��.Left = vsfCurve.CellLeft + vsfCurve.Left + 15
                picδ��.Top = vsfCurve.CellTop + vsfCurve.Top + vsfCurve.CellHeight
                If lstδ��.Height < vsfCurve.CellHeight + 20 Then lstδ��.Height = vsfCurve.CellHeight + 20
                lstδ��.Width = vsfCurve.CellWidth + 20
                picδ��.Height = lstδ��.Height
                picδ��.Width = lstδ��.Width
                
                If picδ��.Top + picδ��.Height > vsfCurve.Top + vsfCurve.Height Then
                    picδ��.Top = vsfCurve.CellTop + vsfTab.Top - picδ��.Height
                End If
                picδ��.Visible = True
                lstδ��.Visible = True: lstδ��.Enabled = True
                lstδ��.SetFocus
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
    Case col_����
        If vsfCurve.TextMatrix(vsfCurve.Row, COL_��Ŀ���) <> 0 Then
            vsfCurve.TextMatrix(vsfCurve.Row, col_����) = IIf(vsfCurve.EditText = "", " ", vsfCurve.EditText)
            If vsfCurve.TextMatrix(vsfCurve.Row, COL_������) <> "2)���±�˵��" Then
                vsfCurve.TextMatrix(vsfCurve.Row, col_��ɫ) = vsfCurve.TextMatrix(vsfCurve.Row, col_����)
            End If
            If vsfCurve.EditText <> "" Then vsfCurve.TextMatrix(vsfCurve.Row, Col_δ��˵��) = ""
        End If
    End Select
End Sub

Private Sub vsfCurve_ComboCloseUp(Row As Long, Col As Long, FinishEdit As Boolean)
    Dim intType As Integer
    Dim blnAllow As Boolean
        
    blnAllow = True
    If Trim(vsfCurve.TextMatrix(Row, COL_������)) = "1)����������Ŀ" Then
        intType = 1
    ElseIf Trim(vsfCurve.TextMatrix(Row, COL_������)) = "2)���±�˵��" Then
        If Val(vsfCurve.TextMatrix(Row, COL_��Ŀ���)) = 4 And vsfCurve.EditText <> "" Then
'            intType = 2
'
'            If Trim(vsfCurve.TextMatrix(Row, col_����)) = "" Then
'                vsfCurve.TextMatrix(Row, col_����) = Format(GetCenterTime(CDate(mstrBegin), CDate(mstrEnd)), "HH:mm")
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
                If .Col = Col_δ��˵�� Then
                    Call vsfCurve_CellButtonClick(.Row, .Col)
                ElseIf (.Col = col_���� Or .Col = col_��ɫ) And .TextMatrix(.Row, COL_������) = "2)���±�˵��" Then
                    vsfCurve.Tag = .TextMatrix(.Row, col_����)
                    PicValue.Top = .CellTop + .CellHeight + .Top
                    If PicValue.Top + PicValue.Height > .Top + .Height Then
                        PicValue.Top = .CellTop - PicValue.Height
                    End If
                    If PicValue.Top < .Top Then PicValue.Top = .Top
                    PicValue.Left = IIf(.Col = col_��ɫ, .CellLeft, .CellLeft + .CellWidth) + .Left
                    PicValue.Visible = True
                    PicValue.ZOrder 0
         
                    usrValue.Left = 0
                    usrValue.Top = -450
                    usrValue.Visible = True
                    usrValue.ZOrder 0
                    PicValue.SetFocus
                    usrValue.Color = Val(.Cell(flexcpBackColor, .Row, col_��ɫ, .Row, col_��ɫ))
                    PicValue.Tag = Val(usrValue.Color)
                    usrValue.Tag = .Row
                End If
            End If
        End If
    End With
End Sub

Private Sub vsfCurve_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, ByVal ComboList As String, KeyCode As Integer, ByVal Shift As Integer, Cancel As Boolean)
    If KeyCode = vbKeyReturn Then
        If Col = Col_δ��˵�� Then
            If InStr(1, "," & mstrδ��˵�� & ",", "," & vsfCurve.EditText & ",") = 0 Then
                vsfCurve.TextMatrix(Row, Col) = ""
                vsfCurve.Cell(flexcpData, Row, Col) = ""
            Else
                vsfCurve.TextMatrix(Row, Col) = vsfCurve.EditText
                vsfCurve.Cell(flexcpData, Row, Col) = vsfCurve.EditText
                vsfCurve.TextMatrix(Row, col_����) = Space(Row) & IIf(vsfCurve.TextMatrix(Row, COL_������) = "2)���±�˵��", "", "") & Space(Row)
                vsfCurve.TextMatrix(Row, col_��ɫ) = Space(Row) & IIf(vsfCurve.TextMatrix(Row, COL_������) = "2)���±�˵��", " ", "") & Space(Row)
                vsfCurve.TextMatrix(Row, COL_��λ) = ""
                vsfCurve.TextMatrix(Row, col_����) = ""
            End If
        End If
    End If
    If KeyCode = vbKeyDown And Shift = vbShiftMask And Col = col_���� Then
        Call vsfCurve_KeyDown(KeyCode, Shift)
        Cancel = True
    End If
End Sub

Private Sub vsfCurve_KeyPress(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer, Cancel As Boolean)
    If KeyAscii = 32 Then '
        If Col = col_���� Then
            If Val(vsfCurve.TextMatrix(Row, col_����)) <> 0 And Val(vsfCurve.TextMatrix(Row, COL_��Ŀ���)) = 1 Then
                If vsfCurve.TextMatrix(Row, Col) = "" Then
                    vsfCurve.TextMatrix(Row, Col) = "��"
                    vsfCurve.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
                Else
                    vsfCurve.TextMatrix(Row, Col) = ""
                End If
                Call UpdateCurveDate(Row, Col, 1)
            End If
        End If
        If Col = col_��ɫ And vsfCurve.TextMatrix(Row, COL_������) = "2)���±�˵��" Then
            Call vsfCurve_KeyDown(vbKeyDown, vbShiftMask)
        End If
    End If
End Sub

Private Sub vsfCurve_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim lngNO As Long
    Dim strDate As String
    
    On Error Resume Next
    lngNO = Val(vsfCurve.TextMatrix(Row, COL_��Ŀ���))
    strDate = vsfCurve.TextMatrix(Row, COL_tab��Ŀ����)
    
    If KeyAscii <> vbKeyReturn Then
        If lngNO <> 0 Then
            If vsfCurve.TextMatrix(Row, COL_������) = "1)����������Ŀ" Then
                If Col <> Col_δ��˵�� Then
                    If lngNO = 2 And mint����Ӧ�� = 2 And mblnEdit���� = True Then
                        If FilterKeyAscii(KeyAscii, 99, "0123456789./") = 0 Then KeyAscii = 0
                    ElseIf lngNO = 1 Then
                        '���²����м��
                    Else
                        If FilterKeyAscii(KeyAscii, 99, "0123456789.") = 0 Then KeyAscii = 0
                    End If
                Else
                    If FilterKeyAscii(KeyAscii, 99, "'") > 0 Then KeyAscii = 0
                End If
            ElseIf vsfCurve.TextMatrix(Row, COL_������) = "2)���±�˵��" And lngNO = 4 Then
'                If Col = col_���� Then
'                    If FilterKeyAscii(KeyAscii, 99, "0123456789:") = 0 Then KeyAscii = 0
'                End If
            End If
        End If
    End If
End Sub

Private Sub vsfCurve_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lng��Ŀ��� As Long, strDate As String
    Dim strName As String
    Dim intRow As Integer
    Dim strData As String
    
    lng��Ŀ��� = Val(vsfCurve.TextMatrix(Row, COL_��Ŀ���))
    strName = vsfCurve.TextMatrix(Row, COL_��Ŀ����)
    vsfCurve.Body.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignLeftCenter
        
    Select Case Col
        Case col_����
            vsfCurve.TextMatrix(Row, Col) = IIf(RTrim(LTrim(vsfCurve.TextMatrix(Row, Col))) = "", " ", RTrim(LTrim(vsfCurve.TextMatrix(Row, Col))))
            If Row <> mOptRow.�ϱ� And Row <> mOptRow.�±� Then
                vsfCurve.TextMatrix(Row, col_��ɫ) = vsfCurve.TextMatrix(Row, Col)
            Else
                vsfCurve.TextMatrix(Row, Col) = RTrim(LTrim(vsfCurve.TextMatrix(Row, Col)))
            End If
            vsfCurve.Body.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000018
            strDate = RTrim(LTrim(vsfCurve.TextMatrix(Row, Col)))
    End Select
    
    vsfCurve.Tag = RTrim(LTrim(vsfCurve.TextMatrix(Row, Col)))
     
    If Col = col_���� Or Col = Col_δ��˵�� Then
        '������Դ
        If InStr(1, ",0,9,", "," & Val(marrDate(Row)) & ",") = 0 Then
            If Col = col_���� Then
                If lng��Ŀ��� = 1 And strDate = "����" Then GoTo NotEdit
                If lng��Ŀ��� = 1 Or (lng��Ŀ��� = 2 And mint����Ӧ�� = 2 And mblnEdit���� = True) Then
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
            If lng��Ŀ��� = 1 Then
                lblStb.Caption = "ͬ��������[" & strName & "]����ֻ�����޸������²���."
            ElseIf lng��Ŀ��� = 2 And mint����Ӧ�� = 2 And mblnEdit���� = True Then
                lblStb.Caption = "ͬ��������[" & strName & "]����ֻ�����޸��������Ჿ��."
            Else
                lblStb.Caption = "�ɻ����¼���������ط�ͬ�����������ݲ����޸�"
            End If
            lblStb.ForeColor = 255
            vsfCurve.Body.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
        End If
    ElseIf col_���� = Col Then
        vsfCurve.Body.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
    End If
GoNext:
    If mblnFileBack = True Then
        Cancel = True
        vsfCurve.Body.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
        lblStb.Caption = "�������������Ѿ��鵵,��������������޸�."
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
    If Col = col_���� Then
        vsfCurve.TextMatrix(Row, col_����) = Space(Row) & vsfCurve.TextMatrix(Row, col_����) & Space(Row)
        vsfCurve.TextMatrix(Row, col_��ɫ) = IIf(vsfCurve.TextMatrix(Row, COL_������) = "2)���±�˵��", Space(Row + 1), vsfCurve.TextMatrix(Row, col_����))
    End If
End Sub

Private Sub vsfCurve_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strֵ�� As String
    Dim lngNO As Long, intС�� As Integer, intType As Integer
    Dim strInfo As String, strText As String, strName As String, strMsg As String, strDate As String
    Dim arrValue() As String
    Dim lngCount As Long, i As Long, strValue As String
    Dim blnOk As Boolean
    
    '������ݺϷ���
    If Col = col_���� Then
        strValue = vsfCurve.Tag
        Select Case vsfCurve.TextMatrix(Row, COL_������)
            Case "1)����������Ŀ"
                strֵ�� = Split(vsfCurve.TextMatrix(Row, COL_�ַ���), ",")(0)
                lngNO = Val(vsfCurve.TextMatrix(Row, COL_��Ŀ���))
                strName = vsfCurve.TextMatrix(Row, COL_��Ŀ����)
                intС�� = Val(Split(vsfCurve.TextMatrix(Row, COL_�ַ���), ",")(2))
                intType = 1
                GoTo CheckPoint
            Case "2)���±�˵��"
                If InStr(1, ",2,6,", "," & Val(vsfCurve.TextMatrix(Row, COL_��Ŀ���)) & ",") <> 0 Then
                    PicValue.Tag = vsfCurve.Cell(flexcpBackColor, Row, col_��ɫ, Row, col_��ɫ)
                    intType = 2: GoTo CheckTag
                End If
        End Select
    End If
    
    Exit Sub
CheckPoint:
    strDate = vsfCurve.EditText
    If Trim(vsfCurve.EditText) <> "" And strֵ�� <> "" Then
        strInfo = vsfCurve.EditText
        
        '���������������/��Ҫ�������������
        If lngNO = 2 And mint����Ӧ�� = 2 And mblnEdit���� = True Then
            If InStr(1, strInfo, "/") > 0 Then
                If Split(Trim(strInfo), "/")(1) = "" Or Split(Trim(strInfo), "/")(0) = "" Then
                    strMsg = strName & "����¼�����" & Space(4) & "��������:����/����"
                    GoTo ErrInfo
                Else
                    If Not IsNumeric(Split(Trim(strInfo), "/")(0)) Or Not IsNumeric(Split(Trim(strInfo), "/")(1)) Then
                        strMsg = strName & "����¼�����" & Space(4) & "��Ч��Χ:" & strֵ��
                        GoTo ErrInfo
                    End If
                End If
            End If
        End If
        
        If lngNO <> 1 And Not (lngNO = 2 And mint����Ӧ�� = 2 And mblnEdit���� = True) Then
            If InStr(1, strInfo, "/") Then
                strMsg = strName & "����¼�����" & Space(4) & "��Ч��Χ:" & strֵ��
                GoTo ErrInfo
            End If
        End If

        If UBound(Split(strInfo, "/")) > 1 Then
            strMsg = strName & "����¼��������飡"
            GoTo ErrInfo
        End If
        
        '�����������Ч��Χ���Ƿ���Ч
        arrValue = Split(strInfo, "/")
        lngCount = UBound(arrValue)
        For i = 0 To lngCount
            blnOk = False
            strText = arrValue(i)
            If i = 0 Then
                '����������Ŀ��Ҫ���˵�δ��˵��
                If InStr(1, strText, ";") <> 0 And UBound(arrValue) = 0 Then strText = Split(strText, ";")(1)
                If InStr(1, IIf(lngNO = 1, ",����,", ""), "," & strText & ",") = 0 Then
                    blnOk = False
                Else
                    blnOk = True
                End If
            End If
            
            If Not blnOk Then
                If Not IsNumeric(strText) Then
                    strMsg = strName & "����¼�����" & Space(4) & "��Ч��Χ:" & strֵ��
                    GoTo ErrInfo
                End If
            End If
            
            If Not blnOk And strText <> "" Then strText = Format(Val(strText), "#0" & IIf(intС�� > 0, ".", "") & String(intС��, "0"))
            If IsNumeric(Split(strֵ��, "��")(0)) And IsNumeric(strText) Then
                If Not (Val(strText) >= Split(strֵ��, "��")(0) And Val(strText) <= Split(strֵ��, "��")(1)) Then
                    strMsg = strName & "������Ч��Χ(" & strֵ�� & "),����!"
                    GoTo ErrInfo
                End If
            End If
        Next i
    End If
    
    '����������Դ<>0,9�� ����,�������� ���б༭(�������º������������¼��������,��������)
    If InStr(1, ",0,9,", "," & Val(marrDate(Row)) & ",") = 0 Then
        If Col = col_���� Then
            If lngNO = 1 Or (lngNO = 2 And mint����Ӧ�� = 2 And mblnEdit���� = True) Then
                '--�����:56853,�޸��ˣ����Σ����������������������¼�뷽ʽ�Ƿ���ȷ������/����
                If (InStr(strValue, "/") > 0 Or InStr(strValue, "/") = 0) And mbln����������ʾ Then
                    If InStr(1, strDate, "/") <> 0 Then
                        strDate = Split(strDate, "/")(1)
                    Else
                        strDate = Split(strDate, "/")(0)
                    End If
                     If strDate <> mArrValue(Row) Then
                        If lngNO = 1 Then
                            strMsg = "ͬ��������[" & strName & "]����ֻ�����޸������²���."
                        Else
                            strMsg = "ͬ��������[" & strName & "]����ֻ�����޸��������Ჿ��."
                        End If
                        
                        vsfCurve.TextMatrix(Row, col_����) = Space(Row) & Trim(CStr(mArrValue(Row))) & Space(Row)
                        vsfCurve.TextMatrix(Row, col_��ɫ) = vsfCurve.TextMatrix(Row, col_����)
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
                                strMsg = "ͬ��������[" & strName & "]����ֻ�����޸������²���."
                            Else
                                strMsg = "ͬ��������[" & strName & "]����ֻ�����޸��������Ჿ��."
                            End If
                            
                            vsfCurve.TextMatrix(Row, col_����) = Space(Row) & Trim(CStr(mArrValue(Row))) & Space(Row)
                            vsfCurve.TextMatrix(Row, col_��ɫ) = vsfCurve.TextMatrix(Row, col_����)
                            GoTo ErrInfo
                        End If
                    Else
                        If mArrModfy(Row) <> 0 Then
                            If strDate <> mArrValue(Row) Then
                                If lngNO = 1 Then
                                    strMsg = "ͬ��������[" & strName & "]�����������������,�������޸�."
                                Else
                                    strMsg = "ͬ��������[" & strName & "]�������������������,�������޸�."
                                End If
                                vsfCurve.TextMatrix(Row, col_����) = Space(Row) & CStr(mArrValue(Row)) & Space(Row)
                                vsfCurve.TextMatrix(Row, col_��ɫ) = vsfCurve.TextMatrix(Row, col_����)
                                GoTo ErrInfo
                            End If
                        Else
                            If strDate <> Split(mArrValue(Row), "/")(0) Then
                                If lngNO = 1 Then
                                    strMsg = "ͬ��������[" & strName & "]����ֻ�����޸������²���."
                                Else
                                    strMsg = "ͬ��������[" & strName & "]����ֻ�����޸��������Ჿ��."
                                End If
                                vsfCurve.TextMatrix(Row, col_����) = Space(Row) & CStr(mArrValue(Row)) & Space(Row)
                                vsfCurve.TextMatrix(Row, col_��ɫ) = vsfCurve.TextMatrix(Row, col_����)
                                GoTo ErrInfo
                            End If
                        End If
                    End If
                End If
            End If
        End If
    End If

    '��ʾȱʡ��λ
    If vsfCurve.TextMatrix(Row, COL_��λ) = "" And Trim(vsfCurve.TextMatrix(Row, col_����)) <> "" Then
        mrsPart.Filter = "��Ŀ���=" & lngNO & " and ȱʡ��=1"
        If mrsPart.RecordCount > 0 Then
            vsfCurve.TextMatrix(Row, COL_��λ) = CStr(zlCommFun.Nvl(mrsPart!��λ))
        End If
    End If
    
    GoTo ErrSaveData
    Exit Sub
CheckTag:
    GoTo ErrSaveData
    Exit Sub
ErrInfo:    '������Ϣ���
    'stbThis.Panels(2).Text = StrMsg
    lblStb.Caption = strMsg
    lblStb.ForeColor = 255
    vsfCurve.TextMatrix(Row, col_����) = Space(Row) & strValue & Space(Row)
    vsfCurve.TextMatrix(Row, col_��ɫ) = vsfCurve.TextMatrix(Row, col_����)
    Cancel = True
    Exit Sub
ErrSaveData:
     Call UpdateCurveDate(Row, Col, intType)
End Sub

Private Function UpdateCurveDate(ByVal intRow As Integer, ByVal intCOl As Integer, ByVal intType As Integer, _
    Optional blnComList As Boolean = False, Optional blnOper As Boolean = False) As Boolean
'------------------------------------------------------------------------
'����:����������Ŀ.����.���±�����ݱ���
'------------------------------------------------------------------------
    Dim lngNO As Long, strName As String, strTime As String, lngID As Long
    Dim strValue As String, int��� As Integer, strδ�� As String
    Dim str��λ As String
    Dim strData As String
    On Error GoTo Errhand:
    
    If Not blnOper Then
        lngNO = Val(vsfCurve.TextMatrix(intRow, COL_��Ŀ���))
        If UBound(Split(vsfCurve.TextMatrix(intRow, COL_��Ŀ����), "(")) = -1 Then
            strName = vsfCurve.TextMatrix(intRow, COL_��Ŀ����)
        Else
            strName = Split(vsfCurve.TextMatrix(intRow, COL_��Ŀ����), "(")(0)
        End If
        
        If blnComList = True Then
            str��λ = vsfCurve.EditText
            If str��λ = "" Then str��λ = vsfCurve.TextMatrix(intRow, COL_��λ)
        Else
            str��λ = vsfCurve.TextMatrix(intRow, COL_��λ)
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
        str��λ = ""
        
    End If
    If intType = 1 Then '�������ݴ���
        strValue = Trim(vsfCurve.TextMatrix(intRow, col_����))
       If lngNO = 2 And mint����Ӧ�� = 2 And mblnEdit���� = True Then
            '��ת��������������
            If mbln����������ʾ And InStr(strValue, "/") > 0 Then
                strData = Split(Trim(vsfCurve.TextMatrix(intRow, col_����)), "/")(1) & "/" & Split(Trim(vsfCurve.TextMatrix(intRow, col_����)), "/")(0)
                strValue = strData
            End If
        End If
        
        strδ�� = Trim(vsfCurve.TextMatrix(intRow, Col_δ��˵��))
        If strValue <> "" Then strδ�� = ""
        '�������ݸ��´���
        mrsCurve.Filter = "��Ŀ���=" & lngNO & " and ʱ��='" & Format(mArrdkpTime(dkpTime.Tag), "YYYY-MM-DD HH:mm:ss") & "'"
        
        If mrsCurve.RecordCount <> 0 Then
            If Val(mrsCurve!״̬) <> 1 And Val(mrsCurve!״̬) <> 3 Then
                mrsCurve!״̬ = 2
                mrsCurve!��ֵ = strValue
                mrsCurve!��λ = str��λ
                mrsCurve!���� = IIf(vsfCurve.TextMatrix(intRow, col_����) = "��", 1, 0)
                mrsCurve!�޸� = 0
                mArrModfy(intRow) = 0
                mrsCurve!δ��˵�� = strδ��
                
            Else
                If strValue = "" And strδ�� = "" Then
                    mrsCurve!״̬ = 3
                Else
                    mrsCurve!״̬ = 1
                End If

                mrsCurve!��ֵ = strValue
                mrsCurve!��λ = str��λ
                mrsCurve!���� = IIf(vsfCurve.TextMatrix(intRow, col_����) = "��", 1, 0)
                mrsCurve!δ��˵�� = strδ��
            End If
            mrsCurve.Update
        Else '��������
            If strValue <> "" Or strδ�� <> "" Then
                gstrFields = "���|������|��ֵ|��λ|���|ʱ��|��Ŀ���|��Ŀ����|����|δ��˵��|������Դ|�޸�|��ʾ|��ԴID|����|״̬|�к�|��¼����"
                gstrValues = GetMaxID & "|1)����������Ŀ|" & strValue & "|" & str��λ & "|" & _
                    int��� & "|" & Format(mArrdkpTime(dkpTime.Tag), "YYYY-MM-DD HH:mm:ss") & "|" & lngNO & "|" & strName & "|" & _
                    Val(vsfCurve.TextMatrix(intRow, col_����)) & "|" & strδ�� & "|0|0|0|0|0|1|0|1"
                Call Record_Add(mrsCurve, gstrFields, gstrValues)
            End If
        End If
        
    ElseIf intType = 2 Then '�������±괦��
    
        If Not blnOper Then
            strValue = LTrim(RTrim(vsfCurve.TextMatrix(intRow, col_����)))
            mrsNote.Filter = "��¼����=" & lngNO
            If mrsNote.RecordCount <> 0 Then
                If Val(mrsNote!״̬) <> 1 And Val(mrsNote!״̬) <> 3 Then 'his��ȡ������
                    mrsNote!״̬ = 2
                    mrsNote!���� = LTrim(RTrim(vsfCurve.TextMatrix(intRow, col_����)))
                    mrsNote!δ��˵�� = IIf(mrsNote!���� = "", "", vsfCurve.Cell(flexcpBackColor, intRow, col_��ɫ, intRow, col_��ɫ))
                Else
                    If strValue = "" Then
                        mrsNote!״̬ = 3
                        mrsNote!���� = strValue
                        mrsNote!δ��˵�� = ""
                    Else
                        mrsNote!״̬ = 1
                        mrsNote!���� = strValue
                        mrsNote!δ��˵�� = IIf(mrsNote!���� = "", "", vsfCurve.Cell(flexcpBackColor, intRow, col_��ɫ, intRow, col_��ɫ))
                    End If
                End If
                mrsNote.Update
            Else
                If lngNO = 2 Then
                    strName = "�ϱ�˵��"
                ElseIf lngNO = 6 Then
                    strName = "�±�˵��"
                End If
                strTime = GetCenterTime(CDate(mstrBegin), CDate(mstrEnd))
                
                If strValue <> "" Then
                    lngID = GetMaxID
                    gstrFields = "���|��Ŀ���|ʱ��|ԭʼʱ��|��¼����|����|��Ŀ����|δ��˵��|��¼���|������Դ|��ʾ|��ԴID|����|״̬"
                    gstrValues = lngID & "|" & 0 & "|" & strTime & "|" & strTime & "|" & lngNO & "|" & strValue & "|" & strName & "|" & IIf(lngNO = 4, "", vsfCurve.Cell(flexcpBackColor, intRow, col_��ɫ, intRow, col_��ɫ)) & "|0|0|0|0|0|1"
                    Call Record_Add(mrsNote, gstrFields, gstrValues)
                End If
            End If
        Else
            mrsOper.Filter = "��¼����=" & lngNO & " And ���=" & Val(vsfOper.RowData(intRow))
            If mrsOper.RecordCount <> 0 Then
                If Val(mrsOper!״̬) <> 1 And Val(mrsOper!״̬) <> 3 Then 'his��ȡ������
                    mrsOper!״̬ = 2
                    If Trim(strTime) = "" Or strName = "" Then
                       mrsOper!��Ŀ���� = ""
                       mrsOper!���� = ""
                    ElseIf Trim(strTime) <> "" And strName <> "" Then
                        mrsOper!��Ŀ���� = strName
                        mrsOper!���� = strName
                    End If
                    If Trim(strTime) <> "" Then mrsOper!ʱ�� = SetDate(Format(Format(mstrBegin, "YYYY-MM-DD") & " " & Trim(strTime) & ":00", "YYYY-MM-DD HH:mm:ss"))
                Else
                    If Trim(strTime) = "" Or strName = "" Then
                        mrsOper!״̬ = 3
                        mrsOper!��Ŀ���� = ""
                        mrsOper!���� = ""
                    Else
                        mrsOper!״̬ = 1
                        mrsOper!��Ŀ���� = strName
                        mrsOper!���� = strName
                    End If
                    If Trim(strTime) <> "" Then mrsOper!ʱ�� = SetDate(Format(Format(mstrBegin, "YYYY-MM-DD") & " " & Trim(strTime) & ":00", "YYYY-MM-DD HH:mm:ss"))
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
                    gstrFields = "���|��Ŀ���|ʱ��|ԭʼʱ��|��¼����|����|��Ŀ����|δ��˵��|��¼���|������Դ|��ʾ|��ԴID|����|״̬"
                    gstrValues = lngID & "|" & 0 & "|" & strTime & "|" & strTime & "|" & lngNO & "|" & strValue & "|" & strName & "|" & IIf(lngNO = 4, "", vsfCurve.Cell(flexcpBackColor, intRow, col_��ɫ, intRow, col_��ɫ)) & "|0|0|0|0|0|1"
                    vsfOper.RowData(intRow) = lngID
                    Call Record_Add(mrsOper, gstrFields, gstrValues)
                End If
            End If
        End If
        
    End If
    
    If intCOl = col_���� And Trim(vsfCurve.Tag) <> Trim(vsfCurve.TextMatrix(intRow, col_����)) Then
        mblnChage = True
        mblnCurveChange = True
    ElseIf intCOl = COL_��λ And Trim(vsfCurve.Tag) <> str��λ Then
        mblnChage = True
        mblnCurveChange = True
    ElseIf intType = 1 And intCOl = Col_δ��˵�� And Trim(vsfCurve.Tag) <> Trim(vsfCurve.TextMatrix(intRow, Col_δ��˵��)) Then
        mblnChage = True
        mblnCurveChange = True
    ElseIf intType = 2 And intCOl = col_���� And PicValue.Visible = True And PicValue.Tag <> vsfCurve.Cell(flexcpBackColor, intRow, col_��ɫ) Then
        mblnChage = True
        mblnCurveChange = True
    ElseIf lngNO = 1 And intCOl = col_���� Then
        mblnChage = True
        mblnCurveChange = True
    ElseIf lngNO = 4 Then '������Ϣ
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
    '����Ƿ���ͬ������������
    Dim lngID As Long, intState As Integer
    lngID = Val(vsfOper.RowData(Row))
    If lngID > 0 Then
        mrsOper.Filter = "��¼����=4 And ���=" & lngID
        intState = mrsOper!״̬
        If InStr(1, ",0,9,", "," & Val(Nvl(mrsOper!������Դ, 0)) & ",") = 0 Then
            Cancel = True
            lblStb.Caption = "ͬ������������,�������������ɾ��."
            lblStb.ForeColor = 255
            vsfOper.Body.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
        End If
        
        '������ݵ�ɾ������
        If intState = 0 Or intState = 2 Then '��ʾ��ԭ������
            mrsOper!���� = ""
            mrsOper!��Ŀ���� = ""
            mrsOper!״̬ = 2
        Else '��ʾ��������
            mrsOper.Delete
        End If
        mrsOper.Update
        mblnChage = True
    End If
End Sub

Private Sub vsfOper_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    Dim intRow As Integer
    '�����һ��û��¼��ʱ���������Ϣ ���ܽ�����һ��
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
        lblStb.Caption = "�������������Ѿ��鵵,��������������޸�."
        lblStb.ForeColor = 255
        vsfOper.Body.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
    End If
    
    '����Ƿ���ͬ������������
    If Val(vsfOper.RowData(Row)) > 0 Then
        mrsOper.Filter = "��¼����=4 And ���=" & Val(vsfOper.RowData(Row))
        If InStr(1, ",0,9,", "," & Val(Nvl(mrsOper!������Դ, 0)) & ",") = 0 Then
            Cancel = True
            lblStb.Caption = "ͬ������������,��������������޸�."
            lblStb.ForeColor = 255
            vsfOper.Body.Cell(flexcpBackColor, Row, Col, Row, Col) = &H80000005
        End If
    End If
End Sub

Private Sub vsfOper_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '�������ݺϷ��Լ��
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
            
            '�Ϸ��Լ��
            If Mid(strText, 3, 1) <> ":" Then
                strInfo = "¼���ʱ���ʽ�Ƿ���[Сʱ:����]"
                GoTo ErrInfo
            End If
            If Mid(strText, 1, 2) < 0 Or Mid(strText, 1, 2) > 23 Then
                strInfo = "¼���ʱ���ʽ�Ƿ���[СʱӦ��0��23֮��]"
                GoTo ErrInfo
            End If
            If Mid(strText, 4, 2) < 0 Or Mid(strText, 4, 2) > 59 Then
                strInfo = "¼���ʱ���ʽ�Ƿ���[����Ӧ��0��59֮��]"
                GoTo ErrInfo
            End If
            .EditText = Format(strText, "HH:mm")
            
            '���¼���ʱ���Ƿ��Ѿ�������������Ϣ
            strDate = Format(dkpDate.Value & " " & strText, "YYYY-MM-DD HH:mm:ss")
            gstrSQL = "select 1 from ���˻����ļ� A,���˻������� B,���˻�����ϸ C" & _
                " Where A.ID=B.�ļ�ID And B.ID=C.��¼ID And A.ID=[1] And B.����ʱ��=[2] And C.��¼����=4"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "�Ƿ��������", mT_Patient.lng�ļ�ID, CDate(strDate))
            If rsTemp.RecordCount > 0 Then
                strInfo = "��ʱ���Ѿ�����������Ϣ�����飡 ʱ��[" & strDate & "]"
                GoTo ErrInfo
            End If
            
            If Not CheckDateTime(Row, "ʱ��", Format(dkpDate.Value & " " & strText, "YYYY-MM-DD HH:mm:ss")) Then Cancel = True
        End If
    End With
ErrEnd:
    '��֤ͨ���������ݱ������
    Call UpdateCurveDate(Row, Col, 2, IIf(Col = Col_OperType, True, False), True)
    Exit Sub
ErrInfo:
    lblStb.Caption = strInfo
    lblStb.ForeColor = 255
    Cancel = True
End Sub

Private Sub vsfTab_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strInfo As String
    Dim lngNO As Long, strName As String, strTmp As String, strֵ�� As String
    Dim arrStr() As String
    Dim cbrControl As CommandBarControl
    Dim blnCheck As Boolean
    Dim strText As String
    
    Call AdjustRowFlag(vsfTab, NewRow)
    
    If mblnInit = False Then Exit Sub
    If NewRow < vsfTab.FixedRows Or NewCol < vsfTab.FixedCols Then Exit Sub
    
    With vsfTab
        lngNO = Val(.TextMatrix(NewRow, COL_tab��Ŀ���))
        strTmp = .TextMatrix(NewRow, COL_tab��Ŀ����)
        If strTmp = "" Then strTmp = "("
        strName = Split(strTmp, "(")(0)
        strTmp = .TextMatrix(NewRow, COL_tab�ַ���)
        strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
        arrStr = Split(strTmp, ",")
        strֵ�� = arrStr(0)
        
        If strֵ�� = "" Then
            strInfo = ""
        Else
            strInfo = strName & "��Ч��Χ:" & strֵ��
        End If
        
        If lngNO = 4 And strName = "Ѫѹ" Then 'Ѫѹ
            strInfo = strInfo & Space(4) & "¼�����:����ѹ/����ѹ��(�����δ�⡢�ܲ⡢���)"
        End If
        
        If Val(arrStr(4)) = 4 Then strInfo = strInfo & Space(4) & "������Ŀ" & Space(4) & "¼�����:����¼��" & IIf(mbln���ܵ��� = True, "����", "����") & "�����ݡ�"
    End With
    
    lblStb.Caption = strInfo
    lblStb.ForeColor = &H80000012
    
    '��������Ƿ������޸�
    mrsCurve.Filter = "��Ŀ���=" & lngNO & " and ��Ŀ����='" & strName & "'" & _
        "   and �к�=" & NewCol - vsfTab.FixedCols + 1
    If mrsCurve.RecordCount > 0 Then
        If InStr(1, ",0,9,", "," & Val(mrsCurve!������Դ) & ",") = 0 Then
            lblStb.Caption = "������Դ�ڻ����¼����PDA�����ݲ��ܽ����޸ġ�ɾ������"
            lblStb.ForeColor = 255
            Exit Sub
        End If
    End If
    
    If NewCol < vsfTab.FixedCols + Val(arrStr(3)) Then
        'ȷ��������Һ����
        strText = Trim(vsfTab.TextMatrix(NewRow, NewCol))
        blnCheck = False
        For Each cbrControl In mcbrToolBar.Controls(5).CommandBar.Controls
            cbrControl.Checked = False
            If lngNO = gint��� Then
                Select Case cbrControl.Id
                    Case conMenu_Edit_Append * 10 + 1
                        cbrControl.Checked = (InStr(1, UCase(strText), "/E") = 0 And InStr(1, UCase(strText), "E") > 0)
                    Case conMenu_Edit_Append * 10 + 2
                        cbrControl.Checked = (InStr(1, UCase(strText), "/E") > 0)
                    Case conMenu_Edit_Append * 10 + 3
                        cbrControl.Checked = (UCase(strText) = "*" Or UCase(strText) = "��")
                    Case conMenu_Edit_Append * 10 + 4
                        cbrControl.Checked = (UCase(strText) = "��")
                End Select
               
            ElseIf lngNO = gint��Һ Then
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
            If NewCol < .FixedCols + (Split(.TextMatrix(NewRow, COL_tab�ַ���), ",")(3)) Then
                mrsCurve.Filter = "��Ŀ���=" & Val(.TextMatrix(NewRow, COL_tab��Ŀ���)) & " and ��Ŀ����='" & Split(.TextMatrix(NewRow, COL_tab��Ŀ����), "(")(0) & "'" & _
                    "   and �к�=" & NewCol - .FixedCols + 1
                If mrsCurve.RecordCount > 0 Then
                    If InStr(1, ",0,9,", "," & Val(mrsCurve!������Դ) & ",") = 0 Then
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
    Dim intType As Integer, intƵ�� As Integer
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
    '������ݺϷ���
    '--51282,������,2012-08-03,ȫ�������ʾ¼��ʱ��(DYEYҪ���ֹ�¼�����ʱ��H)
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
     
    '���ݲ��Ϸ�
    If blnAllow = False Then
        If vsfTab.Row <> intRow Then vsfTab.Row = intRow
        If vsfTab.Col <> intCOl Then vsfTab.Col = intCOl
        GoTo ErrFouce
        Exit Sub
    End If
    
    If vsfTab.Row < vsfTab.FixedRows And vsfTab.Col < vsfTab.FixedCols Then Exit Sub
    If Not vsfTab.RowIsVisible(vsfTab.Row) Then Exit Sub
    If Not mblnScroll And vsfTab.Visible Then vsfTab.SetFocus
    
    '�������б༭�ؼ�
    picδ��.Visible = False
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
        strInfo = "�������������Ѿ��鵵,��������������޸�."
        mblnEdit = False
        GoTo ErrInfo
    End If
        
    If mblnEdit = False Then Exit Sub
    
    With vsfTab
        If .Row > .FixedRows - 1 And .Col > .FixedCols - 1 And vsfTab.Col < .FixedCols + Val(Split(vsfTab.TextMatrix(vsfTab.Row, COL_tab�ַ���), ",")(3)) Then
            
            intType = Val(Split(.TextMatrix(vsfTab.Row, COL_tab�ַ���), ",")(4))
            intƵ�� = Val(Split(.TextMatrix(vsfTab.Row, COL_tab�ַ���), ",")(3))
            '���¼�����Ŀʱ���Ƿ񳬳��û����õ�ʱ�䷶Χ��¼ʱ�䷶Χ
            Call GetAnimalItemTime(.Row, .Col, strInfo)
            If strInfo <> "" Then
                mblnEdit = False
                GoTo ErrInfo
            End If
            '��鲨����Ŀ
            If IsWaveItem(Val(.TextMatrix(.Row, COL_tab��Ŀ���))) And InStr(1, Trim(.TextMatrix(.Row, .Col)), "-") <> 0 Then
                strInfo = "������ֵ�Ѿ��γɲ�����Χ�Ĳ�����Ŀ���ܽ����޸ġ�ɾ������"
                GoTo ErrInfo
            End If
            
            '���������Դ�Ƿ����Ի����¼����PDA
            mrsCurve.Filter = "��Ŀ���=" & Val(.TextMatrix(.Row, COL_tab��Ŀ���)) & " and ��Ŀ����='" & Split(.TextMatrix(.Row, COL_tab��Ŀ����), "(")(0) & "'" & _
                "   and �к�=" & .Col - .FixedCols + 1
            If mrsCurve.RecordCount > 0 Then
                If InStr(1, ",0,9,", "," & Val(mrsCurve!������Դ) & ",") = 0 Then
                    blnEdit = False
                End If
                cmdColor.Tag = Val(mrsCurve!δ��˵��)
            End If
            '--51282,������,2012-08-03,ȫ�������ʾ¼��ʱ��(DYEYҪ���ֹ�¼�����ʱ��H)
            If blnEdit = False And Not (intType = 4 And intƵ�� = 1 And mbln¼��Сʱ = True) Then
                strInfo = "������Դ�ڻ����¼����PDA�����ݲ��ܽ����޸ġ�ɾ������"
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
            '������Ŀ�������������͵Ļ��Ŀ����������������ɫ
             If Val(Split(.TextMatrix(vsfTab.Row, COL_tab�ַ���), ",")(1)) = 1 And intType = 0 And Val(Split(.TextMatrix(vsfTab.Row, COL_tab�ַ���), ",")(5)) = 2 Then
                cmdColor.Top = 0
                cmdColor.Height = picEdit.Height
                cmdColor.Width = 300
                cmdColor.Left = picEdit.Width - cmdColor.Width
                txtEdit.Width = cmdColor.Left
                cmdColor.Enabled = True
                cmdColor.Visible = True
                GoTo ShowText
            ElseIf intType = 4 And intƵ�� = 1 And mbln¼��Сʱ = True Then
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
                
                strTmp = vsfTab.TextMatrix(vsfTab.Row, COL_tab�ַ���)
                lngItemNO = Val(vsfTab.TextMatrix(vsfTab.Row, COL_tab��Ŀ���))
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
            ElseIf intType = 2 Or intType = 3 Then '��ѡ
                
                '51600,������,2012-07-16,��ѡ��Ŀ�ṩ����ѡ���¼�빦��
                strValue = Split(.TextMatrix(vsfTab.Row, COL_tab�ַ���), ",")(0)
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
                    If Left(arrValue(i), 1) = "��" Then arrValue(i) = Mid(arrValue(i), 2): strValue1 = arrValue(i)
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
                
                '�ؼ���ʾ
                '51600,������,2012-07-16,��ѡ��Ŀ�ṩ����ѡ���¼�빦��
                If intType = 0 Then
                    PicLst.FontName = .FontName
                    PicLst.FontSize = .FontSize
                    PicLst.Left = .CellLeft + .Left + 15
                    PicLst.Top = .CellTop + vsfTab.Top
                    PicLst.Height = 80 + (.CellHeight - 5) + PicLst.TextHeight("��") * 2 + lstSelect(intType).ListCount * (PicLst.TextHeight("��") + PicLst.TextHeight("��") \ 4)
                    If PicLst.Height < .CellHeight + 20 Then PicLst.Height = .CellHeight + 20
                    PicLst.Width = LenB(StrConv(lstSelect(intType).List(lstSelect(intType).ListCount \ 2), vbFromUnicode)) * 120 + 500    '���м���ĳ���Ϊ����
                    If PicLst.Width < .CellWidth + 20 Then PicLst.Width = .CellWidth + 20
                    If PicLst.Height > vsfTab.Height Then PicLst.Height = vsfTab.Height
                    If PicLst.Top + PicLst.Height > vsfTab.Height Then PicLst.Top = .CellTop + .Top + .CellHeight + 20 - PicLst.Height
                    If PicLst.Top < 0 Then PicLst.Top = vsfTab.Top
                    PicLst.Visible = True
                    PicLst.ZOrder 0
                    
                    lbllst(2).Left = 20
                    lbllst(2).Top = 20
                    If lbllst(2).Width > PicLst.Width Then
                        PicLst.Width = lbllst(2).Width + PicLst.TextWidth("��")
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
                    strTmp = vsfTab.TextMatrix(vsfTab.Row, COL_tab�ַ���)
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
                    lstSelect(intType).Height = lstSelect(intType).ListCount * (PicLst.TextHeight("��") + PicLst.TextHeight("��") \ 4)
                    If lstSelect(intType).Height < .CellHeight + 20 Then lstSelect(intType).Height = .CellHeight + 20
                    lstSelect(intType).Width = LenB(StrConv(lstSelect(intType).List(lstSelect(intType).ListCount \ 2), vbFromUnicode)) * 120 + 500    '���м���ĳ���Ϊ����
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
            ElseIf intType = 5 Then 'ѡ��
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
                strTmp = vsfTab.TextMatrix(vsfTab.Row, COL_tab�ַ���)
                lngItemNO = Val(vsfTab.TextMatrix(vsfTab.Row, COL_tab��Ŀ���))
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
        lblStb.Caption = "�������������Ѿ��鵵,��������������޸�."
        lblStb.ForeColor = 255
    End If
End Sub

Private Sub vsfTab_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    vsfTab.Cell(flexcpAlignment, Row, Col, Row, Col) = flexAlignCenterCenter
End Sub

Private Sub vsfTab_KeyDown(KeyCode As Integer, Shift As Integer)

    Dim intCols As Integer
    Dim intType As Integer, intƵ�� As Integer
    Dim blnTrue As Boolean
    Dim blnEdit As Boolean
    Dim strText As String
    
    If vsfTab.Row < vsfTab.FixedRows And vsfTab.Col < vsfTab.FixedCols Then Exit Sub
    
    '���ε�ĳЩ���ܼ�
    If KeyCode = vbKeyEscape Or KeyCode = vbKeyBack Or Shift <> 0 Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or KeyCode = vbKeyRight Then Exit Sub
    
    If KeyCode = vbKeyLeft And (picEdit.Visible = False And lstSelect(0).Visible = False And lstSelect(1).Visible = False) Then Exit Sub
    
    intCols = Val(Split(vsfTab.TextMatrix(vsfTab.Row, COL_tab�ַ���), ",")(3)) + vsfTab.FixedCols
    
    With vsfTab
        If KeyCode = vbKeyReturn Then
NextCol2: '������һ��
            If .Col < vsfTab.FixedCols Then
                .Col = .Col + 1: GoTo NextCol2
            End If
            If .Col < intCols - 1 Then
                .Col = .Col + 1
                If .ColHidden(.Col) = True Then GoTo NextCol2
            Else
NextRow2: '������һ��
                If .Row < .Rows - 1 Then
                    .Col = vsfTab.FixedCols: .Row = .Row + 1
                    If .RowHidden(.Row) = True Then GoTo NextRow2
                Else
                    Call txtEdit_KeyPress(vbKeyEscape)
                    .Row = .FixedRows
                    .Col = .FixedCols
                End If
            End If
            '������л��в��ɼ����Զ���ʾ����
            If .ColIsVisible(.Col) = False Then
                .LeftCol = .Col
            End If
            If .RowIsVisible(.Row) = False Then
                .TopRow = .Row
            End If
        
            Exit Sub
        End If
        '���
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
                    .Col = Val(Split(vsfTab.TextMatrix(vsfTab.Row, COL_tab�ַ���), ",")(3)) + vsfTab.FixedCols
                    GoTo PreCol2
                End If
            End If
            '������л��в��ɼ����Զ���ʾ����
            If .ColIsVisible(.Col) = False Then
                .LeftCol = .Col
            End If
            If .RowIsVisible(.Row) = False Then
                .TopRow = .Row
            End If
            Exit Sub
        End If
        
        'ɾ����Ϣ
        If KeyCode = vbKeyDelete Then
            If Shift = 0 And .Col > .FixedCols - 1 And .Col < intCols Then
                blnEdit = True
                If .TextMatrix(.Row, .Col) <> "" Then
                    '�����Ŀ�Ƿ��ǲ�����Ŀ
                    If IsWaveItem(Val(.TextMatrix(.Row, COL_tab��Ŀ���))) And InStr(1, Trim(.TextMatrix(.Row, .Col)), "-") <> 0 Then
                        lblStb.Caption = "������ֵ�Ѿ��γɲ�����Χ�Ĳ�����Ŀ���ܽ����޸ġ�ɾ������"
                        lblStb.ForeColor = 255
                        GoTo ErrExit
                    End If
                    '���������Դ�Ƿ����Ի����¼����PDA
                    mrsCurve.Filter = "��Ŀ���=" & Val(.TextMatrix(.Row, COL_tab��Ŀ���)) & " and ��Ŀ����='" & Split(.TextMatrix(.Row, COL_tab��Ŀ����), "(")(0) & "'" & _
                        "   and �к�=" & .Col - .FixedCols + 1
                    If mrsCurve.RecordCount > 0 Then
                        If InStr(1, ",0,9,", "," & Val(mrsCurve!������Դ) & ",") = 0 Then
                            blnEdit = False
                        End If
                    End If
                    intƵ�� = Val(Split(vsfTab.TextMatrix(vsfTab.Row, COL_tab�ַ���), ",")(3))
                    intType = Val(Split(vsfTab.TextMatrix(vsfTab.Row, COL_tab�ַ���), ",")(4))
                    If blnEdit = False And Not (intType = 4 And intƵ�� = 1 And mbln¼��Сʱ = True) Then
                        lblStb.Caption = "������Դ�ڻ����¼����PDA�����ݲ��ܽ����޸ġ�ɾ������"
                        lblStb.ForeColor = 255
                        GoTo ErrExit
                    End If
                    picTab.Tag = .Row & "|" & .Col
                    FraTable.Tag = .TextMatrix(.Row, .Col)
                    strText = ""
                    If blnEdit = False Then '������ȫ�������Ŀ������mbln¼��Сʱ=true
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
'����:���������޸ı���
'--------------------------------------------------------
    Dim strSQL As String, arrSQL() As String, arrSQLTime() As String
    Dim strTime As String, strEnd As String, strMarkTime As String, strOldTime As String
    Dim lngItemCode As Long, strValue As String, strδ�� As String, strTmp As String
    Dim arrTmp() As String
    Dim intModify As Integer
    Dim blnEdit As Boolean
    Dim blnSave As Boolean
    Dim strName As String, strInfo As String
    Dim lngRow As Long, lng��¼ID As Long, lngOldID As Long
    Dim i As Integer, int��Ŀ�״� As Integer
    Dim blnTran As Boolean
    Dim int������ As Integer
    
    On Error GoTo Errhand
    
    mrsCurve.Filter = 0
    mrsCurve.Sort = "ʱ��,��Ŀ���"
    Screen.MousePointer = 11
    
    ReDim Preserve arrSQL(1 To 1)
    ReDim Preserve arrSQLTime(1 To 1)
    
    mrsRecodeID.Filter = 0
    '�������ݱ���
    With mrsCurve
        Do While Not .EOF
            lngItemCode = Val(!��Ŀ���)
            strValue = Nvl(!��ֵ)
            '�����:53505,�޸��ˣ�����,Ѫѹ����¼������:�����δ���.
            If lngItemCode = 4 And zlCommFun.Nvl(!��Ŀ����) = "Ѫѹ" And Nvl(!��ֵ) <> "" Then
                strValue = Nvl(!��ֵ) & "/" & Nvl(!��ֵ)
            End If
            intModify = Val(zlCommFun.Nvl(!�޸�))
            blnEdit = False
            If intModify = 1 And InStr(1, ",0,9,", Val(zlCommFun.Nvl(!������Դ))) = 0 Then
                blnEdit = False
            Else
                blnEdit = True
            End If
            blnSave = False
            If Val(!״̬) <> 3 And Val(!״̬) <> 0 Then
               '����������Ŀ����
                If !������ = "1)����������Ŀ" Then
                    strTime = !ʱ��
                    strOldTime = Trim(zlCommFun.Nvl(!ԭʼʱ��))
                    If strTime = "" Then
                        'ʱ��Ϊ�վ���ȡ����ʱ����е�ʱ��
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
                    int������ = IIf(ISCheckDept(strMarkTime) = True, 1, 0)
                    strMarkTime = "To_Date('" & strMarkTime & "','yyyy-mm-dd hh24:mi:ss')"
                    
                    '�����޸Ĳ��˻������ݷ���ʱ��
                    If strOldTime <> strTime And strOldTime <> "" Then
                        mrsRecodeID.Filter = "ʱ��='" & strOldTime & "'"
                        If mrsRecodeID.RecordCount > 0 Then
                            lng��¼ID = Val(mrsRecodeID!��¼ID)
                            
                            '��ͬ��¼�޸Ĺ����ٴν����޸�
                            If lng��¼ID <> lngOldID Then
                                strSQL = "ZL_���µ�����_����ʱ��("
                                'ID_IN       IN ���˻�������.ID%TYPE,
                                strSQL = strSQL & lng��¼ID & ","
                                '����ʱ��_IN IN ���˻�������.����ʱ��%TYPE
                                strSQL = strSQL & strMarkTime & ")"
                                
                                arrSQLTime(ReDimArray(arrSQLTime)) = strSQL
                            End If
                        End If
                    End If
                    
                    lngOldID = lng��¼ID
                    
                    If strValue = "����" And lngItemCode = Item���� Then
                        strδ�� = ""
                    Else
                        strδ�� = !δ��˵��
                    End If
                    
                    '״̬=4ֻ�Ƕ�ʱ��������޸�(�����Ѿ�����)
                    If Val(!״̬) <> 4 Then
                        '����������Ϣ
                        strSQL = "Zl_���µ�����_Update("
                        '�ļ�id_In   In ���˻����ļ�.Id%Type,  --���˻����ļ�ID
                        strSQL = strSQL & Val(mT_Patient.lng�ļ�ID) & ","
                        '����ʱ��_In In ���˻�������.����ʱ��%Type, --�������ݵķ���ʱ��
                        strSQL = strSQL & strMarkTime & ","
                        '��¼����_In In ���˻�����ϸ.��¼����%Type, --������Ŀ=1���ϱ�˵��=2�����ת���=3�������ձ��=4,�±�˵��=6
                        strSQL = strSQL & "1,"
                        '��Ŀ���_In In ���˻�����ϸ.��Ŀ���%Type, --������Ŀ����ţ��ǻ�����Ŀ�̶�Ϊ0
                        strSQL = strSQL & lngItemCode & ","
                        '��¼����_In In ���˻�����ϸ.��¼����%Type := Null, --��¼���ݣ��������Ϊ�գ��������ǰ������  36��36/37
                        strSQL = strSQL & "'" & strValue & "',"
                        '���²�λ_In In ���˻�����ϸ.���²�λ%Type := Null, --ɾ������ʱ������д��λ �����Ŀ��
                        strSQL = strSQL & IIf(strValue <> "", "'" & Nvl(!��λ) & "'", "NULL") & ","
                        '���Ժϸ�_In In Number := 0,
                        strSQL = strSQL & IIf(lngItemCode = Item���� And strValue <> "", Val(!����), "0") & ","
                        'δ��˵��_In In ���˻�����ϸ.δ��˵��%Type := Null, --δ��˵��
                        strSQL = strSQL & "'" & strδ�� & "',"
                        '���˼�¼_In In Number := 1,
                        strSQL = strSQL & "1,"
                        '������Դ_In In ���˻�����ϸ.������Դ%Type := 0,
                        strSQL = strSQL & "0,"
                        '��Դid_In   In ���˻�����ϸ.��Դid%Type := Null,
                        strSQL = strSQL & IIf(Val(!��ԴID) = 0, "NULL", !��ԴID) & ","
                        '����_In     In ���˻�����ϸ.����%Type := 0,
                        strSQL = strSQL & Val(!����)
                        '  ��Ŀ�״�_In In Number := 0,--������Ŀʹ�ã���������ǰ�Ƿ���ɾ��һ��ʱ���ڵ�������Ϣ�� 1 ɾ��
                        '  ��ʼʱ��_In In ���˻�������.����ʱ��%Type := Null, --����¼��Ч��ȵĿ�ʼʱ��
                        '  ����ʱ��_In In ���˻�������.����ʱ��%Type := Null, --����¼��Ч��ȵ���ֹʱ�䣬������¼Ϊÿ���ӣ����±�Ϊ4Сʱ,ʱ�����ڵ���ͬ��Ŀ��¼Ҫɾ��
                        '  ����Ա_IN  IN ���˻�������.������%TYPE := NULL,
                        '  ������_IN IN Number :=1
                        strSQL = strSQL & ",0,NULL,NULL,NULL,"
                        strSQL = strSQL & int������ & ")"
                        
                        arrSQL(ReDimArray(arrSQL)) = strSQL
                    End If
                '���±����Ŀ����
                ElseIf !������ = "2)���±����Ŀ" Then
                    int��Ŀ�״� = 0
                    strName = zlCommFun.Nvl(!��Ŀ����)
                    strTmp = GetItemInfo(lngItemCode, strName, lngRow)
                    strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
                    arrTmp = Split(strTmp, ",")
                    
                    strTime = Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")
                    strEnd = strTime
                    strMarkTime = strTime

                    '���ڿ���¼��Ļ�����Ŀ,��Ҫ���ݻ���ʱ��ɾ����ʱ���ڵ���������
                    If Val(arrTmp(4)) = 4 Then
                        strTmp = GetAnimalItemTime(lngRow, !�к� + vsfCurve.FixedCols - 1, strInfo, 1)
                        If strInfo <> "" Then Exit Function
                        strTime = Split(strTmp, ";")(0)
                        strEnd = Split(strTmp, ";")(1)
                        If CDate(strMarkTime) < CDate(mstrBTime) Then strMarkTime = Format(mstrBTime, "YYYY-MM-DD HH:mm:ss")
                        If CDate(strMarkTime) > CDate(mstrETime) Then strMarkTime = Format(mstrETime, "YYYY-MM-DD HH:mm:ss")
                        int��Ŀ�״� = 1
                    End If
                    int������ = IIf(ISCheckDept(strMarkTime) = True, 1, 0)
                    strMarkTime = "To_Date('" & strMarkTime & "','yyyy-mm-dd hh24:mi:ss')"
                    
                    '����������Ϣ
                    strSQL = "Zl_���µ�����_Update("
                    '�ļ�id_In   In ���˻����ļ�.Id%Type,  --���˻����ļ�ID
                    strSQL = strSQL & Val(mT_Patient.lng�ļ�ID) & ","
                    '����ʱ��_In In ���˻�������.����ʱ��%Type, --�������ݵķ���ʱ��
                    strSQL = strSQL & strMarkTime & ","
                    '��¼����_In In ���˻�����ϸ.��¼����%Type, --������Ŀ=1���ϱ�˵��=2�����ת���=3�������ձ��=4,�±�˵��=6
                    strSQL = strSQL & Val(Nvl(!��¼����, 1)) & ","
                    '��Ŀ���_In In ���˻�����ϸ.��Ŀ���%Type, --������Ŀ����ţ��ǻ�����Ŀ�̶�Ϊ0
                    strSQL = strSQL & lngItemCode & ","
                    '��¼����_In In ���˻�����ϸ.��¼����%Type := Null, --��¼���ݣ��������Ϊ�գ��������ǰ������  36��36/37
                    strSQL = strSQL & "'" & strValue & "',"
                    '���²�λ_In In ���˻�����ϸ.���²�λ%Type := Null, --ɾ������ʱ������д��λ �����Ŀ��
                    strSQL = strSQL & IIf(Val(arrTmp(5)) = 2, "'" & Nvl(!��λ) & "'", "NULL") & ","
                    '���Ժϸ�_In In Number := 0,
                    strSQL = strSQL & IIf(lngItemCode = Item���� And strValue <> "", Val(!����), "0") & ","
                    'δ��˵��_In In ���˻�����ϸ.δ��˵��%Type := Null, --δ��˵��
                    If Val(arrTmp(1)) = 1 And Val(arrTmp(5)) = 2 Then
                        strSQL = strSQL & "'" & IIf(strValue = "", "", Val(!δ��˵��)) & "',"
                    Else
                        strSQL = strSQL & "NUll,"
                    End If
                    '���˼�¼_In In Number := 1,
                    strSQL = strSQL & "1,"
                    '������Դ_In In ���˻�����ϸ.������Դ%Type := 0,
                    strSQL = strSQL & Val(!������Դ) & ","
                    '��Դid_In   In ���˻�����ϸ.��Դid%Type := Null,
                    strSQL = strSQL & IIf(Val(!��ԴID) = 0, "NULL", !��ԴID) & ","
                    '����_In     In ���˻�����ϸ.����%Type := 0,
                    strSQL = strSQL & Val(!����) & ","
                    '��Ŀ�״�_In In Number := 0,--������Ŀʹ�ã���������ǰ�Ƿ���ɾ��һ��ʱ���ڵ�������Ϣ�� 1 ɾ��
                    strSQL = strSQL & int��Ŀ�״� & ","
                    '��ʼʱ��_In In ���˻�������.����ʱ��%Type := Null,
                    strSQL = strSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'),"
                    '����ʱ��_In In ���˻�������.����ʱ��%Type := Null --����¼��Ч��ȵ���ֹʱ�䣬������¼Ϊÿ���ӣ����±�Ϊ4Сʱ,ʱ�����ڵ���ͬ��Ŀ��¼Ҫɾ��
                    strSQL = strSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss')"
                    '  ����Ա_IN  IN ���˻�������.������%TYPE := NULL,
                    '  ������_IN IN Number :=1
                    strSQL = strSQL & ",NULL," & int������ & ")"
                    
                    arrSQL(ReDimArray(arrSQL)) = strSQL
                End If
            End If
        .MoveNext
        Loop
    End With
    
    
    '�������������±�˵����Ϣ
    mrsOper.Filter = 0
    '��ɾ�����޸ĵ�������Ϣ,һ��������ö���������������ʱ�����������ʱ����ͬ����������ʱ��Ļ���Ҳ�ᵶ�Ӻ���������ʱ�䷢���仯
    mrsOper.Sort = "ʱ��"
    With mrsOper
        Do While Not .EOF
            If Val(!״̬) <> 3 And Val(!״̬) <> 0 Then
                lngItemCode = 4
                If Val(!״̬) = 2 Then
                    strTime = Format(!ԭʼʱ��, "YYYY-MM-DD HH:mm:ss")
                    strEnd = strTime
                    strMarkTime = strTime
                    int������ = IIf(ISCheckDept(strMarkTime) = True, 1, 0)
                    strMarkTime = "To_Date('" & strMarkTime & "','yyyy-mm-dd hh24:mi:ss')"
                    
                    '����������Ϣ
                    strSQL = "Zl_���µ�����_Update("
                    '�ļ�id_In   In ���˻����ļ�.Id%Type,  --���˻����ļ�ID
                    strSQL = strSQL & Val(mT_Patient.lng�ļ�ID) & ","
                    '����ʱ��_In In ���˻�������.����ʱ��%Type, --�������ݵķ���ʱ��
                    strSQL = strSQL & strMarkTime & ","
                    '��¼����_In In ���˻�����ϸ.��¼����%Type, --������Ŀ=1���ϱ�˵��=2�����ת���=3�������ձ��=4,�±�˵��=6
                    strSQL = strSQL & lngItemCode & ","
                    '��Ŀ���_In In ���˻�����ϸ.��Ŀ���%Type, --������Ŀ����ţ��ǻ�����Ŀ�̶�Ϊ0
                    strSQL = strSQL & 0 & ","
                    '��¼����_In In ���˻�����ϸ.��¼����%Type := Null, --��¼���ݣ��������Ϊ�գ��������ǰ������  36��36/37
                    strSQL = strSQL & "NULL" & ","
                    '���²�λ_In In ���˻�����ϸ.���²�λ%Type := Null, --ɾ������ʱ������д��λ �����Ŀ��
                    strSQL = strSQL & "NULL,"
                    '���Ժϸ�_In In Number := 0,
                    strSQL = strSQL & "NULL,"
                    'δ��˵��_In In ���˻�����ϸ.δ��˵��%Type := Null, --δ��˵��
                    strSQL = strSQL & "NULL" & ","
                    '���˼�¼_In In Number := 1,
                    strSQL = strSQL & "1,"
                    '������Դ_In In ���˻�����ϸ.������Դ%Type := 0,
                    strSQL = strSQL & Val(!������Դ) & ","
                    '��Դid_In   In ���˻�����ϸ.��Դid%Type := Null,
                    strSQL = strSQL & IIf(Val(!��ԴID) = 0, "NULL", !��ԴID) & ","
                    '����_In     In ���˻�����ϸ.����%Type := 0,
                    strSQL = strSQL & Val(!����) & ","
                    '��Ŀ�״�_In In Number := 0,--������Ŀʹ�ã���������ǰ�Ƿ���ɾ��һ��ʱ���ڵ�������Ϣ�� 1 ɾ��
                    strSQL = strSQL & 0 & ","
                    '��ʼʱ��_In In ���˻�������.����ʱ��%Type := Null,
                    strSQL = strSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'),"
                    '����ʱ��_In In ���˻�������.����ʱ��%Type := Null --����¼��Ч��ȵ���ֹʱ�䣬������¼Ϊÿ���ӣ����±�Ϊ4Сʱ,ʱ�����ڵ���ͬ��Ŀ��¼Ҫɾ��
                    strSQL = strSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss')"
                    '  ����Ա_IN  IN ���˻�������.������%TYPE := NULL,
                    '  ������_IN IN Number :=1
                    strSQL = strSQL & ",NULL," & int������ & ")"
                    arrSQL(ReDimArray(arrSQL)) = strSQL
                End If
                
                strTime = Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")
                strEnd = strTime
                strMarkTime = strTime
                int������ = IIf(ISCheckDept(strMarkTime) = True, 1, 0)
                strMarkTime = "To_Date('" & strMarkTime & "','yyyy-mm-dd hh24:mi:ss')"
                strValue = Trim(zlCommFun.Nvl(!����))
                If strValue <> "" Then
                    '����������Ϣ
                    strSQL = "Zl_���µ�����_Update("
                    '�ļ�id_In   In ���˻����ļ�.Id%Type,  --���˻����ļ�ID
                    strSQL = strSQL & Val(mT_Patient.lng�ļ�ID) & ","
                    '����ʱ��_In In ���˻�������.����ʱ��%Type, --�������ݵķ���ʱ��
                    strSQL = strSQL & strMarkTime & ","
                    '��¼����_In In ���˻�����ϸ.��¼����%Type, --������Ŀ=1���ϱ�˵��=2�����ת���=3�������ձ��=4,�±�˵��=6
                    strSQL = strSQL & lngItemCode & ","
                    '��Ŀ���_In In ���˻�����ϸ.��Ŀ���%Type, --������Ŀ����ţ��ǻ�����Ŀ�̶�Ϊ0
                    strSQL = strSQL & 0 & ","
                    '��¼����_In In ���˻�����ϸ.��¼����%Type := Null, --��¼���ݣ��������Ϊ�գ��������ǰ������  36��36/37
                    strSQL = strSQL & "'" & strValue & "',"
                    '���²�λ_In In ���˻�����ϸ.���²�λ%Type := Null, --ɾ������ʱ������д��λ �����Ŀ��
                    strSQL = strSQL & "NULL,"
                    '���Ժϸ�_In In Number := 0,
                    strSQL = strSQL & "NULL,"
                    'δ��˵��_In In ���˻�����ϸ.δ��˵��%Type := Null, --δ��˵��
                    strSQL = strSQL & IIf(lngItemCode <> 4, "'" & Nvl(!δ��˵��) & "'", "NULL") & ","
                    '���˼�¼_In In Number := 1,
                    strSQL = strSQL & "1,"
                    '������Դ_In In ���˻�����ϸ.������Դ%Type := 0,
                    strSQL = strSQL & Val(!������Դ) & ","
                    '��Դid_In   In ���˻�����ϸ.��Դid%Type := Null,
                    strSQL = strSQL & IIf(Val(!��ԴID) = 0, "NULL", !��ԴID) & ","
                    '����_In     In ���˻�����ϸ.����%Type := 0,
                    strSQL = strSQL & Val(!����) & ","
                    '��Ŀ�״�_In In Number := 0,--������Ŀʹ�ã���������ǰ�Ƿ���ɾ��һ��ʱ���ڵ�������Ϣ�� 1 ɾ��
                    strSQL = strSQL & int��Ŀ�״� & ","
                    '��ʼʱ��_In In ���˻�������.����ʱ��%Type := Null,
                    strSQL = strSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'),"
                    '����ʱ��_In In ���˻�������.����ʱ��%Type := Null --����¼��Ч��ȵ���ֹʱ�䣬������¼Ϊÿ���ӣ����±�Ϊ4Сʱ,ʱ�����ڵ���ͬ��Ŀ��¼Ҫɾ��
                    strSQL = strSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss')"
                    '  ����Ա_IN  IN ���˻�������.������%TYPE := NULL,
                    '  ������_IN IN Number :=1
                    strSQL = strSQL & ",NULL," & int������ & ")"
                    arrSQL(ReDimArray(arrSQL)) = strSQL
                End If
            End If
        .MoveNext
        Loop
    End With
    
    '�±���Ϣ
    mrsNote.Filter = 0
    mrsNote.Sort = "ʱ��"
    With mrsNote
        Do While Not .EOF
        lngItemCode = Val(!��¼����)
        
        If Val(!״̬) <> 3 And Val(!״̬) <> 0 Then
            strTime = Format(mstrBegin, "YYYY-MM-DD HH:mm:ss")
            strEnd = Format(mstrEnd, "YYYY-MM-DD HH:mm:ss")
            strMarkTime = Format(!ʱ��, "YYYY-MM-DD HH:mm:ss")
            strValue = zlCommFun.Nvl(!����)
            int��Ŀ�״� = 1
            int������ = IIf(ISCheckDept(strMarkTime) = True, 1, 0)
            strMarkTime = "To_Date('" & strMarkTime & "','yyyy-mm-dd hh24:mi:ss')"
            
             '����������Ϣ
            strSQL = "Zl_���µ�����_Update("
            '�ļ�id_In   In ���˻����ļ�.Id%Type,  --���˻����ļ�ID
            strSQL = strSQL & Val(mT_Patient.lng�ļ�ID) & ","
            '����ʱ��_In In ���˻�������.����ʱ��%Type, --�������ݵķ���ʱ��
            strSQL = strSQL & strMarkTime & ","
            '��¼����_In In ���˻�����ϸ.��¼����%Type, --������Ŀ=1���ϱ�˵��=2�����ת���=3�������ձ��=4,�±�˵��=6
            strSQL = strSQL & lngItemCode & ","
            '��Ŀ���_In In ���˻�����ϸ.��Ŀ���%Type, --������Ŀ����ţ��ǻ�����Ŀ�̶�Ϊ0
            strSQL = strSQL & 0 & ","
            '��¼����_In In ���˻�����ϸ.��¼����%Type := Null, --��¼���ݣ��������Ϊ�գ��������ǰ������  36��36/37
            strSQL = strSQL & "'" & strValue & "',"
            '���²�λ_In In ���˻�����ϸ.���²�λ%Type := Null, --ɾ������ʱ������д��λ �����Ŀ��
            strSQL = strSQL & "NULL,"
            '���Ժϸ�_In In Number := 0,
            strSQL = strSQL & "NULL,"
            'δ��˵��_In In ���˻�����ϸ.δ��˵��%Type := Null, --δ��˵��
            strSQL = strSQL & IIf(lngItemCode <> 4, "'" & Nvl(!δ��˵��) & "'", "NULL") & ","
            '���˼�¼_In In Number := 1,
            strSQL = strSQL & "1,"
            '������Դ_In In ���˻�����ϸ.������Դ%Type := 0,
            strSQL = strSQL & Val(!������Դ) & ","
            '��Դid_In   In ���˻�����ϸ.��Դid%Type := Null,
            strSQL = strSQL & IIf(Val(!��ԴID) = 0, "NULL", !��ԴID) & ","
            '����_In     In ���˻�����ϸ.����%Type := 0,
            strSQL = strSQL & Val(!����) & ","
            '��Ŀ�״�_In In Number := 0,--������Ŀʹ�ã���������ǰ�Ƿ���ɾ��һ��ʱ���ڵ�������Ϣ�� 1 ɾ��
            strSQL = strSQL & int��Ŀ�״� & ","
            '��ʼʱ��_In In ���˻�������.����ʱ��%Type := Null,
            strSQL = strSQL & "To_Date('" & strTime & "','yyyy-mm-dd hh24:mi:ss'),"
            '����ʱ��_In In ���˻�������.����ʱ��%Type := Null --����¼��Ч��ȵ���ֹʱ�䣬������¼Ϊÿ���ӣ����±�Ϊ4Сʱ,ʱ�����ڵ���ͬ��Ŀ��¼Ҫɾ��
            strSQL = strSQL & "To_Date('" & strEnd & "','yyyy-mm-dd hh24:mi:ss')"
            '  ����Ա_IN  IN ���˻�������.������%TYPE := NULL,
            '  ������_IN IN Number :=1
            strSQL = strSQL & ",NULL," & int������ & ")"
            arrSQL(ReDimArray(arrSQL)) = strSQL
        End If
        .MoveNext
        Loop
    End With
    
     '------------------------------------------------------------------------------------------------------------------
    'ѭ��ִ��SQL��������
    'Debug.Print "--�������ݿ�ʼ:" & Now
     
    gcnOracle.BeginTrans
    blnTran = True
    '��ִ��ʱ��仯
    For i = 1 To UBound(arrSQLTime)
        If arrSQLTime(i) <> "" Then Call zlDatabase.ExecuteProcedure(CStr(arrSQLTime(i)), "������������"):  'Debug.Print CStr(arrSQLTime(i))
    Next
    '��ִ�����ݱ仯
    For i = 1 To UBound(arrSQL)
        If arrSQL(i) <> "" Then Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "������������"):  'Debug.Print CStr(arrSQL(i))
    Next
    gcnOracle.CommitTrans
    
    'Debug.Print "--�������ݽ���:" & Now
     
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

Private Function ISCheckDept(ByVal str����ʱ�� As String) As Boolean
'���ܣ��Ƿ���Zl_���µ�����_Update�н��п��Ҽ��
    'mstrOverDate<=mstrETime ���Ҳ����Ѿ���Ժ���϶��ǲ��˳�Ժʱ�����Ժʱ����һ�У��������Ľ����
    If mbln��Ժ = True And Format(mstrOverDate, "YYYY-MM-DD HH:mm:ss") < Format(mstrETime, "YYYY-MM-DD HH:mm:ss") Then
        If Format(str����ʱ��, "YYYY-MM-DD HH:mm:ss") > Format(mstrOverDate, "YYYY-MM-DD HH:mm:ss") And Format(str����ʱ��, "YYYY-MM-DD HH:mm:ss") <= Format(mstrETime, "YYYY-MM-DD HH:mm:ss") Then
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
'����:��ȡ��Ŀ��Ϣ
'---------------------------------------------------------------
    Dim intRow As Integer
    Dim strValue As String
    
    For intRow = vsfTab.FixedRows To vsfTab.Rows - 1
        If Val(vsfTab.TextMatrix(intRow, COL_tab��Ŀ���)) = lngItemNO And Split(vsfTab.TextMatrix(intRow, COL_tab��Ŀ����), "(")(0) = strName Then
            Exit For
        End If
    Next intRow
    
    If intRow >= vsfTab.Rows Then
        For intRow = vsfTab.FixedRows To vsfTab.Rows - 1
            If Val(vsfTab.TextMatrix(intRow, COL_tab��Ŀ���)) = lngItemNO Then
                Exit For
            End If
        Next intRow
    End If
    
    If intRow < vsfTab.Rows Then
        strValue = vsfTab.TextMatrix(intRow, COL_tab�ַ���)
    End If
    lngRow = intRow
    GetItemInfo = strValue
End Function

Private Function WriteIntoVfgTab(ByVal strText As String, Optional blnDelete As Boolean = False, Optional ByVal blnVisible As Boolean = True, Optional strErrMsg As String = "") As Boolean
'-------------------------------------------------------------------------
'����:�û��༭������д��vsfTab
'����:strtext �༭���ı���Ϣ   blndelete �Ƿ���VsfTab��Delete ��ɾ����Ϣ
'-------------------------------------------------------------------------
    Dim intRow As Integer, intCOl As Integer
    Dim lng��Ŀ��� As Long, str��Ŀ���� As String, strTmp As String, strPart As String
    Dim arrStr() As String
    Dim strֵ�� As String, intType As Integer, intNum As Integer, lngLen As Long, intƵ�� As Integer, int���� As Integer, int��ʾ As Integer
    Dim lngColor As String
    Dim blnAllow As Boolean, blnTrue As Boolean
    Dim strValue As String, strHour As String, strHourOld As String
    Dim intIndex As Integer, int��¼���� As Integer
    Dim strTime As String
    
    '--�����޸���Ϣ
    Dim int״̬ As Integer
    On Error GoTo Errhand
    
    If Not blnDelete Then
        If picEdit.Visible = True And txtEdit.Tag <> "" Then
            intRow = Split(txtEdit.Tag, "|")(0)
            intCOl = Split(txtEdit.Tag, "|")(1)
            If txtEdit.Visible = True Or lblCheck.Visible = True Then
                strTmp = vsfTab.TextMatrix(intRow, COL_tab�ַ���)
                lng��Ŀ��� = Val(vsfTab.TextMatrix(intRow, COL_tab��Ŀ���))
                str��Ŀ���� = Split(vsfTab.TextMatrix(intRow, COL_tab��Ŀ����), "(")(0)
                strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
                arrStr = Split(strTmp, ",")
                strֵ�� = arrStr(0)
                intType = Val(arrStr(1))
                intNum = Val(arrStr(2))
                intƵ�� = Val(arrStr(3))
                int��ʾ = Val(arrStr(4))
                int���� = Val(arrStr(5))
                lngLen = Val(arrStr(6))
                strPart = arrStr(7)
                
                If intType = 1 Then strֵ�� = ""
                'ȫ�������Ŀ�����Ҳ�����ȫ�������ʾ¼��ʱ�䡱��ѡ
                If int��ʾ = 4 And intƵ�� = 1 And mbln¼��Сʱ = True Then
                    If InStr(1, strText, ")") > 0 Then
                        strHour = Replace(Replace(Split(strText, ")")(0), "(", ""), "h", "")
                        If strHour <> "" Then
                            If Not (Val(strHour) >= 0 And strHour <= 24) Then
                                lblStb.Caption = "����Сʱֻ����0��24֮�䣬������¼�룡": lblStb.ForeColor = 255
                                Exit Function
                            End If
                            strHour = "(" & strHour & "h)"
                        End If
                        strText = Split(strText, ")")(1)
                        If Trim(strText) = "" Then strHour = ""
                    End If
                End If
                If txtEdit.Enabled = True Then
                    blnAllow = CheckValidata(intRow, intCOl, lng��Ŀ���, intType, intNum, strֵ��, int��ʾ, lngLen, strText, strErrMsg)
                End If
            End If
            strValue = Split(IIf(Trim(picEdit.Tag) = "", "/#$&/", Trim(picEdit.Tag)), "/#$&/")(0)
        ElseIf lstSelect(0).Visible = True Or lstSelect(1).Visible = True Then
            If lstSelect(0).Visible = True Then strValue = lstSelect(0).Tag: intIndex = 0
            If lstSelect(1).Visible = True Then strValue = lstSelect(1).Tag: intIndex = 1
            intRow = Split(lbllst(intIndex).Tag, "|")(0)
            intCOl = Split(lbllst(intIndex).Tag, "|")(1)
            lng��Ŀ��� = Val(vsfTab.TextMatrix(intRow, COL_tab��Ŀ���))
            str��Ŀ���� = Split(vsfTab.TextMatrix(intRow, COL_tab��Ŀ����), "(")(0)
            strTmp = vsfTab.TextMatrix(intRow, COL_tab�ַ���)
            strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
            arrStr = Split(strTmp, ",")
            intType = Val(arrStr(1))
            int���� = Val(arrStr(5))
            strPart = arrStr(7)
            
            blnAllow = True
        End If
    Else
        blnAllow = True
        If InStr(1, picTab.Tag, "|") = 0 Then Exit Function
        intRow = Split(picTab.Tag, "|")(0)
        intCOl = Split(picTab.Tag, "|")(1)
        lng��Ŀ��� = Val(vsfTab.TextMatrix(intRow, COL_tab��Ŀ���))
        str��Ŀ���� = Split(vsfTab.TextMatrix(intRow, COL_tab��Ŀ����), "(")(0)
        strTmp = vsfTab.TextMatrix(intRow, COL_tab�ַ���)
        strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
        arrStr = Split(strTmp, ",")
        intType = Val(arrStr(1))
        intƵ�� = Val(arrStr(3))
        int��ʾ = Val(arrStr(4))
        int���� = Val(arrStr(5))
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
    int��¼���� = 1
    blnTrue = False
    '���������޸ı�־
    If blnAllow = True Then
        '--51282,������,2012-08-03,ȫ�������ʾ¼��ʱ��(DYEYҪ���ֹ�¼�����ʱ��H)
        strHour = Replace(Replace(strHour, "(", ""), "h)", "")
        If int��ʾ = 4 And intƵ�� = 1 And mbln¼��Сʱ = True Then
            If InStr(1, strValue, ")") > 0 Then
                strHourOld = Replace(Replace(Split(strValue, ")")(0), "(", ""), "h", "")
                strValue = Split(strValue, ")")(1)
            End If
            '�����û�¼��Ļ���Сʱ
            If Val(strHour) <> Val(strHourOld) Then
                blnTrue = True
                int��¼���� = 11
                GoTo ErrHour
            End If
        End If
ErrBegin:
        If strValue <> strText Then
ErrHour:
            mrsCurve.Filter = "��Ŀ���=" & lng��Ŀ��� & " and ��Ŀ����='" & str��Ŀ���� & "' And ��¼����=" & int��¼���� & " And �к�=" & intCOl - vsfTab.FixedCols + 1
            'Call OutputRsData(mrsCurve, True)
            If mrsCurve.RecordCount > 0 Then
                mrsCurve!δ��˵�� = lngColor
                If mrsCurve!״̬ <> 1 And mrsCurve!״̬ <> 3 Then 'ԭ�е����� �޸ġ�ɾ�����״̬ʼ��Ϊ2
                    mrsCurve!״̬ = 2
                    mrsCurve!��ֵ = IIf(blnTrue = True, strHour, strText)
                Else '�����������ݵĴ���
                    If Trim(IIf(blnTrue = True, strHour, strText)) = "" Then
                        mrsCurve!״̬ = 3
                        mrsCurve!��ֵ = ""
                    Else
                        mrsCurve!״̬ = 1
                        mrsCurve!��ֵ = IIf(blnTrue = True, strHour, strText)
                    End If
                End If
                mrsCurve.Update
            Else '�����ڼ�¼����������
                If Trim(strText) <> "" Then
                    strTime = GetAnimalItemTime(intRow, intCOl, strErrMsg)
                    If strErrMsg <> "" Then GoTo ErrInfo

                    gstrFields = "���|������|��ֵ|��λ|���|ʱ��|��Ŀ���|��Ŀ����|����|δ��˵��|������Դ|�޸�|��ʾ|��ԴID|����|״̬|�к�|��¼����"
                    gstrValues = GetMaxID & "|2)���±����Ŀ|" & IIf(blnTrue = True, strHour, strText) & "|" & strPart & "|" & _
                        0 & "|" & strTime & "|" & lng��Ŀ��� & "|" & str��Ŀ���� & "|0|" & lngColor & "|0|0|0|0|0|1|" & intCOl - vsfTab.FixedCols + 1 & "|" & int��¼����
                    Call Record_Add(mrsCurve, gstrFields, gstrValues)
                End If
            End If
            mblnChage = True
            If blnTrue = True Then
                blnTrue = False
                int��¼���� = 1
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
'����:��ȡ���±����ĿĳƵ�ε�ʱ��
'arrTime ������Ϣ ���� ��ʼʱ�� �е�ʱ�� ����ʱ��
'IntMode 0 �����м��ʱ�� 1,���ؿ�ʼʱ��ͽ���ʱ�� 2 ���ؿ�ʼʱ��;�м��ʱ��;����ʱ��
'---------------------------------------------------------------------------------
    Dim strTmp As String, lng��Ŀ��� As Long, str��Ŀ���� As String, intƵ�� As Integer, _
        int��Ŀ��ʾ As String, intType As Integer, intNO As Integer
    Dim arrStr() As String
    Dim strTime As String
    Dim rsTmp As New ADODB.Recordset
    Dim strBegin As String, strEnd As String
    Dim intHour As Integer
    Dim lngRow As Long
    Dim strCurrDate As String, strDate As String
    Dim strReturn As String
    Dim bln���� As Boolean
    
    On Error GoTo Errhand
    
    strDate = mstrBegin
    strInfo = ""
    lngRow = intRow - vsfTab.FixedRows + 1
    strTmp = vsfTab.TextMatrix(intRow, COL_tab�ַ���)
    lng��Ŀ��� = Val(vsfTab.TextMatrix(intRow, COL_tab��Ŀ���))
    str��Ŀ���� = vsfTab.TextMatrix(intRow, COL_tab��Ŀ����)
    strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
    arrStr = Split(strTmp, ",")
    intƵ�� = Val(arrStr(3))
    int��Ŀ��ʾ = Val(arrStr(4))
    
    bln���� = IsWaveItem(lng��Ŀ���)
    
    '����/������Ŀ����=2
    If int��Ŀ��ʾ = 4 Or bln���� Then
        intType = 2
        If intƵ�� = 0 Then
            intƵ�� = 2
        ElseIf intƵ�� > 2 Then
            intƵ�� = 2
        End If
        
        '�ɲ���ȷ������/������Ŀ����¼����������ݻ��ǵ��������
        If Not mbln���ܵ��� Then strDate = CDate(mstrBegin) - 1
    Else
        intType = 1
    End If
    
    '��ȡ��ǰ��¼��Ƶ��
    intNO = intCOl - vsfTab.FixedCols + 1
    
    '�������ͣ�Ƶ�κ���� �������Ҳ�����Ϣ
    mrsTabTime.Filter = "����=" & intType & " and Ƶ��=" & intƵ�� & " and ���=" & intNO
    If mrsTabTime.RecordCount = 0 Then
        strInfo = "���ڻ�����Ŀ����������[" & IIf(intType = 2, "������Ŀ", "���±����Ŀ") & "]ʱ����Ϣ!"
        Exit Function
    End If
    
    With mrsTabTime
        .MoveFirst
        intHour = CInt(24 / intƵ��)
        strBegin = Format(IIf(IsDate(Trim(Nvl(!��ʼ))) = False, (Val(Nvl(!���)) - 1) * intHour & ":00:00", !��ʼ), "HH:mm:ss")
        strEnd = Format(IIf(IsDate(Trim(Nvl(!����))) = False, Val(Nvl(!���)) * intHour - 1 & ":59:59", !����), "HH:mm:ss")
        If intNO = intƵ�� Then
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
    
    '��ȡϵͳ��ǰʱ��
    strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
    '��ȡ�е�ʱ��
    intHour = DateDiff("H", CDate(strBegin), CDate(strEnd) + 0.00001) / 2
    strTime = DateAdd("H", intHour, CDate(strBegin)) '�е�ʱ��
    
    '������Ŀ���⴦��
'    If int��Ŀ��ʾ = 4 Or bln���� = True Then
'        '���µ���ʼ���첻���ڻ�������¼��
'        If Format(mstrBegin, "YYYY-MM-DD") = Format(mstrBTime, "YYYY-MM-DD") Then
'            strInfo = "����/������Ŀ[" & str��Ŀ���� & "]�����µ���ʼ���첻����¼������[���µ���ʼʱ�䣺" & Format(mstrBTime, "YYYY-MM-DD HH:mm:ss") & "]��"
'            GoTo ExitFunction
'        End If
'        GoTo ErrNext
'    End If
    
    '����¼�뵱������� �Ե�ǰʱ��Ϊ׼(�ڵ�ǰʱ�������Ŀ��ʱ�䷶Χʱ)
    If CDate(strCurrDate) >= CDate(strBegin) And CDate(strCurrDate) <= CDate(strEnd) Then
        strTime = strCurrDate
    End If
    
    If CDate(strTime) < CDate(mstrBTime) Then
        strTime = Format(mstrBTime, "YYYY-MM-DD HH:mm:ss")
        If CDate(strTime) > CDate(strEnd) Then
           strInfo = "��" & lngRow & "��[" & str��Ŀ���� & "]�Ľ���ʱ�䣺" & Format(strEnd, "YYYY-MM-DD HH:mm:ss") & "������С��[���µ���ʼʱ�䣺" & Format(mstrBTime, "YYYY-MM-DD HH:mm:ss") & "]��"
           GoTo ExitFunction
        End If
    End If
    
    If CDate(strTime) > CDate(mstrETime) Then
        strTime = Format(mstrETime, "YYYY-MM-DD HH:mm:ss")
        If CDate(strTime) < CDate(strBegin) Then
            If mbln��Ժ = False Then
                strInfo = "��" & lngRow & "��[" & str��Ŀ���� & "]�Ŀ�ʼʱ�䣺" & Format(strBegin, "YYYY-MM-DD HH:mm:ss") & "���ѳ�������[����¼��������" & mintPreDays & "��]��ָ���ķ�Χ��"
            Else
                strInfo = "��" & lngRow & "��[" & str��Ŀ���� & "]�Ŀ�ʼʱ�䣺" & Format(strBegin, "YYYY-MM-DD HH:mm:ss") & "�����ܴ���[���˳�Ժʱ�䣺" & Format(mstrETime, "YYYY-MM-DD HH:mm:ss") & "]!"
            End If
            GoTo ExitFunction
        End If
    End If
    
ErrNext:
    '��鲡��ת�ƺ�Ĳ�¼ʱ��
    If Not IsAllowInput(mT_Patient.lng����ID, mT_Patient.lng��ҳID, strEnd, strCurrDate) Then
        strInfo = "��¼����ʱ��[" & strTime & "]����[�������ݲ�¼����Чʱ��:" & mlngHours & "Сʱ]"
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
'����:��ʼ����¼�� ������λ��Ϣ��������Ŀʱ�Σ���¼Ƶ��ʱ��
'----------------------------------------------------------------
    On Error GoTo Errhand
    '��ȡ���в�λ��Ϣ
    mstrSQL = "Select ��Ŀ���,��λ,ȱʡ�� From ���²�λ"
    Call zlDatabase.OpenRecordset(mrsPart, mstrSQL, Me.Caption)
    
    '��ȡ���ü�¼����Ϣ
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
'����:��ȡ��¼mrsCurve�е�������
'----------------------------------------------------
    mrsCurve.Filter = 0
    mrsCurve.Sort = "��� Desc"
    If mrsCurve.RecordCount = 0 Then
        GetMaxID = 1
    Else
        GetMaxID = Val(mrsCurve!���) + 1
    End If
End Function

Private Function CheckValidata(ByVal intRow As Integer, ByVal intCOl As Integer, ByVal lngNO As Long, ByVal intType As Integer, ByVal intС�� As Integer, ByVal strֵ�� As String, _
    ByVal int��ʾ As Integer, ByVal lngLen As Long, strInfo As String, Optional strErrMsg As String = "") As Boolean
'----------------------------------------------------------------------------------------------------------------
'����:������ݺϷ���(��������������Ŀ�ͱ����Ŀ�ļ��)
'����:introw����һ�� intCol�� ��һ��  lngNo:��Ŀ��� intype�� ��Ŀ���� 0�������� 1 �������� strֵ����Ŀֵ��
'   lngLen����Ŀ����  strInfo��ҪУ����ı�ֵ
'----------------------------------------------------------------------------------------------------------------
    Dim strName As String, strMsg As String
    Dim lngRow As Long
    Dim arrValue() As String
    Dim lngCount As Long, i As Integer, blnOk As Boolean, strText As String
    Dim blnAllow As Boolean
    
    strName = Split(vsfTab.TextMatrix(intRow, COL_tab��Ŀ����), "(")(0)
    lngRow = intRow - vsfTab.FixedRows + 1
    
    If strInfo = "" Then
        CheckValidata = True
        Exit Function
    End If
    
    blnAllow = True
    
    If strName = "����" Or strName = "���" Then
        If IsNumeric(strInfo) Then
            blnAllow = True
        Else
            blnAllow = False
        End If
    End If
    
    '���������ų��������
    If blnAllow = True Then blnAllow = IIf(InStr(1, "," & gint��� & "," & gint��Һ & ",", "," & lngNO & ",") > 0, False, True)
    
    If Not (intType = 0 And InStr(1, "0,4", int��ʾ) <> 0) Then
        If LenB(StrConv(strInfo, vbFromUnicode)) > lngLen Then
            strMsg = "��" & lngRow & "��[" & strName & "]��ֵ����(��󳤶�:" & lngLen & "),����!"
            GoTo ErrInfo
        End If
    Else
        If intType = 0 Then
            If int��ʾ = 4 Or strֵ�� = "" Then
                strֵ�� = "0��" & IIf(lngLen - intС�� > 0, String(lngLen - intС��, "9"), "0") & IIf(intС�� > 0, "." & String(intС��, "9"), "")
            End If
            If lngNO <> 4 And lngNO <> 5 And blnAllow = True Then
                If Not IsNumeric(strInfo) Then
                    strMsg = strName & "����¼�����" & Space(4) & "��Ч��Χ:" & strֵ��
                    GoTo ErrInfo
                End If
            End If
            
            If lngNO = 4 And strName = "Ѫѹ" Then
                '--�����:53505,�޸��ˣ�����,Ѫѹ����¼������˵���������δ���
                If strInfo = "���" Or strInfo = "δ��" Or strInfo = "�ܲ�" Or strInfo = "���" Then
                    CheckValidata = True
                    Exit Function
                Else
                    If InStr(1, strInfo, "/") = 0 Then
                        strMsg = "��" & lngRow & "��[Ѫѹ]���ݵĸ�ʽ��������ѹ/����ѹ��"
                        GoTo ErrInfo
                    End If
                    If Trim(Split(strInfo, "/")(0)) = "" Or Trim(Split(strInfo, "/")(1)) = "" Then
                        strMsg = "��" & lngRow & "��[Ѫѹ]����¼���������ѹ/����ѹ��"
                        GoTo ErrInfo
                    End If
                End If
            End If
            
            If UBound(Split(strInfo, "/")) > 1 And blnAllow = True Then
                strMsg = "��" & lngRow & "��[" & strName & "]����¼��������飡"
                GoTo ErrInfo
            End If
            
            '�����������Ч��Χ���Ƿ���Ч
            arrValue = Split(strInfo, "/")
            lngCount = UBound(arrValue)
            For i = 0 To lngCount
                blnOk = False
                strText = arrValue(i)
                If Not blnOk Then
                    If Not IsNumeric(strText) And blnAllow = True Then
                        strMsg = "��" & lngRow & "��[" & strName & "]����¼�����" & Space(4) & "��Ч��Χ:" & strֵ��
                        GoTo ErrInfo
                    End If
                End If
                
                If Not blnOk And strText <> "" And blnAllow = True Then strText = Format(Val(strText), "#0" & IIf(intС�� > 0, ".", "") & String(intС��, "0"))
                
                If int��ʾ <> 4 And blnAllow = True Then
                    If Len(Replace(strText, ".", "")) > lngLen Then
                        strMsg = "��" & lngRow & "��[" & strName & "]��ֵ����(��󳤶�:" & lngLen & "),����!"
                        GoTo ErrInfo
                    End If
                End If
                
                If IsNumeric(Split(strֵ��, "��")(0)) And IsNumeric(strText) Then
                    If blnAllow = True Then   '��������������Ч��Χ���
                        If Not (Val(strText) >= Split(strֵ��, "��")(0) And Val(strText) <= Split(strֵ��, "��")(1)) Then
                            strMsg = strName & "������Ч��Χ(" & strֵ�� & "),����!"
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
'����:����û��޸���������ʱ��ʱ���Ƿ�Ϸ�
'-----------------------------------------------------------
    Dim strBegin As String, strEnd As String, strTime As String
    strEnd = Format(mstrEnd, "HH:mm")
    strBegin = Format(mstrBegin, "HH:mm")
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo Errhand
    
    If Format(dkpTime.Value, "HH:mm") = Format(mArrdkpTime(dkpTime.Tag), "HH:mm") Then ChangeCurveTime = True: Exit Function
    
    If Format(dkpTime.Value, "HH:mm") < strBegin And Format(dkpTime.Value, "HH:mm") > strEnd Then
        lblStb.Caption = "��������ʱ��ֻ���� " & strBegin & "" & strEnd & " ֮��"
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
    
    '����޸ĵ�ʱ���Ƿ��Ѿ���������
    mstrSQL = "select 1 From ���˻����ļ� a,���˻������� b" & vbNewLine & _
        " where A.ID=B.�ļ�ID and A.ID=[1] and A.����ID=[2] and A.��ҳID=[3] And nvl(A.Ӥ��,0)=[4]" & vbNewLine & _
        " and B.����ʱ��=[5]"
        
    If mblnMove Then
        mstrSQL = Replace(mstrSQL, "���˻����ļ�", "H���˻����ļ�")
        mstrSQL = Replace(mstrSQL, "���˻�������", "H���˻�������")
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(mstrSQL, "���ʱ��", mT_Patient.lng�ļ�ID, mT_Patient.lng����ID, mT_Patient.lng��ҳID, mT_Patient.lngӤ��, CDate(strTime))
    
    If rsTemp.RecordCount > 0 Then
        lblStb.Caption = "��ʱ���Ѿ���������,����������ʱ��."
        lblStb.ForeColor = 255
        dkpTime.Value = Format(mArrdkpTime(dkpTime.Tag), "HH:mm")
        If dkpTime.Enabled = True Then dkpTime.SetFocus
        Exit Function
    End If
    
    '����Ƿ񳬳�����ʱ��
    If Not CheckDateTime(0, "", strTime) Then
        dkpTime.Value = Format(mArrdkpTime(dkpTime.Tag), "HH:mm")
        If dkpTime.Enabled = True Then dkpTime.SetFocus
        Exit Function
    End If
    
    '�޸ı�ʱ����ڵ�����������������ʱ��
    mrsCurve.Filter = 0
    mrsCurve.Filter = "������='1)����������Ŀ' And ʱ��='" & Format(mArrdkpTime(dkpTime.Tag), "YYYY-MM-DD HH:mm:ss") & "'"
    If mrsCurve.RecordCount > 0 Then mblnChage = True: mblnCurveChange = True
    
    '״̬ 1���� ,2 �޸� ,3������ɾ��(δ����),4 ֻ���޸�ʱ��
    With mrsCurve
        Do While Not .EOF
            !ʱ�� = strTime
             If Val(!״̬) <> 1 And Val(!״̬) <> 3 Then
                If Val(!״̬) = 2 Then
                    mrsCurve!״̬ = 2
                Else
                    mrsCurve!״̬ = 4
                End If
            Else
                If mrsCurve!��ֵ = "" And mrsCurve!δ��˵�� = "" Then
                    mrsCurve!״̬ = 3
                Else
                    mrsCurve!״̬ = 1
                End If
            End If
            .Update
        .MoveNext
        Loop
    End With
   
    '����ʱ�������ֵ
    mArrdkpTime(dkpTime.Tag) = Format(strTime, "YYYY-MM-DD HH:mm:ss")
    
    ChangeCurveTime = True
    
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function is������Һ(ByVal intType As Integer) As Boolean
'����Ƿ��Ǵ����Ŀ����ҹ��Ŀ  �����Ŀ���=10 ��ҹ=9
'intType=1 Ϊ�����Ŀ ����Ϊ��Һ��Ŀ
    Dim lngItemNO As Long
    Dim strKey As String
    Dim rsObj As New ADODB.Recordset
    Dim strTmp As String, strName As String, arrStr
    On Error GoTo Errhand
    
    If vsfTab.Col < vsfTab.FixedCols Or vsfTab.Row < vsfTab.FixedRows Then Exit Function
    If mblnInit = False Then Exit Function
    
    '��ȡ��Ŀ���
    lngItemNO = Val(vsfTab.TextMatrix(vsfTab.Row, COL_tab��Ŀ���))
    If intType = 1 Then
        If lngItemNO <> 10 Then Exit Function
    Else
        If lngItemNO <> 9 Then Exit Function
    End If
    strName = vsfTab.TextMatrix(vsfTab.Row, COL_tab��Ŀ����)
    strTmp = vsfTab.TextMatrix(vsfTab.Row, COL_tab�ַ���)
    strTmp = strTmp & String(8 - UBound(Split(strTmp, ",")), ",")
    arrStr = Split(strTmp, ",")
    
    '����¼Ƶ�κ���Ŀ��ʾ
    If vsfTab.Col > vsfTab.FixedCols + Val(arrStr(3)) - 1 Then Exit Function
    If InStr(1, ",2,3,5,", "," & Val(arrStr(4)) & ",") > 0 Then Exit Function
    
    '����Ƿ���ͬ��������
    mrsCurve.Filter = "��Ŀ���=" & lngItemNO & " and ��Ŀ����='" & strName & "'" & _
        "   and �к�=" & vsfTab.Col - vsfTab.FixedCols + 1
    If mrsCurve.RecordCount > 0 Then
        If InStr(1, ",0,9,", "," & Val(mrsCurve!������Դ) & ",") = 0 Then
            Exit Function
        End If
    End If
    
    is������Һ = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


