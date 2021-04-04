VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "Msmask32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.2#0"; "MSCOMCTL.OCX"
Object = "*\A..\ZLPATIADDRESS\ZlPatiAddress.vbp"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{7CAC59E5-B703-4CCF-B326-8B956D962F27}#9.60#0"; "Codejock.ReportControl.9600.ocx"
Begin VB.Form frmCertifyStation 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   Caption         =   "实名认证工作站"
   ClientHeight    =   12390
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   20040
   Icon            =   "frmCertifyStation.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   12390
   ScaleWidth      =   20040
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.StatusBar stbBar 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   75
      Top             =   12030
      Width           =   20040
      _ExtentX        =   35348
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCertifyStation.frx":6852
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   30824
            Key             =   "Info"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Bevel           =   0
            Object.Width           =   1587
            MinWidth        =   1587
            Text            =   "病人状态"
            TextSave        =   "病人状态"
            Key             =   "病人状态"
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
   Begin VB.PictureBox picPati 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   11295
      Left            =   120
      ScaleHeight     =   11295
      ScaleWidth      =   6240
      TabIndex        =   28
      Top             =   720
      Width           =   6240
      Begin XtremeReportControl.ReportControl rptPati 
         Height          =   6570
         Left            =   0
         TabIndex        =   5
         Top             =   1800
         Width           =   6135
         _Version        =   589884
         _ExtentX        =   10821
         _ExtentY        =   11589
         _StockProps     =   0
         MultipleSelection=   0   'False
         EditOnClick     =   0   'False
         AutoColumnSizing=   0   'False
      End
      Begin VB.PictureBox picFilter 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00FFFFF0&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   0
         ScaleHeight     =   1455
         ScaleWidth      =   6135
         TabIndex        =   30
         Top             =   120
         Width           =   6135
         Begin VB.PictureBox picstrFilter 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   2400
            ScaleHeight     =   240
            ScaleWidth      =   2235
            TabIndex        =   81
            Top             =   352
            Width           =   2260
            Begin VB.ComboBox cbostrFilter 
               Appearance      =   0  'Flat
               Height          =   300
               ItemData        =   "frmCertifyStation.frx":70E4
               Left            =   -30
               List            =   "frmCertifyStation.frx":70E6
               TabIndex        =   0
               Text            =   "cbostrFilter"
               Top             =   -30
               Width           =   2295
            End
         End
         Begin VB.PictureBox picCboFilter 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   840
            ScaleHeight     =   240
            ScaleWidth      =   1515
            TabIndex        =   79
            Top             =   352
            Width           =   1540
            Begin VB.ComboBox cboFilter 
               Appearance      =   0  'Flat
               Height          =   300
               Left            =   -30
               Style           =   2  'Dropdown List
               TabIndex        =   80
               Top             =   -30
               Width           =   1575
            End
         End
         Begin VB.CommandButton cmdFilter 
            Appearance      =   0  'Flat
            Caption         =   "确 定"
            Height          =   350
            Left            =   5400
            TabIndex        =   4
            Top             =   810
            Width           =   615
         End
         Begin VB.CommandButton cmdDate 
            Appearance      =   0  'Flat
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
            Left            =   5030
            Picture         =   "frmCertifyStation.frx":70E8
            Style           =   1  'Graphical
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   870
            Width           =   270
         End
         Begin VB.CommandButton cmdDate 
            Appearance      =   0  'Flat
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
            Left            =   2645
            Picture         =   "frmCertifyStation.frx":71DE
            Style           =   1  'Graphical
            TabIndex        =   70
            TabStop         =   0   'False
            Top             =   870
            Width           =   270
         End
         Begin VB.CheckBox chkOption 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "是否陪诊人"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   4800
            TabIndex        =   1
            Top             =   360
            Width           =   1215
         End
         Begin MSMask.MaskEdBox mskDate 
            Height          =   255
            Index           =   0
            Left            =   840
            TabIndex        =   2
            Tag             =   "####-##-## ##:##"
            Top             =   870
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            AutoTab         =   -1  'True
            MaxLength       =   16
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####-##-## ##:##"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtDate 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   275
            Index           =   0
            Left            =   840
            TabIndex        =   68
            Top             =   855
            Visible         =   0   'False
            Width           =   1740
         End
         Begin MSMask.MaskEdBox mskDate 
            Height          =   255
            Index           =   1
            Left            =   3240
            TabIndex        =   3
            Tag             =   "####-##-## ##:##"
            Top             =   870
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            AutoTab         =   -1  'True
            MaxLength       =   16
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####-##-## ##:##"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtDate 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   275
            Index           =   1
            Left            =   3240
            TabIndex        =   69
            Top             =   855
            Visible         =   0   'False
            Width           =   1740
         End
         Begin VB.Label lblType 
            AutoSize        =   -1  'True
            Caption         =   " 精 确↓"
            Height          =   180
            Left            =   0
            TabIndex        =   72
            Top             =   397
            Width           =   720
         End
         Begin VB.Label lbl建档 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "至"
            Height          =   180
            Left            =   3000
            TabIndex        =   54
            Top             =   900
            Width           =   180
         End
      End
   End
   Begin MSComctlLib.ImageList imgIcons 
      Left            =   1680
      Top             =   120
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
            Picture         =   "frmCertifyStation.frx":72D4
            Key             =   "Certify_StateSure"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCertifyStation.frx":DB36
            Key             =   "Certify_StateStop"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCertifyStation.frx":14398
            Key             =   "Certify_StateFalse"
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picFormation 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   12375
      Left            =   6600
      ScaleHeight     =   12375
      ScaleWidth      =   13095
      TabIndex        =   29
      Top             =   120
      Width           =   13095
      Begin VB.VScrollBar vsbMain 
         Height          =   7335
         LargeChange     =   100
         Left            =   12840
         Max             =   1000
         SmallChange     =   10
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   1200
         Width           =   255
      End
      Begin VB.PictureBox picDetailInfo 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   13575
         Left            =   120
         ScaleHeight     =   13545
         ScaleWidth      =   12735
         TabIndex        =   31
         Top             =   120
         Width           =   12765
         Begin VB.Frame frmInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   5
            Left            =   10875
            TabIndex        =   64
            Top             =   4320
            Width           =   1690
            Begin VB.ComboBox cboInfo 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   5
               Left            =   -30
               Style           =   2  'Dropdown List
               TabIndex        =   19
               Top             =   -30
               Width           =   1670
            End
         End
         Begin VB.Frame frmInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   8
            Left            =   10875
            TabIndex        =   63
            Top             =   4695
            Width           =   1660
            Begin VB.ComboBox cboInfo 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   8
               Left            =   -30
               Style           =   2  'Dropdown List
               TabIndex        =   23
               Top             =   -30
               Width           =   1670
            End
         End
         Begin VB.Frame frmInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   7
            Left            =   1200
            TabIndex        =   62
            Top             =   4695
            Width           =   1820
            Begin VB.ComboBox cboInfo 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   7
               Left            =   -30
               Style           =   2  'Dropdown List
               TabIndex        =   22
               Top             =   -30
               Width           =   1810
            End
         End
         Begin VB.Frame frmInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   6
            Left            =   4800
            TabIndex        =   61
            Top             =   4695
            Width           =   1480
            Begin VB.ComboBox cboInfo 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   6
               Left            =   -30
               Style           =   2  'Dropdown List
               TabIndex        =   21
               Top             =   -30
               Width           =   1455
            End
         End
         Begin VB.Frame frmInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   4
            Left            =   4800
            TabIndex        =   60
            Top             =   4335
            Width           =   1480
            Begin VB.ComboBox cboInfo 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   4
               Left            =   -30
               Style           =   2  'Dropdown List
               TabIndex        =   18
               Top             =   -30
               Width           =   1455
            End
         End
         Begin VB.Frame frmInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   3
            Left            =   1200
            TabIndex        =   59
            Top             =   2400
            Width           =   1820
            Begin VB.ComboBox cboInfo 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   3
               Left            =   -30
               Style           =   2  'Dropdown List
               TabIndex        =   12
               Top             =   -30
               Width           =   1810
            End
         End
         Begin VB.Frame frmInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   2
            Left            =   4800
            TabIndex        =   58
            Top             =   2400
            Width           =   1480
            Begin VB.ComboBox cboInfo 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   2
               Left            =   -30
               Style           =   2  'Dropdown List
               TabIndex        =   11
               Top             =   -30
               Width           =   1455
            End
         End
         Begin VB.Frame frmInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   0
            Left            =   4800
            TabIndex        =   57
            Top             =   2040
            Width           =   1480
            Begin VB.ComboBox cboInfo 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   0
               Left            =   -30
               Style           =   2  'Dropdown List
               TabIndex        =   8
               Top             =   -30
               Width           =   1455
            End
         End
         Begin VB.Frame frmInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   1
            Left            =   10875
            TabIndex        =   56
            Top             =   2040
            Width           =   1690
            Begin VB.ComboBox cboInfo 
               Appearance      =   0  'Flat
               Height          =   300
               Index           =   1
               Left            =   -30
               Style           =   2  'Dropdown List
               TabIndex        =   9
               Top             =   -30
               Width           =   1670
            End
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   270
            Index           =   5
            Left            =   4800
            TabIndex        =   27
            Top             =   5760
            Width           =   7695
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   270
            Index           =   4
            Left            =   1200
            TabIndex        =   26
            Top             =   5760
            Width           =   1815
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   270
            Index           =   3
            Left            =   7515
            TabIndex        =   20
            Top             =   4680
            Width           =   1815
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   270
            Index           =   2
            Left            =   1200
            TabIndex        =   17
            Top             =   4320
            Width           =   1815
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   270
            Index           =   1
            Left            =   7590
            TabIndex        =   10
            Top             =   2392
            Width           =   1815
         End
         Begin VB.TextBox txtInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   270
            Index           =   0
            Left            =   1200
            TabIndex        =   7
            Top             =   2025
            Width           =   1815
         End
         Begin ZlPatiAddress.PatiAddress patiAdressInfo 
            Height          =   270
            Index           =   1
            Left            =   7500
            TabIndex        =   15
            Top             =   2745
            Width           =   5025
            _ExtentX        =   8864
            _ExtentY        =   476
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            Style           =   1
         End
         Begin VB.TextBox txtAdressInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   270
            Index           =   1
            Left            =   7515
            TabIndex        =   16
            Top             =   2745
            Width           =   5025
         End
         Begin ZlPatiAddress.PatiAddress patiAdressInfo 
            Height          =   270
            Index           =   0
            Left            =   1200
            TabIndex        =   13
            Top             =   2760
            Width           =   5025
            _ExtentX        =   8864
            _ExtentY        =   476
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   15659001
            Items           =   3
            Style           =   1
         End
         Begin VB.TextBox txtAdressInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   270
            Index           =   0
            Left            =   1200
            TabIndex        =   14
            Top             =   2760
            Width           =   4995
         End
         Begin ZlPatiAddress.PatiAddress patiAdressInfo 
            Height          =   270
            Index           =   2
            Left            =   1200
            TabIndex        =   24
            Top             =   5040
            Width           =   5025
            _ExtentX        =   8864
            _ExtentY        =   476
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BackColor       =   -2147483643
            Style           =   1
         End
         Begin VB.TextBox txtAdressInfo 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   270
            Index           =   2
            Left            =   1200
            TabIndex        =   25
            Top             =   5040
            Width           =   5040
         End
         Begin VSFlex8Ctl.VSFlexGrid vfgCert 
            Height          =   2295
            Left            =   360
            TabIndex        =   65
            Top             =   7200
            Width           =   12015
            _cx             =   21193
            _cy             =   4048
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
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   9
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   325
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
         Begin MSMask.MaskEdBox txt出生日期 
            Height          =   255
            Index           =   0
            Left            =   7590
            TabIndex        =   73
            Tag             =   "####-##-## ##:##"
            Top             =   2040
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            AutoTab         =   -1  'True
            MaxLength       =   16
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####-##-## ##:##"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtInfoDate 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   270
            Index           =   0
            Left            =   7590
            TabIndex        =   66
            Top             =   2025
            Width           =   1740
         End
         Begin MSMask.MaskEdBox txt出生日期 
            Height          =   255
            Index           =   1
            Left            =   7515
            TabIndex        =   74
            Tag             =   "####-##-## ##:##"
            Top             =   4335
            Width           =   1740
            _ExtentX        =   3069
            _ExtentY        =   450
            _Version        =   393216
            BorderStyle     =   0
            Appearance      =   0
            AutoTab         =   -1  'True
            MaxLength       =   16
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   10.5
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Mask            =   "####-##-## ##:##"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtInfoDate 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   270
            Index           =   1
            Left            =   7515
            TabIndex        =   67
            Top             =   4320
            Width           =   1740
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfInterface 
            Height          =   2295
            Left            =   360
            TabIndex        =   77
            Top             =   10560
            Width           =   12015
            _cx             =   21193
            _cy             =   4048
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
            BackColorSel    =   -2147483643
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   9
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   325
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
         Begin VB.Label lblTitles 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "三方认证信息"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   3
            Left            =   5760
            TabIndex        =   78
            Top             =   9960
            Width           =   1815
         End
         Begin VB.Label lblTitles 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "证件信息"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   2
            Left            =   5760
            TabIndex        =   55
            Top             =   6600
            Width           =   1215
         End
         Begin VB.Label lblTitles 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "病人信息"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   0
            Left            =   5760
            TabIndex        =   53
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label lblTitles 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "陪诊人信息"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   14.25
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000008&
            Height          =   375
            Index           =   1
            Left            =   5640
            TabIndex        =   52
            Top             =   3480
            Width           =   1455
         End
         Begin VB.Label lblFeild 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "备    注"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   19
            Left            =   3870
            TabIndex        =   51
            Top             =   5775
            Width           =   735
         End
         Begin VB.Label lblFeild 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "手 机 号"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   18
            Left            =   360
            TabIndex        =   50
            Top             =   5775
            Width           =   735
         End
         Begin VB.Label lblFeild 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "关    系"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   16
            Left            =   10035
            TabIndex        =   49
            Top             =   4695
            Width           =   735
         End
         Begin VB.Label lblFeild 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "住    址"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   17
            Left            =   360
            TabIndex        =   48
            Top             =   5055
            Width           =   720
         End
         Begin VB.Label lblFeild 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "身份证类型"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   15
            Left            =   195
            TabIndex        =   47
            Top             =   4695
            Width           =   900
         End
         Begin VB.Label lblFeild 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "身份证号"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   13
            Left            =   6675
            TabIndex        =   46
            Top             =   4695
            Width           =   735
         End
         Begin VB.Label lblFeild 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "民    族"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   14
            Left            =   3870
            TabIndex        =   45
            Top             =   4695
            Width           =   735
         End
         Begin VB.Label lblFeild 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "国    籍"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   12
            Left            =   10035
            TabIndex        =   44
            Top             =   4335
            Width           =   735
         End
         Begin VB.Label lblFeild 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "出生日期"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   11
            Left            =   6675
            TabIndex        =   43
            Top             =   4335
            Width           =   735
         End
         Begin VB.Label lblFeild 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "性    别"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   10
            Left            =   3870
            TabIndex        =   42
            Top             =   4335
            Width           =   735
         End
         Begin VB.Label lblFeild 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "姓    名"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   9
            Left            =   360
            TabIndex        =   41
            Top             =   4335
            Width           =   735
         End
         Begin VB.Label lblFeild 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "身份证类型"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   195
            TabIndex        =   40
            Top             =   2400
            Width           =   900
         End
         Begin VB.Label lblFeild 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "身份证号"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   6675
            TabIndex        =   39
            Top             =   2400
            Width           =   735
         End
         Begin VB.Label lblFeild 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "住    址"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   8
            Left            =   6675
            TabIndex        =   38
            Top             =   2760
            Width           =   720
         End
         Begin VB.Label lblFeild 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "出生地点"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   7
            Left            =   360
            TabIndex        =   37
            Top             =   2760
            Width           =   720
         End
         Begin VB.Label lblFeild 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "民    族"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   3870
            TabIndex        =   36
            Top             =   2400
            Width           =   735
         End
         Begin VB.Label lblFeild 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "国    籍"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   10035
            TabIndex        =   35
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label lblFeild 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "出生日期"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   2
            Left            =   6675
            TabIndex        =   34
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label lblFeild 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "性    别"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   1
            Left            =   3870
            TabIndex        =   33
            Top             =   2040
            Width           =   735
         End
         Begin VB.Label lblFeild 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "姓    名"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   0
            Left            =   360
            TabIndex        =   32
            Top             =   2040
            Width           =   735
         End
      End
   End
   Begin MSComCtl2.MonthView monInfo 
      Height          =   2160
      Left            =   0
      TabIndex        =   76
      TabStop         =   0   'False
      Top             =   0
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   3810
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   0
      StartOfWeek     =   200671233
      TitleBackColor  =   8421504
      TitleForeColor  =   16777215
      CurrentDate     =   38003
   End
   Begin VB.Image img图片 
      Height          =   240
      Left            =   3360
      Picture         =   "frmCertifyStation.frx":1ABFA
      Top             =   360
      Width           =   240
   End
   Begin XtremeCommandBars.ImageManager imgManager 
      Left            =   2640
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmCertifyStation.frx":2144C
   End
   Begin XtremeCommandBars.CommandBars cmbMain 
      Left            =   720
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Bindings        =   "frmCertifyStation.frx":2C066
      Left            =   240
      Top             =   120
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmCertifyStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mlng病人ID As Long, mlng实名id As Long
Private mblnEdit As Boolean '判断控件是否锁定
Private mbln精确 As Boolean '是否景区查找
Private mstrFindType As String '查找方式
Private mintFindType As Integer
Private mstrFilter As String  '过滤字符串
Private mstrInput As String   '最终传入的过滤方式
Private mstrInputB As String  '范围查找的开始时间
Private mstrInputE As String  '范围查找的结束时间
Private mintDate As Integer   '时间控件的索引
Private mblnStop As Boolean   '是否停用
Private mbytSize As Byte      '字体大小
Private mlngSource As Long
Private mlngTopVsc As Long
Private mstrPrePati As String
Private mrsCertType As New ADODB.Recordset
Private Const SM_CXVSCROLL = 2
 
Private Enum Pati_Clum
    COL_TAG = 0
    COL_病人Id
    COL_实名ID
    COL_姓名
    COL_性别
    COL_出生日期
    COL_国籍
    COL_民族
    COL_身份证号
    COL_陪诊人姓名
    COL_陪诊人关系
    COL_手机号
    COL_备注
    COL_建档时间
    COL_建档人
    COL_更新时间
    COL_更新人
    COL_是否停用
    COL_停用时间
    COL_是否认证
End Enum

Private Enum TXT_Info
    TXT_姓名 = 0
    TXT_身份证号 = 1
    TXT_陪诊人姓名 = 2
    TXT_陪诊人身份证号 = 3
    txt_手机号 = 4
    TXT_备注 = 5
End Enum

Private Enum PatiAdress_Info
    ADRS_出生地点 = 0
    ADRS_住址 = 1
    ADRS_陪诊人住址 = 2
End Enum

Private Enum VSF_COL
    COLS_证件ID = 0
    COLS_证件号码
    COLS_证件类型
    CLOS_备注
    COLS_所有者
    COLS_图片
End Enum

Private Enum VSFInterface_COL
    INT_接口ID = 0
    INT_名称
    INT_部件名
    INT_说明
    INT_认证结果
End Enum

Private Sub cboFilter_Click()
    Dim i As Long
    
    mintFindType = cbo.FindIndex(cboFilter, cboFilter.Text)
    mstrFindType = cboFilter.Text
    If mstrFindType = "姓名" Or mstrFindType = "二代身份证" Or mstrFindType = "其它证件类型" Or mstrFindType = "证件号码" Then
        chkOption.Visible = True
    Else
        chkOption.Visible = False
    End If
    
    If mstrFindType = "其它证件类型" Then
        mrsCertType.MoveFirst
        cbostrFilter.Clear
        If Not mrsCertType.EOF Then
            For i = 0 To mrsCertType.RecordCount - 1
                cbostrFilter.AddItem mrsCertType!名称
                mrsCertType.MoveNext
            Next
        End If
    ElseIf mstrFindType = "认证状态" Then
        cbostrFilter.Clear
        cbostrFilter.AddItem "已认证"
        cbostrFilter.AddItem "未认证"
    Else
        cbostrFilter.Clear
    End If
End Sub

Private Sub cbostrFilter_Click()
    Dim strPati As String, vRect As RECT, strName As String, strIF As String
    Dim rsTmp As ADODB.Recordset
    Dim strTag As String, strFilter As String, strInput As String
    
    On Error GoTo Errhand
    strName = Trim(cbostrFilter.Text)
    If zlCommFun.GetNeedName(cboFilter.Text, "-") Like "其它证件类型" And strName <> "" And InStr("-*+/", Left(Trim(cbostrFilter.Text), 1)) = 0 Then
        If chkOption.Value = 1 Then
            strFilter = "b.证件类型=[1] And b.所有者=2"
        Else
            strFilter = "b.证件类型=[1] And b.所有者=1"
        End If
        strInput = strName
    ElseIf zlCommFun.GetNeedName(cboFilter.Text, "-") Like "认证状态" And strName <> "" Then
        strFilter = "a.认证状态=[1]"
        strInput = IIf(strName = "已认证", 1, 0)
    End If
    
    If strFilter <> "" And strInput <> "" Then
        mstrFilter = strFilter
        mstrInput = strInput
        Screen.MousePointer = 11
        Call LoadPatients(1)
        Screen.MousePointer = 0
    ElseIf strFilter <> "" Or strInput = "" Then
        MsgBox "请输入查询条件！", vbInformation, gstrSysName
        Exit Sub
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cbostrFilter_KeyPress(KeyAscii As Integer)
    Dim strPati As String, vRect As RECT, strName As String, strIF As String
    Dim rsTmp As ADODB.Recordset
    Dim strTag As String, strFilter As String, strInput As String
    
    On Error GoTo Errhand
    If KeyAscii = 13 Then
        strName = Trim(cbostrFilter.Text)
        If zlCommFun.GetNeedName(cboFilter.Text, "-") Like "姓名" And strName <> "" And InStr("-*+/", Left(Trim(cbostrFilter.Text), 1)) = 0 Then
            strFilter = IIf(chkOption.Value = 1, "a.陪诊人姓名=[1]", "a.姓名=[1]")
            strInput = strName
        ElseIf zlCommFun.GetNeedName(cboFilter.Text, "-") Like "二代身份证" And strName <> "" And InStr("-*+/", Left(Trim(cbostrFilter.Text), 1)) = 0 Then
            strFilter = IIf(chkOption.Value = 1, "a.陪诊人身份证号=[1]", "a.身份证号=[1]")
            strInput = strName
        ElseIf zlCommFun.GetNeedName(cboFilter.Text, "-") Like "其它证件类型" And strName <> "" And InStr("-*+/", Left(Trim(cbostrFilter.Text), 1)) = 0 Then
            If chkOption.Value = 1 Then
                strFilter = "b.证件类型=[1] And b.所有者=2"
            Else
                strFilter = "b.证件类型=[1] And b.所有者=1"
            End If
            strInput = strName
        ElseIf zlCommFun.GetNeedName(cboFilter.Text, "-") Like "证件号码" And strName <> "" And InStr("-*+/", Left(Trim(cbostrFilter.Text), 1)) = 0 Then
            If chkOption.Value = 1 Then
                strFilter = "b.证件号码=[1] And b.所有者=2"
            Else
                strFilter = "b.证件号码=[1] And b.所有者=1"
            End If
            strInput = strName
        ElseIf zlCommFun.GetNeedName(cboFilter.Text, "-") Like "认证状态" And strName <> "" And IsNumeric(Trim(cbostrFilter.Text)) Then
            strFilter = "a.认证状态=[1]"
            strInput = IIf(strName = "已认证", 1, 0)
        ElseIf zlCommFun.GetNeedName(cboFilter.Text, "-") Like "手机号" And strName <> "" And IsNumeric(Trim(cbostrFilter.Text)) Then
            strFilter = "a.手机号=[1]"
            strInput = Val(strName)
        End If
        
        If strFilter <> "" And strInput <> "" Then
            mstrFilter = strFilter
            mstrInput = strInput
            Screen.MousePointer = 11
            Call LoadPatients(1)
            Screen.MousePointer = 0
        ElseIf strFilter <> "" Or strInput = "" Then
            MsgBox "请输入查询条件！", vbInformation, gstrSysName
            Exit Sub
        End If
    Else
        Select Case mstrFindType
            Case "姓名", "其他证件类型"
                If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then
                    KeyAscii = 0
                End If
            Case "身份证号"
                If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8) & Chr(13), UCase(Chr(KeyAscii))) = 0 Then
                    KeyAscii = 0
                End If
            Case "手机号"
                If InStr("0123456789" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
            Case "认证状态"
                If InStr("01" & Chr(8) & Chr(13), Chr(KeyAscii)) = 0 Then KeyAscii = 0
        End Select
    End If
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub chkOption_Click()
    If cbostrFilter.Text <> "" Then
        Call cbostrFilter_KeyPress(13)
    End If
End Sub

Private Sub chkOption_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Trim(cbostrFilter.Text) = "" Then
        MsgBox "请输入查询条件！", vbInformation, gstrSysName
        Exit Sub
    End If
End Sub

Private Sub cmbMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim i As Long
    Dim objControl As CommandBarControl
    Dim lng病人ID As Long, lng实名ID As Long
    
    Select Case Control.ID
        Case conMenu_Certify_Add
            frmCertifyRegist.ShowMe Me, 0, lng病人ID, lng实名ID
            Screen.MousePointer = 11
            Call LoadPatients(1, lng实名ID)
            Screen.MousePointer = 0
        Case conMenu_Certify_Modify
            If rptPati.SelectedRows.Count > 0 Then
                If rptPati.FocusedRow.Record(COL_是否认证).Value = 1 Then
                    If MsgBox("姓名为【" & rptPati.FocusedRow.Record(COL_姓名).Value & "】的实名信息已经认证，确定要修改吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Sub
                    End If
                End If
                frmCertifyRegist.ShowMe Me, 1, mlng病人ID, mlng实名id
                Screen.MousePointer = 11
                Call LoadPatients(1)
                Screen.MousePointer = 0
            End If
        Case conMenu_Certify_Verify
            frmCertifyRegist.ShowMe Me, 1, mlng病人ID, mlng实名id
            Screen.MousePointer = 11
            Call LoadPatients(1)
            Screen.MousePointer = 0
        Case conMenu_Certify_Stop
            If rptPati.SelectedRows.Count > 0 Then
                If MsgBox("确定要停用姓名为【" & rptPati.FocusedRow.Record(COL_姓名).Value & "】的实名信息吗？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    mblnStop = True
                Else
                    mblnStop = True
                    Exit Sub
                End If
            End If
            If mblnStop Then
                Call UpdateCertSate
                Screen.MousePointer = 11
                Call LoadPatients(1)
                Screen.MousePointer = 0
            End If
        Case conMenu_Certify_Start
            mblnStop = False
            Call UpdateCertSate
            Screen.MousePointer = 11
            Call LoadPatients(1)
            Screen.MousePointer = 0
        Case conMenu_Certify_Refresh
            Screen.MousePointer = 11
            Call LoadPatients(1)
            Screen.MousePointer = 0
        Case conMenu_Certify_Record
            If rptPati.SelectedRows.Count > 0 Then
                frmCertifyRecord.ShowMe Me, Val(rptPati.FocusedRow.Record(COL_实名ID).Value)
            End If
        Case conMenu_CertifyHelp_Help
            ShowHelp App.ProductName, Me.hwnd, Me.Name
        Case conMenu_Certify_Quit
             Unload Me
        Case conMenu_CertifyView_FontSize_L
'            If mbytSize <> 1 Then
'                mlngSource = 0
'                mbytSize = 1
'                Call zlDatabase.SetPara("字体", mbytSize, glngSys, p病人实名信息管理, True)
'                Call SetFontSize(True)
'                Me.cmbMain.RecalcLayout
'            End If
        Case conMenu_CertifyView_FontSize_S
            If mbytSize <> 0 Then
                mlngSource = 999
                mbytSize = 0
                Call zlDatabase.SetPara("字体", mbytSize, glngSys, p病人实名信息管理, IIf(InStr(gstrPrivs, ";参数设置;") > 0, True, False))
                Call SetFontSize(True)
                Me.cmbMain.RecalcLayout
            End If
        Case conMenu_CertifyView_StatusBar
            Me.stbBar.Visible = Not Me.stbBar.Visible
            Me.cmbMain.RecalcLayout
        Case conMenu_CertifyView_ToolBar_Button
            For i = 2 To cmbMain.Count
                Me.cmbMain(i).Visible = Not Me.cmbMain(i).Visible
            Next
        Me.cmbMain.RecalcLayout
        Case conMenu_CertifyView_ToolBar_Size
            Me.cmbMain.Options.LargeIcons = Not Me.cmbMain.Options.LargeIcons
            Me.cmbMain.RecalcLayout
        Case conMenu_CertifyView_ToolBar_Text
            For i = 2 To cmbMain.Count
                For Each objControl In Me.cmbMain(i).Controls
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                Next
            Next
            Me.cmbMain.RecalcLayout
        Case conMenu_CertifyHelp_About
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case conMenu_CertifyHelp_Web_Home
            Call zlHomePage(Me.hwnd)
        Case conMenu_Certify_ParSet
            Call zlCommFun.DeviceSetup(Me, glngSys, glngModul)
        Case conMenu_Certify_Identify_Set
            frmIdentifySet.ShowMe Me
    End Select
End Sub

Private Sub cmbMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case conMenu_Certify_Modify, conMenu_Certify_Start, conMenu_Certify_Stop, conMenu_Certify_Verify, conMenu_Certify_Record
            With rptPati
                If .SelectedRows.Count > 0 Then
                    If Control.ID = conMenu_Certify_Modify Then
                        Control.Visible = IIf(.FocusedRow.Record(COL_是否停用).Value = 1, False, True)
                    ElseIf Control.ID = conMenu_Certify_Verify Then
                        If .FocusedRow.Record(COL_是否停用).Value = 1 Then
                            Control.Visible = False
                        Else
                            Control.Visible = IIf(.FocusedRow.Record(COL_是否认证).Value = 0, True, False)
                        End If
                    ElseIf Control.ID = conMenu_Certify_Stop Then
                        Control.Visible = IIf(.FocusedRow.Record(COL_是否停用).Value = 1, False, True)
                    ElseIf Control.ID = conMenu_Certify_Start Then
                        Control.Visible = IIf(.FocusedRow.Record(COL_是否停用).Value = 1, True, False)
                    Else
                        Control.Visible = True
                    End If
                Else
                    Control.Visible = False
                End If
            End With
    Case conMenu_Certify_Identify_Set
        If InStr(gstrPrivs, ";认证接口配置;") > 0 Then
            Control.Visible = True
        Else
            Control.Visible = False
        End If
    Case conMenu_CertifyView_ToolBar_Size
        Control.Checked = Me.cmbMain.Options.LargeIcons
    Case conMenu_CertifyView_ToolBar_Button '工具栏
        If cmbMain.Count >= 2 Then
            Control.Checked = Me.cmbMain(2).Visible
        End If
    Case conMenu_CertifyView_StatusBar '状态栏
        Control.Checked = Me.stbBar.Visible
    Case conMenu_CertifyView_FontSize_S '小字体
        Control.Checked = Not (mbytSize = 1)
    Case conMenu_CertifyView_FontSize_L '大字体
        Control.Checked = (mbytSize = 1)
    Case conMenu_CertifyView_ToolBar_Text '图标文字
        If cmbMain.Count >= 2 Then
            Control.Checked = Not (Me.cmbMain(2).Controls(GetFirstCommandBar(cmbMain(2).Controls)).Style = xtpButtonIcon)
        End If
    End Select
End Sub

Public Function GetFirstCommandBar(ByRef objControls As CommandBarControls) As Long
'功能：获取工具栏打印预览按钮后的第一个按钮的index
    Dim objControl As CommandBarControl, idx As Long
    
    For Each objControl In objControls
        If objControl.ID = conMenu_Certify_Add Then
            idx = objControl.Index
        End If
    Next
    GetFirstCommandBar = idx
End Function

Private Sub cmdDate_Click(Index As Integer)
    Dim objmonInfo As MonthView  '方便调用控件属性
    Dim objCmd As CommandButton
    Dim objMSK As MaskEdBox
    Dim datStart As Date
    Dim DateEnd As Date
    Dim datTmp As Date
    On Error GoTo errH
    
    mintDate = Index
    Set objmonInfo = monInfo
    Set objCmd = cmdDate(Index)
    Set objMSK = mskDate(Index)
    datStart = zlDatabase.Currentdate
    objmonInfo.MinDate = 0
    objmonInfo.MaxDate = zlDatabase.Currentdate
    If IsDate(objMSK.Text) Then
        datTmp = CDate(objMSK.Text)
        If datTmp > objmonInfo.MaxDate Then
            datTmp = objmonInfo.MaxDate
        ElseIf datTmp < objmonInfo.MinDate Then
            datTmp = objmonInfo.MinDate
        End If
        objmonInfo.Value = datTmp
    End If
    objmonInfo.Left = objCmd.Left + objCmd.Width - objmonInfo.Width + objMSK.Container.Left + picPati.Left
    objmonInfo.Top = objCmd.Top - objmonInfo.Height - 20 + objMSK.Container.Top + 1280
    objmonInfo.ZOrder
    objmonInfo.Visible = True
    objmonInfo.SetFocus
    Exit Sub
errH:
    If ErrCenter() <> 1 Then
        Resume
    End If
End Sub

Private Sub cmdFilter_Click()
    Dim objCtl As Object
    Dim strTmp As String

    If Not IsDate(mskDate(0).Text) Then
        Set objCtl = mskDate(0)
        strTmp = "开始时间"
    ElseIf Not IsDate(mskDate(1).Text) Then
        Set objCtl = mskDate(1)
        strTmp = "结束时间"
    End If
    If strTmp <> "" Then
        Call ShowMessage(objCtl, strTmp & "不是有效的日期格式。", False)
        Exit Sub
    Else
        mstrFilter = "a.建档时间 Between TO_Date([1],'yyyy-mm-dd hh24:mi:ss') And To_Date([2],'yyyy-mm-dd hh24:mi:ss')"
        mstrInputB = Format(mskDate(0).Text, "YYYY-MM-DD HH:MM:SS")
        mstrInputE = Format(mskDate(1).Text, "YYYY-MM-DD HH:MM:SS")
        Screen.MousePointer = 11
        Call LoadPatients(1)
        Screen.MousePointer = 0
    End If
End Sub

Private Sub Form_Activate()
    glngPreHWnd = GetWindowLong(picFormation.hwnd, GWL_WNDPROC)
    SetWindowLong picFormation.hwnd, GWL_WNDPROC, AddressOf FlexScroll
End Sub

Private Sub Form_Deactivate()
    SetWindowLong picFormation.hwnd, GWL_WNDPROC, glngPreHWnd
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCur As Long, lngMin As Long, lngMax As Long
    
    lngCur = vsbMain.Value
    lngMin = vsbMain.Min
    lngMax = vsbMain.Max
    
    If KeyCode = vbKeyPageDown Then '下
        If Between(lngCur + (lngMax - lngMin) / 10, lngMin, lngMax) Then
            vsbMain.Value = lngCur + (lngMax - lngMin) / 10
        Else
            vsbMain.Value = lngMax
        End If
    ElseIf KeyCode = vbKeyPageUp Then  '上
        If Between(lngCur - (lngMax - lngMin) / 10, lngMin, lngMax) Then
            vsbMain.Value = lngCur - (lngMax - lngMin) / 10
        Else
            vsbMain.Value = lngMin
        End If
    End If
End Sub

Private Sub Form_Load()
    Dim objPane As Pane
    Dim Curdate As Date, datTmp As Date
    Dim strDate As String
    Dim arrDate As Variant
    Dim i As Long
    Dim objFile As New FileSystemObject
                       
    mbytSize = Val(zlDatabase.GetPara("字体", glngSys, p病人实名信息管理, "0"))
    'CommandBar
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cmbMain.VisualTheme = xtpThemeOffice2003
    With Me.cmbMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
'        .UseFadedIcons = True '放在VisualTheme后有效
        .IconsWithShadow = True '放在VisualTheme后有效
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Set cmbMain.Icons = imgManager.Icons
    cmbMain.EnableCustomization False

    'DockingPane
    Me.dkpMain.SetCommandBars Me.cmbMain
    Me.dkpMain.Options.UseSplitterTracker = False '实时拖动
    Me.dkpMain.Options.ThemedFloatingFrames = True
    Me.dkpMain.Options.AlphaDockingContext = True
    Me.dkpMain.VisualTheme = ThemeOffice2003
    Set objPane = Me.dkpMain.CreatePane(1, IIf(mbytSize <> 0, 400, 410), Me.ScaleHeight, DockLeftOf, Nothing)
    objPane.Title = "病人列表"
    objPane.Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
    
    mbln精确 = Val(zlDatabase.GetPara("实名信息查询方式", glngSys, glngModul, 0)) = 0
    lblType.Caption = IIf(mbln精确, " 精 确↓", " 范 围↓")
    
    mintFindType = Val(zlDatabase.GetPara("实名信息查询条件", glngSys, glngModul, 0))
    
    Curdate = zlDatabase.Currentdate
    strDate = zlDatabase.GetPara("实名信息时间范围", glngSys, glngModul)
    If strDate <> "" Then
        arrDate = Split(strDate, ";")
        For i = LBound(arrDate) To UBound(arrDate)
            If i = 0 Then
                mstrInputB = arrDate(i)
            ElseIf i = 1 Then
                mstrInputE = arrDate(i)
            End If
        Next
    Else
        mstrInputB = Format(Curdate - 3, "YYYY-MM-DD")
        mstrInputE = Format(Curdate, "YYYY-MM-DD")
        mstrInputB = Format(mstrInputB & " 00:00", "YYYY-MM-DD HH:MM")
        mstrInputE = Format(mstrInputE & " 23:59", "YYYY-MM-DD  HH:MM")
    End If
    
    '初始化地址控件
    patiAdressInfo(ADRS_出生地点).Visible = gbln启用结构化地址
    patiAdressInfo(ADRS_住址).Visible = gbln启用结构化地址
    patiAdressInfo(ADRS_陪诊人住址).Visible = gbln启用结构化地址
    txtAdressInfo(ADRS_出生地点).Visible = Not gbln启用结构化地址
    txtAdressInfo(ADRS_住址).Visible = Not gbln启用结构化地址
    txtAdressInfo(ADRS_陪诊人住址).Visible = Not gbln启用结构化地址
    If gbln启用结构化地址 Then
        patiAdressInfo(ADRS_出生地点).ShowTown = gbln显示乡镇
        patiAdressInfo(ADRS_住址).ShowTown = gbln显示乡镇
        patiAdressInfo(ADRS_陪诊人住址).ShowTown = gbln显示乡镇
    End If
    
    '初始化菜单
    Call MainDefCommandBar
    
    '初始化实名认证列表
    Call InitReportColumn
    
    '初始化证件信息表格
    Call InitVsfGridHeader
    
    '初始化数据
    Call InitCboData
    
    '加载病人列表
    Screen.MousePointer = 11
    Call LoadPatients(0)
    Screen.MousePointer = 0
    
    '刷新实名认证信息
    Call zlRefresh(mlng病人ID, mlng实名id, True)

    '画线
    Call DrawLin
    
    '设置控件的可用性
    Call SetCtlEnable
    
    mstrInputB = Format(mstrInputB, decode(mskDate(0).Mask, "####-##-##", "yyyy-MM-dd", "####-##-## ##:##", "yyyy-MM-dd HH:mm", "####-##-## ##:##:##", "yyyy-MM-dd HH:mm:ss", "##:##", "HH:mm"))
    mstrInputE = Format(mstrInputE, decode(mskDate(1).Mask, "####-##-##", "yyyy-MM-dd", "####-##-## ##:##", "yyyy-MM-dd HH:mm", "####-##-## ##:##:##", "yyyy-MM-dd HH:mm:ss", "##:##", "HH:mm"))
    If IsDate(mstrInputB) Then
        mskDate(0).Text = mstrInputB
    End If
    If IsDate(mstrInputE) Then
        mskDate(1).Text = mstrInputE
    End If
    
    cboFilter.ListIndex = IIf(mintFindType = -1, 0, mintFindType)
    
    If Val(zlDatabase.GetPara("使用个性化风格")) = 1 Then
        dkpMain.LoadStateFromString GetSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name & dkpMain.PanesCount, "")
        Call RestoreWinState(Me, App.ProductName, , True)
    End If
    mlngSource = IIf(mbytSize = 1, 0, 999)
    
    If Not objFile.FolderExists(App.Path & "\CertImg") Then
        objFile.CreateFolder App.Path & "\CertImg"
    End If
End Sub

Public Sub SetFontSize(ByVal bytSize As Byte)
'功能:设置医嘱清单的字体大小
'入参:bytSize：0-小(缺省)，1-大
    LockWindowUpdate Me.hwnd
    mbytSize = IIf(bytSize = 0, 9, 12)
    Call zlControl.SetPubFontSize(Me, bytSize)
    Call grid.SetFontSize(vfgCert, mbytSize)
    If gbln启用结构化地址 Then
        patiAdressInfo(ADRS_出生地点).Font.Size = mbytSize
        patiAdressInfo(ADRS_住址).Font.Size = mbytSize
        patiAdressInfo(ADRS_陪诊人住址).Font.Size = mbytSize
    End If
    Call Form_Resize
    LockWindowUpdate 0
End Sub

Private Sub MainDefCommandBar()
'功能：主窗口菜单定义部份
'说明：
'1.其中固有的菜单和按钮必须有，作为子窗体处理菜单的基准
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl

    '菜单定义
    '-----------------------------------------------------
    Set objMenu = cmbMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_Certify_FilePopup, "文件(&F)", -1, False) '固有
    objMenu.ID = conMenu_Certify_FilePopup '对xtpControlPopup类型的命令ID需重新赋值
    With objMenu.CommandBar.Controls
'        Set objControl = .Add(xtpControlButton, conMenu_CertifyFile_PrintSet, "打印设置(&S)…") '固有
'        Set objControl = .Add(xtpControlButton, conMenu_CertifyFile_Preview, "预览(&V)") '固有
'        objControl.IconId = 7
'        Set objControl = .Add(xtpControlButton, conMenu_CertifyFile_Print, "打印(&P)") '固有
'        objControl.IconId = 8
        Set objControl = .Add(xtpControlButton, conMenu_Certify_ParSet, "设备配置(&S)")
        Set objControl = .Add(xtpControlButton, conMenu_Certify_Quit, "退出(&X)"): objControl.BeginGroup = True '固有
        objControl.IconId = 9
    End With
    
    Set objMenu = cmbMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_Certify_Edit, "编辑(&E)", -1, False)
    objMenu.ID = conMenu_Certify_Edit '对xtpControlPopup类型的命令ID需重新赋值
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Certify_Add, "新增")
        objControl.IconId = 1
        Set objControl = .Add(xtpControlButton, conMenu_Certify_Modify, "修改")
        objControl.IconId = 2
        Set objControl = .Add(xtpControlButton, conMenu_Certify_Verify, "审核")
        objControl.IconId = 4
        Set objControl = .Add(xtpControlButton, conMenu_Certify_Stop, "停用")
        objControl.IconId = 3
        Set objControl = .Add(xtpControlButton, conMenu_Certify_Start, "取消停用"): objControl.BeginGroup = True
        objControl.IconId = 5
    End With
    
    Set objMenu = cmbMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_Certify_ViewPopup, "查看(&V)", -1, False) '固有
    objMenu.ID = conMenu_Certify_ViewPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_CertifyView_ToolBar, "工具栏(&T)") '固有
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_CertifyView_ToolBar_Button, "标准按钮(&S)", -1, False '固有
            .Add xtpControlButton, conMenu_CertifyView_ToolBar_Text, "文本标签(&T)", -1, False '固有
            .Add xtpControlButton, conMenu_CertifyView_ToolBar_Size, "大图标(&B)", -1, False '固有
        End With
        
        Set objControl = .Add(xtpControlButton, conMenu_CertifyView_StatusBar, "状态栏(&S)") '固有
        objControl.BeginGroup = True
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_CertifyView_FontSize, "字体大小(&N)") '固有
        objPopup.BeginGroup = True
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_CertifyView_FontSize_S, "小字体(&S)", -1, False '固有
            .Add xtpControlButton, conMenu_CertifyView_FontSize_L, "大字体(&L)", -1, False '固有
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Certify_Refresh, "刷新(&F5)") '固有
        objControl.IconId = 11
        objControl.BeginGroup = True
    End With
    
    Set objMenu = cmbMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_CertifyHelpPopup, "帮助(&H)", -1, False) '固有
    objMenu.ID = conMenu_CertifyHelpPopup
    With objMenu.CommandBar.Controls
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_CertifyHelp_Web, "&WEB上的" & gstrProductName) '固有
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_CertifyHelp_Web_Home, gstrProductName & "主页(&H)", -1, False '固有
        End With
        Set objControl = .Add(xtpControlButton, conMenu_CertifyHelp_About, "关于(&A)…"): objControl.BeginGroup = True '固有
    End With

    '工具栏定义:包括公共部份
    '-----------------------------------------------------
    Set objBar = cmbMain.Add("工具栏", xtpBarTop)
    With objBar.Controls
'        Set objControl = .Add(xtpControlButton, conMenu_CertifyFile_Print, "打印") '固有
'        objControl.IconId = 8
'        Set objControl = .Add(xtpControlButton, conMenu_CertifyFile_Preview, "预览") '固有
        objControl.IconId = 7
        Set objControl = .Add(xtpControlButton, conMenu_Certify_Add, "新增")
        objControl.IconId = 1
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Certify_Modify, "修改")
        objControl.IconId = 2
        Set objControl = .Add(xtpControlButton, conMenu_Certify_Verify, "审核")
        objControl.IconId = 3
        Set objControl = .Add(xtpControlButton, conMenu_Certify_Start, "启用")
        objControl.IconId = 4
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Certify_Stop, "停用")
        objControl.IconId = 5
        Set objControl = .Add(xtpControlButton, conMenu_Certify_Record, "实名认证记录")
        objControl.IconId = 6
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Certify_Identify_Set, "认证接口配置")
        objControl.IconId = 12
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Certify_Refresh, "刷新") '固有
        objControl.IconId = 11
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_CertifyHelp_Help, "帮助") '固有
        objControl.IconId = 10
        objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Certify_Quit, "退出") '固有
        objControl.IconId = 9
    End With
    
    For Each objControl In objBar.Controls
      If objControl.type = xtpControlButton Then
          objControl.Style = xtpButtonIconAndCaption
      End If
    Next
        
    With cmbMain.KeyBindings
        .Add 0, vbKeyF5, conMenu_Certify_Refresh
        .Add 0, vbKeyF1, conMenu_CertifyHelp_Help
    End With
    
End Sub

Private Function DrawLin(Optional ByVal bytModel As Byte = 0)
'功能：给控件画线
    Dim objText As Object
    
    For Each objText In Me.Controls
        If TypeName(objText) = "TextBox" Or TypeName(objText) = "Frame" Then
            If objText.Name <> "txtAdressInfo" Then
                DrawLineCTL objText
            ElseIf objText.Name = "txtAdressInfo" Then
                If Not gbln启用结构化地址 Then
                    DrawLineCTL objText
                End If
            End If
        End If
    Next
End Function

Private Function UpdateCertSate() As Boolean
'功能：点击启用停用时更新实名信息状态
    Dim strSql As String
    Dim blnTrans As Boolean
    On Error GoTo errH:
    strSql = "Zl_病人实名信息_状态_Update(1," & mlng实名id & "," & mlng病人ID & ",Null," & IIf(mblnStop, 1, 0) & ")"
    Screen.MousePointer = 11
    gcnOracle.BeginTrans: blnTrans = True
    zlDatabase.ExecuteProcedure strSql, "更新认证状态"
    gcnOracle.CommitTrans: blnTrans = False
    Screen.MousePointer = 0
    Exit Function
errH:
    Screen.MousePointer = 0
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub InitVsfGridHeader()
    Dim strHeader As String
    strHeader = "证件ID;证件号码,2500,1;证件类型,2500,1;备注,4000,1;所有者,2000,4;图片,900,4"
    Call grid.Init(vfgCert, strHeader, , , 1)
    
    strHeader = "接口ID;名称,3000,1;部件名;说明,6000,1;认证结果,2900,4"
    Call grid.Init(vsfInterface, strHeader)
End Sub

Public Function zlRefresh(ByVal lng病人ID As Long, ByVal lng实名ID As Long, ByVal blnEdit As Boolean) As Boolean
'功能：公共接中，用于刷新

    mlng病人ID = lng病人ID
    mblnEdit = blnEdit
    If lng实名ID = 0 Then
        mblnEdit = True
    End If
    Call ClearPatiInfo
    If lng实名ID <> 0 Then
        Screen.MousePointer = 11
        Call LoadPatiInfo(lng实名ID)
        Screen.MousePointer = 0
    End If
    Call SetCtlBackColor(mblnEdit)
End Function
Private Sub LoadPatiInfo(ByVal lng实名ID As Long)
'功能：加载病人详细信息
    Dim rsPati As New ADODB.Recordset
    Dim rsPatiCert As New ADODB.Recordset
    Dim rsPatiInterface As New ADODB.Recordset
    
    Set rsPati = LoadPatiInfoByID(lng实名ID)
    Set rsPatiCert = LoadPatiCert(0, lng实名ID)
    Set rsPatiInterface = LoadCertInterface(0, lng实名ID)
    If Not rsPati.EOF Then
        SetCtlValues rsPati, 0
    End If
    If Not rsPatiCert.EOF Then
        SetCtlValues rsPatiCert, 1
    End If
    If Not rsPatiInterface.EOF Then
        SetCtlValues rsPatiInterface, 2
    End If
End Sub

Private Function SetCtlValues(ByVal rsTmp As ADODB.Recordset, ByVal intTYPE As Integer)
  '功能：给控件赋值
    Dim objCtl As Object
    Dim intIndex As Integer
    Dim strValue As String, strFMT As String, strPictrue As String
    Dim i As Long
    
    If intTYPE = 0 Then
        If Not rsTmp.EOF Then
             For Each objCtl In Me.Controls
                Select Case objCtl.Name
                    Case "txtInfo"
                        Select Case objCtl.Index
                            Case TXT_姓名
                                objCtl.Text = rsTmp!姓名 & ""
                            Case TXT_陪诊人姓名
                                objCtl.Text = rsTmp!陪诊人姓名 & ""
                            Case TXT_身份证号
                                objCtl.Text = rsTmp!身份证号 & ""
                            Case TXT_陪诊人身份证号
                                objCtl.Text = rsTmp!陪诊人身份证号 & ""
                            Case txt_手机号
                                objCtl.Text = rsTmp!手机号 & ""
                            Case TXT_备注
                                objCtl.Text = rsTmp!备注 & ""
                        End Select
                    Case "cboInfo"
                        Select Case objCtl.Index
                            Case CBO_国籍
                                intIndex = cbo.FindIndex(objCtl, "" & rsTmp!国籍)
                                objCtl.ListIndex = intIndex
                            Case CBO_陪诊人国籍
                                intIndex = cbo.FindIndex(objCtl, "" & rsTmp!陪诊人国籍)
                                objCtl.ListIndex = intIndex
                            Case CBO_民族
                                intIndex = cbo.FindIndex(objCtl, "" & rsTmp!民族)
                                objCtl.ListIndex = intIndex
                            Case CBO_陪诊人民族
                                intIndex = cbo.FindIndex(objCtl, "" & rsTmp!陪诊人民族)
                                objCtl.ListIndex = intIndex
                            Case CBO_身份证类型
                                intIndex = Val("" & rsTmp!身份证类型)
                                objCtl.ListIndex = intIndex
                            Case CBO_陪诊人身份证类型
                                intIndex = Val("" & rsTmp!陪诊人身份证类型)
                                objCtl.ListIndex = intIndex
                            Case CBO_性别
                                intIndex = cbo.FindIndex(objCtl, "" & rsTmp!性别)
                                objCtl.ListIndex = intIndex
                            Case CBO_关系
                                intIndex = cbo.FindIndex(objCtl, "" & rsTmp!陪诊人关系)
                                objCtl.ListIndex = intIndex
                        End Select
                    Case "patiAdressInfo"
                        If gbln启用结构化地址 Then
                            Call SetStructAddress(mlng病人ID, 0, objCtl, decode(objCtl.Index, 0, 1, 1, 3, 2, 5))
                            If objCtl.Value = "" Then
                                Select Case objCtl.Index
                                    Case ADRS_出生地点
                                        objCtl.Value = "" & rsTmp!出生地点
                                    Case ADRS_住址
                                        objCtl.Value = "" & rsTmp!住址
                                    Case ADRS_陪诊人住址
                                        objCtl.Value = "" & rsTmp!陪诊人住址
                                End Select
                            End If
                        End If
                    Case "txtAdressInfo"
                        If Not gbln启用结构化地址 Then
                            Select Case objCtl.Index
                                Case ADRS_出生地点
                                    objCtl.Text = "" & rsTmp!出生地点
                                Case ADRS_住址
                                    objCtl.Text = "" & rsTmp!住址
                                Case ADRS_陪诊人住址
                                    objCtl.Text = "" & rsTmp!陪诊人住址
                            End Select
                        End If
                    Case "txt出生日期"
                        strFMT = objCtl.Mask
                        Select Case objCtl.Index
                            Case 0
                                strValue = "" & rsTmp!出生日期
                            Case 1
                                strValue = "" & rsTmp!陪诊人出生日期
                        End Select
                        If IsDate(strValue) Then
                            strValue = Format(strValue, decode(strFMT, "####-##-##", "yyyy-MM-dd", "####-##-## ##:##", "yyyy-MM-dd HH:mm", "####-##-## ##:##:##", "yyyy-MM-dd HH:mm:ss", "##:##", "HH:mm"))
                        Else
                            strValue = Replace(strFMT, "#", "_")
                        End If
                        objCtl.Text = strValue
                End Select
             Next
        End If
    ElseIf intTYPE = 1 Then
        i = vfgCert.FixedRows
        If Not rsTmp.EOF Then
            With vfgCert
                Do While Not rsTmp.EOF
                    .AddItem "", i
                    .TextMatrix(i, COLS_证件ID) = "" & rsTmp!ID
                    .TextMatrix(i, COLS_证件类型) = "" & rsTmp!证件类型
                    .TextMatrix(i, COLS_证件号码) = "" & rsTmp!证件号码
                    .TextMatrix(i, CLOS_备注) = "" & rsTmp!备注
                    .TextMatrix(i, COLS_所有者) = IIf(Val("" & rsTmp!所有者) = 1, "病人本身", "陪诊人")
                    .Cell(flexcpPicture, i, COLS_图片, i, COLS_图片) = img图片
                    .Cell(flexcpPictureAlignment, i, COLS_图片, i, COLS_图片) = 4
                    i = i + 1
                    rsTmp.MoveNext
                Loop
            End With
        End If
    Else
        i = vsfInterface.FixedRows
        If Not rsTmp.EOF Then
            With vsfInterface
                Do While Not rsTmp.EOF
                    .AddItem "", i
                    .TextMatrix(i, INT_部件名) = "" & rsTmp!部件名
                    .TextMatrix(i, INT_接口ID) = "" & rsTmp!ID
                    .TextMatrix(i, INT_名称) = "" & rsTmp!接口名
                    .TextMatrix(i, INT_说明) = "" & rsTmp!说明
                    .TextMatrix(i, INT_认证结果) = IIf(Val("" & rsTmp!认证结果) = 1, "认证成功", "认证失败")
                    i = i + 1
                    rsTmp.MoveNext
                Loop
            End With
        End If
    End If
End Function

Private Sub DrawLineCTL(ByRef objCtl As Object, Optional ByVal bytModel As Byte = 0)
'功能:给指定对象画一条线或清除此原有线条
'objCtl-传入控件对象，根据该控件对象获取对应坐标值
'bytModel=0-画线;1-清除线
    Dim objPic As Object  '容器
    Dim x1 As Long, y1 As Long, x2 As Long, y2 As Long
    
    Select Case TypeName(objCtl)
        Case "TextBox"
            If objCtl.Name <> "txtFilter" Then
                '在每个TextBox 下面画一条线
                x1 = objCtl.Left
                y1 = objCtl.Top + objCtl.Height + 3
                x2 = objCtl.Left + objCtl.Width
                y2 = y1
            End If
        Case "Frame"
            x1 = objCtl.Left
            y1 = objCtl.Top + objCtl.Height + 3
            x2 = objCtl.Left + objCtl.Width - 60
            y2 = y1
    End Select
    Set objPic = objCtl.Container
    objPic.DrawWidth = 1
    If bytModel = 0 Then
        objPic.Line (x1, y1)-(x2, y2)
    Else
        objPic.Line (x1, y1)-(x2, y2), objPic.BackColor '清除线条
    End If
End Sub

Private Sub InitReportColumn()
    Dim objCol As ReportColumn, lngIdx As Long, i As Long

    With rptPati
        '当列顺序或数量(代码或人为隐藏)改变后,要用Find(列号)查找列,但仍可用Record(列号)访问数据行
        Set objCol = .Columns.Add(COL_TAG, "", 20, True)
        Set objCol = .Columns.Add(COL_病人Id, "病人ID", 0, False)
        objCol.Sortable = False: objCol.Visible = False
        Set objCol = .Columns.Add(COL_实名ID, "实名ID", 0, False)
        objCol.Sortable = False: objCol.Visible = False
        Set objCol = .Columns.Add(COL_姓名, "姓名", 75, True)
        Set objCol = .Columns.Add(COL_性别, "性别", 75, True)
        Set objCol = .Columns.Add(COL_出生日期, "出生日期", 106, True)
        Set objCol = .Columns.Add(COL_国籍, "国籍", 75, True): objCol.Visible = False
        Set objCol = .Columns.Add(COL_民族, "民族", 75, True): objCol.Visible = False
        Set objCol = .Columns.Add(COL_身份证号, "身份证号", 106, True)
        Set objCol = .Columns.Add(COL_陪诊人姓名, "陪诊人姓名", 120, True)
        Set objCol = .Columns.Add(COL_陪诊人关系, "陪诊人关系", 120, True)
        Set objCol = .Columns.Add(COL_建档人, "建档人", 75, True)
        Set objCol = .Columns.Add(COL_建档时间, "建档时间", 106, True)
        Set objCol = .Columns.Add(COL_更新人, "更新人", 75, True)
        Set objCol = .Columns.Add(COL_更新时间, "更新时间", 106, True)
        Set objCol = .Columns.Add(COL_手机号, "手机号", 106, True)
        Set objCol = .Columns.Add(COL_备注, "备注", 200, True)
        Set objCol = .Columns.Add(COL_是否认证, "是否认证", 30, True): objCol.Visible = False
        Set objCol = .Columns.Add(COL_是否停用, "是否停用", 75, True): objCol.Visible = False
        Set objCol = .Columns.Add(COL_停用时间, "停用时间", 106, True)
        
        With .PaintManager
            .ColumnStyle = xtpColumnFlat
            .MaxPreviewLines = 1
            .TreeIndent = 0 '有分组列时，树形线边上会再有一根边线
            .GroupForeColor = &HC00000
            .GridLineColor = RGB(225, 225, 225)
            .VerticalGridStyle = xtpGridSolid
            .NoGroupByText = "拖动列标题到这里,按该列分组..."
            .NoItemsText = "没有可显示的病人..."
            .HighlightBackColor = &HFFA000
'            .HighlightForeColor = vbWhite
        End With
        .PreviewMode = True
        .AllowColumnRemove = False
        .MultipleSelection = False '会引发SelectionChanged事件
        .ShowItemsInGroups = False
        .SetImageList Me.imgIcons

        '分组之后可能失去记录集中的顺序,因此强行加入排序列
        .SortOrder.Add .Columns(COL_病人Id)
        .SortOrder(0).SortAscending = True
        .SortOrder.Add .Columns(COL_姓名)
        .SortOrder(1).SortAscending = True
        .SortOrder.Add .Columns(COL_陪诊人姓名)
        .SortOrder(2).SortAscending = True
    End With
    
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    If Item.ID = 1 Then
        Item.Handle = picPati.hwnd
    ElseIf Item.ID = 2 Then
        Item.Handle = picFormation.hwnd
    End If
End Sub

Private Sub Form_Resize()
    If Me.WindowState = vbMinimized Then Exit Sub
    Call cmbMain_Resize
End Sub

Private Sub cmbMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbBar.Visible Then Bottom = Me.stbBar.Height
End Sub

Private Sub cmbMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long

    Call Me.cmbMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)

    On Error Resume Next
    With Me.picFormation
        .Left = lngLeft: .Top = lngTop
        .Width = lngRight - lngLeft
        .Height = lngBottom - lngTop
    End With
    
    vsbMain.Top = 0
    vsbMain.Left = picFormation.ScaleWidth - vsbMain.Width
    vsbMain.Height = picFormation.ScaleHeight
    vsbMain.Width = GetSystemMetrics(SM_CXVSCROLL) * Screen.TwipsPerPixelX
    vsbMain.Max = (picFormation.ScaleHeight - picDetailInfo.Height) / Screen.TwipsPerPixelY - 200
    vsbMain.Min = 0
    vsbMain.SmallChange = 5
    vsbMain.LargeChange = 50
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call zlDatabase.SetPara("实名信息查询方式", IIf(mbln精确, "0", "1"), glngSys, glngModul, IIf(InStr(gstrPrivs, ";参数设置;") > 0, True, False))
    Call zlDatabase.SetPara("实名信息查询条件", "" & mintFindType, glngSys, glngModul, IIf(InStr(gstrPrivs, ";参数设置;") > 0, True, False))
    Call zlDatabase.SetPara("实名信息时间范围", mstrInputB & ";" & mstrInputE, glngSys, glngModul, IIf(InStr(gstrPrivs, ";参数设置;") > 0, True, False))
    If Val(zlDatabase.GetPara("使用个性化风格")) = 1 Then
        Call SaveSetting("ZLSOFT", "私有模块\" & gstrDBUser & "\界面设置\" & App.ProductName & "\" & Me.Name & "\" & TypeName(dkpMain), dkpMain.Name & dkpMain.PanesCount, dkpMain.SaveStateToString)
        Call SaveWinState(Me, App.ProductName)
    End If
    mlng病人ID = 0
    mlng实名id = 0
    mblnEdit = False
    mbln精确 = False
    mstrFindType = ""
    mintFindType = 0
    mstrFilter = ""
    mstrInput = ""
    mstrInputB = ""
    mstrInputE = ""
    mintDate = 0
    mblnStop = False
    mbytSize = 0
    mlngSource = 0
    mlngTopVsc = 0
    mstrPrePati = ""
    Set mrsCertType = Nothing
End Sub

Private Sub lblType_Click()
    mbln精确 = Not mbln精确
    lblType.Caption = IIf(mbln精确, " 精 确↓", " 范 围↓")
    SetCtlEnable
    Call RefreshFilter
End Sub

Private Sub RefreshFilter()
    Dim strName As String
    Dim objCtl As Object
    Dim strTmp As String
    Dim strFilter As String
    Dim strInput As String

    If mbln精确 Then
        strName = Trim(cbostrFilter.Text)
        If mstrFindType Like "姓名" Then
            strFilter = IIf(chkOption.Value = 1, "a.陪诊人姓名=[1]", "a.姓名=[1]")
            strInput = strName
        ElseIf mstrFindType Like "二代身份证" Then
            strFilter = IIf(chkOption.Value = 1, "a.陪诊人身份证号=[1]", "a.身份证号=[1]")
            strInput = strName
        ElseIf mstrFindType Like "其它证件类型" Then
            If chkOption.Value = 1 Then
                strFilter = "b.证件类型=[1] And b.所有者=2"
            Else
                strFilter = "b.证件类型=[1] And b.所有者=1"
            End If
            strInput = strName
        ElseIf mstrFindType Like "证件号码" Then
            If chkOption.Value = 1 Then
                strFilter = "b.证件号码=[1] And b.所有者=2"
            Else
                strFilter = "b.证件号码=[1] And b.所有者=1"
            End If
            strInput = strName
        ElseIf mstrFindType Like "认证状态" Then
            strFilter = "a.认证状态=[1]"
            strInput = Val(strName)
        ElseIf mstrFindType Like "手机号" Then
            strFilter = "a.手机号=[1]"
            strInput = Val(strName)
        End If
        If strFilter <> "" Then
            mstrFilter = strFilter
            mstrInput = strInput
        End If
    Else
        If Not IsDate(mskDate(0).Text) Then
            Set objCtl = mskDate(0)
            strTmp = "开始时间"
        ElseIf Not IsDate(mskDate(1).Text) Then
            Set objCtl = mskDate(1)
            strTmp = "结束时间"
        End If
        If strTmp = "" Then
            mstrFilter = "a.建档时间 Between TO_Date([1],'yyyy-mm-dd hh24:mi:ss') And To_Date([2],'yyyy-mm-dd hh24:mi:ss')"
            mstrInputB = Format(mskDate(0).Text, "YYYY-MM-DD HH:MM:SS")
            mstrInputE = Format(mskDate(1).Text, "YYYY-MM-DD HH:MM:SS")
        End If
    End If
End Sub

Private Sub monInfo_DateClick(ByVal DateClicked As Date)
'功能：monInfo_DateClick
    Dim strDate As String, strFMT As String
    Dim objMSK As MaskEdBox

    Set objMSK = mskDate(mintDate)
    '获取时分秒数据
    If objMSK.MaxLength >= Len("####-##-## ##:##") Then
        'yyyy-MM-dd HH:mm:ss 格式时间
        If objMSK.MaxLength > Len("####-##-## ##:##") Then
            strFMT = "HH:mm:ss"
        Else
            'yyyy-MM-dd HH:mm 格式时间
            strFMT = "HH:mm"
        End If
        '原时间是时间类型，这取该时间的时分秒数据，否则取当前时间的时分秒
        If IsDate(objMSK.Text) Then
            strDate = " " & Format(objMSK.Text, strFMT)
        Else
            strDate = " " & Format(zlDatabase.Currentdate, strFMT)
        End If
    End If
    '获取时间
    strDate = Format(DateClicked, "yyyy-MM-dd") & strDate
    objMSK.Text = strDate
    mskDate(objMSK.Index).Text = objMSK.Text
    monInfo.Visible = False
    zlControl.ControlSetFocus objMSK
End Sub

Private Sub monInfo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf KeyCode = vbKeyEscape Then
        monInfo.Visible = False
    End If
End Sub

Private Sub monInfo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call monInfo_DateClick(monInfo.Value)
    ElseIf KeyAscii = vbKeyEscape Then
        monInfo.Visible = False
    End If
End Sub

Private Sub monInfo_Validate(Cancel As Boolean)
    monInfo.Visible = False
End Sub

Private Sub mskDate_GotFocus(Index As Integer)
'功能：MskDateInfo_GotFocus
    zlCommFun.OpenIme False
End Sub

Private Sub picDetailInfo_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    vsbMain.SetFocus
End Sub

Private Sub picFormation_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    vsbMain.SetFocus
End Sub

Private Sub picFormation_Resize()
    Dim x1 As Long, x2 As Long

    picDetailInfo.Top = picFormation.ScaleTop
    x1 = picFormation.ScaleWidth / 2 + picFormation.ScaleLeft
    x2 = picDetailInfo.Width / 2
    picDetailInfo.Left = x1 - x2

    vsbMain.Move Me.ScaleWidth - vsbMain.Width, Me.ScaleTop + 930, vsbMain.Width, picFormation.ScaleHeight + picFormation.ScaleTop
    vsbMain.LargeChange = 100
    vsbMain.SmallChange = vsbMain.LargeChange / 2

    Call vsbMain_Change
End Sub

Private Function SetCtlBackColor(ByVal blnEdit As Boolean)
'功能：给控件设置背景色
    Dim objCtl As Object
    For Each objCtl In Me.Controls
        If InStr("," & ",txtInfo,cboInfo,txtAdressInfo,", "," & objCtl.Name & ",") > 0 Then
            objCtl.BackColor = IIf(blnEdit, vbButtonFace, vbWindowBackground)
            objCtl.Locked = True
        ElseIf InStr(",txt陪诊人出生日期,txt陪诊人出生时间,txt出生日期,txt出生时间,", "," & objCtl.Name & ",") > 0 Then
            objCtl.Enabled = False
            objCtl.BackColor = IIf(blnEdit, vbButtonFace, vbWindowBackground)
        ElseIf InStr(",patiAdressInfo,", "," & objCtl.Name & ",") > 0 Then
            objCtl.BackColor = IIf(blnEdit, vbButtonFace, vbWindowBackground)
            objCtl.TextBackColor = IIf(blnEdit, vbButtonFace, vbWindowBackground)
            objCtl.ControlLock = True
        ElseIf InStr(",lblFeild,vfgCert,vsfInterface,", "," & objCtl.Name & ",") > 0 Then
            objCtl.BackColor = IIf(blnEdit, vbButtonFace, vbWindowBackground)
        ElseIf InStr(",lblType,picIdentify,chkOption,lbl建档,mskDate,txtBeginHour,txtDate,", "," & objCtl.Name & ",") > 0 Then
            objCtl.BackColor = &HFFFFF0
        End If
    Next
End Function

Private Sub picPati_Resize()
    Dim lngWidth As Long
    Dim i As Long
    On Error Resume Next
    
    picFilter.Top = picPati.ScaleTop
    picFilter.Left = picPati.ScaleLeft
    picFilter.Width = picPati.Width
    
    rptPati.Move picPati.ScaleLeft, picFilter.Top + picFilter.Height + 10, picFilter.ScaleWidth, picPati.ScaleHeight - picFilter.Top - picFilter.Height
End Sub

Private Sub InitCboData()
'功能：给下拉列表添加数据
    Dim strSql As String
    Dim rsTmp As New ADODB.Recordset
    
    strSql = "Select RowNum As ID, 编码, 简码, 名称, 缺省标志 缺省, '证件类型' 表名 From 证件类型"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, "读取身份证类型")
    
    If Not rsTmp.EOF Then
        Set mrsCertType = rsTmp
    End If
    
    Call SetCboFromRec(Array("国籍", "民族", "性别", "社会关系"), Array(CBO_国籍, CBO_民族, CBO_性别, CBO_关系))
    Call SetCboFromRec(Array("国籍", "民族", "性别"), Array(CBO_陪诊人国籍, CBO_陪诊人民族, CBO_陪诊人性别))
    
    Call SetCboFromList(Array("", "0-居民身份证", "1-港澳台居住证", "2-外国人居留证"), Array(CBO_身份证类型))
    Call SetCboFromList(Array("", "0-居民身份证", "1-港澳台居住证", "2-外国人居留证"), Array(CBO_陪诊人身份证类型))
    
    cboFilter.Clear
    cboFilter.AddItem "姓名"
    cboFilter.AddItem "二代身份证"
    cboFilter.AddItem "其它证件类型"
    cboFilter.AddItem "证件号码"
    cboFilter.AddItem "认证状态"
    cboFilter.AddItem "手机号"
End Sub

Private Sub SetCtlEnable()
'功能：设置控件的可见性
    
    If mbln精确 Then
        picCboFilter.Enabled = True
        cboFilter.Enabled = True
'        txtFilter.Enabled = True
        picstrFilter.Enabled = True
        cbostrFilter.Enabled = True
        chkOption.Enabled = True
        mskDate(0).Enabled = False
        mskDate(1).Enabled = False
        txtDate(0).Enabled = False
        txtDate(1).Enabled = False
        lbl建档.Enabled = False
        cmdDate(0).Enabled = False
        cmdDate(1).Enabled = False
        cmdFilter.Enabled = False
    Else
        picCboFilter.Enabled = False
        cboFilter.Enabled = False
'        txtFilter.Enabled = False
        picstrFilter.Enabled = False
        cbostrFilter.Enabled = False
        chkOption.Enabled = False
        mskDate(0).Enabled = True
        mskDate(1).Enabled = True
        txtDate(0).Enabled = True
        txtDate(1).Enabled = True
        lbl建档.Enabled = True
        cmdDate(0).Enabled = True
        cmdDate(1).Enabled = True
        cmdFilter.Enabled = True
    End If
End Sub
'
Private Sub SetCboFromRec(ByVal arrTab As Variant, ByVal arrCboIdx As Variant, Optional ByVal strAddBeginItems As String = "NULL")
'功能：给下拉列表添加数据
    Dim i As Long, j As Long
    Dim objCboTmp As ComboBox
    Dim arrItem As Variant
    Dim rsTmp As ADODB.Recordset

    For i = 0 To UBound(arrTab)
        Set rsTmp = GetCboData(arrTab(i))
        If Not rsTmp.EOF Then
            Set objCboTmp = cboInfo(arrCboIdx(i))
                objCboTmp.Clear
            If strAddBeginItems <> "NULL" Then
                arrItem = Split(strAddBeginItems, ",")
                For j = LBound(arrItem) To UBound(arrItem)
                    objCboTmp.AddItem arrItem(j)
                Next
            End If
            For j = 1 To rsTmp.RecordCount
                If IsNull(rsTmp!编码) Then
                    objCboTmp.AddItem rsTmp!名称
                Else
                    objCboTmp.AddItem rsTmp!编码 & "-" & rsTmp!名称
                End If
                objCboTmp.ItemData(objCboTmp.NewIndex) = Nvl(rsTmp!ID, 0)
                If Val(rsTmp!缺省 & "") = 1 Then
                    Call zlControl.CboSetIndex(objCboTmp.hwnd, objCboTmp.NewIndex)
                    objCboTmp.Tag = objCboTmp.NewIndex
                End If
                rsTmp.MoveNext
            Next
        End If
    Next
End Sub

Private Sub SetCboFromList(ByVal arrList As Variant, ByVal arrCboIdx As Variant, Optional ByVal intDefault As Integer = -1)
'功能：将指定数据装入指定ComboBox
'参数：arrList=List String数组
'      arrCboIdx=ComboBox索引数组,多个ComboBox时,装入数据相同
'      intDefaut=缺省索引
    Dim i As Long, j As Long

    For i = 0 To UBound(arrCboIdx)
        cboInfo(arrCboIdx(i)).Clear
        For j = 0 To UBound(arrList)
            cboInfo(arrCboIdx(i)).AddItem arrList(j)
        Next
        cboInfo(arrCboIdx(i)).ListIndex = intDefault '缺省为未选中
    Next
End Sub

Private Function LoadPatients(ByVal intTYPE As Integer, Optional lng实名ID As Long) As Boolean
'功能：加载病人实名信息列表数据
    Dim strSql As String
    Dim objRecord As ReportRecord
    Dim rsPati As New ADODB.Recordset
    Dim j As Long, i As Long
    Dim lngPatiRow As Long
    Dim strPatiRow As String
    Dim lngRow As Long
    Dim objRow As ReportRow
    
    If intTYPE = 1 Then
        If mbln精确 Then
            Set rsPati = LoadCertifyPatients(mstrFilter, mstrInput)
        Else
            Set rsPati = LoadCertifyPatients(mstrFilter, mstrInputB, mstrInputE)
        End If
    Else
       Set rsPati = LoadCertifyPatients("", "")
    End If
    If Not rsPati.EOF Then
        '记录现在选中的病人
        If lng实名ID = 0 Then
            If rptPati.SelectedRows.Count > 0 Then
                If Not rptPati.SelectedRows(0).GroupRow Then
                    If rptPati.SelectedRows(0).Record.Tag <> "" Then
                        lngPatiRow = rptPati.SelectedRows(0).Index '用于快速重新定位
                        strPatiRow = rptPati.SelectedRows(0).Record.Tag
                    End If
                End If
            End If
        End If
    End If
    rptPati.Records.DeleteAll
    Do While Not rsPati.EOF
        Set objRecord = rptPati.Records.Add()
        For j = 0 To rptPati.Columns.Count - 1
            objRecord.AddItem ""
        Next
        With objRecord
            .Item(COL_TAG).Value = ""
            .Item(COL_病人Id).Value = "" & rsPati!病人ID
            .Item(COL_实名ID).Value = "" & rsPati!实名ID
            objRecord.Tag = Nvl("" & rsPati!实名ID, "" & rsPati!病人ID)
            .Item(COL_姓名).Value = "" & rsPati!姓名
            .Item(COL_性别).Value = "" & rsPati!性别
            .Item(COL_身份证号).Value = "" & rsPati!身份证号
            .Item(COL_出生日期).Value = Format("" & rsPati!出生日期, "yyyy-MM-dd HH:mm")
            .Item(COL_国籍).Value = "" & rsPati!国籍
            .Item(COL_民族).Value = "" & rsPati!民族
            .Item(COL_陪诊人姓名).Value = "" & rsPati!陪诊人姓名
            .Item(COL_陪诊人关系).Value = "" & rsPati!陪诊人关系
            .Item(COL_手机号).Value = "" & rsPati!手机号
            If Val("" & rsPati!认证状态) = 0 Then
                .Item(COL_姓名).Icon = imgIcons.ListImages("Certify_StateFalse").Index - 1
                .Item(COL_是否认证).Value = Val("" & rsPati!认证状态)
            ElseIf Val("" & rsPati!认证状态) = 1 Then
                .Item(COL_姓名).Icon = imgIcons.ListImages("Certify_StateSure").Index - 1
                .Item(COL_是否认证).Value = Val("" & rsPati!认证状态)
            End If
            .Item(COL_建档人).Value = "" & rsPati!建档人
            .Item(COL_建档时间).Value = Format("" & rsPati!建档时间, "yyyy-MM-dd HH:mm")
            If Val("" & rsPati!是否停用) = 1 Then
                .Item(COL_姓名).Icon = imgIcons.ListImages("Certify_StateStop").Index - 1
            End If
            .Item(COL_更新人).Value = "" & rsPati!更新人
            .Item(COL_更新时间).Value = Format("" & rsPati!更新时间, "yyyy-MM-dd HH:mm")
            .Item(COL_是否停用).Value = Val("" & rsPati!是否停用)
            .Item(COL_停用时间).Value = Format("" & rsPati!停用时间, "yyyy-MM-dd HH:mm")
            .Item(COL_备注).Value = "" & rsPati!备注
        End With
        rsPati.MoveNext
    Loop
    rptPati.Populate
    i = rptPati.Records.Count
    stbBar.Panels(2).Text = "共" & i & "个病人"
    
    If lng实名ID <> 0 Then
        For j = 0 To rptPati.Records.Count - 1
            If Val(rptPati.Rows(j).Record(COL_实名ID).Value) = lng实名ID Then
                lngRow = j
            End If
        Next
    End If
     '定位病人行:在Populate之后
    mstrPrePati = ""
    If rptPati.Rows.Count = 0 Or rsPati.RecordCount > 1 And lngPatiRow = 0 And strPatiRow = "" Then
        If rsPati.RecordCount > 1 Then
            If lngRow = 0 Then
                Set objRow = rptPati.Rows(0)
                Set rptPati.FocusedRow = objRow
            Else
                Set objRow = rptPati.Rows(lngRow)
                Set rptPati.FocusedRow = objRow
            End If
        Else
            Call ClearPatiInfo
        End If
    Else
        '取指定病人行
        If strPatiRow <> "" Then
            '先快速定位
            If lngPatiRow <= rptPati.Rows.Count - 1 Then
                If Not rptPati.Rows(lngPatiRow).GroupRow Then
                    If rptPati.Rows(lngPatiRow).Record.Tag = strPatiRow Then
                        Set objRow = rptPati.Rows(lngPatiRow)
                    End If
                End If
            End If
            '再进行查找
            If objRow Is Nothing Then
                For i = 0 To rptPati.Rows.Count - 1
                    If Not rptPati.Rows(i).GroupRow Then
                        If rptPati.Rows(i).Record.Tag = strPatiRow Then
                            Set objRow = rptPati.Rows(i): Exit For
                        End If
                    End If
                Next
            End If
        End If
        '取第一个非分组行
        If objRow Is Nothing Then
            For i = 0 To rptPati.Rows.Count - 1
                If Not rptPati.Rows(i).GroupRow And rptPati.Rows(i).Childs.Count = 0 Then Set objRow = rptPati.Rows(i): Exit For
            Next
        End If
        Set rptPati.FocusedRow = objRow '该行选中且显示在可见区域,并引发SelectionChanged事件
    End If
End Function

Private Sub rptPati_SelectionChanged()
    Dim blnEdit As Boolean
    Dim strTag As String, strCurPati As String
    If rptPati.SelectedRows.Count <= 0 Then Exit Sub
    With rptPati.SelectedRows(0)
        strCurPati = .Record.Tag
        If strCurPati = mstrPrePati Then Exit Sub
        strTag = mstrPrePati
        mstrPrePati = strCurPati
        mlng病人ID = Val(.Record(COL_病人Id).Value)
        mlng实名id = Val(.Record(COL_实名ID).Value)
        blnEdit = IIf(mlng实名id <> 0, False, True)
        Call zlRefresh(mlng病人ID, mlng实名id, blnEdit)
    End With
End Sub

Private Sub ClearPatiInfo()
'功能:清除界面上病人的基本信息
    Dim objCtl As Object
    Dim i As Long, j As Long
    
    For Each objCtl In Me.Controls
        Select Case objCtl.Name
            Case "txtInfo"
                objCtl.Text = ""
            Case "txtAdressInfo"
                objCtl.Text = ""
            Case "cboInfo"
                If objCtl.Tag <> "" Then '恢复默认值
                    objCtl.ListIndex = Val(objCtl.Tag)
                Else
                    objCtl.ListIndex = -1
                End If
            Case "patiAdressInfo"
                objCtl.Value = ""
            Case "vfgCert"
                i = vfgCert.FixedRows: j = vfgCert.Rows - 1
                Do While i <= j
                    vfgCert.RemoveItem i
                    j = vfgCert.Rows - 1
                Loop
            Case "vsfInterface"
                i = vsfInterface.FixedRows: j = vsfInterface.Rows - 1
                Do While i <= j
                    vsfInterface.RemoveItem i
                    j = vsfInterface.Rows - 1
                Loop
        End Select
    Next
End Sub

Private Sub vfgCert_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim lngRow As Long, lngCol As Long
    Dim vPoint As POINTAPI
    
    vPoint = GetCoordPos(vfgCert.hwnd, vfgCert.Left + vfgCert.Width - 3500, vfgCert.CellTop)
    With vfgCert
        lngRow = .Row
        lngCol = .Col
        If lngCol = COLS_图片 Then
            Set rsTmp = GetCertID(Val(.TextMatrix(lngRow, COLS_证件ID)))
            If rsTmp.EOF Then
                MsgBox "该病人没有证件图片信息！", vbInformation, gstrSysName
                Exit Sub
            Else
                frmCertPicture.ShowMe Me, Val(.TextMatrix(lngRow, COLS_证件ID)), 0, vPoint.X, vPoint.Y, vfgCert.Height, rsTmp!序号
            End If
        End If
    End With
End Sub

Private Function ShowMessage(objTmp As Object, ByVal strMsg As String, Optional ByVal blnAsk As Boolean) As VbMsgBoxResult
'功能：显示提示信息并定位在输入项目上
    Dim lngColor As Long
    
    On Error GoTo errH
 
    If UCase(TypeName(objTmp)) <> UCase("VSFlexGrid") Then
        lngColor = objTmp.BackColor: objTmp.BackColor = &HC0C0FF
    Else
        lngColor = objTmp.CellBackColor: objTmp.CellBackColor = &HC0C0FF
        Call objTmp.ShowCell(objTmp.Row, objTmp.Col)
    End If
    
    If Not blnAsk Then
        MsgBox strMsg, vbInformation, gstrSysName
    Else
        ShowMessage = MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
    End If
    If UCase(TypeName(objTmp)) <> UCase("VSFlexGrid") Then
        objTmp.BackColor = lngColor
    Else
        objTmp.CellBackColor = lngColor
    End If
    If objTmp.Enabled And objTmp.Visible Then
        If TypeName(objTmp) = "TextBox" Then zlControl.TxtSelAll objTmp
        objTmp.SetFocus
    End If
    Me.Refresh
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub vsbMain_Change()
    Call vsbMain_Scroll
End Sub

Private Sub vsbMain_Scroll()
    picDetailInfo.Top = vsbMain.Value * Screen.TwipsPerPixelY
End Sub






