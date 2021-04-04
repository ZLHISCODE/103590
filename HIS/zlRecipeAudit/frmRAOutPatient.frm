VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.10#0"; "zlIDKind.ocx"
Begin VB.Form frmRAOutPatient 
   Caption         =   "门诊处方审查"
   ClientHeight    =   8430
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11760
   Icon            =   "frmRAOutPatient.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   8430
   ScaleWidth      =   11760
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '所有者中心
   Begin VB.Timer timPatient 
      Enabled         =   0   'False
      Left            =   1560
      Top             =   120
   End
   Begin VB.PictureBox picAuditDel 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   10800
      Picture         =   "frmRAOutPatient.frx":076A
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   42
      TabStop         =   0   'False
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picAuditYN 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   1
      Left            =   10320
      Picture         =   "frmRAOutPatient.frx":0AF4
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   41
      TabStop         =   0   'False
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox picAuditYN 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   375
      Index           =   0
      Left            =   9840
      Picture         =   "frmRAOutPatient.frx":107E
      ScaleHeight     =   375
      ScaleWidth      =   375
      TabIndex        =   40
      TabStop         =   0   'False
      Top             =   1560
      Visible         =   0   'False
      Width           =   375
   End
   Begin MSComctlLib.ImageList imgPASSLight 
      Left            =   2280
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRAOutPatient.frx":1608
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRAOutPatient.frx":20FA
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRAOutPatient.frx":2BEC
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRAOutPatient.frx":36DE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmRAOutPatient.frx":41D0
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.PictureBox picNothing 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   495
      Left            =   8880
      ScaleHeight     =   495
      ScaleWidth      =   735
      TabIndex        =   30
      TabStop         =   0   'False
      Top             =   360
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.PictureBox picDetail 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5175
      Left            =   4080
      ScaleHeight     =   5175
      ScaleWidth      =   6015
      TabIndex        =   10
      TabStop         =   0   'False
      Top             =   2640
      Width           =   6015
      Begin VB.PictureBox picLine01_S 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   60
         Left            =   120
         MousePointer    =   7  'Size N S
         ScaleHeight     =   60
         ScaleWidth      =   1695
         TabIndex        =   39
         TabStop         =   0   'False
         Top             =   2280
         Width           =   1695
      End
      Begin VB.PictureBox picAudit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2655
         Left            =   120
         ScaleHeight     =   2655
         ScaleWidth      =   5655
         TabIndex        =   13
         TabStop         =   0   'False
         Top             =   2400
         Width           =   5655
         Begin zl9RecipeAudit.SpeedButton sbnSelect 
            Height          =   300
            Left            =   1320
            TabIndex        =   37
            Top             =   120
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   529
            BackColor       =   -2147483646
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   ""
            Enabled         =   -1  'True
            PictureAlign    =   1
            Picture         =   "frmRAOutPatient.frx":4CC2
            ShowCaption     =   0   'False
         End
         Begin VB.PictureBox picYesNO 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   495
            Left            =   2280
            ScaleHeight     =   465
            ScaleWidth      =   2385
            TabIndex        =   20
            TabStop         =   0   'False
            Top             =   1920
            Width           =   2415
            Begin VB.Label lblDisp_Fixed 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               Caption         =   "合格"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   26.25
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H80000008&
               Height          =   525
               Left            =   240
               TabIndex        =   21
               Top             =   120
               Width           =   1050
            End
         End
         Begin VB.PictureBox picFunc 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1215
            Left            =   2160
            ScaleHeight     =   1215
            ScaleWidth      =   3255
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   600
            Width           =   3255
            Begin VB.CommandButton cmdNo_Fixed 
               Caption         =   "不合格(&N)"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   15.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1080
               Left            =   1350
               Picture         =   "frmRAOutPatient.frx":5014
               Style           =   1  'Graphical
               TabIndex        =   19
               Top             =   0
               Width           =   1695
            End
            Begin VB.CommandButton cmdYes_Fixed 
               Caption         =   "合格(&Y)"
               BeginProperty Font 
                  Name            =   "宋体"
                  Size            =   15.75
                  Charset         =   134
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   1080
               Left            =   0
               Picture         =   "frmRAOutPatient.frx":5CDE
               Style           =   1  'Graphical
               TabIndex        =   18
               Top             =   0
               Width           =   1335
            End
         End
         Begin VB.TextBox txtReason 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            Height          =   375
            Left            =   120
            MaxLength       =   500
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   16
            Top             =   600
            Width           =   1575
         End
         Begin VB.PictureBox picLine02 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   60
            Left            =   120
            ScaleHeight     =   60
            ScaleWidth      =   1695
            TabIndex        =   15
            TabStop         =   0   'False
            Top             =   1440
            Width           =   1695
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfAudit 
            Height          =   735
            Left            =   120
            TabIndex        =   14
            Top             =   1800
            Width           =   1455
            _cx             =   1975978502
            _cy             =   1975977232
            Appearance      =   2
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
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   2
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   0
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
         Begin zl9RecipeAudit.SpeedButton sbnHouse 
            Height          =   300
            Left            =   1800
            TabIndex        =   38
            Top             =   120
            Width           =   300
            _ExtentX        =   529
            _ExtentY        =   529
            BackColor       =   -2147483646
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   ""
            Enabled         =   -1  'True
            PictureAlign    =   1
            Picture         =   "frmRAOutPatient.frx":69A8
            ShowCaption     =   0   'False
         End
         Begin VB.Label lblReason 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "综合理由"
            Height          =   180
            Left            =   120
            TabIndex        =   35
            Top             =   120
            Width           =   720
         End
         Begin VB.Label lblAudit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "审查项目"
            Height          =   180
            Left            =   120
            TabIndex        =   33
            Top             =   1560
            Width           =   720
         End
      End
      Begin VB.PictureBox picRecipe 
         Appearance      =   0  'Flat
         BackColor       =   &H80000002&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1095
         Left            =   120
         ScaleHeight     =   1095
         ScaleWidth      =   5655
         TabIndex        =   12
         TabStop         =   0   'False
         Top             =   960
         Width           =   5655
         Begin zl9RecipeAudit.SpeedButton sbnPASS 
            Height          =   300
            Left            =   4080
            TabIndex        =   36
            Top             =   0
            Width           =   1500
            _ExtentX        =   2646
            _ExtentY        =   529
            BackColor       =   -2147483646
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "宋体"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "查看PASS结果"
            Enabled         =   -1  'True
            PictureAlign    =   0
            Picture         =   "frmRAOutPatient.frx":6CFA
            ShowCaption     =   -1  'True
         End
         Begin VB.CheckBox chk24 
            BackColor       =   &H80000002&
            Caption         =   "显示24小时内所有药嘱(&P)"
            Height          =   180
            Left            =   1680
            TabIndex        =   34
            Top             =   120
            Width           =   2380
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfRecipe 
            Height          =   495
            Left            =   120
            TabIndex        =   31
            Top             =   480
            Width           =   1455
            _cx             =   1975978502
            _cy             =   1975976809
            Appearance      =   2
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
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   2
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   0
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
         Begin VB.Label lblRecipe 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "处方明细"
            Height          =   180
            Left            =   120
            TabIndex        =   32
            Top             =   120
            Width           =   720
         End
      End
      Begin VB.TextBox txtOther 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         Height          =   500
         Left            =   120
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   11
         Text            =   "frmRAOutPatient.frx":704C
         Top             =   120
         Width           =   2055
      End
   End
   Begin VB.PictureBox picTab 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1935
      Left            =   3960
      ScaleHeight     =   1935
      ScaleWidth      =   3135
      TabIndex        =   8
      TabStop         =   0   'False
      Top             =   360
      Width           =   3135
      Begin XtremeSuiteControls.TabControl tbcTab 
         Height          =   855
         Left            =   240
         TabIndex        =   9
         Top             =   840
         Width           =   1335
         _Version        =   589884
         _ExtentX        =   2355
         _ExtentY        =   1508
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picRecFinish 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1215
      Left            =   120
      ScaleHeight     =   1215
      ScaleWidth      =   3135
      TabIndex        =   4
      TabStop         =   0   'False
      Top             =   6120
      Width           =   3135
      Begin VSFlex8Ctl.VSFlexGrid vsfRecFinish 
         Height          =   735
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   1455
         _cx             =   1975978502
         _cy             =   1975977232
         Appearance      =   2
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
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
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
   Begin VB.PictureBox picRecWait 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   120
      ScaleHeight     =   1335
      ScaleWidth      =   3135
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   4560
      Width           =   3135
      Begin VB.CheckBox chkPubPatient 
         Caption         =   "含岗位未明确来源科室的待审数据(&I)"
         Height          =   180
         Left            =   120
         TabIndex        =   7
         Top             =   1080
         Width           =   3500
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfRecWait 
         Height          =   735
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   1455
         _cx             =   1975978502
         _cy             =   1975977232
         Appearance      =   2
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
         BackColorBkg    =   -2147483636
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   0
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   0
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
   Begin VB.PictureBox picRec 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   120
      ScaleHeight     =   3495
      ScaleWidth      =   3135
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   720
      Width           =   3135
      Begin VB.TextBox txtPatient 
         Height          =   300
         Left            =   1200
         TabIndex        =   29
         Top             =   1380
         Width           =   1815
      End
      Begin zlIDKind.IDKindNew iknPatient 
         Height          =   300
         Left            =   120
         TabIndex        =   28
         Top             =   1380
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   529
         IDKindStr       =   "姓|姓名|0|0|0|0|0|0;门|门诊号|0|0|0|0|0|0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   9
         FontName        =   "宋体"
         IDKind          =   -1
         ShowPropertySet =   -1  'True
         DefaultCardType =   "0"
         SmallStyle      =   -1  'True
         AllowAutoICCard =   -1  'True
         AllowAutoIDCard =   -1  'True
         BackColor       =   -2147483633
         SaveRegType     =   4
      End
      Begin VB.ComboBox cboDate 
         Appearance      =   0  'Flat
         Height          =   300
         ItemData        =   "frmRAOutPatient.frx":7063
         Left            =   1200
         List            =   "frmRAOutPatient.frx":7065
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   990
         Width           =   1815
      End
      Begin VB.ComboBox cboDrugstore 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   25
         Top             =   600
         Width           =   1815
      End
      Begin VB.ComboBox cboClinic 
         Appearance      =   0  'Flat
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   210
         Width           =   1815
      End
      Begin XtremeSuiteControls.TabControl tbcRec 
         Height          =   855
         Left            =   120
         TabIndex        =   2
         Top             =   1920
         Width           =   1335
         _Version        =   589884
         _ExtentX        =   2355
         _ExtentY        =   1508
         _StockProps     =   64
      End
      Begin VB.Label lblDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "提交时间(&T)"
         Height          =   180
         Left            =   120
         TabIndex        =   26
         Top             =   1020
         Width           =   990
      End
      Begin VB.Label lblDrugstore 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "发药药房(&D)"
         Height          =   180
         Left            =   120
         TabIndex        =   24
         Top             =   630
         Width           =   990
      End
      Begin VB.Label lblClinic 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "来源科室(&F)"
         Height          =   180
         Left            =   120
         TabIndex        =   22
         Top             =   240
         Width           =   990
      End
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   8070
      Width           =   11760
      _ExtentX        =   20743
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmRAOutPatient.frx":7067
            Text            =   "中联软件"
            TextSave        =   "中联软件"
            Key             =   "ZLFLAG"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15875
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   953
            MinWidth        =   88
            Text            =   "数字"
            TextSave        =   "数字"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   88
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
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   240
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeCommandBars.ImageManager imgMain 
      Bindings        =   "frmRAOutPatient.frx":78F9
      Left            =   720
      Top             =   120
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmRAOutPatient.frx":790D
   End
   Begin XtremeDockingPane.DockingPane dkpMain 
      Left            =   1200
      Top             =   240
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmRAOutPatient"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Const MSTR_PATIENT = "姓|姓名|0;门|门诊号|0;身|身份证号|1"

Private Const MSTR_VSF_WAIT As String = _
        "病人,,3,1000|审方ID,,0,0|病人ID,,0,0|挂号单ID,,0,0|主页ID,,0,0|性别,,3,600|年龄,,3,600|门诊号,,3,1000|提交时间,,3,1600" & _
        "|提交科室,,3,1500|提交人,,3,1000|发药药房,,3,1500"
        
Private Const MSTR_VSF_FINISH As String = _
        "审查结果,,3,1000|病人,,3,1000|审方ID,,0,0|病人ID,,0,0|挂号单ID,,0,0|主页ID,,0,0|性别,,3,600|年龄,,3,600|门诊号,,3,1000|提交时间,,3,1600" & _
        "|提交科室,,3,1500|提交人,,3,1000|发药药房,,3,1500|审查时间,,3,1600|审查人,,3,1000"
        
Private Const MSTR_VSF_RECIPE As String = _
        "临床科室,,3,1500|审方ID,,0,0|医嘱ID,,0,0|超量说明,,3,1500|医生嘱托,,3,1500|组序号,,0,0|组数,,0,0|组,,3,300|PASS,,3,600" & _
        "|药品名称,,3,2000|商品名,,3,1500|规格,,3,1500|单位,,3,600|数量,,3,600,n|单量,,3,800,n|相关ID,,0,0|用法,,3,800|频次,,3,800" & _
        "|单价,,3,800,n|应收金额,,3,1000,n|不合格,,0,0|细目ID,,0,0|标志,,0,0"

Private Const MSTR_VSF_AUDIT As String = _
        "药师审查,,3,1000|自动审查,,3,1000|药品,,3,2500|删除,,3,600|编码,,3,1000|简称,,3,2000|项目内容,,3,4000|医嘱ID,,0,0|审查项目ID,,0,0" & _
        "|新增,,0,0"
        
Private WithEvents mobjMipModule As zl9ComLib.clsMipModule      '消息平台
Attribute mobjMipModule.VB_VarHelpID = -1
'Private mobjArchiveMedRec As zlPublicAdvice.clsArchiveMedRec    '电子病历
Private mobjPubAdvice As zlPublicAdvice.clsPublicAdvice         '临床公共类
Private mobjPASS As Object                                      '合理用药接口部件
Private mfrmAMR As Form

Private mlngModule As Long              '模块号
Private mstrPrivs As String             '权限
Private mblnMemory As Boolean           '个性化
Private mblnNeedAudit As Boolean        'True开启审方参数；False未开启审方参数
Private mblnAuditStart As Boolean       'True开启审方事务；False停止审方事务
Private mblnLocking As Boolean          'True锁定；False未锁定
Private mblnSendBeforeAudit As Boolean
Private mblnExit As Boolean
Private mblnEnter As Boolean
Private mstrPCName As String
Private mblnReadCard As Boolean
Private mbytFontSize As Byte
Private mintParaPass As Integer         '合理用药厂商
Private msngY As Single
Private mbytDrugName As Byte            '0-显示通用名，1-显示商品名，2-同时显示通用名和商品名
Private mintItem As Integer             '1-依据《处方点评管理规范》28项；2-依据《处方管理办法》7项
Private mlngAuditID As Long             '审方ID
Private mblnReasonRefresh As Boolean    '防止程序控制txtReason时触发Change事件
Private mblnRecipeSendAuto As Boolean   '处方自动发送

Private Sub cboClinic_Click()
    If Me.Visible = False Then Exit Sub
    chkPubPatient.Enabled = cboClinic.ListIndex = 0
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim objControl As CommandBarControl
    Dim i As Integer
    Dim l As Long, lngAuditID As Long
    Dim strTmp As String
    
    Select Case Control.ID
        Case enuMenus.打印设置
            Call zlPrintSet
        Case enuMenus.打印预览, enuMenus.打印, enuMenus.输出Excel
            Dim objTmp As Object
            Dim strTitle As String
            
            Set objTmp = Me.ActiveControl
            If TypeName(objTmp) = "VSFlexGrid" Then
                objTmp.Redraw = False
                If UCase(objTmp.Name) = "VSFRECWAIT" Then
                    strTitle = "门诊处方待审查记录"
                ElseIf UCase(objTmp.Name) = "VSFRECFINISH" Then
                    strTitle = "门诊处方已审查记录"
                ElseIf UCase(objTmp.Name) = "VSFRECIPE" Then
                    strTitle = "门诊处方审查药嘱明细"
                ElseIf UCase(objTmp.Name) = "VSFAUDIT" Then
                    strTitle = "门诊处方审查项目明细"
                End If
                If strTitle <> "" Then
                    If Control.ID = enuMenus.打印预览 Then
                        zlRptPrint 0, objTmp, strTitle
                    ElseIf Control.ID = enuMenus.打印 Then
                        zlRptPrint 1, objTmp, strTitle
                    Else
                        zlRptPrint 3, objTmp, strTitle
                    End If
                End If
                objTmp.Redraw = True
            End If
        Case enuMenus.参数设置
            timPatient.Enabled = False      '关闭定时控件
            frmRAParams.ShowMe Me, 0
            Call SetPatientTimer            '过程确定是否开启定时
            Call SetClinicItem
        Case enuMenus.退出
            If mblnAuditStart Then
                If AuditOperate(0) Then
                    Unload Me
                End If
            Else
                Unload Me
            End If
        Case enuMenus.开启审查
            If AuditOperate(IIf(mblnAuditStart, 0, 1)) Then
                mblnAuditStart = True
            End If
            Call RefreshLockControls
'            Call FillVSFData(1)
'            Call FillVSFData(2)
'            Call FillVSFData(3)
        Case enuMenus.停止审查
            If AuditOperate(IIf(mblnAuditStart, 0, 1)) Then
                mblnAuditStart = False
            End If
            Call RefreshLockControls
        Case enuMenus.合格
            If mblnLocking = False Then
                lngAuditID = Val(vsfRecWait.TextMatrix(vsfRecWait.Row, vsfRecWait.ColIndex("审方ID")))
                '加锁
                Call AuditLock(lngAuditID)
            End If
            Call AuditProcess(1)
        Case enuMenus.不合格
            If mblnLocking = False Then
                lngAuditID = Val(vsfRecWait.TextMatrix(vsfRecWait.Row, vsfRecWait.ColIndex("审方ID")))
                '加锁
                Call AuditLock(lngAuditID)
            End If
            Call AuditProcess(2)
        Case enuMenus.查看PASS结果
            Call PassResultView(mobjPASS, True, Val(vsfRecipe.TextMatrix(vsfRecipe.Row, vsfRecipe.ColIndex("医嘱ID"))))
        Case enuMenus.刷新
            If mblnLocking Then
                If MsgBox("正在审查当前病人的药嘱，是否放弃审查？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
                
                Call AuditLock(0, False)
            End If
            
            Call AuditOperate(IIf(mblnAuditStart, 1, 0))
        
            Screen.MousePointer = vbHourglass
            
            Call FillVSFData(1)
            Call FillVSFData(2)
            'Call FillVSFData(3)
            Call RefreshLockControls
            Call SetStatusbar
            Call RefreshAMR
            
            Screen.MousePointer = vbDefault
        Case enuMenus.标准按钮
            For i = 2 To cbsMain.Count
                Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
            Next
            Me.cbsMain.RecalcLayout
        Case enuMenus.文本标签
            For i = 2 To cbsMain.Count
                For Each objControl In Me.cbsMain(i).Controls
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                Next
            Next
            Me.cbsMain.RecalcLayout
        Case enuMenus.大图标
            cbsMain.Options.LargeIcons = Not cbsMain.Options.LargeIcons
            cbsMain.RecalcLayout
        Case enuMenus.小字体
            If mbytFontSize <> 0 Then Call SetControlFontSize(0)
        Case enuMenus.大字体
            If mbytFontSize <> 1 Then Call SetControlFontSize(1)
        Case enuMenus.状态栏
            stbThis.Visible = Not Control.Checked
            cbsMain.RecalcLayout
        Case enuMenus.帮助主题
            Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
        Case enuMenus.中联主页
            Call zlHomePage(Me.hwnd)
        Case enuMenus.中联论坛
            Call zlWebForum(Me.hwnd)
        Case enuMenus.发送反馈
            Call zlMailTo(Me.hwnd)
        Case enuMenus.关于
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case Else
            '报表
            If Between(Control.ID, enuMenus.报表 * 100# + 1, enuMenus.报表 * 100# + 99) And Control.Parameter <> "" Then
                Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, "审方ID=" & mlngAuditID)
            End If
    End Select
End Sub

Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If Me.stbThis.Visible Then
        Bottom = stbThis.Height
    Else
        Bottom = 0
    End If
End Sub

Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case enuMenus.打印设置, enuMenus.打印预览, enuMenus.打印, enuMenus.输出Excel, enuMenus.参数设置
            Control.Enabled = Not mblnLocking
        Case enuMenus.开启审查
            Control.Enabled = mblnNeedAudit
            If mblnNeedAudit = False Then Exit Sub
            Control.Enabled = Not mblnAuditStart
        Case enuMenus.停止审查
            Control.Enabled = mblnNeedAudit
            If mblnNeedAudit = False Then Exit Sub
            Control.Enabled = mblnAuditStart
        Case enuMenus.合格
            Control.Enabled = cmdYes_Fixed.Enabled
        Case enuMenus.不合格
            Control.Enabled = cmdNo_Fixed.Enabled
        Case enuMenus.标准按钮
            Control.Checked = Me.cbsMain(2).Visible
        Case enuMenus.文本标签
            Control.Checked = (Me.cbsMain(2).Controls(1).Style = xtpButtonCaption Or Me.cbsMain(2).Controls(1).Style = xtpButtonIconAndCaption)
        Case enuMenus.大图标
            Control.Checked = cbsMain.Options.LargeIcons
        Case enuMenus.小字体
            Control.Checked = mbytFontSize = 0
            Control.Enabled = Not mblnLocking
        Case enuMenus.大字体
            Control.Checked = mbytFontSize = 1
            Control.Enabled = Not mblnLocking
        Case enuMenus.查看PASS结果
            Control.Enabled = (mintParaPass > 0 And mintParaPass < 5) And vsfRecipe.Rows > 1
        Case enuMenus.状态栏
            Control.Checked = Me.stbThis.Visible
        Case Else
            '报表
            If Between(Control.ID, enuMenus.报表 * 100# + 1, enuMenus.报表 * 100# + 99) And Control.Parameter <> "" Then
                Control.Enabled = Not mblnLocking
            End If
    End Select
End Sub

Private Sub chk24_Click()
    Call FillVSFData(2)
End Sub

Private Sub chkPubPatient_Click()
    Dim objControl As XtremeCommandBars.CommandBarButton
    
    Set objControl = cbsMain.ActiveMenuBar.Controls.Find(, enuMenus.刷新, , True)
    If Not objControl Is Nothing Then Call cbsMain_Execute(objControl)
End Sub

Private Sub cmdNo_Fixed_Click()
    Dim objControl As XtremeCommandBars.CommandBarButton
    
    Set objControl = cbsMain.ActiveMenuBar.FindControl(, enuMenus.不合格, , True)
    If Not objControl Is Nothing Then
        If objControl.Enabled Then
            Call cbsMain_Execute(objControl)
        End If
    End If
End Sub

Private Sub cmdYes_Fixed_Click()
    Dim objControl As XtremeCommandBars.CommandBarButton
    
    Set objControl = cbsMain.ActiveMenuBar.FindControl(, enuMenus.合格, , True)
    If Not objControl Is Nothing Then
        If objControl.Enabled Then
            Call cbsMain_Execute(objControl)
        End If
    End If
End Sub

Private Sub dkpMain_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
        Case 1
            Item.Handle = picRec.hwnd
        Case 2
            Item.Handle = picTab.hwnd
    End Select
End Sub

Private Sub Form_Activate()
    If mblnExit Then Unload Me
End Sub

Private Sub Form_Load()
    Dim strSQL As String, strTmp As String
    Dim rsTemp As ADODB.Recordset
    Dim lngTmp As Long
    
    On Error GoTo errHandle
    
    mblnEnter = False
    mblnExit = False
    mstrPCName = UCase(OS.ComputerName)
    
    '来源科室
'    '检查门诊处方审查条件是否设置来源科室
'    strSQL = "Select a.Id 部门id, a.编码, a.名称 " & vbNewLine & _
'             "From 部门表 A, 处方审查条件 B " & vbNewLine & _
'             "Where a.Id = b.科室id And (a.撤档时间 Is Null Or To_Char(a.撤档时间, 'yyyy') = '3000') And b.类别 = 1 "
'    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取门诊处方审查条件的来源科室")
'    If rsTemp.RecordCount <= 0 Then
'        MsgBox "“处方审查条件”未设置来源科室！", vbInformation, gstrSysName
'        rsTemp.Close
''        mblnExit = True
''        Exit Sub
'    End If
    
    '获取岗位设置的来源科室
    Call SetClinicItem
    
    '发药药房
    strSQL = "Select Distinct a.部门id, c.编码, c.名称 " & vbNewLine & _
             "From 部门人员 A, 部门性质说明 B, 部门表 C " & vbNewLine & _
             "Where a.部门id = b.部门id And a.部门id = c.Id And a.人员id = [1] And b.工作性质 In ('中药房', '西药房', '成药房') " & vbNewLine & _
             "   And b.服务对象 In (1, 3) And (c.撤档时间 Is Null Or To_Char(c.撤档时间, 'yyyy') = '3000') " & vbNewLine & _
             "Order By c.名称 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取操作员的所属门诊类药房", UserInfo.ID)
    If rsTemp.RecordCount <= 0 Then
        MsgBox "操作员无门诊类药房部门的属性，请检查部门管理！", vbInformation, gstrSysName
        rsTemp.Close
        mblnExit = True
        Exit Sub
    End If
    
    With rsTemp
        cboDrugstore.Tag = ""
        cboDrugstore.Clear
        cboDrugstore.AddItem "所有发药药房"
        Do While .EOF = False
            cboDrugstore.AddItem !名称
            cboDrugstore.ItemData(cboDrugstore.NewIndex) = !部门ID
            cboDrugstore.Tag = IIf(cboDrugstore.Tag = "", "", cboDrugstore.Tag & ",") & !部门ID
            .MoveNext
        Loop
        .Close
        cboDrugstore.ListIndex = 0
    End With

    '初始化参数和模块变量
    mblnSendBeforeAudit = Val(zlDatabase.GetPara("门诊审方时机", glngSys)) = 1
    mintParaPass = Val(zlDatabase.GetPara("合理用药监测接口", glngSys))      '0-表示未使用,1-美康接口,2-大通接口（暂不支持）,3-太元通接口,4-保进
    mbytDrugName = Val(zlDatabase.GetPara("药品名称显示"))
    mintItem = Val(zlDatabase.GetPara("处方审查依据", glngSys, , "1"))
    mblnRecipeSendAuto = Val(zlDatabase.GetPara("门诊审方处方自动发送", glngSys)) = 1
    
    lngTmp = Val(zlDatabase.GetPara("处方审查", glngSys))
    mblnNeedAudit = (lngTmp = 1 Or lngTmp = 3)                              '1-门诊启用审查；3-门诊和住院启用审查
    mblnAuditStart = False                                                  '进入窗体审查状态为禁止
    mblnLocking = False
    mlngModule = glngModule
    mstrPrivs = zlStr.FormatString(";[1];", GetPrivFunc(glngSys, mlngModule))
    mblnMemory = Val(zlDatabase.GetPara("使用个性化风格")) = 1
    mbytFontSize = 0
    
    '创建消息平台对象
    On Error Resume Next
    Set mobjMipModule = New zl9ComLib.clsMipModule
    If Not mobjMipModule Is Nothing Then
        mobjMipModule.InitMessage glngSys, mlngModule, mstrPrivs
        zl9ComLib.AddMipModule mobjMipModule
        If Err.Number <> 0 Then Set mobjMipModule = Nothing
    End If
    Err.Clear: On Error GoTo errHandle
    
'    '加载电子病案窗体
'    On Error Resume Next
'    Set mobjArchiveMedRec = New zlPublicAdvice.clsArchiveMedRec
'    If Not mobjArchiveMedRec Is Nothing Then
'        Call mobjArchiveMedRec.InitArchiveMedRec(gcnOracle, glngSys)
'        If Err.Number <> 0 Then
'            Set mobjArchiveMedRec = Nothing
'        Else
'            Set mfrmAMR = mobjArchiveMedRec.zlGetForm(0)
'        End If
'    End If
'    Err.Clear: On Error GoTo errHandle
    
    '临床的公共方法
    On Error Resume Next
    Set mobjPubAdvice = New zlPublicAdvice.clsPublicAdvice
    If Not mobjPubAdvice Is Nothing Then
        Call mobjPubAdvice.InitCommon(gcnOracle, glngSys)
        If Err.Number <> 0 Then
            Set mobjPubAdvice = Nothing
        Else
            Set mfrmAMR = mobjPubAdvice.GetArchiveFrom()
        End If
    End If
    Err.Clear: On Error GoTo errHandle
    
    '合理用药
    If mintParaPass > 0 Then
        On Error Resume Next
        Set mobjPASS = CreateObject("zlPassInterface.clsPass")
        If Not mobjPASS Is Nothing Then
            If mobjPASS.zlPassInit_YF(gcnOracle, glngSys, mlngModule) = False Then
                Set mobjPASS = Nothing
            End If
        End If
        Err.Clear: On Error GoTo errHandle
    End If
    
    '初始化界面布局
    Call InitCommandbars
    Call InitDockPane
    Call InitTBCRec
    Call InitTBCTab
    
    '初始化控件
    Err.Clear: On Error Resume Next
    Call iknPatient.zlInit(Me, glngSys, mlngModule, gcnOracle, UserInfo.用户名, , MSTR_PATIENT, txtPatient)
    Err.Clear: On Error GoTo errHandle
    
    Call InitVSF(vsfRecWait)
    Call InitVSF(vsfRecFinish)
    Call InitVSF(vsfRecipe)
    Call InitVSF(vsfAudit)
    
    Call mdlDefine.SetVSFHead(vsfRecWait, MSTR_VSF_WAIT)
    Call mdlDefine.SetVSFHead(vsfRecFinish, MSTR_VSF_FINISH)
    Call mdlDefine.SetVSFHead(vsfRecipe, MSTR_VSF_RECIPE)
    Call mdlDefine.SetVSFHead(vsfAudit, MSTR_VSF_AUDIT)
    
    If mbytDrugName = 0 Then
        vsfRecipe.ColHidden(vsfRecipe.ColIndex("药品名称")) = False
        vsfRecipe.ColHidden(vsfRecipe.ColIndex("商品名")) = True
    ElseIf mbytDrugName = 1 Then
        vsfRecipe.ColHidden(vsfRecipe.ColIndex("药品名称")) = True
        vsfRecipe.ColHidden(vsfRecipe.ColIndex("商品名")) = False
    Else
        vsfRecipe.ColHidden(vsfRecipe.ColIndex("药品名称")) = False
        vsfRecipe.ColHidden(vsfRecipe.ColIndex("商品名")) = False
    End If
    
    '处理PASS结果的图片（灯）；  0-表示未使用,1-美康接口,2-大通接口,3-太元通接口,4-保进
    vsfRecipe.ColHidden(vsfRecipe.ColIndex("PASS")) = (mintParaPass < 1 Or mintParaPass > 4)
    
    With vsfAudit
        .Editable = flexEDKbdMouse
        .ColComboList(vsfAudit.ColIndex("药品")) = "..."
    End With
    
    Call SetFilterDay(0)
    
    '恢复上次界面
    RestoreWinState Me, App.ProductName
    If mblnMemory Then
        Dim strPane As String
        Dim objControl As XtremeCommandBars.CommandBarButton
        
        '字体大小
        lngTmp = Val(GetSetting("ZLSOFT", zlStr.FormatString("私有模块\[1]\界面设置\[2]\[3]\Form", "ZLHIS", App.ProductName, Me.Name), "FontSize"))
        Set objControl = cbsMain.FindControl(, IIf(lngTmp = 1, enuMenus.大字体, enuMenus.小字体), , True)
        If Not objControl Is Nothing Then
            Call cbsMain_Execute(objControl)
        End If
        
        'DockingPane
        strPane = GetSetting("ZLSOFT", zlStr.FormatString("私有模块\[1]\界面设置\[2]\[3]\Form", "ZLHIS", App.ProductName, Me.Name), "布局")
        iknPatient.IDKind = Val(GetSetting("ZLSOFT", zlStr.FormatString("私有模块\[1]\界面设置\[2]\[3]\Form", "ZLHIS", App.ProductName, Me.Name), "iknPatient"))
        dkpMain.LoadStateFromString strPane
    End If
    
    '解锁自己锁定的审方记录
    Call AuditLock(0, False)
    
    '刷新数据
    Call FillVSFData(1)
    Call FillVSFData(2)
    'Call FillVSFData(3)
    
    sbnPASS.Visible = (mintParaPass > 0 And mintParaPass < 5) And Not mobjPASS Is Nothing    '查看PASS结果
    
    If vsfRecWait.Visible And vsfRecWait.Enabled Then vsfRecWait.SetFocus
    
    Call SetPatientTimer
    
    mblnEnter = True

    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then
        Resume
    Else
        mblnExit = True
    End If
End Sub

Private Sub InitCommandbars()
    Dim cbpTmp As CommandBarPopup
    Dim cbcTmp As CommandBarControl
    Dim cbrTmp As CommandBar
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    cbsMain.VisualTheme = xtpThemeOffice2003 'xtpthemeoffice2000有凹凸感
    With cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    With cbsMain
        .EnableCustomization False
        Set .Icons = zlCommFun.GetPubIcons
        .ActiveMenuBar.Title = "菜单"
        .ActiveMenuBar.EnableDocking xtpFlagHideWrap Or xtpFlagStretched
    End With
    
    picLine01_S.BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
    picLine02.BackColor = picLine01_S.BackColor
    
    '文件
    Set cbpTmp = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enuMenus.文件, "文件(&F)", -1, False)
    With cbpTmp
        .ID = enuMenus.文件
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.打印设置, "打印设置(&S)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.打印预览, "打印预览(&V)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.打印, "打印")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.输出Excel, "输出到&Excel...")
        cbcTmp.BeginGroup = True
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.参数设置, "参数设置")
        cbcTmp.BeginGroup = True
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.退出, "退出")
    End With
    
    '编辑
    Set cbpTmp = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enuMenus.编辑, "编辑(&E)", -1, False)
    With cbpTmp
        .ID = enuMenus.编辑
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.开启审查, "开启审查(&S)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.停止审查, "停止审查(&P)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.合格, "合格(&Y)")
        cbcTmp.BeginGroup = True
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.不合格, "不合格(&N)")
    End With
    
    '查看
    Set cbpTmp = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enuMenus.查看, "查看(&V)", -1, False)
    With cbpTmp
        .ID = enuMenus.查看
        Set cbpTmp = .CommandBar.Controls.Add(xtpControlPopup, enuMenus.工具栏, "工具栏(&T)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.标准按钮, "标准按钮(&S)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.文本标签, "文本标签(&T)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.大图标, "大图标(&B)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.状态栏, "状态栏(&S)")
        cbcTmp.BeginGroup = True
        Set cbpTmp = .CommandBar.Controls.Add(xtpControlPopup, enuMenus.字体大小, "字体大小(&F)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.小字体, "小字体(&S)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.大字体, "大字体(&B)")
        cbcTmp.BeginGroup = True
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.查看PASS结果, "查看&PASS结果")
        cbcTmp.BeginGroup = True
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.刷新, "刷新")
        cbcTmp.BeginGroup = True
    End With
    
    '帮助
    Set cbpTmp = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enuMenus.帮助, "帮助(&H)", -1, False)
    With cbpTmp
        .ID = enuMenus.帮助
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.帮助主题, "帮助主题")
        Set cbpTmp = .CommandBar.Controls.Add(xtpControlPopup, enuMenus.WEB上的中联, "&WEB上的中联")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.中联主页, "中联主页(&H)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.中联论坛, "中联论坛(&F)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.发送反馈, "发送反馈(&K)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.关于, "关于(&A)")
        cbcTmp.BeginGroup = True
    End With
    
    '报表接口
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModule, mstrPrivs)
    
    '菜单项的快键绑定
    With cbsMain.KeyBindings
        .Add 8, vbKeyP, enuMenus.打印
        .Add 8, vbKeyX, enuMenus.退出
        .Add 0, vbKeyF12, enuMenus.参数设置
        .Add 0, vbKeyF5, enuMenus.刷新
        .Add 0, vbKeyF1, enuMenus.帮助主题
    End With
    
    '定义工具栏
    Set cbrTmp = cbsMain.Add("工具栏", xtpBarTop)
    With cbrTmp
        .ShowTextBelowIcons = False
        .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
        
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.打印预览, "打印预览")
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.打印, "打印")
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.开启审查, "开启审查")
        cbcTmp.BeginGroup = True
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.停止审查, "停止审查")
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.刷新, "刷新")
        cbcTmp.BeginGroup = True
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.帮助主题, "帮助主题")
        cbcTmp.BeginGroup = True
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.退出, "退出")
    End With
    
    '有图标，有文本的按钮风格
    For Each cbcTmp In cbsMain(2).Controls
        If cbcTmp.Type <> xtpControlLabel Then
            cbcTmp.Style = xtpButtonIconAndCaption
        End If
    Next
    
End Sub

Private Sub InitDockPane()
    Dim panLeft As Pane, panRight As Pane
    
    With dkpMain
        .SetCommandBars cbsMain
        .Options.UseSplitterTracker = False
        .Options.ThemedFloatingFrames = True
        .Options.AlphaDockingContext = True
        .Options.LunaColors = True
        .Options.HideClient = True
        .VisualTheme = ThemeGrippered
        
        Set panLeft = .CreatePane(1, 250, 0, DockLeftOf)
        With panLeft
            .Options = PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
            .Title = "过滤条件"
            .MaxTrackSize.Width = 1000
            .MinTrackSize.Width = 250
        End With
        
        Set panRight = .CreatePane(2, 0, 0, DockRightOf)
        With panRight
            .Options = PaneNoCaption Or PaneNoCloseable Or PaneNoFloatable Or PaneNoHideable
            .MaxTrackSize.Width = 2000
            .MinTrackSize.Width = 450
        End With
    End With
End Sub

Private Sub InitTBCRec()
    
    With tbcRec.PaintManager
        .Appearance = xtpTabAppearancePropertyPage2003
        .ClientFrame = xtpTabFrameSingleLine
        .BoldSelected = True
        .OneNoteColors = True
        .ShowIcons = False
    End With
    
    With tbcRec
        .InsertItem 0, "待审(&A)", picRecWait.hwnd, 0
        .InsertItem 1, "已审(&B)", picRecFinish.hwnd, 0
    End With
    
End Sub

Private Sub InitTBCTab()
    
    With tbcTab.PaintManager
        .Appearance = xtpTabAppearancePropertyPage2003
        .ClientFrame = xtpTabFrameSingleLine
        .BoldSelected = True
        .OneNoteColors = True
        .ShowIcons = False
    End With
    
    With tbcTab
        .InsertItem 0, "处方审查(&1)", picDetail.hwnd, 0
        If Not mfrmAMR Is Nothing Then
            .InsertItem 1, "医嘱和报告(&2)", mfrmAMR.hwnd, 0
        End If
    End With
    
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    If mblnAuditStart Then
        If AuditOperate(0) Then
            Cancel = False
        Else
            Cancel = True
        End If
    Else
        Cancel = False
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Width < 9000 Then Width = 9000
    If Height < 7000 Then Height = 7000
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim strPane As String
    
    If Not mobjPASS Is Nothing Then
        Set mobjPASS = Nothing
    End If
    If Not mobjMipModule Is Nothing Then
        mobjMipModule.CloseMessage
        zl9ComLib.DelMipModule mobjMipModule
        Set mobjMipModule = Nothing
    End If
    If Not mfrmAMR Is Nothing Then
        Unload mfrmAMR
    End If
    If Not mobjPubAdvice Is Nothing Then
        Set mobjPubAdvice = Nothing
    End If
    
    If mblnExit = False Then
        SaveWinState Me, App.ProductName
        strPane = dkpMain.SaveStateToString
        SaveSetting "ZLSOFT", zlStr.FormatString("私有模块\[1]\界面设置\[2]\[3]\Form", "ZLHIS", App.ProductName, Me.Name), "布局", strPane
        SaveSetting "ZLSOFT", zlStr.FormatString("私有模块\[1]\界面设置\[2]\[3]\Form", "ZLHIS", App.ProductName, Me.Name), "iknPatient", iknPatient.IDKind
        SaveSetting "ZLSOFT", zlStr.FormatString("私有模块\[1]\界面设置\[2]\[3]\Form", "ZLHIS", App.ProductName, Me.Name), "FontSize", mbytFontSize
    End If
End Sub

Private Sub iknPatient_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If txtPatient.Enabled And txtPatient.Visible Then
        txtPatient.Text = ""
        txtPatient.SetFocus
    End If
End Sub

Private Sub iknPatient_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    txtPatient.Text = objPatiInfor.卡号
    mblnReadCard = True
    Call txtPatient_KeyPress(0)
End Sub

Private Sub picAudit_Resize()
    On Error Resume Next
    
    With lblReason
        .Left = 120
        .Top = IIf(mbytFontSize = 1, 45, 90)
    End With
    
    With txtReason
        .Left = 0
        .Top = 350
        .Width = picAudit.ScaleWidth - 3090
        .Height = 725
    End With
    
    With sbnHouse
        .Left = txtReason.Width - .Width - 60
        .Top = 30
    End With
    
    With sbnSelect
        .Left = sbnHouse.Left - 60 - .Width
        .Top = 30
    End With
    
    With picFunc
        .Left = txtReason.Width
        .Top = 0
        .Width = picAudit.ScaleWidth - txtReason.Width
        .Height = txtReason.Height + txtReason.Top
    End With
    
    With picYesNO
        .Left = picFunc.Left
        .Top = picFunc.Top
        .Width = picFunc.Width
        .Height = picFunc.Height
    End With
    
    With picLine02
        .Left = 0
        .Top = txtReason.Top + txtReason.Height
        .Width = picAudit.ScaleWidth
    End With
    
    With lblAudit
        .Top = picLine02.Top + picLine02.Height + IIf(mbytFontSize = 1, 45, 90)
        .Left = 120
    End With
    
    With vsfAudit
        .Top = picLine02.Top + picLine02.Height + 350
        .Left = 0
        .Width = picAudit.ScaleWidth
        .Height = picAudit.ScaleHeight - .Top
    End With
    
End Sub

Private Sub picDetail_Resize()
    On Error Resume Next
    
    With txtOther
        .Left = 45
        .Top = 45
        .Width = picDetail.ScaleWidth - 45 * 2
        .Height = IIf(mbytFontSize = 0, 500, 700)
    End With
    
    With picLine01_S
        .Left = 0
        .Width = picDetail.ScaleWidth
    End With
    
    With picRecipe
        .Left = 0
        .Top = txtOther.Height + 45
        .Width = picDetail.ScaleWidth
        .Height = picLine01_S.Top - txtOther.Top - txtOther.Height
    End With
    
    With picAudit
        .Left = 0
        .Top = picLine01_S.Top + picLine01_S.Height
        .Width = picDetail.ScaleWidth
        .Height = picDetail.ScaleHeight - .Top
    End With
End Sub

Private Sub picFunc_Resize()
    On Error Resume Next
    
    cmdYes_Fixed.Move 15, 0
    cmdNo_Fixed.Move cmdYes_Fixed.Left + cmdYes_Fixed.Width + 15, 0
    
End Sub

Private Sub picLine01_S_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    msngY = Y
End Sub

Private Sub picLine01_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> vbLeftButton Then Exit Sub
    With picLine01_S
        If .Top + Y < ScaleHeight * 0.3 Then
            .Top = ScaleHeight * 0.3
            Exit Sub
        End If
        If .Top + Y > ScaleHeight * 0.5 Then
            .Top = ScaleHeight * 0.5
            Exit Sub
        End If
        .Move .Left, .Top + Y - msngY
    End With
End Sub

Private Sub picLine01_S_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call picDetail_Resize
    msngY = 0
End Sub

Private Sub picRec_Resize()
    On Error Resume Next
    
    With cboClinic
        .Left = lblClinic.Left + lblClinic.Width + 60
        .Width = picRec.ScaleWidth - cboClinic.Left - 120
    End With
    
    With cboDrugstore
        .Left = cboClinic.Left
        .Width = cboClinic.Width
    End With
    
    With cboDate
        .Left = cboClinic.Left
        .Width = cboClinic.Width
    End With
    
    With txtPatient
        .Left = cboClinic.Left
        .Width = cboClinic.Width
    End With
    
    With tbcRec
        .Left = 15
        .Top = txtPatient.Top + txtPatient.Height + 120
        .Width = picRec.ScaleWidth - 15 * 2
        .Height = picRec.ScaleHeight - .Top - 15
    End With
    
End Sub

Private Sub picRecFinish_Resize()
    On Error Resume Next
    
    With Me.vsfRecFinish
        .Top = 0
        .Left = 0
        .Width = picRecFinish.ScaleWidth
        .Height = picRecFinish.ScaleHeight
    End With
End Sub

Private Sub picRecipe_Resize()
    On Error Resume Next
    
    With lblRecipe
        .Top = IIf(mbytFontSize = 1, 45, 90)
        .Left = 120
    End With
    
    If sbnPASS.Visible Then
        With sbnPASS
            .Top = 30
            .Left = picRecipe.ScaleWidth - .Width - 30
        End With
        
        With chk24
            .Top = 60
            .Left = sbnPASS.Left - .Width - 60
        End With
    Else
        With chk24
            .Top = 60
            .Left = picRecipe.ScaleWidth - .Width - 30
        End With
    End If
    
    With vsfRecipe
        .Left = 0
        .Top = 350
        .Width = picRecipe.ScaleWidth
        .Height = picRecipe.ScaleHeight - sbnPASS.Height - 30 * 2
    End With
End Sub

Private Sub picRecWait_Resize()
    On Error Resume Next
    
    With Me.vsfRecWait
        .Top = 0
        .Left = 0
        .Width = picRecWait.ScaleWidth
        .Height = picRecWait.ScaleHeight - chkPubPatient.Height - 45 * 2
    End With
    
    With chkPubPatient
        .Top = vsfRecWait.Height + 45
        .Left = 120
    End With
End Sub

Private Sub picTab_Resize()
    On Error Resume Next
    
    With tbcTab
        .Left = 0
        .Top = 0
        .Width = picTab.ScaleWidth
        .Height = picTab.ScaleHeight - .Top
    End With
End Sub

Private Sub picYesNO_Resize()
    On Error Resume Next
    
    With lblDisp_Fixed
        .Left = (picYesNO.Width - .Width) \ 2
        .Top = (picYesNO.Height - .Height) \ 2
    End With
End Sub

Private Sub sbnHouse_Click()
    Dim strSQL As String
    
    If mblnNeedAudit = False Or mblnAuditStart = False Then Exit Sub
    If txtReason.Text = "" Or vsfAudit.Rows <= 1 Then Exit Sub
    
    If zlCommFun.ActualLen(txtReason.Text) > txtReason.MaxLength Then
        MsgBox "“综合理由”内容超出500字符或250汉字！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    On Error GoTo errHandle
    
    strSQL = zlStr.FormatString("Zl_处方审查常用理由_Update(1, '[1]', '[2]')", _
                    UserInfo.用户名, _
                    txtReason.Text)
    Call zlDatabase.ExecuteProcedure(strSQL, "新增处方审查常用理由")
    
    MsgBox "收藏完成！", vbInformation, gstrSysName
    
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub sbnPASS_Click()
    Dim objControl As XtremeCommandBars.CommandBarButton
    
    Set objControl = cbsMain.ActiveMenuBar.Controls.Find(, enuMenus.查看PASS结果, , True)
    If Not objControl Is Nothing Then Call cbsMain_Execute(objControl)
End Sub

Private Sub sbnSelect_Click()
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim blnCancel As Boolean
    Dim i As Integer
    
    If mblnNeedAudit = False Or mblnAuditStart = False Then Exit Sub
    
    On Error GoTo errHandle
    
    strSQL = "Select Null 选择, 内容 From 处方审查常用理由 Where 用户名 = [1] "
    
    '理由选择器，可多选
    Set rsTemp = mdlRecipeAudit.ShowReason(Me, strSQL, blnCancel, UserInfo.用户名)
    
    If blnCancel = False Then
        With rsTemp
            strSQL = ""
            i = 1
            If .RecordCount > 0 Then .MoveFirst
            Do While .EOF = False
                strSQL = strSQL & zlStr.FormatString("[1]、[2][3]", i, !内容, vbNewLine)
                i = i + 1
                .MoveNext
            Loop
            If strSQL <> "" Then
                txtReason.Text = strSQL
            End If
        End With
        rsTemp.Close
    End If
    
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub tbcRec_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    Dim objControl As XtremeCommandBars.CommandBarButton
    
    If Item.Index = 0 Then
        '待审
        picFunc.Visible = True
        picYesNO.Visible = False
        chk24.Visible = True
    Else
        '已审
        picFunc.Visible = False
        picYesNO.Visible = True
        chk24.Visible = False
    End If
    
    SetFilterDay Item.Index
    
    If Me.Visible = False Then Exit Sub
    
    Set objControl = cbsMain.ActiveMenuBar.Controls.Find(, enuMenus.刷新, , True)
    If Not objControl Is Nothing Then
        Call cbsMain_Execute(objControl)
'        If vsfRecFinish.ColKey(0) <> "" Then Call vsfRecFinish_AfterRowColChange(0, 0, vsfRecFinish.Row, 1)
    End If

End Sub

Private Sub InitVSF(ByRef vsfVar As VSFlexGrid)
'功能：初始化窗体的VSFlexGrid控件的风格
'参数：
'  vsfVar：要初始化的VSFlexGrid控件

    With vsfVar
        .Appearance = flexFlat
        .AllowSelection = False
        .AllowUserResizing = flexResizeColumns
        .FixedCols = 0
        .FixedRows = 1
        .SelectionMode = flexSelectionByRow
        .SheetBorder = .BackColor
        .WordWrap = True
        .AutoSizeMode = flexAutoSizeRowHeight
        .BackColorBkg = .BackColor
        .AutoResize = True
    End With
End Sub

Private Sub FillVSFData(ByVal bytMode As Byte)
'功能：填充VSF控件的数据
'参数：
'  bytMode：1-病人审查数据；2-处方明细数据；3-审查项目数据

    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String, strTmp As String
    Dim datBegin As Date, datEnd As Date
    Dim lngClinicID As Long, lngDrugstoreID As Long, lngRAID As Long, lngPatientID As Long
    Dim strPatient As String
    Dim vsfTmp As VSFlexGrid
    Dim l As Long
    Dim dblPrice As Double, dblAmount As Double
    Dim intDay As Integer
    Dim blnReadID As Boolean, blnAllPass As Boolean
    Dim objPatiInfo As PatiInfor
    
    On Error GoTo errHandle
    
    intDay = cboDate.ListIndex
    If intDay < 0 Then intDay = 0
    
    Select Case bytMode
        Case 1      '1-病人审查数据
            '来源科室
            If cboClinic.ListCount > 0 Then
                lngClinicID = cboClinic.ItemData(cboClinic.ListIndex)
            Else
                '未设置来源科室
                lngClinicID = -1
            End If
            
            '发药药房
            lngDrugstoreID = cboDrugstore.ItemData(cboDrugstore.ListIndex)
            
            '提交时间
            datEnd = Format(zlDatabase.Currentdate, "yyyy-MM-dd 23:59:59")
            datBegin = Format(datEnd - intDay, "yyyy-MM-dd 00:00:00")
            
            '病人信息
            If Trim(txtPatient.Text) <> "" Then
                Select Case iknPatient.GetCurCard.名称
                    Case "姓名"
                        strPatient = " And Upper(b.姓名) = [5] "
                    Case "门诊号"
                        strPatient = " And b.门诊号 = [5] "
                    Case "身份证号"
                        strPatient = " And b.身份证号 = [5] "
                    Case Else
                        Call iknPatient.zlFindPatient(Me.txtPatient.Text, objPatiInfo)
                        If Not objPatiInfo Is Nothing Then
                            lngPatientID = objPatiInfo.病人ID
                        End If
                        strPatient = " And b.病人ID = [5] "
                        blnReadID = True
                End Select
            End If
            
            If lngClinicID < 0 Then
                '未设置来源科室
                strTmp = "Select -1 来源科室 From Dual"
            ElseIf lngClinicID = 0 Then
                '所有被勾选的来源科室
                strTmp = "Select Column_Value 来源科室 " & vbNewLine & _
                         "From Table(f_Num2list((Select 来源科室 From 处方审查参数 Where Upper(机器名) = [6] And 服务对象 = 0), ',')) "
'                strTmp = "Select a.Id 来源科室  " & vbNewLine & _
'                         "From 部门表 A, 处方审查条件 B, " & vbNewLine & _
'                         "    Table(f_Num2list((Select 来源科室 From 处方审查参数 Where 机器名 = [6] And 服务对象 = 0))) C " & vbNewLine & _
'                         "Where a.Id = b.科室id And a.Id = c.Column_Value And (a.撤档时间 Is Null Or To_Char(a.撤档时间, 'yyyy') = '3000') " & vbNewLine & _
'                         "    And b.类别 = 1 And (b.科室id Is Not Null Or b.科室id > 0)"
            End If
            
            If Me.tbcRec.Item(0).Selected Then
                '待审查
                If chkPubPatient.Value = 1 And lngClinicID <= 0 Then
                    '岗位未设置的来源科室
                    strTmp = strTmp & _
                             "Union All Select 科室ID From (" & vbNewLine & _
                             "Select 科室id From 处方审查条件 Where 类别 = 1 " & vbNewLine & _
                             "Minus " & vbNewLine & _
                             "Select Distinct Column_Value 来源科室 " & vbNewLine & _
                             "From Table(f_Num2list((Select f_List2str(Cast(Collect(来源科室) As t_Strlist), ',') 来源科室 " & vbNewLine & _
                             "                       From 处方审查参数 " & vbNewLine & _
                             "                       Where 服务对象 = 0 And 最后操作时间 >= Sysdate - 60 And 来源科室 Is Not Null), ',')) ) "
                End If
                
                strSQL = "Select b.姓名 病人, b.病人id, b.主页id, b.性别, b.年龄, b.门诊号, c.Id 审方id, c.挂号id 挂号单Id, " & vbNewLine & _
                         "    To_Char(c.提交时间, 'yyyy-mm-dd hh24:mi') 提交时间, c.提交人, D1.名称 提交科室, D2.名称 发药药房 " & vbNewLine & _
                         "From " & IIf(lngClinicID <= 0, zlStr.FormatString("(Select 来源科室 From ([1])) A,", strTmp), "") & vbNewLine & _
                         "    病人信息 B, 处方审查记录 C, 部门表 D1, 部门表 D2 " & vbNewLine & _
                         IIf(lngDrugstoreID <= 0, ", Table(f_Num2list([4], ',')) E ", "") & vbNewLine & _
                         "Where b.病人id = c.病人id And c.提交科室id = D1.Id And c.发药药房id = D2.Id " & vbNewLine & _
                         IIf(lngClinicID <= 0, " And a.来源科室 = c.提交科室id ", "") & vbNewLine & _
                         IIf(lngDrugstoreID <= 0, " And c.发药药房id = e.Column_Value ", " And c.发药药房id = [4] ") & vbNewLine & _
                         "    And c.审查时间 Is Null And c.状态 = 0 And c.提交时间 Between [1] And [2] And c.挂号ID Is Not Null "
                strSQL = strSQL & IIf(lngClinicID <= 0, "", " And c.提交科室id = [3] ")
                strSQL = strSQL & IIf(Trim(txtPatient.Text) = "", "", strPatient)
                strSQL = strSQL & vbNewLine & "Order By c.提交时间 "
                
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询门诊类待审查数据", _
                                        datBegin, datEnd, _
                                        lngClinicID, _
                                        IIf(lngDrugstoreID <= 0, cboDrugstore.Tag, CStr(lngDrugstoreID)), _
                                        IIf(blnReadID = False, txtPatient.Text, lngPatientID), _
                                        mstrPCName)
                Call mdlDefine.FillVSFData(vsfRecWait, rsTemp)
                
                If vsfRecWait.Rows > 1 Then
                    vsfRecWait.Row = 1
                End If
            
            Else
                '已审查
                strSQL = "Select b.姓名 病人, b.病人id, b.主页id, b.性别, b.年龄, b.门诊号, c.Id 审方id, c.挂号id 挂号单Id, " & vbNewLine & _
                         "    To_Char(c.提交时间, 'yyyy-mm-dd hh24:mi') 提交时间, c.提交人, Decode(c.审查结果, 1, '合格', '不合格') 审查结果, " & vbNewLine & _
                         "    To_Char(c.审查时间, 'yyyy-mm-dd hh24:mi') 审查时间, c.审查人, D1.名称 提交科室, D2.名称 发药药房 " & vbNewLine & _
                         "From " & IIf(lngClinicID <= 0, zlStr.FormatString("(Select 来源科室 From ([1])) A,", strTmp), "") & vbNewLine & _
                         "    病人信息 B, 处方审查记录 C, 部门表 D1, 部门表 D2 " & vbNewLine & _
                         IIf(lngDrugstoreID <= 0, ", Table(f_Num2list([4], ',')) E ", "") & vbNewLine & _
                         "Where b.病人id = c.病人id And c.提交科室id = D1.Id And c.发药药房id = D2.Id " & vbNewLine & _
                         IIf(lngClinicID <= 0, " And a.来源科室 = c.提交科室id ", "") & vbNewLine & _
                         IIf(lngDrugstoreID <= 0, " And c.发药药房id = e.Column_Value ", " And c.发药药房id = [4] ") & vbNewLine & _
                         "    And c.审查时间 Is Not Null And c.状态 = 1 And c.提交时间 Between [1] And [2] And c.挂号ID Is Not Null "
                strSQL = strSQL & IIf(lngClinicID <= 0, "", " And c.提交科室id = [3] ")
                strSQL = strSQL & IIf(Trim(txtPatient) = "", "", strPatient)
                strSQL = strSQL & vbNewLine & "Order By c.提交时间 "
                
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询门诊类已审查数据", _
                                        datBegin, datEnd, _
                                        lngClinicID, _
                                        IIf(lngDrugstoreID <= 0, cboDrugstore.Tag, CStr(lngDrugstoreID)), _
                                        IIf(blnReadID = False, txtPatient.Text, lngPatientID), _
                                        mstrPCName)
                Call mdlDefine.FillVSFData(vsfRecFinish, rsTemp)
                
                If vsfRecFinish.Rows > 1 Then
                    vsfRecFinish.Row = 1
                End If
                
            End If
            rsTemp.Close
            
        Case 2      '2-处方明细数据
            
            If Me.tbcRec.Item(0).Selected Then
                Set vsfTmp = vsfRecWait
            Else
                Set vsfTmp = vsfRecFinish
            End If
            
            vsfRecipe.Rows = 1
            vsfRecipe.Clear 1
            
            If vsfTmp.Rows <= 1 Then Exit Sub
            
            lngRAID = Val(vsfTmp.TextMatrix(vsfTmp.Row, vsfTmp.ColIndex("审方ID")))
            lngPatientID = Val(vsfTmp.TextMatrix(vsfTmp.Row, vsfTmp.ColIndex("病人ID")))
            
            strSQL = "Select Null 标志, e.名称 临床科室, b.审方id, b.医嘱id, A1.超量说明, A1.医生嘱托, A1.相关id, A1.审查结果 PASS, " & vbNewLine & _
                     "    d.名称 药品名称, f.名称 商品名, d.规格, d.计算单位 单位, A1.总给予量 数量, A1.单次用量 || Nvl(g.计算单位, '') 单量, A2.医嘱内容 用法, " & vbNewLine & _
                     "    A1.执行频次 频次, c.标准单价 单价, c.应收金额, d.ID 细目ID, " & vbNewLine & _
                     "    Row_Number() Over(Partition By A1.相关id Order By b.医嘱id) 组序号, Count(1) Over(Partition By A1.相关id) 组数 " & vbNewLine & _
                     "From 病人医嘱记录 A1, 病人医嘱记录 A2, 处方审查明细 B, (" & vbNewLine & _
                     "       Select a.医嘱序号, Max(a.标准单价) 标准单价, Sum(a.应收金额) 应收金额 " & vbNewLine & _
                     "       From 门诊费用记录 A, 处方审查明细 B " & vbNewLine & _
                     "       Where a.医嘱序号 = b.医嘱id And b.审方id = [1] " & vbNewLine & _
                     "       Group By a.医嘱序号 " & vbNewLine & _
                     "    ) C, 收费项目目录 D, 部门表 E, 收费项目别名 F, 诊疗项目目录 G " & vbNewLine & _
                     "Where A1.相关id = A2.Id And A1.Id = b.医嘱id And A1.Id = c.医嘱序号(+) And A1.收费细目id = d.Id And A1.开嘱科室id = e.Id And " & vbNewLine & _
                     "    A1.收费细目id = f.收费细目id(+) And b.审方id = [1] And f.性质(+) = 3 And f.码类(+) = 1 And A1.诊疗项目ID = g.ID(+) " & vbNewLine & _
                     "Order By A1.相关id, 组序号 "
            
            If chk24.Value = 1 And Me.chk24.Visible = True Then
                '24小时内的所有药嘱（有病人医嘱发送记录）
'                strTmp = "Select e.名称 临床科室, Null 审方id, A1.Id 医嘱id, A1.超量说明, A1.医生嘱托, A1.相关id, A1.审查结果, d.名称 药品名称, " & vbNewLine & _
'                         "    f.名称 商品名, d.规格, d.计算单位 单位, A1.总给予量 数量, A1.单次用量 单量, A2.医嘱内容 用法, A1.执行频次 频次, " & vbNewLine & _
'                         "    c.标准单价 单价, c.应收金额, d.ID 细目ID, Row_Number() Over(Partition By A1.相关id Order By A1.Id) 组序号,  " & vbNewLine & _
'                         "    Count(1) Over(Partition By A1.相关id) 组数 " & vbNewLine & _
'                         "From 病人医嘱记录 A1, 病人医嘱记录 A2, 病人医嘱发送 B, 门诊费用记录 C, 收费项目目录 D, 部门表 E, 收费项目别名 F " & vbNewLine & _
'                         "Where A1.相关id = A2.Id And A1.Id = b.医嘱id And A1.Id = c.医嘱序号(+) And A1.收费细目id = d.Id And A1.开嘱科室id = e.Id " & vbNewLine & _
'                         "    And A1.收费细目id = f.收费细目id(+) And f.性质(+) = 3 And f.码类(+) = 1 And b.发送时间 Between Sysdate - 1 And Sysdate " & vbNewLine & _
'                         "    And b.审方ID <> [1] And A1.病人ID = [2] And A1.病人来源 = 1 And Not A1.Id In (Select 医嘱id From W_A) "
                
                strTmp = "Select 1 标志, e.名称 临床科室, b.审方id, b.医嘱id, A1.超量说明, A1.医生嘱托, A1.相关id, A1.审查结果 PASS, " & vbNewLine & _
                         "    d.名称 药品名称, f.名称 商品名, d.规格, d.计算单位 单位, A1.总给予量 数量, A1.单次用量 || Nvl(g.计算单位, '') 单量, A2.医嘱内容 用法, " & vbNewLine & _
                         "    A1.执行频次 频次, c.标准单价 单价, c.应收金额, d.ID 细目ID, " & vbNewLine & _
                         "    Row_Number() Over(Partition By A1.相关id Order By b.医嘱id) 组序号, Count(1) Over(Partition By A1.相关id) 组数 " & vbNewLine & _
                         "From 病人医嘱记录 A1, 病人医嘱记录 A2, 病人医嘱发送 A3, 处方审查明细 B, (" & vbNewLine & _
                         "    Select a.医嘱序号, Max(a.标准单价) 标准单价, Sum(a.应收金额) 应收金额 " & vbNewLine & _
                         "    From 门诊费用记录 A, 病人医嘱记录 B " & vbNewLine & _
                         "    Where a.医嘱序号 = b.id And a.收费类别 in ('5','6','7') " & vbNewLine & _
                         "        And b.病人id = [2] And b.病人来源 = 1 And a.登记时间 Between Sysdate - 1 And Sysdate " & vbNewLine & _
                         "    Group By a.医嘱序号 " & vbNewLine & _
                         "    ) C, 收费项目目录 D, 部门表 E, 收费项目别名 F, 诊疗项目目录 G " & vbNewLine & _
                         "Where A1.相关id = A2.Id And A1.ID = A3.医嘱ID And A1.Id = b.医嘱id(+) And A1.Id = c.医嘱序号(+) " & vbNewLine & _
                         "    And A1.收费细目id = d.Id And A1.开嘱科室id = e.Id And A1.收费细目id = f.收费细目id(+) And A1.诊疗项目ID = g.ID(+) " & vbNewLine & _
                         "    And b.审方id(+) <> [1] And f.性质(+) = 3 And f.码类(+) = 1 And A1.病人id = [2] And A1.病人来源 = 1 " & vbNewLine & _
                         "    And Not A1.Id In (Select 医嘱id From W_A) And A3.发送时间 Between Sysdate - 1 And Sysdate " & vbNewLine & _
                         "Order By A1.相关id, 组序号 "
                         
                strSQL = "Select * From (With W_A As (" & strSQL & ") " & vbNewLine & _
                         "  Select * From W_A " & vbNewLine & _
                         "  Union All " & vbNewLine & _
                         "  Select * From (" & strTmp & ") ) "
            End If
            
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询审方记录的医嘱信息", lngRAID, lngPatientID)
            Call mdlDefine.FillVSFData(vsfRecipe, rsTemp)
            
            If vsfRecipe.Rows > 1 Then
                With vsfRecipe
'                    '用法、频次合并单元格；要分组合并，必须在最左
'                    .MergeCells = flexMergeRestrictAll
'                    .MergeCol(.ColIndex("相关ID")) = True
'                    .MergeCol(.ColIndex("用法")) = True
'                    .MergeCol(.ColIndex("频次")) = True
                        
                    For l = 1 To .Rows - 1
                        '处理医嘱的组
                        If Val(.TextMatrix(l, .ColIndex("组数"))) >= 3 Then
                            If Val(.TextMatrix(l, .ColIndex("组序号"))) = 1 Then
                                '组首
                                .TextMatrix(l, .ColIndex("组")) = "┏"
                            ElseIf Val(.TextMatrix(l, .ColIndex("组序号"))) = Val(.TextMatrix(l, .ColIndex("组数"))) Then
                                '组尾
                                .TextMatrix(l, .ColIndex("组")) = "┗"
                            Else
                                '组体
                                .TextMatrix(l, .ColIndex("组")) = "┃"
                            End If
                        ElseIf Val(.TextMatrix(l, .ColIndex("组数"))) = 2 Then
                            If Val(.TextMatrix(l, .ColIndex("组序号"))) = 1 Then
                                '组首
                                .TextMatrix(l, .ColIndex("组")) = "┏"
                            Else
                                '组尾
                                .TextMatrix(l, .ColIndex("组")) = "┗"
                            End If
                        End If
                        
                        '处理PASS结果的图片（灯）
                        If .ColHidden(.ColIndex("PASS")) = False And Not mobjPASS Is Nothing Then
                            Set .Cell(flexcpPicture, l, .ColIndex("PASS")) = mobjPASS.zlPassSetWarnLight_YF(Val(.TextMatrix(l, .ColIndex("PASS"))))
                            If Not .Cell(flexcpPicture, l, .ColIndex("PASS")) Is Nothing Then
                                .Cell(flexcpPictureAlignment, l, .ColIndex("PASS")) = flexPicAlignCenterCenter
                            End If
                            .TextMatrix(l, .ColIndex("PASS")) = ""      '不显示文本，只显示图片
                        End If
                        
                        '待审查处方与24小时内的处方区分（灰色背景）
                        .Cell(flexcpBackColor, l, 0, l, .Cols - 1) = IIf(Val(.TextMatrix(l, .ColIndex("标志"))) = 1, &H8000000F, .BackColor)
                        
                        '药品单价与应收金额
                        If Not mobjPubAdvice Is Nothing Then
                            Call mobjPubAdvice.GetDurgPrice(Val(.TextMatrix(l, .ColIndex("医嘱ID"))), dblPrice)
                            dblAmount = dblPrice * Val(.TextMatrix(l, .ColIndex("数量")))
                            .TextMatrix(l, .ColIndex("单价")) = Format(dblPrice, "#0.000")
                            .TextMatrix(l, .ColIndex("应收金额")) = Format(dblAmount, "#0.00")
                        End If
                    Next
                    
                    .Row = 1
                End With
            End If
        
        Case 3      '3-审查项目数据
        
            If Me.tbcRec.Item(0).Selected Then
                Set vsfTmp = vsfRecWait
            Else
                Set vsfTmp = vsfRecFinish
            End If
            
            vsfAudit.ColHidden(vsfAudit.ColIndex("删除")) = tbcRec.Item(0).Selected = False
            vsfAudit.Rows = 1
            vsfAudit.Clear 1
            
            If vsfTmp.Rows <= 1 Then
                mblnReasonRefresh = True
                txtReason.Text = ""
                mblnReasonRefresh = False
                Exit Sub
            End If
            
            lngRAID = Val(vsfRecipe.TextMatrix(vsfRecipe.Row, vsfRecipe.ColIndex("审方ID")))
            
            If lngRAID <= 0 Then
                Exit Sub
            End If
            
            mblnReasonRefresh = True
            If Me.tbcRec(1).Selected Or Val(vsfRecipe.TextMatrix(vsfRecipe.Row, vsfRecipe.ColIndex("标志"))) = 1 Then
                strSQL = "Select 综合理由 From 处方审查记录 Where ID = [1] "
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询处方审查记录信息", lngRAID)
                If rsTemp.EOF = False Then
                    txtReason.Text = NVL(rsTemp!综合理由)
                End If
                rsTemp.Close
            Else
                txtReason.Text = ""
            End If
            mblnReasonRefresh = False
            
            strSQL = "Select Decode(b.药师审查, 2, '不合格', 1, '合格', Decode(b.自动审查, 2, '不合格', '合格')) 药师审查, " & vbNewLine & _
                     "    Decode(b.自动审查, 2, '不合格', 1, '合格', '未审查') 自动审查, " & vbNewLine & _
                     "    b.医嘱id, c.名称 药品, d.id 审查项目ID, d.类别, d.编码, d.简称, d.内容 项目内容 " & vbNewLine & _
                     "From 病人医嘱记录 A, 处方审查结果 B, 收费项目目录 C, 处方审查项目 D " & vbNewLine & _
                     "Where a.Id(+) = b.医嘱id And a.收费细目id = c.Id(+) And b.审查项目id(+) = d.Id And b.审方id(+) = [1] " & vbNewLine & _
                     "    And d.是否门诊启用 = 1 And d.类别 in ([2], 3, 4)" & vbNewLine & _
                     "Order By d.类别, d.编码"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "查询审方审查项目信息", lngRAID, IIf(mintItem = 1, 2, 1))
            Call mdlDefine.FillVSFData(vsfAudit, rsTemp)
            
            If vsfAudit.Rows > 1 Then
                With vsfAudit
                    '编码、简称、项目内容合并单元格
                    .MergeCells = flexMergeRestrictColumns
                    .MergeCol(.ColIndex("编码")) = True
                    .MergeCol(.ColIndex("简称")) = True
                    .MergeCol(.ColIndex("项目内容")) = True
                    
                    blnAllPass = True
                    For l = 1 To .Rows - 1
                        '合格、不合格的图片
                        If .TextMatrix(l, .ColIndex("药师审查")) = "合格" Then
                            .Cell(flexcpPicture, l, .ColIndex("药师审查")) = picAuditYN(0).Picture
                        Else
                            blnAllPass = False
                            .Cell(flexcpPicture, l, .ColIndex("药师审查")) = picAuditYN(1).Picture
                            '药师审查不合格的记录（浅红色）
                            .Cell(flexcpBackColor, l, 0, l, .Cols - 1) = &HC0C0FF
                        End If
                        
                        If .TextMatrix(l, .ColIndex("自动审查")) = "合格" Then
                            .Cell(flexcpPicture, l, .ColIndex("自动审查")) = picAuditYN(0).Picture
                        ElseIf .TextMatrix(l, .ColIndex("自动审查")) = "不合格" Then
                            .Cell(flexcpPicture, l, .ColIndex("自动审查")) = picAuditYN(1).Picture
                        Else
                            .Cell(flexcpPicture, l, .ColIndex("自动审查")) = Nothing
                            .Cell(flexcpText, l, .ColIndex("自动审查")) = "未审查"
                        End If
                    Next
                    
                    .Row = 1
                End With
            End If
            
            '更新显示lblAudit控件
            Call mdlRecipeAudit.DispCountNG(vsfAudit, lblAudit)
            
'            '所有项目合格，默认“合格、不合格”按钮允许操作
'            If blnAllPass And Me.tbcRec.Item(0).Selected Then
'                '自动加锁
'                Call txtReason_Change
'            End If
        
        Case Else
            '
    End Select
    
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub tbcTab_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Me.Visible And Item.Index = 1 Then
        If Me.tbcRec.Item(0).Selected Then
            Call RefreshAMR(vsfRecWait)
        Else
            Call RefreshAMR(vsfRecFinish)
        End If
    End If
End Sub

Private Sub timPatient_Timer()
    If tbcRec.Item(0).Selected = False Then Exit Sub
    If vsfRecWait.Rows > 1 Or mblnLocking = True Then
        Exit Sub
    End If

    Screen.MousePointer = vbHourglass
            
    Call FillVSFData(1)
    Call FillVSFData(2)
    Call RefreshLockControls
    Call SetStatusbar
    Call RefreshAMR
    
    Screen.MousePointer = vbDefault
End Sub

Private Sub txtPatient_Change()
'    iknPatient.SetAutoReadCard Trim(txtPatient.Text) = ""
End Sub

Private Sub txtPatient_GotFocus()
    Call zlControl.TxtSelAll(txtPatient)
'    iknPatient.SetAutoReadCard Trim(txtPatient.Text) = ""
End Sub

Private Sub txtPatient_KeyPress(KeyAscii As Integer)
    Dim strCard As String
    
    strCard = iknPatient.Cards(iknPatient.IDKind).名称

    If mblnReadCard Or KeyAscii = 13 Then
        Call zlControl.TxtSelAll(txtPatient)
    Else
        Select Case strCard
            Case "门诊号"
                If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
            Case "身份证号"
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
            Case Else
                If InStr(":：;；?？", Chr(KeyAscii)) > 0 Then
                    KeyAscii = 0
                Else
                    KeyAscii = Asc(UCase(Chr(KeyAscii)))
                End If
        End Select
    End If
    mblnReadCard = False
End Sub

Private Sub txtPatient_LostFocus()
'    iknPatient.SetAutoReadCard False
End Sub

Private Sub txtReason_Change()
    Dim lngAuditID As Long
    
    If tbcRec.Item(0).Selected = False Or mblnNeedAudit = False Or mblnAuditStart = False Or mblnReasonRefresh Then Exit Sub
    
    If mblnLocking = False Then
        lngAuditID = Val(vsfRecWait.TextMatrix(vsfRecWait.Row, vsfRecWait.ColIndex("审方ID")))
        Call AuditLock(lngAuditID, True)
    End If
End Sub

Private Sub txtReason_KeyPress(KeyAscii As Integer)
    If InStr("""'", KeyAscii) Then KeyAscii = 0
End Sub

Private Sub vsfAudit_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim blnSign As Boolean
    
    If vsfRecipe.Rows > 1 And vsfRecipe.Row > 0 Then
        blnSign = Val(vsfRecipe.TextMatrix(vsfRecipe.Row, vsfRecipe.ColIndex("标志"))) = 1
    Else
        blnSign = True
    End If

    If Me.tbcRec.Item(0).Selected And mblnNeedAudit And mblnAuditStart And blnSign = False Then
        If Col = vsfAudit.ColIndex("药品") Then
            Cancel = False
        Else
            Cancel = True
        End If
    Else
        Cancel = True
    End If

End Sub

Private Sub vsfAudit_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    If Me.tbcRec.Item(0).Selected = False Then Exit Sub
    
    Dim rsTemp As ADODB.Recordset
    Dim strSQL As String, strTmp As String
    Dim blnCancel As Boolean, blnFind As Boolean
    Dim l As Long, lngItemID As Long, lngAuditID As Long
    Dim intCol As Integer
    Dim colTmp As New Collection
    
    If Col = vsfAudit.ColIndex("药品") Then
    
        lngItemID = Val(vsfAudit.TextMatrix(Row, vsfAudit.ColIndex("审查项目ID")))
        lngAuditID = Val(vsfRecWait.TextMatrix(vsfRecWait.Row, vsfRecWait.ColIndex("审方ID")))
        
        strSQL = "Select b.医嘱id ID, d.名称 药品名称, d.规格, d.计算单位 单位, " & vbNewLine & _
                 "    A1.总给予量 数量, A1.单次用量 || Nvl(f.计算单位, '') 单量, A2.医嘱内容 用法, A1.执行频次 频次, " & vbNewLine & _
                 "    e.名称 临床科室, b.审方id  " & vbNewLine & _
                 "From 病人医嘱记录 A1, 病人医嘱记录 A2, 处方审查明细 B, 收费项目目录 D, 部门表 E, 诊疗项目目录 F " & vbNewLine & _
                 "Where A1.相关id = A2.Id And A1.Id = b.医嘱id And A1.收费细目id = d.Id And A1.开嘱科室id = e.Id And A1.诊疗项目id = f.id(+) " & vbNewLine & _
                 "    And b.审方id = [1] " & vbNewLine & _
                 "Order By A1.相关id, A1.Id "

        Set rsTemp = zlDatabase.ShowSQLMultiSelect(Me, strSQL, 0, "选择审查项目对应的药品", False, "", "", _
                            False, False, False, 0, 0, 2000, blnCancel, False, False, _
                            lngAuditID)
        If blnCancel = False Then
            '将选择的数据写入
            If Not rsTemp Is Nothing Then
                With rsTemp
                    Do While .EOF = False
                        '查找医嘱ID是否存在
                        blnFind = False
                        For l = 1 To vsfAudit.Rows - 1
                            If Val(vsfAudit.TextMatrix(l, vsfAudit.ColIndex("医嘱ID"))) = !ID And Val(vsfAudit.TextMatrix(l, vsfAudit.ColIndex("审查项目ID"))) = lngItemID Then
                                blnFind = True
                                Exit For
                            End If
                        Next
                        
                        '医嘱ID不存在，再插入
                        If blnFind = False Then
                            For l = vsfAudit.Rows - 1 To 1 Step -1
                                If Val(vsfAudit.TextMatrix(l, vsfAudit.ColIndex("审查项目ID"))) = lngItemID Then
                                    '复制当前记录
                                    Set colTmp = Nothing
                                    For intCol = 0 To vsfAudit.Cols - 1
                                        If intCol = vsfAudit.ColIndex("药师审查") Then
                                            colTmp.Add "不合格", CStr(intCol)
                                        ElseIf intCol = vsfAudit.ColIndex("自动审查") Then
                                            colTmp.Add "合格", CStr(intCol)
                                        ElseIf intCol = vsfAudit.ColIndex("药品") Then
                                            colTmp.Add !药品名称, CStr(intCol)
                                        ElseIf intCol = vsfAudit.ColIndex("医嘱ID") Then
                                            colTmp.Add !ID, CStr(intCol)
                                        ElseIf intCol = vsfAudit.ColIndex("新增") Then
                                            colTmp.Add "1", CStr(intCol)
                                        Else
                                            colTmp.Add vsfAudit.TextMatrix(l, intCol), CStr(intCol)
                                        End If
                                    Next
                                    
                                    '插入新记录
                                    If Not colTmp Is Nothing Then
                                        l = l + 1
                                        vsfAudit.AddItem "", l
                                        For intCol = 1 To colTmp.Count
                                            vsfAudit.TextMatrix(l, intCol - 1) = colTmp(intCol)
                                            If intCol - 1 = vsfAudit.ColIndex("药师审查") Or intCol - 1 = vsfAudit.ColIndex("自动审查") Then
                                                vsfAudit.Cell(flexcpPicture, l, intCol - 1) = IIf(colTmp(intCol) = "合格", picAuditYN(0).Picture, picAuditYN(1).Picture)
                                            ElseIf intCol - 1 = vsfAudit.ColIndex("删除") Then
                                                vsfAudit.Cell(flexcpPicture, l, intCol - 1) = picAuditDel.Picture
                                                vsfAudit.Cell(flexcpPictureAlignment, l, intCol - 1) = flexPicAlignCenterCenter
                                            End If
                                        Next
                                        '药师审查不合格的记录（浅红色）
                                        vsfAudit.Cell(flexcpBackColor, l, 0, l, vsfAudit.Cols - 1) = &HC0C0FF
                                    End If
                                                                    
                                    Exit For
                                End If
                            Next
                        End If
                        
                        .MoveNext
                    Loop
                    .Close
                End With
            End If
            
            Call DispCountNG(vsfAudit, lblAudit)
            
            '加锁
            If mblnLocking = False Then
                Call AuditLock(lngAuditID, True)
            End If
            
        End If
    End If
End Sub

Private Sub vsfAudit_Click()
    Dim lngAuditID As Long
    
    If Me.tbcRec.Item(0).Selected = False Or mblnNeedAudit = False Or mblnAuditStart = False Then Exit Sub
    
    If vsfRecipe.Rows > 1 And vsfRecipe.Row > 0 Then
        If Val(vsfRecipe.TextMatrix(vsfRecipe.Row, vsfRecipe.ColIndex("标志"))) = 1 Then
            Exit Sub
        End If
    Else
        Exit Sub
    End If

    With vsfAudit
        If .Rows <= 1 Then Exit Sub
        
        If .ColIndex("药师审查") = .Col Then
            .TextMatrix(.Row, .Col) = IIf(.TextMatrix(.Row, .Col) = "不合格", "合格", "不合格")
            If .TextMatrix(.Row, .Col) = "合格" Then
                .Cell(flexcpPicture, .Row, .Col) = picAuditYN(0).Picture
                .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = .BackColor
            Else
                .Cell(flexcpPicture, .Row, .Col) = picAuditYN(1).Picture
                '药师审查不合格的记录（浅红色）
                .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = &HC0C0FF
            End If
            
            Call DispCountNG(vsfAudit, lblAudit)
            
            If mblnLocking = False Then
                lngAuditID = Val(vsfRecWait.TextMatrix(vsfRecWait.Row, vsfRecWait.ColIndex("审方ID")))
                Call AuditLock(lngAuditID, True)
            End If
        ElseIf .ColIndex("删除") = .Col Then
            If Val(.TextMatrix(.Row, .ColIndex("新增"))) = 1 Then
                If MsgBox("确认删除该记录？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    .RemoveItem .Row
                    Call DispCountNG(vsfAudit, lblAudit)
                End If
            End If
        End If
        
    End With
End Sub

Private Sub vsfAudit_KeyPress(KeyAscii As Integer)
    If vsfAudit.Col = vsfAudit.ColIndex("药师审查") Then
        If KeyAscii = vbKeySpace Then Call vsfAudit_Click
    End If
End Sub

Private Sub vsfRecFinish_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <> NewRow Then
        Dim lngPatientID As Long, lngMedID As Long
        
        If vsfRecFinish.TextMatrix(NewRow, vsfRecFinish.ColIndex("审查结果")) = "合格" Then
            lblDisp_Fixed.Caption = "合格"
            lblDisp_Fixed.ForeColor = &H8000&
        ElseIf vsfRecFinish.TextMatrix(NewRow, vsfRecFinish.ColIndex("审查结果")) = "不合格" Then
            lblDisp_Fixed.Caption = "不合格"
            lblDisp_Fixed.ForeColor = vbRed
        Else
            lblDisp_Fixed.Caption = "未知"
            lblDisp_Fixed.ForeColor = vbBlack
        End If
        lblDisp_Fixed.Left = (picYesNO.ScaleWidth - lblDisp_Fixed.Width) \ 2
        
        Call FillVSFData(2)
        'Call FillVSFData(3)
        
        '刷新电子病案
        If Me.tbcTab.ItemCount > 1 Then
            If Me.tbcTab.Item(1).Selected Then
                Call RefreshAMR(vsfRecFinish)
            End If
        End If
        
        If vsfRecFinish.Visible And vsfRecFinish.Enabled Then vsfRecFinish.SetFocus
    End If
End Sub

Private Sub vsfRecipe_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <> NewRow Then
        If NewRow > 0 Then 'If Screen.MousePointer = vbDefault And NewRow > 0 Then
            Call SetOtherText
            If Val(vsfRecipe.TextMatrix(NewRow, vsfRecipe.ColIndex("标志"))) = 1 Then
                vsfAudit.Clear 1
                vsfAudit.Rows = 1
            Else
                Call FillVSFData(3)
            End If
            Call RefreshLockControls
            If Not mobjPASS Is Nothing Then
                Call mobjPASS.zlPassSetDrug_YF(vsfRecipe.TextMatrix(NewRow, vsfRecipe.ColIndex("细目ID")), _
                                               vsfRecipe.TextMatrix(NewRow, vsfRecipe.ColIndex("药品名称")))
            End If
        Else
            Call SetOtherText
            Call FillVSFData(3)
            Call RefreshLockControls
        End If
    End If
End Sub

Private Sub vsfRecipe_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strTip As String
    
    With vsfRecipe
        If .Rows <= 1 Or .MouseCol < 0 Then Exit Sub
        If .MouseRow < 0 Then Exit Sub
        
        On Error Resume Next
        If (.TextMatrix(.MouseRow, .MouseCol) <> "" And .MouseCol = .ColIndex("超量说明") _
            Or .TextMatrix(.MouseRow, .MouseCol) <> "" And .MouseCol = .ColIndex("医生嘱托")) Then
            strTip = .TextMatrix(.MouseRow, .MouseCol)
            fs.ShowTipInfo .hwnd, strTip, True
        Else
            fs.ShowTipInfo 0, "", True
        End If
        On Error GoTo 0
    End With
End Sub

Private Sub vsfRecWait_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <> NewRow Then
        Dim lngPatientID As Long, lngMedID As Long
        
        mblnLocking = False
        Call FillVSFData(2)
        'Call FillVSFData(3)
        Call RefreshLockControls
        
        '刷新电子病案
        If Me.tbcTab.ItemCount > 1 Then
            If Me.tbcTab.Item(1).Selected Then
                Call RefreshAMR(vsfRecWait)
            End If
        End If
        
        If vsfRecWait.Visible And vsfRecWait.Enabled Then vsfRecWait.SetFocus
    End If
End Sub

Private Sub SetControlFontSize(ByVal bytSize As Byte)
'功能：设置窗体控件的字体大小
'参数：
'  bytSize：0-小字体；1-大字体

    mbytFontSize = bytSize
    
    Call SetPublicFontSize(Me, bytSize)
    Call picRec_Resize
    
End Sub

Private Sub RefreshLockControls()
'功能：刷新审查锁定或非锁定状态的相关控件
    
    Dim blnAllPass As Boolean
    
    blnAllPass = GetAuditResult(vsfAudit)
    
    '所有项目合格，默认“合格、不合格”按钮允许操作
    cmdYes_Fixed.Enabled = mblnAuditStart And (blnAllPass Or mblnLocking)
    cmdNo_Fixed.Enabled = mblnAuditStart And (blnAllPass Or mblnLocking)
    
    txtReason.Enabled = mblnAuditStart And (vsfAudit.Rows > 1) And tbcRec.Item(0).Selected
    If vsfRecipe.ColIndex("标志") >= 0 Then
        txtReason.Enabled = txtReason.Enabled And Val(vsfRecipe.TextMatrix(vsfRecipe.Row, vsfRecipe.ColIndex("标志"))) = 0
    End If
    If tbcRec.ItemCount > 1 Then tbcRec.Item(1).Enabled = Not mblnLocking
    chk24.Enabled = Not mblnLocking
    chkPubPatient.Enabled = Not mblnLocking
    sbnPASS.Enabled = Not mblnLocking And (mintParaPass > 0 And mintParaPass < 5) And Not mobjPASS Is Nothing
    sbnHouse.Enabled = mblnLocking
    sbnSelect.Enabled = mblnAuditStart  'mblnLocking
    
    vsfRecipe.Enabled = Not mblnLocking
    Call cboClinic_Click
End Sub

Private Sub AuditProcess(ByVal bytFun As Byte)
'功能：处方审查处理
'参数：
'  bytFun：1-合格；2-不合格

    Dim strSQL As String, strReason As String, strTmp As String
    Dim lngAuditID As Long, lngMedicalID As Long, lngItemID As Long
    Dim l As Long
    Dim colSQL As New Collection
    Dim blnFind As Boolean, blnNoTrans As Boolean
    Dim strIDs As String
    Dim lngPatientID As Long, lngRegisterID As Long
    
    '提交前的检查
    
    If LenB(StrConv(Trim(txtReason.Text), vbFromUnicode)) > 500 Then
        MsgBox "“综合理由”内容超限（250个汉字或500个字符）！", vbInformation, gstrSysName
        txtReason.SetFocus
        Exit Sub
    End If
    
    strReason = Trim(txtReason.Text)
    strReason = Replace(Replace(Replace(strReason, vbLf, ""), vbCr, ""), vbNewLine, "")
    
    With vsfAudit
        blnFind = False
        For l = 1 To .Rows - 1
            If bytFun = 1 Then
                '检查药师最终判定合格，但明细有不合格的情况
                If Trim(.TextMatrix(l, .ColIndex("药师审查"))) = "不合格" Then
                    MsgBox "“审查项目”的“药师审查”列与最终的结果不相符，请检查！", vbInformation, gstrSysName
                    .Row = l
                    .SetFocus
                    Exit Sub
                End If
            Else
                '检查药师最终判定不合格，但明细没有不合格的情况
                If .TextMatrix(l, .ColIndex("药师审查")) = "不合格" Then
                    blnFind = True
                End If
            End If
            
            '检查综合理由；自动审查或药师审查有“不合格”的必须填写综合理由
            If strReason = "" And (Trim(.TextMatrix(l, .ColIndex("药师审查"))) = "不合格" Or Trim(.TextMatrix(l, .ColIndex("自动审查"))) = "不合格") Then
                MsgBox "请填写审查的“综合理由”！", vbInformation, gstrSysName
                .Row = l
                If txtReason.Enabled And txtReason.Visible Then txtReason.SetFocus
                Exit Sub
            
            '检查审查项目ID<=0的记录
            ElseIf Val(.TextMatrix(l, .ColIndex("审查项目ID"))) <= 0 Then
                MsgBox "审查项目异常！", vbInformation, gstrSysName
                .Row = l
                .SetFocus
                Exit Sub
            End If
        Next
        
        If bytFun = 2 And blnFind = False Then
            MsgBox "“审查项目”的“药师审查”列无不合格，请检查！", vbInformation, gstrSysName
            .Row = 1
            .SetFocus
            Exit Sub
        End If
    End With
    
    lngAuditID = Val(vsfRecWait.TextMatrix(vsfRecWait.Row, vsfRecWait.ColIndex("审方ID")))
    If lngAuditID <= 0 Then
        MsgBox "待审记录异常，审查操作终止！", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '处方审查记录信息
    
    strSQL = zlStr.FormatString("Zl_处方审查_Audit([1], [2], [3], [4])", _
                    lngAuditID, _
                    bytFun, _
                    "'" & UserInfo.姓名 & "'", _
                    IIf(strReason = "", "Null", "'" & Trim(txtReason.Text) & "'"))
    AddArray colSQL, strSQL
    
    '处方审查明细信息
    
    With vsfAudit
        For l = 1 To .Rows - 1
            If .TextMatrix(l, .ColIndex("药师审查")) = "不合格" Or .TextMatrix(l, .ColIndex("自动审查")) = "不合格" Then
                lngMedicalID = Val(.TextMatrix(l, .ColIndex("医嘱ID")))
                lngItemID = Val(.TextMatrix(l, .ColIndex("审查项目ID")))
                
                strSQL = zlStr.FormatString("Zl_处方审查_Audit_Detail([1], [2], [3], [4])", _
                                lngAuditID, _
                                IIf(lngMedicalID <= 0, "Null", lngMedicalID), _
                                lngItemID, _
                                IIf(.TextMatrix(l, .ColIndex("药师审查")) = "合格", 1, 2))
                AddArray colSQL, strSQL
            End If
        Next
    End With
    
    '更新最后操作时间
    strSQL = zlStr.FormatString("ZL_处方审查参数_SAVE(2, '[1]', 0, 1, Null)", mstrPCName)
    AddArray colSQL, strSQL
    
    '执行存储过程
    Err.Clear: On Error GoTo errHandle: blnNoTrans = False
    ExecuteProcedureArray colSQL, Me.Caption, blnNoTrans
    blnNoTrans = True
    
    '获取病人ID、挂号单ID
    With vsfRecWait
        lngPatientID = Val(.TextMatrix(.Row, .ColIndex("病人ID")))
        lngRegisterID = Val(.TextMatrix(.Row, .ColIndex("挂号单ID")))
    End With
    
    '获取相关IDs
    With vsfRecipe
        For l = 1 To .Rows - 1
            strIDs = strIDs & "," & .TextMatrix(l, .ColIndex("相关ID"))
        Next
        If strIDs <> "" Then strIDs = Mid(strIDs, 2)
    End With
    
    mblnReasonRefresh = True
    txtReason.Text = ""
    mblnReasonRefresh = False
    
    '成功完成，状态调整
    mblnLocking = False
'    vsfRecWait.RemoveItem vsfRecWait.Row
    
    Screen.MousePointer = vbHourglass
    Call FillVSFData(1)
    Call FillVSFData(2)
    Call RefreshLockControls
    Call SetStatusbar
    Call RefreshAMR
    
    '门诊医嘱自动发送
    If bytFun = Val("1-合格") And Not mobjPubAdvice Is Nothing Then
        If mblnSendBeforeAudit And mblnRecipeSendAuto Then
            '参数控制：处方发送前审方；审方合格自动发送处方
            If strIDs <> "" And mobjPubAdvice.OutAdviceSendDrug(Me, strIDs, lngPatientID, lngRegisterID) Then
                '门诊医嘱自动发送成功不发送消息
            Else
                '发送消息通知医生
                SendMessage bytFun + 10, lngAuditID, 1, mobjMipModule, mblnSendBeforeAudit
            End If
        Else
            SendMessage bytFun, lngAuditID, 1, mobjMipModule, mblnSendBeforeAudit
        End If
    Else
        SendMessage bytFun, lngAuditID, 1, mobjMipModule, mblnSendBeforeAudit
    End If
    
    Screen.MousePointer = vbDefault
    
    Exit Sub
    
errHandle:
    If ErrCenter() = 1 Then
        Resume
    Else
        If blnNoTrans = False Then gcnOracle.RollbackTrans
    End If
    If Screen.MousePointer = vbHourglass Then Screen.MousePointer = vbDefault
End Sub

Private Sub AuditLock(ByVal lngAuditID As Long, Optional ByVal blnLock As Boolean = True)
'功能：加锁/解锁切换
'参数：
'  lngAuditID：审方ID
'  blnLock：True加锁；False解锁
    
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    strSQL = zlStr.FormatString("Zl_处方审查记录_Lock([1], '[2]', 0, [3])", _
                    IIf(blnLock, 1, 0), _
                    mstrPCName, _
                    IIf(lngAuditID <= 0, "Null", lngAuditID))
    Call zlDatabase.ExecuteProcedure(strSQL, IIf(blnLock, "审查记录加锁", "审查记录解锁"))
    
    mblnLocking = blnLock
    Call RefreshLockControls
    Exit Sub
    
errHandle:
    Call ErrCenter
    If gcnOracle.Errors(0).Description Like "*已被审查*" Or gcnOracle.Errors(0).Description Like "*已被删除*" Then
        vsfRecWait.RemoveItem vsfRecWait.Row
    End If
    Call FillVSFData(3)
    Call RefreshLockControls
End Sub

Private Sub vsfRecWait_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If OldRow <> NewRow Then
        If mblnLocking Then
            If MsgBox("正在审查当前病人的药嘱，是否放弃审查？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True
            Else
                Call AuditLock(Val(vsfRecWait.TextMatrix(OldRow, vsfRecWait.ColIndex("审方ID"))), False)
            End If
        End If
    End If
End Sub

Private Function AuditOperate(ByVal bytFun As Byte) As Boolean
'功能：当前机器开启/停止处方审查
'参数：
'  bytFun：0-停止；1-开启
'返回：True成功；False失败
    
    'Dim lngAuditID As Long
    Dim strSQL As String

    On Error GoTo errHandle

    If bytFun = 1 Then
        '开启
        strSQL = zlStr.FormatString("ZL_处方审查参数_SAVE(2, '[1]', 0, 1, Null)", mstrPCName)
        Call zlDatabase.ExecuteProcedure(strSQL, "更新开启审方、最后操作时间")
    Else
        '停止
        If mblnLocking Then
            '正在审查切换至停止审查
            If MsgBox("正在审查当前病人的药嘱，是否放弃审查？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
        
        '更新是否开启审方、最后操作时间
        strSQL = zlStr.FormatString("ZL_处方审查参数_SAVE(2, '[1]', 0, 0, Null)", mstrPCName)
        Call zlDatabase.ExecuteProcedure(strSQL, "更新停止审方、最后操作时间")
        
        '解锁
        Call AuditLock(0, False)
        
    End If
    
    AuditOperate = True
    Exit Function

errHandle:
End Function

Private Sub SetOtherText()
    Dim vsfTmp As VSFlexGrid
    Dim lngPatientID As Long, lngRegisterID As Long, lngMedicalID As Long
    Dim strDiagnose As String, strTmp As String
    
    'If Me.Visible = False Then Exit Sub
    
    If tbcRec.Item(0).Selected Then
        Set vsfTmp = vsfRecWait
    Else
        Set vsfTmp = vsfRecFinish
    End If
    
    txtOther.Text = zlStr.FormatString("诊断：[1]热量需要量：", vbCrLf)
    
    If vsfTmp.Rows <= 1 Then Exit Sub
    
    lngPatientID = Val(vsfTmp.TextMatrix(vsfTmp.Row, vsfTmp.ColIndex("病人ID")))
    lngRegisterID = Val(vsfTmp.TextMatrix(vsfTmp.Row, vsfTmp.ColIndex("挂号单ID")))
    lngMedicalID = Val(vsfRecipe.TextMatrix(vsfRecipe.Row, vsfRecipe.ColIndex("相关ID")))
    
    '诊断
    If Not mobjPubAdvice Is Nothing Then
        Call mobjPubAdvice.GetAdviceDiag(lngMedicalID, strDiagnose)
    End If
    txtOther.Text = zlStr.FormatString("诊断：[1][2]", strDiagnose, vbCrLf)
    
    '热量需要量
    strTmp = GetCalorie(lngPatientID, lngRegisterID, 0)
    txtOther.Text = txtOther & zlStr.FormatString("热量需要量：[1]", strTmp)
End Sub

Private Sub SetStatusbar()
    Dim vsfTmp As VSFlexGrid
    
    If tbcRec.Item(0).Selected Then
        Set vsfTmp = vsfRecWait
    Else
        Set vsfTmp = vsfRecFinish
    End If
    
    If vsfTmp.Rows <= 1 Then
        stbThis.Panels(2).Text = ""
    Else
        stbThis.Panels(2).Text = zlStr.FormatString("当前[1]审记录数量：[2]条", IIf(tbcRec.Item(0).Selected, "待", "已"), vsfTmp.Rows - 1)
    End If
    
End Sub

Private Sub SetFilterDay(ByVal bytMode As Byte)
    If tbcRec.ItemCount <= 1 Then Exit Sub

    If bytMode = 0 Then
        tbcRec.Item(1).Tag = cboDate.ListIndex
        With cboDate
            .Clear
            .AddItem "当天"
            .AddItem "两天内"
            .AddItem "三天内"
        End With
    Else
        tbcRec.Item(0).Tag = cboDate.ListIndex
        With cboDate
            .Clear
            .AddItem "当天"
            .AddItem "两天内"
            .AddItem "三天内"
            .AddItem "四天内"
            .AddItem "五天内"
            .AddItem "六天内"
            .AddItem "七天内"
        End With
    End If
    
    If Val(tbcRec.Item(bytMode).Tag) >= 0 Then
        cboDate.ListIndex = Val(tbcRec.Item(bytMode).Tag)
    Else
        cboDate.ListIndex = 0
    End If
    
End Sub

Private Sub SetClinicItem()
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim blnDept As Boolean
    
    On Error GoTo errHandle
    
    '处方审查条件是否设置
    strSQL = "Select Count(1) Rec From 处方审查条件 Where Rownum < 2 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取处方审查条件是否设置")
    If rsTemp!Rec <= 0 Then
        rsTemp.Close
        If Me.Visible = False Then
            MsgBox "“处方审查条件”未做任何设置，请检查！", vbInformation, gstrSysName
        End If
        Exit Sub
    End If
    rsTemp.Close
    
    '有科室设置，按设置科室提取
    strSQL = "Select a.Id 部门id, a.编码, a.名称 " & vbNewLine & _
             "From 部门表 A, Table(f_Num2list((Select 来源科室 From 处方审查参数 Where 机器名 = [1] And 服务对象 = 0), ',')) B " & vbNewLine & _
             "Where a.Id = b.Column_Value " & vbNewLine & _
             "Order By a.名称 "
'    strSQL = "Select a.部门id, b.编码, b.名称 " & vbNewLine & _
'             "From 部门性质说明 A, 部门表 B, 处方审查条件 C," & vbNewLine & _
'             "    Table(f_Num2list((Select 来源科室 From 处方审查参数 Where 机器名 = [1] And 服务对象 = 0))) D " & vbNewLine & _
'             "Where a.部门id = b.Id And a.部门id = c.科室id And a.部门id = d.Column_Value And a.工作性质 = '临床' " & vbNewLine & _
'             "    And a.服务对象 In (1, 3) And (b.撤档时间 Is Null Or To_Char(b.撤档时间, 'yyyy') = '3000') " & vbNewLine & _
'             "    And c.类别 = 1 And (c.科室id Is Not Null Or c.科室id > 0) "

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "获取门诊处方审查的来源科室", mstrPCName)
    
    With rsTemp
        cboClinic.Clear
        cboClinic.AddItem "所有来源科室"
        Do While .EOF = False
            cboClinic.AddItem !名称
            cboClinic.ItemData(cboClinic.NewIndex) = !部门ID
            .MoveNext
        Loop
        .Close
        If cboClinic.ListCount > 0 Then cboClinic.ListIndex = 0
    End With
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub RefreshAMR(Optional ByVal vsfVal As VSFlexGrid)
    Dim lngPatientID As Long, lngMediID As Long

    '刷新临床信息
    If Me.Visible Then
        If vsfVal Is Nothing Then
            If Me.tbcRec.Item(0).Selected Then
                lngPatientID = Val(vsfRecWait.TextMatrix(vsfRecWait.Row, vsfRecWait.ColIndex("病人ID")))
                lngMediID = Val(vsfRecWait.TextMatrix(vsfRecWait.Row, vsfRecWait.ColIndex("挂号单ID")))
            Else
                lngPatientID = Val(vsfRecFinish.TextMatrix(vsfRecFinish.Row, vsfRecFinish.ColIndex("病人ID")))
                lngMediID = Val(vsfRecFinish.TextMatrix(vsfRecFinish.Row, vsfRecFinish.ColIndex("挂号单ID")))
            End If
        Else
            lngPatientID = Val(vsfVal.TextMatrix(vsfVal.Row, vsfVal.ColIndex("病人ID")))
            lngMediID = Val(vsfVal.TextMatrix(vsfVal.Row, vsfVal.ColIndex("挂号单ID")))
        End If
        On Error GoTo errHandle
'        If Not mobjArchiveMedRec Is Nothing And Not mfrmAMR Is Nothing Then
'            Call mobjArchiveMedRec.zlRefresh(0, lngPatientID, lngMediID, False)
'        End If
        If Not mobjPubAdvice Is Nothing And Not mfrmAMR Is Nothing Then
            Call mobjPubAdvice.zlArchiveRefresh(lngPatientID, lngMediID)
        End If
    End If
    Exit Sub
    
errHandle:
    Call ErrCenter
End Sub

Private Sub SetPatientTimer()
    With timPatient
        .Enabled = False
        .Interval = Val(zlDatabase.GetPara("自动刷新病人列表", glngSys, mlngModule)) * 1000
        If .Interval > 0 Then .Enabled = True
    End With
End Sub

