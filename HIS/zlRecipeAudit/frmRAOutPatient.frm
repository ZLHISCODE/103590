VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{AF9744ED-CAFC-4877-8437-2C20C14CEA4E}#6.10#0"; "zlIDKind.ocx"
Begin VB.Form frmRAOutPatient 
   Caption         =   "���ﴦ�����"
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
   StartUpPosition =   1  '����������
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
               Name            =   "����"
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
               Caption         =   "�ϸ�"
               BeginProperty Font 
                  Name            =   "����"
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
               Caption         =   "���ϸ�(&N)"
               BeginProperty Font 
                  Name            =   "����"
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
               Caption         =   "�ϸ�(&Y)"
               BeginProperty Font 
                  Name            =   "����"
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
               Name            =   "����"
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
            Caption         =   "�ۺ�����"
            Height          =   180
            Left            =   120
            TabIndex        =   35
            Top             =   120
            Width           =   720
         End
         Begin VB.Label lblAudit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�����Ŀ"
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
               Name            =   "����"
               Size            =   12
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Caption         =   "�鿴PASS���"
            Enabled         =   -1  'True
            PictureAlign    =   0
            Picture         =   "frmRAOutPatient.frx":6CFA
            ShowCaption     =   -1  'True
         End
         Begin VB.CheckBox chk24 
            BackColor       =   &H80000002&
            Caption         =   "��ʾ24Сʱ������ҩ��(&P)"
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
            Caption         =   "������ϸ"
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
         Caption         =   "����λδ��ȷ��Դ���ҵĴ�������(&I)"
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
         IDKindStr       =   "��|����|0|0|0|0|0|0;��|�����|0|0|0|0|0|0"
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         FontSize        =   9
         FontName        =   "����"
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
         Caption         =   "�ύʱ��(&T)"
         Height          =   180
         Left            =   120
         TabIndex        =   26
         Top             =   1020
         Width           =   990
      End
      Begin VB.Label lblDrugstore 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ҩҩ��(&D)"
         Height          =   180
         Left            =   120
         TabIndex        =   24
         Top             =   630
         Width           =   990
      End
      Begin VB.Label lblClinic 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��Դ����(&F)"
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
            Text            =   "�������"
            TextSave        =   "�������"
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
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   88
            Text            =   "��д"
            TextSave        =   "��д"
            Key             =   "STACAPS"
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

Private Const MSTR_PATIENT = "��|����|0;��|�����|0;��|���֤��|1"

Private Const MSTR_VSF_WAIT As String = _
        "����,,3,1000|��ID,,0,0|����ID,,0,0|�Һŵ�ID,,0,0|��ҳID,,0,0|�Ա�,,3,600|����,,3,600|�����,,3,1000|�ύʱ��,,3,1600" & _
        "|�ύ����,,3,1500|�ύ��,,3,1000|��ҩҩ��,,3,1500"
        
Private Const MSTR_VSF_FINISH As String = _
        "�����,,3,1000|����,,3,1000|��ID,,0,0|����ID,,0,0|�Һŵ�ID,,0,0|��ҳID,,0,0|�Ա�,,3,600|����,,3,600|�����,,3,1000|�ύʱ��,,3,1600" & _
        "|�ύ����,,3,1500|�ύ��,,3,1000|��ҩҩ��,,3,1500|���ʱ��,,3,1600|�����,,3,1000"
        
Private Const MSTR_VSF_RECIPE As String = _
        "�ٴ�����,,3,1500|��ID,,0,0|ҽ��ID,,0,0|����˵��,,3,1500|ҽ������,,3,1500|�����,,0,0|����,,0,0|��,,3,300|PASS,,3,600" & _
        "|ҩƷ����,,3,2000|��Ʒ��,,3,1500|���,,3,1500|��λ,,3,600|����,,3,600,n|����,,3,800,n|���ID,,0,0|�÷�,,3,800|Ƶ��,,3,800" & _
        "|����,,3,800,n|Ӧ�ս��,,3,1000,n|���ϸ�,,0,0|ϸĿID,,0,0|��־,,0,0"

Private Const MSTR_VSF_AUDIT As String = _
        "ҩʦ���,,3,1000|�Զ����,,3,1000|ҩƷ,,3,2500|ɾ��,,3,600|����,,3,1000|���,,3,2000|��Ŀ����,,3,4000|ҽ��ID,,0,0|�����ĿID,,0,0" & _
        "|����,,0,0"
        
Private WithEvents mobjMipModule As zl9ComLib.clsMipModule      '��Ϣƽ̨
Attribute mobjMipModule.VB_VarHelpID = -1
'Private mobjArchiveMedRec As zlPublicAdvice.clsArchiveMedRec    '���Ӳ���
Private mobjPubAdvice As zlPublicAdvice.clsPublicAdvice         '�ٴ�������
Private mobjPASS As Object                                      '������ҩ�ӿڲ���
Private mfrmAMR As Form

Private mlngModule As Long              'ģ���
Private mstrPrivs As String             'Ȩ��
Private mblnMemory As Boolean           '���Ի�
Private mblnNeedAudit As Boolean        'True�����󷽲�����Falseδ�����󷽲���
Private mblnAuditStart As Boolean       'True����������Falseֹͣ������
Private mblnLocking As Boolean          'True������Falseδ����
Private mblnSendBeforeAudit As Boolean
Private mblnExit As Boolean
Private mblnEnter As Boolean
Private mstrPCName As String
Private mblnReadCard As Boolean
Private mbytFontSize As Byte
Private mintParaPass As Integer         '������ҩ����
Private msngY As Single
Private mbytDrugName As Byte            '0-��ʾͨ������1-��ʾ��Ʒ����2-ͬʱ��ʾͨ��������Ʒ��
Private mintItem As Integer             '1-���ݡ�������������淶��28�2-���ݡ���������취��7��
Private mlngAuditID As Long             '��ID
Private mblnReasonRefresh As Boolean    '��ֹ�������txtReasonʱ����Change�¼�
Private mblnRecipeSendAuto As Boolean   '�����Զ�����

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
        Case enuMenus.��ӡ����
            Call zlPrintSet
        Case enuMenus.��ӡԤ��, enuMenus.��ӡ, enuMenus.���Excel
            Dim objTmp As Object
            Dim strTitle As String
            
            Set objTmp = Me.ActiveControl
            If TypeName(objTmp) = "VSFlexGrid" Then
                objTmp.Redraw = False
                If UCase(objTmp.Name) = "VSFRECWAIT" Then
                    strTitle = "���ﴦ��������¼"
                ElseIf UCase(objTmp.Name) = "VSFRECFINISH" Then
                    strTitle = "���ﴦ��������¼"
                ElseIf UCase(objTmp.Name) = "VSFRECIPE" Then
                    strTitle = "���ﴦ�����ҩ����ϸ"
                ElseIf UCase(objTmp.Name) = "VSFAUDIT" Then
                    strTitle = "���ﴦ�������Ŀ��ϸ"
                End If
                If strTitle <> "" Then
                    If Control.ID = enuMenus.��ӡԤ�� Then
                        zlRptPrint 0, objTmp, strTitle
                    ElseIf Control.ID = enuMenus.��ӡ Then
                        zlRptPrint 1, objTmp, strTitle
                    Else
                        zlRptPrint 3, objTmp, strTitle
                    End If
                End If
                objTmp.Redraw = True
            End If
        Case enuMenus.��������
            timPatient.Enabled = False      '�رն�ʱ�ؼ�
            frmRAParams.ShowMe Me, 0
            Call SetPatientTimer            '����ȷ���Ƿ�����ʱ
            Call SetClinicItem
        Case enuMenus.�˳�
            If mblnAuditStart Then
                If AuditOperate(0) Then
                    Unload Me
                End If
            Else
                Unload Me
            End If
        Case enuMenus.�������
            If AuditOperate(IIf(mblnAuditStart, 0, 1)) Then
                mblnAuditStart = True
            End If
            Call RefreshLockControls
'            Call FillVSFData(1)
'            Call FillVSFData(2)
'            Call FillVSFData(3)
        Case enuMenus.ֹͣ���
            If AuditOperate(IIf(mblnAuditStart, 0, 1)) Then
                mblnAuditStart = False
            End If
            Call RefreshLockControls
        Case enuMenus.�ϸ�
            If mblnLocking = False Then
                lngAuditID = Val(vsfRecWait.TextMatrix(vsfRecWait.Row, vsfRecWait.ColIndex("��ID")))
                '����
                Call AuditLock(lngAuditID)
            End If
            Call AuditProcess(1)
        Case enuMenus.���ϸ�
            If mblnLocking = False Then
                lngAuditID = Val(vsfRecWait.TextMatrix(vsfRecWait.Row, vsfRecWait.ColIndex("��ID")))
                '����
                Call AuditLock(lngAuditID)
            End If
            Call AuditProcess(2)
        Case enuMenus.�鿴PASS���
            Call PassResultView(mobjPASS, True, Val(vsfRecipe.TextMatrix(vsfRecipe.Row, vsfRecipe.ColIndex("ҽ��ID"))))
        Case enuMenus.ˢ��
            If mblnLocking Then
                If MsgBox("������鵱ǰ���˵�ҩ�����Ƿ������飿", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
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
        Case enuMenus.��׼��ť
            For i = 2 To cbsMain.Count
                Me.cbsMain(i).Visible = Not Me.cbsMain(i).Visible
            Next
            Me.cbsMain.RecalcLayout
        Case enuMenus.�ı���ǩ
            For i = 2 To cbsMain.Count
                For Each objControl In Me.cbsMain(i).Controls
                    objControl.Style = IIf(objControl.Style = xtpButtonIcon, xtpButtonIconAndCaption, xtpButtonIcon)
                Next
            Next
            Me.cbsMain.RecalcLayout
        Case enuMenus.��ͼ��
            cbsMain.Options.LargeIcons = Not cbsMain.Options.LargeIcons
            cbsMain.RecalcLayout
        Case enuMenus.С����
            If mbytFontSize <> 0 Then Call SetControlFontSize(0)
        Case enuMenus.������
            If mbytFontSize <> 1 Then Call SetControlFontSize(1)
        Case enuMenus.״̬��
            stbThis.Visible = Not Control.Checked
            cbsMain.RecalcLayout
        Case enuMenus.��������
            Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
        Case enuMenus.������ҳ
            Call zlHomePage(Me.hwnd)
        Case enuMenus.������̳
            Call zlWebForum(Me.hwnd)
        Case enuMenus.���ͷ���
            Call zlMailTo(Me.hwnd)
        Case enuMenus.����
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case Else
            '����
            If Between(Control.ID, enuMenus.���� * 100# + 1, enuMenus.���� * 100# + 99) And Control.Parameter <> "" Then
                Call ReportOpen(gcnOracle, Split(Control.Parameter, ",")(0), Split(Control.Parameter, ",")(1), Me, "��ID=" & mlngAuditID)
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
        Case enuMenus.��ӡ����, enuMenus.��ӡԤ��, enuMenus.��ӡ, enuMenus.���Excel, enuMenus.��������
            Control.Enabled = Not mblnLocking
        Case enuMenus.�������
            Control.Enabled = mblnNeedAudit
            If mblnNeedAudit = False Then Exit Sub
            Control.Enabled = Not mblnAuditStart
        Case enuMenus.ֹͣ���
            Control.Enabled = mblnNeedAudit
            If mblnNeedAudit = False Then Exit Sub
            Control.Enabled = mblnAuditStart
        Case enuMenus.�ϸ�
            Control.Enabled = cmdYes_Fixed.Enabled
        Case enuMenus.���ϸ�
            Control.Enabled = cmdNo_Fixed.Enabled
        Case enuMenus.��׼��ť
            Control.Checked = Me.cbsMain(2).Visible
        Case enuMenus.�ı���ǩ
            Control.Checked = (Me.cbsMain(2).Controls(1).Style = xtpButtonCaption Or Me.cbsMain(2).Controls(1).Style = xtpButtonIconAndCaption)
        Case enuMenus.��ͼ��
            Control.Checked = cbsMain.Options.LargeIcons
        Case enuMenus.С����
            Control.Checked = mbytFontSize = 0
            Control.Enabled = Not mblnLocking
        Case enuMenus.������
            Control.Checked = mbytFontSize = 1
            Control.Enabled = Not mblnLocking
        Case enuMenus.�鿴PASS���
            Control.Enabled = (mintParaPass > 0 And mintParaPass < 5) And vsfRecipe.Rows > 1
        Case enuMenus.״̬��
            Control.Checked = Me.stbThis.Visible
        Case Else
            '����
            If Between(Control.ID, enuMenus.���� * 100# + 1, enuMenus.���� * 100# + 99) And Control.Parameter <> "" Then
                Control.Enabled = Not mblnLocking
            End If
    End Select
End Sub

Private Sub chk24_Click()
    Call FillVSFData(2)
End Sub

Private Sub chkPubPatient_Click()
    Dim objControl As XtremeCommandBars.CommandBarButton
    
    Set objControl = cbsMain.ActiveMenuBar.Controls.Find(, enuMenus.ˢ��, , True)
    If Not objControl Is Nothing Then Call cbsMain_Execute(objControl)
End Sub

Private Sub cmdNo_Fixed_Click()
    Dim objControl As XtremeCommandBars.CommandBarButton
    
    Set objControl = cbsMain.ActiveMenuBar.FindControl(, enuMenus.���ϸ�, , True)
    If Not objControl Is Nothing Then
        If objControl.Enabled Then
            Call cbsMain_Execute(objControl)
        End If
    End If
End Sub

Private Sub cmdYes_Fixed_Click()
    Dim objControl As XtremeCommandBars.CommandBarButton
    
    Set objControl = cbsMain.ActiveMenuBar.FindControl(, enuMenus.�ϸ�, , True)
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
    
    '��Դ����
'    '������ﴦ����������Ƿ�������Դ����
'    strSQL = "Select a.Id ����id, a.����, a.���� " & vbNewLine & _
'             "From ���ű� A, ����������� B " & vbNewLine & _
'             "Where a.Id = b.����id And (a.����ʱ�� Is Null Or To_Char(a.����ʱ��, 'yyyy') = '3000') And b.��� = 1 "
'    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���ﴦ�������������Դ����")
'    If rsTemp.RecordCount <= 0 Then
'        MsgBox "���������������δ������Դ���ң�", vbInformation, gstrSysName
'        rsTemp.Close
''        mblnExit = True
''        Exit Sub
'    End If
    
    '��ȡ��λ���õ���Դ����
    Call SetClinicItem
    
    '��ҩҩ��
    strSQL = "Select Distinct a.����id, c.����, c.���� " & vbNewLine & _
             "From ������Ա A, ��������˵�� B, ���ű� C " & vbNewLine & _
             "Where a.����id = b.����id And a.����id = c.Id And a.��Աid = [1] And b.�������� In ('��ҩ��', '��ҩ��', '��ҩ��') " & vbNewLine & _
             "   And b.������� In (1, 3) And (c.����ʱ�� Is Null Or To_Char(c.����ʱ��, 'yyyy') = '3000') " & vbNewLine & _
             "Order By c.���� "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����Ա������������ҩ��", UserInfo.ID)
    If rsTemp.RecordCount <= 0 Then
        MsgBox "����Ա��������ҩ�����ŵ����ԣ����鲿�Ź���", vbInformation, gstrSysName
        rsTemp.Close
        mblnExit = True
        Exit Sub
    End If
    
    With rsTemp
        cboDrugstore.Tag = ""
        cboDrugstore.Clear
        cboDrugstore.AddItem "���з�ҩҩ��"
        Do While .EOF = False
            cboDrugstore.AddItem !����
            cboDrugstore.ItemData(cboDrugstore.NewIndex) = !����ID
            cboDrugstore.Tag = IIf(cboDrugstore.Tag = "", "", cboDrugstore.Tag & ",") & !����ID
            .MoveNext
        Loop
        .Close
        cboDrugstore.ListIndex = 0
    End With

    '��ʼ��������ģ�����
    mblnSendBeforeAudit = Val(zlDatabase.GetPara("������ʱ��", glngSys)) = 1
    mintParaPass = Val(zlDatabase.GetPara("������ҩ���ӿ�", glngSys))      '0-��ʾδʹ��,1-�����ӿ�,2-��ͨ�ӿڣ��ݲ�֧�֣�,3-̫Ԫͨ�ӿ�,4-����
    mbytDrugName = Val(zlDatabase.GetPara("ҩƷ������ʾ"))
    mintItem = Val(zlDatabase.GetPara("�����������", glngSys, , "1"))
    mblnRecipeSendAuto = Val(zlDatabase.GetPara("�����󷽴����Զ�����", glngSys)) = 1
    
    lngTmp = Val(zlDatabase.GetPara("�������", glngSys))
    mblnNeedAudit = (lngTmp = 1 Or lngTmp = 3)                              '1-����������飻3-�����סԺ�������
    mblnAuditStart = False                                                  '���봰�����״̬Ϊ��ֹ
    mblnLocking = False
    mlngModule = glngModule
    mstrPrivs = zlStr.FormatString(";[1];", GetPrivFunc(glngSys, mlngModule))
    mblnMemory = Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 1
    mbytFontSize = 0
    
    '������Ϣƽ̨����
    On Error Resume Next
    Set mobjMipModule = New zl9ComLib.clsMipModule
    If Not mobjMipModule Is Nothing Then
        mobjMipModule.InitMessage glngSys, mlngModule, mstrPrivs
        zl9ComLib.AddMipModule mobjMipModule
        If Err.Number <> 0 Then Set mobjMipModule = Nothing
    End If
    Err.Clear: On Error GoTo errHandle
    
'    '���ص��Ӳ�������
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
    
    '�ٴ��Ĺ�������
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
    
    '������ҩ
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
    
    '��ʼ�����沼��
    Call InitCommandbars
    Call InitDockPane
    Call InitTBCRec
    Call InitTBCTab
    
    '��ʼ���ؼ�
    Err.Clear: On Error Resume Next
    Call iknPatient.zlInit(Me, glngSys, mlngModule, gcnOracle, UserInfo.�û���, , MSTR_PATIENT, txtPatient)
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
        vsfRecipe.ColHidden(vsfRecipe.ColIndex("ҩƷ����")) = False
        vsfRecipe.ColHidden(vsfRecipe.ColIndex("��Ʒ��")) = True
    ElseIf mbytDrugName = 1 Then
        vsfRecipe.ColHidden(vsfRecipe.ColIndex("ҩƷ����")) = True
        vsfRecipe.ColHidden(vsfRecipe.ColIndex("��Ʒ��")) = False
    Else
        vsfRecipe.ColHidden(vsfRecipe.ColIndex("ҩƷ����")) = False
        vsfRecipe.ColHidden(vsfRecipe.ColIndex("��Ʒ��")) = False
    End If
    
    '����PASS�����ͼƬ���ƣ���  0-��ʾδʹ��,1-�����ӿ�,2-��ͨ�ӿ�,3-̫Ԫͨ�ӿ�,4-����
    vsfRecipe.ColHidden(vsfRecipe.ColIndex("PASS")) = (mintParaPass < 1 Or mintParaPass > 4)
    
    With vsfAudit
        .Editable = flexEDKbdMouse
        .ColComboList(vsfAudit.ColIndex("ҩƷ")) = "..."
    End With
    
    Call SetFilterDay(0)
    
    '�ָ��ϴν���
    RestoreWinState Me, App.ProductName
    If mblnMemory Then
        Dim strPane As String
        Dim objControl As XtremeCommandBars.CommandBarButton
        
        '�����С
        lngTmp = Val(GetSetting("ZLSOFT", zlStr.FormatString("˽��ģ��\[1]\��������\[2]\[3]\Form", "ZLHIS", App.ProductName, Me.Name), "FontSize"))
        Set objControl = cbsMain.FindControl(, IIf(lngTmp = 1, enuMenus.������, enuMenus.С����), , True)
        If Not objControl Is Nothing Then
            Call cbsMain_Execute(objControl)
        End If
        
        'DockingPane
        strPane = GetSetting("ZLSOFT", zlStr.FormatString("˽��ģ��\[1]\��������\[2]\[3]\Form", "ZLHIS", App.ProductName, Me.Name), "����")
        iknPatient.IDKind = Val(GetSetting("ZLSOFT", zlStr.FormatString("˽��ģ��\[1]\��������\[2]\[3]\Form", "ZLHIS", App.ProductName, Me.Name), "iknPatient"))
        dkpMain.LoadStateFromString strPane
    End If
    
    '�����Լ��������󷽼�¼
    Call AuditLock(0, False)
    
    'ˢ������
    Call FillVSFData(1)
    Call FillVSFData(2)
    'Call FillVSFData(3)
    
    sbnPASS.Visible = (mintParaPass > 0 And mintParaPass < 5) And Not mobjPASS Is Nothing    '�鿴PASS���
    
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
    
    cbsMain.VisualTheme = xtpThemeOffice2003 'xtpthemeoffice2000�а�͹��
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
        .ActiveMenuBar.Title = "�˵�"
        .ActiveMenuBar.EnableDocking xtpFlagHideWrap Or xtpFlagStretched
    End With
    
    picLine01_S.BackColor = cbsMain.GetSpecialColor(STDCOLOR_BTNFACE)
    picLine02.BackColor = picLine01_S.BackColor
    
    '�ļ�
    Set cbpTmp = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enuMenus.�ļ�, "�ļ�(&F)", -1, False)
    With cbpTmp
        .ID = enuMenus.�ļ�
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.��ӡ����, "��ӡ����(&S)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.��ӡԤ��, "��ӡԤ��(&V)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.��ӡ, "��ӡ")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.���Excel, "�����&Excel...")
        cbcTmp.BeginGroup = True
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.��������, "��������")
        cbcTmp.BeginGroup = True
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.�˳�, "�˳�")
    End With
    
    '�༭
    Set cbpTmp = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enuMenus.�༭, "�༭(&E)", -1, False)
    With cbpTmp
        .ID = enuMenus.�༭
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.�������, "�������(&S)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.ֹͣ���, "ֹͣ���(&P)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.�ϸ�, "�ϸ�(&Y)")
        cbcTmp.BeginGroup = True
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.���ϸ�, "���ϸ�(&N)")
    End With
    
    '�鿴
    Set cbpTmp = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enuMenus.�鿴, "�鿴(&V)", -1, False)
    With cbpTmp
        .ID = enuMenus.�鿴
        Set cbpTmp = .CommandBar.Controls.Add(xtpControlPopup, enuMenus.������, "������(&T)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.��׼��ť, "��׼��ť(&S)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.�ı���ǩ, "�ı���ǩ(&T)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.��ͼ��, "��ͼ��(&B)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.״̬��, "״̬��(&S)")
        cbcTmp.BeginGroup = True
        Set cbpTmp = .CommandBar.Controls.Add(xtpControlPopup, enuMenus.�����С, "�����С(&F)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.С����, "С����(&S)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.������, "������(&B)")
        cbcTmp.BeginGroup = True
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.�鿴PASS���, "�鿴&PASS���")
        cbcTmp.BeginGroup = True
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.ˢ��, "ˢ��")
        cbcTmp.BeginGroup = True
    End With
    
    '����
    Set cbpTmp = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, enuMenus.����, "����(&H)", -1, False)
    With cbpTmp
        .ID = enuMenus.����
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.��������, "��������")
        Set cbpTmp = .CommandBar.Controls.Add(xtpControlPopup, enuMenus.WEB�ϵ�����, "&WEB�ϵ�����")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.������ҳ, "������ҳ(&H)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.������̳, "������̳(&F)")
            Set cbcTmp = cbpTmp.CommandBar.Controls.Add(xtpControlButton, enuMenus.���ͷ���, "���ͷ���(&K)")
        Set cbcTmp = .CommandBar.Controls.Add(xtpControlButton, enuMenus.����, "����(&A)")
        cbcTmp.BeginGroup = True
    End With
    
    '����ӿ�
    Call zlDatabase.ShowReportMenu(Me, glngSys, mlngModule, mstrPrivs)
    
    '�˵���Ŀ����
    With cbsMain.KeyBindings
        .Add 8, vbKeyP, enuMenus.��ӡ
        .Add 8, vbKeyX, enuMenus.�˳�
        .Add 0, vbKeyF12, enuMenus.��������
        .Add 0, vbKeyF5, enuMenus.ˢ��
        .Add 0, vbKeyF1, enuMenus.��������
    End With
    
    '���幤����
    Set cbrTmp = cbsMain.Add("������", xtpBarTop)
    With cbrTmp
        .ShowTextBelowIcons = False
        .EnableDocking xtpFlagStretched Or xtpFlagHideWrap
        
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.��ӡԤ��, "��ӡԤ��")
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.��ӡ, "��ӡ")
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.�������, "�������")
        cbcTmp.BeginGroup = True
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.ֹͣ���, "ֹͣ���")
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.ˢ��, "ˢ��")
        cbcTmp.BeginGroup = True
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.��������, "��������")
        cbcTmp.BeginGroup = True
        Set cbcTmp = .Controls.Add(xtpControlButton, enuMenus.�˳�, "�˳�")
    End With
    
    '��ͼ�꣬���ı��İ�ť���
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
            .Title = "��������"
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
        .InsertItem 0, "����(&A)", picRecWait.hwnd, 0
        .InsertItem 1, "����(&B)", picRecFinish.hwnd, 0
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
        .InsertItem 0, "�������(&1)", picDetail.hwnd, 0
        If Not mfrmAMR Is Nothing Then
            .InsertItem 1, "ҽ���ͱ���(&2)", mfrmAMR.hwnd, 0
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
        SaveSetting "ZLSOFT", zlStr.FormatString("˽��ģ��\[1]\��������\[2]\[3]\Form", "ZLHIS", App.ProductName, Me.Name), "����", strPane
        SaveSetting "ZLSOFT", zlStr.FormatString("˽��ģ��\[1]\��������\[2]\[3]\Form", "ZLHIS", App.ProductName, Me.Name), "iknPatient", iknPatient.IDKind
        SaveSetting "ZLSOFT", zlStr.FormatString("˽��ģ��\[1]\��������\[2]\[3]\Form", "ZLHIS", App.ProductName, Me.Name), "FontSize", mbytFontSize
    End If
End Sub

Private Sub iknPatient_ItemClick(Index As Integer, objCard As zlIDKind.Card)
    If txtPatient.Enabled And txtPatient.Visible Then
        txtPatient.Text = ""
        txtPatient.SetFocus
    End If
End Sub

Private Sub iknPatient_ReadCard(ByVal objCard As zlIDKind.Card, objPatiInfor As zlIDKind.PatiInfor, blnCancel As Boolean)
    txtPatient.Text = objPatiInfor.����
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
        MsgBox "���ۺ����ɡ����ݳ���500�ַ���250���֣�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    On Error GoTo errHandle
    
    strSQL = zlStr.FormatString("Zl_������鳣������_Update(1, '[1]', '[2]')", _
                    UserInfo.�û���, _
                    txtReason.Text)
    Call zlDatabase.ExecuteProcedure(strSQL, "����������鳣������")
    
    MsgBox "�ղ���ɣ�", vbInformation, gstrSysName
    
    Exit Sub
    
errHandle:
    If ErrCenter = 1 Then Resume
End Sub

Private Sub sbnPASS_Click()
    Dim objControl As XtremeCommandBars.CommandBarButton
    
    Set objControl = cbsMain.ActiveMenuBar.Controls.Find(, enuMenus.�鿴PASS���, , True)
    If Not objControl Is Nothing Then Call cbsMain_Execute(objControl)
End Sub

Private Sub sbnSelect_Click()
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim blnCancel As Boolean
    Dim i As Integer
    
    If mblnNeedAudit = False Or mblnAuditStart = False Then Exit Sub
    
    On Error GoTo errHandle
    
    strSQL = "Select Null ѡ��, ���� From ������鳣������ Where �û��� = [1] "
    
    '����ѡ�������ɶ�ѡ
    Set rsTemp = mdlRecipeAudit.ShowReason(Me, strSQL, blnCancel, UserInfo.�û���)
    
    If blnCancel = False Then
        With rsTemp
            strSQL = ""
            i = 1
            If .RecordCount > 0 Then .MoveFirst
            Do While .EOF = False
                strSQL = strSQL & zlStr.FormatString("[1]��[2][3]", i, !����, vbNewLine)
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
        '����
        picFunc.Visible = True
        picYesNO.Visible = False
        chk24.Visible = True
    Else
        '����
        picFunc.Visible = False
        picYesNO.Visible = True
        chk24.Visible = False
    End If
    
    SetFilterDay Item.Index
    
    If Me.Visible = False Then Exit Sub
    
    Set objControl = cbsMain.ActiveMenuBar.Controls.Find(, enuMenus.ˢ��, , True)
    If Not objControl Is Nothing Then
        Call cbsMain_Execute(objControl)
'        If vsfRecFinish.ColKey(0) <> "" Then Call vsfRecFinish_AfterRowColChange(0, 0, vsfRecFinish.Row, 1)
    End If

End Sub

Private Sub InitVSF(ByRef vsfVar As VSFlexGrid)
'���ܣ���ʼ�������VSFlexGrid�ؼ��ķ��
'������
'  vsfVar��Ҫ��ʼ����VSFlexGrid�ؼ�

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
'���ܣ����VSF�ؼ�������
'������
'  bytMode��1-����������ݣ�2-������ϸ���ݣ�3-�����Ŀ����

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
        Case 1      '1-�����������
            '��Դ����
            If cboClinic.ListCount > 0 Then
                lngClinicID = cboClinic.ItemData(cboClinic.ListIndex)
            Else
                'δ������Դ����
                lngClinicID = -1
            End If
            
            '��ҩҩ��
            lngDrugstoreID = cboDrugstore.ItemData(cboDrugstore.ListIndex)
            
            '�ύʱ��
            datEnd = Format(zlDatabase.Currentdate, "yyyy-MM-dd 23:59:59")
            datBegin = Format(datEnd - intDay, "yyyy-MM-dd 00:00:00")
            
            '������Ϣ
            If Trim(txtPatient.Text) <> "" Then
                Select Case iknPatient.GetCurCard.����
                    Case "����"
                        strPatient = " And Upper(b.����) = [5] "
                    Case "�����"
                        strPatient = " And b.����� = [5] "
                    Case "���֤��"
                        strPatient = " And b.���֤�� = [5] "
                    Case Else
                        Call iknPatient.zlFindPatient(Me.txtPatient.Text, objPatiInfo)
                        If Not objPatiInfo Is Nothing Then
                            lngPatientID = objPatiInfo.����ID
                        End If
                        strPatient = " And b.����ID = [5] "
                        blnReadID = True
                End Select
            End If
            
            If lngClinicID < 0 Then
                'δ������Դ����
                strTmp = "Select -1 ��Դ���� From Dual"
            ElseIf lngClinicID = 0 Then
                '���б���ѡ����Դ����
                strTmp = "Select Column_Value ��Դ���� " & vbNewLine & _
                         "From Table(f_Num2list((Select ��Դ���� From ���������� Where Upper(������) = [6] And ������� = 0), ',')) "
'                strTmp = "Select a.Id ��Դ����  " & vbNewLine & _
'                         "From ���ű� A, ����������� B, " & vbNewLine & _
'                         "    Table(f_Num2list((Select ��Դ���� From ���������� Where ������ = [6] And ������� = 0))) C " & vbNewLine & _
'                         "Where a.Id = b.����id And a.Id = c.Column_Value And (a.����ʱ�� Is Null Or To_Char(a.����ʱ��, 'yyyy') = '3000') " & vbNewLine & _
'                         "    And b.��� = 1 And (b.����id Is Not Null Or b.����id > 0)"
            End If
            
            If Me.tbcRec.Item(0).Selected Then
                '�����
                If chkPubPatient.Value = 1 And lngClinicID <= 0 Then
                    '��λδ���õ���Դ����
                    strTmp = strTmp & _
                             "Union All Select ����ID From (" & vbNewLine & _
                             "Select ����id From ����������� Where ��� = 1 " & vbNewLine & _
                             "Minus " & vbNewLine & _
                             "Select Distinct Column_Value ��Դ���� " & vbNewLine & _
                             "From Table(f_Num2list((Select f_List2str(Cast(Collect(��Դ����) As t_Strlist), ',') ��Դ���� " & vbNewLine & _
                             "                       From ���������� " & vbNewLine & _
                             "                       Where ������� = 0 And ������ʱ�� >= Sysdate - 60 And ��Դ���� Is Not Null), ',')) ) "
                End If
                
                strSQL = "Select b.���� ����, b.����id, b.��ҳid, b.�Ա�, b.����, b.�����, c.Id ��id, c.�Һ�id �Һŵ�Id, " & vbNewLine & _
                         "    To_Char(c.�ύʱ��, 'yyyy-mm-dd hh24:mi') �ύʱ��, c.�ύ��, D1.���� �ύ����, D2.���� ��ҩҩ�� " & vbNewLine & _
                         "From " & IIf(lngClinicID <= 0, zlStr.FormatString("(Select ��Դ���� From ([1])) A,", strTmp), "") & vbNewLine & _
                         "    ������Ϣ B, ��������¼ C, ���ű� D1, ���ű� D2 " & vbNewLine & _
                         IIf(lngDrugstoreID <= 0, ", Table(f_Num2list([4], ',')) E ", "") & vbNewLine & _
                         "Where b.����id = c.����id And c.�ύ����id = D1.Id And c.��ҩҩ��id = D2.Id " & vbNewLine & _
                         IIf(lngClinicID <= 0, " And a.��Դ���� = c.�ύ����id ", "") & vbNewLine & _
                         IIf(lngDrugstoreID <= 0, " And c.��ҩҩ��id = e.Column_Value ", " And c.��ҩҩ��id = [4] ") & vbNewLine & _
                         "    And c.���ʱ�� Is Null And c.״̬ = 0 And c.�ύʱ�� Between [1] And [2] And c.�Һ�ID Is Not Null "
                strSQL = strSQL & IIf(lngClinicID <= 0, "", " And c.�ύ����id = [3] ")
                strSQL = strSQL & IIf(Trim(txtPatient.Text) = "", "", strPatient)
                strSQL = strSQL & vbNewLine & "Order By c.�ύʱ�� "
                
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ��������������", _
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
                '�����
                strSQL = "Select b.���� ����, b.����id, b.��ҳid, b.�Ա�, b.����, b.�����, c.Id ��id, c.�Һ�id �Һŵ�Id, " & vbNewLine & _
                         "    To_Char(c.�ύʱ��, 'yyyy-mm-dd hh24:mi') �ύʱ��, c.�ύ��, Decode(c.�����, 1, '�ϸ�', '���ϸ�') �����, " & vbNewLine & _
                         "    To_Char(c.���ʱ��, 'yyyy-mm-dd hh24:mi') ���ʱ��, c.�����, D1.���� �ύ����, D2.���� ��ҩҩ�� " & vbNewLine & _
                         "From " & IIf(lngClinicID <= 0, zlStr.FormatString("(Select ��Դ���� From ([1])) A,", strTmp), "") & vbNewLine & _
                         "    ������Ϣ B, ��������¼ C, ���ű� D1, ���ű� D2 " & vbNewLine & _
                         IIf(lngDrugstoreID <= 0, ", Table(f_Num2list([4], ',')) E ", "") & vbNewLine & _
                         "Where b.����id = c.����id And c.�ύ����id = D1.Id And c.��ҩҩ��id = D2.Id " & vbNewLine & _
                         IIf(lngClinicID <= 0, " And a.��Դ���� = c.�ύ����id ", "") & vbNewLine & _
                         IIf(lngDrugstoreID <= 0, " And c.��ҩҩ��id = e.Column_Value ", " And c.��ҩҩ��id = [4] ") & vbNewLine & _
                         "    And c.���ʱ�� Is Not Null And c.״̬ = 1 And c.�ύʱ�� Between [1] And [2] And c.�Һ�ID Is Not Null "
                strSQL = strSQL & IIf(lngClinicID <= 0, "", " And c.�ύ����id = [3] ")
                strSQL = strSQL & IIf(Trim(txtPatient) = "", "", strPatient)
                strSQL = strSQL & vbNewLine & "Order By c.�ύʱ�� "
                
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ���������������", _
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
            
        Case 2      '2-������ϸ����
            
            If Me.tbcRec.Item(0).Selected Then
                Set vsfTmp = vsfRecWait
            Else
                Set vsfTmp = vsfRecFinish
            End If
            
            vsfRecipe.Rows = 1
            vsfRecipe.Clear 1
            
            If vsfTmp.Rows <= 1 Then Exit Sub
            
            lngRAID = Val(vsfTmp.TextMatrix(vsfTmp.Row, vsfTmp.ColIndex("��ID")))
            lngPatientID = Val(vsfTmp.TextMatrix(vsfTmp.Row, vsfTmp.ColIndex("����ID")))
            
            strSQL = "Select Null ��־, e.���� �ٴ�����, b.��id, b.ҽ��id, A1.����˵��, A1.ҽ������, A1.���id, A1.����� PASS, " & vbNewLine & _
                     "    d.���� ҩƷ����, f.���� ��Ʒ��, d.���, d.���㵥λ ��λ, A1.�ܸ����� ����, A1.�������� || Nvl(g.���㵥λ, '') ����, A2.ҽ������ �÷�, " & vbNewLine & _
                     "    A1.ִ��Ƶ�� Ƶ��, c.��׼���� ����, c.Ӧ�ս��, d.ID ϸĿID, " & vbNewLine & _
                     "    Row_Number() Over(Partition By A1.���id Order By b.ҽ��id) �����, Count(1) Over(Partition By A1.���id) ���� " & vbNewLine & _
                     "From ����ҽ����¼ A1, ����ҽ����¼ A2, ���������ϸ B, (" & vbNewLine & _
                     "       Select a.ҽ�����, Max(a.��׼����) ��׼����, Sum(a.Ӧ�ս��) Ӧ�ս�� " & vbNewLine & _
                     "       From ������ü�¼ A, ���������ϸ B " & vbNewLine & _
                     "       Where a.ҽ����� = b.ҽ��id And b.��id = [1] " & vbNewLine & _
                     "       Group By a.ҽ����� " & vbNewLine & _
                     "    ) C, �շ���ĿĿ¼ D, ���ű� E, �շ���Ŀ���� F, ������ĿĿ¼ G " & vbNewLine & _
                     "Where A1.���id = A2.Id And A1.Id = b.ҽ��id And A1.Id = c.ҽ�����(+) And A1.�շ�ϸĿid = d.Id And A1.��������id = e.Id And " & vbNewLine & _
                     "    A1.�շ�ϸĿid = f.�շ�ϸĿid(+) And b.��id = [1] And f.����(+) = 3 And f.����(+) = 1 And A1.������ĿID = g.ID(+) " & vbNewLine & _
                     "Order By A1.���id, ����� "
            
            If chk24.Value = 1 And Me.chk24.Visible = True Then
                '24Сʱ�ڵ�����ҩ�����в���ҽ�����ͼ�¼��
'                strTmp = "Select e.���� �ٴ�����, Null ��id, A1.Id ҽ��id, A1.����˵��, A1.ҽ������, A1.���id, A1.�����, d.���� ҩƷ����, " & vbNewLine & _
'                         "    f.���� ��Ʒ��, d.���, d.���㵥λ ��λ, A1.�ܸ����� ����, A1.�������� ����, A2.ҽ������ �÷�, A1.ִ��Ƶ�� Ƶ��, " & vbNewLine & _
'                         "    c.��׼���� ����, c.Ӧ�ս��, d.ID ϸĿID, Row_Number() Over(Partition By A1.���id Order By A1.Id) �����,  " & vbNewLine & _
'                         "    Count(1) Over(Partition By A1.���id) ���� " & vbNewLine & _
'                         "From ����ҽ����¼ A1, ����ҽ����¼ A2, ����ҽ������ B, ������ü�¼ C, �շ���ĿĿ¼ D, ���ű� E, �շ���Ŀ���� F " & vbNewLine & _
'                         "Where A1.���id = A2.Id And A1.Id = b.ҽ��id And A1.Id = c.ҽ�����(+) And A1.�շ�ϸĿid = d.Id And A1.��������id = e.Id " & vbNewLine & _
'                         "    And A1.�շ�ϸĿid = f.�շ�ϸĿid(+) And f.����(+) = 3 And f.����(+) = 1 And b.����ʱ�� Between Sysdate - 1 And Sysdate " & vbNewLine & _
'                         "    And b.��ID <> [1] And A1.����ID = [2] And A1.������Դ = 1 And Not A1.Id In (Select ҽ��id From W_A) "
                
                strTmp = "Select 1 ��־, e.���� �ٴ�����, b.��id, b.ҽ��id, A1.����˵��, A1.ҽ������, A1.���id, A1.����� PASS, " & vbNewLine & _
                         "    d.���� ҩƷ����, f.���� ��Ʒ��, d.���, d.���㵥λ ��λ, A1.�ܸ����� ����, A1.�������� || Nvl(g.���㵥λ, '') ����, A2.ҽ������ �÷�, " & vbNewLine & _
                         "    A1.ִ��Ƶ�� Ƶ��, c.��׼���� ����, c.Ӧ�ս��, d.ID ϸĿID, " & vbNewLine & _
                         "    Row_Number() Over(Partition By A1.���id Order By b.ҽ��id) �����, Count(1) Over(Partition By A1.���id) ���� " & vbNewLine & _
                         "From ����ҽ����¼ A1, ����ҽ����¼ A2, ����ҽ������ A3, ���������ϸ B, (" & vbNewLine & _
                         "    Select a.ҽ�����, Max(a.��׼����) ��׼����, Sum(a.Ӧ�ս��) Ӧ�ս�� " & vbNewLine & _
                         "    From ������ü�¼ A, ����ҽ����¼ B " & vbNewLine & _
                         "    Where a.ҽ����� = b.id And a.�շ���� in ('5','6','7') " & vbNewLine & _
                         "        And b.����id = [2] And b.������Դ = 1 And a.�Ǽ�ʱ�� Between Sysdate - 1 And Sysdate " & vbNewLine & _
                         "    Group By a.ҽ����� " & vbNewLine & _
                         "    ) C, �շ���ĿĿ¼ D, ���ű� E, �շ���Ŀ���� F, ������ĿĿ¼ G " & vbNewLine & _
                         "Where A1.���id = A2.Id And A1.ID = A3.ҽ��ID And A1.Id = b.ҽ��id(+) And A1.Id = c.ҽ�����(+) " & vbNewLine & _
                         "    And A1.�շ�ϸĿid = d.Id And A1.��������id = e.Id And A1.�շ�ϸĿid = f.�շ�ϸĿid(+) And A1.������ĿID = g.ID(+) " & vbNewLine & _
                         "    And b.��id(+) <> [1] And f.����(+) = 3 And f.����(+) = 1 And A1.����id = [2] And A1.������Դ = 1 " & vbNewLine & _
                         "    And Not A1.Id In (Select ҽ��id From W_A) And A3.����ʱ�� Between Sysdate - 1 And Sysdate " & vbNewLine & _
                         "Order By A1.���id, ����� "
                         
                strSQL = "Select * From (With W_A As (" & strSQL & ") " & vbNewLine & _
                         "  Select * From W_A " & vbNewLine & _
                         "  Union All " & vbNewLine & _
                         "  Select * From (" & strTmp & ") ) "
            End If
            
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�󷽼�¼��ҽ����Ϣ", lngRAID, lngPatientID)
            Call mdlDefine.FillVSFData(vsfRecipe, rsTemp)
            
            If vsfRecipe.Rows > 1 Then
                With vsfRecipe
'                    '�÷���Ƶ�κϲ���Ԫ��Ҫ����ϲ�������������
'                    .MergeCells = flexMergeRestrictAll
'                    .MergeCol(.ColIndex("���ID")) = True
'                    .MergeCol(.ColIndex("�÷�")) = True
'                    .MergeCol(.ColIndex("Ƶ��")) = True
                        
                    For l = 1 To .Rows - 1
                        '����ҽ������
                        If Val(.TextMatrix(l, .ColIndex("����"))) >= 3 Then
                            If Val(.TextMatrix(l, .ColIndex("�����"))) = 1 Then
                                '����
                                .TextMatrix(l, .ColIndex("��")) = "��"
                            ElseIf Val(.TextMatrix(l, .ColIndex("�����"))) = Val(.TextMatrix(l, .ColIndex("����"))) Then
                                '��β
                                .TextMatrix(l, .ColIndex("��")) = "��"
                            Else
                                '����
                                .TextMatrix(l, .ColIndex("��")) = "��"
                            End If
                        ElseIf Val(.TextMatrix(l, .ColIndex("����"))) = 2 Then
                            If Val(.TextMatrix(l, .ColIndex("�����"))) = 1 Then
                                '����
                                .TextMatrix(l, .ColIndex("��")) = "��"
                            Else
                                '��β
                                .TextMatrix(l, .ColIndex("��")) = "��"
                            End If
                        End If
                        
                        '����PASS�����ͼƬ���ƣ�
                        If .ColHidden(.ColIndex("PASS")) = False And Not mobjPASS Is Nothing Then
                            Set .Cell(flexcpPicture, l, .ColIndex("PASS")) = mobjPASS.zlPassSetWarnLight_YF(Val(.TextMatrix(l, .ColIndex("PASS"))))
                            If Not .Cell(flexcpPicture, l, .ColIndex("PASS")) Is Nothing Then
                                .Cell(flexcpPictureAlignment, l, .ColIndex("PASS")) = flexPicAlignCenterCenter
                            End If
                            .TextMatrix(l, .ColIndex("PASS")) = ""      '����ʾ�ı���ֻ��ʾͼƬ
                        End If
                        
                        '����鴦����24Сʱ�ڵĴ������֣���ɫ������
                        .Cell(flexcpBackColor, l, 0, l, .Cols - 1) = IIf(Val(.TextMatrix(l, .ColIndex("��־"))) = 1, &H8000000F, .BackColor)
                        
                        'ҩƷ������Ӧ�ս��
                        If Not mobjPubAdvice Is Nothing Then
                            Call mobjPubAdvice.GetDurgPrice(Val(.TextMatrix(l, .ColIndex("ҽ��ID"))), dblPrice)
                            dblAmount = dblPrice * Val(.TextMatrix(l, .ColIndex("����")))
                            .TextMatrix(l, .ColIndex("����")) = Format(dblPrice, "#0.000")
                            .TextMatrix(l, .ColIndex("Ӧ�ս��")) = Format(dblAmount, "#0.00")
                        End If
                    Next
                    
                    .Row = 1
                End With
            End If
        
        Case 3      '3-�����Ŀ����
        
            If Me.tbcRec.Item(0).Selected Then
                Set vsfTmp = vsfRecWait
            Else
                Set vsfTmp = vsfRecFinish
            End If
            
            vsfAudit.ColHidden(vsfAudit.ColIndex("ɾ��")) = tbcRec.Item(0).Selected = False
            vsfAudit.Rows = 1
            vsfAudit.Clear 1
            
            If vsfTmp.Rows <= 1 Then
                mblnReasonRefresh = True
                txtReason.Text = ""
                mblnReasonRefresh = False
                Exit Sub
            End If
            
            lngRAID = Val(vsfRecipe.TextMatrix(vsfRecipe.Row, vsfRecipe.ColIndex("��ID")))
            
            If lngRAID <= 0 Then
                Exit Sub
            End If
            
            mblnReasonRefresh = True
            If Me.tbcRec(1).Selected Or Val(vsfRecipe.TextMatrix(vsfRecipe.Row, vsfRecipe.ColIndex("��־"))) = 1 Then
                strSQL = "Select �ۺ����� From ��������¼ Where ID = [1] "
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ��������¼��Ϣ", lngRAID)
                If rsTemp.EOF = False Then
                    txtReason.Text = NVL(rsTemp!�ۺ�����)
                End If
                rsTemp.Close
            Else
                txtReason.Text = ""
            End If
            mblnReasonRefresh = False
            
            strSQL = "Select Decode(b.ҩʦ���, 2, '���ϸ�', 1, '�ϸ�', Decode(b.�Զ����, 2, '���ϸ�', '�ϸ�')) ҩʦ���, " & vbNewLine & _
                     "    Decode(b.�Զ����, 2, '���ϸ�', 1, '�ϸ�', 'δ���') �Զ����, " & vbNewLine & _
                     "    b.ҽ��id, c.���� ҩƷ, d.id �����ĿID, d.���, d.����, d.���, d.���� ��Ŀ���� " & vbNewLine & _
                     "From ����ҽ����¼ A, ��������� B, �շ���ĿĿ¼ C, ���������Ŀ D " & vbNewLine & _
                     "Where a.Id(+) = b.ҽ��id And a.�շ�ϸĿid = c.Id(+) And b.�����Ŀid(+) = d.Id And b.��id(+) = [1] " & vbNewLine & _
                     "    And d.�Ƿ��������� = 1 And d.��� in ([2], 3, 4)" & vbNewLine & _
                     "Order By d.���, d.����"
            Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ѯ�������Ŀ��Ϣ", lngRAID, IIf(mintItem = 1, 2, 1))
            Call mdlDefine.FillVSFData(vsfAudit, rsTemp)
            
            If vsfAudit.Rows > 1 Then
                With vsfAudit
                    '���롢��ơ���Ŀ���ݺϲ���Ԫ��
                    .MergeCells = flexMergeRestrictColumns
                    .MergeCol(.ColIndex("����")) = True
                    .MergeCol(.ColIndex("���")) = True
                    .MergeCol(.ColIndex("��Ŀ����")) = True
                    
                    blnAllPass = True
                    For l = 1 To .Rows - 1
                        '�ϸ񡢲��ϸ��ͼƬ
                        If .TextMatrix(l, .ColIndex("ҩʦ���")) = "�ϸ�" Then
                            .Cell(flexcpPicture, l, .ColIndex("ҩʦ���")) = picAuditYN(0).Picture
                        Else
                            blnAllPass = False
                            .Cell(flexcpPicture, l, .ColIndex("ҩʦ���")) = picAuditYN(1).Picture
                            'ҩʦ��鲻�ϸ�ļ�¼��ǳ��ɫ��
                            .Cell(flexcpBackColor, l, 0, l, .Cols - 1) = &HC0C0FF
                        End If
                        
                        If .TextMatrix(l, .ColIndex("�Զ����")) = "�ϸ�" Then
                            .Cell(flexcpPicture, l, .ColIndex("�Զ����")) = picAuditYN(0).Picture
                        ElseIf .TextMatrix(l, .ColIndex("�Զ����")) = "���ϸ�" Then
                            .Cell(flexcpPicture, l, .ColIndex("�Զ����")) = picAuditYN(1).Picture
                        Else
                            .Cell(flexcpPicture, l, .ColIndex("�Զ����")) = Nothing
                            .Cell(flexcpText, l, .ColIndex("�Զ����")) = "δ���"
                        End If
                    Next
                    
                    .Row = 1
                End With
            End If
            
            '������ʾlblAudit�ؼ�
            Call mdlRecipeAudit.DispCountNG(vsfAudit, lblAudit)
            
'            '������Ŀ�ϸ�Ĭ�ϡ��ϸ񡢲��ϸ񡱰�ť�������
'            If blnAllPass And Me.tbcRec.Item(0).Selected Then
'                '�Զ�����
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
    
    strCard = iknPatient.Cards(iknPatient.IDKind).����

    If mblnReadCard Or KeyAscii = 13 Then
        Call zlControl.TxtSelAll(txtPatient)
    Else
        Select Case strCard
            Case "�����"
                If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
            Case "���֤��"
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
                If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
            Case Else
                If InStr(":��;��?��", Chr(KeyAscii)) > 0 Then
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
        lngAuditID = Val(vsfRecWait.TextMatrix(vsfRecWait.Row, vsfRecWait.ColIndex("��ID")))
        Call AuditLock(lngAuditID, True)
    End If
End Sub

Private Sub txtReason_KeyPress(KeyAscii As Integer)
    If InStr("""'", KeyAscii) Then KeyAscii = 0
End Sub

Private Sub vsfAudit_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim blnSign As Boolean
    
    If vsfRecipe.Rows > 1 And vsfRecipe.Row > 0 Then
        blnSign = Val(vsfRecipe.TextMatrix(vsfRecipe.Row, vsfRecipe.ColIndex("��־"))) = 1
    Else
        blnSign = True
    End If

    If Me.tbcRec.Item(0).Selected And mblnNeedAudit And mblnAuditStart And blnSign = False Then
        If Col = vsfAudit.ColIndex("ҩƷ") Then
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
    
    If Col = vsfAudit.ColIndex("ҩƷ") Then
    
        lngItemID = Val(vsfAudit.TextMatrix(Row, vsfAudit.ColIndex("�����ĿID")))
        lngAuditID = Val(vsfRecWait.TextMatrix(vsfRecWait.Row, vsfRecWait.ColIndex("��ID")))
        
        strSQL = "Select b.ҽ��id ID, d.���� ҩƷ����, d.���, d.���㵥λ ��λ, " & vbNewLine & _
                 "    A1.�ܸ����� ����, A1.�������� || Nvl(f.���㵥λ, '') ����, A2.ҽ������ �÷�, A1.ִ��Ƶ�� Ƶ��, " & vbNewLine & _
                 "    e.���� �ٴ�����, b.��id  " & vbNewLine & _
                 "From ����ҽ����¼ A1, ����ҽ����¼ A2, ���������ϸ B, �շ���ĿĿ¼ D, ���ű� E, ������ĿĿ¼ F " & vbNewLine & _
                 "Where A1.���id = A2.Id And A1.Id = b.ҽ��id And A1.�շ�ϸĿid = d.Id And A1.��������id = e.Id And A1.������Ŀid = f.id(+) " & vbNewLine & _
                 "    And b.��id = [1] " & vbNewLine & _
                 "Order By A1.���id, A1.Id "

        Set rsTemp = zlDatabase.ShowSQLMultiSelect(Me, strSQL, 0, "ѡ�������Ŀ��Ӧ��ҩƷ", False, "", "", _
                            False, False, False, 0, 0, 2000, blnCancel, False, False, _
                            lngAuditID)
        If blnCancel = False Then
            '��ѡ�������д��
            If Not rsTemp Is Nothing Then
                With rsTemp
                    Do While .EOF = False
                        '����ҽ��ID�Ƿ����
                        blnFind = False
                        For l = 1 To vsfAudit.Rows - 1
                            If Val(vsfAudit.TextMatrix(l, vsfAudit.ColIndex("ҽ��ID"))) = !ID And Val(vsfAudit.TextMatrix(l, vsfAudit.ColIndex("�����ĿID"))) = lngItemID Then
                                blnFind = True
                                Exit For
                            End If
                        Next
                        
                        'ҽ��ID�����ڣ��ٲ���
                        If blnFind = False Then
                            For l = vsfAudit.Rows - 1 To 1 Step -1
                                If Val(vsfAudit.TextMatrix(l, vsfAudit.ColIndex("�����ĿID"))) = lngItemID Then
                                    '���Ƶ�ǰ��¼
                                    Set colTmp = Nothing
                                    For intCol = 0 To vsfAudit.Cols - 1
                                        If intCol = vsfAudit.ColIndex("ҩʦ���") Then
                                            colTmp.Add "���ϸ�", CStr(intCol)
                                        ElseIf intCol = vsfAudit.ColIndex("�Զ����") Then
                                            colTmp.Add "�ϸ�", CStr(intCol)
                                        ElseIf intCol = vsfAudit.ColIndex("ҩƷ") Then
                                            colTmp.Add !ҩƷ����, CStr(intCol)
                                        ElseIf intCol = vsfAudit.ColIndex("ҽ��ID") Then
                                            colTmp.Add !ID, CStr(intCol)
                                        ElseIf intCol = vsfAudit.ColIndex("����") Then
                                            colTmp.Add "1", CStr(intCol)
                                        Else
                                            colTmp.Add vsfAudit.TextMatrix(l, intCol), CStr(intCol)
                                        End If
                                    Next
                                    
                                    '�����¼�¼
                                    If Not colTmp Is Nothing Then
                                        l = l + 1
                                        vsfAudit.AddItem "", l
                                        For intCol = 1 To colTmp.Count
                                            vsfAudit.TextMatrix(l, intCol - 1) = colTmp(intCol)
                                            If intCol - 1 = vsfAudit.ColIndex("ҩʦ���") Or intCol - 1 = vsfAudit.ColIndex("�Զ����") Then
                                                vsfAudit.Cell(flexcpPicture, l, intCol - 1) = IIf(colTmp(intCol) = "�ϸ�", picAuditYN(0).Picture, picAuditYN(1).Picture)
                                            ElseIf intCol - 1 = vsfAudit.ColIndex("ɾ��") Then
                                                vsfAudit.Cell(flexcpPicture, l, intCol - 1) = picAuditDel.Picture
                                                vsfAudit.Cell(flexcpPictureAlignment, l, intCol - 1) = flexPicAlignCenterCenter
                                            End If
                                        Next
                                        'ҩʦ��鲻�ϸ�ļ�¼��ǳ��ɫ��
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
            
            '����
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
        If Val(vsfRecipe.TextMatrix(vsfRecipe.Row, vsfRecipe.ColIndex("��־"))) = 1 Then
            Exit Sub
        End If
    Else
        Exit Sub
    End If

    With vsfAudit
        If .Rows <= 1 Then Exit Sub
        
        If .ColIndex("ҩʦ���") = .Col Then
            .TextMatrix(.Row, .Col) = IIf(.TextMatrix(.Row, .Col) = "���ϸ�", "�ϸ�", "���ϸ�")
            If .TextMatrix(.Row, .Col) = "�ϸ�" Then
                .Cell(flexcpPicture, .Row, .Col) = picAuditYN(0).Picture
                .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = .BackColor
            Else
                .Cell(flexcpPicture, .Row, .Col) = picAuditYN(1).Picture
                'ҩʦ��鲻�ϸ�ļ�¼��ǳ��ɫ��
                .Cell(flexcpBackColor, .Row, 0, .Row, .Cols - 1) = &HC0C0FF
            End If
            
            Call DispCountNG(vsfAudit, lblAudit)
            
            If mblnLocking = False Then
                lngAuditID = Val(vsfRecWait.TextMatrix(vsfRecWait.Row, vsfRecWait.ColIndex("��ID")))
                Call AuditLock(lngAuditID, True)
            End If
        ElseIf .ColIndex("ɾ��") = .Col Then
            If Val(.TextMatrix(.Row, .ColIndex("����"))) = 1 Then
                If MsgBox("ȷ��ɾ���ü�¼��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    .RemoveItem .Row
                    Call DispCountNG(vsfAudit, lblAudit)
                End If
            End If
        End If
        
    End With
End Sub

Private Sub vsfAudit_KeyPress(KeyAscii As Integer)
    If vsfAudit.Col = vsfAudit.ColIndex("ҩʦ���") Then
        If KeyAscii = vbKeySpace Then Call vsfAudit_Click
    End If
End Sub

Private Sub vsfRecFinish_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If OldRow <> NewRow Then
        Dim lngPatientID As Long, lngMedID As Long
        
        If vsfRecFinish.TextMatrix(NewRow, vsfRecFinish.ColIndex("�����")) = "�ϸ�" Then
            lblDisp_Fixed.Caption = "�ϸ�"
            lblDisp_Fixed.ForeColor = &H8000&
        ElseIf vsfRecFinish.TextMatrix(NewRow, vsfRecFinish.ColIndex("�����")) = "���ϸ�" Then
            lblDisp_Fixed.Caption = "���ϸ�"
            lblDisp_Fixed.ForeColor = vbRed
        Else
            lblDisp_Fixed.Caption = "δ֪"
            lblDisp_Fixed.ForeColor = vbBlack
        End If
        lblDisp_Fixed.Left = (picYesNO.ScaleWidth - lblDisp_Fixed.Width) \ 2
        
        Call FillVSFData(2)
        'Call FillVSFData(3)
        
        'ˢ�µ��Ӳ���
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
            If Val(vsfRecipe.TextMatrix(NewRow, vsfRecipe.ColIndex("��־"))) = 1 Then
                vsfAudit.Clear 1
                vsfAudit.Rows = 1
            Else
                Call FillVSFData(3)
            End If
            Call RefreshLockControls
            If Not mobjPASS Is Nothing Then
                Call mobjPASS.zlPassSetDrug_YF(vsfRecipe.TextMatrix(NewRow, vsfRecipe.ColIndex("ϸĿID")), _
                                               vsfRecipe.TextMatrix(NewRow, vsfRecipe.ColIndex("ҩƷ����")))
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
        If (.TextMatrix(.MouseRow, .MouseCol) <> "" And .MouseCol = .ColIndex("����˵��") _
            Or .TextMatrix(.MouseRow, .MouseCol) <> "" And .MouseCol = .ColIndex("ҽ������")) Then
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
        
        'ˢ�µ��Ӳ���
        If Me.tbcTab.ItemCount > 1 Then
            If Me.tbcTab.Item(1).Selected Then
                Call RefreshAMR(vsfRecWait)
            End If
        End If
        
        If vsfRecWait.Visible And vsfRecWait.Enabled Then vsfRecWait.SetFocus
    End If
End Sub

Private Sub SetControlFontSize(ByVal bytSize As Byte)
'���ܣ����ô���ؼ��������С
'������
'  bytSize��0-С���壻1-������

    mbytFontSize = bytSize
    
    Call SetPublicFontSize(Me, bytSize)
    Call picRec_Resize
    
End Sub

Private Sub RefreshLockControls()
'���ܣ�ˢ����������������״̬����ؿؼ�
    
    Dim blnAllPass As Boolean
    
    blnAllPass = GetAuditResult(vsfAudit)
    
    '������Ŀ�ϸ�Ĭ�ϡ��ϸ񡢲��ϸ񡱰�ť�������
    cmdYes_Fixed.Enabled = mblnAuditStart And (blnAllPass Or mblnLocking)
    cmdNo_Fixed.Enabled = mblnAuditStart And (blnAllPass Or mblnLocking)
    
    txtReason.Enabled = mblnAuditStart And (vsfAudit.Rows > 1) And tbcRec.Item(0).Selected
    If vsfRecipe.ColIndex("��־") >= 0 Then
        txtReason.Enabled = txtReason.Enabled And Val(vsfRecipe.TextMatrix(vsfRecipe.Row, vsfRecipe.ColIndex("��־"))) = 0
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
'���ܣ�������鴦��
'������
'  bytFun��1-�ϸ�2-���ϸ�

    Dim strSQL As String, strReason As String, strTmp As String
    Dim lngAuditID As Long, lngMedicalID As Long, lngItemID As Long
    Dim l As Long
    Dim colSQL As New Collection
    Dim blnFind As Boolean, blnNoTrans As Boolean
    Dim strIDs As String
    Dim lngPatientID As Long, lngRegisterID As Long
    
    '�ύǰ�ļ��
    
    If LenB(StrConv(Trim(txtReason.Text), vbFromUnicode)) > 500 Then
        MsgBox "���ۺ����ɡ����ݳ��ޣ�250�����ֻ�500���ַ�����", vbInformation, gstrSysName
        txtReason.SetFocus
        Exit Sub
    End If
    
    strReason = Trim(txtReason.Text)
    strReason = Replace(Replace(Replace(strReason, vbLf, ""), vbCr, ""), vbNewLine, "")
    
    With vsfAudit
        blnFind = False
        For l = 1 To .Rows - 1
            If bytFun = 1 Then
                '���ҩʦ�����ж��ϸ񣬵���ϸ�в��ϸ�����
                If Trim(.TextMatrix(l, .ColIndex("ҩʦ���"))) = "���ϸ�" Then
                    MsgBox "�������Ŀ���ġ�ҩʦ��顱�������յĽ������������飡", vbInformation, gstrSysName
                    .Row = l
                    .SetFocus
                    Exit Sub
                End If
            Else
                '���ҩʦ�����ж����ϸ񣬵���ϸû�в��ϸ�����
                If .TextMatrix(l, .ColIndex("ҩʦ���")) = "���ϸ�" Then
                    blnFind = True
                End If
            End If
            
            '����ۺ����ɣ��Զ�����ҩʦ����С����ϸ񡱵ı�����д�ۺ�����
            If strReason = "" And (Trim(.TextMatrix(l, .ColIndex("ҩʦ���"))) = "���ϸ�" Or Trim(.TextMatrix(l, .ColIndex("�Զ����"))) = "���ϸ�") Then
                MsgBox "����д���ġ��ۺ����ɡ���", vbInformation, gstrSysName
                .Row = l
                If txtReason.Enabled And txtReason.Visible Then txtReason.SetFocus
                Exit Sub
            
            '��������ĿID<=0�ļ�¼
            ElseIf Val(.TextMatrix(l, .ColIndex("�����ĿID"))) <= 0 Then
                MsgBox "�����Ŀ�쳣��", vbInformation, gstrSysName
                .Row = l
                .SetFocus
                Exit Sub
            End If
        Next
        
        If bytFun = 2 And blnFind = False Then
            MsgBox "�������Ŀ���ġ�ҩʦ��顱���޲��ϸ����飡", vbInformation, gstrSysName
            .Row = 1
            .SetFocus
            Exit Sub
        End If
    End With
    
    lngAuditID = Val(vsfRecWait.TextMatrix(vsfRecWait.Row, vsfRecWait.ColIndex("��ID")))
    If lngAuditID <= 0 Then
        MsgBox "�����¼�쳣����������ֹ��", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '��������¼��Ϣ
    
    strSQL = zlStr.FormatString("Zl_�������_Audit([1], [2], [3], [4])", _
                    lngAuditID, _
                    bytFun, _
                    "'" & UserInfo.���� & "'", _
                    IIf(strReason = "", "Null", "'" & Trim(txtReason.Text) & "'"))
    AddArray colSQL, strSQL
    
    '���������ϸ��Ϣ
    
    With vsfAudit
        For l = 1 To .Rows - 1
            If .TextMatrix(l, .ColIndex("ҩʦ���")) = "���ϸ�" Or .TextMatrix(l, .ColIndex("�Զ����")) = "���ϸ�" Then
                lngMedicalID = Val(.TextMatrix(l, .ColIndex("ҽ��ID")))
                lngItemID = Val(.TextMatrix(l, .ColIndex("�����ĿID")))
                
                strSQL = zlStr.FormatString("Zl_�������_Audit_Detail([1], [2], [3], [4])", _
                                lngAuditID, _
                                IIf(lngMedicalID <= 0, "Null", lngMedicalID), _
                                lngItemID, _
                                IIf(.TextMatrix(l, .ColIndex("ҩʦ���")) = "�ϸ�", 1, 2))
                AddArray colSQL, strSQL
            End If
        Next
    End With
    
    '����������ʱ��
    strSQL = zlStr.FormatString("ZL_����������_SAVE(2, '[1]', 0, 1, Null)", mstrPCName)
    AddArray colSQL, strSQL
    
    'ִ�д洢����
    Err.Clear: On Error GoTo errHandle: blnNoTrans = False
    ExecuteProcedureArray colSQL, Me.Caption, blnNoTrans
    blnNoTrans = True
    
    '��ȡ����ID���Һŵ�ID
    With vsfRecWait
        lngPatientID = Val(.TextMatrix(.Row, .ColIndex("����ID")))
        lngRegisterID = Val(.TextMatrix(.Row, .ColIndex("�Һŵ�ID")))
    End With
    
    '��ȡ���IDs
    With vsfRecipe
        For l = 1 To .Rows - 1
            strIDs = strIDs & "," & .TextMatrix(l, .ColIndex("���ID"))
        Next
        If strIDs <> "" Then strIDs = Mid(strIDs, 2)
    End With
    
    mblnReasonRefresh = True
    txtReason.Text = ""
    mblnReasonRefresh = False
    
    '�ɹ���ɣ�״̬����
    mblnLocking = False
'    vsfRecWait.RemoveItem vsfRecWait.Row
    
    Screen.MousePointer = vbHourglass
    Call FillVSFData(1)
    Call FillVSFData(2)
    Call RefreshLockControls
    Call SetStatusbar
    Call RefreshAMR
    
    '����ҽ���Զ�����
    If bytFun = Val("1-�ϸ�") And Not mobjPubAdvice Is Nothing Then
        If mblnSendBeforeAudit And mblnRecipeSendAuto Then
            '�������ƣ���������ǰ�󷽣��󷽺ϸ��Զ����ʹ���
            If strIDs <> "" And mobjPubAdvice.OutAdviceSendDrug(Me, strIDs, lngPatientID, lngRegisterID) Then
                '����ҽ���Զ����ͳɹ���������Ϣ
            Else
                '������Ϣ֪ͨҽ��
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
'���ܣ�����/�����л�
'������
'  lngAuditID����ID
'  blnLock��True������False����
    
    Dim strSQL As String
    
    On Error GoTo errHandle
    
    strSQL = zlStr.FormatString("Zl_��������¼_Lock([1], '[2]', 0, [3])", _
                    IIf(blnLock, 1, 0), _
                    mstrPCName, _
                    IIf(lngAuditID <= 0, "Null", lngAuditID))
    Call zlDatabase.ExecuteProcedure(strSQL, IIf(blnLock, "����¼����", "����¼����"))
    
    mblnLocking = blnLock
    Call RefreshLockControls
    Exit Sub
    
errHandle:
    Call ErrCenter
    If gcnOracle.Errors(0).Description Like "*�ѱ����*" Or gcnOracle.Errors(0).Description Like "*�ѱ�ɾ��*" Then
        vsfRecWait.RemoveItem vsfRecWait.Row
    End If
    Call FillVSFData(3)
    Call RefreshLockControls
End Sub

Private Sub vsfRecWait_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If OldRow <> NewRow Then
        If mblnLocking Then
            If MsgBox("������鵱ǰ���˵�ҩ�����Ƿ������飿", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = True
            Else
                Call AuditLock(Val(vsfRecWait.TextMatrix(OldRow, vsfRecWait.ColIndex("��ID"))), False)
            End If
        End If
    End If
End Sub

Private Function AuditOperate(ByVal bytFun As Byte) As Boolean
'���ܣ���ǰ��������/ֹͣ�������
'������
'  bytFun��0-ֹͣ��1-����
'���أ�True�ɹ���Falseʧ��
    
    'Dim lngAuditID As Long
    Dim strSQL As String

    On Error GoTo errHandle

    If bytFun = 1 Then
        '����
        strSQL = zlStr.FormatString("ZL_����������_SAVE(2, '[1]', 0, 1, Null)", mstrPCName)
        Call zlDatabase.ExecuteProcedure(strSQL, "���¿����󷽡�������ʱ��")
    Else
        'ֹͣ
        If mblnLocking Then
            '��������л���ֹͣ���
            If MsgBox("������鵱ǰ���˵�ҩ�����Ƿ������飿", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
        
        '�����Ƿ����󷽡�������ʱ��
        strSQL = zlStr.FormatString("ZL_����������_SAVE(2, '[1]', 0, 0, Null)", mstrPCName)
        Call zlDatabase.ExecuteProcedure(strSQL, "����ֹͣ�󷽡�������ʱ��")
        
        '����
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
    
    txtOther.Text = zlStr.FormatString("��ϣ�[1]������Ҫ����", vbCrLf)
    
    If vsfTmp.Rows <= 1 Then Exit Sub
    
    lngPatientID = Val(vsfTmp.TextMatrix(vsfTmp.Row, vsfTmp.ColIndex("����ID")))
    lngRegisterID = Val(vsfTmp.TextMatrix(vsfTmp.Row, vsfTmp.ColIndex("�Һŵ�ID")))
    lngMedicalID = Val(vsfRecipe.TextMatrix(vsfRecipe.Row, vsfRecipe.ColIndex("���ID")))
    
    '���
    If Not mobjPubAdvice Is Nothing Then
        Call mobjPubAdvice.GetAdviceDiag(lngMedicalID, strDiagnose)
    End If
    txtOther.Text = zlStr.FormatString("��ϣ�[1][2]", strDiagnose, vbCrLf)
    
    '������Ҫ��
    strTmp = GetCalorie(lngPatientID, lngRegisterID, 0)
    txtOther.Text = txtOther & zlStr.FormatString("������Ҫ����[1]", strTmp)
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
        stbThis.Panels(2).Text = zlStr.FormatString("��ǰ[1]���¼������[2]��", IIf(tbcRec.Item(0).Selected, "��", "��"), vsfTmp.Rows - 1)
    End If
    
End Sub

Private Sub SetFilterDay(ByVal bytMode As Byte)
    If tbcRec.ItemCount <= 1 Then Exit Sub

    If bytMode = 0 Then
        tbcRec.Item(1).Tag = cboDate.ListIndex
        With cboDate
            .Clear
            .AddItem "����"
            .AddItem "������"
            .AddItem "������"
        End With
    Else
        tbcRec.Item(0).Tag = cboDate.ListIndex
        With cboDate
            .Clear
            .AddItem "����"
            .AddItem "������"
            .AddItem "������"
            .AddItem "������"
            .AddItem "������"
            .AddItem "������"
            .AddItem "������"
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
    
    '������������Ƿ�����
    strSQL = "Select Count(1) Rec From ����������� Where Rownum < 2 "
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������������Ƿ�����")
    If rsTemp!Rec <= 0 Then
        rsTemp.Close
        If Me.Visible = False Then
            MsgBox "���������������δ���κ����ã����飡", vbInformation, gstrSysName
        End If
        Exit Sub
    End If
    rsTemp.Close
    
    '�п������ã������ÿ�����ȡ
    strSQL = "Select a.Id ����id, a.����, a.���� " & vbNewLine & _
             "From ���ű� A, Table(f_Num2list((Select ��Դ���� From ���������� Where ������ = [1] And ������� = 0), ',')) B " & vbNewLine & _
             "Where a.Id = b.Column_Value " & vbNewLine & _
             "Order By a.���� "
'    strSQL = "Select a.����id, b.����, b.���� " & vbNewLine & _
'             "From ��������˵�� A, ���ű� B, ����������� C," & vbNewLine & _
'             "    Table(f_Num2list((Select ��Դ���� From ���������� Where ������ = [1] And ������� = 0))) D " & vbNewLine & _
'             "Where a.����id = b.Id And a.����id = c.����id And a.����id = d.Column_Value And a.�������� = '�ٴ�' " & vbNewLine & _
'             "    And a.������� In (1, 3) And (b.����ʱ�� Is Null Or To_Char(b.����ʱ��, 'yyyy') = '3000') " & vbNewLine & _
'             "    And c.��� = 1 And (c.����id Is Not Null Or c.����id > 0) "

    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ���ﴦ��������Դ����", mstrPCName)
    
    With rsTemp
        cboClinic.Clear
        cboClinic.AddItem "������Դ����"
        Do While .EOF = False
            cboClinic.AddItem !����
            cboClinic.ItemData(cboClinic.NewIndex) = !����ID
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

    'ˢ���ٴ���Ϣ
    If Me.Visible Then
        If vsfVal Is Nothing Then
            If Me.tbcRec.Item(0).Selected Then
                lngPatientID = Val(vsfRecWait.TextMatrix(vsfRecWait.Row, vsfRecWait.ColIndex("����ID")))
                lngMediID = Val(vsfRecWait.TextMatrix(vsfRecWait.Row, vsfRecWait.ColIndex("�Һŵ�ID")))
            Else
                lngPatientID = Val(vsfRecFinish.TextMatrix(vsfRecFinish.Row, vsfRecFinish.ColIndex("����ID")))
                lngMediID = Val(vsfRecFinish.TextMatrix(vsfRecFinish.Row, vsfRecFinish.ColIndex("�Һŵ�ID")))
            End If
        Else
            lngPatientID = Val(vsfVal.TextMatrix(vsfVal.Row, vsfVal.ColIndex("����ID")))
            lngMediID = Val(vsfVal.TextMatrix(vsfVal.Row, vsfVal.ColIndex("�Һŵ�ID")))
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
        .Interval = Val(zlDatabase.GetPara("�Զ�ˢ�²����б�", glngSys, mlngModule)) * 1000
        If .Interval > 0 Then .Enabled = True
    End With
End Sub

