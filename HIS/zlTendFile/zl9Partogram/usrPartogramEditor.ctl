VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.UserControl usrPartogramEditor 
   AutoRedraw      =   -1  'True
   ClientHeight    =   5550
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8565
   KeyPreview      =   -1  'True
   ScaleHeight     =   5550
   ScaleWidth      =   8565
   Begin VB.PictureBox picTmp 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
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
      Height          =   300
      Left            =   840
      ScaleHeight     =   300
      ScaleWidth      =   1005
      TabIndex        =   31
      TabStop         =   0   'False
      Top             =   120
      Width           =   1000
      Begin VB.ComboBox cboBaby 
         Height          =   300
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   0
         Width           =   1005
      End
   End
   Begin MSComctlLib.ImageList imgRow 
      Left            =   7200
      Top             =   600
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
            Picture         =   "usrPartogramEditor.ctx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "usrPartogramEditor.ctx":039A
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.TextBox txtLength 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   1005
      Left            =   2340
      MultiLine       =   -1  'True
      TabIndex        =   15
      Top             =   3090
      Visible         =   0   'False
      Width           =   2025
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4515
      Left            =   60
      ScaleHeight     =   4515
      ScaleWidth      =   8385
      TabIndex        =   10
      Top             =   510
      Width           =   8385
      Begin VB.PictureBox picBaby 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1695
         Left            =   5040
         ScaleHeight     =   1695
         ScaleWidth      =   1965
         TabIndex        =   34
         Top             =   120
         Visible         =   0   'False
         Width           =   1965
         Begin VB.CommandButton cmdAddBaby 
            Height          =   315
            Left            =   960
            Picture         =   "usrPartogramEditor.ctx":0734
            Style           =   1  'Graphical
            TabIndex        =   37
            ToolTipText     =   "���"
            Top             =   1320
            Width           =   450
         End
         Begin VB.CommandButton cmdBabyCancle 
            Height          =   315
            Left            =   1440
            Picture         =   "usrPartogramEditor.ctx":0CBE
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "ȡ��"
            Top             =   1320
            Width           =   450
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfBaby 
            Height          =   1300
            Left            =   0
            TabIndex        =   35
            Top             =   0
            Width           =   1935
            _cx             =   3413
            _cy             =   2311
            Appearance      =   0
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
            BackColor       =   -2147483624
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16764057
            ForeColorSel    =   0
            BackColorBkg    =   -2147483624
            BackColorAlternate=   -2147483624
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   0
            GridLinesFixed  =   0
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   1
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   300
            ColWidthMin     =   1900
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   0   'False
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
            Begin VB.CommandButton cmdDelBaby 
               Height          =   300
               Left            =   1600
               Picture         =   "usrPartogramEditor.ctx":1248
               Style           =   1  'Graphical
               TabIndex        =   36
               ToolTipText     =   "ɾ��"
               Top             =   300
               Width           =   300
            End
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfHead 
         Height          =   795
         Left            =   0
         TabIndex        =   39
         Top             =   915
         Width           =   4305
         _cx             =   7594
         _cy             =   1402
         Appearance      =   0
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
         AllowUserResizing=   0
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
         FormatString    =   $"usrPartogramEditor.ctx":17D2
         ScrollTrack     =   -1  'True
         ScrollBars      =   0
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
      Begin VB.PictureBox picSign 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H00C0FFFF&
         ForeColor       =   &H80000008&
         Height          =   225
         Left            =   4380
         ScaleHeight     =   195
         ScaleWidth      =   945
         TabIndex        =   40
         Tag             =   "225"
         Top             =   3345
         Visible         =   0   'False
         Width           =   975
         Begin VB.Label lblCheckSign 
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "��֤ǩ��"
            ForeColor       =   &H00FF0000&
            Height          =   180
            Left            =   210
            TabIndex        =   41
            Top             =   0
            Width           =   720
         End
         Begin VB.Image imgSign 
            Height          =   240
            Left            =   -30
            Picture         =   "usrPartogramEditor.ctx":1834
            Tag             =   "240"
            Top             =   -30
            Width           =   240
         End
      End
      Begin VB.CheckBox chkSwitch 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ȫѡ"
         ForeColor       =   &H80000008&
         Height          =   210
         Left            =   0
         TabIndex        =   22
         TabStop         =   0   'False
         Top             =   930
         Visible         =   0   'False
         Width           =   195
      End
      Begin VB.PictureBox picDoubleChoose 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   6330
         ScaleHeight     =   240
         ScaleWidth      =   900
         TabIndex        =   24
         TabStop         =   0   'False
         Top             =   3300
         Visible         =   0   'False
         Width           =   930
         Begin VB.PictureBox picChooseRight 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   540
            ScaleHeight     =   255
            ScaleWidth      =   375
            TabIndex        =   27
            TabStop         =   0   'False
            Top             =   0
            Width           =   375
            Begin VB.ComboBox cboChoose 
               BackColor       =   &H80000018&
               Height          =   300
               Index           =   1
               ItemData        =   "usrPartogramEditor.ctx":8086
               Left            =   -30
               List            =   "usrPartogramEditor.ctx":8096
               Style           =   2  'Dropdown List
               TabIndex        =   28
               Top             =   -30
               Width           =   1605
            End
         End
         Begin VB.PictureBox picChooseLeft 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   435
            TabIndex        =   25
            TabStop         =   0   'False
            Top             =   0
            Width           =   435
            Begin VB.ComboBox cboChoose 
               BackColor       =   &H80000018&
               Height          =   300
               Index           =   0
               ItemData        =   "usrPartogramEditor.ctx":80A8
               Left            =   -30
               List            =   "usrPartogramEditor.ctx":80B8
               Style           =   2  'Dropdown List
               TabIndex        =   26
               Top             =   -30
               Width           =   1605
            End
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "/"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   435
            TabIndex        =   29
            Top             =   30
            Width           =   105
         End
      End
      Begin VB.PictureBox picMutilInput 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   435
         Left            =   6060
         ScaleHeight     =   405
         ScaleWidth      =   1575
         TabIndex        =   8
         Top             =   3720
         Visible         =   0   'False
         Width           =   1600
         Begin VB.TextBox txt 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Height          =   225
            Index           =   0
            Left            =   810
            TabIndex        =   9
            Top             =   90
            Width           =   675
         End
         Begin VB.Label lbl 
            Alignment       =   1  'Right Justify
            Appearance      =   0  'Flat
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "������¼"
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   15
            TabIndex        =   13
            Top             =   112
            Width           =   720
         End
      End
      Begin VB.CommandButton cmdWord 
         BackColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6120
         Picture         =   "usrPartogramEditor.ctx":80CA
         Style           =   1  'Graphical
         TabIndex        =   21
         TabStop         =   0   'False
         Top             =   1290
         Visible         =   0   'False
         Width           =   285
      End
      Begin VB.PictureBox picDouble 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   270
         Left            =   6330
         ScaleHeight     =   240
         ScaleWidth      =   900
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   2910
         Visible         =   0   'False
         Width           =   930
         Begin VB.PictureBox picDnInput 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   540
            ScaleHeight     =   255
            ScaleWidth      =   375
            TabIndex        =   18
            TabStop         =   0   'False
            Top             =   0
            Width           =   375
            Begin VB.Label lblDnInput 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H80000008&
               Height          =   195
               Left            =   60
               TabIndex        =   20
               Top             =   30
               Width           =   315
            End
         End
         Begin VB.PictureBox picUpInput 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   0
            ScaleHeight     =   255
            ScaleWidth      =   435
            TabIndex        =   17
            TabStop         =   0   'False
            Top             =   0
            Width           =   435
            Begin VB.Label lblUpInput 
               Alignment       =   2  'Center
               Appearance      =   0  'Flat
               BackColor       =   &H80000005&
               BackStyle       =   0  'Transparent
               ForeColor       =   &H80000008&
               Height          =   165
               Left            =   60
               TabIndex        =   19
               Top             =   30
               Width           =   315
            End
         End
         Begin VB.TextBox txtDnInput 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   240
            Left            =   525
            MaxLength       =   12
            TabIndex        =   7
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
            TabIndex        =   6
            Top             =   30
            Width           =   375
         End
         Begin VB.Label lblSplit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "/"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   435
            TabIndex        =   16
            Top             =   30
            Width           =   105
         End
      End
      Begin VB.ListBox lstSelect 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   2550
         Index           =   1
         ItemData        =   "usrPartogramEditor.ctx":840C
         Left            =   6660
         List            =   "usrPartogramEditor.ctx":8422
         Style           =   1  'Checkbox
         TabIndex        =   4
         Top             =   1590
         Visible         =   0   'False
         Width           =   1485
      End
      Begin VB.PictureBox picInput 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000018&
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5790
         ScaleHeight     =   225
         ScaleWidth      =   585
         TabIndex        =   1
         Top             =   1290
         Visible         =   0   'False
         Width           =   615
         Begin VB.TextBox txtInput 
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            BorderStyle     =   0  'None
            Height          =   180
            Left            =   0
            MultiLine       =   -1  'True
            TabIndex        =   2
            Top             =   30
            Width           =   315
         End
         Begin VB.Label lblInput 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H80000018&
            Caption         =   "��"
            ForeColor       =   &H80000008&
            Height          =   180
            Left            =   0
            TabIndex        =   14
            Top             =   30
            Width           =   315
         End
      End
      Begin VB.ListBox lstSelect 
         Appearance      =   0  'Flat
         BackColor       =   &H80000018&
         Height          =   2550
         Index           =   0
         ItemData        =   "usrPartogramEditor.ctx":845A
         Left            =   5790
         List            =   "usrPartogramEditor.ctx":8470
         TabIndex        =   3
         Top             =   1590
         Visible         =   0   'False
         Width           =   1215
      End
      Begin VSFlex8Ctl.VSFlexGrid VsfData 
         Height          =   2655
         Left            =   0
         TabIndex        =   0
         Top             =   945
         Width           =   4305
         _cx             =   7594
         _cy             =   4683
         Appearance      =   0
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
         AllowUserResizing=   0
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   4
         FixedRows       =   0
         FixedCols       =   1
         RowHeightMin    =   255
         RowHeightMax    =   5000
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"usrPartogramEditor.ctx":84A8
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
      Begin VSFlex8Ctl.VSFlexGrid vsTest 
         Height          =   495
         Left            =   1920
         TabIndex        =   30
         Top             =   930
         Visible         =   0   'False
         Width           =   1845
         _cx             =   3254
         _cy             =   873
         Appearance      =   0
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
         AllowUserResizing=   0
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
         FormatString    =   $"usrPartogramEditor.ctx":850A
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
      Begin VB.PictureBox picSignCheck 
         Appearance      =   0  'Flat
         AutoRedraw      =   -1  'True
         BackColor       =   &H80000018&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2865
         Left            =   4365
         ScaleHeight     =   2865
         ScaleWidth      =   4815
         TabIndex        =   42
         Top             =   3600
         Visible         =   0   'False
         Width           =   4815
         Begin VB.CommandButton cmdSignAll 
            Caption         =   "ȫ��"
            Height          =   350
            Left            =   270
            TabIndex        =   45
            ToolTipText     =   "ȷ��"
            Top             =   2370
            Width           =   840
         End
         Begin VB.CommandButton cmdSignCur 
            Caption         =   "��֤"
            Height          =   350
            Left            =   2790
            TabIndex        =   44
            ToolTipText     =   "ȷ��"
            Top             =   2370
            Width           =   840
         End
         Begin VB.CommandButton cmdCancl 
            Caption         =   "ȡ��"
            Height          =   350
            Left            =   3690
            TabIndex        =   43
            ToolTipText     =   "ȡ��"
            Top             =   2370
            Width           =   840
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfSignData 
            Height          =   1635
            Left            =   0
            TabIndex        =   46
            Top             =   630
            Width           =   4755
            _cx             =   8387
            _cy             =   2884
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
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   4
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   255
            RowHeightMax    =   5000
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"usrPartogramEditor.ctx":856C
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
         Begin VB.Label Label4 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "������ǩ����ʷ��¼����ѡ������֤��Ҳ�ɽ���ȫ����֤��"
            ForeColor       =   &H80000008&
            Height          =   360
            Left            =   810
            TabIndex        =   47
            Top             =   150
            Width           =   3720
         End
         Begin VB.Image imgNote 
            Height          =   480
            Left            =   120
            Picture         =   "usrPartogramEditor.ctx":85CE
            Top             =   90
            Width           =   480
         End
      End
      Begin VB.Label lblSubEnd 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ע:##"
         Height          =   180
         Left            =   1320
         TabIndex        =   33
         Top             =   480
         Width           =   630
      End
      Begin VB.Label lblCurPage 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "P333"
         BeginProperty Font 
            Name            =   "����"
            Size            =   12
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C00000&
         Height          =   240
         Left            =   7650
         TabIndex        =   23
         Top             =   240
         Width           =   555
      End
      Begin VB.Label lblTitle 
         Alignment       =   2  'Center
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "һ�㻤���¼��"
         Height          =   180
         Left            =   3450
         TabIndex        =   12
         Top             =   30
         Width           =   1275
      End
      Begin VB.Label lblSubhead 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����:##"
         Height          =   180
         Left            =   390
         TabIndex        =   11
         Top             =   540
         Width           =   720
         WordWrap        =   -1  'True
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
Attribute VB_Name = "usrPartogramEditor"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'��������:
'1.�����¼ͬһʱ��ֻ���ܴ���һ����¼
'2.�����¼�в���Ҫ�����µ����� , ��¼�����Ƿ����, �ܲ������, �����˵����ݲż�¼
'3.¼�뻤���¼����ʱ,�����¼������ݴ�����������, ����ȡ����
'4.�����¼�в���Ҫ¼�������¼�������׾����ȷ��Ҫ��¼���ڻ���ժҪ�������͵�����
'#ʵ��ԭ��:
'1.���Ӽ�¼����¼��Щҳ��Щ��Ԫ���û��޸Ĺ�
'2.�κα༭(ճ��,�������),����Ҫ���¼���ÿ�����ݵ�ռ����

Public mblnEditable As Boolean
'Public objFileSys As New FileSystemObject
'Public objStream As TextStream

Private Const mintҳ�� As Integer = 1
Private mFrmParent As Object
Private mblnInit As Boolean
Private mblnShow As Boolean                 '�Ƿ���ʾ¼���
Private mblnVerify As Boolean               '�Ƿ���ǩģʽ(���޸�,����������и���ճ������Ȳ���,ֻ���޸�)
Private mstrVerify As String                '�ȴ���ǩ��ID��
Private mintVerify As Integer               '��ǰ����Ա����߼���
Private mintVerify_Last As Integer          '��ѡ��ǩ��¼����߼���
Private mblnBlowup As Boolean               '�Ŵ�񣿷Ŵ�1/3��������9�ŷŴ�Ϊ12��
Private mblnChange As Boolean               '�Ƿ��޸�����
Private mstrData As String                  '����༭״̬ǰ����֮ǰ������
Private mintPreDays As Long
Private mstrMaxDate As String

Private mlng�ļ�ID As Long
Private mlng��ʽID As Long
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mlng����id As Long
Private mlng����ID As Long
Private mintӤ�� As Integer
Private mlngFileIndex As Long
Private mstrPrivs As String

Private mintSymbol As Integer               '��ǰ�ؼ�����
Private mstrSymbol As String                '�����ַ�
Private mstrCOLNothing As String            'δ�󶨵��м���
Private mstrCatercorner As String           '�жԽ��߼���
Private mblnEditAssistant As Boolean        '��ǰѡ�����Ŀ�Ƿ�������дʾ�ѡ��
Private mlngRowCount As Long                '��ǰ��¼������
Private mlngDate As Long                    '����
Private mlngTime As Long                    'ʱ��
Private mlngSpread As Long                  '��������
Private mlngJust As Long                    '��¶����
Private mlngProduce As Long                 '����
Private mlngChoose As Long                  'ѡ����
Private mlngOperator As Long                '��ʿ
Private mlngSignLevel As Long               'ǩ������
Private mlngSigner As Long                  'ǩ����Ϣ
Private mlngSignName As Long                'ǩ����
Private mlngSignTime As Long                'ǩ��ʱ��
Private mlngRecord As Long                  '��¼ID
Private mlngNoEditor As Long                '��ֹ�༭��,���ڻ�ʿ�����Ի�ʿ��Ϊ׼,�����ڻ�ʿ������ǩ����Ϊ׼
Private mlngDemo As Long                    '������
Private mlngActTime As Long                 '����ʱ��

Private mblnSign As Boolean                 '�Ƿ�ǩ��
Private mblnArchive As Boolean              '�Ƿ�鵵
Private mintType As Integer                 '��¼��ǰ�ı༭ģʽ
Private mblnDateAd As Boolean               '������д?
Private mblnDate As Boolean                 '�Ƿ����������
Private mstr��ʼʱ�� As String              '��ǰ�ļ��Ŀ�ʼʱ��
Private mstr����ʱ�� As String              '��ǰ�ļ��Ľ���ʱ��
Private mstrBeginTime As String             '������ʼʱ��
Private CellRect As RECT

Private rsTemp As New ADODB.Recordset
Private mrsItems As New ADODB.Recordset             '���л����¼��Ŀ�嵥
Private mrsPartogram As New ADODB.Recordset         '����Ҫ�ؼ�¼�嵥
Private mrsSelItems As New ADODB.Recordset          '��ǰ¼��Ļ����¼��Ŀ�嵥
Private mrsDataMap As New ADODB.Recordset           '��ǰ����Ա¼������ݾ���,�����ݱ༭��ʽһ��,���������ȫ�������Ա�Ѹ�ٻָ�
Private mrsCellMap As New ADODB.Recordset           '�༭�������ݾ���,�ֶ���:ҳ��,�к�,�к�,��¼ID,����,��λ,ɾ��
Private mrsCopyMap As New ADODB.Recordset           '����������

Private Enum ColIcon
    ǩ�� = 1
    ��ǩ = 2
End Enum
Private Enum SignLevel
    ���� = 1
    ���� = 2
    �м� = 3
    ʦ�� = 4
    Աʿ = 5
    δ���� = 9
End Enum

Private Const conMenu_Save = 1

Private Const WS_MAXIMIZE = &H1000000
Private Const WS_MAXIMIZEBOX = &H10000
Private Const WS_MINIMIZEBOX = &H20000
Private Const WS_CAPTION = &HC00000
Private Const WS_SYSMENU = &H80000
Private Const WS_THICKFRAME = &H40000
Private Const WS_CHILD = &H40000000
Private Const WS_POPUP = &H80000000
Private Const SWP_NOZORDER = &H4
Private Const SWP_FRAMECHANGED = &H20
Private Const SWP_NOOWNERZORDER = &H200
Private Const SWP_NOREPOSITION = SWP_NOOWNERZORDER

Private Const madLongVarCharDefault As Integer = 10          '�ַ����ֶ�ȱʡ����
Private Const madDoubleDefault As Integer = 18               '�������ֶ�ȱʡ����
Private Const madDbDateDefault As Integer = 20               '�������ֶ�ȱʡ����

Public Event AfterDataChanged(ByVal blnChange As Boolean)
Public Event AfterRefresh()
Public Event AfterRowColChange(ByVal strInfo As String, ByVal blnImportant As Boolean, ByVal blnSign As Boolean, ByVal blnArchive As Boolean)
Public Event AfterDataSave(ByVal blnSave As Boolean)
Public Event AfterFileIndex(ByVal lngFileIndex As Long)
Public Event AfterPartogramInfo(ByVal lngFlieId As Long, ByVal lngFileIndex As Long, ByVal lngFileFormatID As Long, ByVal rsPartogram As ADODB.Recordset)

Private mstrFields As String
Private mstrValues As String
Private mstrTag As String           '�ݴ�

'�����ļ���ʽ�������
Private mintTabTiers As Integer     '��ͷ���
Private mintTagFormHour As Integer  '��ʼʱ������
Private mintTagToHour As Integer    '��ֹʱ������
Private mobjTagFont As New StdFont  '������ʽ����
Private mlngTagColor As Long        '������ʽ��ɫ
Private mstrPaperSet As String      '��ʽ
Private mblnChildForm As Boolean
Private mstrSubHead As String       '���ϱ�ǩ
Private mstrSubEnd As String        '���±�ǩ
Private mstrTabHead As String       '��ͷ��Ԫ
Private mstrColWidth As String      '�п����д�
Private mstrColumns As String       '��ǰ�����ļ����ж�Ӧ����Ŀ
Private lngCurColor As Long, strCurFont As String, objFont As StdFont
'����򿪻����¼�ļ���SQL���������ط�Ҳ��ʹ�ã������޸�
Private mstrSQL�� As String
Private mstrSQL�� As String
Private mstrSQL�� As String
Private mstrSQL���� As String
Private mstrSQL As String

'######################################################################################################################
'**********************************************************************************************************************
'��#�ָ��������ڵĴ��붼���ͼ���,û�±�
Private Declare Function OleTranslateColor Lib "olepro32.dll" (ByVal OLE_COLOR As Long, ByVal HPALETTE As Long, pccolorref As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32" (ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function GetStockObject Lib "gdi32" (ByVal nIndex As Long) As Long
Private Declare Function SetTextColor Lib "gdi32" (ByVal hDC As Long, ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hDC As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, lpPoint As POINTAPI) As Long

Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Const WHITE_BRUSH = 0    '��ɫ����
Private Const cdblWidth As Double = 6          'һ��Ӣ���ַ��Ŀ��
Private Const cHideCols = 3         'ǰ׺������:����,ʱ��,ѡ��
Private Const cControlFields = 2    '��¼��������:ҳ��,�к�

Private Function GetRBGFromOLEColor(ByVal dwOleColour As Long) As Long
    '��VB����ɫת��ΪRGB��ʾ
    Dim clrref As Long
    Dim r As Long, g As Long, b As Long
    
    OleTranslateColor dwOleColour, 0, clrref
    
    b = (clrref \ 65536) And &HFF
    g = (clrref \ 256) And &HFF
    r = clrref And &HFF
    
    GetRBGFromOLEColor = RGB(r, g, b)
End Function

Private Function GetSymbolWidth(ByVal strPara As String) As Double
    'ȱʡ������9��,�������Сͬ�ȷŴ�
    Dim sinFontSize As Single
    Dim i As Integer, j As Integer
    
    j = Len(strPara)
    sinFontSize = VsfData.FontSize
    For i = 1 To j
        GetSymbolWidth = GetSymbolWidth + IIf(Asc(Mid(strPara, i, 1)) > 0, 1, 2) * cdblWidth * sinFontSize / 9
    Next
End Function

Private Sub DrawCell(ByVal hDC As Long, ByVal ROW As Long, ByVal COL As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Dim strText As String
    Dim strLeft As String
    Dim strRight As String
    Dim lngLeft As Long
    Dim lngRight As Long
    Dim dblWidth As Double
    Dim lngBackColor As Long
    Dim lngForeColor As Long
    Dim blnDraw As Boolean
    '��ͼ���
    Dim lngPen As Long
    Dim lngOldPen As Long
    Dim lngBrush As Long
    Dim lngOldBrush As Long
    Dim lpPoint As POINTAPI
    Dim T_ClientRect As RECT
    On Error GoTo errHand
    '******************************************
    '�ڴ��¼��в��ܶԵ�Ԫ����κ����Ը�ֵ,����Celldata,�����������¼�����ѭ��,���¹��������ʱ���޷�����������
    '******************************************
    'ʹ��ƥ��ı���ɫ��ǰ��ɫ����������ı������
    If Not mblnInit Then Exit Sub
    If VsfData.RowHidden(ROW) Then Exit Sub
    Done = False
    
    strText = VsfData.TextMatrix(ROW, COL)
    If IsDiagonal(COL) And InStr(1, strText, "/") <> 0 Then
        blnDraw = True
        '����ֵ
        strLeft = Split(strText, "/")(0)
        strRight = Mid(strText, InStr(1, strText, "/") + 1)
        lngLeft = LenB(StrConv(strLeft, vbFromUnicode))
        lngRight = LenB(StrConv(strRight, vbFromUnicode))
        'ȡ�ַ����
        dblWidth = GetSymbolWidth(strRight)
        '�趨�ͻ������С
        With T_ClientRect
            .Left = Left + 1
            .Top = Top + 1
            .Right = Right - 1
            .Bottom = Bottom - 1
        End With
        
        '1���������
        '�����뱳��ɫ��ͬ��ˢ��
        If ROW < VsfData.FixedRows Then
            lngBackColor = GetRBGFromOLEColor(VsfData.BackColorFixed)
            lngForeColor = GetRBGFromOLEColor(VsfData.ForeColorFixed)
        Else
            If ROW = VsfData.RowSel Then
                lngBackColor = GetRBGFromOLEColor(VsfData.BackColorSel)
                lngForeColor = RGB(0, 0, 0)
            Else
                lngBackColor = RGB(255, 255, 255)
                lngForeColor = GetRBGFromOLEColor(VsfData.Cell(flexcpForeColor, ROW, COL))
            End If

        End If
        lngBrush = CreateSolidBrush(lngBackColor)
        'ʹ�ø�ˢ����䱳��ɫ
        lngOldBrush = SelectObject(hDC, lngBrush)
        Call FillRect(hDC, T_ClientRect, lngBrush)
        '����������ʱʹ�õ�ˢ�Ӳ���ԭˢ��
        Call SelectObject(hDC, lngOldBrush)
        Call DeleteObject(lngBrush)
        
        '2��׼������
        '�����»���
        Call SetTextColor(hDC, lngForeColor)
        lngPen = CreatePen(0, 1, lngForeColor)
        lngOldPen = SelectObject(hDC, lngPen)
        '����
        Call MoveToEx(hDC, Left, Bottom - 2, lpPoint)
        Call LineTo(hDC, Right, Top)
        '����ı�
        Call TextOut(hDC, Left, Top, strLeft, lngLeft)
        Call TextOut(hDC, IIf(Right - dblWidth >= Left, Right - dblWidth, Left), Bottom - 16, strRight, lngRight)
        
        '��ԭ���ʲ�����
        Call SelectObject(hDC, lngOldPen)
        Call DeleteObject(lngPen)
        
        '�������ͼ
        Done = True
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub
'######################################################################################################################
'**********************************************************************************************************************
'��#�ָ��������ڵĴ��붼��������,û�±�
Private Function GetData(ByVal strInput As String) As Variant
    Dim arrData
    Dim strData As String
    Dim strLine(256) As Byte
    Dim lngRow As Long, lngRows As Long, lngLen As Long
    
    GetData = ""
    lngRows = SendMessage(txtLength.Hwnd, EM_GETLINECOUNT, 0&, 0&)
    For lngRow = 1 To lngRows
        Call ClearArray(strLine)
        lngLen = SendMessage(txtLength.Hwnd, EM_GETLINE, lngRow - 1, strLine(0))
        Call ClearArray(strLine, lngLen)
        strData = StrConv(strLine, vbUnicode)
        strData = TruncZero(strData)
        GetData = GetData & IIf(GetData = "", "", "|LPF.ZLSOFT|") & strData & IIf(lngRow < lngRows, vbCrLf, "")
    Next
    GetData = Split(GetData, "|LPF.ZLSOFT|")
End Function

Private Sub ClearArray(strLine() As Byte, Optional ByVal lngPos As Long = 0)
    Dim intDo As Integer, intMax As Integer
    intMax = UBound(strLine)
    For intDo = lngPos To intMax
        strLine(intDo) = 0
        If lngPos > 0 Then Exit Sub     '��Ϊ��,��ʾ�������ַ���������
    Next
    strLine(1) = 1
End Sub

Private Function TrimStr(ByVal str As String) As String
'���ܣ�ȥ���ַ�����\0�Ժ���ַ�������ȥ�����˵Ŀո�

    If InStr(str, Chr(0)) > 0 Then
        TrimStr = Trim(Left(str, InStr(str, Chr(0)) - 1))
    Else
        TrimStr = Trim(str)
    End If
End Function

Private Function TruncZero(ByVal strInput As String) As String
'���ܣ�ȥ���ַ�����\0�Ժ���ַ�
    Dim lngPos As Long
    
    lngPos = InStr(strInput, Chr(0))
    If lngPos > 0 Then
        TruncZero = Mid(strInput, 1, lngPos - 1)
    Else
        TruncZero = strInput
    End If
End Function

'**********************************************************************************************************************
'######################################################################################################################

Private Sub LoadBabyNum()
    Dim i As Integer
    On Error GoTo errHand
    
    '���ÿؼ����꣨��߻��ұ߳�����Ļ��С���һ�����ʾ����������Ϊ������ʾ��
    With picBaby
        .Left = VsfData.Left
        .Top = 0
        If .Height + .Top + picMain.Top > ScaleHeight Then
            .Top = ScaleHeight - picMain.Top - .Height
        End If
        If .Left + .Width > ScaleWidth Then
            .Left = ScaleWidth - .Width
        End If
        If .Left < VsfData.Left Then
            .Left = VsfData.Left
        End If
        If cboBaby.ListCount > 0 Then
            .Visible = True
            .ZOrder 0
        Else
            .Visible = False
            RaiseEvent AfterRowColChange("���ٴ���һ��Ӥ�������뿪������ϵ��", True, mblnSign, mblnArchive)
        End If
    End With
    
    '����Ӥ��������Ϣ
    With vsfBaby
        .FixedCols = 0
        .FixedRows = 0
        .Rows = cboBaby.ListCount
        For i = 0 To cboBaby.ListCount - 1
            .RowData(i) = cboBaby.ItemData(i)
            .TextMatrix(i, 0) = "Ӥ��" & .RowData(i)
        Next i
        .FocusRect = flexFocusHeavy
        .COL = .FixedCols: .ROW = .Rows - 1
        Call vsfBaby_AfterRowColChange(.FixedRows, .FixedCols, .ROW, .COL)
    End With
     
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function GetPeriod() As String
    On Error GoTo errHand
    gstrSQL = " Select   ��Ժ���� AS ��ʼʱ�� From ������ҳ Where ����ID=[1] And ��ҳID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��Ժ���ڻ��������", mlng����ID, mlng��ҳID)
    GetPeriod = Format(rsTemp!��ʼʱ��, "yyyy-MM-dd HH:mm:ss") & "��" & Format(mstr����ʱ��, "yyyy-MM-dd HH:mm:ss")
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ReadStruDef() As Boolean
    Dim lngCol As Long
    Dim strSQLHead As String, strSQLRow As String
    On Error GoTo errHand
    
    '��ȡ�ļ�����
    mstrCOLNothing = ""
    mblnDateAd = False
    mblnDate = False
    Call GetFileProperty
    
    '��ȡ�����ļ���ʽ����
    gstrSQL = "Select   d.�������, d.�����ı�, d.Ҫ������" & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '�����ʽ'" & _
        " Order By d.�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ļ���ʽ����", mlng��ʽID)
    With rsTemp
        Do While Not .EOF
            Select Case "" & !Ҫ������
            Case "��ͷ����": mintTabTiers = Val("" & !�����ı�)
            Case "������"
                VsfData.Cols = Val("" & !�����ı�)
                vsfHead.Cols = VsfData.Cols
            Case "��С�и�"
                VsfData.RowHeightMin = BlowUp(Val("" & !�����ı�))
                vsfHead.RowHeightMin = VsfData.RowHeightMin
            Case "�ı�����"
                strCurFont = "" & !�����ı�
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = BlowUp(Val(Split(strCurFont, ",")(1)))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "��") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "б") > 0 Then .Italic = True
                End With
                Set VsfData.Font = objFont
                Set vsfHead.Font = objFont
                Set lblSubhead.Font = VsfData.Font
                Set lblSubEnd.Font = VsfData.Font
                Set Font = lblSubhead.Font
                
            Case "�ı���ɫ"
                VsfData.ForeColor = Val("" & !�����ı�)
                vsfHead.ForeColor = VsfData.ForeColor
            Case "�����ɫ"
                VsfData.GridColor = Val("" & !�����ı�): VsfData.GridColorFixed = VsfData.GridColor
                vsfHead.GridColor = VsfData.GridColor: vsfHead.GridColorFixed = VsfData.GridColorFixed
            Case "�����ı�"
                lblTitle.Caption = "" & !�����ı�
                lblTitle.AutoSize = True
            Case "��������"
                strCurFont = "" & !�����ı�
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = BlowUp(Val(Split(strCurFont, ",")(1)))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "��") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "б") > 0 Then .Italic = True
                End With
                Set lblTitle.Font = objFont
                lblTitle.AutoSize = False
            
            Case "��ʼʱ��": mintTagFormHour = Val("" & !�����ı�)
            Case "��ֹʱ��": mintTagToHour = Val("" & !�����ı�)
            Case "��������"
                strCurFont = "" & !�����ı�
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = BlowUp(Val(Split(strCurFont, ",")(1)))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "��") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "б") > 0 Then .Italic = True
                End With
                Set mobjTagFont = objFont
            Case "������ɫ"
                mlngTagColor = Val("" & !�����ı�)
            End Select
            .MoveNext
        Loop
    End With
    
    gstrSQL = "Select ��ʽ From ����ҳ���ʽ Where ���� = 3 And ��� In (Select ҳ�� From �����ļ��б� Where Id = [1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����ҳ���ʽ", mlng��ʽID)
    If Not rsTemp.EOF Then
        mstrPaperSet = "" & rsTemp!��ʽ
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select   d.�������, d.�����ı�, d.Ҫ������, Nvl(d.�Ƿ���, 0) As �Ƿ���" & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '���ϱ�ǩ'" & _
        " Order By d.�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ϱ�ǩ����", mlng��ʽID)
    With rsTemp
        mstrSubHead = ""
        Do While Not .EOF
            mstrSubHead = mstrSubHead & "|" & IIf(!�Ƿ��� = 0, "", vbCrLf) & !�����ı� & "{" & !Ҫ������ & "}"
            .MoveNext
        Loop
        If mstrSubHead <> "" Then mstrSubHead = Replace(Mid(mstrSubHead, 2), Chr(1), " ")
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select   d.�������, d.�����ı�, d.Ҫ������, Nvl(d.�Ƿ���, 0) As �Ƿ���" & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '���±�ǩ'" & _
        " Order By d.�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���ϱ�ǩ����", mlng��ʽID)
    With rsTemp
        mstrSubEnd = ""
        Do While Not .EOF
            mstrSubEnd = mstrSubEnd & "|" & IIf(!�Ƿ��� = 0, "", vbCrLf) & !�����ı� & "{" & !Ҫ������ & "}"
            .MoveNext
        Loop
        If mstrSubEnd <> "" Then mstrSubEnd = Replace(Mid(mstrSubEnd, 2), Chr(1), " ")
    End With
    '------------------------------------------------------------------------------------------------------------------
    '����Ƿ����������
    gstrSQL = "Select  d.�������, d.��������, d.�����д�, d.�����ı�, d.Ҫ������, d.Ҫ�ص�λ,d.Ҫ�ر�ʾ" & vbNewLine & _
            "        From �����ļ��ṹ d, �����ļ��ṹ p" & vbNewLine & _
            "        Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '���м���'" & vbNewLine & _
            "        And D.Ҫ������='����'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���м��϶���", mlng��ʽID)
    mblnDate = (rsTemp.RecordCount > 0)
    If mblnDate = False Then
        VsfData.Cols = VsfData.Cols + 1
        vsfHead.Cols = VsfData.Cols
    End If
    '------------------------------------------------------------------------------------------------------------------
    '������������Ĭ�����������
    If mblnDate = False Then
        gstrSQL = "SELECT �������, �����д�, �����ı�" & vbNewLine & _
                "FROM (SELECT 1 �������, 1 �����д�, '����' �����ı�" & vbNewLine & _
                "       FROM DUAL" & vbNewLine & _
                "       UNION ALL" & vbNewLine & _
                "       SELECT D.�������+1 �������, D.�����д�, D.�����ı�" & vbNewLine & _
                "       FROM �����ļ��ṹ D, �����ļ��ṹ P" & vbNewLine & _
                "       WHERE P.ID = D.��ID AND P.�ļ�ID = [1] AND P.�������� = 1 AND P.�����ı� = '��ͷ��Ԫ')" & vbNewLine & _
                "ORDER BY �������"
    Else
        gstrSQL = "Select   d.�������, d.�����д�, d.�����ı�" & _
            " From �����ļ��ṹ d, �����ļ��ṹ p" & _
            " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '��ͷ��Ԫ'" & _
            " Order By d.�������"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��ͷ��Ԫ����", mlng��ʽID)
    With rsTemp
        mstrTabHead = ""
        Do While Not .EOF
            mstrTabHead = mstrTabHead & "|" & !�����д� - 1 & "," & !������� & "," & !�����ı�
            .MoveNext
        Loop
        If mstrTabHead <> "" Then mstrTabHead = Mid(mstrTabHead, 2)
    End With
    
    '��ѯ�����֯
    '------------------------------------------------------------------------------------------------------------------
    Dim strSql�� As String, str��ʽ As String, strSqlNull As String
    Dim bln���� As Boolean, blnʱ�� As Boolean, bln��ʿ As Boolean
    Dim blnǩ���� As Boolean, blnǩ��ʱ�� As Boolean, blnǩ������ As Boolean
    Dim bln�Խ��� As Boolean         '�����һ���ǶԽ�����ѡ����,��ֱ����ȡ��������,ƴ��ͷʱ����ֵ�����/
    Dim lngColumn As Long
      
    If mblnDate = False Then
        gstrSQL = "SELECT �������, ��������, �����д�, �����ı�, Ҫ������, Ҫ�ص�λ, Ҫ�ر�ʾ" & vbNewLine & _
            "FROM (SELECT 1 �������, '0`4' ��������, 1 �����д�, '' �����ı�, '����' Ҫ������, '' Ҫ�ص�λ, 0 Ҫ�ر�ʾ" & vbNewLine & _
            "       FROM DUAL" & vbNewLine & _
            "       UNION ALL" & vbNewLine & _
            "       SELECT D.�������+1 �������, D.��������, D.�����д�, D.�����ı�, D.Ҫ������, D.Ҫ�ص�λ, D.Ҫ�ر�ʾ" & vbNewLine & _
            "       FROM �����ļ��ṹ D, �����ļ��ṹ P" & vbNewLine & _
            "       WHERE P.ID = D.��ID AND P.�ļ�ID = [1] AND P.�������� = 1 AND P.�����ı� = '���м���')" & vbNewLine & _
            "ORDER BY �������, �����д�"
    Else
        gstrSQL = "Select   d.�������, d.��������, d.�����д�, d.�����ı�, d.Ҫ������, d.Ҫ�ص�λ,d.Ҫ�ر�ʾ " & _
            " From �����ļ��ṹ d, �����ļ��ṹ p" & _
            " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '���м���'" & _
            " Order By d.�������, d.�����д�"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ���м��϶���", mlng��ʽID)
    With rsTemp
        lngColumn = 0: mstrColumns = "": mstrColWidth = "": mstrCatercorner = ""
        mstrSQL�� = "": mstrSQL�� = "": strSql�� = "": mstrSQL�� = "": mstrSQL���� = "": strSqlNull = ""
        bln���� = False: blnʱ�� = False: bln��ʿ = False
        blnǩ���� = False: blnǩ��ʱ�� = False: blnǩ������ = False
        Do While Not .EOF
            If lngColumn <> !������� Then
                mstrColumns = mstrColumns & IIf(mstrColumns = "", "", "'1'" & str��ʽ) & "|" & !������� & "'" & !Ҫ������
                mstrColWidth = mstrColWidth & "," & !�������� & "`" & !������� & "`" & !Ҫ�ر�ʾ
                If !Ҫ�ر�ʾ = 1 Then mstrCatercorner = mstrCatercorner & "," & !�������
                str��ʽ = ""
                If !Ҫ������ <> "" Then str��ʽ = "{" & NVL(!�����ı�) & "[" & !Ҫ������ & "]" & NVL(!Ҫ�ص�λ) & "}"
                
                If Mid(strSqlNull, 3) = "" Then
                    strSqlNull = "''"
                Else
                    strSqlNull = Mid(strSqlNull, 3)
                End If
                mstrSQL�� = mstrSQL�� & "," & IIf(Mid(strSql��, 3) = "", "''", "Decode(" & Mid(strSql��, 3) & "," & strSqlNull & ",''," & Mid(strSql��, 3) & ")") & " As C" & Format(lngColumn, "00")
                
                strSql�� = ""
                strSqlNull = ""
                lngColumn = !�������
                bln�Խ��� = (NVL(!Ҫ�ر�ʾ, 0) = 1)
            Else
                mstrColumns = mstrColumns & "," & !Ҫ������
                str��ʽ = str��ʽ & "{" & NVL(!�����ı�) & "[" & !Ҫ������ & "]" & NVL(!Ҫ�ص�λ) & "}"
            End If
            
            Select Case !Ҫ������
            Case "����"
                bln���� = True
                mblnDateAd = (NVL(!Ҫ�ر�ʾ, 0) = 1)
                mstrSQL�� = mstrSQL�� & ",����"
                mstrSQL�� = mstrSQL�� & ",To_Char(l.����ʱ��, " & IIf(mblnDateAd, "'dd/MM'", "'yyyy-mm-dd'") & ") As ����"
                strSql�� = strSql�� & "||" & !Ҫ������
            Case "ʱ��"
                blnʱ�� = True
                mstrSQL�� = mstrSQL�� & ",ʱ��"
                mstrSQL�� = mstrSQL�� & ",To_Char(l.����ʱ��, 'hh24:mi') As ʱ��"
                strSql�� = strSql�� & "||" & !Ҫ������
                
            Case "ǩ����"
                blnǩ���� = True
                mstrSQL�� = mstrSQL�� & ",ǩ����"
                mstrSQL�� = mstrSQL�� & ",l.ǩ����"
                strSql�� = strSql�� & "||" & !Ҫ������
                
            Case "ǩ��ʱ��"
                blnǩ��ʱ�� = True
                mstrSQL�� = mstrSQL�� & ",ǩ��ʱ��"
                mstrSQL�� = mstrSQL�� & ",l.ǩ��ʱ��"
                strSql�� = strSql�� & "||" & !Ҫ������
                
            Case "��ʿ"
                bln��ʿ = True
                mstrSQL�� = mstrSQL�� & ",��ʿ"
                mstrSQL�� = mstrSQL�� & ",l.������ As ��ʿ"
                strSql�� = strSql�� & "||" & !Ҫ������
            Case Else
                If !Ҫ������ <> "" Then
                    mstrSQL�� = mstrSQL�� & ",Max(""" & !Ҫ������ & """) As """ & !Ҫ������ & """"
                    mstrSQL���� = mstrSQL���� & " Or """ & !Ҫ������ & """ Is Not Null"
                    
                    strSql�� = strSql�� & "||'" & !�����ı� & "'||""" & !Ҫ������ & """||'" & !Ҫ�ص�λ & "'"
                    strSqlNull = strSqlNull & "||" & "'" & !�����ı� & "'||'" & !Ҫ�ص�λ & "'"
                    mstrSQL�� = mstrSQL�� & ", Decode(c.��Ŀ����, '" & !Ҫ������ & "', Nvl(c.δ��˵��,c.��¼����), '') As """ & !Ҫ������ & """"

                Else
                    mstrCOLNothing = mstrCOLNothing & "," & Val(Format(!�������, "00"))
                End If
            End Select
            .MoveNext
        Loop
        
        mstrCOLNothing = Mid(mstrCOLNothing, 2)
        mstrCatercorner = Mid(mstrCatercorner, 2)
        mstrColWidth = Mid(mstrColWidth, 2)
        '�������һ�еĸ�ʽ
        mstrColumns = mstrColumns & IIf(mstrColumns = "", "", "'1'" & str��ʽ) '& "|" & !������� & "'" & !Ҫ������
        mstrColumns = Mid(mstrColumns, 2)     '��ʽ��:�к�;��Ŀ����1,��Ŀ����2|�к�...,ʵ��;1;����|2;����|3...
        If Mid(strSql��, 3) <> "" Then
            mstrSQL�� = mstrSQL�� & "," & Mid(strSql��, 3) & " As C" & Format(lngColumn, "00")
        Else
            mstrSQL�� = mstrSQL�� & ",'' As C" & Format(lngColumn, "00")
        End If
        
        If mstrSQL���� <> "" Then mstrSQL���� = "(" & Mid(mstrSQL����, 5) & ")"
        
        '���û�г������ڣ�ʱ�䣬��ʿ�����ڲ���Ҫ���䣬�Ա�֤�в�����������
        If bln���� = False Then mstrSQL�� = mstrSQL�� & ",To_Char(l.����ʱ��, 'yyyy-mm-dd') As ����"
        If blnʱ�� = False Then mstrSQL�� = mstrSQL�� & ",To_Char(l.����ʱ��, 'hh24:mi') As ʱ��"
        
        If blnǩ���� = False Then mstrSQL�� = mstrSQL�� & ",l.ǩ���� As ǩ����"
        If blnǩ��ʱ�� = False Then mstrSQL�� = mstrSQL�� & ",l.ǩ��ʱ��"
        
        If Mid(mstrSQL��, 2) = "" Then
            MsgBox "�Բ�����û�ж��嵱ǰ��������ʾ����Ϣ�����ڲ����ļ������ж��壡", vbInformation, gstrSysName
            Exit Function
        End If
        
        '�����ڲ��������ӹ̶���
        mstrSQL�� = mstrSQL�� & ",MAX(ǩ������) AS ǩ������,MAX(ǩ����Ϣ) AS ǩ����Ϣ,MAX(��¼ID) AS ��¼ID,MAX(����) AS ����"
        mstrSQL�� = mstrSQL�� & ",l.ǩ������,l.ǩ���� AS ǩ����Ϣ,C.��¼ID,'' AS ����"
        mstrSQL�� = mstrSQL�� & ",ǩ������,ǩ����Ϣ,��¼ID,����"
        
        If bln��ʿ = False Then
            'ǿ����ӻ�ʿ��,Ϊ�˱����޸�����������(����¼�������,����������Ҳ������)
            mstrSQL�� = mstrSQL�� & ",��ʿ"
            mstrSQL�� = mstrSQL�� & ",l.������ As ��ʿ"
            mstrSQL�� = mstrSQL�� & ",��ʿ"
        End If
        
        '�����Ŀ���뵽SQL��
        Call SQLCombination
    End With
    
    ReadStruDef = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub SQLCombination(Optional ByVal lng��¼ID As Long = 0)
    Dim str���� As String
    str���� = mstrSQL���� & IIf(lng��¼ID = 0, "", IIf(mstrSQL���� = "", "", " And") & " ��¼ID=[6]")
    
    mstrSQL = "Select '' ����,to_char(����ʱ��,'yyyy-MM-dd hh24:mi:ss') AS ����ʱ��,'' AS ѡ��," & Mid(mstrSQL��, 12) & vbCrLf & _
                " From (Select ����ʱ��," & Mid(mstrSQL��, 2) & vbCrLf & _
                "        From (Select l.����ʱ��," & Mid(mstrSQL��, 2) & vbCrLf & _
                "               From ���˻������� l, ���˻�����ϸ c,���˻����ļ� f " & vbCrLf & _
                "               Where l.Id = c.��¼id And l.�ļ�ID=f.ID" & _
                "               And c.��ֹ�汾 Is Null And c.��¼����<>5  " & _
                "               And f.id=[1] And f.����id = [2] And f.��ҳid = [3] And Nvl(f.Ӥ��,0)=[4] And l.�������=[5])" & vbCrLf & _
                IIf(str���� <> "", "Where " & str����, "") & _
                "       Group By ����, ʱ��, ����ʱ��,��ʿ,ǩ����,ǩ��ʱ��" & _
                                "       Order By ����ʱ��,��ʿ,ǩ����,ǩ��ʱ��)"
End Sub

Public Sub zlRefresh(Optional ByVal blnRefresh As Boolean = True)
'-----------------------------------------------------------------------
'���ܣ��������ˢ��,blnRefresh=false ��ʾֻˢ�±��Ϻͱ�ǩ������Ϣ
'-----------------------------------------------------------------------
    Dim aryRow() As String, aryItem() As String, arrItemEnd() As String
    Dim strPrefix As String, strItemName As String
    Dim lngRow As Long, lngCol As Long, lngCount As Long, strCell As String, i As Integer
    Dim strTmpSQL As String
    Dim aryPeriod() As String
    Dim strtmp As String, str��λ As String
    
    Err = 0: On Error GoTo errHand
    'Debug.Print Now & "zlRefresh"
    Call InitCons
    '���ϱ�ǩ��ȡ
    lblSubhead.Caption = ""
    lblSubhead.Tag = ""
    lblSubEnd.Caption = ""
    lblSubEnd.Tag = ""
    aryPeriod = Split(GetPeriod, "��")
    gstrSQL = "Select [1] || Zl_Replace_Element_Value([2],[3],[4],2,NULL,[5]) as ��Ϣ From Dual"
    aryItem = Split(mstrSubHead, "|")
    arrItemEnd = Split(mstrSubEnd, "|")
    For i = 0 To 1
        For lngCount = 0 To IIf(i = 0, UBound(aryItem), UBound(arrItemEnd))
            If i = 0 Then
                strPrefix = Left(aryItem(lngCount), InStr(1, aryItem(lngCount), "{") - 1)
                strItemName = Mid(aryItem(lngCount), InStr(1, aryItem(lngCount), "{") + 1, InStr(1, aryItem(lngCount), "}") - InStr(1, aryItem(lngCount), "{") - 1)
            Else
                strPrefix = Left(arrItemEnd(lngCount), InStr(1, arrItemEnd(lngCount), "{") - 1)
                strItemName = Mid(arrItemEnd(lngCount), InStr(1, arrItemEnd(lngCount), "{") + 1, InStr(1, arrItemEnd(lngCount), "}") - InStr(1, arrItemEnd(lngCount), "{") - 1)
            End If
            mrsPartogram.Filter = 0
            mrsPartogram.Filter = "������='" & strItemName & "'"
            '�������Ҳ����������ֹ��޸�����
            If mrsPartogram.RecordCount = 0 Then GoTo ErrNext
            str��λ = Trim(NVL(mrsPartogram!��λ))
            If Val(NVL(mrsPartogram!�滻��)) = 1 Then
                '���̶̹�Ҫ����Ϣ
                strtmp = strPrefix
                Select Case strItemName
                Case "��ǰ����"
                
                    strTmpSQL = "Select   b.����" & vbNewLine & _
                                "From (Select ����id, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                                "            From ���˱䶯��¼" & vbNewLine & _
                                "            Where ����id = [1] And ��ҳid = [2] And ����id = [3]) a,���ű� b " & vbNewLine & _
                                "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [4] >= a.��ʼʱ��) And a.����id Is Not Null And b.ID=a.����id" & vbNewLine & _
                                "Order By a.��ʼʱ��"
                                
                    Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "��ǰ����", mlng����ID, mlng��ҳID, mlng����id, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
                    
                Case "��ǰ����"
                
                    strTmpSQL = "Select   a.����" & vbNewLine & _
                                "From (Select ����, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                                "            From ���˱䶯��¼" & vbNewLine & _
                                "            Where ����id = [1] And ��ҳid = [2] And ����id = [3]) a" & vbNewLine & _
                                "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [4] >= a.��ʼʱ��) And a.���� Is Not Null" & vbNewLine & _
                                "Order By a.��ʼʱ��"
        
                    Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "��ǰ����", mlng����ID, mlng��ҳID, mlng����id, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
                    If rsTemp.BOF = False Then rsTemp.MoveLast
                    
                Case "��ǰ����"
                
                    strTmpSQL = "Select   ���� From ���ű� a Where a.ID=[1]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "��ǰ����", mlng����id)
                    
                Case "סԺҽʦ"
                    strTmpSQL = "Select   a.����ҽʦ" & vbNewLine & _
                                "From (Select ����ҽʦ, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                                "            From ���˱䶯��¼" & vbNewLine & _
                                "            Where ����id = [1] And ��ҳid = [2] And ����id = [3]) a" & vbNewLine & _
                                "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [4] >= a.��ʼʱ��) And a.����ҽʦ Is Not Null" & vbNewLine & _
                                "Order By a.��ʼʱ��"
                    Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "סԺҽʦ", mlng����ID, mlng��ҳID, mlng����id, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
                    If rsTemp.BOF = False Then rsTemp.MoveLast
                Case "���λ�ʿ"
                
                    strTmpSQL = "Select   a.���λ�ʿ" & vbNewLine & _
                                "From (Select ���λ�ʿ, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                                "            From ���˱䶯��¼" & vbNewLine & _
                                "            Where ����id = [1] And ��ҳid = [2] And ����id = [3]) a" & vbNewLine & _
                                "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [4] >= a.��ʼʱ��) And a.���λ�ʿ Is Not Null" & vbNewLine & _
                                "Order By a.��ʼʱ��"
                    Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "���λ�ʿ", mlng����ID, mlng��ҳID, mlng����id, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
                    If rsTemp.BOF = False Then rsTemp.MoveLast
                    
                Case "����ȼ�"
                    strTmpSQL = "Select   b.����" & vbNewLine & _
                                "From (Select ����ȼ�ID, ��ʼʱ��, Nvl(��ֹʱ��, To_Date('3000-01-01', 'yyyy-mm-dd')) As ��ֹʱ��" & vbNewLine & _
                                "            From ���˱䶯��¼" & vbNewLine & _
                                "            Where ����id = [1] And ��ҳid = [2] And ����id = [3]) a,����ȼ� b" & vbNewLine & _
                                "Where ([4] Between a.��ʼʱ�� And a.��ֹʱ�� Or [4] >= a.��ʼʱ��) And a.����ȼ�ID Is Not Null And b.���=a.����ȼ�ID" & vbNewLine & _
                                "Order By a.��ʼʱ��"
                    Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "����ȼ�", mlng����ID, mlng��ҳID, mlng����id, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
                    If rsTemp.BOF = False Then rsTemp.MoveLast
                Case "������"
                    strtmp = ""
                    gstrSQL = "Select [1] || Zl_Replace_Element_Value([2],[3],[4],2,NULL,[5],[6]) as ��Ϣ From Dual"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡҪ��", strPrefix, strItemName, mlng����ID, mlng��ҳID, mintӤ��, CDate(aryPeriod(0)))
                Case Else
                    strtmp = ""
                    gstrSQL = "Select [1] || Zl_Replace_Element_Value([2],[3],[4],2,NULL,[5]) as ��Ϣ From Dual"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡҪ��", strPrefix, strItemName, mlng����ID, mlng��ҳID, mintӤ��)
                End Select
            Else
                '����¼��Ҫ����Ϣ
                strtmp = strPrefix
                gstrSQL = "SELECT ���� From ����Ҫ������" & _
                    "   Where �ļ�ID = [1] And Ӥ�� = [2] And ���� =[3]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȥҪ��", mlng�ļ�ID, mlngFileIndex, strItemName)
            End If
            If rsTemp.BOF = False Then
                If i = 0 Then
                    If strtmp <> "" Then
                        lblSubhead.Tag = lblSubhead.Tag & " " & strtmp & rsTemp.Fields(0).Value & str��λ
                    Else
                        lblSubhead.Tag = lblSubhead.Tag & " " & rsTemp.Fields(0).Value & str��λ
                    End If
                Else
                    If strtmp <> "" Then
                        lblSubEnd.Tag = lblSubEnd.Tag & " " & strtmp & rsTemp.Fields(0).Value & str��λ
                    Else
                        lblSubEnd.Tag = lblSubEnd.Tag & " " & rsTemp.Fields(0).Value & str��λ
                    End If
                End If
            Else
            If i = 0 Then
                If strtmp <> "" Then
                        lblSubhead.Tag = lblSubhead.Tag & " " & strtmp
                    Else
                        lblSubhead.Tag = lblSubhead.Tag & " "
                    End If
                Else
                    If strtmp <> "" Then
                        lblSubEnd.Tag = lblSubEnd.Tag & " " & strtmp
                    Else
                        lblSubEnd.Tag = lblSubEnd.Tag & " "
                    End If
                End If
            End If
ErrNext:
        Next
    Next i
    lblSubhead.Tag = Trim(lblSubhead.Tag)
    lblSubEnd.Tag = Trim(lblSubEnd.Tag)
    '���ϱ�ǩ��ɢ����
    Call zlLableBruit
    '�����м�¼��
    Call InitRecords
    
    If blnRefresh = False Then Exit Sub
    'װ������
    Call SQLCombination
    gstrSQL = mstrSQL
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������", mlng�ļ�ID, mlng����ID, mlng��ҳID, mintӤ��, cboBaby.ItemData(cboBaby.ListIndex))
    '�����������¼���ṹ
    Call DataMap_Init(rsTemp)
    '�����ݲ����ø�ʽ,ͬʱʵ��һ�����ݷ�����ʾ�Ĺ���
    Call PreTendFormat(rsTemp)
    
    lblCurPage.Caption = ""
    Exit Sub

errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub DataMap_Init(ByVal rsSource As ADODB.Recordset)
    '��ʼ���ڴ����ݼ�
    
    '���ݼ�¼��,���ڿ��ٻָ�
    Set mrsDataMap = CopyNewRec(rsSource)
    mrsDataMap.Sort = "ҳ��,�к�"
    '�޸ĵ�Ԫ���¼,���ڱ���
    Call Record_Init(mrsCellMap, "ID," & adLongVarChar & ",50|ҳ��," & adDouble & ",18|�к�," & adDouble & ",18|" & _
            "�к�," & adDouble & ",18|��¼ID," & adDouble & ",18|����," & adLongVarChar & ",4000|��λ," & adLongVarChar & ",100|" & _
            "����," & adDouble & ",1|ɾ��," & adDouble & ",1")
    mrsCellMap.Sort = "ҳ��,�к�,�к�"
    '���Ƽ�¼��
    Set mrsCopyMap = New ADODB.Recordset
    Set mrsCopyMap = CopyNewRec(mrsDataMap, False)
End Sub

Private Function DataMap_Save() As Boolean
    '����ǰҳ�����û��༭�������ݱ�������,ҳ���л��򱣴�ǰ����
    Dim blnExit As Boolean
    Dim lngRow As Long, lngCol As Long, lngRows As Long, lngCols As Long
    Dim strDate As String, strTime As String, strDatetime As String, strCurrDate As String
    On Error GoTo errHand
    
    '�����Ƿ�༭��������
    If Not CheckFlip Then Exit Function
    
    '��ɾ��ָ��ҳ�ŵ�����������
    mrsDataMap.Filter = "ҳ��=" & mintҳ��
    Do While True
        If mrsDataMap.RecordCount = 0 Then Exit Do
        mrsDataMap.Delete
        mrsDataMap.MoveNext
    Loop
    mrsDataMap.Filter = 0
    
    strCurrDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm")
    '����ָ��ҳ�ŵ�����������
    lngRows = VsfData.Rows - 1
    lngCols = VsfData.Cols - 1
    For lngRow = VsfData.FixedRows To lngRows
        mrsDataMap.AddNew
        mrsDataMap!ҳ�� = mintҳ��
        mrsDataMap!�к� = lngRow
        mrsDataMap!ɾ�� = IIf(VsfData.RowHidden(lngRow), 1, 0)
        strDatetime = ""
        For lngCol = 0 To lngCols - VsfData.FixedCols
            If lngCol + VsfData.FixedCols = mlngChoose Then
                mrsDataMap.Fields(cControlFields + lngCol).Value = VsfData.Cell(flexcpChecked, lngRow, mlngChoose)
            ElseIf InStr(1, "," & mlngRecord & ",", "," & lngCol + VsfData.FixedCols & ",") <> 0 Then
                mrsDataMap.Fields(cControlFields + lngCol).Value = Val(VsfData.TextMatrix(lngRow, lngCol + VsfData.FixedCols))
            Else
                mrsDataMap.Fields(cControlFields + lngCol).Value = IIf(VsfData.TextMatrix(lngRow, lngCol + VsfData.FixedCols) = "", Null, VsfData.TextMatrix(lngRow, lngCol + VsfData.FixedCols))
            End If
            
            If lngCol + VsfData.FixedCols = mlngDate Then
                  strDate = Trim(VsfData.TextMatrix(lngRow, lngCol + VsfData.FixedCols))
                  If strDate <> "" Then
                      If mblnDateAd Then
                          strDate = Mid(strCurrDate, 1, 5) & ToStandDate(strDate)
                      Else
                          strDate = Format(strDate, "yyyy-MM-dd")
                      End If
                  End If
            ElseIf lngCol + VsfData.FixedCols = mlngTime Then
                  strTime = Trim(VsfData.TextMatrix(lngRow, lngCol + VsfData.FixedCols))
            End If
        Next
        If strDate <> "" And strTime <> "" Then
            strDatetime = strDate & " " & strTime & ":00"
            If mblnDateAd Then
                strDatetime = GetDateAdCurrDate(strDatetime)
            End If
            strDatetime = Format(strDatetime, "YYYY-MM-DD HH:mm:ss")
        End If
        mrsDataMap!�������� = Format(strDatetime, "YYYY-MM-DD HH:mm:ss")
        mrsDataMap.Update
    Next
    
    DataMap_Save = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub CellMap_Update(ByVal lngStart As Long, ByVal lngDeff As Long)
    Dim lngPos As Long
    Dim intCOl As Integer
    
    '���µ�ǰҳ�����д�����ʼ�е��к�����
    With mrsCellMap
        If .RecordCount <> 0 Then .MoveLast
        If .BOF Then Exit Sub
        Do While Not .BOF
            If !ҳ�� = mintҳ�� And !�к� > lngStart Then
                intCOl = !�к�
                lngPos = .AbsolutePosition
                !�к� = !�к� + lngDeff
                !ID = mintҳ�� & "," & !�к� & "," & !�к�
                .Update
                .MoveFirst
                .Move lngPos - 2
            Else
                .MovePrevious
            End If
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
End Sub

Public Function CopyNewRec(ByVal rsSource As ADODB.Recordset, Optional ByVal blnAddPage As Boolean = True) As ADODB.Recordset
    'ֻ������¼���Ľṹ,ͬʱ����ҳ��,�к��ֶ�
    Dim rsTarget As New ADODB.Recordset
    Dim intFields As Integer
    
    Set rsTarget = New ADODB.Recordset
    With rsTarget
        If blnAddPage Then
            .Fields.Append "ҳ��", adDouble, 18
            .Fields.Append "�к�", adDouble, 18
        End If
        For intFields = 0 To rsSource.Fields.Count - 1
            If rsSource.Fields(intFields).Name = "��������" Then
                .Fields.Append rsSource.Fields(intFields).Name, adLongVarChar, 50, adFldIsNullable      '0:��ʾ����
            ElseIf rsSource.Fields(intFields).Type = 200 Then       '�����ʹ���Ϊ�ַ���
                .Fields.Append rsSource.Fields(intFields).Name, adLongVarChar, rsSource.Fields(intFields).DefinedSize, adFldIsNullable     '0:��ʾ����
            Else
                .Fields.Append rsSource.Fields(intFields).Name, IIf(rsSource.Fields(intFields).Type = adNumeric, adDouble, rsSource.Fields(intFields).Type), rsSource.Fields(intFields).DefinedSize, adFldIsNullable    '0:��ʾ����
            End If
        Next
        If blnAddPage Then
            .Fields.Append "ɾ��", adDouble, 1
            .Fields.Append "��������", adLongVarChar, 20, adFldIsNullable
        End If
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
    
    Set CopyNewRec = rsTarget
End Function

Private Sub PreTendMutilRows()
    Dim lngRow As Long, lngCol As Long, lngMaxRowCount As Long
    Dim arrData
    Dim intData As Integer, intDatas As Integer
    
    On Error GoTo errHand
    '���һ����ʾ�����������ʾ(���ݵ�ǰʵ�����ݼ���ռ������,�����������Ӻ͸�ֵ,Ȼ���������������ݴ���)
    lngRow = VsfData.FixedRows
    Do While True
        If lngRow > VsfData.Rows - 1 Then Exit Do
        '���������������,�����������Ӻ͸�ֵ
        lngMaxRowCount = 1
        For lngCol = mlngTime + 1 To mlngNoEditor - 1
            If VsfData.ColHidden(lngCol) = False And lngCol <> mlngRowCount Then
                With txtLength
                   .Width = VsfData.ColWidth(lngCol)
                   .Text = VsfData.TextMatrix(lngRow, lngCol)
                   .FontName = VsfData.CellFontName
                   .FontSize = VsfData.CellFontSize
                   .FontBold = VsfData.CellFontBold
                   .FontItalic = VsfData.CellFontItalic
                End With
                arrData = GetData(txtLength.Text)
                intDatas = UBound(arrData)
                If intDatas > 0 Then
                    'ѭ����ֵ
                    For intData = 0 To intDatas
                        If intData > 0 Then
                            VsfData.Rows = VsfData.Rows + 1
                            VsfData.RowPosition(VsfData.Rows - 1) = lngRow + intData
                        End If
                        VsfData.TextMatrix(lngRow + intData, lngCol) = CStr(arrData(intData))
                    Next
                End If
                If lngMaxRowCount < intDatas + 1 Then lngMaxRowCount = intDatas + 1
            End If
        Next lngCol
        
        If lngMaxRowCount > 1 Then
            'ѭ������ǰ������
            For lngCol = VsfData.FixedCols To VsfData.Cols - 1
                If VsfData.ColHidden(lngCol) And lngCol <> mlngRowCount Then
                    'ѭ����ֵ
                    For intData = 2 To lngMaxRowCount
                        VsfData.TextMatrix(lngRow + intData - 1, lngCol) = VsfData.TextMatrix(lngRow, lngCol)
                    Next
                ElseIf lngCol = mlngNoEditor Then
                    '����ֵ��Ϊ��1��ʼ,������4������,����4|1
                    For intData = 1 To lngMaxRowCount
                        VsfData.TextMatrix(lngRow + intData - 1, mlngRowCount) = lngMaxRowCount & "|" & intData
                    Next
                    '���һ����Ҫ��д���ǩ��
                    If mlngSignName > 0 Then VsfData.TextMatrix(lngRow + lngMaxRowCount - 1, mlngSignName) = VsfData.TextMatrix(lngRow, mlngSignName)
                    If mlngSignTime > 0 Then VsfData.TextMatrix(lngRow + lngMaxRowCount - 1, mlngSignTime) = VsfData.TextMatrix(lngRow, mlngSignTime)
                End If
            Next
        Else
            VsfData.TextMatrix(lngRow, mlngRowCount) = "1|1"
        End If
        lngRow = lngRow + lngMaxRowCount
    Loop
    
    VsfData.Rows = VsfData.Rows + 1
    VsfData.ROW = VsfData.Rows - 1
    VsfData.TopRow = VsfData.ROW
    VsfData.COL = mlngDate
    If VsfData.Enabled And VsfData.Visible Then VsfData.SetFocus
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub PreTendFormat(ByVal rsTemp As ADODB.Recordset)
    Dim aryItem() As String
    Dim lngRow As Long, lngCol As Long, lngCount As Long, strCell As String
    Dim blnAlign As Boolean
    On Error GoTo errHand
    
    '���û������ݱ༭����ĸ�ʽ
    With vsfHead
        .Redraw = flexRDNone
        .Clear
        Set .DataSource = rsTemp
        .Rows = 3
        
        '��ͷ��д
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeRow(2) = True
        
        '�����ڲ�����������
        .ColHidden(1) = True
        .ColHidden(2) = True
        .ColHidden(mlngChoose) = Not mblnVerify
        .ColHidden(mlngRowCount) = True
        .ColHidden(mlngRecord) = True
        .ColHidden(mlngSigner) = True
        .ColHidden(mlngSignLevel) = True
        If mlngOperator = -1 Then .ColHidden(.Cols - 1) = True
        .ColWidth(0) = 250
        .ColWidth(mlngChoose) = 250      'ѡ����
        
        .FrozenCols = mlngTime
        .SheetBorder = &H40C0&
        
        '������ͷ
        aryItem = Split(mstrTabHead, "|")
        For lngCount = 0 To UBound(aryItem)
            strCell = aryItem(lngCount)
            lngRow = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            lngCol = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            .TextMatrix(lngRow, lngCol + cHideCols + .FixedCols - 1) = strCell
        Next
        
        '���ù̶��м�ѡ����
        .TextMatrix(0, 0) = " "
        .TextMatrix(1, 0) = " "
        .TextMatrix(2, 0) = " "
        .TextMatrix(0, mlngChoose) = " "
        .TextMatrix(1, mlngChoose) = " "
        .TextMatrix(2, mlngChoose) = " "
        
        '�п�����
        aryItem = Split(mstrColWidth, ",")
        For lngCount = cHideCols + .FixedCols To .Cols - 1
            If Not .ColHidden(lngCount) Then
                .ColWidth(lngCount) = BlowUp(Val(Split(aryItem(lngCount - cHideCols - .FixedCols), "`")(0)))
                If InStr(1, aryItem(lngCount - cHideCols - .FixedCols), "`") <> 0 Then
                    blnAlign = True
                    .ColAlignment(lngCount) = Val(Split(aryItem(lngCount - cHideCols - .FixedCols), "`")(1))
                End If
            End If
        Next
        
        '�̶��и�ʽΪ����
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        '�ٰ��кϲ�
        For lngCount = 0 To .Cols - 1
            .MergeCol(lngCount) = True
        Next
        
        If blnAlign = False Then
            '��Ϊ�����û���������ʾ�ж��뷽ʽ
            If .FixedRows < .Rows Then .Cell(flexcpAlignment, .FixedRows, 0, .Rows - 1, .Cols - 1) = flexAlignGeneralCenter
        End If
        For lngCount = 0 To .Rows - 1
            If .RowHeight(lngCount) <> .RowHeightMin Then .RowHeight(lngCount) = .RowHeightMin
        Next
        Select Case mintTabTiers
        Case 1
            .RowHidden(0) = False
            .RowHidden(1) = True
            .RowHidden(2) = True
        Case 2
            .RowHidden(0) = False
            .RowHidden(1) = False
            .RowHidden(2) = True
        Case 3
            .RowHidden(0) = False
            .RowHidden(1) = False
            .RowHidden(2) = False
        End Select
        For lngCount = 0 To .Cols - 1
            .MergeCol(lngCount) = True
        Next
        
        '���̶ܹ��е��и߲���ȷ��Ҫ�Զ�������
        .AutoResize = True
        .WordWrap = True
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 0, .Cols - 1
        .AutoResize = False
        
        .Height = 0
        For lngCount = 0 To .FixedRows - 1
            If Not .RowHidden(lngCount) Then
                If .RowHeight(lngCount) < .RowHeightMin Then
                    .Height = .Height + .RowHeightMin
                Else
                    .Height = .Height + .RowHeight(lngCount)
                End If
            End If
        Next
        .Height = .Height - 20
        .Redraw = flexRDDirect
    End With
    
    With VsfData
        .Redraw = flexRDNone
        .Clear
        Set .DataSource = rsTemp

        '�����ڲ�����������
        .ColHidden(1) = True
        .ColHidden(2) = True
        .ColHidden(mlngChoose) = Not mblnVerify
        .ColHidden(mlngRowCount) = True
        .ColHidden(mlngRecord) = True
        .ColHidden(mlngSigner) = True
        .ColHidden(mlngSignLevel) = True
        If mlngOperator = -1 Then .ColHidden(.Cols - 1) = True
        .ColWidth(0) = 250
        .ColWidth(mlngChoose) = 250      'ѡ����
        
        .FrozenCols = mlngTime
        .SheetBorder = &H40C0&
        
        '�п�����
        aryItem = Split(mstrColWidth, ",")
        For lngCount = cHideCols + .FixedCols To .Cols - 1
            If Not .ColHidden(lngCount) Then
                .ColWidth(lngCount) = BlowUp(Val(Split(aryItem(lngCount - cHideCols - .FixedCols), "`")(0)))
                If InStr(1, aryItem(lngCount - cHideCols - .FixedCols), "`") <> 0 Then
                    blnAlign = True
                    .ColAlignment(lngCount) = Val(Split(aryItem(lngCount - cHideCols - .FixedCols), "`")(1))
                End If
            End If
        Next
        
        If blnAlign = False Then
            '��Ϊ�����û���������ʾ�ж��뷽ʽ
            If .FixedRows < .Rows Then .Cell(flexcpAlignment, .FixedRows, 0, .Rows - 1, .Cols - 1) = flexAlignGeneralCenter
        End If
        
        Call PreTendMutilRows
        Call WriteColor
        
        '���ǹ̶��е��и�����Ϊ��С�и�
        For lngCount = .FixedRows To .Rows - 1
            .RowHeight(lngCount) = .RowHeightMin
        Next
        
        .Redraw = flexRDDirect
    End With

    With chkSwitch
        .Value = 0
        .Top = vsfHead.Top + vsfHead.Height - .Height - 50
        .Left = vsfHead.Left + vsfHead.Cell(flexcpLeft, mintTabTiers - 1, mlngChoose) + 50
        .Visible = mblnVerify
    End With
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub WriteColor()
    Dim blnTag As Boolean
    Dim lngCount As Long, lngCol As Long
    '����Ժ�ɫ��ʾ��ͬʱ������ʼ������ΪNoCheckBox������ͼ��
    With VsfData
        For lngCount = .FixedRows To .Rows - 1
            If .TextMatrix(lngCount, 2) <> "" Then
                '����Ժ�ɫ��ʾ
                blnTag = False
                If mintTagFormHour < mintTagToHour Then
                    blnTag = (Hour(.TextMatrix(lngCount, 2)) >= mintTagFormHour And Hour(.TextMatrix(lngCount, 2)) < mintTagToHour)
                Else
                    blnTag = (Hour(.TextMatrix(lngCount, 2)) >= mintTagFormHour Or Hour(.TextMatrix(lngCount, 2)) < mintTagToHour)
                End If
                If blnTag Then
                    Set .Cell(flexcpFont, lngCount, 0, lngCount, .Cols - 1) = mobjTagFont
                    .Cell(flexcpForeColor, lngCount, 0, lngCount, .Cols - 1) = mlngTagColor
                End If
            End If
            
            
            '������ʼ������ΪNoCheckBox
            If Not VsfData.RowHidden(lngCount) Then
                If Not VsfData.TextMatrix(lngCount, mlngRowCount) Like "*|1" Then
                    VsfData.Cell(flexcpChecked, lngCount, mlngChoose) = flexNoCheckbox
                Else
                    If VsfData.Cell(flexcpChecked, lngCount, mlngChoose) <> flexTSChecked Then
                        VsfData.Cell(flexcpChecked, lngCount, mlngChoose) = flexTSUnchecked
                    End If
                    
                    '����ͼ��
                    If VsfData.TextMatrix(lngCount, mlngSigner) = "" Then
                        VsfData.Cell(flexcpPicture, lngCount, 0) = Nothing
                    Else
                        If InStr(1, VsfData.TextMatrix(lngCount, mlngSigner), "/") <> 0 Then
                            VsfData.Cell(flexcpPicture, lngCount, 0) = imgRow.ListImages(��ǩ).Picture
                        Else
                            VsfData.Cell(flexcpPicture, lngCount, 0) = imgRow.ListImages(ǩ��).Picture
                        End If
                    End If
                End If
            End If
        Next
    End With
End Sub

Private Sub zlLableBruit()
    Dim lngScaleLeft As Long, lngScaleTop As Long, lngScaleRight As Long, lngScaleBottom As Long
    
    lblSubhead.Caption = lblSubhead.Tag
    lblSubEnd.Caption = lblSubEnd.Tag
    lblSubhead.Top = lblTitle.Top + lblTitle.Height + 120
    Call cbsThis.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
    vsfHead.Move lngScaleLeft + 210, lblSubhead.Top + lblSubhead.Height + 45, lngScaleRight - lngScaleLeft - 210 * 2
    VsfData.Move vsfHead.Left, vsfHead.Top + vsfHead.Height - 20, vsfHead.Width
    VsfData.Height = picMain.Height - vsfHead.Height - vsfHead.Top

    lblSubEnd.Move lblSubhead.Left, VsfData.Top + VsfData.Height + 45
End Sub

Private Sub GetFileProperty()
    '��ȡ�ļ�����
    Dim strEnd As String
    On Error GoTo errHand
    
    gstrSQL = " Select   ��ʼʱ��,����ʱ��,��ʽID,����ID,�鵵�� From ���˻����ļ� " & _
              " Where ����ID=[1] And ��ҳID=[2] And Ӥ��=[3] And ID=[4] And Rownum<2"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ļ�����", mlng����ID, mlng��ҳID, mintӤ��, mlng�ļ�ID)
    If rsTemp.RecordCount <> 0 Then
        mlng��ʽID = rsTemp!��ʽID
        mlng����id = rsTemp!����ID
        mblnArchive = (NVL(rsTemp!�鵵��) <> "")
        mstr��ʼʱ�� = Format(rsTemp!��ʼʱ��, "yyyy-MM-dd HH:mm")
        mstr����ʱ�� = Format(rsTemp!����ʱ��, "yyyy-MM-dd HH:mm")
        mstrBeginTime = mstr��ʼʱ��
        strEnd = DateAdd("n", -1, CDate(Format(CDate(mstr��ʼʱ��) + 1, "yyyy-MM-dd HH:mm:ss")))
        If mstr����ʱ�� = "" Then
            mstr����ʱ�� = Format(strEnd, "YYYY-MM-DD HH:mm:ss")
        Else
            If (mstr����ʱ�� <> "" And CDate(mstr����ʱ��) > CDate(strEnd)) Then mstr����ʱ�� = Format(strEnd, "yyyy-MM-dd HH:mm:ss")
        End If
    End If
    
    '�ڶ����ļ�������ȡ��ʼʱ��
    If mlngFileIndex > 1 Then
        gstrSQL = "SELECT Max(B.����ʱ��) ����ʱ��" & vbNewLine & _
            "FROM ���˻����ļ� A,���˻������� B,���˻�����ϸ C,�����¼��Ŀ D" & vbNewLine & _
            "WHERE A.ID=B.�ļ�ID AND B.ID=C.��¼ID AND A.ID=[1] And ����ID=[2] And ��ҳID=[3] And Ӥ��=[4] AND B.�������<[5] AND C.��Ŀ���=D.��Ŀ���" & vbNewLine & _
            "AND NVL(D.��Ŀ����,'')='����' AND NVL(D.������Ŀ,1)=1 ORDER BY B.����ʱ��"
        Call SQLDIY(gstrSQL)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ļ�����", mlng�ļ�ID, mlng����ID, mlng��ҳID, mintӤ��, mlngFileIndex)
        If rsTemp.RecordCount <> 0 Then
            mstr��ʼʱ�� = DateAdd("n", 1, CDate(Format(rsTemp!����ʱ��, "yyyy-MM-dd HH:mm")))
        End If
    End If
    RaiseEvent AfterRowColChange("", False, mblnSign, mblnArchive)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitEnv()
    Dim rs As New ADODB.Recordset
    On Error GoTo errHand
    
    glngHours = Val(zlDatabase.GetPara("���ݲ�¼ʱ��", glngSys))
    '���ִ��ڵ����л����¼��Ŀ
    gstrSQL = " Select   ��Ŀ���,��Ŀ����,��Ŀ����,��Ŀ����,��Ŀ����,��ĿС��,��Ŀ��ʾ,��Ŀ��λ,��Ŀֵ��,����ȼ�,Ӧ�÷�ʽ" & _
              " From �����¼��Ŀ B" & _
              " Order by ��Ŀ���"
    Set mrsItems = zlDatabase.OpenSQLRecord(gstrSQL, "���ִ��ڵ����л����¼��Ŀ")
    '��ȡ���в���Ҫ����Ϣ
    gstrSQL = "Select ������,�滻��,����,����,С��,��λ,��ʾ��,��ֵ��,����" & vbNewLine & _
        "From (Select i.����id, i.����, i.������, nvl(i.�滻��,0) �滻��,i.����,i.����,i.С��,i.��λ,i.��ʾ��,i.��ֵ��,i.����" & vbNewLine & _
        "       From ����������Ŀ I, ������������ K" & vbNewLine & _
        "       Where k.Id = i.����id And k.���� In ('02', '03', '05', '06') And i.�滻�� = 1 And k.���� = 1" & vbNewLine & _
        "       Union" & vbNewLine & _
        "       Select i.����id, i.����, i.������, nvl(i.�滻��,0) �滻��,i.����,i.����,i.С��,i.��λ,i.��ʾ��,i.��ֵ��,i.����" & vbNewLine & _
        "       From ����������Ŀ I, ������������ K" & vbNewLine & _
        "       Where k.Id = i.����id And k.���� In ('04', '05') And k.���� = 2)" & vbNewLine & _
        "Order By ����id, ����, �滻��"

    Set mrsPartogram = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ����Ҫ����Ϣ")
    'ȡ��ǰ����Ա�ļ���
    mintVerify = δ����
    mintVerify_Last = δ����
    gstrSQL = "select  Ƹ�μ���ְ�� from ��Ա�� p where ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mRichEPR", glngUserId)
    If Not rs.EOF Then
        mintVerify = NVL(rs("Ƹ�μ���ְ��"), δ����)
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitRecords()
    Dim i As Integer, j As Integer, k As Integer, l As Integer
    Dim lngCol As Long, lngOrder As Long, strName As String, intImmovable As Integer, strFormat As String
    Dim arrColumn, arrItem, strColumns As String
    Dim blnSet As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    strColumns = mstrColumns
    If Not mblnInit Then
        '��ʼ���ڴ��¼��(δ��Ӧ��Ŀ����Ϊ���Ŀ,�����о�Ϊ�̶���)
        mstrFields = "��," & adDouble & ",18|���," & adDouble & ",2|��Ŀ���," & adDouble & ",18|��Ŀ����," & adLongVarChar & ",20|�̶�," & adDouble & ",2|��ʽ," & adLongVarChar & ",2000"
        Call Record_Init(mrsSelItems, mstrFields)
        mstrFields = "��|���|��Ŀ���|��Ŀ����|�̶�|��ʽ"
    End If
    
    '�����ж���
    If Not mblnInit Then
        arrColumn = Split(strColumns, "|")
        j = UBound(arrColumn)
        For i = 0 To j
            lngCol = Split(arrColumn(i), "'")(0)
            arrItem = Split(Split(arrColumn(i), "'")(1), ",")
            blnSet = False   '����������Դ���ֵΪ׼'�����Ҳ�����Ŀ���ǻ��Ŀ
            If UBound(Split(arrColumn(i), "'")) > 1 Then
                blnSet = True
                intImmovable = Split(arrColumn(i), "'")(2)
            End If
            If UBound(Split(arrColumn(i), "'")) > 2 Then
                strFormat = Split(arrColumn(i), "'")(3)
            End If
            
            k = UBound(arrItem)
            For l = 0 To k
                strName = arrItem(l)
                mrsItems.Filter = "��Ŀ����='" & strName & "'"
                If mrsItems.RecordCount <> 0 Then
                    lngOrder = mrsItems!��Ŀ���
                    If Not blnSet Then intImmovable = 1   '�̶��������޸�
                    Select Case strName
                        Case "��������"
                            mlngSpread = i + cHideCols + VsfData.FixedCols
                        Case "��¶�ߵ�"
                            mlngJust = i + cHideCols + VsfData.FixedCols
                        Case "����"
                            mlngProduce = i + cHideCols + VsfData.FixedCols
                    End Select
                Else
                    lngOrder = 0
                    If Not blnSet Then intImmovable = 0
                    
                    '��¼������
                    Select Case strName
                    Case "����"
                        mlngDate = i + cHideCols + VsfData.FixedCols
                    Case "ʱ��"
                        mlngTime = i + cHideCols + VsfData.FixedCols
                    Case "��ʿ"
                        mlngOperator = i + cHideCols + VsfData.FixedCols
                    Case "ǩ����"
                        mlngSignName = i + cHideCols + VsfData.FixedCols
                    Case "ǩ��ʱ��"
                        mlngSignTime = i + cHideCols + VsfData.FixedCols
                    End Select
                End If
                mstrValues = lngCol & "|" & l + 1 & "|" & lngOrder & "|" & strName & "|" & intImmovable & "|" & strFormat
                Call Record_Add(mrsSelItems, mstrFields, mstrValues)
            Next
        Next
        
        'Call OutputRsData(mrsSelItems)
        
        '��������ڲ�������(�����ڶ�ȡ���ݺ��ʱ���ӵ�,��ʱֻ��Ԥ������)
        mlngDemo = 1
        mlngActTime = 2
        mlngChoose = 2 + VsfData.FixedCols
        mlngSignLevel = VsfData.Cols + cHideCols + VsfData.FixedCols '����������
        mlngSigner = mlngSignLevel + 1
        mlngRecord = mlngSigner + 1
        mlngRowCount = mlngRecord + 1
        
        If mlngOperator <> -1 And mlngSignName <> -1 Then
            mlngNoEditor = IIf(mlngOperator < mlngSignName, mlngOperator, mlngSignName)
        Else
            mlngNoEditor = IIf(mlngOperator <> -1, mlngOperator, mlngSignName)
        End If
    End If
    
    mrsItems.Filter = 0
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Public Function SignMe(Optional ByVal bln��ǩ As Boolean = False) As Boolean
    Dim blnSign As Boolean          '�Ƿ�ǩ���ɹ�
    Dim blnRefresh As Boolean
    Dim strSignTime As String       '��֤����ǩ����ǩ��ʱ��һ��,����ȡ��ǩ��ʱ��ǩ��ʱ��ͳһȡ��
    Dim str״̬ As String           '����ǩ��ѡ��,����ѭ��ǩ��ʱ��ͣ�ĵ���ǩ������
    Dim str�д��� As String
    Dim str���� As String
    Dim intRow As Integer, intRows As Integer
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '������ʱ��ѭ��������δǩ�����ݽ���ǩ��
    
    If mlng����ID = 0 Then Exit Function
    
    '��ǩ:������δǩ�������ݽ���ǩ��
    '��ǩ:��������ǩ�������ݽ�����ǩ
    If bln��ǩ Then
        If Not mblnVerify Then
            '��������ҲҪǩ��,���ȥ������: And B.�������=0
            gstrSQL = " Select  distinct B.����ʱ�� " & vbNewLine & _
                      " From ���˻�����ϸ A,���˻������� B,���˻����ļ� C" & vbNewLine & _
                      " Where A.��¼ID=B.ID And B.�ļ�ID=C.ID And A.������Դ=0 And MOD(A.��¼����,10)=5 AND A.��ֹ�汾 Is NULL And C.ID=[1] "
            Call SQLDIY(gstrSQL)
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ����δǩ��������", mlng�ļ�ID)
            If rsTemp.RecordCount = 0 Then
                RaiseEvent AfterRowColChange("��������ǩ�������ݣ�", True, mblnSign, mblnArchive)
                Exit Function
            End If
        
            '������ǩģʽ,���޸�����,�ɹ�ѡ����
            mblnVerify = True
            chkSwitch.Visible = mblnVerify
            vsfHead.ColHidden(mlngChoose) = Not mblnVerify
            VsfData.ColHidden(mlngChoose) = Not mblnVerify
            VsfData.Cell(flexcpChecked, VsfData.FixedRows, mlngChoose, VsfData.Rows - 1, mlngChoose) = flexTSUnchecked
            Call WriteColor
            RaiseEvent AfterDataChanged(mblnChange Or mblnVerify)
            Exit Function
        Else
            '��ȡ����ǩ������
            '��������ҲҪǩ��,���ȥ������: And B.�������=0
            gstrSQL = " Select /*+ RULE */ distinct B.����ʱ�� " & vbNewLine & _
                      " From ���˻�����ϸ A,���˻������� B,���˻����ļ� C,(SELECT COLUMN_VALUE FROM TABLE(CAST(F_NUM2LIST([2]) AS ZLTOOLS.T_NUMLIST))) G" & vbNewLine & _
                      " Where A.��¼ID=B.ID And B.ID=G.COLUMN_VALUE And B.�ļ�ID=C.ID And MOD(A.��¼����,10)=5 AND A.��ֹ�汾 Is NULL And C.ID=[1] "
            Call SQLDIY(gstrSQL)
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ����δǩ��������", mlng�ļ�ID, mstrVerify)
        End If
    Else
        '���Ա����޸ĵ����ݽ���ǩ��(��ȡδǩ������-��ǩ������)
        '��������ҲҪǩ��,���ȥ������: And B.�������=0
        gstrSQL = "" & _
                "SELECT  DISTINCT B.����ʱ��" & vbNewLine & _
                "FROM ���˻�����ϸ A,���˻������� B" & vbNewLine & _
                "WHERE A.��¼ID=B.ID And A.������Դ=0 AND A.��ֹ�汾 IS NULL AND A.��¼���� =1 AND instr(NVL(B.ǩ����,'QMR'),'/',1)=0 AND A.��¼��=[2] AND B.�ļ�ID=[1]" & vbNewLine & _
                "MINUS" & vbNewLine & _
                "SELECT DISTINCT B.����ʱ��" & vbNewLine & _
                "FROM ���˻�����ϸ A,���˻������� B" & vbNewLine & _
                "WHERE A.��¼ID=B.ID And A.������Դ=0 AND A.��ֹ�汾 IS NULL AND A.��¼���� =5 AND B.�ļ�ID=[1]"
        Call SQLDIY(gstrSQL)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����Ƿ����δǩ��������", mlng�ļ�ID, gstrUserName)
        If rsTemp.RecordCount = 0 Then
            RaiseEvent AfterRowColChange("û���ҵ���Ҫǩ�������ݣ�ֻ�ܶ��Լ��Ǽǻ��޸ĵ����ݽ���ǩ������", True, mblnSign, mblnArchive)
            Exit Function
        End If
    End If
    
    '׼��ǩ��
    strSignTime = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    str�д��� = ""
    With rsTemp
        Do While Not .EOF
            blnSign = SignName(Format(!����ʱ��, "yyyy-MM-dd HH:mm:ss"), strSignTime, bln��ǩ, str״̬, str�д���)
            If Not blnSign Then Exit Do
            If Not blnRefresh Then blnRefresh = blnSign
'            If str�д��� <> "" Then
'                str���� = str���� & vbCrLf & "����ʱ��=[" & Format(!����ʱ��, "yyyy-MM-dd HH:mm:ss") & "]" & str�д���
'            End If
            .MoveNext
        Loop
    End With
    
    If blnRefresh And Not mblnVerify Then Call ShowMe(mFrmParent, mlng�ļ�ID, mlng����ID, mlng��ҳID, mlng����ID, mintӤ��, mstrPrivs, mblnEditable, IIf(mblnBlowup = True, 1, 0))
    If str���� <> "" Then MsgBox "ǩ��ʱ�������´���" & str����, vbInformation, gstrSysName
    SignMe = blnRefresh
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub UnSignMe(Optional ByVal bln��ǩ As Boolean = False)
    Dim intPos As Integer
    Dim lngStart As Long                '��ʼ��
    Dim lngRecord As Long
    Dim blnOK As Boolean
    Dim strSignTime As String           'ǩ��ʱ��
    Dim blnClear As Boolean             'ȡ��ǩ��ʱ�Ƿ�����ð汾�����ݻ��˵��ϴ�ǩ�����״̬
    Dim blnTrans As Boolean
    Dim strSignName As String
    Dim clsSign As Object
    Dim rsTemp As New ADODB.Recordset
    Dim rsSign As New ADODB.Recordset
    Dim lngRecordCount As Long
    Dim strSQLTime() As String
    ReDim Preserve strSQLTime(1 To 1)
    On Error GoTo errHand
    '�������һ���Ǳ��˵�ǩ�������ݵ�ǰѡ�����ݵ�ǩ��ʱ�䣬����ȡ��ǩ��
    
    If mlng����ID = 0 Then Exit Sub
    
    '��Ҫ�Լ��
    '��ǰ��¼���¼�¼���˳�
    If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "" Then Exit Sub
    lngStart = GetStartRow(VsfData.ROW)
    lngRecord = Val(VsfData.TextMatrix(lngStart, mlngRecord))
    If lngRecord = 0 Then
        RaiseEvent AfterRowColChange("������¼������ȡ��ǩ����", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    '��ǰ��¼δǩ�����˳�
    If VsfData.TextMatrix(lngStart, mlngSigner) = "" Then
        RaiseEvent AfterRowColChange("��ǰ��¼��δǩ����", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    '��ǩ����ǰ��¼δ��ǩ���˳���ƽǩ����ǰ��¼����ǩ���˳�
    intPos = InStr(1, VsfData.TextMatrix(lngStart, mlngSigner), "/")
    If bln��ǩ Then
        If intPos = 0 Then
            RaiseEvent AfterRowColChange("��ǰ��¼δ��ǩ���޷�ִ��ȡ����ǩ������", True, mblnSign, mblnArchive)
            Exit Sub
        End If
    Else
        If intPos <> 0 Then
            RaiseEvent AfterRowColChange("��ǰ��¼����ǩ����ȡ����ǩ���ٲ�����", True, mblnSign, mblnArchive)
            Exit Sub
        End If
    End If
    'ȡ��ǩ��ʱ�����Ի����Լ���ǩ�������Ի��˴�ǩ������
    gstrSQL = "" & _
              " SELECT  A.��¼��,A.��¼ʱ��,A.��Ŀ���� AS ǩ��ʱ��,NVL(A.��ʼ�汾,1) ��ʼ�汾,B.ǩ����" & vbNewLine & _
              " FROM ���˻�����ϸ A,���˻������� B" & vbNewLine & _
              " WHERE A.��¼ID=B.ID AND B.�ļ�ID=[1] AND A.��¼ID=[2] AND A.��¼����=" & IIf(bln��ǩ, 15, 5) & vbNewLine & _
              " ORDER BY A.��Ŀ���� DESC"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ǰ��¼�����ǩ���˲��Ǳ������˳�", mlng�ļ�ID, lngRecord)
    
    If rsTemp.RecordCount = 0 Then
        RaiseEvent AfterRowColChange("�޷��ҵ���" & IIf(bln��ǩ, "��ǩ", "ǩ��") & "�����ݣ����������ݱ仯δˢ�µ��£���ˢ�����ݺ����ԣ�", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    
    If bln��ǩ = False And InStr(1, NVL(rsTemp!ǩ����), "/") <> 0 Then
        RaiseEvent AfterRowColChange("��ǰ��¼����ǩ�����������ݱ仯δˢ�µ��£���ˢ�����ݡ�ȡ����ǩ���ٲ�����", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    
    If NVL(rsTemp!��¼��) <> gstrUserName Then
        strSignName = NVL(rsTemp!��¼��)
        '������Ǳ���ǩ��������Ƿ��Ǵ�ǩ
        gstrSQL = "" & _
              " SELECT A.��¼��" & vbNewLine & _
              " FROM ���˻�����ϸ A,���˻������� B" & vbNewLine & _
              " WHERE A.��¼ID=B.ID AND A.��¼����=1 AND B.�ļ�ID=[1] AND A.��¼ID=[2] And nvl(A.��ʼ�汾,1)=[3]"
        Call SQLDIY(gstrSQL)
        Set rsSign = zlDatabase.OpenSQLRecord(gstrSQL, "��ǰ��¼�����ǩ���˲��Ǳ������˳�", mlng�ļ�ID, lngRecord, Val(NVL(rsTemp!��ʼ�汾, 1)))
        lngRecordCount = rsSign.RecordCount
        rsSign.Filter = "��¼��='" & gstrUserName & "'"
        If rsSign.RecordCount = 0 And lngRecordCount > 0 Then
            RaiseEvent AfterRowColChange("���������ǩ����[" & strSignName & "]���ǩ��[" & gstrUserName & "]������ִ�б�������", True, mblnSign, mblnArchive)
            Exit Sub
        End If
    Else
        strSignName = gstrUserName
    End If
    
    '��ȡ��������׼��ȡ��ǩ������ǩ(��¼ʱ��<>""˵�����°�ǩ��)
    '��������ҲҪǩ��,���ȥ������: And B.�������=0
    If Not IsNull(rsTemp!��¼ʱ��) Then
        gstrSQL = "" & _
                  " SELECT  A.��ĿID AS ֤��ID,B.����ʱ��,B.ǩ����" & vbNewLine & _
                  " FROM ���˻�����ϸ A,���˻������� B" & vbNewLine & _
                  " WHERE A.��¼ID=B.ID AND B.�ļ�ID=[1] And A.��¼��=[2] And A.��¼ʱ��=[3] " & _
                  " AND A.��¼����=" & IIf(bln��ǩ, 15, 5)
        Call SQLDIY(gstrSQL)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������׼��ȡ��ǩ������ǩ", mlng�ļ�ID, strSignName, CDate(rsTemp!��¼ʱ��))
    Else
        gstrSQL = "" & _
                  " SELECT  A.��ĿID AS ֤��ID,B.����ʱ��,B.ǩ����" & vbNewLine & _
                  " FROM ���˻�����ϸ A,���˻������� B" & vbNewLine & _
                  " WHERE A.��¼ID=B.ID AND B.�ļ�ID=[1] And A.��¼��=[2] And A.��Ŀ����=[3] " & _
                  " AND A.��¼����=" & IIf(bln��ǩ, 15, 5) & vbNewLine & _
                  " ORDER BY A.��Ŀ���� DESC"
        Call SQLDIY(gstrSQL)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������׼��ȡ��ǩ������ǩ", mlng�ļ�ID, strSignName, CStr(rsTemp!ǩ��ʱ��))
    End If
    
    'ǩ���������޸ģ������޸ı������ǩ�������ȡ����ǩʱ��������ʾ�Ƿ�������ݵ����⣬��ǩ�Զ����ˣ�����ȡ����ʾ ѯ���Ƿ���Ҫ�������
'    If Not bln��ǩ Then
'        blnClear = (MsgBox("ȡ��ǩ��ʱ�Ƿ�ð汾�����ݻ��˵��ϴ�ǩ�����״̬��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
'    End If
    blnClear = True
    Do While Not rsTemp.EOF
        If (bln��ǩ = False And InStr(1, NVL(rsTemp!ǩ����), "/") = 0) Or bln��ǩ = True Then
            If NVL(rsTemp!֤��ID, 0) > 0 Then
                '����ǩ����֤��ֻ��֤һ��
                Err.Clear
                If clsSign Is Nothing Then
                    If gobjESign Is Nothing Then
                        On Error Resume Next
                        Set gobjESign = CreateObject("zl9ESign.clsESign")
                        If Err <> 0 Then Err.Clear
                        On Error GoTo 0
                        If Not gobjESign Is Nothing Then Call gobjESign.Initialize(gcnOracle, glngSys)
                    End If
                    Set clsSign = gobjESign
                    
                    If Not clsSign Is Nothing Then
                        If Not clsSign.CheckCertificate(gstrDBUser) Then
                            gcnOracle.RollbackTrans
                            Exit Sub
                        End If
                    Else
                        gcnOracle.RollbackTrans
                        RaiseEvent AfterRowColChange("����ǩ������δ����ȷ�������˲������ܼ�����", True, mblnSign, mblnArchive)
                        Exit Sub
                    End If
                End If
            End If
            
            'ȡ��ǩ��
            gstrSQL = "ZL_���˻�������_UNSIGNNAME("
            gstrSQL = gstrSQL & mlng�ļ�ID & ","
            gstrSQL = gstrSQL & "To_Date('" & Format(rsTemp!����ʱ��, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss')," & IIf(blnClear, "1", "0") & ")"
            strSQLTime(ReDimArray(strSQLTime)) = gstrSQL
        End If
        rsTemp.MoveNext
    Loop
    
    gcnOracle.BeginTrans
    blnTrans = True
    
    For intPos = 1 To UBound(strSQLTime)
        If strSQLTime(intPos) <> "" Then
            Call zlDatabase.ExecuteProcedure(strSQLTime(intPos), "ִ��ȡ��ǩ��")
        End If
    Next intPos
    
    gcnOracle.CommitTrans
    blnTrans = False
    
    'ˢ������
    Call ShowMe(mFrmParent, mlng�ļ�ID, mlng����ID, mlng��ҳID, mlng����ID, mintӤ��, mstrPrivs, mblnEditable, IIf(mblnBlowup = True, 1, 0))
    Exit Sub
errHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function SignName(ByVal strStart As String, ByVal strSignTime As String, ByVal bln��ǩ As Boolean, _
    str״̬ As String, Optional str���� As String) As Boolean
    '******************************************************************************************************************
    '����:
    '
    '
    '******************************************************************************************************************
    Dim oSign As cPartogramSign
    Dim strSource As String             '��ǩԴ���ݴ�
    Dim lngLoop As Long
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    '��ʼ����
    '------------------------------------------------------------------------------------------------------------------
    strSource = ""
    
    '��ȡҪǩ��������(��������ҲҪǩ��,���ȥ������: And B.�������=0)
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = " Select  a.��¼����,a.��Ŀ����,a.��Ŀ���,a.��Ŀ����,a.��Ŀ����,a.��¼����,a.��Ŀ��λ,a.��¼���,a.���²�λ,a.��¼���,a.���Ժϸ�,a.δ��˵��,a.��¼��,a.��¼ʱ��  " & _
              " From ���˻�����ϸ a,���˻������� b,���˻����ļ� c " & _
              " Where a.��¼id=b.ID And b.�ļ�ID=c.ID AND MOD(A.��¼����,10)<>5 And a.��ֹ�汾 Is Null And C.ID=[1] And b.����ʱ��=[2]" & _
              " Order by a.��Ŀ���"
    Call SQLDIY(gstrSQL)
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡҪǩ��������", mlng�ļ�ID, CDate(strStart))
    If rs.BOF = False Then
        Do While Not rs.EOF
            For lngLoop = 0 To rs.Fields.Count - 1
                strSource = strSource & CStr(zlCommFun.NVL(rs.Fields(lngLoop).Value, ""))
            Next
            rs.MoveNext
        Loop
    End If
    'Debug.Print "��ʼǩ����" & Now & vbCrLf & strSource
    If strSource = "" Then
        RaiseEvent AfterRowColChange("��ǰû����Ҫǩ������Ϣ��", True, mblnSign, mblnArchive)
        Exit Function
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    Err = 0
    '76223:������,2014-08-05,����ǩ������ʱ�����Ϣ
    Set oSign = frmPartogramSign.ShowMe(Me, mstrPrivs, mlng�ļ�ID, mlng����ID, mintVerify_Last, strSource, bln��ǩ, str״̬, str����)
    On Error GoTo errHand
    
    If Not oSign Is Nothing Then
        gstrSQL = "ZL_���˻�������_SIGNNAME("
        gstrSQL = gstrSQL & mlng�ļ�ID & ","
        gstrSQL = gstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss')," & IIf(bln��ǩ, 1, 0) & ","
        gstrSQL = gstrSQL & "'" & oSign.���� & "',"
        gstrSQL = gstrSQL & "'" & oSign.ǩ����Ϣ & "'," & oSign.ǩ������ & ","
        gstrSQL = gstrSQL & oSign.֤��ID & ","
        gstrSQL = gstrSQL & oSign.ǩ����ʽ & ",'" & oSign.ʱ��� & "',0,'" & oSign.ʱ�����Ϣ & "',"
        gstrSQL = gstrSQL & "To_Date('" & strSignTime & "','yyyy-mm-dd hh24:mi:ss'))"
        
        'Debug.Print gstrSQL
        Call zlDatabase.ExecuteProcedure(gstrSQL, "ִ��ǩ��")
        SignName = True
    End If
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Function CancelMe() As Boolean
    CancelMe = True
    mblnVerify = False
    mblnChange = False
    Call ShowMe(mFrmParent, mlng�ļ�ID, mlng����ID, mlng��ҳID, mlng����ID, mintӤ��, mstrPrivs, mblnEditable, IIf(mblnBlowup = True, 1, 0))
End Function

Public Function SaveME() As Boolean
    If Not CheckData Then Exit Function
    If Not SaveData Then Exit Function
    
    mblnShow = False
    Call InitCons
    SaveME = True
    
    Call ShowMe(mFrmParent, mlng�ļ�ID, mlng����ID, mlng��ҳID, mlng����ID, mintӤ��, mstrPrivs, mblnEditable, IIf(mblnBlowup = True, 1, 0))
End Function

Public Function ShowMe(ByVal frmParent As Form, ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, _
    ByVal lngDeptID As Long, ByVal intBaby As Integer, Optional ByVal strPrivs As String, Optional ByVal blnEditable As Boolean = True, _
    Optional ByVal bytSize As Byte = 0) As Boolean
    '******************************************************************************************************************
    '���ܣ� ��ʾ�����¼�ļ�����
    '������ frmParent           �ϼ��������
    '       lngPatiID           ����id
    '       lngPageID           ��ҳid
    '       lngDeptID           Ҫ��ʾ�����¼�Ŀ���
    '       intBaby             Ӥ����־
    '       blnEditable         ���Ϊ��,˵������Ϊ��ѯ�Ӵ�����ʹ��,ȡ����༭��صĹ���
    '       bytSize             0-С����,1������
    '���أ� ��
    '******************************************************************************************************************
    Dim lngRow As Long
    Dim lngCount As Long, i As Long
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    Err = 0
    
    mblnInit = False
    mblnChange = False
    mlng�ļ�ID = lngFileID
    mlng����ID = lngPatiID
    mlng��ҳID = lngPageId
    mlng����ID = lngDeptID
    mintӤ�� = intBaby
    mlngFileIndex = frmParent.FileNumIndex
    mstrPrivs = strPrivs
    mblnBlowup = (bytSize = 1)
    Set mFrmParent = frmParent
    
    mintPreDays = Val(zlDatabase.GetPara("����¼�뻤����������", glngSys, 1255, "1"))
    
    If mrsItems.State = 0 Then
        Call InitMenuBar
        Call InitEnv            '��ʼ������
        cbsThis.ActiveMenuBar.Visible = False
        cbsThis.RecalcLayout
    End If
    Call ReSetFontSize
    
    '��ȡ�����ļ�����
    lngCount = GetFileCount(mlng�ļ�ID, mlng����ID, mlng��ҳID)
    If mlngFileIndex < 1 Or mlngFileIndex > lngCount Then mlngFileIndex = 1
    With cboBaby
        .Tag = ""
        .Clear
        For i = 1 To lngCount
            .AddItem i: .ItemData(.NewIndex) = i
            If i = mlngFileIndex Then
                .ListIndex = i - 1
            End If
        Next i
        If .ListIndex = -1 Then .ListIndex = 0
        mlngFileIndex = .ItemData(.ListIndex)
    End With
    If mlngFileIndex <= 0 Then Exit Function
    
    mblnEditable = blnEditable And Not gblnMoved And Not mblnArchive
    
    RaiseEvent AfterDataChanged(mblnChange Or mblnVerify)
    RaiseEvent AfterRefresh
    
'    Call OutputRsData(mrsSelItems)
    ShowMe = True
    Exit Function
    
errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ReSetFontSize()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���������С
    '���:bytSize��0-С(ȱʡ)��1-��
    '����:������
    '����:2012-06-20 15:15:00
    '����:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCtrl As Control
    Dim CtlFont As StdFont
    Dim bytFontSize As Byte
    bytFontSize = BlowUp(9)
    
    UserControl.FontSize = bytFontSize
    UserControl.FontName = "����"
    
    Set CtlFont = cbsThis.Options.Font
    If CtlFont Is Nothing Then
        Set CtlFont = UserControl.Font
    End If
    CtlFont.Size = bytFontSize
    Set cbsThis.Options.Font = CtlFont
    
    cboBaby.FontSize = bytFontSize
    picTmp.Height = cboBaby.Height
    vsfBaby.FontSize = bytFontSize
End Sub

Private Function CheckFlip() As Boolean
    Dim blnExit As Boolean
    Dim lngRow As Long, lngCol As Long, lngRows As Long, lngCols As Long
    '���ݱ���ʱ����Ƿ�¼�������ʱ��
    
    lngRows = VsfData.Rows - 1
    lngCols = VsfData.Cols - 1
    For lngRow = VsfData.FixedRows To lngRows
        mrsCellMap.Filter = "ҳ��=" & mintҳ�� & " And �к�=" & lngRow & " And �к�>" & mlngTime
        If mrsCellMap.RecordCount <> 0 Then
            If Not VsfData.RowHidden(lngRow) Then
                blnExit = (VsfData.TextMatrix(lngRow, mlngDate) = "" Or VsfData.TextMatrix(lngRow, mlngTime) = "")
                If blnExit Then
                    mrsCellMap.Filter = 0
                    VsfData.ROW = lngRow
                    If Not VsfData.RowIsVisible(lngRow) Then VsfData.TopRow = lngRow
                    If mblnDate = False Then
                        lngCol = mlngTime
                    Else
                        lngCol = IIf(VsfData.TextMatrix(lngRow, mlngDate) = "", mlngDate, mlngTime)
                    End If
                    VsfData.COL = lngCol
                    If Not VsfData.ColIsVisible(lngCol) Then VsfData.LeftCol = lngCol
                    RaiseEvent AfterRowColChange("�벹������ʱ�䣡", True, mblnSign, mblnArchive)
                    CheckFlip = False
                    Exit Function
                End If
            End If
        End If
    Next
    
    mrsCellMap.Filter = 0
    CheckFlip = True
End Function

Private Function CheckProduce() As Boolean
    Dim rsTemp As New ADODB.Recordset
    Dim blnExit As Boolean, blnProduce As Boolean
    Dim lngRow As Long, lngCol As Long
    Dim lngPage As Long, lngCount As Long
    Dim strDatetime As String
    
    '����ʱ���,1��Ӥ������ʱ������ڶ�Ӧ�Ĺ�����̥ͷ����,2.�������Ӥ���ļ���>1������ȡ����һ�ļ�Ӥ��������־
    '˵�����������DataMap_Save֮��ʹ��
    
    On Error GoTo errHand
    
    '--1��Ӥ������ʱ������ڶ�Ӧ�Ĺ�����̥ͷ����
    mrsDataMap.Filter = ""
    mrsDataMap.Filter = "ɾ��=0 And ��������<>''"
    mrsDataMap.Sort = "��������"
    'Call OutputRsData(mrsDataMap)
    Do While Not mrsDataMap.EOF
        If blnExit = False Then blnExit = (NVL(mrsDataMap.Fields(cControlFields + mlngSpread - VsfData.FixedCols).Value) <> "") And (NVL(mrsDataMap.Fields(cControlFields + mlngJust - VsfData.FixedCols).Value) <> "")
        If Mid(NVL(mrsDataMap.Fields(cControlFields + mlngProduce - VsfData.FixedCols).Value, 0), 1, 1) = "��" Then
            blnProduce = True
            strDatetime = Format(mrsDataMap!��������, "YYYY-MM-DD HH:mm:ss")
            lngPage = mrsDataMap!ҳ��
            lngRow = mrsDataMap!�к�
            If NVL(mrsDataMap.Fields(cControlFields + mlngSpread - VsfData.FixedCols).Value) = "" Then
                lngCol = mlngSpread
            Else
                lngCol = mlngJust
            End If
            GoTo ErrProduce
        End If
    mrsDataMap.MoveNext
    Loop
ErrProduce:
    If blnExit = False And blnProduce = True Then
        lngCount = 0
        If lngCount = 0 Then
            lngRow = IIf(lngPage = mintҳ��, lngRow, VsfData.FixedRows)
            VsfData.ROW = lngRow
            VsfData.COL = lngCol
            If Not VsfData.RowIsVisible(lngRow) Then VsfData.TopRow = lngRow
            If Not VsfData.ColIsVisible(mlngSpread) Then VsfData.LeftCol = mlngSpread
            RaiseEvent AfterRowColChange("Ӥ����������ͬʱ���ڹ����������¶�½����ݣ����飡", True, mblnSign, mblnArchive)
            CheckProduce = False
            Exit Function
        End If
    End If
    
    '2.�������Ӥ���ļ���>1������ȡ����һ�ļ�Ӥ��������־
    If cboBaby.ListCount > 1 And cboBaby.ListIndex < cboBaby.ListCount - 1 Then
        If blnProduce = False Then '��ʾ������Ӥ���������
            RaiseEvent AfterRowColChange("��ǰӤ���ļ�С�����ļ���ʱ�����������Ӧ��������־�����飡", True, mblnSign, mblnArchive)
            CheckProduce = False
            Exit Function
        End If
    End If
    CheckProduce = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckData() As Boolean
    Dim intLevel As Integer
    Dim lngPage As Long
    On Error GoTo errHand
    '�������
    
    '����޸������ݶ�����ʱ�䲻ȫ����ʾ�����ݺϷ�����¼��ʱ�Ѿ���飩
    If Not DataMap_Save Then Exit Function
    If Not CheckProduce Then Exit Function
    
    '�������ǩģʽ,������ѡ�����Ƿ���ڲ�����ǩ�����
    If mblnVerify Then
        mstrVerify = ""
        mintVerify_Last = δ����
        '��ǩ��������������
        mrsDataMap.Filter = "ҳ��=" & mintҳ��
        Do While Not mrsDataMap.EOF
            If NVL(mrsDataMap!ѡ��, 0) = flexTSChecked Then
                mstrVerify = mstrVerify & "," & mrsDataMap!��¼ID
                
                If IsNull(mrsDataMap!ǩ������) Then
                    intLevel = NVL(mrsDataMap!ǩ������, δ����)
                Else
                    intLevel = Val(mrsDataMap!ǩ������) + 1
                End If
                If mintVerify < intLevel Then mintVerify_Last = intLevel
            End If
            mrsDataMap.MoveNext
        Loop
        mrsDataMap.Filter = 0
        
        If mstrVerify = "" Then
            RaiseEvent AfterRowColChange("����Ҫѡ��һ�����ݲ��������ǩ������", True, mblnSign, mblnArchive)
            Exit Function
        End If
        mstrVerify = Mid(mstrVerify, 2)
    End If
    
    CheckData = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function SaveData() As Boolean
    Dim arrValue, arrOrder, arrPart
    Dim strSQL() As String
    Dim intAllow As Integer
    Dim lngRecord As Long
    Dim blnTrans As Boolean
    Dim intPos As Integer, intMax As Integer, intRow As Integer
    Dim strReturn As String, strCellData As String, strPart As String
    Dim strMonth As String, strDay As String
    Dim strDate As String, strTime As String, strTemp As String
    Dim strDatetime As String, strCurrDate As String
    
    ReDim Preserve strSQL(1 To 1)
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    intAllow = IIf(InStr(mstrPrivs, "���˻����¼") > 0, 1, 0)
    strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    With mrsCellMap
        '����Ч���ݹ��˳���:��¼ID>0����ʷ����+��������Ч����
        .Filter = "��¼ID>0 or (��¼ID=0 And ɾ��=0)"
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If (InStr(1, "," & mstrVerify & ",", "," & NVL(!��¼ID, 0) & ",") <> 0 And mblnVerify = True) Or mblnVerify = False Then
                If intRow <> !�к� Then
                    '����ֵ
                    intRow = !�к�
                    strDate = ""
                    strDatetime = ""
                    lngRecord = NVL(!��¼ID, 0)
                End If
                
                If !�к� = mlngDate Then
                    strDate = NVL(!����)
                    If strDate <> "" Then
                        If mblnDateAd Then
                            strDate = Mid(strCurrDate, 1, 5) & ToStandDate(strDate)
                        Else
                            strDate = Format(strDate, "yyyy-MM-dd")
                        End If
                    End If
                ElseIf !�к� = mlngTime Then
                    strTime = NVL(!����)
                    If strDate = "" Then strDate = Mid(strCurrDate, 1, 10)
                    strDatetime = strDate & " " & strTime & ":00"
                    If mblnDateAd Then
                        strDatetime = GetDateAdCurrDate(strDatetime)
                    End If
                    
                    If lngRecord <> 0 Then
                        '���·���ʱ��
                        gstrSQL = "ZL_����ͼ����_����ʱ��(" & lngRecord & ",to_date('" & strDatetime & "','yyyy-MM-dd hh24:mi:ss'))"
                        strSQL(ReDimArray(strSQL)) = gstrSQL
                    End If
                Else
                    If !�к� > mlngTime Then
                        'ȡָ����Ԫ�������
                        strCellData = NVL(!����)
                        strPart = NVL(!��λ)
                        strReturn = ShowInput(!�к�, strCellData, True)
                        'strOrders��ʽ����Ŀ���,��Ŀ���...
                        'strValues��ʽ��ֵ'ֵ'ֵ...
                        arrOrder = Split(Split(strReturn, "||")(0), ",")
                        arrValue = Split(Split(strReturn, "||")(1) & "'", "'")
                        arrPart = Split(strPart & "/////", "/")
                        
                        intMax = UBound(arrOrder)
                        For intPos = 0 To intMax
        '                    �ļ�ID_IN IN ���˻�������.�ļ�ID%TYPE,
        '                    ����ʱ��_IN IN ���˻�������.����ʱ��%TYPE,
        '                    ��¼����_IN IN ���˻�����ϸ.��¼����%TYPE,          --������Ŀ=1���ϱ�˵��=2�������ձ��=4��ǩ����¼=5���±�˵��=6�����������=9
        '                    ��Ŀ���_IN IN ���˻�����ϸ.��Ŀ���%TYPE,          --������Ŀ����ţ��ǻ�����Ŀ�̶�Ϊ0
        '                    ��¼����_IN IN ���˻�����ϸ.��¼����%TYPE := NULL,  --��¼���ݣ��������Ϊ�գ��������ǰ�����ݣ�37��38/37
        '                    ���²�λ_IN IN ���˻�����ϸ.���²�λ%TYPE := NULL,
        '                    ��¼���_IN IN ���˻�������.�������%TYPE := 1,             --����Ǹ�Ӥ��������
        '                    ���˼�¼_IN IN NUMBER := 1,
                            gstrSQL = "ZL_����ͼ����_UPDATE(" & mlng�ļ�ID & ",to_date('" & strDatetime & "','yyyy-MM-dd hh24:mi:ss'),1," & _
                                    arrOrder(intPos) & ",'" & arrValue(intPos) & "','" & arrPart(intPos) & "'," & mlngFileIndex & "," & intAllow & ",0," & IIf(mblnVerify, 1, 0) & ")"
                            strSQL(ReDimArray(strSQL)) = gstrSQL
                        Next
                        mrsItems.Filter = 0
                    End If
                End If
            End If
            .MoveNext
        Loop
        
        mrsDataMap.Filter = 0
    End With
    
    'ѭ��ִ��SQL��������
    intMax = UBound(strSQL)
    gcnOracle.BeginTrans
    blnTrans = True
    If intMax > 0 Then
        For intPos = 1 To intMax
            If strSQL(intPos) <> "" Then
                'Debug.Print strSQL(intPos)
                Call zlDatabase.ExecuteProcedure(strSQL(intPos), "�������ͼ����")
            End If
        Next
    End If
    If mblnVerify Then
        If Not SignMe(mblnVerify) Then
            gcnOracle.RollbackTrans
            Exit Function
        End If
    End If
    
    gcnOracle.CommitTrans
    SaveData = True
    blnTrans = False
    mblnChange = False
    mblnVerify = False
    
    RaiseEvent AfterDataChanged(mblnChange Or mblnVerify)
    RaiseEvent AfterRefresh
    RaiseEvent AfterDataSave(True)
    Exit Function
errHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cboBaby_Click()
    Dim i As Integer
    If Val(cboBaby.Tag) = cboBaby.ItemData(cboBaby.ListIndex) Then Exit Sub
    mblnInit = False
    If mblnChange Then
        If MsgBox("��ǰ���˵����ݻ�δ���棬�㡰�ǡ��ֹ����б��棬�㡰�񡱽����������޸ģ�", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            For i = 0 To cboBaby.ListCount - 1
                If cboBaby.ItemData(i) = mlngFileIndex Then
                    Call zlControl.CboSetIndex(cboBaby.Hwnd, i)
                    Exit For
                End If
            Next i
            Call InitCons
            Exit Sub
        Else
            mblnChange = False
        End If
    End If
    mlngFileIndex = cboBaby.ItemData(cboBaby.ListIndex)
    cboBaby.Tag = mlngFileIndex
    'Debug.Print Now & "Begin"
    Call InitVariable
    Call InitCons
    If Not ReadStruDef Then Exit Sub
    Call zlRefresh
    'Debug.Print Now & "Over"
    mblnInit = True
    RaiseEvent AfterFileIndex(mlngFileIndex)
    RaiseEvent AfterDataChanged(mblnChange)
End Sub

Private Sub cboChoose_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Index = 0 Then
            cboChoose(1).SetFocus
        Else
            Call MoveNextCell
        End If
    End If
End Sub


Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim strDate As String, strTime As String
    Dim strLockItem As String                   'ͬ������������,�������޸Ļ�ɾ��
    Dim lngTop As Long, lngHeight As Long
    Dim intMax As Integer                       'ͬ������������ռ�õ��������
    Dim intNULL As Integer, lngStartRow As Long
    Dim lngRow As Long, lngCol As Long, lngRows As Long, lngCols As Long
    Dim strKey As String, strField As String, strValue As String
    Dim blnShow As Boolean
    On Error GoTo err_exit
    
    Select Case Control.ID
    'ճ��,���ʱ��Ҫͬ��mrsCellMap����
    Case conMenu_Edit_FileMan
        '�ļ����
        Call LoadBabyNum
    Case conMenu_Edit_Copy
        '����ָ�������е�����
        If VsfData.ROW < VsfData.FixedRows Then Exit Sub
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then Exit Sub
        lngRow = GetStartRow(VsfData.ROW)
        
        '���Ƽ�¼��
        Set mrsCopyMap = New ADODB.Recordset
        Set mrsCopyMap = CopyNewRec(mrsDataMap, False)
        
        '�õ�ָ�������е���ʼ��,������
        lngCols = VsfData.Cols - 1
        lngRows = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
        lngRows = lngRow + lngRows - 1
        For lngRow = lngRow To lngRows
            mrsCopyMap.AddNew
            mrsCopyMap!ҳ�� = mintҳ��
            mrsCopyMap!�к� = lngRow
            For lngCol = 0 To lngCols - VsfData.FixedCols    '����һ���̶���
                mrsCopyMap.Fields(cControlFields + lngCol).Value = IIf(VsfData.TextMatrix(lngRow, lngCol + VsfData.FixedCols) = "", Null, VsfData.TextMatrix(lngRow, lngCol + VsfData.FixedCols))
            Next
            mrsCopyMap.Update
        Next
    Case conMenu_Edit_PASTE
        'ճ��ʱ����Ŀ�������帲�ǣ�ͬ�������������У���г���
        '���Ŀ���ܲ�ͬҳ����Ŀ��ͬ����λ��ͬ�����Բ����ǻ��Ŀ
        'ͬ������ռ�õ��������䣬�粻������ӿհ��У�����ճ��
        If VsfData.ROW < VsfData.FixedRows Then Exit Sub
        If mrsCopyMap.RecordCount = 0 Then Exit Sub
        
        '��ҳ�����в���������н���ճ��,ɾ��,ֻ�ܱ༭�����Ŀ�����
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then
            lngStartRow = VsfData.ROW
        Else
            lngStartRow = GetStartRow(VsfData.ROW)
        End If
        
        '���Ŀ���������Ƿ����ͬ������������,�����������ͬ���ļ�¼
        strLockItem = GetSynItems(2, intMax)        '1.������Ŀ���;2.�����к�
        
        '�õ�Ŀ�������е���ʼ��,������
        strField = "ID|ҳ��|�к�|�к�|��¼ID|����|ɾ��"
        lngCols = VsfData.Cols - 1
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then
            lngRow = VsfData.ROW
            lngStartRow = lngRow
            If mlngDate > -1 Then strDate = VsfData.TextMatrix(lngRow, mlngDate)
            strTime = VsfData.TextMatrix(lngRow, mlngTime)
        Else
            'ɾ�������������,����һ��
            lngRow = GetStartRow(VsfData.ROW)
            lngStartRow = lngRow
            strDate = VsfData.TextMatrix(lngRow, mlngDate)
            strTime = VsfData.TextMatrix(lngRow, mlngTime)
            lngRows = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0)) - 1
            For intNULL = 1 To lngRows
                VsfData.RemoveItem lngRow + 1
            Next
        End If
        
        '������������,�������������������������ӵ�����
        intNULL = mrsCopyMap.RecordCount - 1
        For lngRow = 1 To mrsCopyMap.RecordCount - 1
            '��֤��ǰ�����������һҳ����ʾȫ
            If lngRow + VsfData.ROW > VsfData.Rows - 1 Then Exit For
            
            If Val(VsfData.TextMatrix(lngRow + VsfData.ROW, mlngRecord)) = 0 And VsfData.TextMatrix(lngRow + VsfData.ROW, mlngRowCount) = "" Then
                intNULL = intNULL - 1
            Else
                Exit For
            End If
        Next
        '�����ӿ���
        If intNULL > 0 Then
            VsfData.Rows = VsfData.Rows + intNULL
            '�ӵ�ǰ�м�¼�Ŀհ��п�ʼ��ÿ�е�λ��+�����ӵĿհ�����
            For lngRow = 1 To intNULL
                VsfData.RowPosition(VsfData.Rows - 1) = lngStartRow + 1
            Next
        End If
        
        '��ԭ���ڣ�ʱ�䣬ǿ�Ʋ������޸�
        VsfData.TextMatrix(lngStartRow, mlngDate) = strDate
        VsfData.TextMatrix(lngStartRow, mlngTime) = strTime
        '��¼�û��޸Ĺ��ĵ�Ԫ��
        If mlngDate <> -1 Then
            strKey = mintҳ�� & "," & lngStartRow & "," & mlngDate
            strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & mlngDate & "|" & _
                Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & VsfData.TextMatrix(lngStartRow, mlngDate) & "|0"
            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        End If
        '2\ʱ��
        strKey = mintҳ�� & "," & lngStartRow & "," & mlngTime
        strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & mlngTime & "|" & _
            Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & VsfData.TextMatrix(lngStartRow, mlngTime) & "|0"
        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        '�����������
        With mrsCopyMap
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                For lngCol = 0 To lngCols - VsfData.FixedCols
                    Select Case lngCol + VsfData.FixedCols
                    Case 1, mlngDate, mlngTime, mlngOperator, mlngSigner, mlngSignTime, mlngRecord
                    Case Else
                        If InStr(1, "," & strLockItem & ",", "," & lngCol - (cHideCols - 1) & ",") = 0 And InStr(1, "," & mstrCOLNothing & ",", "," & lngCol - (cHideCols - 1) & ",") = 0 Then
                            VsfData.TextMatrix(lngStartRow + .AbsolutePosition - 1, lngCol + VsfData.FixedCols) = NVL(.Fields(cControlFields + lngCol).Value)
                            
                            '�޸ı�־
                            If .AbsolutePosition = 1 Then
                                strKey = mintҳ�� & "," & lngStartRow & "," & lngCol + VsfData.FixedCols
                                strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & lngCol + VsfData.FixedCols & "|" & _
                                    Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & GetMutilData(lngStartRow, lngCol + VsfData.FixedCols, lngTop, lngHeight) & "|0"
                                Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                            End If
                        End If
                    End Select
                Next
                .MoveNext
            Loop
        End With
        mblnChange = True
        RaiseEvent AfterDataChanged(mblnChange Or mblnVerify)
    
    Case conMenu_Edit_Clear
        Dim lngRowCount As Long
        If VsfData.ROW < VsfData.FixedRows Then Exit Sub
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then Exit Sub
        If VsfData.TextMatrix(VsfData.ROW, mlngSigner) <> "" Then
            RaiseEvent AfterRowColChange("��ǩ�������ݲ�����ɾ����", True, mblnSign, mblnArchive)
            Exit Sub
        End If
        
        '��ҳ�����в���������н���ճ��,ɾ��,ֻ�ܱ༭�����Ŀ�����
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then
            lngStartRow = VsfData.ROW
        Else
            lngStartRow = GetStartRow(VsfData.ROW)
        End If
        
        lngRowCount = Split(VsfData.TextMatrix(lngStartRow, mlngRowCount), "|")(0)
        '���Ŀ���������Ƿ����ͬ������������,�����������ͬ���ļ�¼
        strLockItem = GetSynItems(2, intMax)        '1.������Ŀ���;2.�����к�
        
        '׼��ɾ��
        strField = "ID|ҳ��|�к�|�к�|��¼ID|����|����|ɾ��"
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then
            lngRow = VsfData.ROW
        Else
            lngRow = GetStartRow(VsfData.ROW)
            lngStartRow = lngRow
            If VsfData.TextMatrix(lngStartRow, mlngSigner) <> "" Then
                RaiseEvent AfterRowColChange("��ǩ�������ݲ�����ɾ����", True, mblnSign, mblnArchive)
                Exit Sub
            End If
            
            'ɾ������������
            lngRows = Val(Split(VsfData.TextMatrix(lngStartRow, mlngRowCount), "|")(0))
            For intNULL = 2 To lngRows
                VsfData.RowHidden(lngRow + intNULL - 1) = True
            Next
        End If
        
        Select Case mintType
        Case 0, 3
            picInput.Visible = False
        Case 1, 2
            lstSelect(mintType - 1).Visible = False
        Case 4, 5
            picDouble.Visible = False
        Case 6
            picMutilInput.Visible = False
        Case 7
            picDoubleChoose.Visible = False
        End Select
        cmdWord.Visible = False
        mintType = -1
        blnShow = mblnShow
        mblnShow = False
        '��¼�û��޸Ĺ��ĵ�Ԫ��
        strDate = VsfData.TextMatrix(lngStartRow, mlngDate)
        strTime = VsfData.TextMatrix(lngStartRow, mlngTime)
            
        If mblnDateAd Then
            If InStr(1, strDate, "/") <> 0 Then
                strDate = Mid(zlDatabase.Currentdate, 1, 5) & Split(strDate, "/")(1) & "-" & Split(strDate, "/")(0)
            End If
            strDate = Mid(strDate, 9, 2) & "/" & Mid(strDate, 6, 2)
        End If
        
        strField = "ID|ҳ��|�к�|�к�|��¼ID|����|����|ɾ��"
        strKey = mintҳ�� & "," & lngStartRow & "," & mlngDate
        strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & mlngDate & "|" & _
            Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & strDate & "|0|1"
        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        '2\ʱ��
        strKey = mintҳ�� & "," & lngStartRow & "," & mlngTime
        strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & mlngTime & "|" & _
            Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & strTime & "|0|1"
        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        
        strField = "ID|ҳ��|�к�|�к�|��¼ID|����|����|ɾ��"
        
        'ɾ����ʼ���з�ͬ��������
        If strLockItem = "" Then
            VsfData.RowHidden(lngRow) = True
            '��д�޸ı�־
            For lngCol = mlngTime + 1 To mlngNoEditor - 1
                strKey = mintҳ�� & "," & lngStartRow & "," & lngCol
                strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & lngCol & "|" & _
                    Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "||0|1"
                Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
            Next
        Else
            '��д�޸ı�־(����ͬ������,������ʱ���в��������)``
            For lngCol = mlngTime + 1 To mlngNoEditor - 1
                If InStr(1, "," & strLockItem & ",", "," & lngCol - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 And lngCol <> mlngDate And lngCol <> mlngTime Then
                    VsfData.TextMatrix(lngStartRow, lngCol) = ""
                    
                    strKey = mintҳ�� & "," & lngStartRow & "," & lngCol
                    strValue = strKey & "|" & mintҳ�� & "|" & lngStartRow & "|" & lngCol & "|" & _
                        Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "||0|1"
                    Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                End If
            Next
            VsfData.TextMatrix(lngStartRow, mlngRowCount) = "1|1"
        End If
        
        If lngStartRow + lngRowCount < VsfData.Rows - 1 Then
            VsfData.ROW = lngStartRow + lngRowCount
        End If
        
        mblnShow = blnShow
        mblnChange = True
        RaiseEvent AfterDataChanged(mblnChange Or mblnVerify)
        
    Case conMenu_Edit_SPECIALCHAR
        
        '��鵱ǰ¼��ؼ�
        On Error Resume Next
        Dim objTXT As TextBox
        Dim strText As String
        Dim intPos As Integer, intLen As Integer
        
        mstrSymbol = frmInsSymbol.ShowMe(False, 0)
        If mintSymbol = -1 Then
            Set objTXT = txtInput
        Else
            Set objTXT = txt(mintSymbol)
        End If
        objTXT.SetFocus
        strText = objTXT.Text
        intPos = objTXT.SelStart
        intLen = Len(objTXT)
        objTXT.Text = Mid(strText, 1, intPos) & mstrSymbol & Mid(strText, intPos + 1)
    
        If mintSymbol = -1 Then
            Call txtInput_KeyDown(vbKeyReturn, 0)
        Else
            Call txt_KeyDown(Val(txt(mintSymbol)), vbKeyReturn, 0)
        End If
    Case conMenu_Edit_Append '����Ҫ��
        RaiseEvent AfterPartogramInfo(mlng�ļ�ID, mlngFileIndex, mlng��ʽID, mrsPartogram)
    Case conMenu_Edit_Word
        Call cmdWord_Click
    End Select
    
err_exit:
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim arrData
    Dim blnFind As Boolean
    Dim strItem As String
    Dim intDo  As Integer, intCount As Integer
    On Error GoTo errHand
    
    If Not mblnInit Then Exit Sub
    Select Case Control.ID
    Case conMenu_Edit_FileMan '���Ӥ��
        Control.Enabled = Not mblnArchive And mblnEditable And Not mblnVerify And Not mblnChange
        If picBaby.Visible = True Then
            picBaby.Visible = Control.Enabled
        End If
    Case conMenu_Edit_Copy
        Control.Enabled = Not mblnShow And Not mblnArchive And Not mblnVerify And mblnEditable
    Case conMenu_Edit_PASTE
        Control.Enabled = False
        If mrsCopyMap.State = 0 Then Exit Sub
        'ǩ�����ݲ�����ճ��
        If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) <> "" Then
            intDo = GetStartRow(VsfData.ROW)
        Else
            intDo = VsfData.ROW
        End If
        If VsfData.TextMatrix(intDo, mlngSigner) <> "" Then Exit Sub
        
        Control.Enabled = Not mblnShow And Not mblnArchive And mblnEditable And mrsCopyMap.RecordCount
    Case conMenu_Edit_Clear
        Control.Enabled = False
        If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) <> "" Then
            intDo = GetStartRow(VsfData.ROW)
        Else
            intDo = VsfData.ROW
        End If
        If VsfData.TextMatrix(intDo, mlngSigner) <> "" Then Exit Sub
        
        Control.Enabled = Not mblnArchive And Not mblnVerify And mblnEditable
    Case conMenu_Edit_SPECIALCHAR
        Control.Enabled = mblnShow And Not mblnArchive And mblnEditable And (mintType = 0 Or mintType = 6)
    Case conMenu_Edit_Append '����Ҫ��
        Control.Enabled = Not mblnArchive And mblnEditable
    Case conMenu_Edit_Word
        Control.Enabled = mblnEditAssistant And mblnShow And Not mblnArchive And mblnEditable
    End Select
errHand:
End Sub

Private Sub chkSwitch_Click()
    Dim blnSel As Boolean            '�Ƿ�ȫ��ѡ��
    Dim blnUpdate As Boolean
    Dim intLevel As Integer
    Dim lngRow As Long, lngRows As Long
    Dim strKey As String, strField As String, strValue As String
    '��������ȫ��ѡ�л�ȡ��ѡ�У����������
    
    If Not mblnInit Then Exit Sub
    lngRows = VsfData.Rows - 1
    strField = "ID|ҳ��|�к�|�к�|��¼ID|����|ɾ��"
    
    blnSel = chkSwitch.Value
    For lngRow = VsfData.FixedRows To lngRows
        If Not VsfData.RowHidden(lngRow) Then
            If VsfData.TextMatrix(lngRow, mlngRowCount) Like "*|1" Then
                blnUpdate = False
                If blnSel Then
                    '���,ǩ�����ļ�¼,�ҵ�ǰ����Ա������ϴ�ǩ�������
                    If VsfData.TextMatrix(lngRow, mlngSignLevel) = "" Then
                        intLevel = δ����
                    Else
                        intLevel = Val(VsfData.TextMatrix(lngRow, mlngSignLevel)) + 1
                    End If
                    If mintVerify < intLevel And intLevel <> δ���� Then
                        blnUpdate = (VsfData.Cell(flexcpChecked, lngRow, mlngChoose) <> flexTSChecked)
                        VsfData.Cell(flexcpChecked, lngRow, mlngChoose) = flexTSChecked
                    End If
                Else
                    blnUpdate = (VsfData.Cell(flexcpChecked, lngRow, mlngChoose) <> flexTSUnchecked)
                    VsfData.Cell(flexcpChecked, lngRow, mlngChoose) = flexTSUnchecked
                End If
                
                If blnUpdate Then
                    '�����޸ļ�¼�Ա�ͬ��
                    strKey = mintҳ�� & "," & lngRow & "," & mlngChoose
                    strValue = strKey & "|" & mintҳ�� & "|" & lngRow & "|" & mlngChoose & "|" & _
                        Val(VsfData.TextMatrix(lngRow, mlngRecord)) & "|" & VsfData.Cell(flexcpChecked, lngRow, mlngChoose) & "|1"
                    Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
                End If
            End If
        End If
    Next
End Sub

Private Sub cmdAddBaby_Click()
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    Dim lngGroupID As Long
    If picBaby.Visible = False Then Exit Sub
    
    On Error GoTo errHand
    
    lngGroupID = Val(cboBaby.ItemData(cboBaby.ListCount - 1))
    '��ӹ���Ϊ��һ��Ӥ���������������һ��Ӥ��
    strSQL = " SELECT 1" & vbNewLine & _
            " FROM ���˻����ļ� A, ���˻������� B, ���˻�����ϸ C,�����¼��Ŀ D" & vbNewLine & _
            " WHERE A.ID = B.�ļ�ID AND B.ID = C.��¼ID AND A.ID = [1] AND A.����ID = [2] AND A.��ҳID = [3] AND B.������� = [4]" & vbNewLine & _
            " AND substr(nvl(C.��¼����,''),1,1)='��' AND C.��Ŀ���=D.��Ŀ��� AND D.��Ŀ����='����' AND NVL(D.������Ŀ,0)=1"
    Call SQLDIY(strSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�ļ�ɾ�����", mlng�ļ�ID, mlng����ID, mlng��ҳID, lngGroupID)
    If rsTemp.RecordCount = 0 Then
        RaiseEvent AfterRowColChange("��ӹ���Ϊ��һӤ���ļ���Ӥ���Ѿ����������������һ����", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    
    cboBaby.AddItem lngGroupID + 1
    cboBaby.ItemData(cboBaby.NewIndex) = lngGroupID + 1
    cboBaby.ListIndex = cboBaby.ListCount - 1
    cboBaby.Refresh
    
    picBaby.Visible = False
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdBabyCancle_Click()
    picBaby.Visible = False
End Sub

Private Sub cmdDelBaby_Click()
    Dim intRow As Integer
    Dim lngFileIndex As Long, lngFileOldIndex As Long
    If picBaby.Visible = False Or Val(vsfBaby.RowData(vsfBaby.ROW)) < 2 Then Exit Sub
    
    lngFileIndex = Val(vsfBaby.RowData(vsfBaby.ROW))
    'Ϊ�˱�֤ɾ��ֻ�ܴӺ���ǰ���˴��ٴν����ж�
    For intRow = vsfBaby.FixedRows To vsfBaby.Rows - 1
        If lngFileOldIndex < Val(vsfBaby.RowData(intRow)) Then
            lngFileOldIndex = Val(vsfBaby.RowData(intRow))
        End If
    Next intRow
    
    If lngFileIndex < lngFileOldIndex Then
       RaiseEvent AfterRowColChange("ɾ��ֻ�ܴ����һ��Ӥ����ʼ�����飡", True, mblnSign, mblnArchive)
       Exit Sub
    End If
    
    If MsgBox("�˲�����ɾ�����Ӥ����ص�����������Ϣ���������Ƿ�Ҫɾ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    '���Ӥ����Ӧ�ļ����ݵ�ɾ��
'    zl_������������_DelBaby
    mstrSQL = "zl_������������_DelBaby("
'    �ļ�ID_IN ���˻�������.�ļ�ID%TYPE,
    mstrSQL = mstrSQL & mlng�ļ�ID & ","
'    Ӥ��_IN   ���˻�������.�������%TYPE
    mstrSQL = mstrSQL & lngFileIndex & ")"
    Call zlDatabase.ExecuteProcedure(mstrSQL, "zl_������������_DelBaby")
    '�������ˢ��
    mblnVerify = False
    mblnChange = False
    lngFileIndex = lngFileIndex - 1
    If lngFileIndex < 1 Then lngFileIndex = 1
    RaiseEvent AfterFileIndex(mlngFileIndex)
    RaiseEvent AfterDataSave(True)
    Call ShowMe(mFrmParent, mlng�ļ�ID, mlng����ID, mlng��ҳID, mlng����ID, mintӤ��, mstrPrivs, mblnEditable, IIf(mblnBlowup = True, 1, 0))
End Sub

Private Sub cmdWord_Click()
    Dim strInput As String
    '�����ʾ�ѡ����
    
    If cmdWord.Tag = -1 Then
        strInput = txtInput.Text
    Else
        strInput = txt(Val(cmdWord.Tag)).Text
    End If
    strInput = frmEditAssistant.ShowMe(Me, mlng����ID, mlng��ҳID, mintӤ��, strInput)
    
    If cmdWord.Tag = -1 Then
        txtInput.Text = strInput
        Call txtInput_KeyDown(vbKeyReturn, 0)
    Else
        txt(Val(cmdWord.Tag)).Text = strInput
        Call txt_KeyDown(Val(cmdWord.Tag), vbKeyReturn, 0)
    End If
End Sub

Private Sub imgSign_Click()
    Call picSign_Click
End Sub

Private Sub lblCheckSign_Click()
    Call picSign_Click
End Sub

Private Sub picSign_Click()
    '����ǩ����ʷ��¼
    Dim str����ʱ�� As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    vsfSignData.Clear
    str����ʱ�� = VsfData.TextMatrix(VsfData.ROW, mlngActTime)
    gstrSQL = "" & _
        " SELECT A.��¼�� AS ǩ����,NVL(to_char(A.��¼ʱ��,'yyyy-MM-dd hh24:mi:ss'),A.��Ŀ����) AS ǩ��ʱ��,A.��¼���� AS ǩ����Ϣ,A.��¼��� AS ǩ������,A.ID,DECODE(A.��ĿID,NULL,'��Ч','δ��֤') AS ��Ч��,A.��ʼ�汾,NVL(A.��Ŀ���,2) AS ǩ������汾" & vbNewLine & _
        " FROM ���˻�����ϸ A,���˻������� B,���˻����ļ� C" & vbNewLine & _
        " WHERE A.��¼ID=B.ID And B.�ļ�ID=C.ID AND MOD(A.��¼����,10)=5" & vbNewLine & _
        " AND C.ID=[1] AND B.����ʱ��=[2] " & vbNewLine & _
        " Order by A.��Ŀ���� Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡǩ����ʷ��¼", mlng�ļ�ID, CDate(str����ʱ��))
    
    Set vsfSignData.DataSource = rsTemp
    With vsfSignData
        .ColWidth(0) = 1000
        .ColWidth(1) = 1800
        .ColWidth(2) = 0
        .ColWidth(3) = 0
        .ColWidth(4) = 0
        .ColWidth(6) = 0
        .ColWidth(7) = 0
        .ROW = 1: .COL = 5
    End With
    
    rsTemp.Filter = "��Ч��='δ��֤'"
    cmdSignAll.Enabled = rsTemp.RecordCount > 0
    
    picSign.Visible = False
    With picSignCheck
        .Left = VsfData.Left + (VsfData.Width - .Width) / 2
        .Top = VsfData.Top + (VsfData.Height - .Height) / 2
        .ZOrder 0
        .Visible = True
    End With
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdCancl_Click()
    picSignCheck.Visible = False
End Sub

Private Sub cmdSignCur_Click()
    '������֤
    Dim lngLoop As Long
    Dim int�汾 As Integer
    Dim strSource As String, str����ʱ�� As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    If (Val(vsfSignData.TextMatrix(vsfSignData.ROW, 4)) = 0) Then Exit Sub
    '��ȡҪǩ��������
    '------------------------------------------------------------------------------------------------------------------
    int�汾 = vsfSignData.TextMatrix(vsfSignData.ROW, 6)
    str����ʱ�� = VsfData.TextMatrix(VsfData.ROW, mlngActTime)
    Set rsTemp = GetSignData(str����ʱ��, int�汾)
    Do While Not rsTemp.EOF
        For lngLoop = 0 To rsTemp.Fields.Count - 1
            strSource = strSource & CStr(zlCommFun.NVL(rsTemp.Fields(lngLoop).Value, ""))
        Next
        rsTemp.MoveNext
    Loop
    'Debug.Print "��֤ǩ����" & Now & vbCrLf & strSource
    
    '����ǩ��
    Err.Clear
    If gobjESign Is Nothing Then
        On Error Resume Next
        Set gobjESign = CreateObject("zl9ESign.clsESign")
        If Err <> 0 Then Err.Clear
        On Error GoTo 0
        If Not gobjESign Is Nothing Then
            Call gobjESign.Initialize(gcnOracle, glngSys)
        End If
    End If
    If gobjESign Is Nothing Then
        MsgBox "����ǩ������δ����ȷ��װ����֤�������ܼ�����", vbInformation, gstrSysName
        Exit Sub
    End If
    If gobjESign.VerifySignature(strSource, Val(vsfSignData.TextMatrix(vsfSignData.ROW, 4)), 6) Then
        vsfSignData.TextMatrix(vsfSignData.ROW, 5) = "��Ч"
        Call vsfSignData_EnterCell
    End If
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cmdSignAll_Click()
    Dim lngSel As Long
    Dim lngRow As Long, lngRows As Long
    'ȫ����֤
    
    lngSel = vsfSignData.ROW
    vsfSignData.Redraw = flexRDNone
    lngRows = vsfSignData.Rows - 1
    For lngRow = 1 To lngRows
        If (vsfSignData.TextMatrix(lngRow, 5) <> "��Ч") Then
            vsfSignData.ROW = lngRow
            Call cmdSignCur_Click
        End If
    Next
    vsfSignData.ROW = lngSel
    vsfSignData.Redraw = flexRDDirect
End Sub

Private Function ShowSignMarker(Optional ByVal bln�ⲿ As Boolean = False) As Boolean
    '��ʾ��ʷǩ�����
    
    picSign.Visible = False
    picSignCheck.Visible = False
    If Not bln�ⲿ Then
        If VsfData.COL <> mlngSignName Then Exit Function
    End If
    If VsfData.TextMatrix(VsfData.ROW, mlngSigner) = "" Then Exit Function
    
    With picSign
        .Top = VsfData.Top + VsfData.CellTop + VsfData.CellHeight - .Height
        .Left = VsfData.Left + VsfData.CellLeft + 500
        .ZOrder 0
        .Visible = True
    End With
    ShowSignMarker = True
End Function

Private Sub vsfSignData_EnterCell()
    cmdSignCur.Enabled = (vsfSignData.TextMatrix(vsfSignData.ROW, 5) <> "��Ч")
End Sub

Private Function GetSignData(ByVal str����ʱ�� As String, ByVal int�汾 As Integer) As ADODB.Recordset
    On Error GoTo errHand
    Dim rsTemp As New ADODB.Recordset
    
    If int�汾 = 1 Then
        gstrSQL = "" & _
            "Select a.��¼����,a.��Ŀ����,a.��Ŀ���,a.��Ŀ����,a.��Ŀ����,a.��¼����,a.��Ŀ��λ,a.��¼���,a.���²�λ,a.��¼���,a.���Ժϸ�,a.δ��˵��,a.��¼��,a.��¼ʱ��" & vbNewLine & _
            "  From ���˻�����ϸ a, ���˻������� b,���˻����ļ� C" & vbNewLine & _
            " Where c.ID=[1] And b.����ʱ�� =[2]" & vbNewLine & _
            "   And a.��¼id = b.ID and B.�ļ�ID=C.ID and MOD(A.��¼����,10) <>5 and A.��ʼ�汾=1" & vbNewLine & _
            " ORDER BY ��Ŀ���"
    Else
        gstrSQL = "" & _
            "Select a.��¼����,a.��Ŀ����,a.��Ŀ���,a.��Ŀ����,a.��Ŀ����,a.��¼����,a.��Ŀ��λ,a.��¼���,a.���²�λ,a.��¼���,a.���Ժϸ�,a.δ��˵��,a.��¼��,a.��¼ʱ��" & vbNewLine & _
            "  From ���˻�����ϸ a, ���˻������� b,���˻����ļ� C" & vbNewLine & _
            " Where c.ID=[1] And b.����ʱ�� =[2]" & vbNewLine & _
            "   And a.��¼id = b.ID and B.�ļ�ID=C.ID and MOD(A.��¼����,10) <>5" & vbNewLine & _
            "   and (A.��ʼ�汾=[3] or (A.��ʼ�汾 <[3] and A.��ֹ�汾 IS NULL) or (A.��ʼ�汾<[3] and A.��ֹ�汾>[3]))" & vbNewLine & _
            " ORDER BY ��Ŀ���"
    End If
    Set GetSignData = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡָ���汾������", mlng�ļ�ID, CDate(str����ʱ��), int�汾)
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub SignMarker()
    '���ⲿ���������
    If Not ShowSignMarker(True) Then Exit Sub
    Call picSign_Click
End Sub

Private Sub vsfBaby_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim CellRect As RECT
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo errHand
    
    If picBaby.Visible = False Or Val(vsfBaby.RowData(NewRow)) <= 0 Then Exit Sub
    With vsfBaby
        CellRect.Left = .CellLeft + .Left
        CellRect.Top = .CellTop + .Top
        CellRect.Bottom = .CellHeight + CellRect.Top
        CellRect.Right = .CellWidth + CellRect.Left
        cmdDelBaby.Top = CellRect.Top
        cmdDelBaby.Left = CellRect.Right - cmdDelBaby.Width
        cmdDelBaby.Height = CellRect.Bottom - CellRect.Top
        cmdDelBaby.Visible = True
        cmdDelBaby.Enabled = True
        '��һ���ļ�����ɾ��
        If .RowData(NewRow) = 1 Then cmdDelBaby.Visible = False: cmdDelBaby.Enabled = False: Exit Sub
        '�ļ�ֻ�ܴӺ���ǰɾ��
        strSQL = " SELECT 1" & vbNewLine & _
            " FROM ���˻����ļ� A, ���˻������� B, ���˻�����ϸ C" & vbNewLine & _
            " WHERE A.ID = B.�ļ�ID AND B.ID = C.��¼ID AND A.ID = [1] AND A.����ID = [2] AND A.��ҳID = [3] AND B.������� > [4]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "�ļ�ɾ�����", mlng�ļ�ID, mlng����ID, mlng��ҳID, Val(.RowData(NewRow)))
        If rsTemp.RecordCount > 0 Then cmdDelBaby.Visible = False: cmdDelBaby.Enabled = False
    End With
    
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub VsfData_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    If Not mblnInit Then Exit Sub
    Call InitCons
    If OldLeftCol <> NewLeftCol Then
        vsfHead.LeftCol = NewLeftCol
        VsfData.LeftCol = vsfHead.LeftCol
    End If
End Sub

Private Sub VsfData_DblClick()
    Call vsfdata_KeyDown(Asc("A"), 0)
End Sub

Private Sub VsfData_EnterCell()
    Dim strCols As String
    Dim strName As String
    Dim intMax As Integer
    Dim lngStart As Long
    On Error Resume Next
    
    '��������ʾ��¼��ؼ�
    Select Case mintType
    Case 0, 3
        picInput.Visible = False
    Case 1, 2
        lstSelect(mintType - 1).Visible = False
    Case 4, 5
        picDouble.Visible = False
    Case 6
        picMutilInput.Visible = False
    Case 7
        picDoubleChoose.Visible = False
    End Select
    cmdWord.Visible = False
    
    'δ������в�����¼������
    mintType = -1
    If mblnInit = False Then Exit Sub
    
    Call ShowSignMarker
    
    If InStr(1, mstrPrivs, "����ͼ��ͼ") = 0 Then Exit Sub
    
    If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "" Then
        lngStart = VsfData.ROW
    Else
        lngStart = GetStartRow(VsfData.ROW)
    End If
    
    If mblnVerify Then  '�������mblnShow�ж���������
        If VsfData.COL = mlngChoose Then Call vsfdata_KeyDown(vbKeySpace, 0): Exit Sub
        If VsfData.COL = mlngDate Or VsfData.COL = mlngTime Then Exit Sub
        If Val(VsfData.TextMatrix(lngStart, mlngRecord)) = 0 Then Exit Sub
        If VsfData.TextMatrix(lngStart, mlngSigner) = "" Then Exit Sub
        If VsfData.Cell(flexcpChecked, lngStart, mlngChoose) = flexTSUnchecked Then Exit Sub 'û��ѡ�еļ�¼���ܱ༭
    Else
        '��ǩ��������ֻ������ǩ״̬���޸�
        If InStr(1, VsfData.TextMatrix(lngStart, mlngSigner), "/") <> 0 Then
            RaiseEvent AfterRowColChange("����ǩ�����ݲ�����༭��", True, mblnSign, mblnArchive)
            Exit Sub
        End If
        
        '--ֻҪ��ǩ�������ݾͲ������޸�
        If VsfData.TextMatrix(lngStart, mlngSigner) <> "" Then
            RaiseEvent AfterRowColChange("��ǩ�������ݲ������ٴα༭����ȡ��ǩ�������ԣ�", True, mblnSign, mblnArchive)
            Exit Sub
        End If
        
        'Ĭ��ǩ�����뱣������ͬ,�������޸����˻����¼Ȩ�޵Ĳ���Ա,�������޸����˵�����
        strName = VsfData.TextMatrix(lngStart, IIf(mlngOperator = -1, VsfData.Cols - 1, mlngOperator))
        If strName <> "" Then
            If strName <> gstrUserName And _
                InStr(1, mstrPrivs, "���˻����¼") = 0 Then
                RaiseEvent AfterRowColChange("��û���޸����˻����¼���ݵ�Ȩ�ޣ�ԭ����Ա:" & strName, True, mblnSign, mblnArchive)
                Exit Sub
            End If
        End If
    End If
    If mblnArchive Then Exit Sub
    If Not mblnShow Or Not mblnEditable Then Exit Sub

    'δ����Ŀ�в������޸�
    If InStr(1, "," & mstrCOLNothing & ",", "," & VsfData.COL & ",") <> 0 Then Exit Sub
    
    'ͬ�������в�����༭
    strCols = GetSynItems(2, intMax)
    If strCols <> "" Then
        '����ͬ�����ݵ���,������ʱ���ǲ������޸ĵ�
        If VsfData.COL = mlngDate Or VsfData.COL = mlngTime Then Exit Sub
        strCols = "," & strCols & ","
        If InStr(1, strCols, "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") <> 0 Then Exit Sub
    End If
    
    If VsfData.COL <= mlngNoEditor - 1 Then Call ShowInput
    'Debug.Print txtInput.Text
    '�ÿؼ���ý���
    Select Case mintType
    Case 0, 3
        picInput.SetFocus
    Case 1, 2
        lstSelect(mintType - 1).SetFocus
    Case 4, 5
        picDouble.SetFocus
    Case 6
        picMutilInput.SetFocus
    End Select
    'Debug.Print txtInput.Text
End Sub

Private Sub VsfData_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strInfo As String
    Dim strCols As String
    Dim intMax As Integer
    If mblnInit = False Then Exit Sub
    If mblnEditable = False Then Exit Sub
    If OldRow = NewRow And OldCol = NewCol Then Exit Sub
    On Error GoTo errHand
    
    'ѡ����,ͬ��������ֱ���˳�,����˴������ʾ��Ϣ
    If NewCol = mlngChoose Then Exit Sub
    strCols = GetSynItems(2, intMax)
    If strCols <> "" Then
        strCols = "," & strCols & ","
        If InStr(1, strCols, "," & NewCol - (cHideCols + VsfData.FixedCols - 1) & ",") <> 0 Then Exit Sub
    End If
    
    '��ʾ��ǰ��Ŀ�������Ϣ
    mrsSelItems.Filter = "��=" & NewCol - (cHideCols + VsfData.FixedCols - 1)
    If mrsSelItems.RecordCount <> 0 Then
        mrsItems.Filter = "��Ŀ���=" & mrsSelItems!��Ŀ���
        If mrsItems.RecordCount <> 0 Then
            If NVL(mrsItems!��Ŀֵ��) <> "" Then
                If mrsItems!��Ŀ���� = 0 Then
                    strInfo = "��Ч��Χ:" & Split(mrsItems!��Ŀֵ��, ";")(0) & "��" & Split(mrsItems!��Ŀֵ��, ";")(1)
                Else
                    strInfo = "��Ч��Χ:" & mrsItems!��Ŀֵ��
                End If
            Else
                strInfo = ""
            End If
        End If
    End If
    mrsSelItems.Filter = 0
    mrsItems.Filter = 0
    
    '����Ƿ���ǩ��
    If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "" Then
        intMax = VsfData.ROW
    Else
        intMax = GetStartRow(VsfData.ROW)
    End If
    mblnSign = (VsfData.TextMatrix(intMax, mlngSigner) <> "")
    
    RaiseEvent AfterRowColChange(strInfo, False, mblnSign, mblnArchive)
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub vsfData_DrawCell(ByVal hDC As Long, ByVal ROW As Long, ByVal COL As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
    Call DrawCell(hDC, ROW, COL, Left, Top, Right, Bottom, Done)
End Sub

Private Sub vsfdata_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngStart As Long
    Dim intLevel As Integer
    Dim strField As String, strKey As String, strValue As String
    On Error GoTo errHand
    
    If KeyCode = vbKeyReturn Then
        If Not mblnShow And VsfData.COL = mlngDate Then
            mblnShow = True
            Call VsfData_EnterCell
        Else
            Call MoveNextCell
        End If
    ElseIf KeyCode = vbKeySpace And mblnVerify Then
        'ֻ��ѡ��ʼ��
        lngStart = GetStartRow(VsfData.ROW)
        If VsfData.TextMatrix(lngStart, mlngTime) = "" Then Exit Sub
        
        '��ǩʱ,��ǰ��¼��ǩ��,�Ҳ���Ա��ǩ��������ϴ�ǩ������߲�����
        If VsfData.TextMatrix(lngStart, mlngSignLevel) = "" Then
            RaiseEvent AfterRowColChange("�����ݻ�δǩ�������ܽ�����ǩ��", True, mblnSign, mblnArchive)
            Exit Sub
        Else
            intLevel = Val(VsfData.TextMatrix(lngStart, mlngSignLevel)) + 1
        End If
        If mintVerify >= intLevel Then
            RaiseEvent AfterRowColChange("���ļ���Ҫ���ϴ���ǩ�˵ļ���߲��ܹ�ѡ�ü�¼��", True, mblnSign, mblnArchive)
            Exit Sub
        End If
        
        VsfData.Cell(flexcpChecked, lngStart, mlngChoose) = IIf(VsfData.Cell(flexcpChecked, lngStart, mlngChoose) = flexTSChecked, flexTSUnchecked, flexTSChecked)
        '�����޸ļ�¼�Ա�ͬ��
        strField = "ID|ҳ��|�к�|�к�|��¼ID|����|ɾ��"
        strKey = mintҳ�� & "," & lngStart & "," & mlngChoose
        strValue = strKey & "|" & mintҳ�� & "|" & lngStart & "|" & mlngChoose & "|" & Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & VsfData.Cell(flexcpChecked, lngStart, mlngChoose) & "|1"
        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
    Else
        If Not (KeyCode = vbKeyLeft Or KeyCode = vbKeyRight Or KeyCode = vbKeyUp Or KeyCode = vbKeyDown Or Shift <> 0) Then
            mblnShow = True
            Call VsfData_EnterCell
        End If
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub InitVariable()
    '�������
    mlngDate = -1
    mlngTime = -1
    mlngOperator = -1
    mlngSigner = -1
    mlngSignTime = -1
    mlngSignName = -1
    mlngRecord = -1
    mlngNoEditor = -1
    
    mblnChange = False
    mblnShow = False
    mblnSign = False
    mblnArchive = False
    mblnEditAssistant = False
    mblnVerify = False
    
End Sub

Private Sub InitCons()
    '��������ؼ�
    picInput.Visible = False
    lstSelect(0).Visible = False
    lstSelect(1).Visible = False
    picDouble.Visible = False
    picMutilInput.Visible = False
    cmdWord.Visible = False
    picBaby.Visible = False
End Sub

Private Function InitMenuBar() As Boolean
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    Dim cbrMenuBar As CommandBarControl
    Dim cbrControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    Dim cbrPop As CommandBarControl
    Dim cbrToolBar As CommandBar
    Dim rs As ADODB.Recordset
    Dim objExtendedBar As CommandBar
    
    On Error GoTo errHand
    
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    cbsThis.ActiveMenuBar.Title = "�˵���"
    cbsThis.ActiveMenuBar.Visible = False
    
    Set cbsThis.Icons = zlCommFun.GetPubIcons
    With cbsThis.Options
        .AlwaysShowFullMenus = False
        .ShowExpandButtonAlways = False
        .UseDisabledIcons = True
        .SetIconSize True, 16, 16
        .LargeIcons = True
    End With

    '------------------------------------------------------------------------------------------------------------------
    '�˵�����
    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    
        '------------------------------------------------------------------------------------------------------------------
        '����������
        Set cbrToolBar = cbsThis.Add("��׼", xtpBarTop)
        cbrToolBar.ShowTextBelowIcons = False
        cbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
        cbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
        With cbrToolBar.Controls
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_FileMan, "Ӥ��"): cbrControl.ToolTipText = "�ļ�����"
            Set cbrCustom = .Add(xtpControlCustom, conMenu_View_Option, "")
            cbrCustom.Flags = xtpFlagAlignLeft
            picTmp.Visible = True
            cbrCustom.Handle = picTmp.Hwnd
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Copy, "����"): cbrControl.ToolTipText = "����(Ctrl+C)": cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PASTE, "ճ��"):  cbrControl.ToolTipText = "ճ��(Ctrl+V)"
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Clear, "���"):   cbrControl.ToolTipText = "���"
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SPECIALCHAR, "�������"):  cbrControl.ToolTipText = "�����������(Ctrl+D)": cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Word, "�ʾ�ѡ��"):  cbrControl.ToolTipText = "�ʾ�ѡ��(Ctrl+W)"
            'Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Brief, "С��"): cbrControl.ToolTipText = "С��"
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Append, "������Ϣ"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "������Ϣ"
        End With
    
        For Each cbrControl In cbrToolBar.Controls
            If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel Then
                cbrControl.Style = xtpButtonIconAndCaption
            End If
        Next
    
         '�����
        With cbsThis.KeyBindings
            .Add FCONTROL, Asc("C"), conMenu_Edit_Copy
            .Add FCONTROL, Asc("V"), conMenu_Edit_PASTE
            .Add FCONTROL, Asc("D"), conMenu_Edit_SPECIALCHAR
            .Add FCONTROL, Asc("W"), conMenu_Edit_Word
            .Add FCONTROL, Asc("S"), conMenu_Save
        End With
    
    InitMenuBar = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function CheckTxtTime(objText As TextBox) As String
    Dim strInput As String
    Dim strHour As String, strMin As String
    '���ʱ��¼���Ƿ�Ϸ�������������
    
    If Trim(objText.Text) <> "" Then
        strInput = Trim(objText.Text)
        If InStr(1, strInput, ":") > 0 Then
            strHour = Split(strInput, ":")(0)
            strMin = Split(strInput, ":")(1)
        ElseIf InStr(1, strInput, "��") > 0 Then
            strHour = Split(strInput, "��")(0)
            strMin = Split(strInput, "��")(1)
        Else
            strHour = strInput
            strMin = "00"
        End If
        strHour = Format(strHour, "00")
        strMin = Format(strMin, "00")
        If Not IsNumeric(strHour) Then
            RaiseEvent AfterRowColChange("��ʼʱ���к��зǷ��ַ���", True, mblnSign, mblnArchive)
            Exit Function
        End If
        If Val(strHour) < 0 Or Val(strHour) > 23 Then
            RaiseEvent AfterRowColChange("��ʼʱ�㲻��ȷ��СʱֵӦ��>0��С��24��", True, mblnSign, mblnArchive)
            Exit Function
        End If
        If Not IsNumeric(strMin) Then
            RaiseEvent AfterRowColChange("��ʼʱ���к��зǷ��ַ���", True, mblnSign, mblnArchive)
            Exit Function
        End If
        If Val(strMin) < 0 Or Val(strMin) > 59 Then
            RaiseEvent AfterRowColChange("��ʼʱ�㲻��ȷ������ֵӦ��>0��С��60��", True, mblnSign, mblnArchive)
            Exit Function
        End If
        strInput = strHour & ":" & strMin
    End If
    CheckTxtTime = strInput
End Function

Private Function CheckTime(ByVal lngRow As Long, ByVal lng����ID As Long, ByVal lng��ҳID As Long, _
    ByVal strTime As String, ByVal strCurTime As String, ByRef strMsg As String) As Boolean
    Dim blnMsg As Boolean
    Dim blnExist As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '���ݷ���ʱ������ڲ���ʱ��Ͳ��˷�Χ��
    '˵��:���ڲ��˿���Ժǰ�Ϳ�ʼ���������Բ���ʱ�����С�ڲ�����Ժʱ������
    blnMsg = (strMsg <> "")
    
    '����ļ���ʼ,����ʱ��
    If strTime < Format(mstr��ʼʱ��, "yyyy-MM-dd HH:mm") Or strTime > Format(mstr����ʱ��, "yyyy-MM-dd HH:mm") Then
        strMsg = "����ʱ��[" & strTime & "]���ڿ�ʼʱ��[" & Format(mstr��ʼʱ��, "YYYY-MM-DD HH:mm") & "]������ʱ��[" & Format(mstr����ʱ��, "YYYY-MM-DD HH:mm") & "]֮��"
        GoTo exitHand
    End If
    
    '������ڶ���ļ�,��һ�ļ���ʱ�䲻�ܴ�����һ�ļ���ʼʱ��
    If cboBaby.ListCount > 1 And cboBaby.ListIndex < cboBaby.ListCount - 1 Then
        gstrSQL = "Select 1 From ���˻����ļ� A,���˻������� B" & _
            "   Where A.ID=B.�ļ�ID And A.ID=[1] And A.����ID=[2] And A.��ҳID=[3] And A.Ӥ��=[4] AND B.�������=[5] And B.����ʱ��<=[6]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "���ݼ��", mlng�ļ�ID, mlng����ID, mlng��ҳID, mintӤ��, cboBaby.ItemData(cboBaby.ListIndex + 1), CDate(strTime))
        If rsTemp.RecordCount > 0 Then
            strMsg = "��" & lngRow & "�еķ���ʱ��" & Format(strTime, "YYYY-MM-DD HH:mm") & "���󣬲��ܴ�����һӤ���ļ��Ŀ�ʼʱ�䣡"
            GoTo exitHand
        End If
    End If
    
    
    If Format(strTime, "YYYY-MM-DD HH:mm") > Format(DateAdd("d", mintPreDays, CDate(strCurTime)), "YYYY-MM-DD HH:mm") Then
        strMsg = "¼��������ѳ�������[����¼��������" & mintPreDays & "��]��ָ���ķ�Χ��"
        GoTo exitHand
    End If
            
    '���ݲ��˱䶯��¼���м��
    gstrSQL = " Select   ��ʼԭ��,����ID,to_char(��ʼʱ��,'yyyy-MM-dd hh24:mi') AS ��ʼʱ��,to_char(NVL(��ֹʱ��,sysDate+" & mintPreDays & "),'yyyy-MM-dd hh24:mi') AS ��ֹʱ�� " & _
              " From ���˱䶯��¼ " & _
              " Where ����ID=[1] And ��ҳID=[2]" & _
              " Order by ��ʼʱ��,��ʼԭ��"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ��ǰ������Чʱ�䷶Χ", lng����ID, lng��ҳID)
    With rsTemp
        .Filter = "����ID=" & mlng����ID
        Do While Not .EOF
            If Format(strTime, "YYYY-MM-DD HH:mm") <= Format(!��ֹʱ��, "YYYY-MM-DD HH:mm") Then
                blnExist = True
                Exit Do
            End If
            .MoveNext
        Loop
        .Filter = 0
        '�ҵ��˾��˳�
        If blnExist Then
            If Not IsAllowInput(lng����ID, lng��ҳID, strTime, strCurTime) Then
                strMsg = "��" & lngRow & "�еķ���ʱ��" & strTime & "����[�������ݲ�¼����Чʱ��:" & glngHours & "Сʱ]"
                GoTo exitHand
            End If
            
            CheckTime = True
            Exit Function
        End If
        
        .Filter = 0
        '˵��¼�������ʱ�䲻�ڲ���ʱ��Ͳ��˵�ǰ�������ʱ�䷶Χ��
        strMsg = "��" & lngRow & "�еķ���ʱ��" & strTime & "����[���ڲ��̿�ʼʱ������ǰ��������ʱ�����Ч��Χ��]"
        GoTo exitHand
    End With
    
    CheckTime = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
exitHand:
    rsTemp.Filter = 0
    If blnMsg Then MsgBox strMsg, vbInformation, gstrSysName
End Function

Private Function CheckInput(strReturn As String, strInfo As String) As Boolean
    Dim i As Integer, j As Integer
    Dim strOrders As String, strText As String
    '���¼�����ݵĺϷ���(����Ҳ��Ϊ��һ���ַ�,���ǵ�������Ŀ�ȴ��ڲ���\�������Ϣ)
    '���ص�����,���һ�а󶨶����Ŀ,�Ե�������Ϊ�ָ���
    
    'mintType:0=�ı���¼��;1=��ѡ;2=��ѡ;3=ѡ��;4-Ѫѹ��һ�а���������Ŀ,���ʽ����Ѫѹ��������Ŀ;5=һ�а���������Ŀ�Ҿ���ѡ����Ŀ;
    '6=һ�а�N����Ŀ,�ֹ�¼��
    Select Case mintType
    Case 0
        strText = txtInput.Text
        strOrders = txtInput.Tag
    Case 1, 2   '���
        If mintType = 1 Then
            If InStr(1, lstSelect(mintType - 1).Text, "-") <> 0 Then
                strText = Split(lstSelect(mintType - 1).Text, "-")(1)
            Else
                strText = ""
            End If
        Else
            j = lstSelect(mintType - 1).ListCount
            For i = 1 To j
                If lstSelect(mintType - 1).Selected(i - 1) Then
                    strText = strText & "," & Split(lstSelect(mintType - 1).List(i - 1), "-")(1)
                End If
            Next
            If strText <> "" Then strText = Mid(strText, 2)
        End If
        strOrders = lstSelect(mintType - 1).Tag
    Case 4
        strText = txtUpInput.Text & "'" & txtDnInput.Text
        strOrders = txtUpInput.Tag & "'" & txtDnInput.Tag
    Case 6
        j = txt.Count
        For i = 1 To j
            strText = strText & "'" & txt(i - 1).Text
            strOrders = strOrders & "'" & txt(i - 1).Tag
        Next
        If strText <> "" Then
            strText = Mid(strText, 2)
            strOrders = Mid(strOrders, 2)
        End If
    Case 3      '���
        strText = lblInput.Caption
    Case 5      '���
        strText = lblUpInput.Caption & "/" & lblDnInput.Caption
    Case 7
        strText = cboChoose(0).Text & "/" & cboChoose(1).Text
    End Select
    If Val(strOrders) <> 0 Then
        If Not CheckValid(strText, strOrders, strInfo) Then Exit Function
    ElseIf VsfData.COL = mlngDate Or VsfData.COL = mlngTime Then
        If Not CheckDateTime(strText, strInfo) Then Exit Function
    End If
    
    strReturn = strText
    CheckInput = True
End Function

Private Function CheckDateTime(strText As String, strInfo As String) As Boolean
    Dim blnCheck As Boolean, blnExist As Boolean
    Dim strCurrDate As String
    Dim strDate As String
    Dim lngRow As Long, strTmpTime As String
    
    On Error GoTo errHand
    
    strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm")
    
    If VsfData.COL = mlngDate Then
        If mblnDateAd Then
            If Trim(strText) = "" Then
                strInfo = "���ڲ���Ϊ�գ�"
                Exit Function
            End If
            If InStr(1, strText, "/") = 0 Then
                strInfo = "���ڸ�ʽ������1��12�գ�12/01"
                Exit Function
            End If
            
            strDate = Mid(strCurrDate, 1, 5) & ToStandDate(strText)
                
            If Not IsDate(strDate) Then
                strInfo = "¼������ݲ��ǺϷ������ڣ���1��12�գ�12/01"
                Exit Function
            End If
        Else
            If Trim(strText) = "" Then
                strInfo = "���ڲ���Ϊ�գ�"
                Exit Function
            End If
            If Not IsDate(strText) Then
                strInfo = "¼������ݲ��ǺϷ������ڣ��磺2011-01-12"
                Exit Function
            End If
            strDate = Format(strText, "yyyy-MM-dd")
        End If
        
        If VsfData.TextMatrix(VsfData.ROW, mlngTime) <> "" Then
            If mblnDateAd Then
                strDate = Mid(GetDateAdCurrDate(VsfData.TextMatrix(VsfData.ROW, mlngTime)), 1, 5) & ToStandDate(strText)
            End If
            strDate = Format(strDate & " " & VsfData.TextMatrix(VsfData.ROW, mlngTime), "YYYY-MM-DD HH:mm")
            blnCheck = True
        End If
    Else
        If Trim(strText) = "" Then
            strInfo = "ʱ�䲻��Ϊ�գ�"
            Exit Function
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
            Exit Function
        End If
        If Mid(strText, 1, 2) < 0 Or Mid(strText, 1, 2) > 23 Then
            strInfo = "¼���ʱ���ʽ�Ƿ���[СʱӦ��0��23֮��]"
            Exit Function
        End If
        If Mid(strText, 4, 2) < 0 Or Mid(strText, 4, 2) > 59 Then
            strInfo = "¼���ʱ���ʽ�Ƿ���[����Ӧ��0��59֮��]"
            Exit Function
        End If
        
        'û������Ĭ�Ͻ������ڼ���
        If mblnDate = False Then
            If Format(strText, "HH:mm") >= Format(mstr��ʼʱ��, "HH:mm") Then
                strDate = Format(mstr��ʼʱ��, "YYYY-MM-DD")
            Else
                strDate = Format(CDate(mstr��ʼʱ��) + 1, "YYYY-MM-DD")
            End If
            VsfData.TextMatrix(VsfData.ROW, mlngDate) = strDate
        End If
        '���кϷ��Լ��
        If VsfData.TextMatrix(VsfData.ROW, mlngDate) <> "" Then
            strTmpTime = GetDateAdCurrDate(strText)
            strDate = VsfData.TextMatrix(VsfData.ROW, mlngDate)
            If mblnDateAd Then
                strDate = Mid(strTmpTime, 1, 5) & ToStandDate(strDate)
            Else
                strDate = Format(VsfData.TextMatrix(VsfData.ROW, mlngDate), "yyyy-MM-dd")
            End If
            strDate = Format(strDate & " " & strText, "YYYY-MM-DD HH:mm")
            
            If Not IsDate(strDate) Then
                strInfo = "¼������ݲ��ǺϷ������ڣ��磺12:01"
                Exit Function
            End If
        
            blnCheck = True
        End If
    End If
    
    If blnCheck Then
        '���¼��������Ƿ��Ѿ�����
        For lngRow = VsfData.FixedRows To VsfData.Rows - 1
            If VsfData.TextMatrix(lngRow, mlngRowCount) Like "*|1" And lngRow <> VsfData.ROW And VsfData.RowHidden(lngRow) = False Then
                If VsfData.TextMatrix(lngRow, mlngDate) <> "" And VsfData.TextMatrix(lngRow, mlngTime) <> "" Then
                    If mblnDateAd Then
                        strTmpTime = Mid(GetDateAdCurrDate(VsfData.TextMatrix(lngRow, mlngTime)), 1, 5) & ToStandDate(VsfData.TextMatrix(lngRow, mlngDate))
                        strTmpTime = strTmpTime & " " & VsfData.TextMatrix(lngRow, mlngTime)
                    Else
                        strTmpTime = VsfData.TextMatrix(lngRow, mlngDate) & " " & VsfData.TextMatrix(lngRow, mlngTime)
                    End If
                    If Format(strTmpTime, "YYYY-MM-DD HH:mm") = Format(strDate, "YYYY-MM-DD HH:mm") Then
                        strInfo = "��¼���ʱ���Ѿ�������ʷ���ݣ�����λ�ã��� " & (lngRow - VsfData.FixedRows + 1) & " ��!"
                        Exit Function
                    End If
                End If
            End If
        Next lngRow
        '���ݷ���ʱ�䲻���ڵ�ǰ����Ա�������ҵ���Чʱ����ǰ
        If Not CheckTime(VsfData.ROW, mlng����ID, mlng��ҳID, strDate, strCurrDate, strInfo) Then
            Exit Function
        End If
    End If
    
    CheckDateTime = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetDateAdCurrDate(ByVal strTime As String) As String
'���ڴ��ڶԽ��ߣ���ȡ��ǰʱ��
    Dim strDate As String
    If Format(strTime, "HH:mm") >= Format(mstr��ʼʱ��, "HH:mm") Then
        strDate = Format(mstr��ʼʱ��, "YYYY-MM-DD")
    Else
        strDate = Format(CDate(mstr��ʼʱ��) + 1, "YYYY-MM-DD")
    End If
    GetDateAdCurrDate = Format(strDate & " " & Format(strTime, "HH:mm") & ":00", "yyyy-MM-dd HH:mm")
End Function

Private Function CheckValid(strReturn As String, ByVal strOrders As String, strInfo As String) As Boolean
    Dim arrData, arrOrder
    Dim blnCheck As Boolean
    Dim i As Integer, j As Integer
    Dim dblMin As Double, dblMax As Double
    Dim strText As String, strName As String, strFormat As String, strFormat1 As String
    
    '���и�ʽ��װ����
    mrsSelItems.Filter = "��=" & VsfData.COL - (cHideCols + VsfData.FixedCols - 1)
    If mrsSelItems.RecordCount <> 0 Then
        '�д��е�δ���ж���
        strFormat = NVL(mrsSelItems!��ʽ)   '{P[����]C}{...}
        strFormat1 = strFormat
    End If
    mrsSelItems.Filter = 0
    
    '�������
    arrData = Split(strReturn, "'")
    arrOrder = Split(strOrders, "'")
    j = UBound(arrData)
    For i = 0 To j
        strText = arrData(i)
        strOrders = arrOrder(i)
        If Val(strOrders) <> 0 Then
            mrsItems.Filter = "��Ŀ���=" & strOrders
            strName = mrsItems!��Ŀ����
            If strText <> "" Then
                blnCheck = True
                '�����������Ŀ,�������Ĳ����������򲻼��
                If mrsItems!��Ŀ��� >= 1 And mrsItems!��Ŀ��� <= 3 Then
                    If Not IsNumeric(Trim(strText)) Then
                        blnCheck = False
                    End If
                End If
                
                If blnCheck Then
                    If mrsItems!��Ŀ���� = 0 And InStr(1, "0,4", mrsItems!��Ŀ��ʾ) <> 0 Then
                        strText = Val(strText)
                        If NVL(mrsItems!��ĿС��, 0) <> 0 Then   '��������ͨ���ؼ���MaxLength�����Ƶ�
                            If InStr(1, strText, ".") <> 0 Then strText = Mid(strText, 1, InStr(1, strText, ".") - 1)
                            If Len(strText) > mrsItems!��Ŀ���� Then
                                mrsItems.Filter = 0
                                strInfo = "[" & strName & "]¼������ݳ����˺Ϸ����ȣ�"
                                Exit Function
                            End If
                            
                            strText = Val(arrData(i))
                            If InStr(1, strText, ".") <> 0 Then
                                strText = Mid(strText, InStr(1, strText, ".") + 1)
                                If Len(strText) > mrsItems!��ĿС�� Then
                                    mrsItems.Filter = 0
                                    strInfo = "[" & strName & "]¼���С�����ֳ����˺Ϸ����ȣ�"
                                    Exit Function
                                End If
                            End If
                            strText = Val(arrData(i))
                        End If
                        If mrsItems!��Ŀ��ʾ = 0 Then
                            If Not IsNull(mrsItems!��Ŀֵ��) Then
                                dblMin = Val(Split(mrsItems!��Ŀֵ��, ";")(0))
                                dblMax = Val(Split(mrsItems!��Ŀֵ��, ";")(1))
                                If Not (Val(strText) >= dblMin And Val(strText) <= dblMax) Then
                                    mrsItems.Filter = 0
                                    strInfo = "[" & strName & "]¼������ݲ���" & Format(dblMin, "#0.00") & "��" & Format(dblMax, "#0.00") & "����Ч��Χ��"
                                    Exit Function
                                End If
                            End If
                        End If
                    Else
                        If LenB(StrConv(strText, vbFromUnicode)) > mrsItems!��Ŀ���� Then
                            strInfo = "[" & strName & "]¼������ݳ�������󳤶ȣ�" & mrsItems!��Ŀ���� & "��"
                            mrsItems.Filter = 0
                            Exit Function
                        End If
                    End If
                End If
                strFormat = Replace(strFormat, "[" & strName & "]", strText)
            Else
                'ɾ������Ŀ
                If InStr(1, "," & mstrCOLNothing & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 Then
                    Call SubstrPro(strFormat, strName)
                Else
                    '����Ŀ������ʱ,�����ǰ�о��жԽ�������,�����
                    strFormat = Replace(strFormat, "[" & strName & "]", strText)
                End If
            End If
        Else
            strFormat = strReturn
        End If
    Next
    If j = -1 Then
        strOrders = arrOrder(i)
        If Val(strOrders) <> 0 Then
            mrsItems.Filter = "��Ŀ���=" & strOrders
            strName = mrsItems!��Ŀ����
            strFormat = Replace(strFormat, "[" & strName & "]", strText)
        End If
    End If
    mrsItems.Filter = 0
    
    strFormat = Replace(strFormat, "{", "")
    strFormat = Replace(strFormat, "}", "")
    If InStr(1, "," & mstrCOLNothing & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 Then
        If strFormat = SubstrFormat(strFormat1, arrOrder) Then strFormat = ""
    End If
    
    If IsNumeric(strFormat) Then
        If Val(strFormat) < 1 And Val(strFormat) > 0 Then strFormat = "0" & strFormat
    End If
    strReturn = strFormat
    
    CheckValid = True
End Function

Public Function SubstrFormat(ByVal strData As String, ByVal arrOrder As Variant) As String
    '��ȡ����Ŀ��ǰ��׺����
    Dim i As Integer
    Dim strOrders As String, strName As String
    For i = 0 To UBound(arrOrder)
        strOrders = CStr(arrOrder(i))
        If Val(strOrders) <> 0 Then
            mrsItems.Filter = "��Ŀ���=" & strOrders
            strName = mrsItems!��Ŀ����
        End If
        strData = Replace(strData, "[" & strName & "]", "")
    Next i
    strData = Replace(strData, "{", "")
    strData = Replace(strData, "}", "")
    
    SubstrFormat = strData
End Function

Public Function SubstrVal(ByVal strData As String, ByVal strFormat As String, ByVal strName As String, intPos As Integer) As String
    Dim i As Integer, j As Integer, l As Integer, r As Integer
    Dim strQZ As String, strHZ As String
    '����ǰһ����Ŀ�ĺ�׺����+��ǰ��Ŀ��ǰ׺���ŵ�λ��
    
    If strData = "" Then Exit Function
    strData = strData
    j = Len(strFormat)
    l = InStr(1, strFormat, "[" & strName & "]")
    If l = 0 Then Exit Function
    '�õ�ǰ׺
    For i = l To 1 Step -1
        If Mid(strFormat, i, 1) = "{" Then Exit For
    Next
    strQZ = Mid(strFormat, i + 1, l - i - 1)
    '�ҵ�����Ŀ��ʽ���еĽ�������
    i = l + Len(strName) + 2
    For r = i To j
        If Mid(strFormat, r, 1) = "}" Then Exit For
    Next
    '�õ���׺
    strHZ = Mid(strFormat, i, r - i)
    '�����׺Ϊ��,�������Ѱ����һ����Ŀ��ǰ׺����
    If strHZ = "" And r < j Then
        For r = r + 1 To j
            If Mid(strFormat, r, 1) = "[" Then Exit For
        Next
        strHZ = Mid(strFormat, InStr(i, strFormat, "{") + 1, r - InStr(i, strFormat, "{") - 1)
    End If
    'ȡ��ָ����Ŀ���������ݴ�
    If strHZ <> "" Then
        j = InStr(intPos, strData, strHZ) '��Ϊ������ȡ��,���ǵ��ָ���������ͬ�����,��¼��һ�ε����λ��,�´δ����λ������ȡ����
        If j = 0 Then
            '�п����м���ڻس����з�
            j = InStr(intPos, Replace(strData, vbCrLf, ""), strHZ)
            If j = 0 Then Exit Function
        End If
    End If
    strData = Mid(strData, intPos)
    'ǰ׺Ϊ��,������ǰѰ����һ����Ŀ�ĺ�׺����
'    If strQZ = "" And i > 1 And intPos > 1 Then
'        For i = i - 1 To 1 Step -1
'            If Mid(strFormat, i, 1) = "]" Then Exit For
'        Next
'        strQZ = Mid(strFormat, i + 1, InStr(i, strFormat, "}") - i - 1)
'    End If
    
    SubstrVal = SubstrAnaly(strData, strHZ, strQZ)
    intPos = intPos + Len(strQZ & SubstrVal & strHZ)
    '�������������ȥ���س����з�����,������ַ�����ԭ������
'    If strHZ <> "" Then
'
'        strData = Mid(strData, 1, InStr(1, Replace(strData, vbCrLf, ""), strHZ) - 1) '��������Ŀ�������
'        intPOS = i + Len(strHZ)
'    End If
'    If strQZ <> "" Then strData = Mid(strData, InStr(1, strData, strQZ) + Len(strQZ)) '��������Ŀ�������
'    SubstrVal = strData ' Replace(strData, vbCrLf, "")
End Function

Private Function SubstrAnaly(ByVal strData As String, ByVal strHZ As String, ByVal strQZ As String) As String
    Dim strText As String
    Dim strCompare As String           '�Աȴ�
    Dim intLen As Integer, intActLen As Integer           'ǰ׺/��׺�ĳ���
    Dim intPos As Integer, intEnd As Integer
    Dim lngASC As Long
    Dim blnFind As Boolean
    '�����س����з�����,�ո����±ȶ�
    
    strText = strData
    If strHZ <> "" Then
        '�Ѻ�׺ȥ��
        strHZ = Replace(strHZ, vbCrLf, "")
        intEnd = Len(strText)
        intLen = Len(strHZ)
        For intPos = 1 To intEnd
            lngASC = Asc(Mid(strText, intPos, 1))
            intActLen = intActLen + 1
            If Not (lngASC = 13 Or lngASC = 10) Then
                If lngASC = 32 Then
                    strCompare = ""
                    intActLen = 0
                Else
                    strCompare = strCompare & Mid(strText, intPos, 1)
                End If
                If Len(strCompare) = intLen Then
                    If strCompare = strHZ Then
                        blnFind = True
                        intPos = intPos - intActLen
                    Else
                        strCompare = ""
                        intPos = intPos - intActLen + 1
                        intActLen = 0
                    End If
                End If
            End If
            If blnFind Then Exit For
        Next
        '�϶���
        strText = Mid(strText, 1, intPos)
    End If
    
    '��ȥ��ǰ׺
    If strQZ <> "" Then
        If InStr(1, strText, strQZ) = 0 Then strText = strQZ & strText
        strQZ = Replace(strQZ, vbCrLf, "")
        intEnd = Len(strText)
        intLen = Len(strQZ)
        strCompare = ""
        intActLen = 0
        blnFind = False
        For intPos = 1 To intEnd
            lngASC = Asc(Mid(strText, intPos, 1))
            intActLen = intActLen + 1
            If Not (lngASC = 13 Or lngASC = 10) Then
                If lngASC = 32 Then
                    strCompare = ""
                    intActLen = 0
                Else
                    strCompare = strCompare & Mid(strText, intPos, 1)
                End If
                If Len(strCompare) = intLen Then
                    If strCompare = strQZ Then
                        blnFind = True
                        intPos = intPos + 1
                    Else
                        strCompare = ""
                        intPos = intPos + 1
                        intActLen = 0
                    End If
                End If
            End If
            If blnFind Then Exit For
        Next
        strText = Mid(strText, intPos)
    End If
    
    If IsNumeric(Replace(strText, vbCrLf, "")) Then
        SubstrAnaly = Replace(strText, vbCrLf, "")
    Else
        SubstrAnaly = strText
    End If
End Function

Public Sub SubstrPro(strFormat As String, ByVal strName As String, Optional ByVal intType As Integer = 0)
    Dim i As Integer, j As Integer, l As Integer, r As Integer, strHZ As String, strQZ As String
    'intType=0-ɾ��ָ����ʽ��;1-�õ�ָ����ʽ��
    j = Len(strFormat)
    i = InStr(1, strFormat, "[" & strName & "]")
    If i = 0 Then Exit Sub
    
    For l = i To 1 Step -1
        If Mid(strFormat, l, 1) = "{" Then Exit For
    Next
    strQZ = Mid(strFormat, l + 1, i - l - 1)
    For r = i To j
        If Mid(strFormat, r, 1) = "}" Then Exit For
    Next
    strHZ = Mid(strFormat, i + Len(strName) + 2, r - i - Len(strName) - 2)
    If intType = 0 Then
        'strFormat = Mid(strFormat, 1, l - 1) & strQZ & strHZ & Mid(strFormat, r + 1)
        If Mid(strFormat, 1, l - 1) = "" And Mid(strFormat, r + 1) = "" Then
            strFormat = ""
        Else
            strFormat = Mid(strFormat, 1, l - 1) & strQZ & strHZ & Mid(strFormat, r + 1)
        End If
    Else
        strFormat = Mid(strFormat, l, r - l + 1)
    End If
End Sub

Private Sub MoveNextCell(Optional ByVal blnNext As Boolean = True)
    Dim arrData
    Dim blnNULL As Boolean                      '�Ƿ�Ϊ����
    Dim strDate As String, strTime As String
    Dim strReturn As String, strMsg As String, strPart As String
    Dim lngStart As Long, lngMutilRows As Long, lngDeff As Long
    Dim intRow As Integer, intCount As Integer, intNULL As Integer  '����ж��ٿ���
    Dim blnTrue As Boolean
    '��ֵȻ���ƶ�����һ����Ч��Ԫ��
    Dim strKey As String, strField As String, strValue As String
    
    On Error GoTo errHand
    
    '�������,���ϸ���ٴε���Ҫ��¼��
    If mintType >= 0 Then
        If Not CheckInput(strReturn, strMsg) Then
            RaiseEvent AfterRowColChange(strMsg, True, mblnSign, mblnArchive)
            Exit Sub
        End If
        
        blnTrue = False
        If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "" Then VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "1|1"
        
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") <> 0 Then
            lngMutilRows = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
        End If
        lngStart = GetStartRow(VsfData.ROW)

        '׼����ֵ
        With txtLength
            '������ʱ���еĿ�Ȳ���,Ϊ�˱��ⷵ�ض���,ǿ������Ϊ5000
            .Width = IIf(VsfData.COL = mlngDate Or VsfData.COL = mlngTime, 5000, VsfData.CellWidth)
            .Text = strReturn
            .FontName = VsfData.CellFontName
            .FontSize = VsfData.CellFontSize
            .FontBold = VsfData.CellFontBold
            .FontItalic = VsfData.CellFontItalic
        End With
        arrData = GetData(txtLength.Text)
        intCount = UBound(arrData)
        
        If intCount > lngMutilRows - 1 Then
            '������������,�������������������������ӵ�����
            intNULL = intCount - (lngMutilRows - 1)
            For intRow = lngMutilRows To intCount
                '��֤��ǰ�����������һҳ����ʾȫ
                If intRow + lngStart > VsfData.Rows - 1 Then Exit For
                
                If Val(VsfData.TextMatrix(intRow + lngStart, mlngRecord)) = 0 And VsfData.TextMatrix(intRow + lngStart, mlngRowCount) = "" Then
                    intNULL = intNULL - 1
                Else
                    Exit For
                End If
            Next
            '�����ӿ���
            If intNULL > 0 Then
                lngDeff = intNULL
                VsfData.Rows = VsfData.Rows + intNULL
                '�ӵ�ǰ�м�¼�Ŀհ��п�ʼ��ÿ�е�λ��+�����ӵĿհ�����
                For intRow = VsfData.Rows - intNULL - 1 To lngStart + intCount - intNULL + 1 Step -1
                    VsfData.RowPosition(intRow) = intRow + intNULL
                Next
            End If
            'ѭ����ֵ
            intCount = UBound(arrData)
            For intRow = 0 To intCount
                VsfData.TextMatrix(lngStart + intRow, VsfData.COL) = Replace(Replace(arrData(intRow), Chr(10), ""), Chr(13), "")
                VsfData.TextMatrix(lngStart + intRow, mlngRowCount) = intCount + 1 & "|" & intRow + 1
            Next
            '���������н��и�ֵ
            lngMutilRows = lngStart + intCount
            For intRow = lngStart + 1 To lngMutilRows
                For intCount = 0 To VsfData.Cols - 1
                    VsfData.Cell(flexcpForeColor, intRow, intCount) = VsfData.Cell(flexcpForeColor, lngStart, intCount)
                    If VsfData.ColHidden(intCount) And mlngRowCount <> intCount Then
                        VsfData.TextMatrix(intRow, intCount) = VsfData.TextMatrix(lngStart, intCount)
                    End If
                Next
            Next
        Else
            '�Ը������¸�ֵ����ֻ����һ������ʱ����֪Ϊ�λ�����ַ�ASCII��Ϊ1�ķ��ţ�
            For intRow = 0 To intCount
                VsfData.TextMatrix(lngStart + intRow, VsfData.COL) = Replace(Replace(Replace(arrData(intRow), Chr(10), ""), Chr(13), ""), Chr(1), "")
            Next
            For intRow = intCount + 1 To lngMutilRows - 1
                VsfData.TextMatrix(lngStart + intRow, VsfData.COL) = ""
            Next
            
            '����������������д������,intNULL��¼���һ����Ϊ���е��к�
            intNULL = lngStart + lngMutilRows - 1
            For intRow = lngMutilRows To 1 Step -1
                blnNULL = True
                For intCount = 0 To VsfData.Cols - 1
                    If Not VsfData.ColHidden(intCount) Then
                        If VsfData.TextMatrix(intRow + lngStart - 1, intCount) <> "" Then
                            blnNULL = False
                            Exit For
                        End If
                    End If
                Next
                
                If Not blnNULL Then Exit For
                intNULL = intNULL - 1
            Next
            
            '������д�����
            For intRow = lngStart To intNULL
                VsfData.TextMatrix(intRow, mlngRowCount) = (intNULL - lngStart + 1) & "|" & intRow - lngStart + 1
            Next
            For intRow = intNULL + 1 To lngStart + lngMutilRows - 1
                VsfData.Cell(flexcpText, intRow, 0, intRow, VsfData.Cols - 1) = ""
            Next
        End If
        
        '���кŷ����仯����ͬ������mrsCellMap�д��ڸ��кŵ��к�����
        If lngDeff <> 0 Then
            Call CellMap_Update(lngStart, lngDeff)
        End If
        
        If mstrData <> strReturn Then
            mblnChange = True
            
            'ͬ������������ʱ���е�����
            strDate = VsfData.TextMatrix(lngStart, mlngDate)
            strTime = VsfData.TextMatrix(lngStart, mlngTime)
            
            '1\����
            strField = "ID|ҳ��|�к�|�к�|��¼ID|����|ɾ��"
            If mlngDate <> -1 Then
                strKey = mintҳ�� & "," & lngStart & "," & mlngDate
                strValue = strKey & "|" & mintҳ�� & "|" & lngStart & "|" & mlngDate & "|" & _
                    Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & strDate & "|0"
                Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
            End If
            '2\ʱ��
            strKey = mintҳ�� & "," & lngStart & "," & mlngTime
            strValue = strKey & "|" & mintҳ�� & "|" & lngStart & "|" & mlngTime & "|" & _
                Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & strTime & "|0"
            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
            
            '��¼�û��޸Ĺ��ĵ�Ԫ��
            If InStr(1, "," & mstrCatercorner & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 Then
                strPart = ""
            Else
                strPart = "/"
            End If
            
            strField = "ID|ҳ��|�к�|�к�|��¼ID|����|��λ|ɾ��"
            strKey = mintҳ�� & "," & lngStart & "," & VsfData.COL
            strValue = strKey & "|" & mintҳ�� & "|" & lngStart & "|" & VsfData.COL & "|" & _
                Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & strReturn & "|" & strPart & "|" & IIf(strReturn = "", "1", "0")
            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        End If
        RaiseEvent AfterDataChanged(mblnChange Or mblnVerify)
    End If
    
    'ʼ��Ҫ����һ�п��У��Ա�¼����������
    For intCount = VsfData.FixedCols To VsfData.Cols - 1
        If VsfData.ColHidden(intCount) = False Then
            If Trim(VsfData.TextMatrix(VsfData.Rows - 1, intCount)) <> "" Then
                VsfData.Rows = VsfData.Rows + 1
                Exit For
            End If
        End If
    Next
    
    If blnNext Then
toMoveNextCol:
        If VsfData.COL < mlngNoEditor - 1 Then
            VsfData.COL = VsfData.COL + 1
            If VsfData.ColWidth(VsfData.COL) = 0 Or VsfData.ColHidden(VsfData.COL) Then GoTo toMoveNextCol
        Else
toMoveNextRow:
            '������һ��
            If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") <> 0 Then
                intRow = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
                intRow = intRow - Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(1)) + 1
            Else
                intRow = 1
            End If
            If VsfData.ROW + intRow < VsfData.Rows Then
                VsfData.ROW = VsfData.ROW + intRow
            End If
            If VsfData.RowHidden(VsfData.ROW) Then GoTo toMoveNextRow
            VsfData.COL = IIf(mlngDate > 0 And mblnDate = True, mlngDate, mlngTime)
        End If
    Else
toMovePrevCol:
        If VsfData.COL > mlngDate Then
            VsfData.COL = VsfData.COL - 1
            If VsfData.ColWidth(VsfData.COL) = 0 Or VsfData.ColHidden(VsfData.COL) Then GoTo toMovePrevCol
        Else
toMovePrevRow:
'            '������һ��
'            intRow = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
'            intRow = intRow - Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(1)) + 1
'            If VsfData.ROW + intRow < VsfData.Rows Then
'                VsfData.ROW = VsfData.ROW + intRow
'            End If
'            If VsfData.RowHidden(VsfData.ROW) Then GoTo toMovePrevRow
'            VsfData.COL = IIf(mlngDate > 0, mlngDate, mlngTime)
        End If
    End If
    If VsfData.ColIsVisible(VsfData.COL) = False Then
        VsfData.LeftCol = VsfData.COL
    End If
    If VsfData.RowIsVisible(VsfData.ROW) = False Then
        VsfData.TopRow = VsfData.ROW
    End If
    Exit Sub
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function GetStartRow(ByVal lngRow As Long) As Long
    Dim lngStart As Long
    Dim lngCurRows As Long, lngRows As Long
    '��ȡ������ʼ��,������ҳ�򷵻�0
    '�����ҳδ��ʾȫ,��˵��������ҳ,Ҳ����0
    '���������������������в�������
    
    lngRows = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0))    '������
    lngCurRows = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(1)) '��ǰ��
    If lngCurRows = 1 Then
        GetStartRow = lngRow
        Exit Function
    End If
    
    'Ѱ����ʼ��
    For lngRow = lngRow To 3 Step -1
        If VsfData.TextMatrix(lngRow, mlngRowCount) = lngRows & "|1" Then
            lngStart = lngRow
            Exit For
        End If
    Next
    
    GetStartRow = lngStart
End Function

Private Function GetMutilData(ByVal lngRow As Long, ByVal lngCol As Long, dblTop As Long, dblHeight As Long) As String
    Dim lngCurRow As Long
    Dim lngCount As Long
    Dim lngStart As Long    '��ʼ��
    Dim lngRecordId As Long
    Dim strReturn As String
    Dim blnAdjust As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '���ص�һ�е�����
    '������ֱ��ȡ������ʱ��������ҳ��ʾȫ��ƴ�ӣ�����ӿ��ж�ȡ
    
    If VsfData.TextMatrix(lngRow, mlngRowCount) = "" Then
        GetMutilData = VsfData.TextMatrix(lngRow, lngCol)
        Exit Function
    End If
    lngRecordId = Val(VsfData.TextMatrix(lngRow, mlngRecord))
    lngCount = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0))
    lngCurRow = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(1))
    
    If lngCount > 1 Then
        lngStart = GetStartRow(lngRow)
    Else
        lngStart = lngRow
    End If
    If lngRecordId <> 0 And (lngStart + lngCount > VsfData.Rows) Then
        '�����ݿ�����ȡ
        Call SQLCombination(lngRecordId)
        gstrSQL = mstrSQL
        Call SQLDIY(gstrSQL)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ��������", mlng�ļ�ID, mlng����ID, mlng��ҳID, mintӤ��, cboBaby.ItemData(cboBaby.ListIndex), lngRecordId)
        strReturn = NVL(rsTemp.Fields(lngCol - VsfData.FixedCols).Value)
        blnAdjust = True
    Else
        For lngRow = lngStart To lngStart + lngCount - 1
            strReturn = strReturn & VsfData.TextMatrix(lngRow, lngCol)
        Next
    End If
    
    'ȡ�и�
    VsfData.ROW = lngStart
    dblHeight = lngCount * VsfData.RowHeightMin + 20
    dblTop = VsfData.Top + VsfData.CellTop
    
    GetMutilData = strReturn
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function ShowInput(Optional ByVal intCOl As Integer = -1, Optional ByVal strCellData As String = "", Optional ByVal blnAnalyse As Boolean = False) As String
    Dim arrData, arrValue
    Dim lngOrder As Long
    Dim i As Integer, j As Integer, intPos As Integer, intIndex As Integer
    Dim strFormat As String, strText As String, strValue As String  '��ʽ��,���ݴ�,��ֵ��
    Dim strOrders As String, strTypes As String, strBounds As String, strLen As String, strName As String
    Dim strCurDate As String
    Const txtHeight = 300
    On Error GoTo errHand
    
    '�����ļ��������ģ����Ҫ����:
    '1��һ�а�һ����Ŀ�Ĳ��ù�
    '2��һ�а�������Ŀ�ģ�Ѫѹ����ɶԣ�Ҫô����¼�룬Ҫô����ѡ�񣬲���������֣�Ҳ��������ֵ�ѡ����ѡ
    '3��һ�а󶨶����Ŀ�ģ�ֻ����¼����Ŀ
    '���������������ƣ�ֻȡ��һ����Ŀ�����ʼ���
    
    '����Ǳ��洦�����������´���
    If intCOl = -1 Then intCOl = VsfData.COL
    If blnAnalyse Then
        strText = strCellData
    Else
        'ȡ��ǰ��Ԫ�������
        CellRect.Left = VsfData.CellLeft + VsfData.Left
        CellRect.Top = VsfData.CellTop + VsfData.Top
        CellRect.Bottom = VsfData.CellHeight + 20
        CellRect.Right = VsfData.CellWidth + 20
        strText = GetMutilData(VsfData.ROW, intCOl, CellRect.Top, CellRect.Bottom)
    End If
    mstrData = strText
    mintType = 0
    intIndex = 0
    
    'ȡ��ǰ�еİ���Ŀ
    intPos = 1
    mrsSelItems.Filter = "��=" & intCOl - cHideCols
    Do While Not mrsSelItems.EOF
        lngOrder = mrsSelItems!��Ŀ���
        If lngOrder = 0 Then
            strLen = 0
            strValue = strText
            Exit Do
        End If
        
        '��Ŀ��ʾ:2��ѡ;3-��ѡ;4-����;5-ѡ��
        '��Ŀֵ��:��Ŀ��ʾΪ0-��ʾ��Сֵ;���ֵ;��Ŀ��ʾΪ2,3-��ʾ��ĿA;��ĿB,ǰ�й��ı�ʾȱʡ��
        strFormat = NVL(mrsSelItems!��ʽ)
        strOrders = strOrders & "," & lngOrder
        If lngOrder <> 0 Then
            mrsItems.Filter = "��Ŀ���=" & lngOrder
            strName = strName & "," & mrsItems!��Ŀ����
            strLen = strLen & "," & mrsItems!��Ŀ���� & ";" & NVL(mrsItems!��ĿС��)
            strTypes = strTypes & "," & mrsItems!��Ŀ��ʾ
            strBounds = strBounds & "," & mrsItems!��Ŀֵ��
            strValue = strValue & "'" & SubstrVal(strText, strFormat, mrsItems!��Ŀ����, intPos)
            
            Select Case mrsItems!��Ŀ��ʾ
            Case 0  '�ı�¼����
                If mrsSelItems.RecordCount = 2 Then
                    mintType = 4
                ElseIf mrsSelItems.RecordCount > 2 Then
                    mintType = 6
                End If
            Case 2  '��ѡ
                If mrsSelItems.RecordCount = 1 Then
                    mintType = 1
                ElseIf mrsSelItems.RecordCount = 2 Then
                    mintType = 7
                End If
            Case 3  '��ѡ
                mintType = 2
            Case 4  '����
                If mrsSelItems.RecordCount = 2 Then
                    mintType = 4
                ElseIf mrsSelItems.RecordCount > 2 Then
                    mintType = 6
                End If
            Case 5  'ѡ��
                If mrsSelItems.RecordCount = 1 Then
                    mintType = 3
                Else
                    mintType = 5
                End If
            End Select
        Else
            strTypes = strTypes & ","
            strBounds = strBounds & ","
            strLen = strLen & ","
            strName = strName & ","
        End If
        
        intIndex = intIndex + 1
        mrsSelItems.MoveNext
    Loop
    If strOrders <> "" Then
        strOrders = Mid(strOrders, 2)
        strName = Mid(strName, 2)
        strLen = Mid(strLen, 2)
        strTypes = Mid(strTypes, 2)
        strBounds = Mid(strBounds, 2)
        strValue = Mid(strValue, 2)
    End If
    mrsSelItems.Filter = 0
    mrsItems.Filter = 0
    
    If blnAnalyse Then
        ShowInput = strOrders & "||" & strValue
        Exit Function
    End If
    
    '���4����У��,�����ͷ�ı�����/����Ϊ6
    If mintType = 4 Then
        If Not IsDiagonal(intCOl) Then
            mintType = 6
        End If
    End If
    
    '�жϵ�ǰ�е�����
    'mintType:0=�ı���¼��;1=��ѡ;2=��ѡ;3=ѡ��;4-Ѫѹ��һ�а���������Ŀ,���ʽ����Ѫѹ��������Ŀ;5=һ�а���������Ŀ�Ҿ���ѡ����Ŀ;
    '6=һ�а�2����������Ŀ,�ֹ�¼��,7=һ�а���������ѡ��Ŀ
    arrValue = Split(strValue & "'", "'")
    Select Case mintType
    Case 0, 3
        With picInput
            .Left = CellRect.Left
            .Top = CellRect.Top
            .Width = CellRect.Right
            .Height = CellRect.Bottom
            If .Height + .Top + picMain.Top > ScaleHeight Then
                .Top = ScaleHeight - picMain.Top - .Height
            End If
            .Visible = True
        End With
        If mintType = 0 Then
            txtInput.Visible = True
            If Val(strLen) <> 0 And Val(strOrders) <> 10 Then
                txtInput.MaxLength = Val(Split(strLen, ";")(0)) + IIf(Val(Split(strLen, ";")(1)) = 0, 0, 1) 'С��λ��Ҫ����С����
            Else
                txtInput.MaxLength = 0
            End If
            txtInput.Tag = lngOrder
        Else
            txtInput.Visible = False
        End If
        With txtInput
            .Top = 0
            .Text = strValue
            .Width = CellRect.Right
            .Height = CellRect.Bottom
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Width = .Width - (180 + IIf(mblnBlowup, 180 * 1 / 3, 0)) / 2 '����9��ʱ��ȥ90,����Խ��۳��ı߾�ԽС,�Ա�֤�ı��������ʵ��һ��
        End With
        With lblInput
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Height = CellRect.Bottom
            .Width = CellRect.Right
            .Top = 50
            .Tag = lngOrder
            .Caption = strValue
            .Visible = (mintType = 3)
        End With
        
        '��������ڻ�ʱ���У��趨�̶�ֵ
        strCurDate = zlDatabase.Currentdate
        If mintType = 0 And txtInput.Text = "" Then
            If intCOl = mlngDate Or intCOl = mlngTime Then
                txtInput.Text = Format(strCurDate, "YYYY-MM-DD HH:mm")
                If Format(strCurDate, "YYYY-MM-DD HH:mm") >= Format(mstr����ʱ��, "YYYY-MM-DD HH:mm") Then
                    txtInput.Text = Format(mstr����ʱ��, "YYYY-MM-DD HH:mm")
                End If
                If Format(strCurDate, "YYYY-MM-DD HH:mm") <= Format(mstr��ʼʱ��, "YYYY-MM-DD HH:mm") Then
                    txtInput.Text = Format(mstr��ʼʱ��, "YYYY-MM-DD HH:mm")
                End If
            End If
            If intCOl = mlngDate Then
                If mblnDateAd Then
                    txtInput.Text = Format(txtInput.Text, "d-M")
                    txtInput.Text = Replace(txtInput.Text, "-", "/")
                Else
                    txtInput.Text = Format(txtInput.Text, "yyyy-MM-dd")
                End If
            ElseIf intCOl = mlngTime Then
                txtInput.Text = Format(txtInput.Text, "HH:mm")
            End If
        End If
    Case 1, 2
        '��������
        lstSelect(mintType - 1).Clear
        If mintType = 1 Then
            lstSelect(mintType - 1).AddItem "0-"
            If mlngProduce = intCOl Then lstSelect(mintType - 1).ListIndex = 0
        End If
        arrData = Split(strBounds, ";")
        j = UBound(arrData)
        For i = 0 To j
            If arrData(i) <> "" Then
                If Mid(arrData(i), 1, 1) = "��" Then
                    lstSelect(mintType - 1).AddItem lstSelect(mintType - 1).NewIndex + 1 & "-" & Mid(arrData(i), 2)
                    If strText = "" And lstSelect(mintType - 1).ListIndex = -1 Then lstSelect(mintType - 1).ListIndex = lstSelect(mintType - 1).NewIndex
                Else
                    lstSelect(mintType - 1).AddItem lstSelect(mintType - 1).NewIndex + 1 & "-" & arrData(i)
                End If
            End If
        Next
        '��ѡ����¼�����ݵ������
        If strValue <> "" Then
            strValue = Replace(strValue, vbCrLf, "")
            j = lstSelect(mintType - 1).ListCount - 1
            For i = 0 To j
                If InStr(1, "," & strValue & ",", "," & Split(lstSelect(mintType - 1).List(i), "-")(1) & ",") <> 0 Then
                    lstSelect(mintType - 1).Selected(i) = True
                End If
            Next
        End If
        '��ʾ
        With lstSelect(mintType - 1)
            .Left = CellRect.Left
            .Top = CellRect.Top
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Height = .ListCount * 300
            If .Height < CellRect.Bottom Then .Height = CellRect.Bottom
            .Width = LenB(StrConv(.List(.ListCount \ 2), vbFromUnicode)) * 120 + 500    '���м���ĳ���Ϊ����
            If .Width < CellRect.Right Then .Width = CellRect.Right
            If .Height + .Top + picMain.Top > ScaleHeight Then
                .Top = ScaleHeight - picMain.Top - .Height
            End If
            .Tag = lngOrder
            .Visible = True
        End With
    Case 4, 5
        With picDouble
            .Left = CellRect.Left
            .Top = CellRect.Top
            .Height = CellRect.Bottom
            If .Height < 280 Then .Height = 280
            .Width = CellRect.Right
            If .Width < 820 Then .Width = 820
            If .Height + .Top + picMain.Top > ScaleHeight Then
                .Top = ScaleHeight - picMain.Top - .Height
            End If
            .Visible = True
        End With
        lblSplit.FontName = VsfData.FontName
        lblSplit.FontSize = VsfData.FontSize
        lblSplit.Left = (picDouble.Width - lblSplit.Width) / 2
        If mblnBlowup Then
            lblSplit.Width = 150
        Else
            lblSplit.Width = 105
        End If
        
        With txtUpInput
            .Text = arrValue(0)
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Width = (picDouble.Width - lblSplit.Width) * 0.4
            .ZOrder IIf(mintType = 4, 0, 1)
            .Locked = Not (mintType = 4)
            .Tag = Split(strOrders, ",")(0)
        End With
        With picUpInput
            .Left = txtUpInput.Left
            .Width = txtUpInput.Width
            .Height = CellRect.Bottom
            .ZOrder IIf(mintType = 5, 0, 1)
            .Tag = Split(strOrders, ",")(0)
        End With
        With lblUpInput
            .Alignment = 2
            .Caption = arrValue(0)
            .Left = 0
            .Top = 50
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Width = txtUpInput.Width
            .Height = CellRect.Bottom
            .Tag = Split(strOrders, ",")(0)
        End With
        With txtDnInput
            .Text = arrValue(1)
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Left = lblSplit.Left + lblSplit.Width
            .Width = picDouble.Width - .Left
            .ZOrder IIf(mintType = 4, 0, 1)
            .Locked = Not (mintType = 4)
            .Tag = Split(strOrders, ",")(1)
        End With
        With picDnInput
            .Left = txtDnInput.Left
            .Height = CellRect.Bottom
            .Width = txtDnInput.Width
            .ZOrder IIf(mintType = 5, 0, 1)
            .Tag = Split(strOrders, ",")(1)
        End With
        With lblDnInput
            .Alignment = 2
            .Caption = arrValue(1)
            .Left = 0
            .Top = 50
            .Height = CellRect.Bottom
            .Width = txtDnInput.Width
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Tag = Split(strOrders, ",")(1)
        End With
        
        If mintType = 4 Then
            If strLen <> "" Then txtUpInput.MaxLength = Val(Split(Split(strLen, ",")(0), ";")(0)) + IIf(Val(Split(Split(strLen, ",")(0), ";")(1)) = 0, 0, 1) 'С��λ��Ҫ����С����
            If strLen <> "" Then txtDnInput.MaxLength = Val(Split(Split(strLen, ",")(1), ";")(0)) + IIf(Val(Split(Split(strLen, ",")(1), ";")(1)) = 0, 0, 1) 'С��λ��Ҫ����С����
        End If
    Case 6
        '��ɾ����ǰ�Ŀؼ�
        j = txt.Count - 1
        For i = 1 To j
            Unload lbl(i)
            Unload txt(i)
        Next
        '�趨����
        With picMutilInput
            .Left = CellRect.Left
            .Top = CellRect.Top
            .Width = IIf(CellRect.Right < 1600, 1600, CellRect.Right)
        End With
        '��ȱʡ�ؼ���ֵ
        arrData = Split(strOrders, ",")
        j = UBound(arrData)
        lbl(0).Top = 130
        lbl(0).Caption = Split(strName, ",")(0)
        lbl(0).FontName = VsfData.FontName
        lbl(0).FontSize = VsfData.FontSize
        txt(0).Tag = arrData(0)
        txt(0).FontName = VsfData.FontName
        txt(0).FontSize = VsfData.FontSize
        txt(0).Width = picMutilInput.Width - txt(0).Left - 100
        txt(0).MaxLength = Val(Split(Split(strLen, ",")(0), ";")(0)) + IIf(Val(Split(Split(strLen, ",")(0), ";")(1)) = 0, 0, 1)  'С��λ��Ҫ����С����
        txt(0).Text = arrValue(0)
        If Not mblnBlowup Then
            txt(0).Height = 225
        End If
        
        '���ؿؼ�
        For i = 1 To j
            Load lbl(i)
            With lbl(i)
                .Caption = Split(strName, ",")(i)
                .Left = lbl(0).Left + lbl(0).Width - .Width
                .Top = lbl(i - 1).Top + txtHeight + IIf(mblnBlowup, txtHeight * 1 / 3, 0)
                .Visible = True
            End With
            Load txt(i)
            With txt(i)
                .TabIndex = txt(i - 1).TabIndex + 1
                .Left = txt(0).Left
                .Top = txt(i - 1).Top + txtHeight + IIf(mblnBlowup, txtHeight * 1 / 3, 0)
                .Tag = arrData(i)
                If strLen <> "" Then
                    .MaxLength = Val(Split(Split(strLen, ",")(i), ";")(0)) + IIf(Val(Split(Split(strLen, ",")(i), ";")(1)) = 0, 0, 1) 'С��λ��Ҫ����С����
                End If
                .Text = arrValue(i)
                .Visible = True
            End With
        Next
        
        With picMutilInput
            .Height = txt(j).Top + txt(j).Height + 120
            If .Height < CellRect.Bottom Then .Height = CellRect.Bottom
            If .Height + .Top + picMain.Top > ScaleHeight Then
                .Top = ScaleHeight - picMain.Top - .Height
            End If
            .Visible = True
        End With
    Case 7
        cboChoose(0).Clear
        arrData = Split(Split(strBounds, ",")(0), ";")
        j = UBound(arrData)
        For i = 0 To j
            If Mid(arrData(i), 1, 1) = "��" Then
                cboChoose(0).AddItem Mid(arrData(i), 2)
                If strValue = "" Then
                    cboChoose(0).ListIndex = i
                Else
                    If Mid(arrData(i), 2) = Split(strValue, "'")(0) Then
                        cboChoose(0).ListIndex = i
                    End If
                End If
            Else
                cboChoose(0).AddItem arrData(i)
                If strValue <> "" Then
                    If arrData(i) = Split(strValue, "'")(0) Then
                        cboChoose(0).ListIndex = i
                    End If
                End If
            End If
        Next
        cboChoose(1).Clear
        arrData = Split(Split(strBounds, ",")(1), ";")
        j = UBound(arrData)
        For i = 0 To j
            If Mid(arrData(i), 1, 1) = "��" Then
                cboChoose(1).AddItem Mid(arrData(i), 2)
                If strValue = "" Then
                    cboChoose(1).ListIndex = i
                Else
                    If Mid(arrData(i), 2) = Split(strValue, "'")(1) Then
                        cboChoose(1).ListIndex = i
                    End If
                End If
            Else
                cboChoose(1).AddItem arrData(i)
                If strValue <> "" Then
                    If arrData(i) = Split(strValue, "'")(1) Then
                        cboChoose(1).ListIndex = i
                    End If
                End If
            End If
        Next
        With picDoubleChoose
            .Left = CellRect.Left
            .Top = CellRect.Top
            .Height = CellRect.Bottom
            If .Height < 280 Then .Height = 280
            .Width = CellRect.Right
            If .Width < 820 Then .Width = 820
            If .Height + .Top + picMain.Top > ScaleHeight Then
                .Top = ScaleHeight - picMain.Top - .Height
            End If
            .Visible = True
        End With
        lblSplit.FontName = VsfData.FontName
        lblSplit.FontSize = VsfData.FontSize
        lblSplit.Left = (picDoubleChoose.Width - lblSplit.Width) / 2
        If mblnBlowup Then
            lblSplit.Width = 150
        Else
            lblSplit.Width = 105
        End If
        
        cboChoose(0).SetFocus
    End Select
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub CheckFormat(ByVal strNames As String, ByVal strFormat As String)
    '�����ʽ��Ѫѹ�ķ�ʽ��ͬ,����ʽ����Ϊ6
    
    'ȥ��ǰ׺����жԱ�
    strFormat = Mid(strFormat, InStr(1, strFormat, "["))
    strFormat = Replace(strFormat, "[", "")
    strFormat = Replace(strFormat, "]", "")
    If Not (strFormat Like Split(strNames, ",")(0) & "/}*" Or strFormat Like "{/*" & Split(strNames, ",")(1)) Then
        mintType = 6
    End If
End Sub

Private Function IsDiagonal(ByVal intCOl As Integer) As Boolean
    Dim arrCol, arrData
    Dim intDo As Integer, intCount As Integer
    '�ж�ָ�����Ƿ��������жԽ��ߣ�mstrColWidth�ĸ�ʽ��765`11`1`1,765`11`2`1,...����������`�������`�жԽ��ߣ�
    
    IsDiagonal = (InStr(1, "," & mstrCatercorner & ",", "," & intCOl - (cHideCols + VsfData.FixedCols - 1) & ",") <> 0)
End Function

Private Sub ISAssistant(ByVal lngOrder As Long, ByVal objTXT As TextBox)
    Dim intIndex As Integer
    Dim objParent As Object
    
    '������Ŀ�ĳ��Ⱦ����Ƿ�������дʾ�ѡ��
    mblnEditAssistant = False
    cmdWord.Visible = mblnEditAssistant
    
    mrsItems.Filter = "��Ŀ���=" & lngOrder
    If mrsItems.RecordCount = 0 Then
        mrsItems.Filter = 0
        Exit Sub
    End If
    mblnEditAssistant = (mrsItems!��Ŀ���� = 1 And mrsItems!��Ŀ���� >= 100)
    mrsItems.Filter = 0
    
    '�������ʾ�ѡ��,��ʾ����λ
    If mblnEditAssistant Then
        If UCase(objTXT.Name) = "TXTINPUT" Then
            intIndex = -1 '��ʾtxtInput
            Set objParent = picInput
        Else
            intIndex = objTXT.Index
            Set objParent = picMutilInput
        End If
        With cmdWord
            .Tag = intIndex
            .Top = objParent.Top + objTXT.Top + 25
            .Left = objParent.Left + objTXT.Left + objTXT.Width - .Width + 25
            .Visible = True
        End With
    End If
End Sub

Public Function GetSynItems(ByVal intType As Integer, ByRef intMax As Integer) As String
    Dim arrCols
    Dim strItems As String
    Dim strCols As String
    Dim strNames As String
    Dim lngRecord As Long, lngStartRow As Long, lngEndRow As Long
    Dim intIn As Integer, intOut As Integer, intInMAX As Integer, intOutMax As Integer, intCount As Integer
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    'intType��ȷ����������ֵ��1)������Ŀ���;2)�����к�
    'intMAX������ͬ����������ռ�õ��и�
    '����ͬ��������(һ���ļ��в����ܳ����ظ�����Ŀ,����,�ж�ʱ���ؼ���к�)
    
    lngRecord = Val(VsfData.TextMatrix(VsfData.ROW, mlngRecord))
    If lngRecord = 0 Then Exit Function
    
    gstrSQL = "" & _
        " SELECT  B.��Ŀ���,B.��Ŀ����,A.������� AS �к�" & vbNewLine & _
        " FROM �����ļ��ṹ A,���˻�����ϸ B" & vbNewLine & _
        " WHERE A.Ҫ������=B.��Ŀ���� AND A.��ID=" & vbNewLine & _
        "      (SELECT A.ID FROM �����ļ��ṹ A,���˻����ļ� B " & vbNewLine & _
        "       WHERE B.ID=[2] And A.�ļ�ID=B.��ʽID AND A.�������=4 AND A.��ID IS NULL)" & vbNewLine & _
        " AND B.������Դ>0 AND B.��¼ID=[1]"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "����ͬ��������", lngRecord, mlng�ļ�ID)
    If rsTemp.RecordCount = 0 Then Exit Function
    
    '��ȡͬ�������Ϣ
    Do While Not rsTemp.EOF
        If InStr(1, "," & strCols & ",", "," & rsTemp!�к� & ",") = 0 Then strCols = strCols & "," & rsTemp!�к�
        strItems = strItems & "," & rsTemp!��Ŀ���
        strNames = strNames & "," & rsTemp!��Ŀ����
        rsTemp.MoveNext
    Loop
    strCols = Mid(strCols, 2)
    strItems = Mid(strItems, 2)
    strNames = Mid(strNames, 2)
    
    '������ѭ�����������ռ�и�
    If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then
        lngStartRow = VsfData.ROW
        lngEndRow = VsfData.ROW
        intInMAX = 1
    Else
        lngStartRow = GetStartRow(VsfData.ROW)
        intInMAX = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
        lngEndRow = lngStartRow + intInMAX - 1
    End If
    
    intCount = 1    'ͬ����ֻ������������Ŀ������ռ����ֻ������1�У��������ݲ�����Ҫ���
'    '����ռ�ó���1�вż��
'    If intInMAX > 1 Then
'        arrCols = Split(strCols, ",")
'        intOutMax = UBound(arrCols)
'        For intOut = 0 To intOutMax
'            For intIn = 2 To intInMAX
'                If VsfData.TextMatrix(intIn + lngStartRow - 1, arrCols(intOut) + 1) <> "" Then
'                    If intIn > intCount Then intCount = intIn
'                End If
'            Next
'        Next
'    End If
    
    intMax = intCount
    GetSynItems = IIf(intType = 1, strItems, strCols)
    If strNames <> "" Then
        RaiseEvent AfterRowColChange("������,ʱ����,�Լ� " & strNames & " ��ͬ�����������ݣ��������޸Ļ�ɾ����", True, mblnSign, mblnArchive)
    End If
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

'######################################################################################################################
'**********************************************************************************************************************
'�����ǻ������������

Private Sub lblDnInput_Click()
    txtDnInput.SetFocus
End Sub

Private Sub lblUpInput_Click()
    txtUpInput.SetFocus
End Sub

Private Sub lstSelect_DblClick(Index As Integer)
    Call lstSelect_KeyDown(Index, vbKeyReturn, 0)
End Sub

Private Sub lstSelect_GotFocus(Index As Integer)
    mblnEditAssistant = False
End Sub

Private Sub txtDnInput_GotFocus()
    txtDnInput.SelStart = 0
    txtDnInput.SelLength = 100
    Call ISAssistant(Val(txtDnInput.Tag), txtDnInput)
End Sub

Private Sub txtInput_GotFocus()
    txtInput.SelStart = 0
    txtInput.SelLength = 100
    mintSymbol = -1
    Call ISAssistant(Val(txtInput.Tag), txtInput)
End Sub

Private Sub txtUpInput_GotFocus()
    txtUpInput.SelStart = 0
    txtUpInput.SelLength = 100
    Call ISAssistant(Val(txtUpInput.Tag), txtUpInput)
End Sub

Private Sub txt_GotFocus(Index As Integer)
    txt(Index).SelStart = 0
    txt(Index).SelLength = 100
    mintSymbol = Index
    Call ISAssistant(Val(txt(Index).Tag), txt(Index))
End Sub

Private Sub lblUpInput_DblClick()
    lblUpInput.Caption = IIf(lblUpInput.Caption = "", "��", "")
    txtUpInput.SetFocus
End Sub

Private Sub lblDnInput_DblClick()
    lblDnInput.Caption = IIf(lblDnInput.Caption = "", "��", "")
    txtDnInput.SetFocus
End Sub

Private Sub lblInput_DblClick()
    lblInput.Caption = IIf(lblInput.Caption = "", "��", "")
End Sub

Private Sub txtUpInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtDnInput.SetFocus
    ElseIf KeyCode = vbKeyRight Then
        If txtUpInput.SelStart = Len(txtUpInput.Text) Then txtDnInput.SetFocus
    ElseIf KeyCode = vbKeyLeft And txtUpInput.SelStart = 1 Then
        Call MoveNextCell(False)
    ElseIf KeyCode = vbKeySpace And txtUpInput.Locked Then
        lblUpInput.Caption = IIf(lblUpInput.Caption = "", "��", "")
    End If
End Sub

Private Sub txtDnInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyRight And txtDnInput.SelStart = Len(txtDnInput.Text)) Then
        Call picDouble_KeyDown(KeyCode, Shift)
    ElseIf KeyCode = vbKeyLeft Then
        If txtDnInput.SelStart = 0 Then txtUpInput.SetFocus
    ElseIf KeyCode = vbKeySpace And txtDnInput.Locked Then
        lblDnInput.Caption = IIf(lblDnInput.Caption = "", "��", "")
    End If
End Sub

Private Sub picMutilInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call MoveNextCell
End Sub

Private Sub picDouble_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Then
        Call MoveNextCell(Not (KeyCode = vbKeyLeft))
    End If
End Sub

Private Sub picInput_GotFocus()
    If txtInput.Visible Then
        txtInput.SetFocus
    End If
End Sub

Private Sub picInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If Not txtInput.Visible Then
        If KeyCode = vbKeySpace Then
            Call lblInput_DblClick
        End If
    End If
    
    If KeyCode = vbKeyReturn Or KeyCode = vbKeyRight Or KeyCode = vbKeyLeft Then
        '�ƶ�����һ����Ԫ��
        Call MoveNextCell(Not (KeyCode = vbKeyLeft))
    End If
End Sub

Private Sub lstSelect_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call MoveNextCell
End Sub

Private Sub picMutilInput_GotFocus()
    On Error Resume Next
    txt(0).SetFocus
End Sub

Private Sub txt_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Index < txt.Count - 1 Then
            txt(Index + 1).SetFocus
        Else
            Call picMutilInput_KeyDown(KeyCode, Shift)
        End If
    End If
End Sub

Private Sub picDouble_GotFocus()
    If txtUpInput.Visible Then
        txtUpInput.SetFocus
    End If
End Sub

Private Sub picMain_Resize()
    picMain.Left = 0
    VsfData.Width = picMain.Width
    VsfData.Height = IIf(picMain.Height - VsfData.Top < 0, 0, picMain.Height - VsfData.Top)
End Sub

Private Sub txtInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn And Shift = vbCtrlMask Then Exit Sub
    
    If KeyCode = vbKeyReturn Or _
        (KeyCode = vbKeyRight And txtInput.SelStart = Len(txtInput.Text)) Or _
        (KeyCode = vbKeyLeft And txtInput.SelStart = 1) Then
        Call picInput_KeyDown(KeyCode, Shift)
    End If
End Sub

Private Sub txtUpInput_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("/") Then
        KeyAscii = 0
        txtDnInput.SetFocus
    End If
End Sub
 

Private Sub txtС������_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Or KeyAscii = Asc(";") Then KeyAscii = 0
End Sub

Private Sub UserControl_GotFocus()
    On Error Resume Next
    VsfData.SetFocus
End Sub

Private Sub UserControl_Initialize()
    mblnShow = False
    mblnChange = False
    mblnInit = False
    
'    Set objStream = objFileSys.OpenTextFile("C:\WORKLOG.txt", ForAppending, True)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    '�����ַ���Ϊ���ݷָ�������¼�¼���ķָ�������˲�����¼��
    If KeyAscii = 39 Or KeyAscii = 13 Or KeyAscii = Asc("|") Then KeyAscii = 0: Exit Sub
    If KeyAscii = vbKeyEscape And mblnShow Then
        mblnShow = False
        Call InitCons
    End If
End Sub

Private Sub cbsThis_Resize()
    Dim lngScaleLeft As Long, lngScaleTop As Long, lngScaleRight As Long, lngScaleBottom As Long
    Call cbsThis.GetClientRect(lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom)
    
    Err = 0: On Error Resume Next
    lblTitle.Move lngScaleLeft, lngScaleTop + 120, lngScaleRight - lngScaleLeft
    With lblSubhead
        .Left = lngScaleLeft + 210: .Width = lngScaleRight - lngScaleLeft - 210 * 2
        .Top = lblTitle.Top + lblTitle.Height + 120
    End With
    picMain.Move lngScaleLeft, lngScaleTop, lngScaleRight, lngScaleBottom - lngScaleTop
    vsfHead.Move lngScaleLeft + 210, lblSubhead.Top + lblSubhead.Height + 45, lngScaleRight - lngScaleLeft - 210 * 2
    VsfData.Move vsfHead.Left, vsfHead.Top + vsfHead.Height - 20, vsfHead.Width
    VsfData.Height = picMain.Height - vsfHead.Height - vsfHead.Top
    
    lblSubEnd.Move lblSubhead.Left, VsfData.Top + VsfData.Height + 45
    
    lblCurPage.Top = picMain.Top
    lblCurPage.Left = picMain.Width - lblCurPage.Width
    
    '���ϱ�ǩ��ɢ����
    Call zlLableBruit
End Sub

Private Sub UserControl_Terminate()
'    objStream.Close
End Sub

Private Sub SetDockRight(BarToDock As CommandBar, BarOnLeft As CommandBar)
    '******************************************************************************************************************
    '���ܣ�
    '������
    '���أ�
    '******************************************************************************************************************
    
    Dim Left As Long
    Dim Top As Long
    Dim Right As Long
    Dim Bottom As Long
    
    cbsThis.RecalcLayout
    BarOnLeft.GetWindowRect Left, Top, Right, Bottom
    cbsThis.DockToolBar BarToDock, Right, (Bottom + Top) / 2, BarOnLeft.Position
End Sub

Public Sub Record_Add(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String)
    Dim arrFields, arrValues, intField As Integer
    '��Ӽ�¼
    'strFields:�ֶ���|�ֶ���
    'strValues:ֵ|ֵ
    
    '���ӣ�
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|��ĿID|ժҪ"
    'strValues = "5188|6666|��Ŀ����"
    'Call Record_Update(rsVoucher, strFields, strValues)

    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField = 0 Then Exit Sub

    With rsObj
        .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Public Sub Record_Update(ByRef rsObj As ADODB.Recordset, ByVal strFields As String, ByVal strValues As String, ByVal strPrimary As String, Optional ByVal blnDelete As Boolean = False)
    Dim arrFields, arrValues, intField As Integer
    '���¼�¼,���������,������
    'strPrimary:�ֶ���|ֵ
    'strFields:�ֶ���|�ֶ���
    'strValues:ֵ|ֵ
    
    '���ӣ�
    'Dim strFields As String, strValues As String, strPrimary As String
    'strFields = "RecordID|��ĿID|ժҪ"
    'strValues = "5188|6666|��Ŀ����"
    'strPrimary = "RecordID|5188"
    'Call Record_Update(rsVoucher, strFields, strValues, strPrimary, True)

    If strValues = "" Then strValues = " "
    arrFields = Split(strFields, "|")
    arrValues = Split(strValues, "|")
    intField = UBound(arrFields)
    If intField < 0 Then Exit Sub

    With rsObj
        If Record_Locate(rsObj, strPrimary, blnDelete) = False Then .AddNew
        For intField = 0 To intField
            .Fields(arrFields(intField)).Value = IIf(UCase(arrValues(intField)) = "NULL", Null, arrValues(intField))
        Next
        .Update
    End With
End Sub

Public Function Record_Locate(ByRef rsObj As ADODB.Recordset, ByVal strPrimary As String, Optional ByVal blnDelete As Boolean = False) As Boolean
    Dim arrTmp
    '��λ��ָ����¼
    'strPrimary:����,ֵ
    'blnDelete=True,��ü�¼������"ɾ��"�ֶ�
    Record_Locate = False
    
    arrTmp = Split(strPrimary, "|")
    With rsObj
        If .RecordCount = 0 Then Exit Function
        .MoveFirst
        .Find arrTmp(0) & "='" & arrTmp(1) & "'"
        If .EOF Then Exit Function
        If blnDelete Then
            Do While Not .EOF
                If !ɾ�� = 0 Then Record_Locate = True: Exit Do
                .MoveNext
            Loop
        Else
            Record_Locate = True
        End If
    End With
End Function

Public Sub Record_Init(ByRef rsObj As ADODB.Recordset, ByVal strFields As String)
    Dim arrFields, intField As Integer
    Dim strFieldName As String, intType As Integer, lngLength As Long
    '��ʼ��ӳ���¼��
    'strFields:�ֶ���,����,����|�ֶ���,����,����    �������Ϊ��,��ȡĬ�ϳ���
    '�ַ���:adLongVarChar;������:adDouble;������:adDBDate
    
    '���ӣ�
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|��ĿID," & adDouble & ",18|ժҪ, " & adLongVarChar & ",50|" & _
    '"ɾ��," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intType = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '��ȡ�ֶ�ȱʡ����
            If lngLength = 0 Then
                Select Case intType
                Case adDouble
                    lngLength = madDoubleDefault
                Case adVarChar
                    lngLength = madLongVarCharDefault
                Case adLongVarChar
                    lngLength = madLongVarCharDefault
                Case Else
                    lngLength = madDbDateDefault
                End Select
            End If
            .Fields.Append strFieldName, intType, lngLength, adFldIsNullable
        Next
        
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open
    End With
End Sub

Private Sub OutputRsData(ByVal rsObj As ADODB.Recordset)
    Dim intCOl As Integer, intCols As Integer
    Dim strValues As String
    With rsObj
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            strValues = ""
            intCols = .Fields.Count - 1
            For intCOl = 0 To intCols
                strValues = strValues & "," & .Fields(intCOl).Name & ":" & .Fields(intCOl).Value
            Next
            Debug.Print Mid(strValues, 2)
            .MoveNext
        Loop
        If .RecordCount <> 0 Then .MoveFirst
    End With
End Sub

Private Function BlowUp(ByRef dblChange As Double) As Double
    '�Ŵ����壬��Ԫ����
    BlowUp = dblChange
    If Not mblnBlowup Then Exit Function
    BlowUp = dblChange + (dblChange * 1 / 3)
End Function



