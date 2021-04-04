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
         Name            =   "宋体"
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
            ToolTipText     =   "添加"
            Top             =   1320
            Width           =   450
         End
         Begin VB.CommandButton cmdBabyCancle 
            Height          =   315
            Left            =   1440
            Picture         =   "usrPartogramEditor.ctx":0CBE
            Style           =   1  'Graphical
            TabIndex        =   38
            ToolTipText     =   "取消"
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
               Name            =   "宋体"
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
               ToolTipText     =   "删除"
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
            Caption         =   "验证签名"
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
         Caption         =   "全选"
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
            Caption         =   "体温体录"
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
            Caption         =   "√"
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
            Caption         =   "全部"
            Height          =   350
            Left            =   270
            TabIndex        =   45
            ToolTipText     =   "确认"
            Top             =   2370
            Width           =   840
         End
         Begin VB.CommandButton cmdSignCur 
            Caption         =   "验证"
            Height          =   350
            Left            =   2790
            TabIndex        =   44
            ToolTipText     =   "确认"
            Top             =   2370
            Width           =   840
         End
         Begin VB.CommandButton cmdCancl 
            Caption         =   "取消"
            Height          =   350
            Left            =   3690
            TabIndex        =   43
            ToolTipText     =   "取消"
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
            Caption         =   "以下是签名历史记录，可选择单行验证，也可进行全部验证。"
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
         Caption         =   "备注:##"
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
            Name            =   "宋体"
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
         Caption         =   "一般护理记录单"
         Height          =   180
         Left            =   3450
         TabIndex        =   12
         Top             =   30
         Width           =   1275
      End
      Begin VB.Label lblSubhead 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "姓名:##"
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
'基础条件:
'1.护理记录同一时点只可能存在一条记录
'2.护理记录中不需要像体温单那样 , 记录病人是否外出, 拒测的数据, 测试了的数据才记录
'3.录入护理记录数据时,如果所录入的数据存在体温数据, 则提取过来
'4.护理记录中不需要录入物理降温及脉搏短拙，如确需要可录入在护理摘要等文字型的列中
'#实现原理:
'1.增加记录集记录哪些页哪些单元格被用户修改过
'2.任何编辑(粘贴,清空数据),都需要重新计算每行数据的占用行

Public mblnEditable As Boolean
'Public objFileSys As New FileSystemObject
'Public objStream As TextStream

Private Const mint页码 As Integer = 1
Private mFrmParent As Object
Private mblnInit As Boolean
Private mblnShow As Boolean                 '是否显示录入框
Private mblnVerify As Boolean               '是否审签模式(可修改,但不允许进行复制粘贴清除等操作,只能修改)
Private mstrVerify As String                '等待审签的ID串
Private mintVerify As Integer               '当前操作员的最高级别
Private mintVerify_Last As Integer          '所选审签记录中最高级别
Private mblnBlowup As Boolean               '放大否？放大1/3，如字体9号放大为12号
Private mblnChange As Boolean               '是否修改数据
Private mstrData As String                  '进入编辑状态前保存之前的数据
Private mintPreDays As Long
Private mstrMaxDate As String

Private mlng文件ID As Long
Private mlng格式ID As Long
Private mlng病人ID As Long
Private mlng主页ID As Long
Private mlng科室id As Long
Private mlng病区ID As Long
Private mint婴儿 As Integer
Private mlngFileIndex As Long
Private mstrPrivs As String

Private mintSymbol As Integer               '当前控件索引
Private mstrSymbol As String                '特殊字符
Private mstrCOLNothing As String            '未绑定的列集合
Private mstrCatercorner As String           '列对角线集合
Private mblnEditAssistant As Boolean        '当前选择的项目是否允许进行词句选择
Private mlngRowCount As Long                '当前记录总行数
Private mlngDate As Long                    '日期
Private mlngTime As Long                    '时间
Private mlngSpread As Long                  '宫口扩大
Private mlngJust As Long                    '先露到底
Private mlngProduce As Long                 '生产
Private mlngChoose As Long                  '选择列
Private mlngOperator As Long                '护士
Private mlngSignLevel As Long               '签名级别
Private mlngSigner As Long                  '签名信息
Private mlngSignName As Long                '签名人
Private mlngSignTime As Long                '签名时间
Private mlngRecord As Long                  '记录ID
Private mlngNoEditor As Long                '禁止编辑列,存在护士列则以护士列为准,不存在护士列则以签名列为准
Private mlngDemo As Long                    '备用列
Private mlngActTime As Long                 '发生时间

Private mblnSign As Boolean                 '是否签名
Private mblnArchive As Boolean              '是否归档
Private mintType As Integer                 '记录当前的编辑模式
Private mblnDateAd As Boolean               '日期缩写?
Private mblnDate As Boolean                 '是否存在日期列
Private mstr开始时间 As String              '当前文件的开始时间
Private mstr结束时间 As String              '当前文件的结束时间
Private mstrBeginTime As String             '宫缩开始时间
Private CellRect As RECT

Private rsTemp As New ADODB.Recordset
Private mrsItems As New ADODB.Recordset             '所有护理记录项目清单
Private mrsPartogram As New ADODB.Recordset         '产程要素记录清单
Private mrsSelItems As New ADODB.Recordset          '当前录入的护理记录项目清单
Private mrsDataMap As New ADODB.Recordset           '当前操作员录入的数据镜像,与数据编辑格式一致,相关行数据全部保存以便迅速恢复
Private mrsCellMap As New ADODB.Recordset           '编辑过的数据镜像,字段有:页号,行号,列号,记录ID,数据,部位,删除
Private mrsCopyMap As New ADODB.Recordset           '复制行数据

Private Enum ColIcon
    签名 = 1
    审签 = 2
End Enum
Private Enum SignLevel
    正高 = 1
    副高 = 2
    中级 = 3
    师级 = 4
    员士 = 5
    未定义 = 9
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

Private Const madLongVarCharDefault As Integer = 10          '字符型字段缺省长度
Private Const madDoubleDefault As Integer = 18               '数字型字段缺省长度
Private Const madDbDateDefault As Integer = 20               '日期型字段缺省长度

Public Event AfterDataChanged(ByVal blnChange As Boolean)
Public Event AfterRefresh()
Public Event AfterRowColChange(ByVal strInfo As String, ByVal blnImportant As Boolean, ByVal blnSign As Boolean, ByVal blnArchive As Boolean)
Public Event AfterDataSave(ByVal blnSave As Boolean)
Public Event AfterFileIndex(ByVal lngFileIndex As Long)
Public Event AfterPartogramInfo(ByVal lngFlieId As Long, ByVal lngFileIndex As Long, ByVal lngFileFormatID As Long, ByVal rsPartogram As ADODB.Recordset)

Private mstrFields As String
Private mstrValues As String
Private mstrTag As String           '暂存

'病历文件格式定义相关
Private mintTabTiers As Integer     '表头层次
Private mintTagFormHour As Integer  '开始时间条件
Private mintTagToHour As Integer    '截止时间条件
Private mobjTagFont As New StdFont  '条件样式字体
Private mlngTagColor As Long        '条件样式颜色
Private mstrPaperSet As String      '格式
Private mblnChildForm As Boolean
Private mstrSubHead As String       '表上标签
Private mstrSubEnd As String        '表下标签
Private mstrTabHead As String       '表头单元
Private mstrColWidth As String      '列宽序列串
Private mstrColumns As String       '当前护理文件各列对应的项目
Private lngCurColor As Long, strCurFont As String, objFont As StdFont
'保存打开护理记录文件的SQL，在其它地方也有使用，不能修改
Private mstrSQL内 As String
Private mstrSQL中 As String
Private mstrSQL列 As String
Private mstrSQL条件 As String
Private mstrSQL As String

'######################################################################################################################
'**********************************************************************************************************************
'以#分隔的区域内的代码都与绘图相关,没事别动
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

Private Const WHITE_BRUSH = 0    '白色画笔
Private Const cdblWidth As Double = 6          '一个英文字符的宽度
Private Const cHideCols = 3         '前缀隐藏列:备用,时间,选择
Private Const cControlFields = 2    '记录集控制列:页号,行号

Private Function GetRBGFromOLEColor(ByVal dwOleColour As Long) As Long
    '将VB的颜色转换为RGB表示
    Dim clrref As Long
    Dim r As Long, g As Long, b As Long
    
    OleTranslateColor dwOleColour, 0, clrref
    
    b = (clrref \ 65536) And &HFF
    g = (clrref \ 256) And &HFF
    r = clrref And &HFF
    
    GetRBGFromOLEColor = RGB(r, g, b)
End Function

Private Function GetSymbolWidth(ByVal strPara As String) As Double
    '缺省是宋体9号,按字体大小同比放大
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
    '绘图相关
    Dim lngPen As Long
    Dim lngOldPen As Long
    Dim lngBrush As Long
    Dim lngOldBrush As Long
    Dim lpPoint As POINTAPI
    Dim T_ClientRect As RECT
    On Error GoTo errHand
    '******************************************
    '在此事件中不能对单元格的任何属性赋值,包括Celldata,否则会引起该事件的死循环,导致工具栏或计时器无法正常工作。
    '******************************************
    '使用匹配的背景色，前景色与字体进行文本输出。
    If Not mblnInit Then Exit Sub
    If VsfData.RowHidden(ROW) Then Exit Sub
    Done = False
    
    strText = VsfData.TextMatrix(ROW, COL)
    If IsDiagonal(COL) And InStr(1, strText, "/") <> 0 Then
        blnDraw = True
        '赋初值
        strLeft = Split(strText, "/")(0)
        strRight = Mid(strText, InStr(1, strText, "/") + 1)
        lngLeft = LenB(StrConv(strLeft, vbFromUnicode))
        lngRight = LenB(StrConv(strRight, vbFromUnicode))
        '取字符宽度
        dblWidth = GetSymbolWidth(strRight)
        '设定客户区域大小
        With T_ClientRect
            .Left = Left + 1
            .Top = Top + 1
            .Right = Right - 1
            .Bottom = Bottom - 1
        End With
        
        '1、清空内容
        '创建与背景色相同的刷子
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
        '使用该刷子填充背景色
        lngOldBrush = SelectObject(hDC, lngBrush)
        Call FillRect(hDC, T_ClientRect, lngBrush)
        '立即销毁临时使用的刷子并还原刷子
        Call SelectObject(hDC, lngOldBrush)
        Call DeleteObject(lngBrush)
        
        '2、准备画线
        '创建新画笔
        Call SetTextColor(hDC, lngForeColor)
        lngPen = CreatePen(0, 1, lngForeColor)
        lngOldPen = SelectObject(hDC, lngPen)
        '画线
        Call MoveToEx(hDC, Left, Bottom - 2, lpPoint)
        Call LineTo(hDC, Right, Top)
        '输出文本
        Call TextOut(hDC, Left, Top, strLeft, lngLeft)
        Call TextOut(hDC, IIf(Right - dblWidth >= Left, Right - dblWidth, Left), Bottom - 16, strRight, lngRight)
        
        '还原画笔并销毁
        Call SelectObject(hDC, lngOldPen)
        Call DeleteObject(lngPen)
        
        '已完成作图
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
'以#分隔的区域内的代码都与分行相关,没事别动
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
        If lngPos > 0 Then Exit Sub     '不为零,表示仅设置字符串结束符
    Next
    strLine(1) = 1
End Sub

Private Function TrimStr(ByVal str As String) As String
'功能：去掉字符串中\0以后的字符，并且去掉两端的空格

    If InStr(str, Chr(0)) > 0 Then
        TrimStr = Trim(Left(str, InStr(str, Chr(0)) - 1))
    Else
        TrimStr = Trim(str)
    End If
End Function

Private Function TruncZero(ByVal strInput As String) As String
'功能：去掉字符串中\0以后的字符
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
    
    '设置控件坐标（左边或右边超出屏幕大小则靠右或靠左显示，否则以列为中心显示）
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
            RaiseEvent AfterRowColChange("至少存在一个婴儿，请与开发商联系！", True, mblnSign, mblnArchive)
        End If
    End With
    
    '加载婴儿数据信息
    With vsfBaby
        .FixedCols = 0
        .FixedRows = 0
        .Rows = cboBaby.ListCount
        For i = 0 To cboBaby.ListCount - 1
            .RowData(i) = cboBaby.ItemData(i)
            .TextMatrix(i, 0) = "婴儿" & .RowData(i)
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
    gstrSQL = " Select   入院日期 AS 开始时间 From 病案主页 Where 病人ID=[1] And 主页ID=[2]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取入院日期或出生日期", mlng病人ID, mlng主页ID)
    GetPeriod = Format(rsTemp!开始时间, "yyyy-MM-dd HH:mm:ss") & "～" & Format(mstr结束时间, "yyyy-MM-dd HH:mm:ss")
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
    
    '读取文件属性
    mstrCOLNothing = ""
    mblnDateAd = False
    mblnDate = False
    Call GetFileProperty
    
    '读取病历文件格式定义
    gstrSQL = "Select   d.对象序号, d.内容文本, d.要素名称" & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表格样式'" & _
        " Order By d.对象序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取病历文件格式定义", mlng格式ID)
    With rsTemp
        Do While Not .EOF
            Select Case "" & !要素名称
            Case "表头层数": mintTabTiers = Val("" & !内容文本)
            Case "总列数"
                VsfData.Cols = Val("" & !内容文本)
                vsfHead.Cols = VsfData.Cols
            Case "最小行高"
                VsfData.RowHeightMin = BlowUp(Val("" & !内容文本))
                vsfHead.RowHeightMin = VsfData.RowHeightMin
            Case "文本字体"
                strCurFont = "" & !内容文本
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = BlowUp(Val(Split(strCurFont, ",")(1)))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
                End With
                Set VsfData.Font = objFont
                Set vsfHead.Font = objFont
                Set lblSubhead.Font = VsfData.Font
                Set lblSubEnd.Font = VsfData.Font
                Set Font = lblSubhead.Font
                
            Case "文本颜色"
                VsfData.ForeColor = Val("" & !内容文本)
                vsfHead.ForeColor = VsfData.ForeColor
            Case "表格颜色"
                VsfData.GridColor = Val("" & !内容文本): VsfData.GridColorFixed = VsfData.GridColor
                vsfHead.GridColor = VsfData.GridColor: vsfHead.GridColorFixed = VsfData.GridColorFixed
            Case "标题文本"
                lblTitle.Caption = "" & !内容文本
                lblTitle.AutoSize = True
            Case "标题字体"
                strCurFont = "" & !内容文本
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = BlowUp(Val(Split(strCurFont, ",")(1)))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
                End With
                Set lblTitle.Font = objFont
                lblTitle.AutoSize = False
            
            Case "开始时间": mintTagFormHour = Val("" & !内容文本)
            Case "终止时间": mintTagToHour = Val("" & !内容文本)
            Case "条件字体"
                strCurFont = "" & !内容文本
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = BlowUp(Val(Split(strCurFont, ",")(1)))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "粗") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "斜") > 0 Then .Italic = True
                End With
                Set mobjTagFont = objFont
            Case "条件颜色"
                mlngTagColor = Val("" & !内容文本)
            End Select
            .MoveNext
        Loop
    End With
    
    gstrSQL = "Select 格式 From 病历页面格式 Where 种类 = 3 And 编号 In (Select 页面 From 病历文件列表 Where Id = [1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取病历页面格式", mlng格式ID)
    If Not rsTemp.EOF Then
        mstrPaperSet = "" & rsTemp!格式
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select   d.对象序号, d.内容文本, d.要素名称, Nvl(d.是否换行, 0) As 是否换行" & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表上标签'" & _
        " Order By d.对象序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取表上标签定义", mlng格式ID)
    With rsTemp
        mstrSubHead = ""
        Do While Not .EOF
            mstrSubHead = mstrSubHead & "|" & IIf(!是否换行 = 0, "", vbCrLf) & !内容文本 & "{" & !要素名称 & "}"
            .MoveNext
        Loop
        If mstrSubHead <> "" Then mstrSubHead = Replace(Mid(mstrSubHead, 2), Chr(1), " ")
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select   d.对象序号, d.内容文本, d.要素名称, Nvl(d.是否换行, 0) As 是否换行" & _
        " From 病历文件结构 d, 病历文件结构 p" & _
        " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表下标签'" & _
        " Order By d.对象序号"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取表上标签定义", mlng格式ID)
    With rsTemp
        mstrSubEnd = ""
        Do While Not .EOF
            mstrSubEnd = mstrSubEnd & "|" & IIf(!是否换行 = 0, "", vbCrLf) & !内容文本 & "{" & !要素名称 & "}"
            .MoveNext
        Loop
        If mstrSubEnd <> "" Then mstrSubEnd = Replace(Mid(mstrSubEnd, 2), Chr(1), " ")
    End With
    '------------------------------------------------------------------------------------------------------------------
    '检查是否绑定了日期列
    gstrSQL = "Select  d.对象序号, d.对象属性, d.内容行次, d.内容文本, d.要素名称, d.要素单位,d.要素表示" & vbNewLine & _
            "        From 病历文件结构 d, 病历文件结构 p" & vbNewLine & _
            "        Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表列集合'" & vbNewLine & _
            "        And D.要素名称='日期'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取表列集合定义", mlng格式ID)
    mblnDate = (rsTemp.RecordCount > 0)
    If mblnDate = False Then
        VsfData.Cols = VsfData.Cols + 1
        vsfHead.Cols = VsfData.Cols
    End If
    '------------------------------------------------------------------------------------------------------------------
    '不存在日期列默认添加日期列
    If mblnDate = False Then
        gstrSQL = "SELECT 对象序号, 内容行次, 内容文本" & vbNewLine & _
                "FROM (SELECT 1 对象序号, 1 内容行次, '日期' 内容文本" & vbNewLine & _
                "       FROM DUAL" & vbNewLine & _
                "       UNION ALL" & vbNewLine & _
                "       SELECT D.对象序号+1 对象序号, D.内容行次, D.内容文本" & vbNewLine & _
                "       FROM 病历文件结构 D, 病历文件结构 P" & vbNewLine & _
                "       WHERE P.ID = D.父ID AND P.文件ID = [1] AND P.对象类型 = 1 AND P.内容文本 = '表头单元')" & vbNewLine & _
                "ORDER BY 对象序号"
    Else
        gstrSQL = "Select   d.对象序号, d.内容行次, d.内容文本" & _
            " From 病历文件结构 d, 病历文件结构 p" & _
            " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表头单元'" & _
            " Order By d.对象序号"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取表头单元定义", mlng格式ID)
    With rsTemp
        mstrTabHead = ""
        Do While Not .EOF
            mstrTabHead = mstrTabHead & "|" & !内容行次 - 1 & "," & !对象序号 & "," & !内容文本
            .MoveNext
        Loop
        If mstrTabHead <> "" Then mstrTabHead = Mid(mstrTabHead, 2)
    End With
    
    '查询语句组织
    '------------------------------------------------------------------------------------------------------------------
    Dim strSql外 As String, str格式 As String, strSqlNull As String
    Dim bln日期 As Boolean, bln时间 As Boolean, bln护士 As Boolean
    Dim bln签名人 As Boolean, bln签名时间 As Boolean, bln签名日期 As Boolean
    Dim bln对角线 As Boolean         '如果上一列是对角线且选择项,则直接提取各项数据,拼列头时在数值间加上/
    Dim lngColumn As Long
      
    If mblnDate = False Then
        gstrSQL = "SELECT 对象序号, 对象属性, 内容行次, 内容文本, 要素名称, 要素单位, 要素表示" & vbNewLine & _
            "FROM (SELECT 1 对象序号, '0`4' 对象属性, 1 内容行次, '' 内容文本, '日期' 要素名称, '' 要素单位, 0 要素表示" & vbNewLine & _
            "       FROM DUAL" & vbNewLine & _
            "       UNION ALL" & vbNewLine & _
            "       SELECT D.对象序号+1 对象序号, D.对象属性, D.内容行次, D.内容文本, D.要素名称, D.要素单位, D.要素表示" & vbNewLine & _
            "       FROM 病历文件结构 D, 病历文件结构 P" & vbNewLine & _
            "       WHERE P.ID = D.父ID AND P.文件ID = [1] AND P.对象类型 = 1 AND P.内容文本 = '表列集合')" & vbNewLine & _
            "ORDER BY 对象序号, 内容行次"
    Else
        gstrSQL = "Select   d.对象序号, d.对象属性, d.内容行次, d.内容文本, d.要素名称, d.要素单位,d.要素表示 " & _
            " From 病历文件结构 d, 病历文件结构 p" & _
            " Where p.Id = d.父id And p.文件id = [1] And p.对象类型 = 1 And p.内容文本 = '表列集合'" & _
            " Order By d.对象序号, d.内容行次"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "读取表列集合定义", mlng格式ID)
    With rsTemp
        lngColumn = 0: mstrColumns = "": mstrColWidth = "": mstrCatercorner = ""
        mstrSQL内 = "": mstrSQL中 = "": strSql外 = "": mstrSQL列 = "": mstrSQL条件 = "": strSqlNull = ""
        bln日期 = False: bln时间 = False: bln护士 = False
        bln签名人 = False: bln签名时间 = False: bln签名日期 = False
        Do While Not .EOF
            If lngColumn <> !对象序号 Then
                mstrColumns = mstrColumns & IIf(mstrColumns = "", "", "'1'" & str格式) & "|" & !对象序号 & "'" & !要素名称
                mstrColWidth = mstrColWidth & "," & !对象属性 & "`" & !对象序号 & "`" & !要素表示
                If !要素表示 = 1 Then mstrCatercorner = mstrCatercorner & "," & !对象序号
                str格式 = ""
                If !要素名称 <> "" Then str格式 = "{" & NVL(!内容文本) & "[" & !要素名称 & "]" & NVL(!要素单位) & "}"
                
                If Mid(strSqlNull, 3) = "" Then
                    strSqlNull = "''"
                Else
                    strSqlNull = Mid(strSqlNull, 3)
                End If
                mstrSQL列 = mstrSQL列 & "," & IIf(Mid(strSql外, 3) = "", "''", "Decode(" & Mid(strSql外, 3) & "," & strSqlNull & ",''," & Mid(strSql外, 3) & ")") & " As C" & Format(lngColumn, "00")
                
                strSql外 = ""
                strSqlNull = ""
                lngColumn = !对象序号
                bln对角线 = (NVL(!要素表示, 0) = 1)
            Else
                mstrColumns = mstrColumns & "," & !要素名称
                str格式 = str格式 & "{" & NVL(!内容文本) & "[" & !要素名称 & "]" & NVL(!要素单位) & "}"
            End If
            
            Select Case !要素名称
            Case "日期"
                bln日期 = True
                mblnDateAd = (NVL(!要素表示, 0) = 1)
                mstrSQL中 = mstrSQL中 & ",日期"
                mstrSQL内 = mstrSQL内 & ",To_Char(l.发生时间, " & IIf(mblnDateAd, "'dd/MM'", "'yyyy-mm-dd'") & ") As 日期"
                strSql外 = strSql外 & "||" & !要素名称
            Case "时间"
                bln时间 = True
                mstrSQL中 = mstrSQL中 & ",时间"
                mstrSQL内 = mstrSQL内 & ",To_Char(l.发生时间, 'hh24:mi') As 时间"
                strSql外 = strSql外 & "||" & !要素名称
                
            Case "签名人"
                bln签名人 = True
                mstrSQL中 = mstrSQL中 & ",签名人"
                mstrSQL内 = mstrSQL内 & ",l.签名人"
                strSql外 = strSql外 & "||" & !要素名称
                
            Case "签名时间"
                bln签名时间 = True
                mstrSQL中 = mstrSQL中 & ",签名时间"
                mstrSQL内 = mstrSQL内 & ",l.签名时间"
                strSql外 = strSql外 & "||" & !要素名称
                
            Case "护士"
                bln护士 = True
                mstrSQL中 = mstrSQL中 & ",护士"
                mstrSQL内 = mstrSQL内 & ",l.保存人 As 护士"
                strSql外 = strSql外 & "||" & !要素名称
            Case Else
                If !要素名称 <> "" Then
                    mstrSQL中 = mstrSQL中 & ",Max(""" & !要素名称 & """) As """ & !要素名称 & """"
                    mstrSQL条件 = mstrSQL条件 & " Or """ & !要素名称 & """ Is Not Null"
                    
                    strSql外 = strSql外 & "||'" & !内容文本 & "'||""" & !要素名称 & """||'" & !要素单位 & "'"
                    strSqlNull = strSqlNull & "||" & "'" & !内容文本 & "'||'" & !要素单位 & "'"
                    mstrSQL内 = mstrSQL内 & ", Decode(c.项目名称, '" & !要素名称 & "', Nvl(c.未记说明,c.记录内容), '') As """ & !要素名称 & """"

                Else
                    mstrCOLNothing = mstrCOLNothing & "," & Val(Format(!对象序号, "00"))
                End If
            End Select
            .MoveNext
        Loop
        
        mstrCOLNothing = Mid(mstrCOLNothing, 2)
        mstrCatercorner = Mid(mstrCatercorner, 2)
        mstrColWidth = Mid(mstrColWidth, 2)
        '加入最后一列的格式
        mstrColumns = mstrColumns & IIf(mstrColumns = "", "", "'1'" & str格式) '& "|" & !对象序号 & "'" & !要素名称
        mstrColumns = Mid(mstrColumns, 2)     '格式如:列号;项目名称1,项目名称2|列号...,实例;1;体温|2;脉搏|3...
        If Mid(strSql外, 3) <> "" Then
            mstrSQL列 = mstrSQL列 & "," & Mid(strSql外, 3) & " As C" & Format(lngColumn, "00")
        Else
            mstrSQL列 = mstrSQL列 & ",'' As C" & Format(lngColumn, "00")
        End If
        
        If mstrSQL条件 <> "" Then mstrSQL条件 = "(" & Mid(mstrSQL条件, 5) & ")"
        
        '如果没有出现日期，时间，护士，则内层需要补充，以保证中层分组的正常：
        If bln日期 = False Then mstrSQL内 = mstrSQL内 & ",To_Char(l.发生时间, 'yyyy-mm-dd') As 日期"
        If bln时间 = False Then mstrSQL内 = mstrSQL内 & ",To_Char(l.发生时间, 'hh24:mi') As 时间"
        
        If bln签名人 = False Then mstrSQL内 = mstrSQL内 & ",l.签名人 As 签名人"
        If bln签名时间 = False Then mstrSQL内 = mstrSQL内 & ",l.签名时间"
        
        If Mid(mstrSQL中, 2) = "" Then
            MsgBox "对不起，您没有定义当前护理单的显示列信息，请在病历文件管理中定义！", vbInformation, gstrSysName
            Exit Function
        End If
        
        '程序内部控制增加固定列
        mstrSQL中 = mstrSQL中 & ",MAX(签名级别) AS 签名级别,MAX(签名信息) AS 签名信息,MAX(记录ID) AS 记录ID,MAX(行数) AS 行数"
        mstrSQL内 = mstrSQL内 & ",l.签名级别,l.签名人 AS 签名信息,C.记录ID,'' AS 行数"
        mstrSQL列 = mstrSQL列 & ",签名级别,签名信息,记录ID,行数"
        
        If bln护士 = False Then
            '强制添加护士列,为了避免修改他人数据行(他人录入的数据,增加新数据也不允许)
            mstrSQL中 = mstrSQL中 & ",护士"
            mstrSQL内 = mstrSQL内 & ",l.保存人 As 护士"
            mstrSQL列 = mstrSQL列 & ",护士"
        End If
        
        '将活动项目加入到SQL中
        Call SQLCombination
    End With
    
    ReadStruDef = True
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub SQLCombination(Optional ByVal lng记录ID As Long = 0)
    Dim str条件 As String
    str条件 = mstrSQL条件 & IIf(lng记录ID = 0, "", IIf(mstrSQL条件 = "", "", " And") & " 记录ID=[6]")
    
    mstrSQL = "Select '' 备用,to_char(发生时间,'yyyy-MM-dd hh24:mi:ss') AS 发生时间,'' AS 选择," & Mid(mstrSQL列, 12) & vbCrLf & _
                " From (Select 发生时间," & Mid(mstrSQL中, 2) & vbCrLf & _
                "        From (Select l.发生时间," & Mid(mstrSQL内, 2) & vbCrLf & _
                "               From 病人护理数据 l, 病人护理明细 c,病人护理文件 f " & vbCrLf & _
                "               Where l.Id = c.记录id And l.文件ID=f.ID" & _
                "               And c.终止版本 Is Null And c.记录类型<>5  " & _
                "               And f.id=[1] And f.病人id = [2] And f.主页id = [3] And Nvl(f.婴儿,0)=[4] And l.汇总类别=[5])" & vbCrLf & _
                IIf(str条件 <> "", "Where " & str条件, "") & _
                "       Group By 日期, 时间, 发生时间,护士,签名人,签名时间" & _
                                "       Order By 发生时间,护士,签名人,签名时间)"
End Sub

Public Sub zlRefresh(Optional ByVal blnRefresh As Boolean = True)
'-----------------------------------------------------------------------
'功能：完成数据刷新,blnRefresh=false 表示只刷新表上和表签内容信息
'-----------------------------------------------------------------------
    Dim aryRow() As String, aryItem() As String, arrItemEnd() As String
    Dim strPrefix As String, strItemName As String
    Dim lngRow As Long, lngCol As Long, lngCount As Long, strCell As String, i As Integer
    Dim strTmpSQL As String
    Dim aryPeriod() As String
    Dim strtmp As String, str单位 As String
    
    Err = 0: On Error GoTo errHand
    'Debug.Print Now & "zlRefresh"
    Call InitCons
    '表上标签获取
    lblSubhead.Caption = ""
    lblSubhead.Tag = ""
    lblSubEnd.Caption = ""
    lblSubEnd.Tag = ""
    aryPeriod = Split(GetPeriod, "～")
    gstrSQL = "Select [1] || Zl_Replace_Element_Value([2],[3],[4],2,NULL,[5]) as 信息 From Dual"
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
            mrsPartogram.Filter = "中文名='" & strItemName & "'"
            '不可能找不到，除非手工修改数据
            If mrsPartogram.RecordCount = 0 Then GoTo ErrNext
            str单位 = Trim(NVL(mrsPartogram!单位))
            If Val(NVL(mrsPartogram!替换域)) = 1 Then
                '产程固定要素信息
                strtmp = strPrefix
                Select Case strItemName
                Case "当前病区"
                
                    strTmpSQL = "Select   b.名称" & vbNewLine & _
                                "From (Select 病区id, 开始时间, Nvl(终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) As 终止时间" & vbNewLine & _
                                "            From 病人变动记录" & vbNewLine & _
                                "            Where 病人id = [1] And 主页id = [2] And 科室id = [3]) a,部门表 b " & vbNewLine & _
                                "Where ([4] Between a.开始时间 And a.终止时间 Or [4] >= a.开始时间) And a.病区id Is Not Null And b.ID=a.病区id" & vbNewLine & _
                                "Order By a.开始时间"
                                
                    Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "当前病区", mlng病人ID, mlng主页ID, mlng科室id, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
                    
                Case "当前床号"
                
                    strTmpSQL = "Select   a.床号" & vbNewLine & _
                                "From (Select 床号, 开始时间, Nvl(终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) As 终止时间" & vbNewLine & _
                                "            From 病人变动记录" & vbNewLine & _
                                "            Where 病人id = [1] And 主页id = [2] And 科室id = [3]) a" & vbNewLine & _
                                "Where ([4] Between a.开始时间 And a.终止时间 Or [4] >= a.开始时间) And a.床号 Is Not Null" & vbNewLine & _
                                "Order By a.开始时间"
        
                    Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "当前床号", mlng病人ID, mlng主页ID, mlng科室id, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
                    If rsTemp.BOF = False Then rsTemp.MoveLast
                    
                Case "当前科室"
                
                    strTmpSQL = "Select   名称 From 部门表 a Where a.ID=[1]"
                    Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "当前科室", mlng科室id)
                    
                Case "住院医师"
                    strTmpSQL = "Select   a.经治医师" & vbNewLine & _
                                "From (Select 经治医师, 开始时间, Nvl(终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) As 终止时间" & vbNewLine & _
                                "            From 病人变动记录" & vbNewLine & _
                                "            Where 病人id = [1] And 主页id = [2] And 科室id = [3]) a" & vbNewLine & _
                                "Where ([4] Between a.开始时间 And a.终止时间 Or [4] >= a.开始时间) And a.经治医师 Is Not Null" & vbNewLine & _
                                "Order By a.开始时间"
                    Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "住院医师", mlng病人ID, mlng主页ID, mlng科室id, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
                    If rsTemp.BOF = False Then rsTemp.MoveLast
                Case "责任护士"
                
                    strTmpSQL = "Select   a.责任护士" & vbNewLine & _
                                "From (Select 责任护士, 开始时间, Nvl(终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) As 终止时间" & vbNewLine & _
                                "            From 病人变动记录" & vbNewLine & _
                                "            Where 病人id = [1] And 主页id = [2] And 科室id = [3]) a" & vbNewLine & _
                                "Where ([4] Between a.开始时间 And a.终止时间 Or [4] >= a.开始时间) And a.责任护士 Is Not Null" & vbNewLine & _
                                "Order By a.开始时间"
                    Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "责任护士", mlng病人ID, mlng主页ID, mlng科室id, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
                    If rsTemp.BOF = False Then rsTemp.MoveLast
                    
                Case "护理等级"
                    strTmpSQL = "Select   b.名称" & vbNewLine & _
                                "From (Select 护理等级ID, 开始时间, Nvl(终止时间, To_Date('3000-01-01', 'yyyy-mm-dd')) As 终止时间" & vbNewLine & _
                                "            From 病人变动记录" & vbNewLine & _
                                "            Where 病人id = [1] And 主页id = [2] And 科室id = [3]) a,护理等级 b" & vbNewLine & _
                                "Where ([4] Between a.开始时间 And a.终止时间 Or [4] >= a.开始时间) And a.护理等级ID Is Not Null And b.序号=a.护理等级ID" & vbNewLine & _
                                "Order By a.开始时间"
                    Set rsTemp = zlDatabase.OpenSQLRecord(strTmpSQL, "护理等级", mlng病人ID, mlng主页ID, mlng科室id, CDate(aryPeriod(0)), CDate(aryPeriod(1)))
                    If rsTemp.BOF = False Then rsTemp.MoveLast
                Case "最后诊断"
                    strtmp = ""
                    gstrSQL = "Select [1] || Zl_Replace_Element_Value([2],[3],[4],2,NULL,[5],[6]) as 信息 From Dual"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取要素", strPrefix, strItemName, mlng病人ID, mlng主页ID, mint婴儿, CDate(aryPeriod(0)))
                Case Else
                    strtmp = ""
                    gstrSQL = "Select [1] || Zl_Replace_Element_Value([2],[3],[4],2,NULL,[5]) as 信息 From Dual"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取要素", strPrefix, strItemName, mlng病人ID, mlng主页ID, mint婴儿)
                End Select
            Else
                '产程录入要素信息
                strtmp = strPrefix
                gstrSQL = "SELECT 内容 From 产程要素内容" & _
                    "   Where 文件ID = [1] And 婴儿 = [2] And 名称 =[3]"
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "去要素", mlng文件ID, mlngFileIndex, strItemName)
            End If
            If rsTemp.BOF = False Then
                If i = 0 Then
                    If strtmp <> "" Then
                        lblSubhead.Tag = lblSubhead.Tag & " " & strtmp & rsTemp.Fields(0).Value & str单位
                    Else
                        lblSubhead.Tag = lblSubhead.Tag & " " & rsTemp.Fields(0).Value & str单位
                    End If
                Else
                    If strtmp <> "" Then
                        lblSubEnd.Tag = lblSubEnd.Tag & " " & strtmp & rsTemp.Fields(0).Value & str单位
                    Else
                        lblSubEnd.Tag = lblSubEnd.Tag & " " & rsTemp.Fields(0).Value & str单位
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
    '表上标签分散处理
    Call zlLableBruit
    '产生列记录集
    Call InitRecords
    
    If blnRefresh = False Then Exit Sub
    '装入数据
    Call SQLCombination
    gstrSQL = mstrSQL
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取护理数据", mlng文件ID, mlng病人ID, mlng主页ID, mint婴儿, cboBaby.ItemData(cboBaby.ListIndex))
    '清除并拷贝记录集结构
    Call DataMap_Init(rsTemp)
    '绑定数据并设置格式,同时实现一行数据分行显示的功能
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
    '初始化内存数据集
    
    '数据记录集,用于快速恢复
    Set mrsDataMap = CopyNewRec(rsSource)
    mrsDataMap.Sort = "页号,行号"
    '修改单元格记录,用于保存
    Call Record_Init(mrsCellMap, "ID," & adLongVarChar & ",50|页号," & adDouble & ",18|行号," & adDouble & ",18|" & _
            "列号," & adDouble & ",18|记录ID," & adDouble & ",18|数据," & adLongVarChar & ",4000|部位," & adLongVarChar & ",100|" & _
            "汇总," & adDouble & ",1|删除," & adDouble & ",1")
    mrsCellMap.Sort = "页号,行号,列号"
    '复制记录集
    Set mrsCopyMap = New ADODB.Recordset
    Set mrsCopyMap = CopyNewRec(mrsDataMap, False)
End Sub

Private Function DataMap_Save() As Boolean
    '将当前页面中用户编辑过的数据保存起来,页面切换或保存前触发
    Dim blnExit As Boolean
    Dim lngRow As Long, lngCol As Long, lngRows As Long, lngCols As Long
    Dim strDate As String, strTime As String, strDatetime As String, strCurrDate As String
    On Error GoTo errHand
    
    '不管是否编辑过都保存
    If Not CheckFlip Then Exit Function
    
    '先删除指定页号的所有数据行
    mrsDataMap.Filter = "页号=" & mint页码
    Do While True
        If mrsDataMap.RecordCount = 0 Then Exit Do
        mrsDataMap.Delete
        mrsDataMap.MoveNext
    Loop
    mrsDataMap.Filter = 0
    
    strCurrDate = Format(zlDatabase.Currentdate, "YYYY-MM-DD HH:mm")
    '复制指定页号的所有数据行
    lngRows = VsfData.Rows - 1
    lngCols = VsfData.Cols - 1
    For lngRow = VsfData.FixedRows To lngRows
        mrsDataMap.AddNew
        mrsDataMap!页号 = mint页码
        mrsDataMap!行号 = lngRow
        mrsDataMap!删除 = IIf(VsfData.RowHidden(lngRow), 1, 0)
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
        mrsDataMap!产程日期 = Format(strDatetime, "YYYY-MM-DD HH:mm:ss")
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
    
    '更新当前页面所有大于起始行的行号数据
    With mrsCellMap
        If .RecordCount <> 0 Then .MoveLast
        If .BOF Then Exit Sub
        Do While Not .BOF
            If !页号 = mint页码 And !行号 > lngStart Then
                intCOl = !列号
                lngPos = .AbsolutePosition
                !行号 = !行号 + lngDeff
                !ID = mint页码 & "," & !行号 & "," & !列号
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
    '只拷贝记录集的结构,同时增加页号,行号字段
    Dim rsTarget As New ADODB.Recordset
    Dim intFields As Integer
    
    Set rsTarget = New ADODB.Recordset
    With rsTarget
        If blnAddPage Then
            .Fields.Append "页号", adDouble, 18
            .Fields.Append "行号", adDouble, 18
        End If
        For intFields = 0 To rsSource.Fields.Count - 1
            If rsSource.Fields(intFields).Name = "汇总日期" Then
                .Fields.Append rsSource.Fields(intFields).Name, adLongVarChar, 50, adFldIsNullable      '0:表示新增
            ElseIf rsSource.Fields(intFields).Type = 200 Then       '日期型处理为字符型
                .Fields.Append rsSource.Fields(intFields).Name, adLongVarChar, rsSource.Fields(intFields).DefinedSize, adFldIsNullable     '0:表示新增
            Else
                .Fields.Append rsSource.Fields(intFields).Name, IIf(rsSource.Fields(intFields).Type = adNumeric, adDouble, rsSource.Fields(intFields).Type), rsSource.Fields(intFields).DefinedSize, adFldIsNullable    '0:表示新增
            End If
        Next
        If blnAddPage Then
            .Fields.Append "删除", adDouble, 1
            .Fields.Append "产程日期", adLongVarChar, 20, adFldIsNullable
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
    '如果一行显示不完则分行显示(根据当前实际数据计算占用行数,并完成行数添加和赋值,然后再行其它列数据处理)
    lngRow = VsfData.FixedRows
    Do While True
        If lngRow > VsfData.Rows - 1 Then Exit Do
        '计算数据最大行数,并完成行数添加和赋值
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
                    '循环赋值
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
            '循环处理当前行数据
            For lngCol = VsfData.FixedCols To VsfData.Cols - 1
                If VsfData.ColHidden(lngCol) And lngCol <> mlngRowCount Then
                    '循环赋值
                    For intData = 2 To lngMaxRowCount
                        VsfData.TextMatrix(lngRow + intData - 1, lngCol) = VsfData.TextMatrix(lngRow, lngCol)
                    Next
                ElseIf lngCol = mlngNoEditor Then
                    '将行值改为从1开始,比如有4行数据,就是4|1
                    For intData = 1 To lngMaxRowCount
                        VsfData.TextMatrix(lngRow + intData - 1, mlngRowCount) = lngMaxRowCount & "|" & intData
                    Next
                    '最后一行需要填写封闭签名
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
    
    '设置护理数据编辑界面的格式
    With vsfHead
        .Redraw = flexRDNone
        .Clear
        Set .DataSource = rsTemp
        .Rows = 3
        
        '表头填写
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFree
        .MergeRow(0) = True
        .MergeRow(1) = True
        .MergeRow(2) = True
        
        '程序内部控制列隐藏
        .ColHidden(1) = True
        .ColHidden(2) = True
        .ColHidden(mlngChoose) = Not mblnVerify
        .ColHidden(mlngRowCount) = True
        .ColHidden(mlngRecord) = True
        .ColHidden(mlngSigner) = True
        .ColHidden(mlngSignLevel) = True
        If mlngOperator = -1 Then .ColHidden(.Cols - 1) = True
        .ColWidth(0) = 250
        .ColWidth(mlngChoose) = 250      '选择列
        
        .FrozenCols = mlngTime
        .SheetBorder = &H40C0&
        
        '设置列头
        aryItem = Split(mstrTabHead, "|")
        For lngCount = 0 To UBound(aryItem)
            strCell = aryItem(lngCount)
            lngRow = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            lngCol = Left(strCell, InStr(1, strCell, ",") - 1): strCell = Mid(strCell, InStr(1, strCell, ",") + 1)
            .TextMatrix(lngRow, lngCol + cHideCols + .FixedCols - 1) = strCell
        Next
        
        '设置固定列及选择列
        .TextMatrix(0, 0) = " "
        .TextMatrix(1, 0) = " "
        .TextMatrix(2, 0) = " "
        .TextMatrix(0, mlngChoose) = " "
        .TextMatrix(1, mlngChoose) = " "
        .TextMatrix(2, mlngChoose) = " "
        
        '列宽设置
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
        
        '固定行格式为居中
        .Cell(flexcpAlignment, 0, 0, .FixedRows - 1, .Cols - 1) = flexAlignCenterCenter
        '再按列合并
        For lngCount = 0 To .Cols - 1
            .MergeCol(lngCount) = True
        Next
        
        If blnAlign = False Then
            '改为根据用户的设置显示列对齐方式
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
        
        '可能固定行的行高不正确需要自动调整下
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

        '程序内部控制列隐藏
        .ColHidden(1) = True
        .ColHidden(2) = True
        .ColHidden(mlngChoose) = Not mblnVerify
        .ColHidden(mlngRowCount) = True
        .ColHidden(mlngRecord) = True
        .ColHidden(mlngSigner) = True
        .ColHidden(mlngSignLevel) = True
        If mlngOperator = -1 Then .ColHidden(.Cols - 1) = True
        .ColWidth(0) = 250
        .ColWidth(mlngChoose) = 250      '选择列
        
        .FrozenCols = mlngTime
        .SheetBorder = &H40C0&
        
        '列宽设置
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
            '改为根据用户的设置显示列对齐方式
            If .FixedRows < .Rows Then .Cell(flexcpAlignment, .FixedRows, 0, .Rows - 1, .Cols - 1) = flexAlignGeneralCenter
        End If
        
        Call PreTendMutilRows
        Call WriteColor
        
        '将非固定行的行高设置为最小行高
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
    '晚班以红色显示，同时将非起始行设置为NoCheckBox，设置图标
    With VsfData
        For lngCount = .FixedRows To .Rows - 1
            If .TextMatrix(lngCount, 2) <> "" Then
                '晚班以红色显示
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
            
            
            '将非起始行设置为NoCheckBox
            If Not VsfData.RowHidden(lngCount) Then
                If Not VsfData.TextMatrix(lngCount, mlngRowCount) Like "*|1" Then
                    VsfData.Cell(flexcpChecked, lngCount, mlngChoose) = flexNoCheckbox
                Else
                    If VsfData.Cell(flexcpChecked, lngCount, mlngChoose) <> flexTSChecked Then
                        VsfData.Cell(flexcpChecked, lngCount, mlngChoose) = flexTSUnchecked
                    End If
                    
                    '设置图标
                    If VsfData.TextMatrix(lngCount, mlngSigner) = "" Then
                        VsfData.Cell(flexcpPicture, lngCount, 0) = Nothing
                    Else
                        If InStr(1, VsfData.TextMatrix(lngCount, mlngSigner), "/") <> 0 Then
                            VsfData.Cell(flexcpPicture, lngCount, 0) = imgRow.ListImages(审签).Picture
                        Else
                            VsfData.Cell(flexcpPicture, lngCount, 0) = imgRow.ListImages(签名).Picture
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
    '提取文件属性
    Dim strEnd As String
    On Error GoTo errHand
    
    gstrSQL = " Select   开始时间,结束时间,格式ID,科室ID,归档人 From 病人护理文件 " & _
              " Where 病人ID=[1] And 主页ID=[2] And 婴儿=[3] And ID=[4] And Rownum<2"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取护理文件数据", mlng病人ID, mlng主页ID, mint婴儿, mlng文件ID)
    If rsTemp.RecordCount <> 0 Then
        mlng格式ID = rsTemp!格式ID
        mlng科室id = rsTemp!科室ID
        mblnArchive = (NVL(rsTemp!归档人) <> "")
        mstr开始时间 = Format(rsTemp!开始时间, "yyyy-MM-dd HH:mm")
        mstr结束时间 = Format(rsTemp!结束时间, "yyyy-MM-dd HH:mm")
        mstrBeginTime = mstr开始时间
        strEnd = DateAdd("n", -1, CDate(Format(CDate(mstr开始时间) + 1, "yyyy-MM-dd HH:mm:ss")))
        If mstr结束时间 = "" Then
            mstr结束时间 = Format(strEnd, "YYYY-MM-DD HH:mm:ss")
        Else
            If (mstr结束时间 <> "" And CDate(mstr结束时间) > CDate(strEnd)) Then mstr结束时间 = Format(strEnd, "yyyy-MM-dd HH:mm:ss")
        End If
    End If
    
    '第二份文件重新提取开始时间
    If mlngFileIndex > 1 Then
        gstrSQL = "SELECT Max(B.发生时间) 发生时间" & vbNewLine & _
            "FROM 病人护理文件 A,病人护理数据 B,病人护理明细 C,护理记录项目 D" & vbNewLine & _
            "WHERE A.ID=B.文件ID AND B.ID=C.记录ID AND A.ID=[1] And 病人ID=[2] And 主页ID=[3] And 婴儿=[4] AND B.汇总类别<[5] AND C.项目序号=D.项目序号" & vbNewLine & _
            "AND NVL(D.项目名称,'')='生产' AND NVL(D.保留项目,1)=1 ORDER BY B.发生时间"
        Call SQLDIY(gstrSQL)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取护理文件数据", mlng文件ID, mlng病人ID, mlng主页ID, mint婴儿, mlngFileIndex)
        If rsTemp.RecordCount <> 0 Then
            mstr开始时间 = DateAdd("n", 1, CDate(Format(rsTemp!发生时间, "yyyy-MM-dd HH:mm")))
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
    
    glngHours = Val(zlDatabase.GetPara("数据补录时限", glngSys))
    '打开现存在的所有护理记录项目
    gstrSQL = " Select   项目序号,项目名称,项目类型,项目性质,项目长度,项目小数,项目表示,项目单位,项目值域,护理等级,应用方式" & _
              " From 护理记录项目 B" & _
              " Order by 项目序号"
    Set mrsItems = zlDatabase.OpenSQLRecord(gstrSQL, "打开现存在的所有护理记录项目")
    '提取所有产程要素信息
    gstrSQL = "Select 中文名,替换域,类型,长度,小数,单位,表示法,数值域,必填" & vbNewLine & _
        "From (Select i.分类id, i.编码, i.中文名, nvl(i.替换域,0) 替换域,i.类型,i.长度,i.小数,i.单位,i.表示法,i.数值域,i.必填" & vbNewLine & _
        "       From 诊治所见项目 I, 诊治所见分类 K" & vbNewLine & _
        "       Where k.Id = i.分类id And k.编码 In ('02', '03', '05', '06') And i.替换域 = 1 And k.性质 = 1" & vbNewLine & _
        "       Union" & vbNewLine & _
        "       Select i.分类id, i.编码, i.中文名, nvl(i.替换域,0) 替换域,i.类型,i.长度,i.小数,i.单位,i.表示法,i.数值域,i.必填" & vbNewLine & _
        "       From 诊治所见项目 I, 诊治所见分类 K" & vbNewLine & _
        "       Where k.Id = i.分类id And k.编码 In ('04', '05') And k.性质 = 2)" & vbNewLine & _
        "Order By 分类id, 编码, 替换域"

    Set mrsPartogram = zlDatabase.OpenSQLRecord(gstrSQL, "提取产程要素信息")
    '取当前操作员的级别
    mintVerify = 未定义
    mintVerify_Last = 未定义
    gstrSQL = "select  聘任技术职务 from 人员表 p where ID=[1]"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "mRichEPR", glngUserId)
    If Not rs.EOF Then
        mintVerify = NVL(rs("聘任技术职务"), 未定义)
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
        '初始化内存记录集(未对应项目的列为活动项目,其它列均为固定项)
        mstrFields = "列," & adDouble & ",18|序号," & adDouble & ",2|项目序号," & adDouble & ",18|项目名称," & adLongVarChar & ",20|固定," & adDouble & ",2|格式," & adLongVarChar & ",2000"
        Call Record_Init(mrsSelItems, mstrFields)
        mstrFields = "列|序号|项目序号|项目名称|固定|格式"
    End If
    
    '加入列定义
    If Not mblnInit Then
        arrColumn = Split(strColumns, "|")
        j = UBound(arrColumn)
        For i = 0 To j
            lngCol = Split(arrColumn(i), "'")(0)
            arrItem = Split(Split(arrColumn(i), "'")(1), ",")
            blnSet = False   '如果已设置以传入值为准'否则找不到项目就是活动项目
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
                mrsItems.Filter = "项目名称='" & strName & "'"
                If mrsItems.RecordCount <> 0 Then
                    lngOrder = mrsItems!项目序号
                    If Not blnSet Then intImmovable = 1   '固定不允许修改
                    Select Case strName
                        Case "宫口扩大"
                            mlngSpread = i + cHideCols + VsfData.FixedCols
                        Case "先露高低"
                            mlngJust = i + cHideCols + VsfData.FixedCols
                        Case "生产"
                            mlngProduce = i + cHideCols + VsfData.FixedCols
                    End Select
                Else
                    lngOrder = 0
                    If Not blnSet Then intImmovable = 0
                    
                    '记录特殊列
                    Select Case strName
                    Case "日期"
                        mlngDate = i + cHideCols + VsfData.FixedCols
                    Case "时间"
                        mlngTime = i + cHideCols + VsfData.FixedCols
                    Case "护士"
                        mlngOperator = i + cHideCols + VsfData.FixedCols
                    Case "签名人"
                        mlngSignName = i + cHideCols + VsfData.FixedCols
                    Case "签名时间"
                        mlngSignTime = i + cHideCols + VsfData.FixedCols
                    End Select
                End If
                mstrValues = lngCol & "|" & l + 1 & "|" & lngOrder & "|" & strName & "|" & intImmovable & "|" & strFormat
                Call Record_Add(mrsSelItems, mstrFields, mstrValues)
            Next
        Next
        
        'Call OutputRsData(mrsSelItems)
        
        '加入程序内部控制列(列是在读取数据后绑定时增加的,此时只有预处理下)
        mlngDemo = 1
        mlngActTime = 2
        mlngChoose = 2 + VsfData.FixedCols
        mlngSignLevel = VsfData.Cols + cHideCols + VsfData.FixedCols '加上隐藏列
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

Public Function SignMe(Optional ByVal bln审签 As Boolean = False) As Boolean
    Dim blnSign As Boolean          '是否签名成功
    Dim blnRefresh As Boolean
    Dim strSignTime As String       '保证所有签名的签名时间一致,便于取消签名时按签名时间统一取消
    Dim str状态 As String           '保存签名选项,避免循环签名时不停的弹出签名窗口
    Dim str行错误 As String
    Dim str错误 As String
    Dim intRow As Integer, intRows As Integer
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '按发生时间循环对所有未签名数据进行签名
    
    If mlng病人ID = 0 Then Exit Function
    
    '普签:对所有未签名的数据进行签名
    '审签:对所有已签名的数据进行审签
    If bln审签 Then
        If Not mblnVerify Then
            '汇总数据也要签名,因此去掉条件: And B.汇总类别=0
            gstrSQL = " Select  distinct B.发生时间 " & vbNewLine & _
                      " From 病人护理明细 A,病人护理数据 B,病人护理文件 C" & vbNewLine & _
                      " Where A.记录ID=B.ID And B.文件ID=C.ID And A.数据来源=0 And MOD(A.记录类型,10)=5 AND A.终止版本 Is NULL And C.ID=[1] "
            Call SQLDIY(gstrSQL)
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查是否存在未签名的数据", mlng文件ID)
            If rsTemp.RecordCount = 0 Then
                RaiseEvent AfterRowColChange("不存在已签名的数据！", True, mblnSign, mblnArchive)
                Exit Function
            End If
        
            '进入审签模式,可修改数据,可勾选数据
            mblnVerify = True
            chkSwitch.Visible = mblnVerify
            vsfHead.ColHidden(mlngChoose) = Not mblnVerify
            VsfData.ColHidden(mlngChoose) = Not mblnVerify
            VsfData.Cell(flexcpChecked, VsfData.FixedRows, mlngChoose, VsfData.Rows - 1, mlngChoose) = flexTSUnchecked
            Call WriteColor
            RaiseEvent AfterDataChanged(mblnChange Or mblnVerify)
            Exit Function
        Else
            '提取待审签的数据
            '汇总数据也要签名,因此去掉条件: And B.汇总类别=0
            gstrSQL = " Select /*+ RULE */ distinct B.发生时间 " & vbNewLine & _
                      " From 病人护理明细 A,病人护理数据 B,病人护理文件 C,(SELECT COLUMN_VALUE FROM TABLE(CAST(F_NUM2LIST([2]) AS ZLTOOLS.T_NUMLIST))) G" & vbNewLine & _
                      " Where A.记录ID=B.ID And B.ID=G.COLUMN_VALUE And B.文件ID=C.ID And MOD(A.记录类型,10)=5 AND A.终止版本 Is NULL And C.ID=[1] "
            Call SQLDIY(gstrSQL)
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查是否存在未签名的数据", mlng文件ID, mstrVerify)
        End If
    Else
        '仅对本人修改的数据进行签名(提取未签名数据-已签名数据)
        '汇总数据也要签名,因此去掉条件: And B.汇总类别=0
        gstrSQL = "" & _
                "SELECT  DISTINCT B.发生时间" & vbNewLine & _
                "FROM 病人护理明细 A,病人护理数据 B" & vbNewLine & _
                "WHERE A.记录ID=B.ID And A.数据来源=0 AND A.终止版本 IS NULL AND A.记录类型 =1 AND instr(NVL(B.签名人,'QMR'),'/',1)=0 AND A.记录人=[2] AND B.文件ID=[1]" & vbNewLine & _
                "MINUS" & vbNewLine & _
                "SELECT DISTINCT B.发生时间" & vbNewLine & _
                "FROM 病人护理明细 A,病人护理数据 B" & vbNewLine & _
                "WHERE A.记录ID=B.ID And A.数据来源=0 AND A.终止版本 IS NULL AND A.记录类型 =5 AND B.文件ID=[1]"
        Call SQLDIY(gstrSQL)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "检查是否存在未签名的数据", mlng文件ID, gstrUserName)
        If rsTemp.RecordCount = 0 Then
            RaiseEvent AfterRowColChange("没有找到需要签名的数据（只能对自己登记或修改的数据进行签名）！", True, mblnSign, mblnArchive)
            Exit Function
        End If
    End If
    
    '准备签名
    strSignTime = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    str行错误 = ""
    With rsTemp
        Do While Not .EOF
            blnSign = SignName(Format(!发生时间, "yyyy-MM-dd HH:mm:ss"), strSignTime, bln审签, str状态, str行错误)
            If Not blnSign Then Exit Do
            If Not blnRefresh Then blnRefresh = blnSign
'            If str行错误 <> "" Then
'                str错误 = str错误 & vbCrLf & "发生时间=[" & Format(!发生时间, "yyyy-MM-dd HH:mm:ss") & "]" & str行错误
'            End If
            .MoveNext
        Loop
    End With
    
    If blnRefresh And Not mblnVerify Then Call ShowMe(mFrmParent, mlng文件ID, mlng病人ID, mlng主页ID, mlng病区ID, mint婴儿, mstrPrivs, mblnEditable, IIf(mblnBlowup = True, 1, 0))
    If str错误 <> "" Then MsgBox "签名时发生以下错误：" & str错误, vbInformation, gstrSysName
    SignMe = blnRefresh
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub UnSignMe(Optional ByVal bln审签 As Boolean = False)
    Dim intPos As Integer
    Dim lngStart As Long                '启始行
    Dim lngRecord As Long
    Dim blnOK As Boolean
    Dim strSignTime As String           '签名时间
    Dim blnClear As Boolean             '取消签名时是否清除该版本的数据回退到上次签名后的状态
    Dim blnTrans As Boolean
    Dim strSignName As String
    Dim clsSign As Object
    Dim rsTemp As New ADODB.Recordset
    Dim rsSign As New ADODB.Recordset
    Dim lngRecordCount As Long
    Dim strSQLTime() As String
    ReDim Preserve strSQLTime(1 To 1)
    On Error GoTo errHand
    '首先最后一次是本人的签名，根据当前选择数据的签名时间，批量取消签名
    
    If mlng病人ID = 0 Then Exit Sub
    
    '必要性检查
    '当前记录是新记录则退出
    If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "" Then Exit Sub
    lngStart = GetStartRow(VsfData.ROW)
    lngRecord = Val(VsfData.TextMatrix(lngStart, mlngRecord))
    If lngRecord = 0 Then
        RaiseEvent AfterRowColChange("新增记录不存在取消签名！", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    '当前记录未签名则退出
    If VsfData.TextMatrix(lngStart, mlngSigner) = "" Then
        RaiseEvent AfterRowColChange("当前记录还未签名！", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    '审签：当前记录未审签则退出；平签：当前记录已审签则退出
    intPos = InStr(1, VsfData.TextMatrix(lngStart, mlngSigner), "/")
    If bln审签 Then
        If intPos = 0 Then
            RaiseEvent AfterRowColChange("当前记录未审签，无法执行取消审签操作！", True, mblnSign, mblnArchive)
            Exit Sub
        End If
    Else
        If intPos <> 0 Then
            RaiseEvent AfterRowColChange("当前记录已审签，请取消审签后再操作！", True, mblnSign, mblnArchive)
            Exit Sub
        End If
    End If
    '取消签名时，可以回退自己的签名，可以回退代签的数据
    gstrSQL = "" & _
              " SELECT  A.记录人,A.记录时间,A.项目名称 AS 签名时间,NVL(A.开始版本,1) 开始版本,B.签名人" & vbNewLine & _
              " FROM 病人护理明细 A,病人护理数据 B" & vbNewLine & _
              " WHERE A.记录ID=B.ID AND B.文件ID=[1] AND A.记录ID=[2] AND A.记录类型=" & IIf(bln审签, 15, 5) & vbNewLine & _
              " ORDER BY A.项目名称 DESC"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "当前记录的最后签名人不是本人则退出", mlng文件ID, lngRecord)
    
    If rsTemp.RecordCount = 0 Then
        RaiseEvent AfterRowColChange("无法找到已" & IIf(bln审签, "审签", "签名") & "的数据，可能是数据变化未刷新导致，请刷新数据后再试！", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    
    If bln审签 = False And InStr(1, NVL(rsTemp!签名人), "/") <> 0 Then
        RaiseEvent AfterRowColChange("当前记录已审签，可能是数据变化未刷新导致，请刷新数据、取消审签后再操作！", True, mblnSign, mblnArchive)
        Exit Sub
    End If
    
    If NVL(rsTemp!记录人) <> gstrUserName Then
        strSignName = NVL(rsTemp!记录人)
        '如果不是本人签名，检查是否是代签
        gstrSQL = "" & _
              " SELECT A.记录人" & vbNewLine & _
              " FROM 病人护理明细 A,病人护理数据 B" & vbNewLine & _
              " WHERE A.记录ID=B.ID AND A.记录类型=1 AND B.文件ID=[1] AND A.记录ID=[2] And nvl(A.开始版本,1)=[3]"
        Call SQLDIY(gstrSQL)
        Set rsSign = zlDatabase.OpenSQLRecord(gstrSQL, "当前记录的最后签名人不是本人则退出", mlng文件ID, lngRecord, Val(NVL(rsTemp!开始版本, 1)))
        lngRecordCount = rsSign.RecordCount
        rsSign.Filter = "记录人='" & gstrUserName & "'"
        If rsSign.RecordCount = 0 And lngRecordCount > 0 Then
            RaiseEvent AfterRowColChange("您不是最后签名人[" & strSignName & "]或代签人[" & gstrUserName & "]，不能执行本操作！", True, mblnSign, mblnArchive)
            Exit Sub
        End If
    Else
        strSignName = gstrUserName
    End If
    
    '提取所有数据准备取消签名或审签(记录时间<>""说明是新版签名)
    '汇总数据也要签名,因此去掉条件: And B.汇总类别=0
    If Not IsNull(rsTemp!记录时间) Then
        gstrSQL = "" & _
                  " SELECT  A.项目ID AS 证书ID,B.发生时间,B.签名人" & vbNewLine & _
                  " FROM 病人护理明细 A,病人护理数据 B" & vbNewLine & _
                  " WHERE A.记录ID=B.ID AND B.文件ID=[1] And A.记录人=[2] And A.记录时间=[3] " & _
                  " AND A.记录类型=" & IIf(bln审签, 15, 5)
        Call SQLDIY(gstrSQL)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取所有数据准备取消签名或审签", mlng文件ID, strSignName, CDate(rsTemp!记录时间))
    Else
        gstrSQL = "" & _
                  " SELECT  A.项目ID AS 证书ID,B.发生时间,B.签名人" & vbNewLine & _
                  " FROM 病人护理明细 A,病人护理数据 B" & vbNewLine & _
                  " WHERE A.记录ID=B.ID AND B.文件ID=[1] And A.记录人=[2] And A.项目名称=[3] " & _
                  " AND A.记录类型=" & IIf(bln审签, 15, 5) & vbNewLine & _
                  " ORDER BY A.项目名称 DESC"
        Call SQLDIY(gstrSQL)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取所有数据准备取消签名或审签", mlng文件ID, strSignName, CStr(rsTemp!签名时间))
    End If
    
    '签名后不允许修改，如需修改必须回退签名，因此取消普签时不存在提示是否回退数据的问题，审签自动回退，所以取消提示 询问是否需要清除数据
'    If Not bln审签 Then
'        blnClear = (MsgBox("取消签名时是否该版本的数据回退到上次签名后的状态？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes)
'    End If
    blnClear = True
    Do While Not rsTemp.EOF
        If (bln审签 = False And InStr(1, NVL(rsTemp!签名人), "/") = 0) Or bln审签 = True Then
            If NVL(rsTemp!证书ID, 0) > 0 Then
                '数字签名验证，只验证一次
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
                        RaiseEvent AfterRowColChange("电子签名部件未能正确安，回退操作不能继续！", True, mblnSign, mblnArchive)
                        Exit Sub
                    End If
                End If
            End If
            
            '取消签名
            gstrSQL = "ZL_病人护理数据_UNSIGNNAME("
            gstrSQL = gstrSQL & mlng文件ID & ","
            gstrSQL = gstrSQL & "To_Date('" & Format(rsTemp!发生时间, "yyyy-MM-dd HH:mm:ss") & "','yyyy-MM-dd hh24:mi:ss')," & IIf(blnClear, "1", "0") & ")"
            strSQLTime(ReDimArray(strSQLTime)) = gstrSQL
        End If
        rsTemp.MoveNext
    Loop
    
    gcnOracle.BeginTrans
    blnTrans = True
    
    For intPos = 1 To UBound(strSQLTime)
        If strSQLTime(intPos) <> "" Then
            Call zlDatabase.ExecuteProcedure(strSQLTime(intPos), "执行取消签名")
        End If
    Next intPos
    
    gcnOracle.CommitTrans
    blnTrans = False
    
    '刷新数据
    Call ShowMe(mFrmParent, mlng文件ID, mlng病人ID, mlng主页ID, mlng病区ID, mint婴儿, mstrPrivs, mblnEditable, IIf(mblnBlowup = True, 1, 0))
    Exit Sub
errHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function SignName(ByVal strStart As String, ByVal strSignTime As String, ByVal bln审签 As Boolean, _
    str状态 As String, Optional str错误 As String) As Boolean
    '******************************************************************************************************************
    '功能:
    '
    '
    '******************************************************************************************************************
    Dim oSign As cPartogramSign
    Dim strSource As String             '审签源数据串
    Dim lngLoop As Long
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    '初始处理
    '------------------------------------------------------------------------------------------------------------------
    strSource = ""
    
    '获取要签名的内容(汇总数据也要签名,因此去掉条件: And B.汇总类别=0)
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = " Select  a.记录类型,a.项目分组,a.项目序号,a.项目名称,a.项目类型,a.记录内容,a.项目单位,a.记录标记,a.体温部位,a.记录组号,a.复试合格,a.未记说明,a.记录人,a.记录时间  " & _
              " From 病人护理明细 a,病人护理数据 b,病人护理文件 c " & _
              " Where a.记录id=b.ID And b.文件ID=c.ID AND MOD(A.记录类型,10)<>5 And a.终止版本 Is Null And C.ID=[1] And b.发生时间=[2]" & _
              " Order by a.项目序号"
    Call SQLDIY(gstrSQL)
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "获取要签名的内容", mlng文件ID, CDate(strStart))
    If rs.BOF = False Then
        Do While Not rs.EOF
            For lngLoop = 0 To rs.Fields.Count - 1
                strSource = strSource & CStr(zlCommFun.NVL(rs.Fields(lngLoop).Value, ""))
            Next
            rs.MoveNext
        Loop
    End If
    'Debug.Print "开始签名：" & Now & vbCrLf & strSource
    If strSource = "" Then
        RaiseEvent AfterRowColChange("当前没有需要签名的信息！", True, mblnSign, mblnArchive)
        Exit Function
    End If
    
    '------------------------------------------------------------------------------------------------------------------
    On Error Resume Next
    Err = 0
    '76223:刘鹏飞,2014-08-05,电子签名增加时间戳信息
    Set oSign = frmPartogramSign.ShowMe(Me, mstrPrivs, mlng文件ID, mlng病区ID, mintVerify_Last, strSource, bln审签, str状态, str错误)
    On Error GoTo errHand
    
    If Not oSign Is Nothing Then
        gstrSQL = "ZL_病人护理数据_SIGNNAME("
        gstrSQL = gstrSQL & mlng文件ID & ","
        gstrSQL = gstrSQL & "To_Date('" & strStart & "','yyyy-mm-dd hh24:mi:ss')," & IIf(bln审签, 1, 0) & ","
        gstrSQL = gstrSQL & "'" & oSign.姓名 & "',"
        gstrSQL = gstrSQL & "'" & oSign.签名信息 & "'," & oSign.签名级别 & ","
        gstrSQL = gstrSQL & oSign.证书ID & ","
        gstrSQL = gstrSQL & oSign.签名方式 & ",'" & oSign.时间戳 & "',0,'" & oSign.时间戳信息 & "',"
        gstrSQL = gstrSQL & "To_Date('" & strSignTime & "','yyyy-mm-dd hh24:mi:ss'))"
        
        'Debug.Print gstrSQL
        Call zlDatabase.ExecuteProcedure(gstrSQL, "执行签名")
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
    Call ShowMe(mFrmParent, mlng文件ID, mlng病人ID, mlng主页ID, mlng病区ID, mint婴儿, mstrPrivs, mblnEditable, IIf(mblnBlowup = True, 1, 0))
End Function

Public Function SaveME() As Boolean
    If Not CheckData Then Exit Function
    If Not SaveData Then Exit Function
    
    mblnShow = False
    Call InitCons
    SaveME = True
    
    Call ShowMe(mFrmParent, mlng文件ID, mlng病人ID, mlng主页ID, mlng病区ID, mint婴儿, mstrPrivs, mblnEditable, IIf(mblnBlowup = True, 1, 0))
End Function

Public Function ShowMe(ByVal frmParent As Form, ByVal lngFileID As Long, ByVal lngPatiID As Long, ByVal lngPageId As Long, _
    ByVal lngDeptID As Long, ByVal intBaby As Integer, Optional ByVal strPrivs As String, Optional ByVal blnEditable As Boolean = True, _
    Optional ByVal bytSize As Byte = 0) As Boolean
    '******************************************************************************************************************
    '功能： 显示护理记录文件内容
    '参数： frmParent           上级窗体对象
    '       lngPatiID           病人id
    '       lngPageID           主页id
    '       lngDeptID           要显示护理记录的科室
    '       intBaby             婴儿标志
    '       blnEditable         如果为假,说明是做为查询子窗体在使用,取消与编辑相关的功能
    '       bytSize             0-小字体,1大字体
    '返回： 无
    '******************************************************************************************************************
    Dim lngRow As Long
    Dim lngCount As Long, i As Long
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    Err = 0
    
    mblnInit = False
    mblnChange = False
    mlng文件ID = lngFileID
    mlng病人ID = lngPatiID
    mlng主页ID = lngPageId
    mlng病区ID = lngDeptID
    mint婴儿 = intBaby
    mlngFileIndex = frmParent.FileNumIndex
    mstrPrivs = strPrivs
    mblnBlowup = (bytSize = 1)
    Set mFrmParent = frmParent
    
    mintPreDays = Val(zlDatabase.GetPara("超期录入护理数据天数", glngSys, 1255, "1"))
    
    If mrsItems.State = 0 Then
        Call InitMenuBar
        Call InitEnv            '初始化环境
        cbsThis.ActiveMenuBar.Visible = False
        cbsThis.RecalcLayout
    End If
    Call ReSetFontSize
    
    '提取产程文件份数
    lngCount = GetFileCount(mlng文件ID, mlng病人ID, mlng主页ID)
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
    '功能:设置字体大小
    '入参:bytSize：0-小(缺省)，1-大
    '编制:刘鹏飞
    '日期:2012-06-20 15:15:00
    '问题:50807
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim objCtrl As Control
    Dim CtlFont As StdFont
    Dim bytFontSize As Byte
    bytFontSize = BlowUp(9)
    
    UserControl.FontSize = bytFontSize
    UserControl.FontName = "宋体"
    
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
    '数据保存时检查是否录入的日期时间
    
    lngRows = VsfData.Rows - 1
    lngCols = VsfData.Cols - 1
    For lngRow = VsfData.FixedRows To lngRows
        mrsCellMap.Filter = "页号=" & mint页码 & " And 行号=" & lngRow & " And 列号>" & mlngTime
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
                    RaiseEvent AfterRowColChange("请补充日期时间！", True, mblnSign, mblnArchive)
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
    
    '保存时检查,1、婴儿出生时必须存在对应的宫缩和胎头数据,2.如果存在婴儿文件数>1，则不能取消上一文件婴儿出生标志
    '说明：必须放在DataMap_Save之后使用
    
    On Error GoTo errHand
    
    '--1、婴儿出生时必须存在对应的宫缩和胎头数据
    mrsDataMap.Filter = ""
    mrsDataMap.Filter = "删除=0 And 产程日期<>''"
    mrsDataMap.Sort = "产程日期"
    'Call OutputRsData(mrsDataMap)
    Do While Not mrsDataMap.EOF
        If blnExit = False Then blnExit = (NVL(mrsDataMap.Fields(cControlFields + mlngSpread - VsfData.FixedCols).Value) <> "") And (NVL(mrsDataMap.Fields(cControlFields + mlngJust - VsfData.FixedCols).Value) <> "")
        If Mid(NVL(mrsDataMap.Fields(cControlFields + mlngProduce - VsfData.FixedCols).Value, 0), 1, 1) = "√" Then
            blnProduce = True
            strDatetime = Format(mrsDataMap!产程日期, "YYYY-MM-DD HH:mm:ss")
            lngPage = mrsDataMap!页号
            lngRow = mrsDataMap!行号
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
            lngRow = IIf(lngPage = mint页码, lngRow, VsfData.FixedRows)
            VsfData.ROW = lngRow
            VsfData.COL = lngCol
            If Not VsfData.RowIsVisible(lngRow) Then VsfData.TopRow = lngRow
            If Not VsfData.ColIsVisible(mlngSpread) Then VsfData.LeftCol = mlngSpread
            RaiseEvent AfterRowColChange("婴儿出生必须同时存在宫口扩大和先露下降数据，请检查！", True, mblnSign, mblnArchive)
            CheckProduce = False
            Exit Function
        End If
    End If
    
    '2.如果存在婴儿文件数>1，则不能取消上一文件婴儿出生标志
    If cboBaby.ListCount > 1 And cboBaby.ListIndex < cboBaby.ListCount - 1 Then
        If blnProduce = False Then '表示不存在婴儿出生标记
            RaiseEvent AfterRowColChange("当前婴儿文件小于总文件数时，必须存在相应的生产标志，请检查！", True, mblnSign, mblnArchive)
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
    '检查数据
    
    '如果修改了数据而日期时间不全则提示（数据合法性在录入时已经检查）
    If Not DataMap_Save Then Exit Function
    If Not CheckProduce Then Exit Function
    
    '如果是审签模式,则检查所选数据是否存在不能审签的情况
    If mblnVerify Then
        mstrVerify = ""
        mintVerify_Last = 未定义
        '审签不允许新增数据
        mrsDataMap.Filter = "页号=" & mint页码
        Do While Not mrsDataMap.EOF
            If NVL(mrsDataMap!选择, 0) = flexTSChecked Then
                mstrVerify = mstrVerify & "," & mrsDataMap!记录ID
                
                If IsNull(mrsDataMap!签名级别) Then
                    intLevel = NVL(mrsDataMap!签名级别, 未定义)
                Else
                    intLevel = Val(mrsDataMap!签名级别) + 1
                End If
                If mintVerify < intLevel Then mintVerify_Last = intLevel
            End If
            mrsDataMap.MoveNext
        Loop
        mrsDataMap.Filter = 0
        
        If mstrVerify = "" Then
            RaiseEvent AfterRowColChange("至少要选择一条数据才能完成审签操作！", True, mblnSign, mblnArchive)
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
    
    intAllow = IIf(InStr(mstrPrivs, "他人护理记录") > 0, 1, 0)
    strCurrDate = Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm:ss")
    
    With mrsCellMap
        '将有效数据过滤出来:记录ID>0的历史数据+新增的有效数据
        .Filter = "记录ID>0 or (记录ID=0 And 删除=0)"
        If .RecordCount <> 0 Then .MoveFirst
        Do While Not .EOF
            If (InStr(1, "," & mstrVerify & ",", "," & NVL(!记录ID, 0) & ",") <> 0 And mblnVerify = True) Or mblnVerify = False Then
                If intRow <> !行号 Then
                    '赋初值
                    intRow = !行号
                    strDate = ""
                    strDatetime = ""
                    lngRecord = NVL(!记录ID, 0)
                End If
                
                If !列号 = mlngDate Then
                    strDate = NVL(!数据)
                    If strDate <> "" Then
                        If mblnDateAd Then
                            strDate = Mid(strCurrDate, 1, 5) & ToStandDate(strDate)
                        Else
                            strDate = Format(strDate, "yyyy-MM-dd")
                        End If
                    End If
                ElseIf !列号 = mlngTime Then
                    strTime = NVL(!数据)
                    If strDate = "" Then strDate = Mid(strCurrDate, 1, 10)
                    strDatetime = strDate & " " & strTime & ":00"
                    If mblnDateAd Then
                        strDatetime = GetDateAdCurrDate(strDatetime)
                    End If
                    
                    If lngRecord <> 0 Then
                        '更新发生时间
                        gstrSQL = "ZL_产程图数据_发生时间(" & lngRecord & ",to_date('" & strDatetime & "','yyyy-MM-dd hh24:mi:ss'))"
                        strSQL(ReDimArray(strSQL)) = gstrSQL
                    End If
                Else
                    If !列号 > mlngTime Then
                        '取指定单元格的数据
                        strCellData = NVL(!数据)
                        strPart = NVL(!部位)
                        strReturn = ShowInput(!列号, strCellData, True)
                        'strOrders格式：项目序号,项目序号...
                        'strValues格式：值'值'值...
                        arrOrder = Split(Split(strReturn, "||")(0), ",")
                        arrValue = Split(Split(strReturn, "||")(1) & "'", "'")
                        arrPart = Split(strPart & "/////", "/")
                        
                        intMax = UBound(arrOrder)
                        For intPos = 0 To intMax
        '                    文件ID_IN IN 病人护理数据.文件ID%TYPE,
        '                    发生时间_IN IN 病人护理数据.发生时间%TYPE,
        '                    记录类型_IN IN 病人护理明细.记录类型%TYPE,          --护理项目=1，上标说明=2，手术日标记=4，签名记录=5，下标说明=6，入出量汇总=9
        '                    项目序号_IN IN 病人护理明细.项目序号%TYPE,          --护理项目的序号，非护理项目固定为0
        '                    记录内容_IN IN 病人护理明细.记录内容%TYPE := NULL,  --记录内容，如果内容为空，即清除以前的内容；37或38/37
        '                    体温部位_IN IN 病人护理明细.体温部位%TYPE := NULL,
        '                    记录组号_IN IN 病人护理数据.汇总类别%TYPE := 1,             --标记那个婴儿的数据
        '                    他人记录_IN IN NUMBER := 1,
                            gstrSQL = "ZL_产程图数据_UPDATE(" & mlng文件ID & ",to_date('" & strDatetime & "','yyyy-MM-dd hh24:mi:ss'),1," & _
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
    
    '循环执行SQL保存数据
    intMax = UBound(strSQL)
    gcnOracle.BeginTrans
    blnTrans = True
    If intMax > 0 Then
        For intPos = 1 To intMax
            If strSQL(intPos) <> "" Then
                'Debug.Print strSQL(intPos)
                Call zlDatabase.ExecuteProcedure(strSQL(intPos), "保存产程图数据")
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
        If MsgBox("当前病人的数据还未保存，点“是”手工进行保存，点“否”将放弃本次修改！", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
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
    Dim strLockItem As String                   '同步过来的数据,不允许修改或删除
    Dim lngTop As Long, lngHeight As Long
    Dim intMax As Integer                       '同步过来的数据占用的最大行数
    Dim intNULL As Integer, lngStartRow As Long
    Dim lngRow As Long, lngCol As Long, lngRows As Long, lngCols As Long
    Dim strKey As String, strField As String, strValue As String
    Dim blnShow As Boolean
    On Error GoTo err_exit
    
    Select Case Control.ID
    '粘贴,清除时需要同步mrsCellMap数据
    Case conMenu_Edit_FileMan
        '文件添加
        Call LoadBabyNum
    Case conMenu_Edit_Copy
        '复制指定数据行的数据
        If VsfData.ROW < VsfData.FixedRows Then Exit Sub
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then Exit Sub
        lngRow = GetStartRow(VsfData.ROW)
        
        '复制记录集
        Set mrsCopyMap = New ADODB.Recordset
        Set mrsCopyMap = CopyNewRec(mrsDataMap, False)
        
        '得到指定数据行的起始行,结束行
        lngCols = VsfData.Cols - 1
        lngRows = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
        lngRows = lngRow + lngRows - 1
        For lngRow = lngRow To lngRows
            mrsCopyMap.AddNew
            mrsCopyMap!页号 = mint页码
            mrsCopyMap!行号 = lngRow
            For lngCol = 0 To lngCols - VsfData.FixedCols    '多了一个固定列
                mrsCopyMap.Fields(cControlFields + lngCol).Value = IIf(VsfData.TextMatrix(lngRow, lngCol + VsfData.FixedCols) = "", Null, VsfData.TextMatrix(lngRow, lngCol + VsfData.FixedCols))
            Next
            mrsCopyMap.Update
        Next
    Case conMenu_Edit_PASTE
        '粘贴时，将目标行整体覆盖，同步过来的数据列，活动列除外
        '活动项目可能不同页面项目不同，部位不同，所以不考虑活动项目
        '同步行所占用的行数不变，如不够再添加空白行，再行粘贴
        If VsfData.ROW < VsfData.FixedRows Then Exit Sub
        If mrsCopyMap.RecordCount = 0 Then Exit Sub
        
        '跨页数据行不允许对整行进行粘贴,删除,只能编辑除活动项目外的列
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then
            lngStartRow = VsfData.ROW
        Else
            lngStartRow = GetStartRow(VsfData.ROW)
        End If
        
        '检查目标数据行是否存在同步过来的数据,如果有则跳过同步的记录
        strLockItem = GetSynItems(2, intMax)        '1.返回项目序号;2.返回列号
        
        '得到目标数据行的起始行,结束行
        strField = "ID|页号|行号|列号|记录ID|数据|删除"
        lngCols = VsfData.Cols - 1
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then
            lngRow = VsfData.ROW
            lngStartRow = lngRow
            If mlngDate > -1 Then strDate = VsfData.TextMatrix(lngRow, mlngDate)
            strTime = VsfData.TextMatrix(lngRow, mlngTime)
        Else
            '删除多余的数据行,仅留一行
            lngRow = GetStartRow(VsfData.ROW)
            lngStartRow = lngRow
            strDate = VsfData.TextMatrix(lngRow, mlngDate)
            strTime = VsfData.TextMatrix(lngRow, mlngTime)
            lngRows = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0)) - 1
            For intNULL = 1 To lngRows
                VsfData.RemoveItem lngRow + 1
            Next
        End If
        
        '往下搜索空行,如果有其它数据行则计算需增加的行数
        intNULL = mrsCopyMap.RecordCount - 1
        For lngRow = 1 To mrsCopyMap.RecordCount - 1
            '保证当前输入的内容在一页中显示全
            If lngRow + VsfData.ROW > VsfData.Rows - 1 Then Exit For
            
            If Val(VsfData.TextMatrix(lngRow + VsfData.ROW, mlngRecord)) = 0 And VsfData.TextMatrix(lngRow + VsfData.ROW, mlngRowCount) = "" Then
                intNULL = intNULL - 1
            Else
                Exit For
            End If
        Next
        '先增加空行
        If intNULL > 0 Then
            VsfData.Rows = VsfData.Rows + intNULL
            '从当前行记录的空白行开始，每行的位置+所增加的空白行数
            For lngRow = 1 To intNULL
                VsfData.RowPosition(VsfData.Rows - 1) = lngStartRow + 1
            Next
        End If
        
        '还原日期，时间，强制不允许修改
        VsfData.TextMatrix(lngStartRow, mlngDate) = strDate
        VsfData.TextMatrix(lngStartRow, mlngTime) = strTime
        '记录用户修改过的单元格
        If mlngDate <> -1 Then
            strKey = mint页码 & "," & lngStartRow & "," & mlngDate
            strValue = strKey & "|" & mint页码 & "|" & lngStartRow & "|" & mlngDate & "|" & _
                Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & VsfData.TextMatrix(lngStartRow, mlngDate) & "|0"
            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        End If
        '2\时间
        strKey = mint页码 & "," & lngStartRow & "," & mlngTime
        strValue = strKey & "|" & mint页码 & "|" & lngStartRow & "|" & mlngTime & "|" & _
            Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & VsfData.TextMatrix(lngStartRow, mlngTime) & "|0"
        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        '向表格填充数据
        With mrsCopyMap
            If .RecordCount <> 0 Then .MoveFirst
            Do While Not .EOF
                For lngCol = 0 To lngCols - VsfData.FixedCols
                    Select Case lngCol + VsfData.FixedCols
                    Case 1, mlngDate, mlngTime, mlngOperator, mlngSigner, mlngSignTime, mlngRecord
                    Case Else
                        If InStr(1, "," & strLockItem & ",", "," & lngCol - (cHideCols - 1) & ",") = 0 And InStr(1, "," & mstrCOLNothing & ",", "," & lngCol - (cHideCols - 1) & ",") = 0 Then
                            VsfData.TextMatrix(lngStartRow + .AbsolutePosition - 1, lngCol + VsfData.FixedCols) = NVL(.Fields(cControlFields + lngCol).Value)
                            
                            '修改标志
                            If .AbsolutePosition = 1 Then
                                strKey = mint页码 & "," & lngStartRow & "," & lngCol + VsfData.FixedCols
                                strValue = strKey & "|" & mint页码 & "|" & lngStartRow & "|" & lngCol + VsfData.FixedCols & "|" & _
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
            RaiseEvent AfterRowColChange("已签名的数据不允许删除！", True, mblnSign, mblnArchive)
            Exit Sub
        End If
        
        '跨页数据行不允许对整行进行粘贴,删除,只能编辑除活动项目外的列
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then
            lngStartRow = VsfData.ROW
        Else
            lngStartRow = GetStartRow(VsfData.ROW)
        End If
        
        lngRowCount = Split(VsfData.TextMatrix(lngStartRow, mlngRowCount), "|")(0)
        '检查目标数据行是否存在同步过来的数据,如果有则跳过同步的记录
        strLockItem = GetSynItems(2, intMax)        '1.返回项目序号;2.返回列号
        
        '准备删除
        strField = "ID|页号|行号|列号|记录ID|数据|汇总|删除"
        If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then
            lngRow = VsfData.ROW
        Else
            lngRow = GetStartRow(VsfData.ROW)
            lngStartRow = lngRow
            If VsfData.TextMatrix(lngStartRow, mlngSigner) <> "" Then
                RaiseEvent AfterRowColChange("已签名的数据不允许删除！", True, mblnSign, mblnArchive)
                Exit Sub
            End If
            
            '删除所有数据行
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
        '记录用户修改过的单元格
        strDate = VsfData.TextMatrix(lngStartRow, mlngDate)
        strTime = VsfData.TextMatrix(lngStartRow, mlngTime)
            
        If mblnDateAd Then
            If InStr(1, strDate, "/") <> 0 Then
                strDate = Mid(zlDatabase.Currentdate, 1, 5) & Split(strDate, "/")(1) & "-" & Split(strDate, "/")(0)
            End If
            strDate = Mid(strDate, 9, 2) & "/" & Mid(strDate, 6, 2)
        End If
        
        strField = "ID|页号|行号|列号|记录ID|数据|汇总|删除"
        strKey = mint页码 & "," & lngStartRow & "," & mlngDate
        strValue = strKey & "|" & mint页码 & "|" & lngStartRow & "|" & mlngDate & "|" & _
            Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & strDate & "|0|1"
        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        '2\时间
        strKey = mint页码 & "," & lngStartRow & "," & mlngTime
        strValue = strKey & "|" & mint页码 & "|" & lngStartRow & "|" & mlngTime & "|" & _
            Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "|" & strTime & "|0|1"
        Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        
        strField = "ID|页号|行号|列号|记录ID|数据|汇总|删除"
        
        '删除启始行中非同步的数据
        If strLockItem = "" Then
            VsfData.RowHidden(lngRow) = True
            '填写修改标志
            For lngCol = mlngTime + 1 To mlngNoEditor - 1
                strKey = mint页码 & "," & lngStartRow & "," & lngCol
                strValue = strKey & "|" & mint页码 & "|" & lngStartRow & "|" & lngCol & "|" & _
                    Val(VsfData.TextMatrix(lngStartRow, mlngRecord)) & "||0|1"
                Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
            Next
        Else
            '填写修改标志(存在同步数据,日期与时间列不允许清除)``
            For lngCol = mlngTime + 1 To mlngNoEditor - 1
                If InStr(1, "," & strLockItem & ",", "," & lngCol - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 And lngCol <> mlngDate And lngCol <> mlngTime Then
                    VsfData.TextMatrix(lngStartRow, lngCol) = ""
                    
                    strKey = mint页码 & "," & lngStartRow & "," & lngCol
                    strValue = strKey & "|" & mint页码 & "|" & lngStartRow & "|" & lngCol & "|" & _
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
        
        '检查当前录入控件
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
    Case conMenu_Edit_Append '产程要素
        RaiseEvent AfterPartogramInfo(mlng文件ID, mlngFileIndex, mlng格式ID, mrsPartogram)
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
    Case conMenu_Edit_FileMan '添加婴儿
        Control.Enabled = Not mblnArchive And mblnEditable And Not mblnVerify And Not mblnChange
        If picBaby.Visible = True Then
            picBaby.Visible = Control.Enabled
        End If
    Case conMenu_Edit_Copy
        Control.Enabled = Not mblnShow And Not mblnArchive And Not mblnVerify And mblnEditable
    Case conMenu_Edit_PASTE
        Control.Enabled = False
        If mrsCopyMap.State = 0 Then Exit Sub
        '签名数据不允许粘贴
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
    Case conMenu_Edit_Append '产程要素
        Control.Enabled = Not mblnArchive And mblnEditable
    Case conMenu_Edit_Word
        Control.Enabled = mblnEditAssistant And mblnShow And Not mblnArchive And mblnEditable
    End Select
errHand:
End Sub

Private Sub chkSwitch_Click()
    Dim blnSel As Boolean            '是否全部选中
    Dim blnUpdate As Boolean
    Dim intLevel As Integer
    Dim lngRow As Long, lngRows As Long
    Dim strKey As String, strField As String, strValue As String
    '将所有列全部选中或取消选中，并保存更新
    
    If Not mblnInit Then Exit Sub
    lngRows = VsfData.Rows - 1
    strField = "ID|页号|行号|列号|记录ID|数据|删除"
    
    blnSel = chkSwitch.Value
    For lngRow = VsfData.FixedRows To lngRows
        If Not VsfData.RowHidden(lngRow) Then
            If VsfData.TextMatrix(lngRow, mlngRowCount) Like "*|1" Then
                blnUpdate = False
                If blnSel Then
                    '检查,签过名的记录,且当前操作员级别比上次签名级别高
                    If VsfData.TextMatrix(lngRow, mlngSignLevel) = "" Then
                        intLevel = 未定义
                    Else
                        intLevel = Val(VsfData.TextMatrix(lngRow, mlngSignLevel)) + 1
                    End If
                    If mintVerify < intLevel And intLevel <> 未定义 Then
                        blnUpdate = (VsfData.Cell(flexcpChecked, lngRow, mlngChoose) <> flexTSChecked)
                        VsfData.Cell(flexcpChecked, lngRow, mlngChoose) = flexTSChecked
                    End If
                Else
                    blnUpdate = (VsfData.Cell(flexcpChecked, lngRow, mlngChoose) <> flexTSUnchecked)
                    VsfData.Cell(flexcpChecked, lngRow, mlngChoose) = flexTSUnchecked
                End If
                
                If blnUpdate Then
                    '保存修改记录以便同步
                    strKey = mint页码 & "," & lngRow & "," & mlngChoose
                    strValue = strKey & "|" & mint页码 & "|" & lngRow & "|" & mlngChoose & "|" & _
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
    '添加规则为上一个婴儿出生才能添加下一个婴儿
    strSQL = " SELECT 1" & vbNewLine & _
            " FROM 病人护理文件 A, 病人护理数据 B, 病人护理明细 C,护理记录项目 D" & vbNewLine & _
            " WHERE A.ID = B.文件ID AND B.ID = C.记录ID AND A.ID = [1] AND A.病人ID = [2] AND A.主页ID = [3] AND B.汇总类别 = [4]" & vbNewLine & _
            " AND substr(nvl(C.记录内容,''),1,1)='√' AND C.项目序号=D.项目序号 AND D.项目名称='生产' AND NVL(D.保留项目,0)=1"
    Call SQLDIY(strSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "文件删除检查", mlng文件ID, mlng病人ID, mlng主页ID, lngGroupID)
    If rsTemp.RecordCount = 0 Then
        RaiseEvent AfterRowColChange("添加规则为上一婴儿文件的婴儿已经出生，才能添加下一个！", True, mblnSign, mblnArchive)
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
    '为了保证删除只能从后往前，此处再次进行判断
    For intRow = vsfBaby.FixedRows To vsfBaby.Rows - 1
        If lngFileOldIndex < Val(vsfBaby.RowData(intRow)) Then
            lngFileOldIndex = Val(vsfBaby.RowData(intRow))
        End If
    Next intRow
    
    If lngFileIndex < lngFileOldIndex Then
       RaiseEvent AfterRowColChange("删除只能从最后一个婴儿开始，请检查！", True, mblnSign, mblnArchive)
       Exit Sub
    End If
    
    If MsgBox("此操作将删除与此婴儿相关的所有数据信息，请问您是否要删除？", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    '完成婴儿对应文件数据的删除
'    zl_产程数据数据_DelBaby
    mstrSQL = "zl_产程数据数据_DelBaby("
'    文件ID_IN 病人护理数据.文件ID%TYPE,
    mstrSQL = mstrSQL & mlng文件ID & ","
'    婴儿_IN   病人护理数据.汇总类别%TYPE
    mstrSQL = mstrSQL & lngFileIndex & ")"
    Call zlDatabase.ExecuteProcedure(mstrSQL, "zl_产程数据数据_DelBaby")
    '完成数据刷新
    mblnVerify = False
    mblnChange = False
    lngFileIndex = lngFileIndex - 1
    If lngFileIndex < 1 Then lngFileIndex = 1
    RaiseEvent AfterFileIndex(mlngFileIndex)
    RaiseEvent AfterDataSave(True)
    Call ShowMe(mFrmParent, mlng文件ID, mlng病人ID, mlng主页ID, mlng病区ID, mint婴儿, mstrPrivs, mblnEditable, IIf(mblnBlowup = True, 1, 0))
End Sub

Private Sub cmdWord_Click()
    Dim strInput As String
    '弹出词句选择器
    
    If cmdWord.Tag = -1 Then
        strInput = txtInput.Text
    Else
        strInput = txt(Val(cmdWord.Tag)).Text
    End If
    strInput = frmEditAssistant.ShowMe(Me, mlng病人ID, mlng主页ID, mint婴儿, strInput)
    
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
    '加载签名历史记录
    Dim str发生时间 As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    vsfSignData.Clear
    str发生时间 = VsfData.TextMatrix(VsfData.ROW, mlngActTime)
    gstrSQL = "" & _
        " SELECT A.记录人 AS 签名人,NVL(to_char(A.记录时间,'yyyy-MM-dd hh24:mi:ss'),A.项目名称) AS 签名时间,A.记录内容 AS 签名信息,A.记录标记 AS 签名规则,A.ID,DECODE(A.项目ID,NULL,'有效','未验证') AS 有效性,A.开始版本,NVL(A.项目序号,2) AS 签名规则版本" & vbNewLine & _
        " FROM 病人护理明细 A,病人护理数据 B,病人护理文件 C" & vbNewLine & _
        " WHERE A.记录ID=B.ID And B.文件ID=C.ID AND MOD(A.记录类型,10)=5" & vbNewLine & _
        " AND C.ID=[1] AND B.发生时间=[2] " & vbNewLine & _
        " Order by A.项目名称 Desc"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取签名历史记录", mlng文件ID, CDate(str发生时间))
    
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
    
    rsTemp.Filter = "有效性='未验证'"
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
    '单行验证
    Dim lngLoop As Long
    Dim int版本 As Integer
    Dim strSource As String, str发生时间 As String
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    
    If (Val(vsfSignData.TextMatrix(vsfSignData.ROW, 4)) = 0) Then Exit Sub
    '获取要签名的内容
    '------------------------------------------------------------------------------------------------------------------
    int版本 = vsfSignData.TextMatrix(vsfSignData.ROW, 6)
    str发生时间 = VsfData.TextMatrix(VsfData.ROW, mlngActTime)
    Set rsTemp = GetSignData(str发生时间, int版本)
    Do While Not rsTemp.EOF
        For lngLoop = 0 To rsTemp.Fields.Count - 1
            strSource = strSource & CStr(zlCommFun.NVL(rsTemp.Fields(lngLoop).Value, ""))
        Next
        rsTemp.MoveNext
    Loop
    'Debug.Print "验证签名：" & Now & vbCrLf & strSource
    
    '数字签名
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
        MsgBox "电子签名部件未能正确安装，验证操作不能继续！", vbInformation, gstrSysName
        Exit Sub
    End If
    If gobjESign.VerifySignature(strSource, Val(vsfSignData.TextMatrix(vsfSignData.ROW, 4)), 6) Then
        vsfSignData.TextMatrix(vsfSignData.ROW, 5) = "有效"
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
    '全部验证
    
    lngSel = vsfSignData.ROW
    vsfSignData.Redraw = flexRDNone
    lngRows = vsfSignData.Rows - 1
    For lngRow = 1 To lngRows
        If (vsfSignData.TextMatrix(lngRow, 5) <> "有效") Then
            vsfSignData.ROW = lngRow
            Call cmdSignCur_Click
        End If
    Next
    vsfSignData.ROW = lngSel
    vsfSignData.Redraw = flexRDDirect
End Sub

Private Function ShowSignMarker(Optional ByVal bln外部 As Boolean = False) As Boolean
    '显示历史签名标记
    
    picSign.Visible = False
    picSignCheck.Visible = False
    If Not bln外部 Then
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
    cmdSignCur.Enabled = (vsfSignData.TextMatrix(vsfSignData.ROW, 5) <> "有效")
End Sub

Private Function GetSignData(ByVal str发生时间 As String, ByVal int版本 As Integer) As ADODB.Recordset
    On Error GoTo errHand
    Dim rsTemp As New ADODB.Recordset
    
    If int版本 = 1 Then
        gstrSQL = "" & _
            "Select a.记录类型,a.项目分组,a.项目序号,a.项目名称,a.项目类型,a.记录内容,a.项目单位,a.记录标记,a.体温部位,a.记录组号,a.复试合格,a.未记说明,a.记录人,a.记录时间" & vbNewLine & _
            "  From 病人护理明细 a, 病人护理数据 b,病人护理文件 C" & vbNewLine & _
            " Where c.ID=[1] And b.发生时间 =[2]" & vbNewLine & _
            "   And a.记录id = b.ID and B.文件ID=C.ID and MOD(A.记录类型,10) <>5 and A.开始版本=1" & vbNewLine & _
            " ORDER BY 项目序号"
    Else
        gstrSQL = "" & _
            "Select a.记录类型,a.项目分组,a.项目序号,a.项目名称,a.项目类型,a.记录内容,a.项目单位,a.记录标记,a.体温部位,a.记录组号,a.复试合格,a.未记说明,a.记录人,a.记录时间" & vbNewLine & _
            "  From 病人护理明细 a, 病人护理数据 b,病人护理文件 C" & vbNewLine & _
            " Where c.ID=[1] And b.发生时间 =[2]" & vbNewLine & _
            "   And a.记录id = b.ID and B.文件ID=C.ID and MOD(A.记录类型,10) <>5" & vbNewLine & _
            "   and (A.开始版本=[3] or (A.开始版本 <[3] and A.终止版本 IS NULL) or (A.开始版本<[3] and A.终止版本>[3]))" & vbNewLine & _
            " ORDER BY 项目序号"
    End If
    Set GetSignData = zlDatabase.OpenSQLRecord(gstrSQL, "提取指定版本的数据", mlng文件ID, CDate(str发生时间), int版本)
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Public Sub SignMarker()
    '供外部主程序调用
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
        '第一份文件不能删除
        If .RowData(NewRow) = 1 Then cmdDelBaby.Visible = False: cmdDelBaby.Enabled = False: Exit Sub
        '文件只能从后往前删除
        strSQL = " SELECT 1" & vbNewLine & _
            " FROM 病人护理文件 A, 病人护理数据 B, 病人护理明细 C" & vbNewLine & _
            " WHERE A.ID = B.文件ID AND B.ID = C.记录ID AND A.ID = [1] AND A.病人ID = [2] AND A.主页ID = [3] AND B.汇总类别 > [4]"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "文件删除检查", mlng文件ID, mlng病人ID, mlng主页ID, Val(.RowData(NewRow)))
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
    
    '隐蔽已显示的录入控件
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
    
    '未定义的列不允许录入数据
    mintType = -1
    If mblnInit = False Then Exit Sub
    
    Call ShowSignMarker
    
    If InStr(1, mstrPrivs, "产程图作图") = 0 Then Exit Sub
    
    If VsfData.TextMatrix(VsfData.ROW, mlngRowCount) = "" Then
        lngStart = VsfData.ROW
    Else
        lngStart = GetStartRow(VsfData.ROW)
    End If
    
    If mblnVerify Then  '必须放在mblnShow判断语句的上面
        If VsfData.COL = mlngChoose Then Call vsfdata_KeyDown(vbKeySpace, 0): Exit Sub
        If VsfData.COL = mlngDate Or VsfData.COL = mlngTime Then Exit Sub
        If Val(VsfData.TextMatrix(lngStart, mlngRecord)) = 0 Then Exit Sub
        If VsfData.TextMatrix(lngStart, mlngSigner) = "" Then Exit Sub
        If VsfData.Cell(flexcpChecked, lngStart, mlngChoose) = flexTSUnchecked Then Exit Sub '没有选中的记录不能编辑
    Else
        '审签过的数据只能在审签状态下修改
        If InStr(1, VsfData.TextMatrix(lngStart, mlngSigner), "/") <> 0 Then
            RaiseEvent AfterRowColChange("已审签的数据不允许编辑！", True, mblnSign, mblnArchive)
            Exit Sub
        End If
        
        '--只要是签名的数据就不允许修改
        If VsfData.TextMatrix(lngStart, mlngSigner) <> "" Then
            RaiseEvent AfterRowColChange("已签名的数据不允许再次编辑，请取消签名后再试！", True, mblnSign, mblnArchive)
            Exit Sub
        End If
        
        '默认签名人与保存人相同,不具有修改他人护理记录权限的操作员,不允许修改他人的数据
        strName = VsfData.TextMatrix(lngStart, IIf(mlngOperator = -1, VsfData.Cols - 1, mlngOperator))
        If strName <> "" Then
            If strName <> gstrUserName And _
                InStr(1, mstrPrivs, "他人护理记录") = 0 Then
                RaiseEvent AfterRowColChange("您没有修改他人护理记录数据的权限！原操作员:" & strName, True, mblnSign, mblnArchive)
                Exit Sub
            End If
        End If
    End If
    If mblnArchive Then Exit Sub
    If Not mblnShow Or Not mblnEditable Then Exit Sub

    '未绑定项目列不允许修改
    If InStr(1, "," & mstrCOLNothing & ",", "," & VsfData.COL & ",") <> 0 Then Exit Sub
    
    '同步数据列不允许编辑
    strCols = GetSynItems(2, intMax)
    If strCols <> "" Then
        '存在同步数据的行,日期与时间是不允许修改的
        If VsfData.COL = mlngDate Or VsfData.COL = mlngTime Then Exit Sub
        strCols = "," & strCols & ","
        If InStr(1, strCols, "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") <> 0 Then Exit Sub
    End If
    
    If VsfData.COL <= mlngNoEditor - 1 Then Call ShowInput
    'Debug.Print txtInput.Text
    '让控件获得焦点
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
    
    '选择列,同步数据列直接退出,避免此处清除提示信息
    If NewCol = mlngChoose Then Exit Sub
    strCols = GetSynItems(2, intMax)
    If strCols <> "" Then
        strCols = "," & strCols & ","
        If InStr(1, strCols, "," & NewCol - (cHideCols + VsfData.FixedCols - 1) & ",") <> 0 Then Exit Sub
    End If
    
    '显示当前项目的相关信息
    mrsSelItems.Filter = "列=" & NewCol - (cHideCols + VsfData.FixedCols - 1)
    If mrsSelItems.RecordCount <> 0 Then
        mrsItems.Filter = "项目序号=" & mrsSelItems!项目序号
        If mrsItems.RecordCount <> 0 Then
            If NVL(mrsItems!项目值域) <> "" Then
                If mrsItems!项目类型 = 0 Then
                    strInfo = "有效范围:" & Split(mrsItems!项目值域, ";")(0) & "～" & Split(mrsItems!项目值域, ";")(1)
                Else
                    strInfo = "有效范围:" & mrsItems!项目值域
                End If
            Else
                strInfo = ""
            End If
        End If
    End If
    mrsSelItems.Filter = 0
    mrsItems.Filter = 0
    
    '检查是否已签名
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
        '只勾选起始行
        lngStart = GetStartRow(VsfData.ROW)
        If VsfData.TextMatrix(lngStart, mlngTime) = "" Then Exit Sub
        
        '审签时,当前记录已签名,且操作员的签名级别比上次签名级别高才允许
        If VsfData.TextMatrix(lngStart, mlngSignLevel) = "" Then
            RaiseEvent AfterRowColChange("该数据还未签名，不能进行审签！", True, mblnSign, mblnArchive)
            Exit Sub
        Else
            intLevel = Val(VsfData.TextMatrix(lngStart, mlngSignLevel)) + 1
        End If
        If mintVerify >= intLevel Then
            RaiseEvent AfterRowColChange("您的级别要比上次审签人的级别高才能勾选该记录！", True, mblnSign, mblnArchive)
            Exit Sub
        End If
        
        VsfData.Cell(flexcpChecked, lngStart, mlngChoose) = IIf(VsfData.Cell(flexcpChecked, lngStart, mlngChoose) = flexTSChecked, flexTSUnchecked, flexTSChecked)
        '保存修改记录以便同步
        strField = "ID|页号|行号|列号|记录ID|数据|删除"
        strKey = mint页码 & "," & lngStart & "," & mlngChoose
        strValue = strKey & "|" & mint页码 & "|" & lngStart & "|" & mlngChoose & "|" & Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & VsfData.Cell(flexcpChecked, lngStart, mlngChoose) & "|1"
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
    '清除常量
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
    '隐藏输入控件
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
    '功能：
    '参数：
    '返回：
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
    cbsThis.ActiveMenuBar.Title = "菜单栏"
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
    '菜单定义
    cbsThis.ActiveMenuBar.Title = "菜单"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    
        '------------------------------------------------------------------------------------------------------------------
        '工具栏定义
        Set cbrToolBar = cbsThis.Add("标准", xtpBarTop)
        cbrToolBar.ShowTextBelowIcons = False
        cbrToolBar.EnableDocking xtpFlagHideWrap + xtpFlagStretched
        cbrToolBar.ModifyStyle XTP_CBRS_GRIPPER, 0
        With cbrToolBar.Controls
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_FileMan, "婴儿"): cbrControl.ToolTipText = "文件份数"
            Set cbrCustom = .Add(xtpControlCustom, conMenu_View_Option, "")
            cbrCustom.Flags = xtpFlagAlignLeft
            picTmp.Visible = True
            cbrCustom.Handle = picTmp.Hwnd
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Copy, "复制"): cbrControl.ToolTipText = "复制(Ctrl+C)": cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_PASTE, "粘贴"):  cbrControl.ToolTipText = "粘贴(Ctrl+V)"
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Clear, "清除"):   cbrControl.ToolTipText = "清除"
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SPECIALCHAR, "特殊符号"):  cbrControl.ToolTipText = "插入特殊符号(Ctrl+D)": cbrControl.BeginGroup = True
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Word, "词句选择"):  cbrControl.ToolTipText = "词句选择(Ctrl+W)"
            'Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Brief, "小结"): cbrControl.ToolTipText = "小结"
            
            Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Append, "产程信息"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "产程信息"
        End With
    
        For Each cbrControl In cbrToolBar.Controls
            If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel Then
                cbrControl.Style = xtpButtonIconAndCaption
            End If
        Next
    
         '快键绑定
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
    '检查时点录入是否合法，并返回数据
    
    If Trim(objText.Text) <> "" Then
        strInput = Trim(objText.Text)
        If InStr(1, strInput, ":") > 0 Then
            strHour = Split(strInput, ":")(0)
            strMin = Split(strInput, ":")(1)
        ElseIf InStr(1, strInput, "：") > 0 Then
            strHour = Split(strInput, "：")(0)
            strMin = Split(strInput, "：")(1)
        Else
            strHour = strInput
            strMin = "00"
        End If
        strHour = Format(strHour, "00")
        strMin = Format(strMin, "00")
        If Not IsNumeric(strHour) Then
            RaiseEvent AfterRowColChange("开始时点中含有非法字符！", True, mblnSign, mblnArchive)
            Exit Function
        End If
        If Val(strHour) < 0 Or Val(strHour) > 23 Then
            RaiseEvent AfterRowColChange("开始时点不正确，小时值应该>0且小于24！", True, mblnSign, mblnArchive)
            Exit Function
        End If
        If Not IsNumeric(strMin) Then
            RaiseEvent AfterRowColChange("开始时点中含有非法字符！", True, mblnSign, mblnArchive)
            Exit Function
        End If
        If Val(strMin) < 0 Or Val(strMin) > 59 Then
            RaiseEvent AfterRowColChange("开始时点不正确，分钟值应该>0且小于60！", True, mblnSign, mblnArchive)
            Exit Function
        End If
        strInput = strHour & ":" & strMin
    End If
    CheckTxtTime = strInput
End Function

Private Function CheckTime(ByVal lngRow As Long, ByVal lng病人ID As Long, ByVal lng主页ID As Long, _
    ByVal strTime As String, ByVal strCurTime As String, ByRef strMsg As String) As Boolean
    Dim blnMsg As Boolean
    Dim blnExist As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '数据发生时间必须在产程时间和病人范围内
    '说明:由于病人可能院前就开始宫缩，所以产程时间存在小于病人入院时间的情况
    blnMsg = (strMsg <> "")
    
    '检查文件开始,结束时间
    If strTime < Format(mstr开始时间, "yyyy-MM-dd HH:mm") Or strTime > Format(mstr结束时间, "yyyy-MM-dd HH:mm") Then
        strMsg = "发生时间[" & strTime & "]不在开始时间[" & Format(mstr开始时间, "YYYY-MM-DD HH:mm") & "]到结束时间[" & Format(mstr结束时间, "YYYY-MM-DD HH:mm") & "]之间"
        GoTo exitHand
    End If
    
    '如果存在多份文件,上一文件的时间不能大于下一文件开始时间
    If cboBaby.ListCount > 1 And cboBaby.ListIndex < cboBaby.ListCount - 1 Then
        gstrSQL = "Select 1 From 病人护理文件 A,病人护理数据 B" & _
            "   Where A.ID=B.文件ID And A.ID=[1] And A.病人ID=[2] And A.主页ID=[3] And A.婴儿=[4] AND B.汇总类别=[5] And B.发生时间<=[6]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "数据检查", mlng文件ID, mlng病人ID, mlng主页ID, mint婴儿, cboBaby.ItemData(cboBaby.ListIndex + 1), CDate(strTime))
        If rsTemp.RecordCount > 0 Then
            strMsg = "第" & lngRow & "行的发生时间" & Format(strTime, "YYYY-MM-DD HH:mm") & "有误，不能大于下一婴儿文件的开始时间！"
            GoTo exitHand
        End If
    End If
    
    
    If Format(strTime, "YYYY-MM-DD HH:mm") > Format(DateAdd("d", mintPreDays, CDate(strCurTime)), "YYYY-MM-DD HH:mm") Then
        strMsg = "录入的日期已超出参数[超期录入天数：" & mintPreDays & "天]所指定的范围！"
        GoTo exitHand
    End If
            
    '根据病人变动记录进行检查
    gstrSQL = " Select   开始原因,病区ID,to_char(开始时间,'yyyy-MM-dd hh24:mi') AS 开始时间,to_char(NVL(终止时间,sysDate+" & mintPreDays & "),'yyyy-MM-dd hh24:mi') AS 终止时间 " & _
              " From 病人变动记录 " & _
              " Where 病人ID=[1] And 主页ID=[2]" & _
              " Order by 开始时间,开始原因"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "取当前科室有效时间范围", lng病人ID, lng主页ID)
    With rsTemp
        .Filter = "病区ID=" & mlng病区ID
        Do While Not .EOF
            If Format(strTime, "YYYY-MM-DD HH:mm") <= Format(!终止时间, "YYYY-MM-DD HH:mm") Then
                blnExist = True
                Exit Do
            End If
            .MoveNext
        Loop
        .Filter = 0
        '找到了就退出
        If blnExist Then
            If Not IsAllowInput(lng病人ID, lng主页ID, strTime, strCurTime) Then
                strMsg = "第" & lngRow & "行的发生时间" & strTime & "有误！[超过数据补录的有效时限:" & glngHours & "小时]"
                GoTo exitHand
            End If
            
            CheckTime = True
            Exit Function
        End If
        
        .Filter = 0
        '说明录入的数据时间不在产程时间和病人当前病区最大时间范围内
        strMsg = "第" & lngRow & "行的发生时间" & strTime & "有误！[不在产程开始时间至当前病区结束时间的有效范围内]"
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
    '检查录入数据的合法性(中文也认为是一个字符,考虑到体温项目等存在不升\外出等信息)
    '返回的数据,如果一列绑定多个项目,以单引号做为分隔符
    
    'mintType:0=文本框录入;1=单选;2=多选;3=选择;4-血压或一列绑定了两个项目,其格式类似血压的输入项目;5=一列绑定了两个项目且均是选择项目;
    '6=一列绑定N个项目,手工录入
    Select Case mintType
    Case 0
        strText = txtInput.Text
        strOrders = txtInput.Tag
    Case 1, 2   '免检
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
    Case 3      '免检
        strText = lblInput.Caption
    Case 5      '免检
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
                strInfo = "日期不能为空！"
                Exit Function
            End If
            If InStr(1, strText, "/") = 0 Then
                strInfo = "日期格式错误，如1月12日：12/01"
                Exit Function
            End If
            
            strDate = Mid(strCurrDate, 1, 5) & ToStandDate(strText)
                
            If Not IsDate(strDate) Then
                strInfo = "录入的数据不是合法的日期，如1月12日：12/01"
                Exit Function
            End If
        Else
            If Trim(strText) = "" Then
                strInfo = "日期不能为空！"
                Exit Function
            End If
            If Not IsDate(strText) Then
                strInfo = "录入的数据不是合法的日期，如：2011-01-12"
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
            strInfo = "时间不能为空！"
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
        
        '合法性检查
        If Mid(strText, 3, 1) <> ":" Then
            strInfo = "录入的时点格式非法！[小时:分钟]"
            Exit Function
        End If
        If Mid(strText, 1, 2) < 0 Or Mid(strText, 1, 2) > 23 Then
            strInfo = "录入的时点格式非法！[小时应在0至23之间]"
            Exit Function
        End If
        If Mid(strText, 4, 2) < 0 Or Mid(strText, 4, 2) > 59 Then
            strInfo = "录入的时点格式非法！[分钟应在0至59之间]"
            Exit Function
        End If
        
        '没有日期默认进行日期计算
        If mblnDate = False Then
            If Format(strText, "HH:mm") >= Format(mstr开始时间, "HH:mm") Then
                strDate = Format(mstr开始时间, "YYYY-MM-DD")
            Else
                strDate = Format(CDate(mstr开始时间) + 1, "YYYY-MM-DD")
            End If
            VsfData.TextMatrix(VsfData.ROW, mlngDate) = strDate
        End If
        '进行合法性检查
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
                strInfo = "录入的数据不是合法的日期，如：12:01"
                Exit Function
            End If
        
            blnCheck = True
        End If
    End If
    
    If blnCheck Then
        '检查录入的数据是否已经存在
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
                        strInfo = "您录入的时点已经存在历史数据，数据位置：第 " & (lngRow - VsfData.FixedRows + 1) & " 行!"
                        Exit Function
                    End If
                End If
            End If
        Next lngRow
        '数据发生时间不能在当前操作员所属科室的有效时间以前
        If Not CheckTime(VsfData.ROW, mlng病人ID, mlng主页ID, strDate, strCurrDate, strInfo) Then
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
'日期存在对角线，获取当前时间
    Dim strDate As String
    If Format(strTime, "HH:mm") >= Format(mstr开始时间, "HH:mm") Then
        strDate = Format(mstr开始时间, "YYYY-MM-DD")
    Else
        strDate = Format(CDate(mstr开始时间) + 1, "YYYY-MM-DD")
    End If
    GetDateAdCurrDate = Format(strDate & " " & Format(strTime, "HH:mm") & ":00", "yyyy-MM-dd HH:mm")
End Function

Private Function CheckValid(strReturn As String, ByVal strOrders As String, strInfo As String) As Boolean
    Dim arrData, arrOrder
    Dim blnCheck As Boolean
    Dim i As Integer, j As Integer
    Dim dblMin As Double, dblMax As Double
    Dim strText As String, strName As String, strFormat As String, strFormat1 As String
    
    '按列格式组装数据
    mrsSelItems.Filter = "列=" & VsfData.COL - (cHideCols + VsfData.FixedCols - 1)
    If mrsSelItems.RecordCount <> 0 Then
        '有此列但未进行定义
        strFormat = NVL(mrsSelItems!格式)   '{P[体温]C}{...}
        strFormat1 = strFormat
    End If
    mrsSelItems.Filter = 0
    
    '检查数据
    arrData = Split(strReturn, "'")
    arrOrder = Split(strOrders, "'")
    j = UBound(arrData)
    For i = 0 To j
        strText = arrData(i)
        strOrders = arrOrder(i)
        If Val(strOrders) <> 0 Then
            mrsItems.Filter = "项目序号=" & strOrders
            strName = mrsItems!项目名称
            If strText <> "" Then
                blnCheck = True
                '如果是曲线项目,如果输入的不是数字型则不检查
                If mrsItems!项目序号 >= 1 And mrsItems!项目序号 <= 3 Then
                    If Not IsNumeric(Trim(strText)) Then
                        blnCheck = False
                    End If
                End If
                
                If blnCheck Then
                    If mrsItems!项目类型 = 0 And InStr(1, "0,4", mrsItems!项目表示) <> 0 Then
                        strText = Val(strText)
                        If NVL(mrsItems!项目小数, 0) <> 0 Then   '等于零是通过控件的MaxLength来控制的
                            If InStr(1, strText, ".") <> 0 Then strText = Mid(strText, 1, InStr(1, strText, ".") - 1)
                            If Len(strText) > mrsItems!项目长度 Then
                                mrsItems.Filter = 0
                                strInfo = "[" & strName & "]录入的数据超过了合法精度！"
                                Exit Function
                            End If
                            
                            strText = Val(arrData(i))
                            If InStr(1, strText, ".") <> 0 Then
                                strText = Mid(strText, InStr(1, strText, ".") + 1)
                                If Len(strText) > mrsItems!项目小数 Then
                                    mrsItems.Filter = 0
                                    strInfo = "[" & strName & "]录入的小数部分超过了合法精度！"
                                    Exit Function
                                End If
                            End If
                            strText = Val(arrData(i))
                        End If
                        If mrsItems!项目表示 = 0 Then
                            If Not IsNull(mrsItems!项目值域) Then
                                dblMin = Val(Split(mrsItems!项目值域, ";")(0))
                                dblMax = Val(Split(mrsItems!项目值域, ";")(1))
                                If Not (Val(strText) >= dblMin And Val(strText) <= dblMax) Then
                                    mrsItems.Filter = 0
                                    strInfo = "[" & strName & "]录入的数据不在" & Format(dblMin, "#0.00") & "～" & Format(dblMax, "#0.00") & "的有效范围！"
                                    Exit Function
                                End If
                            End If
                        End If
                    Else
                        If LenB(StrConv(strText, vbFromUnicode)) > mrsItems!项目长度 Then
                            strInfo = "[" & strName & "]录入的数据超过了最大长度：" & mrsItems!项目长度 & "！"
                            mrsItems.Filter = 0
                            Exit Function
                        End If
                    End If
                End If
                strFormat = Replace(strFormat, "[" & strName & "]", strText)
            Else
                '删除该项目
                If InStr(1, "," & mstrCOLNothing & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 Then
                    Call SubstrPro(strFormat, strName)
                Else
                    '当项目无数据时,如果当前列具有对角线属性,则不清除
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
            mrsItems.Filter = "项目序号=" & strOrders
            strName = mrsItems!项目名称
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
    '获取绑定项目的前后缀符号
    Dim i As Integer
    Dim strOrders As String, strName As String
    For i = 0 To UBound(arrOrder)
        strOrders = CStr(arrOrder(i))
        If Val(strOrders) <> 0 Then
            mrsItems.Filter = "项目序号=" & strOrders
            strName = mrsItems!项目名称
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
    '返回前一个项目的后缀符号+当前项目的前缀符号的位置
    
    If strData = "" Then Exit Function
    strData = strData
    j = Len(strFormat)
    l = InStr(1, strFormat, "[" & strName & "]")
    If l = 0 Then Exit Function
    '得到前缀
    For i = l To 1 Step -1
        If Mid(strFormat, i, 1) = "{" Then Exit For
    Next
    strQZ = Mid(strFormat, i + 1, l - i - 1)
    '找到该项目格式串中的结束符号
    i = l + Len(strName) + 2
    For r = i To j
        If Mid(strFormat, r, 1) = "}" Then Exit For
    Next
    '得到后缀
    strHZ = Mid(strFormat, i, r - i)
    '如果后缀为空,继续向后寻找下一个项目的前缀符号
    If strHZ = "" And r < j Then
        For r = r + 1 To j
            If Mid(strFormat, r, 1) = "[" Then Exit For
        Next
        strHZ = Mid(strFormat, InStr(i, strFormat, "{") + 1, r - InStr(i, strFormat, "{") - 1)
    End If
    '取出指定项目完整的数据串
    If strHZ <> "" Then
        j = InStr(intPos, strData, strHZ) '因为是连续取数,考虑到分隔符可能相同的情况,记录上一次的最后位置,下次从这个位置往后取数据
        If j = 0 Then
            '有可能中间存在回车换行符
            j = InStr(intPos, Replace(strData, vbCrLf, ""), strHZ)
            If j = 0 Then Exit Function
        End If
    End If
    strData = Mid(strData, intPos)
    '前缀为空,继续向前寻找上一个项目的后缀符号
'    If strQZ = "" And i > 1 And intPos > 1 Then
'        For i = i - 1 To 1 Step -1
'            If Mid(strFormat, i, 1) = "]" Then Exit For
'        Next
'        strQZ = Mid(strFormat, i + 1, InStr(i, strFormat, "}") - i - 1)
'    End If
    
    SubstrVal = SubstrAnaly(strData, strHZ, strQZ)
    intPos = intPos + Len(strQZ & SubstrVal & strHZ)
    '如果是数字型则去掉回车换行符返回,如果是字符型则原样返回
'    If strHZ <> "" Then
'
'        strData = Mid(strData, 1, InStr(1, Replace(strData, vbCrLf, ""), strHZ) - 1) '丢弃该项目后的数据
'        intPOS = i + Len(strHZ)
'    End If
'    If strQZ <> "" Then strData = Mid(strData, InStr(1, strData, strQZ) + Len(strQZ)) '丢弃该项目后的数据
'    SubstrVal = strData ' Replace(strData, vbCrLf, "")
End Function

Private Function SubstrAnaly(ByVal strData As String, ByVal strHZ As String, ByVal strQZ As String) As String
    Dim strText As String
    Dim strCompare As String           '对比串
    Dim intLen As Integer, intActLen As Integer           '前缀/后缀的长度
    Dim intPos As Integer, intEnd As Integer
    Dim lngASC As Long
    Dim blnFind As Boolean
    '遇到回车换行符忽略,空格重新比对
    
    strText = strData
    If strHZ <> "" Then
        '把后缀去掉
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
        '肯定有
        strText = Mid(strText, 1, intPos)
    End If
    
    '再去掉前缀
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
    'intType=0-删除指定格式串;1-得到指定格式串
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
    Dim blnNULL As Boolean                      '是否为空行
    Dim strDate As String, strTime As String
    Dim strReturn As String, strMsg As String, strPart As String
    Dim lngStart As Long, lngMutilRows As Long, lngDeff As Long
    Dim intRow As Integer, intCount As Integer, intNULL As Integer  '其后有多少空行
    Dim blnTrue As Boolean
    '赋值然后移动到下一个有效单元格
    Dim strKey As String, strField As String, strValue As String
    
    On Error GoTo errHand
    
    '检查数据,不合格就再次弹出要求录入
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

        '准备赋值
        With txtLength
            '日期与时间列的宽度不管,为了避免返回多行,强制设置为5000
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
            '往下搜索空行,如果有其它数据行则计算需增加的行数
            intNULL = intCount - (lngMutilRows - 1)
            For intRow = lngMutilRows To intCount
                '保证当前输入的内容在一页中显示全
                If intRow + lngStart > VsfData.Rows - 1 Then Exit For
                
                If Val(VsfData.TextMatrix(intRow + lngStart, mlngRecord)) = 0 And VsfData.TextMatrix(intRow + lngStart, mlngRowCount) = "" Then
                    intNULL = intNULL - 1
                Else
                    Exit For
                End If
            Next
            '先增加空行
            If intNULL > 0 Then
                lngDeff = intNULL
                VsfData.Rows = VsfData.Rows + intNULL
                '从当前行记录的空白行开始，每行的位置+所增加的空白行数
                For intRow = VsfData.Rows - intNULL - 1 To lngStart + intCount - intNULL + 1 Step -1
                    VsfData.RowPosition(intRow) = intRow + intNULL
                Next
            End If
            '循环赋值
            intCount = UBound(arrData)
            For intRow = 0 To intCount
                VsfData.TextMatrix(lngStart + intRow, VsfData.COL) = Replace(Replace(arrData(intRow), Chr(10), ""), Chr(13), "")
                VsfData.TextMatrix(lngStart + intRow, mlngRowCount) = intCount + 1 & "|" & intRow + 1
            Next
            '所有隐蔽列进行赋值
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
            '对该列重新赋值（当只输入一个数字时，不知为何会产生字符ASCII码为1的符号）
            For intRow = 0 To intCount
                VsfData.TextMatrix(lngStart + intRow, VsfData.COL) = Replace(Replace(Replace(arrData(intRow), Chr(10), ""), Chr(13), ""), Chr(1), "")
            Next
            For intRow = intCount + 1 To lngMutilRows - 1
                VsfData.TextMatrix(lngStart + intRow, VsfData.COL) = ""
            Next
            
            '根据行数据重新填写行序列,intNULL记录最后一条不为空行的行号
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
            
            '从新填写行序号
            For intRow = lngStart To intNULL
                VsfData.TextMatrix(intRow, mlngRowCount) = (intNULL - lngStart + 1) & "|" & intRow - lngStart + 1
            Next
            For intRow = intNULL + 1 To lngStart + lngMutilRows - 1
                VsfData.Cell(flexcpText, intRow, 0, intRow, VsfData.Cols - 1) = ""
            Next
        End If
        
        '当行号发生变化后，需同步更新mrsCellMap中大于该行号的行号数据
        If lngDeff <> 0 Then
            Call CellMap_Update(lngStart, lngDeff)
        End If
        
        If mstrData <> strReturn Then
            mblnChange = True
            
            '同步保存日期与时间列的数据
            strDate = VsfData.TextMatrix(lngStart, mlngDate)
            strTime = VsfData.TextMatrix(lngStart, mlngTime)
            
            '1\日期
            strField = "ID|页号|行号|列号|记录ID|数据|删除"
            If mlngDate <> -1 Then
                strKey = mint页码 & "," & lngStart & "," & mlngDate
                strValue = strKey & "|" & mint页码 & "|" & lngStart & "|" & mlngDate & "|" & _
                    Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & strDate & "|0"
                Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
            End If
            '2\时间
            strKey = mint页码 & "," & lngStart & "," & mlngTime
            strValue = strKey & "|" & mint页码 & "|" & lngStart & "|" & mlngTime & "|" & _
                Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & strTime & "|0"
            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
            
            '记录用户修改过的单元格
            If InStr(1, "," & mstrCatercorner & ",", "," & VsfData.COL - (cHideCols + VsfData.FixedCols - 1) & ",") = 0 Then
                strPart = ""
            Else
                strPart = "/"
            End If
            
            strField = "ID|页号|行号|列号|记录ID|数据|部位|删除"
            strKey = mint页码 & "," & lngStart & "," & VsfData.COL
            strValue = strKey & "|" & mint页码 & "|" & lngStart & "|" & VsfData.COL & "|" & _
                Val(VsfData.TextMatrix(lngStart, mlngRecord)) & "|" & strReturn & "|" & strPart & "|" & IIf(strReturn = "", "1", "0")
            Call Record_Update(mrsCellMap, strField, strValue, "ID|" & strKey)
        End If
        RaiseEvent AfterDataChanged(mblnChange Or mblnVerify)
    End If
    
    '始终要保留一行空行，以便录入其它数据
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
            '跳到下一行
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
'            '跳到上一行
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
    '提取数据起始行,超出本页则返回0
    '如果本页未显示全,则说明超出本页,也返回0
    '不允许在连续的数据行中插入新行
    
    lngRows = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(0))    '总行数
    lngCurRows = Val(Split(VsfData.TextMatrix(lngRow, mlngRowCount), "|")(1)) '当前行
    If lngCurRows = 1 Then
        GetStartRow = lngRow
        Exit Function
    End If
    
    '寻找起始行
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
    Dim lngStart As Long    '起始行
    Dim lngRecordId As Long
    Dim strReturn As String
    Dim blnAdjust As Boolean
    Dim rsTemp As New ADODB.Recordset
    On Error GoTo errHand
    '返回第一行的坐标
    '不分行直接取，分行时检查如果当页显示全就拼接，否则从库中读取
    
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
        '从数据库中提取
        Call SQLCombination(lngRecordId)
        gstrSQL = mstrSQL
        Call SQLDIY(gstrSQL)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "提取护理数据", mlng文件ID, mlng病人ID, mlng主页ID, mint婴儿, cboBaby.ItemData(cboBaby.ListIndex), lngRecordId)
        strReturn = NVL(rsTemp.Fields(lngCol - VsfData.FixedCols).Value)
        blnAdjust = True
    Else
        For lngRow = lngStart To lngStart + lngCount - 1
            strReturn = strReturn & VsfData.TextMatrix(lngRow, lngCol)
        Next
    End If
    
    '取行高
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
    Dim strFormat As String, strText As String, strValue As String  '格式串,数据串,数值串
    Dim strOrders As String, strTypes As String, strBounds As String, strLen As String, strName As String
    Dim strCurDate As String
    Const txtHeight = 300
    On Error GoTo errHand
    
    '病历文件构造管理模块需要处理:
    '1、一列绑定一个项目的不用管
    '2、一列绑定两个项目的，血压必须成对，要么都是录入，要么都是选择，不允许交叉出现，也不允许出现单选、复选
    '3、一列绑定多个项目的，只能是录入项目
    '由于以上条件限制，只取第一个项目的性质即可
    
    '如果是保存处调用则做如下处理
    If intCOl = -1 Then intCOl = VsfData.COL
    If blnAnalyse Then
        strText = strCellData
    Else
        '取当前单元格的属性
        CellRect.Left = VsfData.CellLeft + VsfData.Left
        CellRect.Top = VsfData.CellTop + VsfData.Top
        CellRect.Bottom = VsfData.CellHeight + 20
        CellRect.Right = VsfData.CellWidth + 20
        strText = GetMutilData(VsfData.ROW, intCOl, CellRect.Top, CellRect.Bottom)
    End If
    mstrData = strText
    mintType = 0
    intIndex = 0
    
    '取当前列的绑定项目
    intPos = 1
    mrsSelItems.Filter = "列=" & intCOl - cHideCols
    Do While Not mrsSelItems.EOF
        lngOrder = mrsSelItems!项目序号
        If lngOrder = 0 Then
            strLen = 0
            strValue = strText
            Exit Do
        End If
        
        '项目表示:2单选;3-多选;4-汇总;5-选择
        '项目值域:项目表示为0-表示最小值;最大值;项目表示为2,3-表示项目A;项目B,前有勾的表示缺省项
        strFormat = NVL(mrsSelItems!格式)
        strOrders = strOrders & "," & lngOrder
        If lngOrder <> 0 Then
            mrsItems.Filter = "项目序号=" & lngOrder
            strName = strName & "," & mrsItems!项目名称
            strLen = strLen & "," & mrsItems!项目长度 & ";" & NVL(mrsItems!项目小数)
            strTypes = strTypes & "," & mrsItems!项目表示
            strBounds = strBounds & "," & mrsItems!项目值域
            strValue = strValue & "'" & SubstrVal(strText, strFormat, mrsItems!项目名称, intPos)
            
            Select Case mrsItems!项目表示
            Case 0  '文本录入项
                If mrsSelItems.RecordCount = 2 Then
                    mintType = 4
                ElseIf mrsSelItems.RecordCount > 2 Then
                    mintType = 6
                End If
            Case 2  '单选
                If mrsSelItems.RecordCount = 1 Then
                    mintType = 1
                ElseIf mrsSelItems.RecordCount = 2 Then
                    mintType = 7
                End If
            Case 3  '多选
                mintType = 2
            Case 4  '汇总
                If mrsSelItems.RecordCount = 2 Then
                    mintType = 4
                ElseIf mrsSelItems.RecordCount > 2 Then
                    mintType = 6
                End If
            Case 5  '选择
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
    
    '针对4进行校对,如果表头文本不含/则处理为6
    If mintType = 4 Then
        If Not IsDiagonal(intCOl) Then
            mintType = 6
        End If
    End If
    
    '判断当前列的性质
    'mintType:0=文本框录入;1=单选;2=多选;3=选择;4-血压或一列绑定了两个项目,其格式类似血压的输入项目;5=一列绑定了两个项目且均是选择项目;
    '6=一列绑定2个及以上项目,手工录入,7=一列绑定了两个单选项目
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
                txtInput.MaxLength = Val(Split(strLen, ";")(0)) + IIf(Val(Split(strLen, ";")(1)) = 0, 0, 1) '小数位数要加上小数点
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
            .Width = .Width - (180 + IIf(mblnBlowup, 180 * 1 / 3, 0)) / 2 '宋体9号时减去90,字体越大扣除的边距越小,以保证文本框分行与实际一致
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
        
        '如果是日期或时间列，设定固定值
        strCurDate = zlDatabase.Currentdate
        If mintType = 0 And txtInput.Text = "" Then
            If intCOl = mlngDate Or intCOl = mlngTime Then
                txtInput.Text = Format(strCurDate, "YYYY-MM-DD HH:mm")
                If Format(strCurDate, "YYYY-MM-DD HH:mm") >= Format(mstr结束时间, "YYYY-MM-DD HH:mm") Then
                    txtInput.Text = Format(mstr结束时间, "YYYY-MM-DD HH:mm")
                End If
                If Format(strCurDate, "YYYY-MM-DD HH:mm") <= Format(mstr开始时间, "YYYY-MM-DD HH:mm") Then
                    txtInput.Text = Format(mstr开始时间, "YYYY-MM-DD HH:mm")
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
        '加载数据
        lstSelect(mintType - 1).Clear
        If mintType = 1 Then
            lstSelect(mintType - 1).AddItem "0-"
            If mlngProduce = intCOl Then lstSelect(mintType - 1).ListIndex = 0
        End If
        arrData = Split(strBounds, ";")
        j = UBound(arrData)
        For i = 0 To j
            If arrData(i) <> "" Then
                If Mid(arrData(i), 1, 1) = "√" Then
                    lstSelect(mintType - 1).AddItem lstSelect(mintType - 1).NewIndex + 1 & "-" & Mid(arrData(i), 2)
                    If strText = "" And lstSelect(mintType - 1).ListIndex = -1 Then lstSelect(mintType - 1).ListIndex = lstSelect(mintType - 1).NewIndex
                Else
                    lstSelect(mintType - 1).AddItem lstSelect(mintType - 1).NewIndex + 1 & "-" & arrData(i)
                End If
            End If
        Next
        '多选且已录入数据的情况下
        If strValue <> "" Then
            strValue = Replace(strValue, vbCrLf, "")
            j = lstSelect(mintType - 1).ListCount - 1
            For i = 0 To j
                If InStr(1, "," & strValue & ",", "," & Split(lstSelect(mintType - 1).List(i), "-")(1) & ",") <> 0 Then
                    lstSelect(mintType - 1).Selected(i) = True
                End If
            Next
        End If
        '显示
        With lstSelect(mintType - 1)
            .Left = CellRect.Left
            .Top = CellRect.Top
            .FontName = VsfData.FontName
            .FontSize = VsfData.FontSize
            .Height = .ListCount * 300
            If .Height < CellRect.Bottom Then .Height = CellRect.Bottom
            .Width = LenB(StrConv(.List(.ListCount \ 2), vbFromUnicode)) * 120 + 500    '以中间项的长度为依据
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
            If strLen <> "" Then txtUpInput.MaxLength = Val(Split(Split(strLen, ",")(0), ";")(0)) + IIf(Val(Split(Split(strLen, ",")(0), ";")(1)) = 0, 0, 1) '小数位数要加上小数点
            If strLen <> "" Then txtDnInput.MaxLength = Val(Split(Split(strLen, ",")(1), ";")(0)) + IIf(Val(Split(Split(strLen, ",")(1), ";")(1)) = 0, 0, 1) '小数位数要加上小数点
        End If
    Case 6
        '先删除以前的控件
        j = txt.Count - 1
        For i = 1 To j
            Unload lbl(i)
            Unload txt(i)
        Next
        '设定坐标
        With picMutilInput
            .Left = CellRect.Left
            .Top = CellRect.Top
            .Width = IIf(CellRect.Right < 1600, 1600, CellRect.Right)
        End With
        '对缺省控件赋值
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
        txt(0).MaxLength = Val(Split(Split(strLen, ",")(0), ";")(0)) + IIf(Val(Split(Split(strLen, ",")(0), ";")(1)) = 0, 0, 1)  '小数位数要加上小数点
        txt(0).Text = arrValue(0)
        If Not mblnBlowup Then
            txt(0).Height = 225
        End If
        
        '加载控件
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
                    .MaxLength = Val(Split(Split(strLen, ",")(i), ";")(0)) + IIf(Val(Split(Split(strLen, ",")(i), ";")(1)) = 0, 0, 1) '小数位数要加上小数点
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
            If Mid(arrData(i), 1, 1) = "√" Then
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
            If Mid(arrData(i), 1, 1) = "√" Then
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
    '如果格式与血压的方式不同,则将样式处理为6
    
    '去掉前缀后进行对比
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
    '判断指定列是否设置了列对角线（mstrColWidth的格式：765`11`1`1,765`11`2`1,...，对象属性`对象序号`列对角线）
    
    IsDiagonal = (InStr(1, "," & mstrCatercorner & ",", "," & intCOl - (cHideCols + VsfData.FixedCols - 1) & ",") <> 0)
End Function

Private Sub ISAssistant(ByVal lngOrder As Long, ByVal objTXT As TextBox)
    Dim intIndex As Integer
    Dim objParent As Object
    
    '根据项目的长度决定是否允许进行词句选择
    mblnEditAssistant = False
    cmdWord.Visible = mblnEditAssistant
    
    mrsItems.Filter = "项目序号=" & lngOrder
    If mrsItems.RecordCount = 0 Then
        mrsItems.Filter = 0
        Exit Sub
    End If
    mblnEditAssistant = (mrsItems!项目类型 = 1 And mrsItems!项目长度 >= 100)
    mrsItems.Filter = 0
    
    '如果允许词句选择,显示并定位
    If mblnEditAssistant Then
        If UCase(objTXT.Name) = "TXTINPUT" Then
            intIndex = -1 '表示txtInput
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
    'intType，确定函数返回值，1)返回项目序号;2)返回列号
    'intMAX，返回同步数据列所占用的行高
    '返回同步数据列(一份文件中不可能出现重复的项目,所以,判断时不必检查列号)
    
    lngRecord = Val(VsfData.TextMatrix(VsfData.ROW, mlngRecord))
    If lngRecord = 0 Then Exit Function
    
    gstrSQL = "" & _
        " SELECT  B.项目序号,B.项目名称,A.对象序号 AS 列号" & vbNewLine & _
        " FROM 病历文件结构 A,病人护理明细 B" & vbNewLine & _
        " WHERE A.要素名称=B.项目名称 AND A.父ID=" & vbNewLine & _
        "      (SELECT A.ID FROM 病历文件结构 A,病人护理文件 B " & vbNewLine & _
        "       WHERE B.ID=[2] And A.文件ID=B.格式ID AND A.对象序号=4 AND A.父ID IS NULL)" & vbNewLine & _
        " AND B.数据来源>0 AND B.记录ID=[1]"
    Call SQLDIY(gstrSQL)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "返回同步数据列", lngRecord, mlng文件ID)
    If rsTemp.RecordCount = 0 Then Exit Function
    
    '获取同步相关信息
    Do While Not rsTemp.EOF
        If InStr(1, "," & strCols & ",", "," & rsTemp!列号 & ",") = 0 Then strCols = strCols & "," & rsTemp!列号
        strItems = strItems & "," & rsTemp!项目序号
        strNames = strNames & "," & rsTemp!项目名称
        rsTemp.MoveNext
    Loop
    strCols = Mid(strCols, 2)
    strItems = Mid(strItems, 2)
    strNames = Mid(strNames, 2)
    
    '根据列循环检查内容所占行高
    If InStr(1, VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|") = 0 Then
        lngStartRow = VsfData.ROW
        lngEndRow = VsfData.ROW
        intInMAX = 1
    Else
        lngStartRow = GetStartRow(VsfData.ROW)
        intInMAX = Val(Split(VsfData.TextMatrix(VsfData.ROW, mlngRowCount), "|")(0))
        lngEndRow = lngStartRow + intInMAX - 1
    End If
    
    intCount = 1    '同步的只能是数字型项目，所以占用行只可能是1行，以下内容不再需要检查
'    '数据占用超过1行才检查
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
        RaiseEvent AfterRowColChange("日期列,时间列,以及 " & strNames & " 是同步过来的数据，不允许修改或删除！", True, mblnSign, mblnArchive)
    End If
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

'######################################################################################################################
'**********************************************************************************************************************
'以下是基础函数或过程

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
    lblUpInput.Caption = IIf(lblUpInput.Caption = "", "√", "")
    txtUpInput.SetFocus
End Sub

Private Sub lblDnInput_DblClick()
    lblDnInput.Caption = IIf(lblDnInput.Caption = "", "√", "")
    txtDnInput.SetFocus
End Sub

Private Sub lblInput_DblClick()
    lblInput.Caption = IIf(lblInput.Caption = "", "√", "")
End Sub

Private Sub txtUpInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtDnInput.SetFocus
    ElseIf KeyCode = vbKeyRight Then
        If txtUpInput.SelStart = Len(txtUpInput.Text) Then txtDnInput.SetFocus
    ElseIf KeyCode = vbKeyLeft And txtUpInput.SelStart = 1 Then
        Call MoveNextCell(False)
    ElseIf KeyCode = vbKeySpace And txtUpInput.Locked Then
        lblUpInput.Caption = IIf(lblUpInput.Caption = "", "√", "")
    End If
End Sub

Private Sub txtDnInput_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Or (KeyCode = vbKeyRight And txtDnInput.SelStart = Len(txtDnInput.Text)) Then
        Call picDouble_KeyDown(KeyCode, Shift)
    ElseIf KeyCode = vbKeyLeft Then
        If txtDnInput.SelStart = 0 Then txtUpInput.SetFocus
    ElseIf KeyCode = vbKeySpace And txtDnInput.Locked Then
        lblDnInput.Caption = IIf(lblDnInput.Caption = "", "√", "")
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
        '移动到下一个单元格
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
 

Private Sub txt小结名称_KeyPress(KeyAscii As Integer)
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
    '以下字符做为数据分隔符或更新记录集的分隔符，因此不允许录入
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
    
    '表上标签分散处理
    Call zlLableBruit
End Sub

Private Sub UserControl_Terminate()
'    objStream.Close
End Sub

Private Sub SetDockRight(BarToDock As CommandBar, BarOnLeft As CommandBar)
    '******************************************************************************************************************
    '功能：
    '参数：
    '返回：
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
    '添加记录
    'strFields:字段名|字段名
    'strValues:值|值
    
    '例子：
    'Dim strFields As String, strValues As String
    'strFields = "RecordID|科目ID|摘要"
    'strValues = "5188|6666|科目名称"
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
    '更新记录,如果不存在,则新增
    'strPrimary:字段名|值
    'strFields:字段名|字段名
    'strValues:值|值
    
    '例子：
    'Dim strFields As String, strValues As String, strPrimary As String
    'strFields = "RecordID|科目ID|摘要"
    'strValues = "5188|6666|科目名称"
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
    '定位到指定记录
    'strPrimary:主健,值
    'blnDelete=True,则该记录集存在"删除"字段
    Record_Locate = False
    
    arrTmp = Split(strPrimary, "|")
    With rsObj
        If .RecordCount = 0 Then Exit Function
        .MoveFirst
        .Find arrTmp(0) & "='" & arrTmp(1) & "'"
        If .EOF Then Exit Function
        If blnDelete Then
            Do While Not .EOF
                If !删除 = 0 Then Record_Locate = True: Exit Do
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
    '初始化映射记录集
    'strFields:字段名,类型,长度|字段名,类型,长度    如果长度为零,则取默认长度
    '字符型:adLongVarChar;数字型:adDouble;日期型:adDBDate
    
    '例子：
    'Dim rsVoucher As New ADODB.Recordset, strFields As String
    'strFields = "RecordID," & adDouble & ",18|科目ID," & adDouble & ",18|摘要, " & adLongVarChar & ",50|" & _
    '"删除," & adDouble & ",1"
    'Call Record_Init(rsVoucher, strFields)

    arrFields = Split(strFields, "|")
    Set rsObj = New ADODB.Recordset

    With rsObj
        If .State = 1 Then .Close
        For intField = 0 To UBound(arrFields)
            strFieldName = Split(arrFields(intField), ",")(0)
            intType = Split(arrFields(intField), ",")(1)
            lngLength = Split(arrFields(intField), ",")(2)

            '获取字段缺省长度
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
    '放大：字体，单元格宽度
    BlowUp = dblChange
    If Not mblnBlowup Then Exit Function
    BlowUp = dblChange + (dblChange * 1 / 3)
End Function



