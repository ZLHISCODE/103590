VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmParameters 
   BackColor       =   &H80000005&
   Caption         =   "ϵͳ��������"
   ClientHeight    =   6288
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   10728
   ControlBox      =   0   'False
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form3"
   MDIChild        =   -1  'True
   Picture         =   "frmParameters.frx":0000
   ScaleHeight     =   6288
   ScaleWidth      =   10728
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExp 
      Caption         =   "����(&E)"
      Height          =   350
      Left            =   7560
      TabIndex        =   21
      Top             =   935
      Width           =   1100
   End
   Begin VB.CommandButton cmdImp 
      Caption         =   "����(&I)"
      Height          =   350
      Left            =   7560
      TabIndex        =   20
      Top             =   560
      Width           =   1100
   End
   Begin VB.Frame fraSplit 
      BackColor       =   &H80000012&
      Height          =   30
      Left            =   0
      MousePointer    =   7  'Size N S
      TabIndex        =   13
      Top             =   3360
      Width           =   9975
   End
   Begin VB.ComboBox cboParType 
      Height          =   300
      Left            =   1605
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   960
      Width           =   2445
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   6600
      Top             =   0
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameters.frx":04F9
            Key             =   "��������ģ��"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameters.frx":0A93
            Key             =   "����˽��ģ��"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameters.frx":102D
            Key             =   "����ģ��"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameters.frx":15C7
            Key             =   "˽��ģ��"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameters.frx":1B61
            Key             =   "˽��ȫ��"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameters.frx":20FB
            Key             =   "����ȫ��"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameters.frx":2695
            Key             =   "���Ų���"
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmdModify 
      Caption         =   "�ı��������(&M)"
      Height          =   350
      Left            =   5640
      TabIndex        =   7
      Top             =   935
      Width           =   1665
   End
   Begin VB.PictureBox picPara 
      BorderStyle     =   0  'None
      Height          =   2025
      Left            =   60
      ScaleHeight     =   2028
      ScaleWidth      =   9828
      TabIndex        =   9
      Top             =   1320
      Width           =   9825
      Begin VSFlex8Ctl.VSFlexGrid vsPara 
         Height          =   1890
         Left            =   30
         TabIndex        =   10
         Top             =   90
         Width           =   7470
         _cx             =   13176
         _cy             =   3334
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
         BackColor       =   -2147483628
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483628
         GridColor       =   12632256
         GridColorFixed  =   -2147483630
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483628
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   17
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmParameters.frx":8EF7
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
         ExplorerBar     =   7
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
   Begin VB.CheckBox chkShowFixed 
      BackColor       =   &H80000005&
      Caption         =   "�̶�����(&H)"
      Height          =   330
      Left            =   4185
      TabIndex        =   6
      Top             =   945
      Width           =   1410
   End
   Begin VB.ComboBox cboModule 
      Height          =   300
      Left            =   4605
      TabIndex        =   5
      Text            =   "cmbModule"
      Top             =   585
      Width           =   2700
   End
   Begin VB.PictureBox picPage 
      BackColor       =   &H00FFEBD7&
      BorderStyle     =   0  'None
      Height          =   2610
      Left            =   0
      ScaleHeight     =   2616
      ScaleWidth      =   10500
      TabIndex        =   11
      Top             =   3720
      Width           =   10500
      Begin VB.PictureBox picDetailParas 
         BorderStyle     =   0  'None
         Height          =   2220
         Left            =   240
         ScaleHeight     =   2220
         ScaleWidth      =   10212
         TabIndex        =   16
         Top             =   120
         Width           =   10215
         Begin VB.Frame fraDetaisModi 
            BorderStyle     =   0  'None
            Height          =   615
            Left            =   120
            TabIndex        =   22
            Top             =   120
            Width           =   10095
            Begin VB.CommandButton cmdDel 
               Caption         =   "ɾ����������(&D)"
               Height          =   350
               Left            =   5535
               TabIndex        =   29
               TabStop         =   0   'False
               Top             =   120
               Width           =   1600
            End
            Begin VB.CommandButton cmdAddNew 
               Caption         =   "������������(&N)"
               Height          =   350
               Left            =   7320
               TabIndex        =   27
               TabStop         =   0   'False
               Top             =   120
               Width           =   1600
            End
            Begin VB.CommandButton cmdSearch 
               Height          =   240
               Left            =   3435
               Picture         =   "frmParameters.frx":9154
               Style           =   1  'Graphical
               TabIndex        =   24
               TabStop         =   0   'False
               Top             =   180
               Width           =   240
            End
            Begin VB.CommandButton cmdModValue 
               Caption         =   "�޸Ĳ���ֵ(&B)"
               Height          =   350
               Left            =   3960
               TabIndex        =   23
               TabStop         =   0   'False
               Top             =   120
               Width           =   1400
            End
            Begin VB.TextBox txtSearch 
               Height          =   300
               Left            =   2160
               MaxLength       =   50
               TabIndex        =   25
               Top             =   150
               Width           =   1545
            End
            Begin VB.Label lblTip 
               AutoSize        =   -1  'True
               Caption         =   "�������ң�"
               Height          =   180
               Index           =   0
               Left            =   0
               TabIndex        =   28
               Top             =   240
               Width           =   900
            End
            Begin VB.Label lblSearch 
               AutoSize        =   -1  'True
               Caption         =   "�û���(&U)��"
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
               Left            =   960
               TabIndex        =   26
               Tag             =   "1"
               Top             =   210
               Width           =   1095
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsDetailParas 
            Height          =   840
            Left            =   240
            TabIndex        =   17
            Top             =   960
            Width           =   7710
            _cx             =   13600
            _cy             =   1482
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
            BackColor       =   -2147483628
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483628
            GridColor       =   12632256
            GridColorFixed  =   -2147483630
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483628
            FocusRect       =   1
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   12
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   300
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParameters.frx":9468
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
            ExplorerBar     =   1
            PicturesOver    =   0   'False
            FillStyle       =   0
            RightToLeft     =   0   'False
            PictureType     =   2
            TabBehavior     =   0
            OwnerDraw       =   0
            Editable        =   2
            ShowComboButton =   0
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
            WallPaperAlignment=   0
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
      End
      Begin VB.PictureBox picParInfo 
         BorderStyle     =   0  'None
         Height          =   2295
         Left            =   3120
         ScaleHeight     =   2292
         ScaleWidth      =   6132
         TabIndex        =   18
         Top             =   1080
         Width           =   6135
         Begin VSFlex8Ctl.VSFlexGrid vsParaInfo 
            Height          =   2160
            Left            =   60
            TabIndex        =   19
            Top             =   120
            Width           =   6015
            _cx             =   10610
            _cy             =   3810
            Appearance      =   2
            BorderStyle     =   0
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
            BackColor       =   -2147483633
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483633
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483633
            BackColorBkg    =   -2147483633
            BackColorAlternate=   -2147483633
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483633
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483628
            FocusRect       =   1
            HighLight       =   0
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   0
            GridLinesFixed  =   0
            GridLineWidth   =   1
            Rows            =   5
            Cols            =   2
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParameters.frx":9604
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
            ExplorerBar     =   7
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
            ForeColorFrozen =   -2147483633
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
      End
      Begin VB.PictureBox picParaChangeLog 
         BorderStyle     =   0  'None
         Height          =   1695
         Left            =   240
         ScaleHeight     =   1692
         ScaleWidth      =   7692
         TabIndex        =   14
         Top             =   1080
         Width           =   7695
         Begin VSFlex8Ctl.VSFlexGrid vsChangeLog 
            Height          =   1320
            Left            =   0
            TabIndex        =   15
            Top             =   120
            Width           =   7470
            _cx             =   13176
            _cy             =   2328
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
            BackColor       =   -2147483628
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483628
            GridColor       =   12632256
            GridColorFixed  =   -2147483630
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483628
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParameters.frx":967F
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
            ExplorerBar     =   7
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
      Begin XtremeSuiteControls.TabControl tbPage 
         Height          =   960
         Left            =   0
         TabIndex        =   12
         Top             =   360
         Width           =   900
         _Version        =   589884
         _ExtentX        =   1587
         _ExtentY        =   1693
         _StockProps     =   64
      End
   End
   Begin VB.ComboBox cboSys 
      Height          =   300
      Left            =   1605
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   585
      Width           =   2445
   End
   Begin MSComctlLib.ImageList ils16 
      Left            =   7050
      Top             =   6300
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameters.frx":9766
            Key             =   "Root"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList imgEdit 
      Left            =   7200
      Top             =   0
      _ExtentX        =   995
      _ExtentY        =   995
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameters.frx":B4F8
            Key             =   "Check"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameters.frx":BA92
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameters.frx":C02C
            Key             =   "ǩ��"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameters.frx":C37E
            Key             =   "Woman"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameters.frx":12BE0
            Key             =   "Man"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameters.frx":19442
            Key             =   "UnCheck"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmParameters.frx":1990A
            Key             =   "AllCheck"
         EndProperty
      EndProperty
   End
   Begin VB.Label lbl���� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "��������"
      Height          =   180
      Left            =   780
      TabIndex        =   2
      Top             =   1020
      Width           =   720
   End
   Begin VB.Label lblģ�� 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ģ��"
      Height          =   180
      Left            =   4185
      TabIndex        =   4
      Top             =   645
      Width           =   360
   End
   Begin VB.Image imgMain 
      Height          =   384
      Left            =   180
      Picture         =   "frmParameters.frx":19DD2
      Top             =   576
      Width           =   384
   End
   Begin VB.Label lblSys 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Ӧ��ϵͳ"
      Height          =   180
      Left            =   840
      TabIndex        =   0
      Top             =   645
      Width           =   720
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "ϵͳ��������"
      BeginProperty Font 
         Name            =   "����"
         Size            =   12
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   195
      TabIndex        =   8
      Top             =   105
      Width           =   1440
   End
   Begin VB.Menu mnuPopuMenu 
      Caption         =   "�����˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuPopuMenuSerch 
         Caption         =   "�û���(&U)"
         Index           =   0
         Shortcut        =   ^U
      End
      Begin VB.Menu mnuPopuMenuSerch 
         Caption         =   "��  Ա(&P)"
         Index           =   1
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuPopuMenuSerch 
         Caption         =   "��  ��(&W)"
         Index           =   2
         Shortcut        =   ^W
      End
      Begin VB.Menu mnuPopuMenuSerch 
         Caption         =   "������(&T)"
         Index           =   3
         Shortcut        =   ^T
      End
      Begin VB.Menu mnuPopuMenuSerch 
         Caption         =   "Ժ  ��(&S)"
         Index           =   4
         Shortcut        =   ^S
      End
      Begin VB.Menu mnuPopuMenuSerch 
         Caption         =   "����ֵ(&R)"
         Index           =   5
         Shortcut        =   ^R
      End
   End
End
Attribute VB_Name = "frmParameters"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'=========================================================
'============                         ģ�����                   ======================
'=========================================================
Private mrsPars As ADODB.Recordset '������Ϣ��¼��
Private mrsModule As ADODB.Recordset 'ģ����Ϣ��¼��
Private mrsSys As New ADODB.Recordset 'ϵͳ��Ϣ
Private mrsDetailParas As ADODB.Recordset '�����š����û�������������ֵ
Private mlngSys As Long '�ϴε�ѡ��ϵͳ
Private mlngModule As Long '�ϴε�ģ��
Private mlngParID As Long '�ϴεĲ���ID
Private mstrParType As String '�ϴεĲ�������
Private mlngModulePreIdx As Long '�����ָ�����ƥ���ָ�ԭ��ѡ���ģ��
Private mblnNotClick As Boolean '������Click�¼�����ʱ����
Private mblnMultiSta As Boolean '�Ƿ����ö�վ��
Private mstrOwner As String '��ǰϵͳ������

Private Enum ChangeCtrl '�����ı�Ŀؼ�
    CT_Sys = 0
    CT_Module = 1
End Enum

Private Enum mPageNum
    Pag_ParaInfo = 0
    Pag_Computer = 1
    Pag_ChangeLog = 2
End Enum

Private Enum ParaInfoRow '����˵����ö��
    PR_Ӱ�����˵�� = 0
    PR_����ֵ���� = 1
    PR_����˵�� = 2
    PR_����˵�� = 3
    PR_����˵�� = 4
End Enum

'������ǩ
Private Enum mnuIndex
    MI_�û��� = 0
    MI_��Ա = 1
    MI_���� = 2
    MI_������ = 3
    MI_վ�� = 4
    MI_����ֵ = 5
End Enum

Private Const GRD_GOTFOCUS_COLORSEL = &H8000000D '16772055 '    '����ؼ�ʱ,ѡ����ʾ��ɫ
Private Const GRD_LOSTFOCUS_COLORSEL = &H80000010  '�뿪����ʱ,ѡ�����ʾ��ɫ
'=========================================================
'============                         �����ӿ�                   ======================
'=========================================================
Public Function SupportPrint() As Boolean
'���ر������Ƿ�֧�ִ�ӡ���������ڵ���
    SupportPrint = True
End Function

Public Sub SubPrint(ByVal bytMode As Byte)
    '����:bytMode=1 ��ӡ;2 Ԥ��;3 �����EXCEL
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    
    If Me.ActiveControl Is vsDetailParas Then
        objPrint.Title.Text = "��Ժ�����û��Ĳ����嵥��ӡ"
    Else
        objPrint.Title.Text = "�����嵥��ӡ"
    End If
    
    objRow.Add "Ӧ��ϵͳ��" & cboSys.Text
    objPrint.UnderAppRows.Add objRow
    If Me.ActiveControl Is vsDetailParas Then
        Set objRow = New zlTabAppRow
        objRow.Add "�������ͣ�" & vsPara.TextMatrix(vsPara.Row, vsPara.ColIndex("��������"))
        objRow.Add "�������ƣ�" & vsPara.TextMatrix(vsPara.Row, vsPara.ColIndex("������")) & "-" & vsPara.TextMatrix(vsPara.Row, vsPara.ColIndex("������"))
        objPrint.UnderAppRows.Add objRow
        
        Set objRow = New zlTabAppRow
        objRow.Add "����˵����" & vsPara.TextMatrix(vsPara.Row, vsPara.ColIndex("����˵��"))
        objPrint.BelowAppRows.Add objRow
    End If
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡʱ�䣺" & Format(CurrentDate, "yyyy��MM��dd��")
    objPrint.BelowAppRows.Add objRow
    If Me.ActiveControl Is vsDetailParas Then
        Set objPrint.Body = vsDetailParas
    Else
        Set objPrint.Body = vsPara
    End If
    If bytMode = 1 Then
      Select Case zlPrintAsk(objPrint)
          Case 1
               zlPrintOrView1Grd objPrint, 1
          Case 2
              zlPrintOrView1Grd objPrint, 2
          Case 3
              zlPrintOrView1Grd objPrint, 3
      End Select
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub cboModule_LostFocus()
    Dim strKey As String
    strKey = cboModule.Text
    
    If cboModule.ListIndex >= 0 Then
        If cboModule.List(cboModule.ListIndex) = strKey Then
            Exit Sub
        End If
        cboModule.Text = cboModule.List(cboModule.ListIndex)
     Else
        If mlngModulePreIdx >= 0 Then
            cboModule.ListIndex = mlngModulePreIdx
        Else
            cboModule.ListIndex = 0
        End If
    End If
End Sub

'=========================================================
'============                         �ؼ��¼�                   ======================
'=========================================================
Private Sub cboSys_Click()
    '�л�ϵͳ�����ˢ������
    If mblnNotClick Then Exit Sub
    If mlngSys <> cboSys.ItemData(cboSys.ListIndex) Or cboSys.Tag = "ǿ��ˢ��" Then
        mlngSys = cboSys.ItemData(cboSys.ListIndex)
        mrsSys.Filter = "���=" & mlngSys
        If Not mrsSys.EOF Then mstrOwner = mrsSys!������ & ""
        If cboSys.Tag <> "ǿ��ˢ��" Then
            mlngModule = -1
            mstrParType = ""
        End If
        Call GetParasInfo(mlngSys)
        Call LoadParas
        Call ResetCtrl
        Call SetParas
    End If
End Sub

Private Sub cmdAddNew_Click()
    Dim strParsType As String, lngParID As Long
    Dim i As Long, strTmp As String, strDetails As String
    Dim StrValue As String, strUsers As String, strPCs As String
    Dim objfrmParaModiSet As New frmParaModiSet
    
    strParsType = vsPara.TextMatrix(vsPara.Row, vsPara.ColIndex("��������"))
    lngParID = Val(vsPara.RowData(vsPara.Row))
    If lngParID = 0 Then Exit Sub
    If Not objfrmParaModiSet.ShowMe(Me, 1, fraDetaisModi.Tag, "", mstrOwner, lngParID, StrValue, strUsers, strPCs) Then Exit Sub
    
    Call ExecuteProcedure("Zlparameters_Add_Details(" & lngParID & ",'" & UCase(strUsers) & "','" & strPCs & "','" & StrValue & "')", "�����޸Ĳ���ֵ")
    Set mrsDetailParas = Nothing
    'ˢ�²���
    Call LoadDetailParas(lngParID, False)
End Sub

Private Sub cmdDel_Click()
    Dim strParsType As String, lngParID As Long
    Dim i As Long, strTmp As String, strValues As String
    Dim strInfo As String
    
    strParsType = vsPara.TextMatrix(vsPara.Row, vsPara.ColIndex("��������"))
    lngParID = Val(vsPara.RowData(vsPara.Row))
    If lngParID = 0 Then Exit Sub

    With vsDetailParas
        If .Tag > 1 Then
            If MsgBox("�Ƿ�ɾ���Ѿ�ѡ�е�" & .Tag & "����¼��", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            End If
        End If
        If .Tag <> 0 Then '����ѡ����
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, .ColIndex("ѡ��"))) = -1 Then
                    If strParsType = "���Ų���" Then
                        strTmp = .TextMatrix(i, .ColIndex("����ID"))
                        strInfo = """" & .TextMatrix(i, .ColIndex("����")) & """"
                    Else
                        strTmp = .TextMatrix(i, .ColIndex("�û���")) & "^" & .TextMatrix(i, .ColIndex("������"))
                        If .ColHidden(.ColIndex("��Ա")) Then
                            strInfo = """" & .TextMatrix(i, .ColIndex("������")) & """"
                        ElseIf .ColHidden(.ColIndex("������")) Then
                            strInfo = """" & .TextMatrix(i, .ColIndex("��Ա")) & """"
                        Else
                            strInfo = """" & .TextMatrix(i, .ColIndex("��Ա")) & """��""" & .TextMatrix(i, .ColIndex("������")) & """��"
                        End If
                    End If
                    If ActualLen(strValues & "#" & strTmp) >= 2000 Then
                        Call ExecuteProcedure("Zlparameters_Del_Details(" & lngParID & ",'" & strValues & "')", "ɾ����������")
                        strValues = strTmp
                    Else
                        strValues = IIf(strValues = "", strTmp, strValues & "#" & strTmp)
                    End If
                End If
            Next
            If .Tag = 1 Then
                If MsgBox("�Ƿ�ɾ��" & strInfo & "�Ĳ������ã�", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            End If
            If strValues <> "" Then
                Call ExecuteProcedure("Zlparameters_Del_Details(" & lngParID & ",'" & strValues & "')", "ɾ����������")
            End If
        Else 'ֻ����ǰ��
            If .RowData(.Row) <> 0 Then
                If strParsType = "���Ų���" Then
                    strValues = .TextMatrix(.Row, .ColIndex("����ID"))
                    strInfo = """" & .TextMatrix(.Row, .ColIndex("����")) & """"
                Else
                    strValues = .TextMatrix(.Row, .ColIndex("�û���")) & "^" & .TextMatrix(.Row, .ColIndex("������"))
                    If .ColHidden(.ColIndex("��Ա")) Then
                        strInfo = """" & .TextMatrix(.Row, .ColIndex("������")) & """"
                    ElseIf .ColHidden(.ColIndex("������")) Then
                        strInfo = """" & .TextMatrix(.Row, .ColIndex("��Ա")) & """"
                    Else
                        strInfo = """" & .TextMatrix(.Row, .ColIndex("��Ա")) & """��""" & .TextMatrix(.Row, .ColIndex("������")) & """��"
                    End If
                End If
                If MsgBox("�Ƿ�ɾ��" & strInfo & "�Ĳ������ã�", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
                Call ExecuteProcedure("Zlparameters_Del_Details(" & lngParID & ",'" & strValues & "')", "ɾ����������")
            End If
        End If
    End With
    Set mrsDetailParas = Nothing
    'ˢ�²���
    Call LoadDetailParas(lngParID, txtSearch.Text <> "")
End Sub

Private Sub cmdExp_Click()
    Dim strSets As String
    Dim arrTmp As Variant

    Dim rsParas As ADODB.Recordset
    strSets = frmParaInOut.ShowMe(PST_Exp, mlngSys)
    If strSets = "" Then Exit Sub
    arrTmp = Split(strSets, "|")
    Set rsParas = GetALLPars(IIf(Val(arrTmp(1)) = 0, mlngSys, -1), Val(arrTmp(2)) = 0, True)
'    Set rsParas = CopyNewRec(rsParas) '�ı�Ϊ�ɱ��¼��
    If gobjFile.FileExists(arrTmp(0)) Then
        Call gobjFile.DeleteFile(arrTmp(0), True)
    End If
    rsParas.Save arrTmp(0), adPersistXML
    MsgBox "���������ɹ���", vbInformation, gstrSysName
End Sub

Private Sub cmdImp_Click()
    Dim strTmp As String, arrSets As Variant, strValues As String
    Dim rsParas As ADODB.Recordset, rsOldPars As ADODB.Recordset, rsComInfo As ADODB.Recordset
    Dim strSQL As String, arrTmp As Variant, arrCols As Variant, strTmpSQL As String
    Dim i As Long, j As Long
    Dim strPre As String, strCur As String, strMsg As String
    Dim strFilter As String, strFilterEx As String, strFilterOld As String, strFilterTmp As String, strFilterNew As String
    Dim lngSys As Long, blnMultiSys As Boolean, blnDetails As Boolean
    Dim strDeptParas As String, strUserParas As String
    Dim blnTrans As Boolean
    Dim dtStart As Date
    Dim cllErrSQL As Collection '����SQL
    
    On Error GoTo errH
    strTmp = frmParaInOut.ShowMe(PST_Imp, mlngSys)
    If strTmp = "" Then Exit Sub
    dtStart = Now
    arrTmp = Split(strTmp, "|")
    lngSys = IIf(Val(arrTmp(1)) = 0, mlngSys, -1): blnDetails = Val(arrTmp(2)) = 0
    '��ȡ��Ӧ���õ�����
    Set rsOldPars = CopyNewRec(GetALLPars(lngSys, False))  '����ȡ��ϸ�Ĳ������ã���Ϊ���ݵĲ���Ҫ
    Set rsParas = New ADODB.Recordset
    rsParas.Open arrTmp(0), , adOpenStatic, adLockOptimistic, adCmdFile
    '��ȡ��������
    If lngSys <> -1 Then strFilter = "ϵͳ=" & mlngSys: strFilterEx = "ϵͳ<>" & mlngSys 'ֻ���뵱ǰ
    If Not blnDetails Then strFilter = strFilter & IIf(strFilter = "", "", " And ") & " ����<1":  strFilterEx = strFilterEx & IIf(strFilterEx = "", "", " OR  ") & " ����>0"
    strFilter = IIf(strFilter = "", "", strFilter & " And  ����<>-99"):  strFilterEx = IIf(strFilterEx = "", "", strFilterEx & " OR  ����=-99")
    'ϵͳ�汾�ŶԱ�  ���� ������, �汾�� ����ֵ, User ȱʡֵ,To_Char(Sysdate, 'yyyy-mm-dd HH24:mi:ss') Ӱ�����˵��
    rsOldPars.Filter = strFilter & IIf(strFilter = "", "", " And ") & "����=-9": rsParas.Filter = strFilter & IIf(strFilter = "", "", " And ") & "����=-9"
    Set rsComInfo = GetCompareRec(rsOldPars, rsParas, "ϵͳ", "����ֵ", "������")
    Debug.Print "ϵͳ�ȶ�=" & DateDiff("s", dtStart, Now)
    strTmp = "": rsComInfo.Filter = "State=-1" '�����д��ڵ�ϵͳ�����ǵ����ļ���û�и�ϵͳ
    Do While Not rsComInfo.EOF
        strTmp = strTmp & vbNewLine & Lpad(rsComInfo!MainKey, 4) & "-" & RPAD(rsComInfo!������, 18)
        strFilterOld = strFilterOld & " OR ϵͳ=" & Val(rsComInfo!MainKey)
        rsComInfo.MoveNext
    Loop
    If strTmp <> "" Then '�����ļ���û�е�ϵͳ�������бȽ�
        strMsg = "�����ļ�ȱʧ����ϵͳ��" & _
                            strTmp & vbNewLine & _
                        "��Щϵͳ�Ĳ����������е��������"
    End If
    strTmp = "": rsComInfo.Filter = "State=1" '�����в����ڵ�ϵͳ�����ǵ����ļ����и�ϵͳ
    Do While Not rsComInfo.EOF
        strTmp = strTmp & vbNewLine & Lpad(rsComInfo!MainKey, 4) & "-" & RPAD(rsComInfo!������_NEW & "", 18)
        strFilterNew = strFilterNew & " OR ϵͳ=" & Val(rsComInfo!MainKey)
        rsComInfo.MoveNext
    Loop
    strTmp = "": strFilterTmp = "": rsComInfo.Filter = "State=2" '�����ļ������ݿ�����ڸ�ϵͳ
    blnMultiSys = rsComInfo.RecordCount > 1
    Do While Not rsComInfo.EOF
        strFilterTmp = strFilterTmp & " OR ϵͳ=" & Val(rsComInfo!MainKey)
        strTmp = strTmp & vbNewLine & Lpad(rsComInfo!MainKey, 4) & "-" & RPAD(rsComInfo!������, 18) & " ��ǰ�汾:" & VerPAD(rsComInfo!����ֵ) & "<>����汾��" & VerPAD(rsComInfo!����ֵ_New)
        rsComInfo.MoveNext
    Loop
    
    If strTmp <> "" Then
        If blnMultiSys Then  '���ϵͳ
            strMsg = strMsg & IIf(strMsg = "", "", vbNewLine) & _
                            "����ϵͳ��" & _
                                strTmp & vbNewLine & _
                            "�汾���ڲ��죬������ܻ�Ӱ��ϵͳ����ʹ�á��Ƿ�����Щϵͳ�Ĳ�����"
        Else
            strMsg = strMsg & IIf(strMsg = "", "", vbNewLine) & _
                            "ϵͳ��" & Mid(strTmp, 2) & vbNewLine & _
                            "�汾���ڲ��죬������ܻ�Ӱ���ϵͳ����ʹ�á��Ƿ����ϵͳ�Ĳ�����"
        End If
        '�汾�в����ϵͳ��ѯ���Ƿ��롣�����룬�򲻱Ƚ���Щϵͳ
        If MsgBox(strMsg, vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            rsComInfo.Filter = "State=0"
            If rsComInfo.RecordCount = 0 Then 'û�пɵ����ϵͳ
                MsgBox "û�пɵ����ϵͳ��", vbInformation + vbYesNo + vbDefaultButton2, gstrSysName
                Exit Sub
            End If
        Else
            strFilterTmp = ""
        End If
        strMsg = ""
    Else
        rsComInfo.Filter = "State=0"
        If rsComInfo.RecordCount = 0 Then 'û�пɵ����ϵͳ
            MsgBox "û�пɵ����ϵͳ��" & IIf(strMsg = "", "", "����������£�" & vbNewLine) & strMsg, vbInformation + vbYesNo + vbDefaultButton2, gstrSysName
            Exit Sub
        End If
    End If
    strTmp = "": rsComInfo.Filter = IIf(strFilterTmp = "", "State=2 OR State=0", "State=0") '�����ļ������ݿ�����ڸ�ϵͳ
    blnMultiSys = rsComInfo.RecordCount > 1
    Do While Not rsComInfo.EOF
        strTmp = strTmp & vbNewLine & Lpad(rsComInfo!MainKey, 4) & "-" & RPAD(rsComInfo!������_NEW, 18)
        rsComInfo.MoveNext
    Loop
    If strTmp <> "" Then
        If blnMultiSys Then '���ϵͳ
            strMsg = strMsg & IIf(strMsg = "", "", vbNewLine) & _
                            "����ϵͳ���ᱻ���룺" & _
                                strTmp & vbNewLine & _
                            IIf(blnDetails, "��Щϵͳ�Ĳ��š�������˽�в���������գ����µ��롣", "") & "�Ƿ������"
        Else
            strMsg = strMsg & IIf(strMsg = "", "", vbNewLine) & _
                            "ϵͳ��" & Mid(strTmp, 2) & vbNewLine & _
                             "���ᱻ���롣" & IIf(blnDetails, "��ϵͳ���š�������˽�в���������գ����µ��롣", "") & "�Ƿ������"
        End If
        '�ٴ�ѯ���Ƿ�������룬��ֹ�����
        If MsgBox(strMsg, vbInformation + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    End If
    Call ShowFlash("���ڽ��в������룬���Ժ�")
    Debug.Print "ϵͳȷ��=" & DateDiff("s", dtStart, Now)
    'ɾ��������ĵ�
    If strFilterOld <> "" Or strFilterTmp <> "" Or strFilterEx <> "" Then
        Call RecDelete(rsOldPars, strFilterEx & IIf(strFilterEx = "", "", IIf(strFilterNew & strFilterTmp <> "", " OR ", "")) & Mid(strFilterOld & strFilterTmp, Len(" OR ") + 1))
    End If
    If strFilterNew <> "" Or strFilterTmp <> "" Or strFilterEx <> "" Then
        Call RecDelete(rsParas, strFilterEx & IIf(strFilterEx = "", "", IIf(strFilterNew & strFilterTmp <> "", " OR ", "")) & Mid(strFilterNew & strFilterTmp, Len(" OR ") + 1))
    End If
    Debug.Print "ɾ��=" & DateDiff("s", dtStart, Now)
    rsOldPars.Filter = "����=0": rsParas.Filter = "����=0"
    Set rsComInfo = GetCompareRec(rsOldPars, rsParas, "MAINKEY", "-SORTKEY,ϵͳ,ģ��,������,����,����id,�û���,������,��ϸ����ֵ", "", Array("SQL", adVarWChar, 20000, Empty))
    Debug.Print "���ͱ䶯=" & DateDiff("s", dtStart, Now)
    
    Set cllErrSQL = New Collection
    gcnOracle.BeginTrans: blnTrans = True
    On Error Resume Next
    With rsComInfo
        .Filter = "State<>0"
        .Sort = "Sort,MainKey"
        Do While Not .EOF
            strSQL = "": arrTmp = Split(!MainKey, "#")
            '�������Դ���
            Select Case rsComInfo!State
                Case 2 '����
                    strTmpSQL = "Set ": arrCols = Split(!DifInfo & "", ",")
                    For i = LBound(arrCols) To UBound(arrCols)
                        If IsType(.Fields(arrCols(i) & "_New").Type, adVarChar) Then '�ַ������ͣ�����Ҫת��
                            strTmpSQL = strTmpSQL & IIf(i = 0, " ", " , ") & arrCols(i) & " = " & SQLAdjust(.Fields(arrCols(i) & "_New").value)
                        Else
                            strTmpSQL = strTmpSQL & IIf(i = 0, " ", " , ") & arrCols(i) & " = " & Val(.Fields(arrCols(i) & "_New").value & "")
                        End If
                    Next
                    strSQL = "Update zlParameters " & _
                                strTmpSQL & _
                                " Where Nvl(ϵͳ, 0) = " & Val(arrTmp(0)) & " And Nvl(ģ��, 0) = " & Val(arrTmp(1)) & " And ������ = '" & arrTmp(2) & "'"
                Case 1 '����
                    strSQL = "Insert Into zlParameters (ID, ϵͳ, ģ��, ˽��, ����, ��Ȩ, �̶�, ����, ����, ������, ������, ����ֵ, ȱʡֵ,Ӱ�����˵��, ����ֵ����, ����˵��, ����˵��, ����˵��) " & _
                                    "  Select zlParameters_ID.NextVal, " & ZVal(arrTmp(0)) & ", " & ZVal(arrTmp(1)) & ", " & Val(!˽��_New & "") & ", " & Val(!����_New & "") & " , " & Val(!��Ȩ_New & "") & ", " & Val(!�̶�_New & "") & " , " & _
                                    Val(!����_New & "") & " ," & Val(!����_New & "") & " ," & !������_New & " , '" & arrTmp(2) & "', " & SQLAdjust(!����ֵ_New) & " , " & SQLAdjust(!ȱʡֵ_New) & ", " & SQLAdjust(!Ӱ�����˵��_New) & _
                                    "  ," & SQLAdjust(!����ֵ����_New) & "  , " & SQLAdjust(!����˵��_New) & " ," & SQLAdjust(!����˵��_New) & " , " & SQLAdjust(!����˵��_New) & " From Dual" & _
                                    " Where Not Exists (Select 1 From zlParameters Where ������ =" & SQLAdjust(arrTmp(2)) & " And Nvl(ģ��,0) = " & Val(arrTmp(1)) & " And Nvl(ϵͳ,0) = " & Val(arrTmp(0)) & ")"
                Case -1 'ɾ��
                    strSQL = "Delete zlParameters Where Nvl(ϵͳ, 0) = " & Val(arrTmp(0)) & " And Nvl(ģ��, 0) = " & Val(arrTmp(1)) & " And ������ = '" & arrTmp(2) & "'"
            End Select
            strSQL = Replace(Trim(strSQL), ChrW(-3979), "")
            'ChrW(-3979),������chr(63)�ʺţ����ǲ�һ�����ᵼ��WriteLine ��������
            If strSQL <> "" Then
                gcnOracle.Execute strSQL
                If err.Number <> 0 Then
                    err.Clear: gcnOracle.Errors.Clear
                    cllErrSQL.Add strSQL
                End If
            End If
            .MoveNext
        Loop
        '����SQL���������ڲ����Ŵ�λ���µģ�������ظ�ִ��һ��
        For i = 0 To 1
            For j = 1 To cllErrSQL.Count
                If cllErrSQL(j) <> "" Then
                    gcnOracle.Execute cllErrSQL(j)
                    If err.Number <> 0 Then
                        err.Clear: gcnOracle.Errors.Clear
                    Else
                        cllErrSQL(j) = ""
                    End If
                End If
            Next
        Next
        On Error GoTo errH
        Debug.Print "�����б��嵥=" & DateDiff("s", dtStart, Now)
        'ɾ����������
        If blnDetails Then
            .Filter = ""
            strDeptParas = "": strUserParas = ""
            Do While Not .EOF
                If (!���� = 1 Or !˽�� = 1 Or !���� = 1) Or (!����_New = 1 Or !˽��_New = 1 Or !����_New = 1) Then          'ɾ��ԭ���в�����ϸ���ã�����û�еĲ�������
                    If !���� = 1 Or !����_New = 1 Then
                        strTmp = Replace(!MainKey, "#", "^")
                         If ActualLen(strDeptParas & "#" & strTmp) >= 2000 Then
                            strSQL = "Zlparameters_Delall_Details('" & strDeptParas & "',1)"
                            Call ExecuteProcedure(strSQL, "���ԭ���Ĳ�������")
                            strDeptParas = strTmp
                        Else
                            strDeptParas = strDeptParas & IIf(strDeptParas <> "", "#", "") & strTmp
                        End If
                    End If
                    If (!˽�� = 1 Or !���� = 1) Or (!˽��_New = 1 Or !����_New = 1) Then
                        strTmp = Replace(!MainKey, "#", "^")
                         If ActualLen(strUserParas & "#" & strTmp) >= 2000 Then
                            strSQL = "Zlparameters_Delall_Details('" & strUserParas & "',0)"
                            Call ExecuteProcedure(strSQL, "���ԭ���Ĳ�������")
                            strUserParas = strTmp
                        Else
                            strUserParas = strUserParas & IIf(strUserParas <> "", "#", "") & strTmp
                        End If
                    End If
                End If
                .MoveNext
            Loop
            If strDeptParas <> "" Then
                strSQL = "Zlparameters_Delall_Details('" & strDeptParas & "',1)"
                Call ExecuteProcedure(strSQL, "���ԭ���Ĳ�������")
            End If
            If strUserParas <> "" Then
                strSQL = "Zlparameters_Delall_Details('" & strUserParas & "',0)"
                Call ExecuteProcedure(strSQL, "���ԭ���Ĳ�������")
            End If
        End If
    End With
    Debug.Print "���ԭ�в���=" & DateDiff("s", dtStart, Now)
    If blnDetails Then
        '�����µĲ�������
        strPre = "": strCur = "": strValues = ""
        With rsParas
            .Filter = "����>0" '�ִ�Ĳ������鶼����Ҫ�����
            .Sort = "SortKey" '��������
            Do While Not rsParas.EOF
                strCur = rsParas!ϵͳ & "," & rsParas!ģ�� & ",'" & Trim(rsParas!������ & "") & "'"
                If strCur <> strPre Then
                    '������һ�������Ĳ�������
                    If strPre <> "" And strValues <> "^^" And strValues <> "^^#^^" Then
                        strSQL = "Zlparameters_Imp_Details(" & strPre & "," & SQLAdjust(strValues) & ")"
                        Call ExecuteProcedure(strSQL, "�����������")
                        strValues = ""
                    End If
                    strPre = strCur
                End If
                
                If Val(!���� & "") = 1 Then
                    strTmp = rsParas!����id & "^^" & rsParas!��ϸ����ֵ
                Else
                    strTmp = rsParas!�û��� & "^" & rsParas!������ & "^" & rsParas!��ϸ����ֵ
                End If
                If ActualLen(strValues & "#" & strTmp) >= 2000 Then
                    strSQL = "Zlparameters_Imp_Details(" & strPre & "," & SQLAdjust(strValues) & ")"
                    Call ExecuteProcedure(strSQL, "�����������")
                    strValues = strTmp
                Else
                    strValues = IIf(strValues = "", strTmp, strValues & "#" & strTmp)
                End If
                rsParas.MoveNext
            Loop
            If strValues <> "" Then
                strSQL = "Zlparameters_Imp_Details(" & strPre & "," & SQLAdjust(strValues) & ")"
                Call ExecuteProcedure(strSQL, "�����������")
            End If
        End With
    End If
    gcnOracle.CommitTrans:  blnTrans = False
    Debug.Print "����ɹ�=" & DateDiff("s", dtStart, Now)
    ShowFlash ("")
    MsgBox "��������ɹ���", vbInformation, gstrSysName
    Exit Sub
errH:
    If 0 = 1 Then
        Resume
    End If
    If blnTrans Then gcnOracle.RollbackTrans
    ShowFlash ("")
    If blnTrans Then
        MsgBox "��������ʧ�ܣ�������Ϣ���£�" & vbNewLine & gcnOracle.Errors(0).Description, vbInformation, "��������"
    Else
        MsgBox "��������ʧ�ܣ�������Ϣ���£�" & vbNewLine & err.Description, vbInformation, "��������"
    End If
End Sub

Private Sub cmdModValue_Click()
    Dim strParsType As String, lngParID As Long
    Dim i As Long, strTmp As String, strDetails As String
    Dim StrValue As String
    Dim objfrmParaModiSet As New frmParaModiSet
    Dim strInfo As String
    
    strParsType = vsPara.TextMatrix(vsPara.Row, vsPara.ColIndex("��������"))
    lngParID = Val(vsPara.RowData(vsPara.Row))
    If lngParID = 0 Then Exit Sub
    
    With vsDetailParas
        If .Tag > 1 Then
            strInfo = "����ѡ�е�" & .Tag & "���������á�"
            i = .FindRow(-1, , .ColIndex("ѡ��"))
            If i <> -1 Then
                StrValue = .TextMatrix(i, .ColIndex("����ֵ"))
            End If
        Else
            If .Tag = 0 Then
                i = .Row
            Else
                i = .FindRow(-1, , .ColIndex("ѡ��"))
            End If
            If i = -1 Then
                MsgBox "��ǰδѡ���κβ������ã�", vbInformation, gstrSysName
                Exit Sub
            Else
                If strParsType = "���Ų���" Then
                    strInfo = """" & .TextMatrix(i, .ColIndex("����")) & """"
                Else
                    If .ColHidden(.ColIndex("��Ա")) Then
                        strInfo = """" & .TextMatrix(i, .ColIndex("������")) & """"
                    ElseIf .ColHidden(.ColIndex("������")) Then
                        strInfo = """" & .TextMatrix(i, .ColIndex("��Ա")) & """"
                    Else
                        strInfo = """" & .TextMatrix(i, .ColIndex("��Ա")) & """��""" & .TextMatrix(i, .ColIndex("������")) & """��"
                    End If
                End If
                strInfo = "����" & strInfo & "�Ĳ���ֵ��"
                StrValue = .TextMatrix(i, .ColIndex("����ֵ"))
            End If
        End If
        If Not objfrmParaModiSet.ShowMe(Me, 0, fraDetaisModi.Tag, strInfo, mstrOwner, lngParID, StrValue) Then Exit Sub

        If .Tag <> 0 Then
            For i = .FixedRows To .Rows - 1
                If Val(.TextMatrix(i, .ColIndex("ѡ��"))) = -1 Then
                    If strParsType = "���Ų���" Then
                        strTmp = .TextMatrix(i, .ColIndex("����ID")) & "^^" & Trim(StrValue)
                    Else
                        strTmp = .TextMatrix(i, .ColIndex("�û���")) & "^" & .TextMatrix(i, .ColIndex("������")) & "^" & Trim(StrValue)
                    End If
                    If ActualLen(strDetails & "#" & strTmp) >= 2000 Then
                        Call ExecuteProcedure("Zlparameters_Update_Details(" & lngParID & ",'" & strDetails & "')", "�����޸Ĳ���ֵ")
                        strDetails = strTmp
                    Else
                        strDetails = IIf(strDetails = "", strTmp, strDetails & "#" & strTmp)
                    End If
                End If
            Next
            If strDetails <> "" Then
                Call ExecuteProcedure("Zlparameters_Update_Details(" & lngParID & ",'" & strDetails & "')", "�����޸Ĳ���ֵ")
            End If
        Else 'ֻ����ǰ��
            If .RowData(.Row) <> 0 Then
                If strParsType = "���Ų���" Then
                    strDetails = .TextMatrix(.Row, .ColIndex("����ID")) & "^^" & Trim(StrValue)
                Else
                    strDetails = .TextMatrix(.Row, .ColIndex("�û���")) & "^" & .TextMatrix(.Row, .ColIndex("������")) & "^" & Trim(StrValue)
                End If
                Call ExecuteProcedure("Zlparameters_Update_Details(" & lngParID & ",'" & strDetails & "')", "�����޸Ĳ���ֵ")
            End If
        End If
    End With

    Set mrsDetailParas = Nothing
    'ˢ�²���
    Call LoadDetailParas(lngParID, txtSearch.Text <> "")
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim strSQL As String, rsTmp As ADODB.Recordset
    Call ShowFlash("���ڼ��ز�����")
    '�ж��Ƿ����ö�վ��
    strSQL = "Select 1 From zlClients Where վ�� Is Not Null and rownum<2"
    Set rsTmp = gclsBase.OpenSQLRecord(gcnOracle, strSQL, Me.Caption)
    mblnMultiSta = Not rsTmp.EOF
    mlngSys = -1: mlngModule = -1
    Call RestoreVsGridWidth(vsChangeLog, Me.Caption, "�����䶯��־")
    Call RestoreVsGridWidth(vsDetailParas, Me.Caption, "վ�㼰�û�")
    Call RestoreVsGridWidth(vsPara, Me.Caption, "ϵͳ�����б�")
    With vsDetailParas
        .Cell(flexcpPicture, 0, .ColIndex("ѡ��")) = imgEdit.ListImages("UnCheck").Picture
        .Cell(flexcpPictureAlignment, 0, .ColIndex("ѡ��")) = flexAlignCenterCenter
    End With
    Call InitFace '�����ʼ��
    mblnNotClick = True
    Call LoadSystems '����Ӧ��ϵͳ
    mblnNotClick = False
    Call cboSys_Click
    Call vsPara_AfterRowColChange(-1, -1, 1, 1)
    ShowFlash ("")
End Sub
'=========================================================
'============                         ˽�з���                   ======================
'=========================================================
Private Sub InitFace()
'���ܣ���ʼ������
    Dim objItem As TabControlItem
    'ҳ��ؼ�����
    Set objItem = tbPage.InsertItem(Pag_ParaInfo, "����˵����Ϣ", picParInfo.hwnd, 0)
    objItem.Tag = Pag_ParaInfo
    Set objItem = tbPage.InsertItem(Pag_Computer, "Ժ�����û�", picDetailParas.hwnd, 0)
    objItem.Tag = Pag_Computer
    Set objItem = tbPage.InsertItem(Pag_ChangeLog, "�����䶯��־", picParaChangeLog.hwnd, 0)
    objItem.Tag = Pag_ChangeLog
    With tbPage
         tbPage.Item(Pag_ParaInfo).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = True
    End With
     '20�������ַ���ȿ�����ʾϵͳ�����б�����
    If cboSys.Width < 20 * Me.TextWidth("��") Then
        Call CboSetWidth(cboSys.hwnd, 20 * Me.TextWidth("��"))
    End If
    mblnNotClick = True
    cboParType.AddItem "��������": cboParType.ItemData(cboParType.NewIndex) = 0
    cboParType.ListIndex = 0
    
    cboModule.AddItem "���в���"
    cboModule.ItemData(cboModule.NewIndex) = -1
    cboModule.ListIndex = 0
    mblnNotClick = False
End Sub

Private Sub LoadSystems()
'���ܣ�����ϵͳ
    Dim strSQL As String
    Dim strVer As String
    '��ȡ�����߰汾��
    strVer = GetToolsVersion
    '���ӹ����������Ҫ�ǽ���ϵͳ����ǰ��
    Set mrsSys = OpenCursor(gcnOracle, "ZLTOOLS.B_Public.Get_Zlsystems", "")
    If gblnDBA Then
        mrsSys.Filter = ""
    Else
        mrsSys.Filter = "������='" & UCase(gstrUserName) & "'"
    End If
    mrsSys.Sort = "���"
    With mrsSys
        '��ӹ�������ʷ��¼�鿴��
        cboSys.Clear
        cboSys.AddItem String(5, " ") & RPAD("������������", 18) & " v" & VerPAD(strVer)
        cboSys.ItemData(cboSys.NewIndex) = 0
        cboSys.ListIndex = cboSys.NewIndex
        Do While Not .EOF
            cboSys.AddItem Lpad(!���, 4) & "-" & RPAD(!���� & "", 18) & " v" & VerPAD(!�汾�� & "")
            cboSys.ItemData(cboSys.NewIndex) = !���
            .MoveNext
        Loop
        'ȱʡΪ������
        If cboSys.ListCount <= 1 Then
            cboSys.Locked = True
        End If
    End With
End Sub

Private Function GetParasInfo(ByVal lngSys As Long) As Boolean
'���ܣ���ȡ���������Ϣ
    Dim rsTmp As New ADODB.Recordset
    Dim strKey As String, strName As String, strCode As String, arrTmp As Variant
    Dim intTotal As Integer, intFixed As Integer
    Dim int����ȫ�� As Integer, int˽��ȫ�� As Integer, int����ģ�� As Integer, int˽��ģ�� As Integer, int�������� As Integer, int����˽�� As Integer, int���Ų��� As Integer
    Dim strParType As String, blnFixed As Boolean
    On Error GoTo errH
    Set mrsPars = New ADODB.Recordset
    Set mrsModule = New ADODB.Recordset
    'Id, ϵͳ,ģ��,˽��,������,������,����ֵ,ȱʡֵ,Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵��,����,��Ȩ,�̶�,����,ģ������,ģ�����
    Set rsTmp = OpenCursor(gcnOracle, "ZLTOOLS.B_Runmana.Get_Parameters", lngSys)
    rsTmp.Filter = "����=0" '����ʾ�����ڲ�������ʹ�õĲ���
    Set mrsPars = CopyNewRec(rsTmp, , , Array("Fixed", adInteger, 1, 0, "ParType", adVarChar, 50, Empty))
    Set mrsModule = CopyNewRec(rsTmp, True, "ϵͳ,ģ��,ģ������,ģ�����", Array("Key", adVarChar, 200, Empty, "Fixed", adInteger, 3, 0, "Total", adInteger, 5, 0, _
                                                                                                                                "����ȫ��", adInteger, 3, 0, "˽��ȫ��", adInteger, 3, 0, _
                                                                                                                                "����ģ��", adInteger, 3, 0, "˽��ģ��", adInteger, 3, 0, _
                                                                                                                                "��������ģ��", adInteger, 3, 0, "����˽��ģ��", adInteger, 3, 0, _
                                                                                                                                "���Ų���", adInteger, 3, 0, "Index", adInteger, 5, 0))
    mrsPars.Filter = ""
    mrsPars.Sort = "ϵͳ,ģ��,������"
    With mrsPars
        Do While Not mrsPars.EOF
            If strKey <> !ϵͳ & "_" & !ģ�� Then
                If strKey <> "" Then
                    arrTmp = Split(strKey, "_")
                    mrsModule.AddNew Array("Key", "ϵͳ", "ģ��", "ģ������", "ģ�����", "Fixed", "Total", "����ȫ��", "˽��ȫ��", "����ģ��", "˽��ģ��", "��������ģ��", "����˽��ģ��", "���Ų���"), _
                                                    Array(strKey, Val(arrTmp(0)), Val(arrTmp(1)), strName, strCode, intFixed, intTotal, int����ȫ��, int˽��ȫ��, int����ģ��, int˽��ģ��, int��������, int����˽��, int���Ų���)
                End If
                strKey = !ϵͳ & "_" & !ģ��
                strName = !ģ������ & "": strCode = !ģ����� & ""
                intFixed = 0: intTotal = 0
                int����ȫ�� = 0: int˽��ȫ�� = 0: int����ģ�� = 0: int˽��ģ�� = 0: int�������� = 0: int����˽�� = 0
            End If
            strParType = GetParaType(Val(Nvl(!ģ��)), Val(Nvl(!˽��)), Val(Nvl(!����)), Val(Nvl(!����)))
            Select Case strParType
                Case "����ȫ��"
                    int����ȫ�� = int����ȫ�� + 1
                Case "˽��ȫ��"
                    int˽��ȫ�� = int˽��ȫ�� + 1
                Case "����ģ��"
                    int����ģ�� = int����ģ�� + 1
                Case "˽��ģ��"
                    int˽��ģ�� = int˽��ģ�� + 1
                Case "��������ģ��"
                    int�������� = int�������� + 1
                Case "����˽��ģ��"
                    int����˽�� = int����˽�� + 1
                Case "���Ų���"
                    int���Ų��� = int���Ų��� + 1
            End Select
            intTotal = intTotal + 1 '�ܼ���+1
            '���ܵ����������͵Ĳ���
            If !�̶� = 1 Or strParType = "����ȫ��" Or strParType = "˽��ȫ��" Then
                intFixed = intFixed + 1: blnFixed = True
            Else
                blnFixed = False
            End If
            .Update Array("Fixed", "ParType"), Array(IIf(blnFixed, 1, 0), strParType)
            .MoveNext
        Loop
        If strKey <> "" Then
            arrTmp = Split(strKey, "_")
            mrsModule.AddNew Array("Key", "ϵͳ", "ģ��", "ģ������", "ģ�����", "Fixed", "Total", "����ȫ��", "˽��ȫ��", "����ģ��", "˽��ģ��", "��������ģ��", "����˽��ģ��", "���Ų���"), _
                                            Array(strKey, Val(arrTmp(0)), Val(arrTmp(1)), strName, strCode, intFixed, intTotal, int����ȫ��, int˽��ȫ��, int����ģ��, int˽��ģ��, int��������, int����˽��, int���Ų���)
        End If
    End With
    Exit Function
errH:
    If 0 = 1 Then
        Resume
    End If
    MsgBox err.Description, vbInformation, gstrSysName
End Function

Private Sub ResetCtrl(Optional ByVal ctInput As ChangeCtrl = CT_Sys)
'�������͡�����ģ�飬�̶����������໥���ã��Լ�
'������intChangeCtrl�������ı�Ŀؼ�
    Dim strParTypes As String, strOldFilter As String
    Dim arrParType As Variant, i As Long, blnMuiltRow As Boolean
    Dim blnShowFixed As Boolean
    
    With mrsModule
        If ctInput = CT_Sys Then
            mblnNotClick = True
            .Filter = ""
            .Sort = "ģ��"
            cboModule.Clear
            If .RecordCount <> 1 Then
                mrsPars.Filter = "": mrsPars.Sort = "ģ��,������,ID"
                cboModule.AddItem "���в���" & "(" & mrsPars.RecordCount & ")"
                cboModule.ItemData(cboModule.NewIndex) = -1
                If .RecordCount = 0 Then
                    cboModule.ListIndex = cboModule.NewIndex
                    Call ResetCtrl(CT_Module)
                    mblnNotClick = False
                    Exit Sub
                End If
            End If
            '����ģ��
            Do While Not .EOF
                If Val(Nvl(!ģ��)) = 0 Then
                    cboModule.AddItem "ϵͳ����" & "(" & !Total & ")"
                    cboModule.ItemData(cboModule.NewIndex) = Nvl(!ģ��)
                Else
                    cboModule.AddItem Nvl(!ģ��) & "-" & Nvl(!ģ������) & "(" & !Total & ")"
                    cboModule.ItemData(cboModule.NewIndex) = Nvl(!ģ��)
                End If
                mrsModule.Update "Index", cboModule.NewIndex '��¼����
                If mlngModule = Val(Nvl(!ģ��)) Then cboModule.ListIndex = cboModule.NewIndex
                .MoveNext
            Loop
            If cboModule.ListIndex < 0 Then cboModule.ListIndex = 0
            Call ResetCtrl(CT_Module)
            mblnNotClick = False
        ElseIf ctInput = CT_Module Then
            chkShowFixed.value = 0
            If cboModule.ItemData(cboModule.ListIndex) = -1 Then
                .Filter = ""
            Else
                .Filter = "ģ��=" & cboModule.ItemData(cboModule.ListIndex)
            End If
            .Sort = "ģ��"
            If .RecordCount = 0 Then
                strParTypes = ""
            Else
                blnMuiltRow = .RecordCount > 1
                strOldFilter = .Filter
                If strOldFilter = "0" Then strOldFilter = ""
                If Not blnMuiltRow Then
                    blnShowFixed = Val(!Fixed) <> 0
                Else
                    .Filter = IIf(strOldFilter <> "", strOldFilter & " And ", "") & "Fixed<>0"
                    blnShowFixed = .RecordCount <> 0
                    .Filter = strOldFilter
                End If
                arrParType = Array("����ȫ��", "˽��ȫ��", "����ģ��", "˽��ģ��", "��������ģ��", "����˽��ģ��", "���Ų���")
                For i = LBound(arrParType) To UBound(arrParType)
                    If blnMuiltRow Then
                        .Filter = IIf(strOldFilter <> "" And strOldFilter <> "0", strOldFilter & " And ", "") & arrParType(i) & "<>0"
                        If .RecordCount <> 0 Then
                            strParTypes = strParTypes & "," & arrParType(i)
                        End If
                    Else
                        If Val(.Fields(arrParType(i))) <> 0 Then
                            strParTypes = strParTypes & "," & arrParType(i)
                        End If
                    End If
                Next
                chkShowFixed.Visible = blnShowFixed
                If blnShowFixed And Not blnMuiltRow Then 'ȫ���ǹ̶���������̶�������ʾ��ѡ
                    If !Fixed = !Total Then
                        chkShowFixed.value = 1
                    End If
                End If
                '���в�������,"����ȫ��","˽��ȫ��","����ģ��","˽��ģ��","��������ģ��","����˽��ģ��", "���Ų���"
                arrParType = Split(strParTypes, ",")
                cboParType.Clear
                '��ʽΪ,���ͣ�ֻ��һ�ֿ�ѡ���ͣ���û�����в�������
                If UBound(arrParType) - LBound(arrParType) + 1 = 2 Then
                    cboParType.AddItem arrParType(UBound(arrParType)): cboParType.ItemData(cboParType.NewIndex) = 0
                    cboParType.ListIndex = 0
                    Exit Sub
                End If
                For i = LBound(arrParType) To UBound(arrParType)
                    If arrParType(i) = "" Then arrParType(i) = "��������"
                    cboParType.AddItem arrParType(i): cboParType.ItemData(cboParType.NewIndex) = i
                    If arrParType(i) = mstrParType Then cboParType.ListIndex = cboParType.NewIndex
                Next
                If cboParType.ListIndex < 0 Then cboParType.ListIndex = 0
            End If
        End If
    End With
End Sub

Private Function LoadParas() As Boolean
    Dim i As Long
    'Id, ϵͳ,ģ��,˽��,������,������,����ֵ,ȱʡֵ,Ӱ�����˵��,����ֵ����,����˵��,����˵��,����˵��,����,��Ȩ,�̶�,ģ������,ģ�����
    With vsPara
        .Redraw = flexRDNone
        .Rows = 2: .Clear 1
        .RowData(1) = 0: .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
        .Cell(flexcpPicture, 1, .ColIndex("��־")) = Nothing
        mrsPars.Filter = "": mrsPars.Sort = "ģ��,������,ID"
        .Rows = IIf(mrsPars.RecordCount = 0, 1, mrsPars.RecordCount) + 1
        For i = 1 To mrsPars.RecordCount
            .RowData(i) = Nvl(mrsPars!Id)
             ' ˽��,����,��Ȩ,�̶�,ģ������,ģ�����
            .TextMatrix(i, .ColIndex("��������")) = mrsPars!ParType
            .TextMatrix(i, .ColIndex("ģ������")) = Nvl(mrsPars!ģ������)
            .TextMatrix(i, .ColIndex("������")) = Nvl(mrsPars!������)
            .TextMatrix(i, .ColIndex("������")) = Nvl(mrsPars!������)
            .TextMatrix(i, .ColIndex("����ֵ")) = Nvl(mrsPars!����ֵ)
            .TextMatrix(i, .ColIndex("��Ȩ")) = IIf(Val(Nvl(mrsPars!��Ȩ)) = 1, "��", "")
            .TextMatrix(i, .ColIndex("ȱʡֵ")) = Nvl(mrsPars!ȱʡֵ)
            .TextMatrix(i, .ColIndex("Ӱ�����˵��")) = Nvl(mrsPars!Ӱ�����˵��)
            .TextMatrix(i, .ColIndex("����")) = Nvl(mrsPars!����)
            .TextMatrix(i, .ColIndex("Fixed")) = Nvl(mrsPars!Fixed)
            .TextMatrix(i, .ColIndex("����ֵ����")) = Nvl(mrsPars!����ֵ����)
            .TextMatrix(i, .ColIndex("����˵��")) = Nvl(mrsPars!����˵��)
            .TextMatrix(i, .ColIndex("����˵��")) = Nvl(mrsPars!����˵��)
            .TextMatrix(i, .ColIndex("����˵��")) = Nvl(mrsPars!����˵��)
            .TextMatrix(i, .ColIndex("ģ��")) = Nvl(mrsPars!ģ��)
            .TextMatrix(i, .ColIndex("ģ�����")) = Nvl(mrsPars!ģ�����)
            If mlngParID = Val(Nvl(mrsPars!Id)) Then .Row = i: .TopRow = .Row
            Set .Cell(flexcpPicture, i, .ColIndex("��־")) = imgList.ListImages(mrsPars!ParType & "").Picture
            .Cell(flexcpPictureAlignment, i, .ColIndex("��־")) = 4
            If Val(mrsPars!Fixed & "") = 1 Then
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &H80000011 ' &H8000000F
            Else
                .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &H80000012
            End If
            mrsPars.MoveNext
        Next
        .Redraw = flexRDBuffered
    End With
    LoadParas = True
End Function

Private Sub chkShowFixed_Click()
    Call SetParas
End Sub

Private Sub chkShowFixed_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub cboModule_Click()
    mlngModulePreIdx = cboModule.ListIndex
    If mblnNotClick Then Exit Sub
    If mlngModule <> cboModule.ItemData(cboModule.ListIndex) Then
        Call ResetCtrl(CT_Module)
    End If
    Call SetParas
    mlngModule = cboModule.ItemData(cboModule.ListIndex)
End Sub

Private Sub SetParas()
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʾ��������صĲ�����Ϣ
    '����:���˺�
    '����:2009-02-19 12:05:34
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim lngModule As Long
    Dim arrData As Variant
    Dim j As Long, lngRow As Long, blnOldRowShow As Boolean
    Dim strParType As String, blnShow As Boolean
    
    
    strParType = cboParType.List(cboParType.ListIndex)
    With vsPara
        .Redraw = flexRDNone
        If cboModule.ListIndex < 0 Then
            cboModule.ListIndex = 0: Exit Sub
        Else
            lngModule = cboModule.ItemData(cboModule.ListIndex)
        End If
        'չʾ������
        For j = 0 To .Cols - 1
           .ColHidden(j) = False
           If j = .ColIndex("ģ������") Then
               .ColHidden(j) = lngModule >= 0
           ElseIf j = .ColIndex("��������") Then
               .ColHidden(j) = strParType <> "��������"
            ElseIf j >= .ColIndex("Ӱ�����˵��") Then
                .ColHidden(j) = True
           End If
        Next
        lngRow = -1
        For i = 1 To .Rows - 1
           blnShow = True '
            If chkShowFixed.value = 0 And Val(.TextMatrix(i, .ColIndex("Fixed"))) = 1 Then
                '����ʾ�̶�����
                blnShow = False
            End If
            If lngModule > -1 Then 'ֻ��ʾ��ǰģ��
                If Val(.TextMatrix(i, .ColIndex("ģ��"))) <> lngModule Then
                    blnShow = False
                End If
            End If
            'ֻ��ʾ��ǰ����
            If strParType <> "��������" Then
                If Trim(.TextMatrix(i, .ColIndex("��������"))) <> strParType Then
                    blnShow = False
                End If
            End If
            .RowHidden(i) = Not blnShow
            If lngRow <= 0 And .RowHidden(i) = False Then lngRow = i
        Next
        If lngRow > 0 Then
           If .RowHidden(.Row) = True Then .Row = lngRow
           If .Row = 0 Then .Row = lngRow
        End If
        If .RowHidden(.Row) Then vsPara.Row = 0
        .Redraw = flexRDBuffered
    End With
End Sub

Private Sub SetDetailPara(ByVal strParType As String)
'���ܣ�չʾ�ֲ�����
'������strParType=��������
    Dim i As Long, lngCol As Long, intDefIdx As Integer
    With vsDetailParas
        'չʾ������
        For i = 0 To .Cols - 1
            .ColHidden(i) = False
            If i = .ColIndex("������") Then '�Ǳ����������ػ�������
                .ColHidden(i) = Not strParType Like "*����*"
            ElseIf i = .ColIndex("��Ա") Or i = .ColIndex("�û���") Then '���Ų����뱾������ģ��������Ա
                .ColHidden(i) = Not strParType Like "*˽��*" Or strParType = "���Ų���"
            ElseIf i = .ColIndex("վ��") Then '�����ڶ�վ�㣬����ʾվ����
                .ColHidden(i) = Not mblnMultiSta
            ElseIf i >= .ColIndex("��Աid") Then
                .ColHidden(i) = True
            End If
        Next
        If lblSearch.Tag = "" Then
            intDefIdx = MI_���� 'Ĭ�ϰ�������������Ϊ������һֱ�ɼ�
            If strParType = "��������ģ��" Then
                intDefIdx = MI_������
            ElseIf strParType = "˽��ȫ��" Or strParType = "˽��ģ��" Or strParType = "����˽��ģ��" Then
                intDefIdx = MI_��Ա
            End If
        Else
            intDefIdx = Val(lblSearch.Tag)
        End If
        '������������
        For i = 0 To MI_����ֵ
            lngCol = Decode(i, MI_�û���, .ColIndex("�û���"), MI_��Ա, .ColIndex("��Ա"), MI_����, .ColIndex("����"), _
                                        MI_������, .ColIndex("������"), MI_վ��, .ColIndex("վ��"), MI_����ֵ, .ColIndex("����ֵ"))
            mnuPopuMenuSerch(i).Enabled = Not .ColHidden(lngCol)
            mnuPopuMenuSerch(i).Visible = Not .ColHidden(lngCol)
            mnuPopuMenuSerch(i).Checked = i = intDefIdx
        Next
        mnuPopuMenuSerch_Click (intDefIdx)
    End With
End Sub

Private Sub cboModule_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim strKey As String
    Dim i As Long
    If KeyCode <> vbKeyReturn Then Exit Sub
    strKey = Replace(cboModule.Text, "'", "")
    i = cboModule.ListIndex
    
    If cboModule.ListIndex >= 0 Then
        If cboModule.List(cboModule.ListIndex) = strKey Then
            SendKeys "{tab}"
            cboModule.ListIndex = i
            Exit Sub
        End If
    End If
    If strKey = "" Then SendKeys "{tab}": Exit Sub
    If ShowSelect(strKey) = False Then
        cboModule.SetFocus
        cboModule.ListIndex = i
        Exit Sub
    End If
End Sub

Private Function ShowSelect(ByVal strKey As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:ѡ����Ӧ������
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-02-20 16:50:12
    '-----------------------------------------------------------------------------------------------------------
  
    Dim lngLeft As Long, lngTop As Long, i As Long
    
    Dim vRect  As RECT
    Dim strSelect As String
    Dim sngHight As Single
    If mrsModule Is Nothing Then Exit Function
    
    
    mrsModule.Filter = 0
    mrsModule.Filter = "ģ��=" & IIf(Val(strKey) = 0, -22, Val(strKey)) & " Or ģ������ like '%" & strKey & "%' or ģ����� like '%" & UCase(strKey) & "%' "
    If mrsModule.RecordCount = 0 Then
        MsgBox "ע��:" & vbCrLf & _
               "    û���ҵ�����������ģ��,���飡", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
        If cboModule.Visible Then cboModule.SetFocus
        mrsModule.Filter = 0
        Exit Function
    End If
    If mrsModule.RecordCount = 1 Then GoTo SelOk:
    vRect = GetControlRect(cboModule.hwnd)
    sngHight = (IIf(mrsModule.RecordCount <= 2, 5, mrsModule.RecordCount) + 1) * 300
    If sngHight > Screen.Height - (vRect.Top + txtSearch.Height) Then
       If sngHight > vRect.Top Then
          sngHight = vRect.Top
          vRect.Top = 0
       Else
          vRect.Top = vRect.Top - sngHight
       End If
    Else
        vRect.Top = vRect.Top + cboModule.Height
    End If
    If frmSelectList.ShowSelect(Nothing, mrsModule, "ģ��,800,0,1;ģ������,2400,0,1;ģ�����,1440,0,0", vRect.Left, vRect.Top, cboModule.Width * 2, sngHight, "", "ϵͳģ��", , strSelect, True) = False Then
        mrsModule.Filter = 0
        Exit Function
    End If
    If mrsModule.EOF Then
        mrsModule.Filter = 0
        Exit Function
    End If
SelOk:
    cboModule.ListIndex = mrsModule!Index
    mrsModule.Filter = 0
    ShowSelect = True
End Function

Private Sub cboSys_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub cboParType_Click()
    If mblnNotClick Then Exit Sub
    If (cboParType.Text = "����ȫ��" Or cboParType.Text = "˽��ȫ��") Then
        chkShowFixed.value = 1
    ElseIf cboParType.Text = "��������" And chkShowFixed.value = 0 And chkShowFixed.Visible Then
        chkShowFixed.value = 1
    Else
        chkShowFixed.value = 0
    End If
    Call SetParas
End Sub

Private Sub cboParType_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{tab}"
End Sub

Private Sub cmdModify_Click()
    '�Ƚ��е�½,
    Dim strUserName As String
    Dim strSystem As String, lng����id As Long
    
    If cboSys.ListIndex < 0 Then Exit Sub
    strSystem = cboSys.ItemData(cboSys.ListIndex)
    With vsPara
        lng����id = .RowData(.Row)
        If lng����id = 0 Then Exit Sub
        If Val(vsPara.TextMatrix(vsPara.Row, vsPara.ColIndex("Fixed"))) = 1 Then Exit Sub
    End With
    If frmUserCheckLogin.ShowLogin(UCT_NormalUser, , strUserName, gstrServer, strSystem) = False Then Exit Sub
    If frmParaChangeSet.ShowEdit(Me, lng����id, strUserName) = False Then Exit Sub
    mlngParID = lng����id
    If cboModule.ListIndex <> 0 Then
        mlngModule = cboModule.ItemData(cboModule.ListIndex)
    End If
    mstrParType = cboParType.Text
    '��Ҫ�������õ�ǰ�еĲ�����Ϣ
    cboSys.Tag = "ǿ��ˢ��"
    Call cboSys_Click
    mlngParID = Val(vsPara.RowData(vsPara.Row))
    If cboModule.ListIndex <> 0 Then
        mlngModule = cboModule.ItemData(cboModule.ListIndex)
    End If
    mstrParType = cboParType.Text
    cboSys.Tag = ""
End Sub
 
Private Sub cmdSearch_Click()
    Dim rsTmp As ADODB.Recordset
    Dim strDistinct As String, strCols As String, strFields As String
    Dim lngLeft As Long, lngTop As Long
    Dim vRect  As RECT
    Dim strSelect As String
    Dim sngHight As Single, sgnWidth As Single
    
    Select Case Val(lblSearch.Tag)
        Case MI_�û���, MI_��Ա
            strDistinct = "�û���,��Ա,��Ա����"
            strCols = "�û���,1000,0,1;��Ա,1500,0,1;��Ա����,1000,0,1"
            strFields = IIf(Val(lblSearch.Tag) = MI_�û���, "�û���", "��Ա")
            sgnWidth = 3530
        Case MI_����
            strDistinct = "����,���ż���"
            strCols = "����,1200,0,1;���ż���,1000,0,1"
            strFields = "����"
            sgnWidth = 2230
        Case MI_������
            strDistinct = "������,����������"
            strCols = "������,2000,0,1;����������,1200,0,1"
            strFields = "������"
            sgnWidth = 3230
        Case MI_վ��
            strDistinct = "վ��"
            strCols = "վ��,2000,0,1"
            strFields = "վ��"
            sgnWidth = 2030
        Case MI_����ֵ
            strDistinct = "����ֵ"
            strCols = "����ֵ,2000,0,1"
            strFields = "����ֵ"
            sgnWidth = 2030
    End Select
    mrsDetailParas.Filter = ""
    Set rsTmp = RecDistinct(mrsDetailParas, strDistinct)
    If rsTmp.RecordCount = 0 Then
        MsgBox "ע��:" & vbCrLf & _
               "    �ò�������ص��û����������Ų�������,���飡", vbInformation + vbDefaultButton1 + vbOKOnly, gstrSysName
        If txtSearch.Visible Then txtSearch.SetFocus
        Exit Sub
    End If
    
    vRect = GetControlRect(txtSearch.hwnd)
    sngHight = (IIf(rsTmp.RecordCount <= 7, 7, rsTmp.RecordCount) + 1) * 300
    If sngHight > Screen.Height - (vRect.Top + txtSearch.Height) Then
       If sngHight > vRect.Top Then
          sngHight = vRect.Top
          vRect.Top = 0
       Else
          vRect.Top = vRect.Top - sngHight
       End If
    Else
        vRect.Top = vRect.Top + txtSearch.Height
    End If
    If sgnWidth > Screen.Width - vRect.Left Then
        sgnWidth = Screen.Width - vRect.Left
    End If
    If frmSelectList.ShowSelect(Nothing, rsTmp, strCols, vRect.Left, vRect.Top, sgnWidth, sngHight, "", strFields & "ѡ��", , strSelect, True) = False Then Exit Sub
    If rsTmp.EOF Then Exit Sub
    txtSearch.Text = Nvl(rsTmp.Fields(strFields).value)
    If txtSearch.Visible Then txtSearch.SetFocus
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If Not mrsModule Is Nothing Then
        If mrsModule.State = 1 Then
            mrsModule.Filter = 0
            mrsModule.Close
        End If
    End If
    If Not mrsDetailParas Is Nothing Then
        If mrsDetailParas.State = 1 Then mrsDetailParas.Close
    End If
    Set mrsModule = Nothing
    Set mrsDetailParas = Nothing
    Call SaveVsGridWidth(vsChangeLog, Me.Caption, "�����䶯��־")
    Call SaveVsGridWidth(vsDetailParas, Me.Caption, "��վ�㼰�û�")
    Call SaveVsGridWidth(vsPara, Me.Caption, "ϵͳ�����б�")
    
End Sub

Private Function LoadDetailParas(ByVal lng����id As Long, Optional blnSearch As Boolean = False) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����ָ���������û�������Ϣ
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-02-17 14:58:37
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim int���� As Integer, int���� As Integer, int˽�� As Integer, strOwner As String
    
    If blnSearch And mrsDetailParas Is Nothing Then
        mstrOwner = ""
        Set mrsDetailParas = GetDetailParas(lng����id, mrsSys, int����, int����, int˽��, mstrOwner)
        fraDetaisModi.Tag = int���� & "," & int˽�� & "," & int����
    End If
    If blnSearch = False Then
        mstrOwner = ""
        Set mrsDetailParas = GetDetailParas(lng����id, mrsSys, int����, int����, int˽��, mstrOwner)
        fraDetaisModi.Tag = int���� & "," & int˽�� & "," & int����
        If UCase(txtSearch.Text) <> "" Then blnSearch = True
    End If
    cmdAddNew.Visible = int���� = 0
    cmdDel.Visible = int���� = 0
    If blnSearch Then
        If UCase(txtSearch.Text) = "" Then
            mrsDetailParas.Filter = 0
        Else
            Select Case Val(lblSearch.Tag)
                Case MI_�û���, MI_��Ա
                    If Val(lblSearch.Tag) = MI_�û��� Then
                        mrsDetailParas.Filter = "�û��� like '" & UCase(txtSearch.Text) & "%'"
                    Else
                        mrsDetailParas.Filter = "��Ա like '" & UCase(txtSearch.Text) & "%' OR ��Ա���� like '" & UCase(txtSearch.Text) & "%'"
                    End If
                Case MI_����
                    mrsDetailParas.Filter = "���� like '" & UCase(txtSearch.Text) & "%' OR ���ż��� like '" & UCase(txtSearch.Text) & "%'"
                Case MI_������
                    mrsDetailParas.Filter = "������ like '" & UCase(txtSearch.Text) & "%' or ���������� like '" & UCase(txtSearch.Text) & "%'"
                Case MI_վ��
                    mrsDetailParas.Filter = "վ�� like '" & UCase(txtSearch.Text) & "%'"
                Case MI_����ֵ
                    mrsDetailParas.Filter = "����ֵ like '" & UCase(txtSearch.Text) & "%'"
            End Select
        End If
    End If
    mrsDetailParas.Sort = "վ��,����,��Ա,������"
    With vsDetailParas
        .Redraw = flexRDNone
        .Rows = 2: .Clear 1
        .RowData(1) = ""
        .Cell(flexcpPicture, 0, .ColIndex("ѡ��")) = imgEdit.ListImages("UnCheck").Picture
        .ColData(.ColIndex("ѡ��")) = 0
        .Cell(flexcpPictureAlignment, 0, .ColIndex("ѡ��")) = flexAlignCenterCenter
        .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
        .Rows = IIf(mrsDetailParas.RecordCount = 0, 1, mrsDetailParas.RecordCount) + 1
        .Tag = 0 '��¼ѡ������
        .RowData(0) = mrsDetailParas.RecordCount '��¼������
        i = 1
        Do While Not mrsDetailParas.EOF  '
            .RowData(i) = Nvl(mrsDetailParas!����id)
            .TextMatrix(i, .ColIndex("վ��")) = Nvl(mrsDetailParas!վ��)
            .TextMatrix(i, .ColIndex("����id")) = Nvl(mrsDetailParas!����id)
            .TextMatrix(i, .ColIndex("����")) = Nvl(mrsDetailParas!����)
            .TextMatrix(i, .ColIndex("���ż���")) = Nvl(mrsDetailParas!���ż���)
            .TextMatrix(i, .ColIndex("�û���")) = Nvl(mrsDetailParas!�û���)
            .TextMatrix(i, .ColIndex("��Աid")) = Nvl(mrsDetailParas!��Աid)
            .TextMatrix(i, .ColIndex("��Ա")) = Nvl(mrsDetailParas!��Ա)
            .TextMatrix(i, .ColIndex("��Ա����")) = Nvl(mrsDetailParas!��Ա����)
            .TextMatrix(i, .ColIndex("������")) = Nvl(mrsDetailParas!������)
            .TextMatrix(i, .ColIndex("����������")) = Nvl(mrsDetailParas!����������)
            .TextMatrix(i, .ColIndex("����ֵ")) = Nvl(mrsDetailParas!����ֵ)
            i = i + 1
            mrsDetailParas.MoveNext
        Loop
        .Redraw = flexRDBuffered
    End With
    Call SetModValue
    LoadDetailParas = True
End Function

Private Function LoadChangeLog(ByVal lng����id As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����ָ���������û�������Ϣ
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-02-17 14:58:37
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim i As Long
     
    Set rsTemp = OpenCursor(gcnOracle, "ZLTOOLS.B_Runmana.Get_Parachangedlog", lng����id)
    '����id, ���, �䶯˵��, �䶯����, �䶯��, �䶯ʱ��, �䶯ԭ��
    With vsChangeLog
        .Redraw = flexRDNone
        .Rows = 2: .Clear 1
        .RowData(1) = ""
        .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        i = 1
        Do While Not rsTemp.EOF  '
            .RowData(i) = Nvl(rsTemp!����id)
            .TextMatrix(i, .ColIndex("���")) = Nvl(rsTemp!���)
            .TextMatrix(i, .ColIndex("�䶯˵��")) = Nvl(rsTemp!�䶯˵��)
            .TextMatrix(i, .ColIndex("�䶯����")) = Nvl(rsTemp!�䶯����)
            .TextMatrix(i, .ColIndex("�䶯��")) = Nvl(rsTemp!�䶯��)
            .TextMatrix(i, .ColIndex("�䶯ʱ��")) = Format(rsTemp!�䶯ʱ��, "yyyy-mm-dd HH:MM:SS")
            .TextMatrix(i, .ColIndex("�䶯ԭ��")) = Nvl(rsTemp!�䶯ԭ��)
            i = i + 1
            rsTemp.MoveNext
        Loop
        .Redraw = flexRDBuffered
    End With
    LoadChangeLog = True
End Function

Private Function GetParaType(ByVal lngģ�� As Long, ByVal int˽�� As Integer, ByVal int���� As Integer, ByVal int���� As Integer) As String
    '-----------------------------------------------------------------------------------------------------------
    '����:��ȡ��������
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-02-17 16:44:21
    '-----------------------------------------------------------------------------------------------------------
    If int���� = 1 Then GetParaType = "���Ų���": Exit Function
    If lngģ�� = 0 Then
        '����ģ��,֤��ֻ����������:����ȫ�ֺ�˽��ȫ��
        GetParaType = IIf(int˽�� = 0, "����ȫ��", "˽��ȫ��")
        Exit Function
    End If
    '��ģ��Ĵ���
    If int���� = 0 Then
        '���Ǳ��������,ֻ����������:����ģ���˽��ģ��
         GetParaType = IIf(int˽�� = 0, "����ģ��", "˽��ģ��")
         Exit Function
    End If
    '�Ա�����ģ����д���Ҳ���������:
    GetParaType = IIf(int˽�� = 0, "��������ģ��", "����˽��ģ��")
End Function

Private Sub Form_Resize()
    On Error Resume Next
    With imgMain
        .Top = 700
        .Left = ScaleLeft + 200
    End With
    If fraSplit.Tag = "" Then 'Ĭ����������
        fraSplit.Top = (Me.ScaleHeight - picPara.Top) * 0.6 + picPara.Top
    End If
    fraSplit.Width = Me.ScaleWidth - fraSplit.Left + 100
    picPara.Height = fraSplit.Top - picPara.Top - 30
    picPara.Width = Me.ScaleWidth - picPara.Left
    picPage.Top = fraSplit.Top + fraSplit.Height + 30
    picPage.Height = Me.ScaleHeight - picPage.Top
    picPage.Width = Me.ScaleWidth - picPage.Left
End Sub

Private Sub fraSplit_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then fraSplit.Top = fraSplit.Top + y
End Sub

Private Sub fraSplit_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If fraSplit.Top - picPara.Top < 1000 Then fraSplit.Top = picPara.Top + 1000
    If fraSplit.Top > picPage.Height + picPage.Top - 1500 Then fraSplit.Top = picPage.Height + picPage.Top - 1500
    fraSplit.Tag = "�϶�"
    Call Form_Resize
End Sub

Private Sub lblSearch_Click()
    Dim i As Long
    '������������
    For i = 0 To MI_����ֵ '���¹�ѡ״̬
        mnuPopuMenuSerch(i).Checked = i = Val(lblSearch.Tag)
    Next
    PopupMenu Me.mnuPopuMenu, , picPage.Left + 30, picPage.Top + picDetailParas.Top + lblSearch.Top + lblSearch.Height + 30
End Sub

Private Sub mnuPopuMenuSerch_Click(Index As Integer)
    lblSearch.Caption = mnuPopuMenuSerch(Index).Caption & "��"
    lblSearch.Tag = Index
    mnuPopuMenuSerch(Index).Checked = True
End Sub

Private Sub picDetailParas_Resize()
    err = 0: On Error Resume Next
    With picDetailParas
        fraDetaisModi.Left = 30: fraDetaisModi.Top = .ScaleTop
        fraDetaisModi.Width = .ScaleWidth - fraDetaisModi.Left
        vsDetailParas.Move .ScaleLeft + 20, fraDetaisModi.Top + fraDetaisModi.Height + 50, .ScaleWidth - 20
        vsDetailParas.Height = .ScaleHeight - vsDetailParas.Top
    End With
End Sub

Private Sub picPage_Resize()
    err = 0: On Error Resume Next
    With picPage
        tbPage.Move .ScaleLeft, .ScaleTop, .ScaleWidth, .ScaleHeight
    End With
End Sub

Private Sub picPara_Resize()
    err = 0: On Error Resume Next
    With picPara
        vsPara.Move .ScaleLeft, .ScaleTop, .ScaleWidth, .ScaleHeight
    End With
End Sub
Private Sub picParaChangeLog_Resize()
    err = 0: On Error Resume Next
    With picParaChangeLog
        vsChangeLog.Move .ScaleLeft, .ScaleTop, .ScaleWidth, .ScaleHeight
    End With
End Sub

Private Sub picParInfo_Resize()
    Dim lngWith As Long
    vsParaInfo.Width = picParInfo.ScaleWidth - vsParaInfo.Left
    vsParaInfo.Height = picParInfo.ScaleHeight - vsParaInfo.Top
    lngWith = vsParaInfo.Width - vsParaInfo.ColWidth(0) - 120
    If lngWith < 10 * Me.TextWidth("��") Then lngWith = 10 * Me.TextWidth("��")
    vsParaInfo.ColWidth(1) = lngWith
    Call vsParaInfo.AutoSize(1)
End Sub

Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    tbPage.Tag = Item.Index
End Sub

Private Sub txtSearch_Change()
    Call LoadDetailParas(0, True)
End Sub

Private Sub txtSearch_KeyDown(KeyCode As Integer, Shift As Integer)
     If KeyCode = vbKeyDelete Then
        txtSearch.Text = ""
     End If
End Sub

Private Sub vsDetailParas_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim intDefIdx As Integer, strParType As String
    
    If OldCol = NewCol Then Exit Sub
    With vsDetailParas
        Select Case NewCol
            Case .ColIndex("�û���")
                Call mnuPopuMenuSerch_Click(MI_�û���)
            Case .ColIndex("��Ա")
                Call mnuPopuMenuSerch_Click(MI_��Ա)
            Case .ColIndex("����")
                Call mnuPopuMenuSerch_Click(MI_����)
            Case .ColIndex("������")
                Call mnuPopuMenuSerch_Click(MI_������)
            Case .ColIndex("վ��")
                Call mnuPopuMenuSerch_Click(MI_վ��)
            Case .ColIndex("����ֵ")
                Call mnuPopuMenuSerch_Click(MI_����ֵ)
            Case Else
                strParType = cboParType.Text
                intDefIdx = MI_���� 'Ĭ�ϰ�������������Ϊ������һֱ�ɼ�
                If strParType = "��������ģ��" Then
                    intDefIdx = MI_������
                ElseIf strParType = "˽��ȫ��" Or strParType = "˽��ģ��" Or strParType = "����˽��ģ��" Then
                    intDefIdx = MI_��Ա
                End If
                Call mnuPopuMenuSerch_Click(intDefIdx)
        End Select
        SetModValue
    End With
End Sub

Private Sub vsDetailParas_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vsDetailParas.ColIndex("ѡ��") Then
        Cancel = True
    End If
End Sub

Private Sub vsDetailParas_Click()
    If vsDetailParas.Col = 0 Then
        vsDetailParas.ExplorerBar = flexExNone
    Else
        vsDetailParas.ExplorerBar = flexExSort
    End If
    If vsDetailParas.Col = vsDetailParas.ColIndex("ѡ��") Then
        Call SelDetailParas(vsDetailParas.MouseRow)
    End If
End Sub

Private Sub vsDetailParas_GotFocus()
    Call zl_VsGridGotFocus(vsDetailParas)
End Sub

Private Sub vsDetailParas_LostFocus()
    Call zl_VsGridLOSTFOCUS(vsDetailParas)
End Sub

Private Sub vsDetailParas_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub vsPara_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim lng����id As Long
    Dim blnShowDetail As Boolean
    
    With vsPara
        cmdModify.Enabled = Not (Val(.TextMatrix(.Row, .ColIndex("Fixed"))) = 1 Or Val(.TextMatrix(.Row, .ColIndex("����"))) = 1)
        If OldRow = NewRow Then Exit Sub
        txtSearch.Text = ""
        lblSearch.Tag = ""
        lng����id = .RowData(.Row)
        If .RowHidden(.Row) Then lng����id = 0
        If .Row > 0 Then
            blnShowDetail = Not (InStr(1, "|����ȫ��|����ģ��|", "|" & Trim(.TextMatrix(.Row, .ColIndex("��������")) & "|")) > 0)
        Else
            blnShowDetail = Not (InStr(1, "|����ȫ��|����ģ��|", "|" & Trim(cboParType.Text) & "|") > 0)
        End If
        '������˵����Ϣ
        vsParaInfo.Cell(flexcpFontBold, PR_Ӱ�����˵��, 0, PR_����˵��, 0) = True
        vsParaInfo.Cell(flexcpText, PR_Ӱ�����˵��, 1, PR_����˵��, 1) = "" '����ϴ���Ϣ
        vsParaInfo.TextMatrix(PR_Ӱ�����˵��, 1) = .TextMatrix(.Row, .ColIndex("Ӱ�����˵��"))
        vsParaInfo.TextMatrix(PR_����ֵ����, 1) = .TextMatrix(.Row, .ColIndex("����ֵ����"))
        vsParaInfo.TextMatrix(PR_����˵��, 1) = .TextMatrix(.Row, .ColIndex("����˵��"))
        vsParaInfo.TextMatrix(PR_����˵��, 1) = .TextMatrix(.Row, .ColIndex("����˵��"))
        vsParaInfo.TextMatrix(PR_����˵��, 1) = .TextMatrix(.Row, .ColIndex("����˵��"))
        Call vsParaInfo.AutoSize(1)
        If blnShowDetail Then
            Call SetDetailPara(Trim(.TextMatrix(.Row, .ColIndex("��������"))))
            Call LoadDetailParas(lng����id)
        End If
    End With
    Call LoadChangeLog(lng����id)
    
    tbPage.Item(Pag_Computer).Visible = blnShowDetail
    If tbPage.Item(Val(tbPage.Tag)).Visible Then
        tbPage.Item(Val(tbPage.Tag)).Selected = True
    Else
        tbPage.Item(Pag_ParaInfo).Selected = True
    End If
    If vsPara.Visible And vsPara.Enabled Then vsPara.SetFocus
End Sub
 
Private Sub zl_VsGridGotFocus(ByVal vsGrid As VSFlexGrid)
    '����ؼ�
    With vsGrid
         .SelectionMode = flexSelectionByRow
         .HighLight = flexHighlightAlways
         .BackColorSel = GRD_GOTFOCUS_COLORSEL
    End With
End Sub
Private Sub zl_VsGridLOSTFOCUS(ByVal vsGrid As VSFlexGrid)
    '�뿪�ؼ�
    With vsGrid
         .SelectionMode = flexSelectionByRow
         .FocusRect = flexFocusHeavy
         .HighLight = flexHighlightAlways
         .BackColorSel = GRD_LOSTFOCUS_COLORSEL
    End With
End Sub
Private Sub vsPara_GotFocus()
    Call zl_VsGridGotFocus(vsPara)
End Sub

Private Sub vsPara_LostFocus()
    Call zl_VsGridLOSTFOCUS(vsPara)
End Sub
Private Sub vsChangeLog_GotFocus()
     
    Call zl_VsGridGotFocus(vsChangeLog)
End Sub

Private Sub vsChangeLog_LostFocus()
    Call zl_VsGridLOSTFOCUS(vsChangeLog)
End Sub

Private Sub SelDetailParas(Optional ByVal lngRow As Long)
'���ܣ�����ѡ��vsDetailParas����ȡ��ѡ��
'          lngRow=0-ѡ���ȡ��ѡ�������У�>0ѡ���ȡ��ѡ��ָ����
    Dim blnSel As Boolean, i As Long
    
    With vsDetailParas
        If lngRow < 0 Or lngRow > .Rows - 1 Then Exit Sub
        If lngRow = 0 Then
            blnSel = Val(.ColData(.ColIndex("ѡ��"))) = 0
            .Cell(flexcpPicture, lngRow, .ColIndex("ѡ��")) = imgEdit.ListImages(IIf(blnSel, "AllCheck", "UnCheck")).Picture
            .ColData(.ColIndex("ѡ��")) = IIf(blnSel, 1, 0) '���ͼ��״̬
            For i = .FixedRows To .Rows - 1
                If Val(.RowData(i)) <> 0 Then
                    .TextMatrix(i, .ColIndex("ѡ��")) = IIf(blnSel, -1, 0)
                End If
            Next
            If blnSel Then
                .Tag = Val(.RowData(0))
            Else
                .Tag = 0
            End If
        Else
            If Val(.RowData(lngRow)) <> 0 Then
                blnSel = Val(.TextMatrix(lngRow, .ColIndex("ѡ��"))) = 0
                .TextMatrix(lngRow, .ColIndex("ѡ��")) = IIf(blnSel, -1, 0)
                .Tag = (Val(.Tag) + IIf(blnSel, 1, -1))
                If Val(.Tag) = 0 Then '���еĶ�δѡ����ͼ�����Ϊ����δ��ѡ
                    .Cell(flexcpPicture, 0, .ColIndex("ѡ��")) = imgEdit.ListImages("UnCheck").Picture
                    .ColData(.ColIndex("ѡ��")) = 0
                ElseIf Val(.Tag) = Val(.RowData(0)) Then '���еĶ�ѡ����ͼ�����Ϊ������ѡ
                    .Cell(flexcpPicture, 0, .ColIndex("ѡ��")) = imgEdit.ListImages("AllCheck").Picture
                    .ColData(.ColIndex("ѡ��")) = 1
                End If
            End If
        End If
    End With
    Call SetModValue
End Sub

Private Sub SetModValue()
'�����޸Ĳ���ֵ�ɼ���
    Dim blnVisible As Boolean
    blnVisible = Val(vsDetailParas.Tag) <> 0
    If Not blnVisible And vsDetailParas.Row >= vsDetailParas.FixedRows Then
        blnVisible = Val(vsDetailParas.RowData(vsDetailParas.Row)) <> 0
    End If
    
    cmdModValue.Enabled = blnVisible
    cmdDel.Enabled = blnVisible
End Sub


