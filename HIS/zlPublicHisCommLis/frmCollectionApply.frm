VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Begin VB.Form frmCollectionApply 
   Caption         =   "�����ֹ����뵥"
   ClientHeight    =   8655
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   13080
   Icon            =   "frmCollectionApply.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   8655
   ScaleWidth      =   13080
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picLeftTop 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   1365
      Left            =   60
      ScaleHeight     =   1335
      ScaleWidth      =   10665
      TabIndex        =   0
      Top             =   390
      Width           =   10695
      Begin VB.TextBox txt����1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   5520
         MaxLength       =   5
         TabIndex        =   5
         Top             =   112
         Width           =   555
      End
      Begin VB.ComboBox cboҽ�� 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   6120
         TabIndex        =   10
         Top             =   487
         Width           =   1605
      End
      Begin VB.ComboBox cbo�������� 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCollectionApply.frx":08CA
         Left            =   3510
         List            =   "frmCollectionApply.frx":08CC
         TabIndex        =   9
         Top             =   487
         Width           =   1635
      End
      Begin VB.TextBox txtID 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   8610
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   127
         Width           =   1635
      End
      Begin VB.TextBox txtPatientDept 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   870
         TabIndex        =   8
         Top             =   495
         Width           =   1785
      End
      Begin VB.TextBox txtBed 
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   6690
         TabIndex        =   6
         Top             =   127
         Width           =   1035
      End
      Begin VB.ComboBox cboAge 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCollectionApply.frx":08CE
         Left            =   4740
         List            =   "frmCollectionApply.frx":08DE
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   112
         Width           =   750
      End
      Begin VB.TextBox txt���� 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   4290
         MaxLength       =   5
         TabIndex        =   3
         Top             =   112
         Width           =   435
      End
      Begin VB.ComboBox cbo�Ա� 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         ItemData        =   "frmCollectionApply.frx":08F4
         Left            =   3165
         List            =   "frmCollectionApply.frx":08F6
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   112
         Width           =   675
      End
      Begin VB.TextBox txt���� 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   870
         MaxLength       =   20
         TabIndex        =   1
         ToolTipText     =   "����Ϊ���￨�š���������ͷΪ����ID��������סԺ�š���*������š���.���Һŵ��š���/���շѵ��ݺ�"
         Top             =   112
         Width           =   1785
      End
      Begin VB.ComboBox cboִ�п��� 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   330
         Left            =   8610
         TabIndex        =   11
         Text            =   "cboִ�п���"
         Top             =   487
         Width           =   1635
      End
      Begin VB.TextBox txtUnit 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Left            =   870
         TabIndex        =   12
         Top             =   900
         Width           =   9375
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   8
         Left            =   2760
         TabIndex        =   25
         Top             =   540
         Width           =   720
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��       ��"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   1
         Left            =   120
         TabIndex        =   18
         Top             =   165
         Width           =   675
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��  ʶ ��"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   0
         Left            =   7890
         TabIndex        =   19
         Top             =   165
         Width           =   675
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ִ�п���"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   6
         Left            =   7845
         TabIndex        =   24
         Top             =   540
         Width           =   720
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��        λ"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   10
         Left            =   120
         TabIndex        =   27
         Top             =   930
         Width           =   720
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ҽ��"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   9
         Left            =   5370
         TabIndex        =   26
         Top             =   540
         Width           =   720
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   5
         Left            =   6300
         TabIndex        =   23
         Top             =   165
         Width           =   360
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���ڿ���"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   4
         Left            =   120
         TabIndex        =   22
         Top             =   540
         Width           =   720
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�Ա�"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   2
         Left            =   2760
         TabIndex        =   20
         Top             =   165
         Width           =   360
      End
      Begin VB.Label lblCap 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Index           =   3
         Left            =   3885
         TabIndex        =   21
         Top             =   165
         Width           =   360
      End
   End
   Begin VB.PictureBox picMain 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   6165
      Left            =   0
      ScaleHeight     =   6135
      ScaleWidth      =   10725
      TabIndex        =   28
      Top             =   2100
      Width           =   10755
      Begin VB.Frame fraWE 
         BorderStyle     =   0  'None
         Height          =   4785
         Left            =   2430
         MousePointer    =   9  'Size W E
         TabIndex        =   35
         Top             =   780
         Width           =   60
      End
      Begin VB.PictureBox picItem 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4155
         Left            =   3300
         ScaleHeight     =   4125
         ScaleWidth      =   4215
         TabIndex        =   32
         Top             =   750
         Width           =   4245
         Begin VB.PictureBox picFind 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   435
            Left            =   30
            ScaleHeight     =   435
            ScaleWidth      =   4215
            TabIndex        =   33
            Top             =   0
            Width           =   4215
            Begin VB.TextBox txtFind 
               ForeColor       =   &H80000011&
               Height          =   315
               Left            =   480
               TabIndex        =   14
               Top             =   45
               Width           =   2355
            End
            Begin VB.Label lblCap 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   9
                  Charset         =   0
                  Weight          =   400
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   225
               Index           =   11
               Left            =   60
               TabIndex        =   34
               Top             =   60
               Width           =   360
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfItem 
            Height          =   3285
            Left            =   420
            TabIndex        =   17
            Top             =   1530
            Width           =   3225
            _cx             =   5689
            _cy             =   5794
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
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
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
      Begin VB.PictureBox picGroup 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   4935
         Left            =   180
         ScaleHeight     =   4905
         ScaleWidth      =   2085
         TabIndex        =   31
         Top             =   720
         Width           =   2115
         Begin VSFlex8Ctl.VSFlexGrid vsfGroup 
            Height          =   3195
            Left            =   180
            TabIndex        =   16
            Top             =   420
            Width           =   1485
            _cx             =   2619
            _cy             =   5636
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   -1  'True
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   10.5
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
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   1
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
      Begin VB.PictureBox picFilter 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   180
         ScaleHeight     =   345
         ScaleWidth      =   10155
         TabIndex        =   29
         Top             =   120
         Width           =   10185
         Begin VB.ComboBox cboSampleType 
            Height          =   300
            Left            =   1140
            TabIndex        =   13
            Text            =   "cboSampleType"
            Top             =   30
            Width           =   1365
         End
         Begin VB.CheckBox chkConcatenation 
            BackColor       =   &H80000005&
            Caption         =   "���浱ǰ��Ŀ��������"
            Height          =   225
            Left            =   2940
            TabIndex        =   15
            Top             =   60
            Width           =   2295
         End
         Begin VB.Label lblCap 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�걾����"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   225
            Index           =   7
            Left            =   240
            TabIndex        =   30
            Top             =   60
            Width           =   720
         End
      End
   End
   Begin VB.PictureBox picRight 
      Height          =   8205
      Left            =   10800
      ScaleHeight     =   8145
      ScaleWidth      =   2145
      TabIndex        =   36
      Top             =   360
      Width           =   2205
      Begin VSFlex8Ctl.VSFlexGrid VSFSeled 
         Height          =   7725
         Left            =   30
         TabIndex        =   38
         Top             =   330
         Width           =   2085
         _cx             =   3678
         _cy             =   13626
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
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16706793
         ForeColorSel    =   0
         BackColorBkg    =   16777215
         BackColorAlternate=   16777215
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483635
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   3
         Cols            =   3
         FixedRows       =   0
         FixedCols       =   0
         RowHeightMin    =   350
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   ""
         ScrollTrack     =   -1  'True
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
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label Label10 
         AutoSize        =   -1  'True
         Caption         =   "��ѡ��(˫��ȡ��ѡ��)"
         BeginProperty Font 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   210
         Left            =   30
         TabIndex        =   37
         Top             =   60
         Width           =   2100
      End
   End
   Begin XtremeCommandBars.CommandBars cbrthis 
      Left            =   30
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmCollectionApply"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long

Private mrsRelativeAdvice As ADODB.Recordset                             '�Ǽǵ����ҽ��
Private PatientType As Integer, mlng����ID As Long, mstrNO As String    '�����շѵ��ݺ�
Private mlngCapID As Long                                               '�ɼ���ĿID
Private mlngReqDept As Long, mstrReqDoctor As String                    'Ĭ�ϵĵǼǿ��Һ�ҽ��
Private mblnSaveAdvice As Boolean                                       '�Ƿ���Ҫ����ҽ���������޸���Ժ���˱걾��Ϣ
Private mstrKeys As String                                              '��ǰ���յ�����ҽ��ID
Private mblnBarCode As Boolean                                          '����
Private miInputType As Integer

Private mlngDeptID As Long                                              '����ID
Private mrsItem As ADODB.Recordset              '�����Ŀ
Private mstrItemSel As String                   'ѡ��������Ŀ
Private mblnFindEOF As Boolean                  '����ʱ���Ƿ��Ѿ������¼��ĩβ
Private mblnEdit As Boolean                     '�Ƿ�༭������
Private mstrSQLPro() As String                  '�ύ�����õ�sql
Private mblnLoad As Boolean                     '�Ƿ��״μ���

Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

Private Const mConst_������Ϣ_���� As String = "a.����id,a.�����,a.סԺ��,a.���￨��,a.����֤��,a.�ѱ�,a.ҽ�Ƹ��ʽ,a.����,a.�Ա�,a.����,a.��������," & _
                                              "a.�����ص�,a.���֤��,a.���,a.ְҵ,a.����,a.����,a.����,a.ѧ��,a.����״��,a.��ͥ��ַ,a.��ͥ�绰,a.��ͥ��ַ�ʱ�," & _
                                              "a.��ϵ�˹�ϵ,a.��ϵ�˵�ַ,a.��ϵ�˵绰,a.��ͬ��λID,a.������λ,a.��λ�绰,a.��λ�ʱ�,a.��λ������,a.��λ�ʺ�," & _
                                              "a.����ʱ��,a.����״̬,a.��������,a.סԺ����,a.��ǰ����ID,a.��ǰ����ID,a.��Ժʱ��,a.��Ժʱ��," & _
                                              "a.IC����,a.������,a.����,a.�Ǽ�ʱ��,a.ͣ��ʱ��,a.��ǰ����,a.ҽ����,a.��ѯ����,a.��Ժ,a.����֤��,a.�໤��,a.����,a.��ҳid"


Private Sub cboAge_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then PressKey vbKeyTab
End Sub

Private Sub cboSampleType_Click()
    Call vsfGroup_RowColChange
End Sub

Private Sub cboSampleType_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then PressKey vbKeyTab
End Sub

Private Sub cbo��������_Click()
    If cbo��������.ListIndex > -1 Then InitDoctors cbo��������.ItemData(cbo��������.ListIndex)
End Sub

Private Sub cbo��������_GotFocus()
    Call TxtSelAll(cbo��������)
End Sub

Private Sub cbo��������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then PressKey vbKeyTab
End Sub

Private Sub cbo��������_Validate(Cancel As Boolean)
          Dim rsTmp As ADODB.Recordset
          Dim strSQL As String, intIdx As Long
          Dim strInput As String
          Dim vRect As RECT, blnCancel As Boolean
              
1         On Error GoTo cbo��������_Validate_Error

2         If cbo��������.ListIndex <> -1 Then mlngReqDept = Me.cbo��������.ItemData(Me.cbo��������.ListIndex): Exit Sub '��ѡ��
3         If cbo��������.Text = "" Then '������
4             Exit Sub
5         End If
          
6         strInput = UCase(NeedName(cbo��������.Text))
          'ȫԺ�ٴ�����
7         strSQL = _
              " Select Distinct A.ID,A.����,A.����,A.����" & _
              " From ���ű� A,��������˵�� B " & _
              " Where B.����ID = A.ID " & _
              " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
              " And (B.�������� IN('�ٴ�','���'))" & _
              " And (Upper(A.����) Like [1] Or Upper(A.����) Like [2] Or Upper(A.����) Like [2])" & _
              " Order by A.����"
          
8         vRect = GetControlRect(cbo��������.hWnd)
9         Set rsTmp = gobjHisDatabase.ShowSQLSelect(Me, strSQL, 0, "��������", False, "", "", False, False, _
              True, vRect.Left, vRect.Top, cbo��������.Height, blnCancel, False, True, strInput & "%", strInput & "%")
10        If Not rsTmp Is Nothing Then
11            If Not CboLocate(cbo��������, rsTmp!����) Then
12                cbo��������.Text = ""
13            End If
14        Else
15            If Not blnCancel Then
16                MsgBox "δ�ҵ���Ӧ�Ŀ��ҡ�", vbInformation, Me.Caption
17            End If
18            Cancel = True: Exit Sub
19        End If
20        If Me.cbo��������.ListIndex > -1 Then mlngReqDept = Me.cbo��������.ItemData(Me.cbo��������.ListIndex)


21        Exit Sub
cbo��������_Validate_Error:
22        Call WriteErrLog("zlPublicHisCommLis", "frmcollectionApply", "ִ��(cbo��������_Validate)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
23        Err.Clear

End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/4/27
'��    ��:���ݲ�ͬ�Ա��ѯ��ͬ�ı걾���ͣ���ֹ��������Ů���о�Һ�걾����Ц����
'��    ��:
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Private Sub cbo�Ա�_Click()
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          
1         On Error GoTo cbo�Ա�_Click_Error
              
2         strSQL = "select ���� from ����걾���� where �����Ա�=[1] or nvl(�����Ա�,0)=0"
3         Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "����걾����", IIf(Me.cbo�Ա�.Text = "��", 1, 2))
4         With cboSampleType
5             .Clear
6             .AddItem "���б걾"
7             Do While Not rsTmp.EOF
8                 .AddItem rsTmp("����") & ""
9                 rsTmp.MoveNext
10            Loop
11            If .ListCount > 0 Then .ListIndex = 0
12        End With
          

13        Exit Sub
cbo�Ա�_Click_Error:
14        Call WriteErrLog("zlPublicHisCommLis", "frmCollectionApply", "ִ��(cbo�Ա�_Click)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
15        Err.Clear
End Sub


Private Sub cbo�Ա�_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        PressKey vbKeyTab
        Exit Sub
    End If
End Sub

Private Sub cboҽ��_Click()
    Call TxtSelAll(cboҽ��)
End Sub

Private Sub cboҽ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then PressKey vbKeyTab
End Sub

Private Sub cboҽ��_Validate(Cancel As Boolean)
          Dim rsTmp As ADODB.Recordset
          Dim strSQL As String, intIdx As Long
          Dim strInput As String
          Dim vRect As RECT, blnCancel As Boolean
              
1         On Error GoTo cboҽ��_Validate_Error

2         If cboҽ��.ListIndex <> -1 Then mstrReqDoctor = Me.cboҽ��.Text: Exit Sub '��ѡ��
3         If cboҽ��.Text = "" Then '������
4             Exit Sub
5         End If
          
6         strInput = UCase(NeedName(cboҽ��.Text))
          'ȫԺҽ��
7         strSQL = "Select Distinct ����ID From ��������˵�� Where ������� IN(1,2,3)"
8         strSQL = "Select Distinct A.ID,A.���,A.����,A.����" & _
              " From ��Ա�� A,������Ա B,��Ա����˵�� C" & _
              " Where A.ID=B.��ԱID And A.ID=C.��ԱID And C.��Ա����='ҽ��'" & _
              " And B.����ID IN(" & strSQL & ")" & _
              " And (Upper(A.���) Like [1] Or Upper(A.����) Like [2] Or Upper(A.����) Like [2])" & _
              " And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) " & _
              " Order by A.����"
          
9         vRect = GetControlRect(cboҽ��.hWnd)
10        Set rsTmp = gobjHisDatabase.ShowSQLSelect(Me, strSQL, 0, "����ҽ��", False, "", "", False, False, _
              True, vRect.Left, vRect.Top, cboҽ��.Height, blnCancel, False, True, strInput & "%", strInput & "%")
11        If Not rsTmp Is Nothing Then
12            cboҽ��.Text = rsTmp!����
13        Else
14            If Not blnCancel Then
15                MsgBox "δ�ҵ���Ӧ��ҽ����", vbInformation, Me.Caption
16            End If
17            Cancel = True: Exit Sub
18        End If
19        If Len(Trim(Me.cboҽ��.Text)) > 0 Then mstrReqDoctor = Me.cboҽ��.Text


20        Exit Sub
cboҽ��_Validate_Error:
21        Call WriteErrLog("zlPublicHisCommLis", "frmcollectionApply", "ִ��(cboҽ��_Validate)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
22        Err.Clear

End Sub

Private Sub cboִ�п���_Click()
    mlngDeptID = cboִ�п���.ItemData(cboִ�п���.ListIndex)
End Sub

Private Sub cboִ�п���_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then PressKey vbKeyTab
End Sub

Private Sub cboִ�п���_Validate(Cancel As Boolean)
          Dim rsTmp As ADODB.Recordset
          Dim strSQL As String
          Dim strInput As String
          Dim vRect As RECT, blnCancel As Boolean

1         On Error GoTo cboִ�п���_Validate_Error

2         If cboִ�п���.ListIndex <> -1 Then mlngReqDept = Me.cboִ�п���.ItemData(Me.cboִ�п���.ListIndex): Exit Sub    '��ѡ��
3         If cboִ�п���.Text = "" Then    '������
4             Exit Sub
5         End If

6         strInput = UCase(NeedName(cboִ�п���.Text))
          'ȫԺ�ٴ�����
7         strSQL = _
        " Select Distinct A.ID,A.����,A.����,A.����" & _
                 " From ���ű� A,��������˵�� B " & _
                 " Where B.����ID = A.ID " & _
                 " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
                 " And (B.�������� IN('����'))" & _
                 " And (Upper(A.����) Like [1] Or Upper(A.����) Like [2] Or Upper(A.����) Like [2])" & _
                 " Order by A.����"


8         vRect = GetControlRect(cboִ�п���.hWnd)
9         Set rsTmp = gobjHisDatabase.ShowSQLSelect(Me, strSQL, 0, "��������", False, "", "", False, False, _
                                               True, vRect.Left, vRect.Top, cboִ�п���.Height, blnCancel, False, True, strInput & "%", strInput & "%")
10        If Not rsTmp Is Nothing Then
11            If Not CboLocate(cboִ�п���, rsTmp!����) Then
12                cboִ�п���.Text = ""
13            End If
14        Else
15            If Not blnCancel Then
16                MsgBox "δ�ҵ���Ӧ�Ŀ��ҡ�", vbInformation, Me.Caption
17            End If
18            Cancel = True: Exit Sub
19        End If
20        If Me.cboִ�п���.ListIndex > -1 Then mlngReqDept = Me.cboִ�п���.ItemData(Me.cboִ�п���.ListIndex)


21        Exit Sub
cboִ�п���_Validate_Error:
22        Call WriteErrLog("zlPublicHisCommLis", "frmCollectionApply", "ִ��(cboִ�п���_Validate)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
23        Err.Clear

End Sub

Private Sub cbrthis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
        Case ConMenu_Browse_Save        '����
            Call getSelItems
        Case ConMenu_Browse_Cancel      'ȡ��
            Call cancelEdit
        Case ConMenu_Appfro_Exit        '�˳�
            Unload Me
    End Select
End Sub

Private Sub cbrthis_ResizeClient(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
    On Error Resume Next
    With Me.picRight
        .Left = Right - .Width
        .Top = Top
        .Height = Bottom - Top
    End With
    
    With Me.picLeftTop
        .Left = Left
        .Top = Top
        .Width = Me.picRight.Left - Left
    End With
    
    With Me.picMain
        .Left = Left
        .Top = Me.picLeftTop.Top + Me.picLeftTop.Height + 50
        .Width = Me.picLeftTop.Width
        .Height = Bottom - .Top
    End With
End Sub


Private Function SaveData(ByRef strNewAdvice As String) As Boolean
    
    SaveData = SaveAdviceData(strNewAdvice)

End Function

Private Sub cancelEdit()
    Me.VSFSeled.Rows = 0
    Call CheckSelItem
End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/5/3
'��    ��:��ȡѡ�����Ŀ
'��    ��:
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Private Sub getSelItems()
          Dim rsTmp As New ADODB.Recordset
          Dim strItemSel As String
          Dim strOldNames As String
          Dim lngRow As Long
          Dim strSampleType As String
          Dim lngLoop As Long
          Dim blnTrs As Boolean
          Dim strCodeBefor As String  '�Թܱ���
          Dim strCode As String   '�Թܱ���
          Dim strNewAdvice As String
          Dim strErr As String

1         On Error GoTo getSelItems_Error

2         ReDim mstrSQLPro(0)

          '����������ݵĺϷ���
3         If Not ValidAdvice Then Exit Sub

          '��ͬ�걾������Ҫ�����ύ
4         With Me.VSFSeled
5             For lngRow = 0 To .Rows - 1
6                 strCodeBefor = GetSampleCode(Val(.TextMatrix(lngRow, .ColIndex("oldid"))))
7                 If (strSampleType = .TextMatrix(lngRow, .ColIndex("�걾")) Or strSampleType = "") And (strCode = strCodeBefor Or strCode = "") Then
8                     strItemSel = strItemSel & "," & .TextMatrix(lngRow, .ColIndex("oldid"))
9                     strOldNames = strOldNames & "," & .TextMatrix(lngRow, .ColIndex("oldName"))
10                Else
11                    If strItemSel <> "" Then strItemSel = Mid(strItemSel, 2) & ";" & strSampleType
12                    If strOldNames <> "" Then strOldNames = Mid(strOldNames, 2) & ";" & strSampleType
13                    mstrItemSel = strOldNames
14                    If mstrItemSel <> "" Then
                          '��ȡ�ɼ���ʽ
15                        Set rsTmp = SelectCap(Split(Split(strItemSel, ";")(0), ",")(0))
16                        If rsTmp Is Nothing Then
17                            MsgBox "û�ж���걾�ɼ���ʽ���뵽������Ŀ���������á�", vbInformation, Me.Caption
18                            Exit Sub
19                        End If
20                        mlngCapID = rsTmp("ID")
21                        Call AdviceSet�������(3, strItemSel)
22                    End If
23                    If Not SaveData(strNewAdvice) Then Exit Sub

24                    strItemSel = ""
25                    strItemSel = strItemSel & "," & .TextMatrix(lngRow, .ColIndex("oldid"))
26                End If
27                strSampleType = .TextMatrix(lngRow, .ColIndex("�걾"))
28                strCode = strCodeBefor
29            Next
30        End With

          '�������һ��
31        If strItemSel <> "" Then
32            strItemSel = Mid(strItemSel, 2) & ";" & strSampleType
33            strOldNames = Mid(strOldNames, 2) & ";" & strSampleType
34            mstrItemSel = strOldNames
35            If mstrItemSel <> "" Then
                  '��ȡ�ɼ���ʽ
36                Set rsTmp = SelectCap(Split(Split(strItemSel, ";")(0), ",")(0))
37                If rsTmp Is Nothing Then
38                    MsgBox "û�ж���걾�ɼ���ʽ���뵽������Ŀ���������á�", vbInformation, Me.Caption
39                    Exit Sub
40                End If
41                mlngCapID = rsTmp("ID")
42                Call AdviceSet�������(3, strItemSel)
43            End If
44            If Not SaveData(strNewAdvice) Then Exit Sub
45        End If

          '�ύ����
46        gcnHisOracle.BeginTrans
47        blnTrs = True
48        For lngLoop = 1 To UBound(mstrSQLPro)
49            Call ComExecuteProc(Sel_His_DB, mstrSQLPro(lngLoop), Me.Caption)
50        Next

          '�����°�ҽ��
51        If SampleBarcodeUpdate(strNewAdvice, "", "", strErr, 0) = False Then
52            gcnHisOracle.RollbackTrans
53            blnTrs = False
54            If strErr <> "" Then
55                MsgBox strErr, vbInformation, gSysInfo.AppName
56            End If
57            Exit Sub
58        End If

59        gcnHisOracle.CommitTrans
60        blnTrs = False

61        If Me.chkConcatenation.value = 1 Then
62            Me.txt����.SetFocus
63        Else
64            Me.txt���� = ""
65            Me.txt���� = "": Me.txt����1 = "": Me.cboAge.ListIndex = 0
66            Me.txtBed = "": Me.txtID = ""
67            Me.txtPatientDept = "": Me.txtPatientDept.Tag = 0
68            Me.cbo��������.ListIndex = -1
69            Me.cboҽ��.ListIndex = -1
70            txtUnit.Text = ""

71            Me.VSFSeled.Rows = 0
72            Call CheckSelItem

73            Me.txt����.SetFocus
74        End If

75        MsgBox "�Ǽǳɹ�����ˢ�²����б�鿴��", vbInformation, Me.Caption

76        Exit Sub
getSelItems_Error:
77        If blnTrs Then gcnHisOracle.RollbackTrans
78        Call WriteErrLog("zlPublicHisCommLis", "frmcollectionApply", "ִ��(getSelItems)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
79        Err.Clear
End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/9/20
'��    ��:��ȡ�Թܱ���
'��    ��:
'��    ��:
'��    ��:
'����Ӱ��:
'---------------------------------------------------------------------------------------
Private Function GetSampleCode(ByVal lngOldID As Long) As String
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          
1         On Error GoTo GetSampleCode_Error

2         strSQL = "Select �Թܱ��� From ������ĿĿ¼ Where ID = [1] "
3         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "������ĿĿ¼", lngOldID)
4         If Not rsTmp.EOF Then
5             GetSampleCode = rsTmp("�Թܱ���") & ""
6         End If


7         Exit Function
GetSampleCode_Error:
8         Call WriteErrLog("zlPublicHisCommLis", "frmCollectionApply", "ִ��(GetSampleCode)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
9         Err.Clear
End Function

Private Sub cbrthis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    If VSFSeled.Rows > 0 And Trim(txt����.Text) <> "" Then
        mblnEdit = True
    Else
        mblnEdit = False
    End If
    Select Case Control.ID
        Case ConMenu_Browse_Save        '����
            Control.Enabled = mblnEdit
        Case ConMenu_Browse_Cancel      'ȡ��
            Control.Enabled = mblnEdit
    End Select
End Sub

Private Sub Form_Activate()
    If mblnLoad Then
        txt����.SetFocus
        mblnLoad = False
    End If
End Sub

Private Sub Form_Load()
    
    '���ܴ���������
    Dim cbrControl As CommandBarControl
    Dim cbrToolBar As CommandBar
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    Me.cbrthis.VisualTheme = xtpThemeOffice2003
    Me.cbrthis.Icons = frmPubIcons.imgPublic.Icons
    With Me.cbrthis.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    Me.cbrthis.EnableCustomization False

    '-----------------------------------------------------
    '�˵�����
    Me.cbrthis.ActiveMenuBar.Title = "�˵�"
    Me.cbrthis.ActiveMenuBar.Visible = False
    Set cbrToolBar = Me.cbrthis.Add("������", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_Save, "����(Crl+S)")
        Set cbrControl = .Add(xtpControlButton, ConMenu_Browse_Cancel, "ȡ��(Crl+U)")
        Set cbrControl = .Add(xtpControlButton, ConMenu_Appfro_Exit, "�˳�(Crl+Q)")
        cbrControl.BeginGroup = True
    End With
    
    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type = xtpControlButton Then
            cbrControl.Style = xtpButtonIconAndCaption
        End If
    Next
    '�����
    With Me.cbrthis.KeyBindings
        .Add FCONTROL, vbKeyS, ConMenu_Browse_Save
        .Add FCONTROL, vbKeyU, ConMenu_Browse_Cancel
        .Add FCONTROL, vbKeyQ, ConMenu_Appfro_Exit
    End With
         
    With VSFSeled
        .ExplorerBar = flexExSortShow
        .Rows = 0
        .Cols = 6
        .ColKey(0) = "ID": .ColHidden(0) = True
        .ColKey(1) = "����"
        .ColKey(2) = "���Ʊ���": .ColHidden(2) = True
        .ColKey(3) = "oldID": .ColHidden(3) = True
        .ColKey(4) = "�걾": .ColHidden(4) = True
        .ColKey(5) = "oldName": .ColHidden(5) = True
    End With
    
    Call InitDepts                      'ȡ�ÿ��Һ��Ա�
    Call intData                        '������Ŀ
    
    '�����ı�����ʾ��
    Call setTxtTip(txtFind, "������롢����������û��س�����")
    
    mblnLoad = True
End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/4/26
'��    ��:��������
'��    ��:
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Private Sub intData()
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          Dim rsGroup As ADODB.Recordset
          Dim strFenLei As String
          Dim blnHasGroup As Boolean  '�Ƿ���ִ��С��


1         On Error GoTo intData_Error


          '��ѯ�����Ŀ
2         If VerCompare(gSysInfo.VersionLIS, "10.35.130") <> -1 Then
3         strSQL = "Select 0 ѡ��, a.Id, a.����, a.����, a.����, b.���뵥id, Nvl(b.����id, 0) ����id, a.���Ʊ���, a.����걾 �걾" & vbCrLf & _
                 " From ���������Ŀ A, �������뵥��ϸ B" & vbCrLf & _
                 " Where a.Id = b.���id And a.ͣ������ Is Null And a.���Ʊ��� Is Not Null And nvl(a.�Ƿ�������Ŀ, 0) = 0"
4         Else
5             strSQL = "Select 0 ѡ��, a.Id, a.����, a.����, a.����, b.���뵥id, Nvl(b.����id, 0) ����id, a.���Ʊ���, a.����걾 �걾" & vbCrLf & _
                 " From ���������Ŀ A, �������뵥��ϸ B" & vbCrLf & _
                 " Where a.Id = b.���id And a.ͣ������ Is Null And a.���Ʊ��� Is Not Null"
6         End If
7         If gUserInfo.NodeNo <> "-" Then
8             strSQL = strSQL & " And (a.վ�� = [1] Or Nvl(a.վ��, 0) = 0)"
9         End If
10        strSQL = strSQL & " Order By b.���뵥id,b.����id,b.����˳��"
11        Set mrsItem = ComOpenSQL(Sel_Lis_DB, strSQL, "���������Ŀ", gUserInfo.NodeNo)

          '��ѯ���뵥����
12        If VerCompare(gSysInfo.VersionLIS, "10.35.130") <> -1 Then
13            strSQL = "Select ID,���뵥ID, ����, ���� ����,ִ��С��" & vbCrLf & _
                     " From (Select Distinct a.���� ����, b.Id,a.ID ���뵥ID, b.����,a.ִ��С��" & vbCrLf & _
                     "        From �������뵥 A, �������뵥���� B, �������뵥��ϸ C" & vbCrLf & _
                     "        Where a.Id = b.���뵥id(+) And a.Id = c.���뵥id And c.����id Is Not Null And" & vbCrLf & _
                     "        nvl(a.�Ƿ��������뵥, 0) = 0 And (a.����id = [1] Or Nvl(a.����id, 0) = 0)" & vbCrLf & _
                     "        Union all" & vbCrLf & _
                     "        Select Distinct a.����, 0 id,a.ID ���뵥ID, 'δ����' ����,a.ִ��С��" & vbCrLf & _
                     "        From �������뵥 A, �������뵥���� B, �������뵥��ϸ C" & vbCrLf & _
                     "        Where a.Id = b.���뵥id(+) And a.Id = c.���뵥id And c.����id Is Null And" & vbCrLf & _
                     "        nvl(a.�Ƿ��������뵥, 0) = 0 And (a.����id = [1] Or Nvl(a.����id, 0) = 0))" & vbCrLf & _
                     " Order By ����,���뵥ID,ID"
14        Else
15            strSQL = "Select ID,���뵥ID, ����, ���� ����,ִ��С��" & vbCrLf & _
                     " From (Select Distinct a.���� ����, b.Id,a.ID ���뵥ID, b.����,a.ִ��С��" & vbCrLf & _
                     "        From �������뵥 A, �������뵥���� B, �������뵥��ϸ C" & vbCrLf & _
                     "        Where a.Id = b.���뵥id(+) And a.Id = c.���뵥id And c.����id Is Not Null And" & vbCrLf & _
                     "              (a.����id = [1] Or Nvl(a.����id, 0) = 0)" & vbCrLf & _
                     "        Union all" & vbCrLf & _
                     "        Select Distinct a.����, 0 id,a.ID ���뵥ID, 'δ����' ����,a.ִ��С��" & vbCrLf & _
                     "        From �������뵥 A, �������뵥���� B, �������뵥��ϸ C" & vbCrLf & _
                     "        Where a.Id = b.���뵥id(+) And a.Id = c.���뵥id And c.����id Is Null And" & vbCrLf & _
                     "              (a.����id = [1] Or Nvl(a.����id, 0) = 0))" & vbCrLf & _
                     " Order By ����,���뵥ID,ID"
16        End If
17        Set rsTmp = ComOpenSQL(Sel_Lis_DB, strSQL, "�������뵥", mlngDeptID)

          '��ѯվ���Ӧ��ִ��С��
18        If gUserInfo.NodeNo <> "-" Then
19            strSQL = "Select Distinct ���� From ����С���¼ Where վ�� = [1] or վ�� is null"
20        Else
21            strSQL = "Select Distinct ���� From ����С���¼"
22        End If
23        Set rsGroup = ComOpenSQL(Sel_Lis_DB, strSQL, "����С���¼", gUserInfo.NodeNo)


24        With Me.vsfGroup
25            .ExplorerBar = flexExSort
26            .Rows = 1
27            .Cols = 4
28            .FixedRows = 1
29            .OutlineBar = flexOutlineBarComplete
30            .OutlineCol = 1
              '        .SubtotalPosition = flexSTAbove
31            .ExtendLastCol = True

              '1.�߿�
32            .Appearance = flex3DLight
33            .BorderStyle = flexBorderFlat
34            .GridLines = flexGridNone
35            .GridColorFixed = flexGridNone

              '2.��ɫ
36            .BackColor = vbWindowBackground    '���ڱ���
37            .BackColorAlternate = vbWindowBackground
38            .BackColorBkg = vbWindowBackground
39            .BackColorFixed = vbButtonFace    '��ť����
40            .BackColorFrozen = &H0&         '��
41            .FloodColor = &HC0&             '��
42            .BackColorSel = &HFFEBD7        'ǳ��
43            .ForeColor = vbWindowText       '�����ı�
44            .ForeColorFixed = vbButtonText  '��ť�ı�
45            .ForeColorFrozen = &H0&         '��
46            .ForeColorSel = vbWindowText

47            .GridColor = vbApplicationWorkspace    'Ӧ�ó�������
48            .GridColorFixed = vbApplicationWorkspace
49            .SheetBorder = vbWindowBackground
50            .TreeColor = vbButtonShadow         '��ť��Ӱ


51            .ColKey(0) = "id": .ColWidth(.ColIndex("id")) = 0: .ColHidden(.ColIndex("id")) = True
52            .ColKey(1) = "���뵥ID": .ColWidth(.ColIndex("���뵥ID")) = 250
53            .ColKey(2) = "����": .ColAlignment(.ColIndex("����")) = flexAlignCenterCenter: .ColHidden(.ColIndex("����")) = True
54            .ColKey(3) = "����": .ColWidth(.ColIndex("����")) = 250: .ColAlignment(.ColIndex("����")) = flexAlignCenterCenter: .TextMatrix(0, .ColIndex("����")) = "����"

55            Do While Not rsTmp.EOF
56                blnHasGroup = False
                  '�ж����뵥�Ƿ���ִ��С��
57                If Not rsGroup Is Nothing Then
58                    If rsGroup.RecordCount > 0 Then
59                        rsGroup.MoveFirst
60                        Do While Not rsGroup.EOF
61                            If InStr("," & rsTmp("ִ��С��") & ",", "," & rsGroup("����") & ",") > 0 Then
62                                blnHasGroup = True
63                            End If
64                            rsGroup.MoveNext
65                        Loop
66                    End If
67                End If

68                If blnHasGroup Then
69                    If InStr(";" & strFenLei & ";", ";" & rsTmp("����") & ";") <= 0 Then
70                        .Rows = .Rows + 2

71                        .TextMatrix(.Rows - 2, .ColIndex("ID")) = rsTmp("����") & ""
72                        .TextMatrix(.Rows - 2, .ColIndex("���뵥ID")) = rsTmp("����") & ""
73                        .TextMatrix(.Rows - 2, .ColIndex("����")) = rsTmp("����") & ""
74                        .TextMatrix(.Rows - 2, .ColIndex("����")) = rsTmp("����") & "": .Cell(flexcpAlignment, .Rows - 2, .ColIndex("����")) = flexAlignLeftCenter

                          '�Ӵ�
75                        .Cell(flexcpFontBold, .Rows - 2, 0, .Rows - 2, .Cols - 1) = True

                          '�ϲ�
76                        .MergeRow(.Rows - 2) = True
77                        .MergeCellsFixed = flexMergeRestrictRows

                          '����
78                        .IsSubtotal(.Rows - 2) = True
79                        .RowOutlineLevel(.Rows - 2) = 1

80                        .TextMatrix(.Rows - 1, .ColIndex("ID")) = rsTmp("ID") & ""
81                        .TextMatrix(.Rows - 1, .ColIndex("���뵥ID")) = rsTmp("���뵥ID") & ""
82                        .TextMatrix(.Rows - 1, .ColIndex("����")) = rsTmp("����") & ""
83                        strFenLei = strFenLei & ";" & rsTmp("����") & ""
84                    Else
85                        .Rows = .Rows + 1
86                        .TextMatrix(.Rows - 1, .ColIndex("ID")) = rsTmp("ID") & ""
87                        .TextMatrix(.Rows - 1, .ColIndex("���뵥ID")) = rsTmp("���뵥ID") & ""
88                        .TextMatrix(.Rows - 1, .ColIndex("����")) = rsTmp("����") & ""
89                    End If
90                End If

91                rsTmp.MoveNext
92            Loop

              'Ĭ��ѡ�е�һ������
93            If .Rows > 2 Then .Row = 2
94        End With



95        Exit Sub
intData_Error:
96        Call WriteErrLog("zlPublicHisCommLis", "frmCollectionApply", "ִ��(intData)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
97        Err.Clear

End Sub

Private Sub Form_Unload(Cancel As Integer)
    SetPara Sel_His_DB, "�ɼ�����վ�Ǽ�", chkConcatenation.value, 100, 1211
    Set mrsItem = Nothing
    Set mrsRelativeAdvice = Nothing
End Sub

Private Sub fraWE_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    On Error Resume Next
    If Button <> 1 Then Exit Sub
    With Me.fraWE
        If .Left + X < 2000 Or picMain.Width - (.Left + X) < 2000 Then Exit Sub
        .Left = .Left + X
        .Tag = .Left
    End With
    With Me.picGroup
        .Width = Me.fraWE.Left
    End With
    
    With Me.picItem
        .Left = Me.fraWE.Left + Me.fraWE.Width
        .Width = Me.picMain.Width - .Left
    End With
    
End Sub

Private Sub picFind_Resize()
    On Error Resume Next
    With Me.txtFind
        .Width = Me.picFind.Width - .Left - 200
    End With
End Sub

Private Sub picGroup_Resize()
    On Error Resume Next
    With Me.vsfGroup
        .Left = 0
        .Top = 0
        .Width = Me.picGroup.Width
        .Height = Me.picGroup.Height
    End With
End Sub

Private Sub picItem_Resize()
    On Error Resume Next
    With Me.picFind
        .Left = 0
        .Top = 0
        .Width = Me.picItem.Width
    End With
    With Me.vsfItem
        .Left = 0
        .Top = picFind.Height
        .Width = Me.picItem.Width
        .Height = Me.picItem.Height - .Top
    End With
End Sub

Private Sub PicMain_Resize()
    On Error Resume Next
    With Me.picFilter
        .Left = 0
        .Top = 0
        .Width = Me.picMain.Width
        .BorderStyle = 0
    End With
    With Me.fraWE
        .Top = picFilter.Height
        .Height = picMain.Height - .Top
    End With
    With Me.picGroup
        .Left = 0
        .Top = picFilter.Height
        .Width = Me.fraWE.Left
        .Height = Me.picMain.Height - .Top
    End With
    With Me.picItem
        .Left = Me.fraWE.Left + Me.fraWE.Width
        .Top = picGroup.Top
        .Width = Me.picMain.Width - .Left
        .Height = picGroup.Height
    End With
End Sub

Private Sub picRight_Resize()
    On Error Resume Next
    With VSFSeled
        .Height = Me.picRight.Height - .Top
        .Width = Me.picRight.Width - 100
    End With
End Sub

Private Sub txtBed_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then PressKey vbKeyTab
End Sub

Private Sub txtFind_GotFocus()
    Call selAllText(txtFind)
End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/5/2
'��    ��:������Ŀ
'��    ��:
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Private Sub txtFind_KeyPress(KeyAscii As Integer)
          Dim strFind As String
          Dim strFilter As String
          Dim strTmp() As String
          Dim strSub() As String
          Dim lngRow As Long
          Dim i As Integer
          
          
1         On Error GoTo txtFind_KeyPress_Error

2         If KeyAscii <> 13 Then
3             vsfItem.Tag = ""
4             mblnFindEOF = False
5             Exit Sub
6         End If
7         If mrsItem Is Nothing Then Exit Sub
          
8         strFind = UCase(Trim(Me.txtFind.Text))
9         If strFind = "" Then Exit Sub
          'ͨ�����ݵ�����ȥ���˼�¼��
10        mrsItem.Filter = "���� like '%" & strFind & "%'"
11        If mrsItem.RecordCount <= 0 Then
12            mrsItem.Filter = "���� like '%" & strFind & "%'"
13            If mrsItem.RecordCount <= 0 Then
14                mrsItem.Filter = "���� like '%" & strFind & "%'"
15            End If
16        End If

          
          '���ڲ�ͬ�����¿��ܴ�����ͬ����Ŀ�����Բ���ʱ���ܻ���ڶ��м�¼����Ҫ����ÿһ�м�¼
17        If mblnFindEOF = False Then  'ֻ�е��ϴι��˵ļ�¼���е����ݶ��Ѿ�������һ��֮��Ž����µļ�¼
18            strFilter = ""
19            Do While Not mrsItem.EOF
20                strFilter = strFilter & ";" & mrsItem("���뵥ID") & "," & mrsItem("����ID") & "," & mrsItem("ID")
21                mrsItem.MoveNext
22            Loop
23            If strFilter <> "" Then Me.txtFind.Tag = Mid(strFilter, 2)
24        End If
              
25        If Trim(Me.txtFind.Tag) = "" Then Exit Sub
26        strTmp = Split(Trim(Me.txtFind.Tag), ";")
          
          '��ʼ��������
27        For i = 0 To UBound(strTmp)
28            strSub = Split(strTmp(i), ",")
29            mblnFindEOF = True
              '��ѡ����Ŀ��Ӧ�ķ���
30            With vsfGroup
31                For lngRow = 1 To .Rows - 1
32                    If Val(.TextMatrix(lngRow, .ColIndex("ID"))) = Val(strSub(1)) And Val(.TextMatrix(lngRow, .ColIndex("���뵥ID"))) = Val(strSub(0)) Then
33                        .Row = lngRow
34                        .ShowCell .Row, 0
35                        vsfItem.Tag = ""
36                        Exit For
                      
37                    End If
38                Next
39            End With
              
              '��ѡ�з����������Ŀ
40            With Me.vsfItem
41                For lngRow = Val(.Tag) + 1 To .Rows - 1
42                    If Val(.TextMatrix(lngRow, .ColIndex("ID"))) = Val(strSub(2)) Then
43                        .Row = lngRow
44                        .ShowCell .Row, 0
45                        If lngRow >= .Rows - 1 Then
46                            .Tag = 0
47                        Else
48                            .Tag = lngRow
49                        End If
50                        Exit For
51                    End If
52                Next
53            End With
              
              '����Ѿ����ҹ�������
54            Me.txtFind.Tag = Replace(Me.txtFind.Tag, strSub(0) & "," & strSub(1) & "," & strSub(2), "")
              '������˵ķֺ�
55            If Mid(Me.txtFind.Tag, 1, 1) = ";" Then
56                Me.txtFind.Tag = Mid(Me.txtFind.Tag, 2)
57            End If
58            If Me.txtFind.Tag <> "" Then
59                If Mid(Me.txtFind.Tag, Len(Me.txtFind.Tag) - 1, 1) = ";" Then
60                    Me.txtFind.Tag = Mid(Me.txtFind.Tag, 1, Len(Me.txtFind.Tag) - 1)
61                End If
62            End If
              
63            Exit For
64        Next
          

65        mrsItem.Filter = ""
66        If Trim(Me.txtFind.Tag) = "" Then
67            vsfItem.Tag = ""
68            mblnFindEOF = False
69        End If


70        Exit Sub
txtFind_KeyPress_Error:
71        Call WriteErrLog("zlPublicHisCommLis", "frmcollectionApply", "ִ��(txtFind_KeyPress)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
72        Err.Clear
End Sub

Private Sub txtFind_LostFocus()
    Call setTxtTip(txtFind, "������롢����������û��س�����")
End Sub

Private Sub txtID_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then PressKey vbKeyTab
End Sub

Private Sub txtPatientDept_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then PressKey vbKeyTab
End Sub

Private Sub txtUnit_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then PressKey vbKeyTab
End Sub

Private Sub txt����_GotFocus()
    TxtSelAll txt����
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        PressKey vbKeyTab
    Else
        KeyAscii = FilterKeyAscii(KeyAscii, 99, "0123456789")
    End If
End Sub

Private Sub txt����1_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then PressKey vbKeyTab
End Sub

Private Sub txt����_GotFocus()
    TxtSelAll txt����
End Sub

Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then
        KeyCode = Asc(UCase(Chr(KeyCode)))
    End If
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If Len(Trim(txt����.Text)) = 0 Then Exit Sub
        Call txt����_Validate(False)
        cbo�Ա�.SetFocus
    End If
End Sub

Private Sub txt����_Validate(Cancel As Boolean)
          Dim rsTmp As New ADODB.Recordset, i As Integer
          Dim strField As String
          Dim strBarCode As String
          Dim rsDept As ADODB.Recordset, strSQL As String
          Dim strAge As String
          Dim aAge() As String

1         On Error GoTo txt����_Validate_Error

2         If Len(Trim(txt����)) = 0 Then Exit Sub
3         If txt���� = txt����.Tag Then Exit Sub

4         Call AdjustEditState(True)

5         mblnSaveAdvice = True
6         Cancel = Not StrIsValid(txt����.Text, txt����.MaxLength)

          '��ʼ������Ϣ
7         Set rsTmp = GetPatient(txt����)
8         strBarCode = txt����
9         If rsTmp.EOF Then
              '�Ǽ��²���
10            mlng����ID = 0
11            mstrKeys = ""
12            Me.txtBed = "": Me.txtID = ""
13            Me.txtPatientDept = "": Me.txtPatientDept.Tag = 0

              '���������Ժ�ڲ��ˣ����������
14            If InStr("+-*./", Left(Me.txt����.Text, 1)) > 0 Or mblnBarCode Then
15                Me.txt����.Text = "": Cancel = True
16                Exit Sub
17            End If
18            PatientType = 1
19        Else
20            Me.txt����.Text = NVL(rsTmp("����"))
21            Me.txt����.Text = "": Me.txt����1.Text = ""
22            strAge = IIf(IsNull(rsTmp("����")), "", rsTmp("����")): If Me.txt���� = "0" Then Me.txt���� = ""

23            strAge = Replace(strAge, "Сʱ", "ʱ")
24            strAge = Replace(strAge, "����", "��")

25            If Trim(Replace(Replace(Replace(Replace(Replace(strAge, "��", ""), "��", ""), "��", ""), "ʱ", ""), "��", "")) <> "" Then
26                If InStr(strAge, "����") > 0 Or InStr(strAge, "Ӥ��") > 0 Then
27                    Me.txt����.Text = ""
28                    Me.cboAge.Text = Trim(strAge)
29                Else
30                    strAge = Replace(Replace(Replace(Replace(Replace(strAge, "��", "��;"), "��", "��;"), "��", "��;"), "ʱ", "ʱ;"), "��", "��;")
31                    aAge = Split(strAge, ";")
32                    If UBound(aAge) = 1 Then
33                        Me.txt����.Text = Val(aAge(0))
34                        Me.cboAge.Text = Replace(Replace(Right(aAge(0), 1), "��", "����"), "ʱ", "Сʱ")
35                    Else
36                        Me.txt����.Text = Val(aAge(0))
37                        Me.cboAge.Text = Replace(Replace(Right(aAge(0), 1), "��", "����"), "ʱ", "Сʱ")
38                        Me.txt����1.Text = Val(aAge(1)) & Replace(Replace(Right(aAge(1), 1), "��", "����"), "ʱ", "Сʱ")
39                    End If
40                End If
41            Else
42                Me.txt����.Text = ""
43                Me.cboAge.ListIndex = 0
44            End If

45            If cboAge.ListIndex = -1 Then cboAge.ListIndex = 0
46            Me.cbo�Ա� = NVL(rsTmp("�Ա�"))    ' CombIndex(cbo�Ա�, Nvl(rsTmp("�Ա�")))

47            mlng����ID = NVL(rsTmp("����ID"), 0): PatientType = NVL(rsTmp("PatientType"), 1)

              '����Ĭ�Ͽ������ҡ�ҽ��
48            cbo��������.ListIndex = FindComboItem(cbo��������, NVL(rsTmp("���˿���"), 0))

              '���˵�λ
49            txtUnit.Text = NVL(rsTmp("������λ"))

50            strField = ""
51            strField = rsTmp.Fields("ҽ��").Name
52            If strField = "ҽ��" Then
53                Me.cboҽ��.Text = NVL(rsTmp("ҽ��"))
54                For i = 0 To Me.cboҽ��.ListCount - 1
55                    If Me.cboҽ��.List(i) Like NVL(rsTmp("ҽ��")) Then
56                        Me.cboҽ��.ListIndex = i
57                        Exit For
58                    End If
59                Next
60            End If

              '��ʾ���˿���
61            strSQL = "Select ���� From ���ű� Where ID=[1]"
62            Set rsDept = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption, CLng(NVL(rsTmp("���˿���"), 0)))
63            If rsDept.EOF Then
64                Me.txtPatientDept = "": Me.txtPatientDept.Tag = 0
65            Else
66                Me.txtPatientDept.Text = rsDept("����") & "": Me.txtPatientDept.Tag = NVL(rsTmp("���˿���"), 0)
67            End If
68            Me.txtID = rsTmp("סԺ��") & "": If Len(Me.txtID) = 0 Then Me.txtID = rsTmp("�����") & ""
69            Me.txtBed = NVL(rsTmp("��ǰ����"))

              '����Ǽǵ�Ĭ�Ͽ��ҡ�ҽ��
70            If Me.cbo��������.ListIndex = -1 And mlngReqDept > 0 Then
71                cbo��������.ListIndex = FindComboItem(cbo��������, mlngReqDept)
72            End If
73        End If

74        txt����.Tag = txt����.Text


75        Exit Sub
txt����_Validate_Error:
76        Call WriteErrLog("zlPublicHisCommLis", "frmcollectionApply", "ִ��(txt����_Validate)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
77        Err.Clear
End Sub

Private Sub InitDoctors(ByVal lng����ID As Long)
      '���ܣ���ȡ��ǰ���������а�����������Ա
          Dim rsTmp As New ADODB.Recordset
          Dim strSQL As String, i As Long
          
1         On Error GoTo InitDoctors_Error

2         Me.cboҽ��.Clear
          
          '����ҽ����ʿ
3         strSQL = _
              "Select Distinct A.ID,B.����ID,A.���,A.����,Upper(A.����) as ����," & _
              " C.��Ա����,Nvl(A.Ƹ�μ���ְ��,0) as ְ��" & _
              " From ��Ա�� A,������Ա B,��Ա����˵�� C" & _
              " Where A.ID=B.��ԱID And A.ID=C.��ԱID" & _
              " And C.��Ա���� IN('ҽ��') And B.����ID=[1] " & _
              " And (a.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or a.����ʱ�� Is Null) "
              
4         strSQL = strSQL & " Order by ����,��Ա���� Desc"
          
5         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption, lng����ID)
          
6         If Not rsTmp.EOF Then
7             For i = 1 To rsTmp.RecordCount
8                 cboҽ��.AddItem rsTmp!����
9                 cboҽ��.ItemData(cboҽ��.ListCount - 1) = rsTmp!����ID
                  
10                If rsTmp!ID = gUserInfo.ID And cboҽ��.ListIndex = -1 Then
11                    cboҽ��.ListIndex = cboҽ��.NewIndex
12                ElseIf cboҽ��.ListCount > 0 Then
13                    cboҽ��.ListIndex = 0
14                End If
15                rsTmp.MoveNext
16            Next
              
17            If cboҽ��.ListCount = 1 And cboҽ��.ListIndex = -1 Then cboҽ��.ListIndex = 0
18        End If


19        Exit Sub
InitDoctors_Error:
20        Call WriteErrLog("zlPublicHisCommLis", "frmcollectionApply", "ִ��(InitDoctors)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
21        Err.Clear
End Sub
Public Sub ShowMe(objFrm As Object)
    Me.Show vbModal, objFrm
End Sub

Private Function InitDepts() As Boolean
      '���ܣ���ʼ��סԺ�ٴ�����
          Dim rsTmp As New ADODB.Recordset
          Dim strSQL As String, i As Long
          Dim strOldText As String
          Dim intloop As Integer
              
1         On Error GoTo InitDepts_Error

2         strSQL = _
              " Select Distinct A.ID,A.����,A.����" & _
              " From ���ű� A,��������˵�� B " & _
              " Where B.����ID = A.ID " & _
              " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
              " And (B.�������� IN('����'))" & _
              " Order by A.����"
3         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption)
          
4         With Me.cboִ�п���
5             Do While Not rsTmp.EOF
6                 .AddItem NVL(rsTmp("����"))
7                 .ItemData(.NewIndex) = rsTmp("ID")
8                 rsTmp.MoveNext
9             Loop
10            If .ListCount > 0 Then
11                .ListIndex = 0
12            End If
13        End With
          
          
14        strOldText = Me.cbo��������.Text
15        Me.cbo��������.Clear
          
16        strSQL = _
              " Select Distinct A.ID,A.����,A.����" & _
              " From ���ű� A,��������˵�� B " & _
              " Where B.����ID = A.ID " & _
              " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL) " & _
              " And (B.�������� IN('�ٴ�','���'))" & _
              " Order by A.����"
17        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption)
          
18        For i = 1 To rsTmp.RecordCount
19            cbo��������.AddItem rsTmp!����
20            cbo��������.ItemData(cbo��������.NewIndex) = rsTmp!ID
              
21            rsTmp.MoveNext
22        Next
          
23        On Error Resume Next
24        Me.cbo��������.Text = strOldText
          
           '�Ա�
26        Set rsTmp = Nothing
27        Set rsTmp = GetDictData("�Ա�")
28        cbo�Ա�.Clear
29        If Not rsTmp Is Nothing Then
30            For intloop = 1 To rsTmp.RecordCount
31                cbo�Ա�.AddItem rsTmp!����
32                If rsTmp!ȱʡ = 1 Then
33                    cbo�Ա�.ItemData(cbo�Ա�.NewIndex) = 1
34                    cbo�Ա�.ListIndex = cbo�Ա�.NewIndex
35                End If
36                rsTmp.MoveNext
37            Next
38        End If
          
39        chkConcatenation.value = GetPara(Sel_His_DB, "�ɼ�����վ�Ǽ�", 100, 1211, 0)

40        InitDepts = True


41        Exit Function
InitDepts_Error:
42        Call WriteErrLog("zlPublicHisCommLis", "frmcollectionApply", "ִ��(InitDepts)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
43        Err.Clear

End Function

Private Sub AdjustEditState(blEnable As Boolean)
    '����:              �����༭״̬
    'Me.txt����.Enabled = blEnable
    cbo�Ա�.Enabled = blEnable
    txt����.Enabled = blEnable
    txt����1.Enabled = blEnable
    cboAge.Enabled = blEnable
    cbo��������.Enabled = blEnable
    cboҽ��.Enabled = blEnable
End Sub
Private Function GetPatient(strCode As String) As ADODB.Recordset
      '���ܣ���ȡ������Ϣ������ʾ�ò��˴��ڵ�ҽ��ʱ��
          Dim strSQL As String
          Dim strNO As String, str���� As String
          Dim strSeek As String
          
          
1         On Error GoTo GetPatient_Error

2         If BlnIsNumber(strCode) Then
          'Ԥ�����뵥������
3             mblnBarCode = True
4             strSQL = "Select Decode(A.��ǰ����id,Null,1,2) As PatientType,B.��ҳID,B.���˿���id As ���˿���,B.����ҽ�� As ҽ��," & mConst_������Ϣ_���� & _
                  " From ������Ϣ A,����ҽ����¼ B,����ҽ������ C Where A.����ID=B.����ID+0 And B.ID=C.ҽ��ID+0" & _
                  " And C.��������=[1]"
5             Set GetPatient = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption, strCode)
6             Exit Function
7         End If
8         mblnBarCode = False
          
9         strSeek = strCode
          '�жϵ�ǰ����ģʽ
10        If IsNumeric(strCode) And IsNumeric(Left(strCode, 1)) Then  'ˢ��
11            miInputType = 0
12        ElseIf (Left(strCode, 1) = "A" Or Left(strCode, 1) = "-") And IsNumeric(Mid(strCode, 2)) Then '����ID
13            miInputType = 1
14            strSeek = Mid(strCode, 2)
15        ElseIf (Left(strCode, 1) = "B" Or Left(strCode, 1) = "+") And IsNumeric(Mid(strCode, 2)) Then 'סԺ��
16            miInputType = 2
17            strSeek = Mid(strCode, 2)
18        ElseIf (Left(strCode, 1) = "D" Or Left(strCode, 1) = "*") And IsNumeric(Mid(strCode, 2)) Then '�����
19            miInputType = 3
20            strSeek = Mid(strCode, 2)
21        ElseIf Left(strCode, 1) = "G" Or Left(strCode, 1) = "." Then '�Һŵ�
22            miInputType = 4
23            strSeek = Mid(strCode, 2)
24        ElseIf Left(strCode, 1) = "/" Then '�շѵ��ݺ�
25            miInputType = 5
26            strSeek = Mid(strCode, 2)
27        ElseIf Not IsNumeric(Mid(strCode, 2)) Then '��������
28            miInputType = 6
29            strSeek = Replace(strCode, "(Ӥ��)", "")
30        End If
          
31        If miInputType = 0 Then 'ˢ��
32            strSQL = "Select Decode(A.��ǰ����id,Null,1,2) As PatientType,A.��ҳID,Decode(A.��ǰ����id,Null,Nvl(B.ִ�в���ID,0),A.��ǰ����id) As ���˿���,B.ִ���� As ҽ��," & mConst_������Ϣ_���� & _
                  " From ������Ϣ A,���˹Һż�¼ B Where A.���￨��=[1] And A.����ID=B.����ID(+) And A.�����=B.�����(+) and (b.����ID is null or (b.��¼���� =1 and b.��¼״̬ =1)) "
      '            " And (A.��ǰ����id IS NOT NULL Or NVL(B.ִ��״̬,1) IN (0,2))"
33        ElseIf miInputType = 1 Then '����ID
34            strSQL = "Select Decode(A.��ǰ����id,Null,1,2) As PatientType,A.��ҳID,Nvl(A.��ǰ����id,0) As ���˿���,'' ҽ��," & mConst_������Ϣ_���� & _
                  " From ������Ϣ A Where A.����ID=[2]"
35        ElseIf miInputType = 2 Then 'סԺ��
36            strSQL = "Select Decode(A.��ǰ����id,Null,1,2) As PatientType,A.��ҳID,Decode(A.��ǰ����id,Null,Nvl(B.��Ժ����ID,0),A.��ǰ����id) As ���˿���,B.סԺҽʦ As ҽ��," & mConst_������Ϣ_���� & _
                  " From ������Ϣ A,������ҳ B Where A.סԺ��=[2] And A.����ID=B.����ID" ' And A.��ǰ����id IS NOT NULL And B.��Ժ���� Is NULL"
37        ElseIf miInputType = 3 Then '�����
38            strSQL = "Select Decode(A.��ǰ����id,Null,1,2) As PatientType,A.��ҳID,Decode(A.��ǰ����id,Null,Nvl(B.ִ�в���ID,0),A.��ǰ����id) As ���˿���,B.ִ���� As ҽ��," & mConst_������Ϣ_���� & _
                  " From ������Ϣ A,���˹Һż�¼ B Where A.�����=[1] And A.����ID=B.����ID(+) And A.�����=B.�����(+) and (b.����ID is null or (b.��¼���� =1 and b.��¼״̬ =1)) "
      '            " And (A.��ǰ����id IS NOT NULL Or NVL(B.ִ��״̬,1) IN (0,2))"
39        ElseIf miInputType = 4 Then '�Һŵ�
40            strNO = GetFullNO(strSeek, 12)
41            strSQL = "Select 1 As PatientType,0 As ��ҳID,Nvl(B.ִ�в���ID,0) As ���˿���,B.ִ���� As ҽ��," & mConst_������Ϣ_���� & _
                  " From ������Ϣ A,������ü�¼ B " & _
                  " Where B.��¼����=4 And B.��¼״̬ IN(1,3) And B.NO=[3] And B.����ID=A.����ID"
42        ElseIf miInputType = 5 Then '�շѵ��ݺ�
43            strNO = GetFullNO(strSeek, 13): mstrNO = strNO
              
44            strSQL = "Select 1 As PatientType,0 As ��ҳID,B.��������ID As ���˿���,B.������ As ҽ��,B.����,B.�Ա�,B.����,a.סԺ��,a.��ǰ����," & _
                  "A.����ID,A.��λ�绰,A.������λ,A.��λ�ʱ�,A.��ͥ��ַ,A.��ͥ�绰,A.��ͥ��ַ�ʱ�,A.�����,A.���֤��,A.�ѱ�,A.ҽ�Ƹ��ʽ," & _
                  "A.����,A.����״��,A.����,A.ְҵ From ������Ϣ A,������ü�¼ B" & _
                  " Where Mod(B.��¼����,10)=1 And B.��¼״̬ IN(1,3) And B.NO=[3] And B.����ID=A.����ID(+) Order By B.����ID" ' And B.ҽ����� Is Null"
45        Else '��������
46            strSQL = "Select Decode(A.��ǰ����id,Null,1,2) As PatientType,A.��ҳID,Nvl(A.��ǰ����id,0) As ���˿���,'' ҽ��," & mConst_������Ϣ_���� & _
                  " From ������Ϣ A Where A.����=[1] and 1 = 2 " '�������������Ĳ��˵��²��˴���
47        End If
          
48        Set GetPatient = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption, strSeek, Val(strSeek), strNO)


49        Exit Function
GetPatient_Error:
50        Call WriteErrLog("zlPublicHisCommLis", "frmcollectionApply", "ִ��(GetPatient)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
51        Err.Clear

End Function

Private Function GetDictData(strDict As String) As ADODB.Recordset
'���ܣ���ָ�����ֵ��ж�ȡ����
'������strDict=�ֵ��Ӧ�ı���
    Dim rsTmp As New ADODB.Recordset
    Dim strSQL As String
    
        
    strSQL = "Select ����,����,Nvl(ȱʡ��־,0) as ȱʡ From " & strDict & " Order by ����"
    Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption)
    If Not rsTmp.EOF Then Set GetDictData = rsTmp

End Function
Private Function GetFullNO(ByVal strNO As String, ByVal intNum As Integer) As String
      '���ܣ����û�����Ĳ��ݵ��ţ�����ȫ���ĵ��š�
      '������intNum=��Ŀ���,Ϊ0ʱ�̶��������
          Dim rsTmp As New ADODB.Recordset
          Dim strSQL As String, intType As Integer
          Dim curDate As Date
          
1         On Error GoTo GetFullNO_Error

2         If Len(strNO) >= 8 Then
3             GetFullNO = Right(strNO, 8)
4             Exit Function
5         ElseIf Len(strNO) = 7 Then
6             GetFullNO = PreFixNO & strNO
7             Exit Function
8         ElseIf intNum = 0 Then
9             GetFullNO = PreFixNO & Format(Right(strNO, 7), "0000000")
10            Exit Function
11        End If
12        GetFullNO = strNO
          
13        strSQL = "Select ��Ź���,Sysdate as ���� From ������Ʊ� Where ��Ŀ���=" & intNum
14        Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption)
15        If Not rsTmp.EOF Then
16            intType = NVL(rsTmp!��Ź���, 0)
17            curDate = rsTmp!����
18        End If

19        If intType = 1 Then
              '���ձ��
20            strSQL = Format(CDate("1992-" & Format(rsTmp!����, "MM-dd")) - CDate("1992-01-01") + 1, "000")
21            GetFullNO = PreFixNO & strSQL & Format(Right(strNO, 4), "0000")
22        Else
              '������
23            GetFullNO = PreFixNO & Format(Right(strNO, 7), "0000000")
24        End If


25        Exit Function
GetFullNO_Error:
26        Call WriteErrLog("zlPublicHisCommLis", "frmcollectionApply", "ִ��(GetFullNO)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
27        Err.Clear

End Function

Private Function SelectCap(Optional ByVal lngItemid As Long = 0) As ADODB.Recordset
      '��ȡ�ɼ���ʽ
          Dim strSQL As String
          Dim rsTmp As New ADODB.Recordset



1         On Error GoTo SelectCap_Error

2         strSQL = "Select Distinct A.ID,A.����,A.���� " + _
                   "From ������ĿĿ¼ A,�����÷����� D Where A.ID=D.�÷�ID" + _
                 " And A.���='E' And A.��������='6'" & _
                 " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL) " + _
                 " And A.������� IN(" & PatientType & ",3) And Nvl(A.�����Ա�,0) IN (" + _
                   IIf(Me.cbo�Ա�.Text Like "*��*", "1,0)", "2,0)") + _
                 " And Nvl(A.ִ��Ƶ��,0) IN(0,1)" + _
                 " And D.��ĿID=" & lngItemid
3         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption)
4         If rsTmp.EOF Then
5             strSQL = "Select Distinct A.ID,A.����,A.���� " + _
                       "From ������ĿĿ¼ A Where " + _
                     " A.���='E' And A.��������='6'" & _
                     " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL) " + _
                     " And A.������� IN(" & PatientType & ",3) And Nvl(A.�����Ա�,0) IN (" + _
                       IIf(Me.cbo�Ա�.Text Like "*��*", "1,0)", "2,0)") + _
                     " And Nvl(A.ִ��Ƶ��,0) IN(0,1)"
6             Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption)
7         End If

8         If Not rsTmp.EOF Then Set SelectCap = rsTmp


9         Exit Function
SelectCap_Error:
10        Call WriteErrLog("zlPublicHisCommLis", "frmcollectionApply", "ִ��(SelectCap)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
11        Err.Clear

End Function

Private Sub AdviceSet�������(ByVal int���� As Integer, ByVal strDataIDs As String)
      '���ܣ�1.��������ָ����������Ŀ�Ĳ�λ��,�����������������Ŀ���޸Ĳ�λ
      '      2.��������ָ��������Ŀ�ĸ���������������Ŀ��,����������������Ŀ��������Ŀ�ĸ���������������Ŀ
      '������int����=1=�����鲿λ��Ŀ,2=������������������Ŀ
      '      strDataIDs=���:������鲿λ��Ϣ,����:��������������������Ŀ��Ϣ,���п���û�и�������������
          Dim strSQL As String

                  
          '���������Ŀ
1         On Error GoTo AdviceSet�������_Error

2         strDataIDs = Mid(strDataIDs, 1, InStr(strDataIDs, ";") - 1)
          
3         If strDataIDs <> "" Then
4             If Not mrsRelativeAdvice Is Nothing Then
5                 mrsRelativeAdvice.Close
6             Else
7                 Set mrsRelativeAdvice = New ADODB.Recordset
8             End If
9             strSQL = "Select ID,����,����,nvl(�걾��λ,' ') As �걾��λ," + _
              "���,nvl(�Ƽ�����,0) As �Ƽ�����,nvl(ִ�п���,0) As ִ�п���,�������� From ������ĿĿ¼ Where ID IN(" & strDataIDs & ")"
10            Set mrsRelativeAdvice = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption)
11        Else
12            If Not mrsRelativeAdvice Is Nothing Then mrsRelativeAdvice.Close: Set mrsRelativeAdvice = Nothing
13        End If


14        Exit Sub
AdviceSet�������_Error:
15        Call WriteErrLog("zlPublicHisCommLis", "frmcollectionApply", "ִ��(AdviceSet�������)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
16        Err.Clear

End Sub

Private Function Get�����������(ByVal int���� As Integer, ByVal txtMainAdvice As String) As String
      '���ܣ��������ɼ���������ݵ�ҽ������
      '������int����=1=�����鲿λ��Ŀ,2=������������������Ŀ
          Dim lngBegin As Long
          Dim strTmp As String

1         On Error GoTo Get�����������_Error

2         If mrsRelativeAdvice Is Nothing Or int���� = 1 Then Get����������� = txtMainAdvice: Exit Function
              
3         mrsRelativeAdvice.MoveFirst
4         Do While Not mrsRelativeAdvice.EOF
5             If Len(Trim(mrsRelativeAdvice("����"))) > 0 Then
6                 strTmp = strTmp & "," & mrsRelativeAdvice("����")
7             End If
              
8             mrsRelativeAdvice.MoveNext
9         Loop
          
10        If strTmp <> "" Then
11            Get����������� = IIf(Len(Trim(txtMainAdvice)) = 0, "", txtMainAdvice & " �� ") & Mid(strTmp, 2)
12        Else
13            Get����������� = txtMainAdvice
14        End If


15        Exit Function
Get�����������_Error:
16        Call WriteErrLog("zlPublicHisCommLis", "frmcollectionApply", "ִ��(Get�����������)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
17        Err.Clear
End Function

'���ҽ�����ݵĺϷ���
Private Function ValidAdvice() As Boolean
1         On Error GoTo ValidAdvice_Error

2         ValidAdvice = True
          
3         On Error Resume Next
4         If txt����.Text = "" Then
5             ValidAdvice = False
6             MsgBox "�����벡�˵�������", vbInformation, Me.Caption: DoEvents
      '        mintFocusItem = FocusItem.����
7             txt����.SetFocus: Exit Function
8         End If
          
      '    If Len(Trim(Me.txtҽ������)) = 0 Then
      '        ValidAdvice = False
      '        MsgBox "��������������Ŀ��", vbInformation, Me.Caption: DoEvents
      ''        mintFocusItem = FocusItem.ҽ������
      '        Me.txtҽ������.SetFocus: Exit Function
      '    End If
9         If Me.cbo��������.ListIndex = -1 Then
10            ValidAdvice = False
11            MsgBox "��ָ���������ң�", vbInformation, Me.Caption: DoEvents
      '        mintFocusItem = FocusItem.��������
12            Me.cbo��������.SetFocus: Exit Function
13        End If
14        If Me.cboִ�п���.ListIndex = -1 Then
15            ValidAdvice = False
16            MsgBox "��ָ��ִ�п���!", vbInformation, Me.Caption: DoEvents
17            Me.cboִ�п���.SetFocus: Exit Function
18        End If
19        If Len(Trim(Me.cboҽ��.Text)) = 0 Then
20            ValidAdvice = False
21            MsgBox "��ָ������ҽ����", vbInformation, Me.Caption: DoEvents
      '        mintFocusItem = FocusItem.ҽ��
22            Me.cboҽ��.SetFocus: Exit Function
23        End If


24        Exit Function
ValidAdvice_Error:
25        Call WriteErrLog("zlPublicHisCommLis", "frmcollectionApply", "ִ��(ValidAdvice)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
26        Err.Clear
End Function

Private Function SaveAdviceData(ByRef strNewAdvice As String) As Boolean
          Dim strSQL As String, strDate As String, strNO As String
          Dim lngAdviceID As Long, lngTmpID As Long, lngSendNO As Long
          Dim iMaxSeq As Integer, iSendSeq As Integer
          Dim rsTmp As New ADODB.Recordset
          Dim lng��������ID As Long, strDoctor As String, i As Integer
          Dim strִ�п���id As String, strִ�п���ID1 As String
          Dim tmpstr��� As String, tmplngClinicID As Long, tmpint�Ƽ����� As Integer
          Dim lngJ As Long, strCostType As String

          Dim strAge As String
          Dim strInfo As String
          Dim lngTmp As Long

1         On Error GoTo SaveAdviceData_Error

          '���没����Ϣ
2         strDate = "To_Date('" & Format(Currentdate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
3         If PatientType = 1 Then    '���ﲡ��
4             If mlng����ID > 0 Then    '���еĲ���
                  '            strSQL = _
                               "zl_�ҺŲ��˲���_INSERT(3," & mlng����ID & ",Null," & _
                               "'',''," & _
                               "'" & txt����.Text & "','" & NeedName(cbo�Ա�.Text) & "','" & txt����.Text & Me.cboAge.Text & Me.txt����1.Text & "'," & _
                               "'�Է�','�Է�'," & _
                               "'','',''," & _
                               "'','','',0,'','','','',''," & strDate & ",NULL)"
5             Else    '�²���
6                 If txt����.Locked = False Then
7                     strAge = txt����.Text
8                     If IsNumeric(strAge) Then strAge = strAge & cboAge.Text & txt����1.Text
9                     strInfo = CheckAge(strAge)
10                    If InStr(1, strInfo, "|") > 0 Then
11                        lngTmp = Val(Split(strInfo, "|")(0))    '1��ֹ,0��ʾ
12                        strInfo = Split(strInfo, "|")(1)
13                        If lngTmp = 1 Then
14                            MsgBox strInfo, vbInformation, Me.Caption
15                            If txt����.Enabled And txt����.Visible Then txt����.SetFocus: Exit Function
16                        End If
17                    End If
18                End If
                  '��ӻ�ȡĬ�Ϸѱ�
19                strSQL = "select ����,ȱʡ��־ from �ѱ� order by ����"
20                Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "mdlLisWork")
21                Do While Not rsTmp.EOF
22                    lngJ = lngJ + 1
23                    If lngJ = 1 Then
24                        strCostType = rsTmp("����")
25                    End If
26                    If rsTmp("ȱʡ��־") = 1 Then
27                        strCostType = rsTmp("����")
28                        Exit Do
29                    End If
30                    rsTmp.MoveNext
31                Loop
32                If strCostType = "" Then strCostType = "�Է�"

33                mlng����ID = GetNextNo(Sel_His_DB, 1)
34                ReDim Preserve mstrSQLPro(UBound(mstrSQLPro) + 1)
35                mstrSQLPro(UBound(mstrSQLPro)) = "zl_�ҺŲ��˲���_INSERT(1," & mlng����ID & ",Null," & _
                                                   "'',''," & _
                                                   "'" & txt����.Text & "','" & NeedName(cbo�Ա�.Text) & "','" & txt����.Text & Me.cboAge.Text & Me.txt����1.Text & "'," & _
                                                   "'" & strCostType & "','" & strCostType & "'," & _
                                                   "'','',''," & _
                                                   "'','','" & Me.txtUnit.Text & "',0,'','','','',''," & strDate & ",NULL)"
36            End If
37        End If
          '����ҽ��������
38        lngAdviceID = GetNextId("����ҽ����¼")
39        iMaxSeq = 0

40        lng��������ID = Me.cbo��������.ItemData(Me.cbo��������.ListIndex)
41        strDoctor = NeedName(Me.cboҽ��.Text)

42        If mrsRelativeAdvice.RecordCount = 0 Then
43            strִ�п���id = mlngDeptID
44        Else
              'PatientType
45            If mlng����ID > 0 Then
46                strSQL = "select  ִ�п���ID from  ����ִ�п��� where ������Դ = [1] and ������ĿID = [2] "
47            End If
48            mrsRelativeAdvice.MoveFirst
49            Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, Me.Caption, PatientType, CLng(mrsRelativeAdvice("Id")))
50            If Not rsTmp.EOF Then strִ�п���id = Val(NVL(rsTmp("ִ�п���ID")))
51        End If

          'ѡ����ִ�п��Ұ�ִ�п��ҽ���
52        If Me.cboִ�п���.Text <> "" Then
53            strִ�п���id = Me.cboִ�п���.ItemData(Me.cboִ�п���.ListIndex)
54        End If

55        iSendSeq = 1
          '������Ŀ���ɼ���ʽ��Ϊ��ҽ��
56        tmplngClinicID = mlngCapID
          'ȡ�ɼ���ʽ��ִ�в���
57        strִ�п���ID1 = gUserInfo.DeptID

58        lngSendNO = GetNextNo(Sel_His_DB, 10)
59        strNO = GetNextNo(Sel_His_DB, IIf(PatientType = 2, 14, 13))

          '�������ҽ��
60        If Not mrsRelativeAdvice Is Nothing Then
61            i = 2
62            mrsRelativeAdvice.MoveFirst
63            Do While Not mrsRelativeAdvice.EOF
64                lngTmpID = GetNextId("����ҽ����¼")
65                With mrsRelativeAdvice
66                    strNewAdvice = strNewAdvice & "," & lngTmpID
67                    ReDim Preserve mstrSQLPro(UBound(mstrSQLPro) + 1)
68                    mstrSQLPro(UBound(mstrSQLPro)) = "ZL_����ҽ����¼_Insert(" & lngTmpID & "," & lngAdviceID & "," & _
                                                       (iMaxSeq + i) & ",3," & mlng����ID & ",NULL," & _
                                                       "0,1," & _
                                                       "1,'" & .Fields("���") & "'," & _
                                                       .Fields("ID") & ",NULL,NULL,NULL,NULL," & _
                                                       "'" & Replace(.Fields("����"), "'", "''") & "',''," & _
                                                       "'" & .Fields("�걾��λ") & "','һ����',NULL,NULL,'',NULL," & _
                                                       .Fields("�Ƽ�����") & "," & _
                                                       strִ�п���id & "," & _
                                                       .Fields("ִ�п���") & ",0," & strDate & ",NULL," & _
                                                       IIf(Val(Me.txtPatientDept.Tag) = 0, lng��������ID, Val(Me.txtPatientDept.Tag)) & "," & lng��������ID & ",'" & strDoctor & "'," & _
                                                       "Sysdate,'',Null)"
69                    iSendSeq = iSendSeq + 1

70                    ReDim Preserve mstrSQLPro(UBound(mstrSQLPro) + 1)
71                    mstrSQLPro(UBound(mstrSQLPro)) = "ZL_����ҽ������_Insert(" & _
                                                       lngTmpID & "," & lngSendNO & "," & PatientType & ",'" & strNO & "'," & _
                                                       iSendSeq & ",NULL,NULL,NULL," & _
                                                       "Sysdate+1/(24*3600)," & _
                                                       "0," & strִ�п���id & ",0,0)"
72                    i = i + 1
73                    .MoveNext
74                End With
75            Loop
76        End If
          '��������Ĳɼ���ʽ�ŵ����
77        iMaxSeq = iMaxSeq + 1
78        strNewAdvice = strNewAdvice & "," & lngAdviceID
79        ReDim Preserve mstrSQLPro(UBound(mstrSQLPro) + 1)
80        mstrSQLPro(UBound(mstrSQLPro)) = "ZL_����ҽ����¼_Insert(" & lngAdviceID & ",NULL," & _
                                           iMaxSeq & ",3," & mlng����ID & ",NULL," & _
                                           "0,1," & _
                                           "1,'E'," & mlngCapID & ",NULL,NULL,NULL,NULL," & _
                                           "'" & Replace(mstrItemSel, "'", "''") & "',''," & _
                                           "'','һ����',NULL,NULL,'',NULL,2," & _
                                           strִ�п���ID1 & ",3,0," & strDate & ",NULL," & _
                                           IIf(Val(Me.txtPatientDept.Tag) = 0, lng��������ID, Val(Me.txtPatientDept.Tag)) & "," & lng��������ID & ",'" & strDoctor & "'," & _
                                           "Sysdate,'',Null)"
81        iSendSeq = iSendSeq + 1
          '������ҽ��
82        ReDim Preserve mstrSQLPro(UBound(mstrSQLPro) + 1)
83        mstrSQLPro(UBound(mstrSQLPro)) = "ZL_����ҽ������_Insert(" & _
                                           lngAdviceID & "," & lngSendNO & "," & PatientType & ",'" & strNO & "'," & _
                                           iSendSeq & ",NULL,NULL,NULL," & _
                                           "Sysdate+1/(24*3600)," & _
                                           "0," & strִ�п���id & ",0,1)"


84        If strNewAdvice <> "" Then strNewAdvice = Mid(strNewAdvice, 2)
85        SaveAdviceData = True

86        Exit Function
SaveAdviceData_Error:
87        Call WriteErrLog("zlPublicHisCommLis", "frmcollectionApply", "ִ��(SaveAdviceData)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
88        Err.Clear

End Function

Private Function PreFixNO(Optional curDate As Date = #1/1/1900#) As String
'���ܣ����ش�д�ĵ��ݺ���ǰ׺
    If curDate = #1/1/1900# Then
        PreFixNO = CStr(CInt(Format(Currentdate, "YYYY")) - 1990)
    Else
        PreFixNO = CStr(CInt(Format(curDate, "YYYY")) - 1990)
    End If
    PreFixNO = IIf(CInt(PreFixNO) < 10, PreFixNO, Chr(55 + CInt(PreFixNO)))
End Function

Private Sub TxtSelAll(ByVal objTxt As Object)
    With objTxt
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Function NeedName(strList As String) As String
    NeedName = Mid(strList, InStr(strList, "-") + 1)
End Function

Private Function GetControlRect(ByVal lnghwnd As Long, Optional ByVal blnTwip As Boolean = True) As RECT
'���ܣ���ȡָ���ؼ�����Ļ�е�λ��(Twip/Pixel)
'���أ�blnTwip=True-����Twip��λ��False-�������ص�λ
    Dim vRect As RECT
    Call GetWindowRect(lnghwnd, vRect)
    If blnTwip Then
        vRect.Left = vRect.Left * Screen.TwipsPerPixelX
        vRect.Right = vRect.Right * Screen.TwipsPerPixelX
        vRect.Top = vRect.Top * Screen.TwipsPerPixelY
        vRect.Bottom = vRect.Bottom * Screen.TwipsPerPixelY
    End If
    GetControlRect = vRect
End Function

Private Function CboLocate(ByVal cboobj As Object, ByVal strValue As String, Optional ByVal blnItem As Boolean = False) As Boolean
'�������ã�ʹ��Cbo.SeekIndex����
'blnItem:True-��ʾ����ItemData��ֵ��λ������;False-��ʾ�����ı������ݶ�λ������
    Dim lngLocate As Long
    CboLocate = False
    For lngLocate = 0 To cboobj.ListCount - 1
        If blnItem Then
            If cboobj.ItemData(lngLocate) = Val(strValue) Then
                cboobj.ListIndex = lngLocate
                CboLocate = True
                Exit For
            End If
        Else
            If Mid(cboobj.List(lngLocate), InStr(1, cboobj.List(lngLocate), "-") + 1) = strValue Then
                cboobj.ListIndex = lngLocate
                CboLocate = True
                Exit For
            End If
        End If
    Next
End Function

Private Function FilterKeyAscii(ByVal KeyAscii As Long, ByVal bytMode As Byte, Optional ByVal KeyCustom As String) As Long
            
    FilterKeyAscii = KeyAscii
    
    If Chr(KeyAscii) = "'" Then
        FilterKeyAscii = 0
        Exit Function
    End If
    
    If KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Or KeyAscii = vbKeyBack Then
        Exit Function
    End If
    
    Select Case bytMode
    Case 1      '������
        If InStr("0123456789<>", Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    Case 2      '��С��
        If InStr("0123456789.-<>+Ee", Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    Case 99
        If InStr(KeyCustom, Chr(KeyAscii)) = 0 Then FilterKeyAscii = 0
    End Select
    
End Function

Private Function StrIsValid(ByVal strInput As String, Optional ByVal intMax As Integer = 0) As Boolean
    '����ַ����Ƿ��зǷ��ַ�������ṩ���ȣ��Գ��ȵĺϷ���Ҳ����⡣
    If InStr(strInput, "'") > 0 Then
        MsgBox "���������ݺ��зǷ��ַ���", vbExclamation, Me.Caption
        Exit Function
    End If
    If intMax > 0 Then
        If LenB(StrConv(strInput, vbFromUnicode)) > intMax Then
            MsgBox "���������ݲ��ܳ���" & Int(intMax / 2) & "������" & "��" & intMax & "���ַ���", vbExclamation, Me.Caption
            Exit Function
        End If
    End If
    StrIsValid = True
End Function

Private Function FindComboItem(objCombox As Object, ByVal lngFind As Long) As Integer
    Dim i As Integer
    
    For i = 0 To objCombox.ListCount - 1
        If objCombox.ItemData(i) = lngFind Then Exit For
    Next
    If i > objCombox.ListCount - 1 Then i = -1
    
    FindComboItem = i
End Function

Private Function BlnIsNumber(ByVal strCode As String) As Boolean
    '���֣��������ж�
     If IsNumeric(strCode) And Len(strCode) >= 12 And InStr("*-+./", Mid(strCode, 1, 1)) = 0 Then
        BlnIsNumber = True
     Else
        BlnIsNumber = False
     End If
End Function

Public Function CheckAge(ByVal strAge As String, Optional ByVal strBirthDay As String = "") As String
    '����:����Ϸ��Լ��
    Dim strSQL As String
    Dim rsTemp As New ADODB.Recordset
    
    strBirthDay = Format(strBirthDay, "YYYY-MM-DD HH:mm")
    If IsDate(strBirthDay) Then
        strSQL = "select Zl_Age_Check([1],[2]) From dual"
        Set rsTemp = ComOpenSQL(Sel_His_DB, strSQL, "Zl_Age_Check", strAge, CDate(strBirthDay))
    Else
        strSQL = "select Zl_Age_Check([1]) From dual"
        Set rsTemp = ComOpenSQL(Sel_His_DB, strSQL, "Zl_Age_Check", strAge)
    End If
    CheckAge = NVL(rsTemp.Fields(0).value)

End Function

Private Sub vsfGroup_RowColChange()
          Dim lngRow As Long
          Dim lngCol As Long
          Dim lngGroupId As Long      '����ID
          Dim lngAppID As Long        '���뵥ID
          Dim strSampleType As String '�걾����
          Dim strErr As String
          
1         On Error GoTo vsfGroup_RowColChange_Error
          
          '��ȡѡ��ı걾
2         strSampleType = IIf(Trim(Me.cboSampleType.Text) = "���б걾", "", Trim(Me.cboSampleType.Text))

3         With Me.vsfGroup
4             lngRow = .Row
5             lngCol = .Col
6             If lngRow <= 1 Or lngCol < 0 Then Exit Sub
7             If .Cell(flexcpFontBold, lngRow, lngCol) = True Then Exit Sub
8             If mrsItem Is Nothing Then Exit Sub
9             If mrsItem.RecordCount <= 0 Then Exit Sub
10            mrsItem.MoveFirst
11            lngAppID = Val(.TextMatrix(lngRow, .ColIndex("���뵥ID")))      '���뵥ID
12            lngGroupId = Val(.TextMatrix(lngRow, .ColIndex("ID")))          '����ID
13            mrsItem.Filter = "���뵥ID=" & lngAppID & " and ����ID=" & lngGroupId & IIf(strSampleType = "", "", " and �걾='" & strSampleType & "'")
14            If vfgLoadFromRecord(vsfItem, mrsItem, strErr) = False Then
15                If strErr <> "" Then
16                    MsgBox strErr, vbInformation, Me.Caption
17                    mrsItem.Filter = ""
18                    Exit Sub
19                End If
20            End If
              
21        End With
22        With Me.vsfItem
23            .ExtendLastCol = True
24            .ColDataType(.ColIndex("ѡ��")) = flexDTBoolean
25            .ExplorerBar = flexExSortShow
              
              
26            .ColHidden(.ColIndex("ID")) = True
27            .ColHidden(.ColIndex("���뵥ID")) = True
28            .ColHidden(.ColIndex("����ID")) = True
29            .ColHidden(.ColIndex("���Ʊ���")) = True
30            .ColHidden(.ColIndex("�걾")) = True
              
31            .ColWidth(.ColIndex("ѡ��")) = 500
32            .ColWidth(.ColIndex("����")) = 1000
33            .ColWidth(.ColIndex("����")) = 4000
34            .ColWidth(.ColIndex("����")) = 1500
              
35            Call CheckSelItem
              
36        End With
              
37        mrsItem.Filter = ""


38        Exit Sub
vsfGroup_RowColChange_Error:
39        Call WriteErrLog("zlPublicHisCommLis", "frmCollectionApply", "ִ��(vsfGroup_RowColChange)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
40        Err.Clear

End Sub

Private Sub VSFItem_Click()
          Dim lngRow As Long
          Dim lngCol As Long
          
1         On Error GoTo VSFItem_Click_Error

2         With Me.vsfItem
3             lngRow = .MouseRow
4             lngCol = .MouseCol
5             .Editable = flexEDNone
              
6             If lngRow > 0 And lngCol = .ColIndex("ѡ��") Then
7                 If .Cell(flexcpChecked, lngRow, lngCol) = 1 Then
8                     .Cell(flexcpChecked, lngRow, lngCol) = 0
9                     Call selOrDelItem(2, lngRow)
10                Else
11                    .Cell(flexcpChecked, lngRow, lngCol) = 1
12                    Call selOrDelItem(1, lngRow)
13                End If
14            End If
15        End With


16        Exit Sub
VSFItem_Click_Error:
17        Call WriteErrLog("zlPublicHisCommLis", "frmcollectionApply", "ִ��(VSFItem_Click)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
18        Err.Clear
End Sub

Private Sub VSFSeled_DblClick()
    With Me.VSFSeled
        If .Row < 0 Or .Col < 0 Then Exit Sub
        Call selOrDelItem(3, .Row)
        Call CheckSelItem
    End With
End Sub


'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/4/27
'��    ��:ѡ�����ȡ��ѡ����Ŀ
'��    ��:
'           intType     1=���ѡ��,2=���ȡ��ѡ��3=˫��VSFSeledȡ��ѡ��
'           lngSelRow   ��ѡ�����
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Private Sub selOrDelItem(ByVal intType As Integer, ByVal lngSelRow As Long)
          Dim lngRow As Long
          
1         On Error GoTo selOrDelItem_Error

2         With VSFSeled
3             If intType = 1 Then
4                 .Rows = .Rows + 1
5                 .TextMatrix(.Rows - 1, .ColIndex("ID")) = vsfItem.TextMatrix(lngSelRow, vsfItem.ColIndex("ID"))
6                 .TextMatrix(.Rows - 1, .ColIndex("����")) = vsfItem.TextMatrix(lngSelRow, vsfItem.ColIndex("����"))
7                 .TextMatrix(.Rows - 1, .ColIndex("���Ʊ���")) = vsfItem.TextMatrix(lngSelRow, vsfItem.ColIndex("���Ʊ���"))
8                 .TextMatrix(.Rows - 1, .ColIndex("oldid")) = GetOldID(.TextMatrix(.Rows - 1, .ColIndex("���Ʊ���")))
9                 .TextMatrix(.Rows - 1, .ColIndex("oldName")) = GetOldName(.TextMatrix(.Rows - 1, .ColIndex("���Ʊ���")))
10                .TextMatrix(.Rows - 1, .ColIndex("�걾")) = vsfItem.TextMatrix(lngSelRow, vsfItem.ColIndex("�걾"))
11            ElseIf intType = 2 Then
12                For lngRow = 0 To .Rows - 1
13                    If Val(.TextMatrix(lngRow, .ColIndex("ID"))) = vsfItem.TextMatrix(lngSelRow, vsfItem.ColIndex("ID")) Then
14                        .RemoveItem lngRow
15                        Exit Sub
16                    End If
17                Next
18            ElseIf intType = 3 Then
19                .RemoveItem lngSelRow
20            End If
              '���ձ걾����
21            If .Rows > 0 Then .Cell(flexcpSort, .FixedRows, .ColIndex("�걾"), .Rows - 1, .ColIndex("�걾")) = 2
22        End With


23        Exit Sub
selOrDelItem_Error:
24        Call WriteErrLog("zlPublicHisCommLis", "frmCollectionApply", "ִ��(selOrDelItem)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
25        Err.Clear
End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/4/27
'��    ��:�л�����ʱ��ѡ�Ѿ�ѡ�����Ŀ
'��    ��:
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Private Sub CheckSelItem()
          Dim lngLoop As Long
          Dim lngRow As Long
          
          '��ȡ��ѡ��
1         On Error GoTo CheckSelItem_Error

2         With vsfItem
3             .Cell(flexcpChecked, 1, .ColIndex("ѡ��"), .Rows - 1, .ColIndex("ѡ��")) = 0
4         End With
          
          '��ѡ��
5         With Me.VSFSeled
6             For lngLoop = 0 To .Rows - 1
7                 For lngRow = 1 To vsfItem.Rows - 1
8                     If vsfItem.TextMatrix(lngRow, vsfItem.ColIndex("ID")) = .TextMatrix(lngLoop, .ColIndex("ID")) Then
9                         vsfItem.Cell(flexcpChecked, lngRow, vsfItem.ColIndex("ѡ��")) = 1
10                    End If
11                Next
12            Next
13        End With


14        Exit Sub
CheckSelItem_Error:
15        Call WriteErrLog("zlPublicHisCommLis", "frmCollectionApply", "ִ��(CheckSelItem)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
16        Err.Clear
End Sub

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/4/27
'��    ��:ͨ�����Ʊ����ȡ�ϰ���ĿID
'��    ��:
'           strCode     ���Ʊ���
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Private Function GetOldID(ByVal strCode As String) As Long
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          
1         On Error GoTo GetOldID_Error

2         strSQL = "select ID from ������ĿĿ¼ where ����=[1]"
3         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "������ĿĿ¼", strCode)
4         If rsTmp.RecordCount > 0 Then GetOldID = Val(rsTmp("ID") & "")


5         Exit Function
GetOldID_Error:
6         Call WriteErrLog("zlPublicHisCommLis", "frmcollectionApply", "ִ��(GetOldID)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
7         Err.Clear
End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/4/27
'��    ��:ͨ�����Ʊ����ȡ�ϰ���Ŀ����
'��    ��:
'           strCode     ���Ʊ���
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Private Function GetOldName(ByVal strCode As String) As String
          Dim strSQL As String
          Dim rsTmp As ADODB.Recordset
          
1         On Error GoTo GetOldID_Error

2         strSQL = "select ���� from ������ĿĿ¼ where ����=[1]"
3         Set rsTmp = ComOpenSQL(Sel_His_DB, strSQL, "������ĿĿ¼", strCode)
4         If rsTmp.RecordCount > 0 Then GetOldName = rsTmp("����") & ""


5         Exit Function
GetOldID_Error:
6         Call WriteErrLog("zlPublicHisCommLis", "frmcollectionApply", "ִ��(GetOldID)ʱ��������,�����:" & Err.Number & " ����ԭ��:" & Err.Description & " �����У�" & Erl, True)
7         Err.Clear
End Function

'---------------------------------------------------------------------------------------
'��    ��:������
'����ʱ��:2018/3/21
'��    ��:�����ı������ʾ��
'��    ��:
'��    ��:
'��    ��:
'---------------------------------------------------------------------------------------
Public Sub setTxtTip(objTxt As TextBox, Optional ByVal strTip As String)
    On Error Resume Next
    With objTxt
        If .Text <> "" Then Exit Sub
        .ToolTipText = strTip
        .Text = strTip
        .ForeColor = &H80000002
        .Tag = "T"
    End With
End Sub



