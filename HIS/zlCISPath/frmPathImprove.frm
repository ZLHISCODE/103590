VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPathImprove 
   BackColor       =   &H80000005&
   Caption         =   "�����Ľ�"
   ClientHeight    =   9390
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15420
   DrawStyle       =   1  'Dash
   HasDC           =   0   'False
   Icon            =   "frmPathImprove.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   9390
   ScaleWidth      =   15420
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraContent 
      BackColor       =   &H80000005&
      Caption         =   "�ſ�"
      Height          =   1695
      Left            =   120
      OLEDropMode     =   1  'Manual
      TabIndex        =   24
      Top             =   840
      Width           =   15015
      Begin VSFlex8Ctl.VSFlexGrid vsStep 
         Height          =   1095
         Left            =   7200
         TabIndex        =   8
         Top             =   480
         Width           =   5655
         _cx             =   9975
         _cy             =   1931
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
         BackColorFixed  =   13430215
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16764057
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   32768
         GridColorFixed  =   32768
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
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   6
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   4000
         ColWidthMin     =   500
         ColWidthMax     =   4500
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmPathImprove.frx":6852
         ScrollTrack     =   -1  'True
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
         BackColorFrozen =   13430215
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "�빴ѡ�����°�·����Ҫ�����Ľ׶Ρ�"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   7200
         TabIndex        =   27
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lblInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "ԭ��׼סԺ��Ϊ��12�죬��ƽ����׼סԺ��Ϊ��10�졣"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   26
         Top             =   600
         Width           =   6255
      End
      Begin VB.Label lblInfo 
         BackColor       =   &H80000005&
         Caption         =   "���ٴ�·�����ܲ�������400�ˣ���������������Ϊ��320�ˡ�"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   6735
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00F0F4E4&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      ScaleHeight     =   755
      ScaleMode       =   0  'User
      ScaleWidth      =   15420
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   8835
      Width           =   15420
      Begin VB.CommandButton cmdSend 
         Caption         =   "�°�·������(&S)"
         Height          =   309
         Left            =   12600
         TabIndex        =   13
         Top             =   120
         Width           =   1500
      End
      Begin VB.Line linSplit 
         BorderColor     =   &H8000000F&
         Index           =   0
         X1              =   0
         X2              =   20400
         Y1              =   0
         Y2              =   0
      End
      Begin VB.Line linSplit 
         BorderColor     =   &H8000000E&
         Index           =   1
         X1              =   0
         X2              =   20640
         Y1              =   0
         Y2              =   0
      End
   End
   Begin VB.Frame fraFilter 
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      Caption         =   "��ѯ����"
      Height          =   855
      Left            =   120
      TabIndex        =   15
      Top             =   0
      Width           =   15375
      Begin VB.ComboBox cboBranch 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   38
         Top             =   480
         Width           =   3075
      End
      Begin VB.ComboBox cboCategory 
         Height          =   300
         Left            =   960
         Style           =   2  'Dropdown List
         TabIndex        =   0
         Top             =   60
         Width           =   1815
      End
      Begin VB.CommandButton cmdPathName 
         Appearance      =   0  'Flat
         Caption         =   "��"
         Height          =   250
         Left            =   6880
         Picture         =   "frmPathImprove.frx":68DF
         TabIndex        =   35
         Top             =   75
         Width           =   300
      End
      Begin VB.TextBox txtPathName 
         Height          =   320
         Left            =   4200
         TabIndex        =   1
         Text            =   "����Ϣ��"
         Top             =   60
         Width           =   3015
      End
      Begin VB.TextBox txtRate 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   0
         Left            =   6600
         MaxLength       =   2
         TabIndex        =   6
         Text            =   "95"
         Top             =   495
         Width           =   400
      End
      Begin VB.CommandButton cmdAnalyse 
         Caption         =   "�������(F)"
         Height          =   320
         Left            =   13440
         TabIndex        =   7
         Top             =   470
         Width           =   1500
      End
      Begin VB.TextBox txtRate 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   2
         Left            =   12120
         MaxLength       =   2
         TabIndex        =   5
         Text            =   "80"
         Top             =   495
         Width           =   400
      End
      Begin VB.TextBox txtRate 
         Alignment       =   2  'Center
         Height          =   270
         Index           =   1
         Left            =   9240
         MaxLength       =   2
         TabIndex        =   4
         Text            =   "95"
         Top             =   495
         Width           =   400
      End
      Begin VB.PictureBox picTime 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   7560
         ScaleHeight     =   300
         ScaleWidth      =   6015
         TabIndex        =   16
         Top             =   60
         Width           =   6015
         Begin MSComCtl2.DTPicker dtpTimeStart 
            Height          =   300
            Left            =   1080
            TabIndex        =   2
            Top             =   0
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   106168323
            CurrentDate     =   41334
         End
         Begin MSComCtl2.DTPicker dtpTimeEnd 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "yyyy/MM/dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   3
            EndProperty
            Height          =   300
            Left            =   3510
            TabIndex        =   3
            Top             =   0
            Width           =   2115
            _ExtentX        =   3731
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   99352579
            CurrentDate     =   41365
         End
         Begin VB.Label lblBetweenTimes 
            AutoSize        =   -1  'True
            BackColor       =   &H80000005&
            Caption         =   "����ʱ�䣺��                        ��                    "
            Height          =   180
            Left            =   0
            TabIndex        =   17
            Top             =   60
            Width           =   5220
         End
      End
      Begin VB.Label lblBranch 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "��֧·��"
         Height          =   180
         Left            =   120
         TabIndex        =   39
         Top             =   540
         Width           =   720
      End
      Begin VB.Label lblRate 
         BackColor       =   &H80000005&
         Caption         =   "�׶���ǰ���Ӻ�ı�����      %"
         Height          =   180
         Index           =   2
         Left            =   4560
         TabIndex        =   23
         Top             =   540
         Width           =   2655
      End
      Begin VB.Label lblRate 
         BackColor       =   &H80000005&
         Caption         =   "·������Ŀ�ı�����      %"
         Height          =   180
         Index           =   1
         Left            =   10440
         TabIndex        =   21
         Top             =   540
         Width           =   2535
      End
      Begin VB.Label lblRate 
         BackColor       =   &H80000005&
         Caption         =   "·����δ���ɱ�����      %"
         Height          =   180
         Index           =   0
         Left            =   7560
         TabIndex        =   20
         Top             =   540
         Width           =   2535
      End
      Begin VB.Label lblName 
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         Caption         =   "·������(&N)"
         Height          =   180
         Left            =   3120
         TabIndex        =   19
         Top             =   120
         Width           =   990
      End
      Begin VB.Label lblCategory 
         BackColor       =   &H80000005&
         Caption         =   "����(&C)"
         Height          =   180
         Left            =   120
         TabIndex        =   18
         Top             =   120
         Width           =   630
      End
   End
   Begin VB.PictureBox picCenter 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   5565
      Left            =   240
      ScaleHeight     =   5535
      ScaleWidth      =   14955
      TabIndex        =   14
      Top             =   2880
      Width           =   14985
      Begin VB.Frame fraSplit 
         Height          =   15
         Left            =   -840
         MousePointer    =   7  'Size N S
         TabIndex        =   34
         Top             =   2400
         Width           =   15735
      End
      Begin VB.Frame fraAdvice 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   3480
         Left            =   120
         TabIndex        =   31
         Top             =   2400
         Width           =   14775
         Begin VB.Frame fraSplitAdvice 
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Height          =   2055
            Left            =   6840
            MousePointer    =   9  'Size W E
            TabIndex        =   36
            Top             =   600
            Width           =   135
         End
         Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
            Height          =   1935
            Index           =   0
            Left            =   120
            TabIndex        =   11
            Tag             =   "ȡ��ҽ��"
            Top             =   720
            Width           =   6375
            _cx             =   11245
            _cy             =   3413
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
            BackColorFixed  =   13430215
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16764057
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   32768
            GridColorFixed  =   32768
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   6
            Cols            =   5
            FixedRows       =   2
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   4000
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmPathImprove.frx":D131
            ScrollTrack     =   -1  'True
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   0   'False
            AutoSizeMode    =   0
            AutoSearch      =   1
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
            AllowUserFreezing=   0
            BackColorFrozen =   13430215
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VSFlex8Ctl.VSFlexGrid vsAdvice 
            Height          =   1935
            Index           =   1
            Left            =   7320
            TabIndex        =   12
            Tag             =   "����ҽ��"
            Top             =   720
            Width           =   6015
            _cx             =   10610
            _cy             =   3413
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
            BackColorFixed  =   13430215
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16764057
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   32768
            GridColorFixed  =   32768
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   6
            Cols            =   5
            FixedRows       =   2
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   4000
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmPathImprove.frx":D21F
            ScrollTrack     =   -1  'True
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   0
            MergeCompare    =   0
            AutoResize      =   -1  'True
            AutoSizeMode    =   0
            AutoSearch      =   1
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
            AllowUserFreezing=   0
            BackColorFrozen =   13430215
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "·��ҽ���� δʹ�ñ���=����ʹ�øý׶ε�δ���ɸ�·��ҽ���Ĳ�����/����ʹ�øý׶εĲ�������       "
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   33
            Top             =   120
            Width           =   8055
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "ʹ�ñ���=�ý׶�������Ӹ�ҽ����·������Ŀ�Ĳ�����/����ʹ�øý׶εĲ�������"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   6
            Left            =   1080
            TabIndex        =   32
            Top             =   360
            Width           =   6735
         End
      End
      Begin VB.Frame fraItem 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2295
         Left            =   0
         TabIndex        =   28
         Top             =   0
         Width           =   14895
         Begin VB.Frame fraSplitItem 
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            Height          =   2055
            Left            =   6960
            MousePointer    =   9  'Size W E
            TabIndex        =   37
            Top             =   720
            Width           =   120
         End
         Begin VSFlex8Ctl.VSFlexGrid vsItem 
            Height          =   1575
            Index           =   0
            Left            =   240
            TabIndex        =   9
            Top             =   600
            Width           =   6615
            _cx             =   11668
            _cy             =   2778
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
            BackColorFixed  =   13430215
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16764057
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   32768
            GridColorFixed  =   32768
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   10
            Cols            =   4
            FixedRows       =   2
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   4000
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmPathImprove.frx":D2F6
            ScrollTrack     =   -1  'True
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   110
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
            BackColorFrozen =   13430215
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VSFlex8Ctl.VSFlexGrid vsItem 
            Height          =   1575
            Index           =   1
            Left            =   7080
            TabIndex        =   10
            Top             =   600
            Width           =   6375
            _cx             =   11245
            _cy             =   2778
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
            BackColorFixed  =   13430215
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16764057
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   32768
            GridColorFixed  =   32768
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   10
            Cols            =   4
            FixedRows       =   2
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   4000
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmPathImprove.frx":D3C8
            ScrollTrack     =   -1  'True
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   110
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
            BackColorFrozen =   13430215
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "ʹ�ñ���=�ý׶�������Ӹ÷�ҽ����·������Ŀ�Ĳ�����/����ʹ�øý׶εĲ�������"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   4
            Left            =   1080
            TabIndex        =   30
            Top             =   360
            Width           =   7095
         End
         Begin VB.Label lblInfo 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            BackStyle       =   0  'Transparent
            Caption         =   "·����Ŀ�� δʹ�ñ���=����ʹ�øý׶ε�δ���ɷ�ҽ������Ŀ�Ĳ�����/����ʹ�øý׶εĲ�������"
            ForeColor       =   &H80000008&
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   29
            Top             =   120
            Width           =   8055
         End
      End
   End
End
Attribute VB_Name = "frmPathImprove"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlng·��ID As Long              '��¼��ǰ·��ID
Private mstrPrivs As String             '����·����Ȩ��
Private mstr���� As String              '��¼��ǰ·����
Private mstr���� As String              '��ǰ·������
Private mlng�汾�� As Long              '·���汾��
Private mrsStep As ADODB.Recordset      '���н׶�
Private mblnSend As Boolean             '���ɳɹ����

'�׶��к�
Private Enum COL_Step
    COL_�׶� = 0
    COL_ѡ��
    COL_����
End Enum
'
Private Enum INDEX_TAG
    Index_DEL = 0
    Index_Add = 1
End Enum
'��Ŀ���к�
Private Enum COL_Item
    COL_Item_�׶� = 0
    COL_Item_ѡ��
    COL_Item_��Ŀ����
    COL_Item_����
    COL_Item_����    '������Ŀ�õ�
End Enum
'ҽ�����к�
Private Enum COL_Advice
    COL_Advice_�׶� = 0
    COL_Advice_ҽ��ID
    COL_Advice_���ID
    COL_Advice_ѡ��
    COL_Advice_��Ч
    COL_Advice_ҽ������ 'ȡ��ҽ��������cellֵ��� ҽ�����;����ҽ��:����cell��� ִ��ID
    COL_Advice_����
    COL_Advice_������� '����ҽ��:cell��ŵ���Ŀ����
    COL_Advice_������ĿID
    COL_Advice_�걾��λ
    COL_Advice_����
    COL_Advice_��Ŀ����
End Enum
'����ֵ�±�����
Private Enum INDEX_RATE
    RATE_STEP = 0
    RATE_UNSEND
    RATE_PATHOUT
End Enum



Private Sub cboBranch_Click()
    Dim lngId As Long
    
    With cboBranch
        If .ListIndex = Val(.Tag) Then Exit Sub '.tag��ʼֵΪ-1
        If mrsStep Is Nothing Then Exit Sub  'δ���з���
        lngId = .ItemData(.ListIndex)
        Call SetVSRowHidden(vsStep, lngId)
        Call SetVSRowHidden(vsItem(Index_Add), lngId)
        Call SetVSRowHidden(vsItem(Index_DEL), lngId)
        Call SetVSRowHidden(vsAdvice(Index_Add), lngId)
        Call SetVSRowHidden(vsAdvice(Index_DEL), lngId)
        .Tag = .ListIndex
    End With
End Sub

Private Sub cboCategory_Click()
    Dim lngCmd As Long
    If Trim(cboCategory.Text) = "" Then Exit Sub
    mstr���� = Trim(cboCategory.Text)
    If cboCategory.Tag = "LOAD" Then
        lngCmd = 0
    Else
        lngCmd = 1
    End If
    Call LoadPathName(lngCmd)
    cboCategory.Tag = ""
End Sub

Private Sub cboCategory_KeyPress(KeyAscii As Integer)
'��������   �ַ����������б�

    Call Cbo.SetIndex(cboCategory.Hwnd, Cbo.MatchIndex(cboCategory.Hwnd, KeyAscii))
    Call cboCategory_Click 'CboSetIndex�������ᴥ��click�¼�
End Sub

Private Sub cboCategory_LostFocus()
    If Trim(cboCategory.Text) = "" Then
      cboCategory.SetFocus
      Call cboCategory_KeyPress(vbKeySpace)  '���Ϊ��ʱ����������������
    End If
End Sub

Private Sub cmdAnalyse_Click()
    Dim i As Integer
    Dim lngAllPati As Long
    
    '��ʼʱ��<����ʱ����
    If DateDiff("s", CDate(dtpTimeStart.Value), CDate(dtpTimeEnd.Value)) < 0 Then
         MsgBox "��ʼʱ�����ڽ���ʱ��,�����µ���ʱ�䡣", vbInformation + vbOKOnly, gstrSysName
         dtpTimeStart.SetFocus
         Exit Sub
    End If
    'ʱ��������
    If DateDiff("m", CDate(dtpTimeStart.Value), CDate(dtpTimeEnd.Value)) >= 3 Then
        If MsgBox("��ѡ������ڼ������3����,��ȷ������������ڼ������ͳ�Ʒ�����?", vbQuestion + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
            Exit Sub
        End If
    End If
    '�������
    If Val(txtRate(RATE_STEP).Text) < 70 Or Val(txtRate(RATE_UNSEND).Text) < 70 Or Val(txtRate(RATE_PATHOUT).Text) < 70 Then
        If MsgBox("������ı���ֵС��70,��ȷ��Ҫ�����������ֵ����ͳ����", vbQuestion + vbOKCancel + vbDefaultButton2, gstrSysName) = vbCancel Then
            txtRate(i).SetFocus
            Exit Sub
        End If
    End If
    '���ؽ׶�
    Call GetPathPhase
    
    '���ػ�����Ϣ
    Call SetSummaryInfo
    
    If Val(lblInfo(0).Tag) = 0 Then
        MsgBox "��ǰû���ҵ��ϸ�Ĳ���,��������������ٽ��б��������", vbOKOnly + vbDefaultButton1 + vbInformation, gstrSysName
        Exit Sub
    End If
    
    '���ؽ׶�
    
    Call LoadPhase
    '������Ŀ
    Call LoadItem
    '����ҽ��
    Call LoadAdvice
    
    'ȱʡ��λ����·��
    cboBranch.Tag = "-1"
    cboBranch.ListIndex = 0
    Call cboBranch_Click
    
End Sub

Private Sub cmdPathName_Click()
'����·������
    Call LoadPathName(2, "")
End Sub

Private Sub cmdSend_Click()
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim intVersion As Integer
    Dim blnTrans As Boolean
    Dim strҽ����ID As String
    Dim blnDo As Boolean
    Dim arrSQL As Variant
    Dim strTmp As String
    Dim i As Long
    
    On Error GoTo errH
    If mrsStep Is Nothing Then Exit Sub
    ' ִ��ǰ������δѡ���κ���Ŀ����ʾ����ֹ
    blnDo = vsAdvice(Index_Add).FindRow("1", , COL_Advice_ѡ��) = -1 And vsAdvice(Index_DEL).FindRow("1", , COL_Advice_ѡ��) = -1 And _
        vsItem(Index_DEL).FindRow("1", , COL_Item_ѡ��) = -1 And vsItem(Index_Add).FindRow("1", , COL_Item_ѡ��) = -1
    If blnDo And vsStep.FindRow("1", , COL_ѡ��) = -1 Then
        If MsgBox("��δѡ���κ�����,�Ƿ���Ҫ�˳�?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            Unload Me: Exit Sub
        Else
            Exit Sub
        End If
    End If
    
    ' ���������Ҫ�����Ľ׶Σ������û���ʾ��Ҫ���û��ֶ������׶Ρ�
    If vsStep.FindRow("1", , COL_ѡ��) <> -1 Then
        MsgBox "�������Ҫ�����׶Σ��뵽·����ƽ����ֶ�������", vbOKOnly + vbInformation + vbDefaultButton1, gstrSysName
    End If
    
    If blnDo Then Exit Sub  'δѡ���κ��޸�����
    
    '����ٴ�·���汾���״̬��δ���ʱ����ɾ��δ��˰汾��������
    arrSQL = Array()
    strSql = "Select ���ʱ��, �汾��" & _
            "   From (Select t.���ʱ��, t.�汾�� From �ٴ�·���汾 T Where t.·��id = [1] Order By t.�汾�� Desc)" & _
            "   Where Rownum < 2"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng·��ID)
    If rsTmp Is Nothing Then Exit Sub
    If rsTmp.RecordCount = 1 Then
        If IsNull(rsTmp!���ʱ��) Then
            intVersion = rsTmp!�汾��
            If MsgBox("��ǰ·������δ��˵��°�," & vbCrLf & "��ȷ��Ҫɾ��δ��˵��°汾������?", vbOKCancel + vbDefaultButton2 + vbQuestion, gstrSysName) = vbOK Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = "Zl_�ٴ�·���汾_Delete(" & mlng·��ID & "," & intVersion & ")"
            Else
                Exit Sub
            End If
        End If
    End If
    
    '����ѡ�񣬸��ƾɰ�·����ɾ�����·����Ŀ��·��ҽ����
    '���Ƶ�ǰѡ��汾���ݲ����°汾����
    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
    arrSQL(UBound(arrSQL)) = "Zl_�ٴ�·���汾_Copy(" & mlng·��ID & "," & mlng�汾�� & "," & mlng·��ID & ",0)"
   
    '��������:��Ŀȡ�� ��������Ŀ ��ȡ��ҽ�� ������ҽ��
    '�ȴ���������Ŀ�ٴ���ɾ����Ŀ,���������׶ε���Ŀɾ�����������ʱ����
    '������Ŀ
    With vsItem(Index_Add)
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, COL_Item_ѡ��) = Checked Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_�ٴ�·���汾_Improve(1," & mlng·��ID & "," & .Cell(flexcpData, i, COL_Item_�׶�) & ",Null,Null,'" & .Cell(flexcpData, i, COL_Item_����) & "'," & .Cell(flexcpData, i, COL_Item_��Ŀ����) & ")"  '�˴�COL_Item_��Ŀ���ƴ����ִ��IDֵ
            End If
        Next
    End With
     '����ҽ��
    With vsAdvice(Index_Add)
        strҽ����ID = ""
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, COL_Advice_ѡ��) = Checked Then
                strTmp = IIf(Val(.TextMatrix(i, COL_Advice_���ID)) = 0, .TextMatrix(i, COL_Advice_ҽ��ID), .TextMatrix(i, COL_Advice_���ID))
                If InStr(strҽ����ID & ",", "," & strTmp & ",") = 0 Then
                    strҽ����ID = strҽ����ID & "," & strTmp
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_�ٴ�·���汾_Improve(3," & mlng·��ID & "," & .Cell(flexcpData, i, COL_Advice_�׶�) & ",'·���Ľ���Ŀ',Null,'" & .Cell(flexcpData, i, COL_Advice_�������) & "'," & .Cell(flexcpData, i, COL_Advice_ҽ������) & "," & Val(strTmp) & ")"
                End If
            End If
        Next
    End With
    
    '��Ŀȡ��
    With vsItem(Index_DEL)
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, COL_Item_ѡ��) = Checked Then
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_�ٴ�·���汾_Improve(0," & mlng·��ID & "," & .Cell(flexcpData, i, COL_Item_�׶�) & ",'" & .TextMatrix(i, COL_Item_��Ŀ����) & "')"
            End If
        Next
    End With
    'ȡ��ҽ��
    With vsAdvice(Index_DEL)
        strҽ����ID = "" '��¼��ӹ�����ID
        For i = .FixedRows To .Rows - 1
            If .Cell(flexcpChecked, i, COL_Advice_ѡ��) = Checked Then
                strTmp = IIf(Val(.TextMatrix(i, COL_Advice_���ID)) = 0, .TextMatrix(i, COL_Advice_ҽ��ID), .TextMatrix(i, COL_Advice_���ID))
                If InStr(strҽ����ID & ",", "," & strTmp & ",") = 0 Then
                    strҽ����ID = strҽ����ID & "," & strTmp
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_�ٴ�·���汾_Improve(2," & mlng·��ID & "," & .Cell(flexcpData, i, COL_Advice_�׶�) & ",'" & .Cell(flexcpData, i, COL_Advice_��Ŀ����) & "'," & .Cell(flexcpData, i, COL_Advice_ҽ������) & ")"
                End If
            End If
        Next
    End With
    
     '�ύ����
    gcnOracle.BeginTrans: blnTrans = True
    For i = LBound(arrSQL) To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next i
    gcnOracle.CommitTrans: blnTrans = False
    mblnSend = True
    '5)  ���ɳɹ���ر��˳����塣
    Unload Me
    Exit Sub
errH:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub dtpTimeEnd_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        txtRate(RATE_STEP).SetFocus
    End If
End Sub

Private Sub dtpTimeStart_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        dtpTimeEnd.SetFocus
    End If
End Sub

Private Sub Form_Load()
    Dim DatCurr As Date
    Dim strTmp As String

    '���ط���
    Call Cbo.SetListHeight(cboCategory, 2000)
    cboCategory.Tag = "LOAD"
    Call LoadCategory
    'Ĭ�Ͽ�ʼʱ��ͽ���ʱ����Ϊ30��
    DatCurr = zlDatabase.Currentdate
    dtpTimeStart.Value = Format(DateAdd("d", -30, DatCurr), "YYYY-MM-DD 00:00:00")
    dtpTimeEnd.Value = Format(DatCurr, "YYYY-MM-DD 23:59:59")
    
    '
    '���г�ʼ��
    Call InitStep
    
    strTmp = "�׶�;ȡ����Ŀ;ȡ����Ŀ;ȡ����Ŀ|�׶�,1500,4;ѡ��,500,4;��Ŀ����,3000,4;δʹ�ñ���(%),1500,4"
    Call InitVSItem(vsItem(Index_DEL), strTmp)
    
    strTmp = "�׶�;������Ŀ;������Ŀ;������Ŀ;������Ŀ|�׶�,1500,4;ѡ��,500,4;��Ŀ����,3000,4;ʹ�ñ���(%),1500,4;����,,"
    Call InitVSItem(vsItem(Index_Add), strTmp)
    
    strTmp = "�׶�;ȡ��ҽ��;ȡ��ҽ��;ȡ��ҽ��;ȡ��ҽ��;ȡ��ҽ��;ȡ��ҽ��;ȡ��ҽ��;ȡ��ҽ��;ȡ��ҽ��;ȡ��ҽ��;ȡ��ҽ��|" & _
            "�׶�,1500,4;���ID;ҽ��ID;ѡ��,500,4;��Ч,500,4;ҽ������,2500,4;δʹ�ñ���(%),1500,4;������ĿID;�������;�걾��λ;����;��Ŀ����,,"
    Call InitVSAdvice(vsAdvice(Index_DEL), strTmp)
    
    strTmp = "�׶�;����ҽ��;����ҽ��;����ҽ��;����ҽ��;����ҽ��;����ҽ��;����ҽ��|" & _
            "�׶�,1500,4;���ID;ҽ��ID;ѡ��,500,4;��Ч,500,4;ҽ������,2500,4;ʹ�ñ���(%),1500,4;�������"
    Call InitVSAdvice(vsAdvice(Index_Add), strTmp)
    
    lblInfo(0).Caption = "���ٴ�·�����ܲ�������0 �ˣ���������������Ϊ��0 �ˡ�"
    lblInfo(1).Caption = "ԭ��׼סԺ��Ϊ��0 �죬��ƽ����׼סԺ��Ϊ��0 �졣"
End Sub

Private Sub Form_Resize()
    Dim lngWidth As Long
    Dim lngLeft As Long

    On Error Resume Next
    lngLeft = 105
    lngWidth = Me.ScaleWidth - lngLeft * 2

    fraFilter.Move lngLeft, 0, lngWidth
    cmdAnalyse.Move fraFilter.Width - cmdAnalyse.Width - 450, (fraFilter.Height - cmdAnalyse.Height) / 2
    fraContent.Move lngLeft, fraFilter.Top + fraFilter.Height, lngWidth
    lblInfo(2).Left = lngLeft + lngWidth / 2
    vsStep.Left = lngLeft + lngWidth / 2
    picCenter.Move lngLeft, fraContent.Top + fraContent.Height, lngWidth, Me.ScaleHeight - picBottom.Height - (fraContent.Top + fraContent.Height)

End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstrPrivs = ""
    
    If Not mrsStep Is Nothing Then
        Set mrsStep = Nothing
    End If
    
End Sub

Private Sub fraSplitAdvice_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        On Error Resume Next

        If fraSplitAdvice.Left + X <= picCenter.ScaleWidth / 10 * 1 Or fraSplitAdvice.Left + X >= picCenter.ScaleWidth / 10 * 9 Then Exit Sub
        vsAdvice(Index_DEL).Width = vsAdvice(Index_DEL).Width + X
        fraSplitAdvice.Left = fraSplitAdvice.Left + X
        vsAdvice(Index_Add).Left = vsAdvice(Index_Add).Left + X
        vsAdvice(Index_Add).Width = vsAdvice(Index_Add).Width - X
    End If
End Sub

Private Sub fraSplit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        On Error Resume Next

        If fraSplit.Top + Y <= picCenter.ScaleHeight / 10 * 1 Or fraSplit.Top + Y >= picCenter.ScaleHeight / 10 * 9 Then Exit Sub
        vsItem(Index_DEL).Height = vsItem(Index_DEL).Height + Y
        vsItem(Index_Add).Height = vsItem(Index_Add).Height + Y
        fraItem.Height = fraItem.Height + Y
        fraSplit.Top = fraSplit.Top + Y
        fraAdvice.Top = fraAdvice.Top + Y
        fraAdvice.Height = fraAdvice.Height - Y
        vsAdvice(Index_DEL).Height = vsAdvice(Index_DEL).Height - Y
        vsAdvice(Index_Add).Height = vsAdvice(Index_Add).Height - Y
    End If
End Sub

Private Sub fraSplitItem_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        On Error Resume Next

        If fraSplitItem.Left + X <= picCenter.ScaleWidth / 10 * 1 Or fraSplitItem.Left + X >= picCenter.ScaleWidth / 10 * 9 Then Exit Sub
        vsItem(Index_DEL).Width = vsItem(Index_DEL).Width + X
        fraSplitItem.Left = fraSplitItem.Left + X
        vsItem(Index_Add).Left = vsItem(Index_Add).Left + X
        vsItem(Index_Add).Width = vsItem(Index_Add).Width - X
    End If
End Sub

Private Sub picBottom_Resize()
    On Error Resume Next
    cmdSend.Move picBottom.ScaleWidth - cmdSend.Width - 500
End Sub

Private Sub picCenter_Resize()
    On Error Resume Next
    Dim lngHeight As Long

    With picCenter
        lngHeight = (.ScaleHeight - 30) / 2
        '�ָ���
        fraSplit.Move 0, lngHeight, .ScaleWidth, 30
        '�ָ�����
        fraItem.Move 0, 0, .ScaleWidth, lngHeight
        vsItem(Index_DEL).Move 0, 600, .ScaleWidth / 2, lngHeight - 600
        fraSplitItem.Move vsItem(Index_DEL).Width, 600, 60, .Height
        vsItem(Index_Add).Move fraSplitItem.Left + 60, 600, .ScaleWidth / 2 - 30, lngHeight - 600
        '�ָ�����
        fraAdvice.Move 0, lngHeight + fraSplit.Height + 60, .ScaleWidth, lngHeight
        vsAdvice(Index_DEL).Move 0, 600, .ScaleWidth / 2, lngHeight - 600
        fraSplitAdvice.Move vsAdvice(Index_DEL).Width, 600, 60, .Height
        vsAdvice(Index_Add).Move fraSplitAdvice.Left + 60, 600, .ScaleWidth / 2 - 30, lngHeight - 600
    End With
End Sub

Public Sub ShowMe(frmParent As Object, ByVal lng·��ID As Long, ByRef str���� As String, ByRef str���� As String, ByRef blnRefresh As Boolean)
'����:
'����:lng·��ID-��ǰĬ��ѡ�е�·��ID
'     str����-ѡ�е�·������
'     blnRefresh=True ��Ҫˢ��������
'     gstrPrivs-����·����Ȩ��
    mlng·��ID = lng·��ID
    mstr���� = str����
    mstrPrivs = gstrPrivs
    mblnSend = False

    Me.Show 1, frmParent
    blnRefresh = mblnSend
    str���� = mstr����
    str���� = mstr����
End Sub

Private Sub LoadPathName(ByVal lngCmd As Long, Optional ByVal strInput As String)
'����:����ĳ�����µ�·�����ƣ������ı仯���仯
'����:lngcmd 0-��ʼ����ʱ�����ݴ��˵�mlng·��ID��λ·������
'            1-ѡ�����ʱ��Ĭ��ѡ��÷����µ�һ��·������
'            2-·������������:���֣����� ������ƥ�䡣
'     strInput -��lngCmd=2ʱ�����ˣ������Ҫƥ����ַ�
    Dim i As Long
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim lngԭ·��ID As Long
    Dim blnOK As Boolean
    
    On Error GoTo errH

    strInput = Trim(strInput)
    lngԭ·��ID = mlng·��ID
    
    If lngCmd = 0 Then
        strTmp = "and ID=" & mlng·��ID
        blnOK = True
    ElseIf lngCmd = 1 Then
        strTmp = "and ����= '" & mstr���� & "' and RowNum <2"
        blnOK = True
    Else
        If strInput <> "" Then
            '��������������жϣ�������Ǻ�����������ƣ�����������ұ���
            If zlCommFun.IsCharChinese(strInput) Then
                '�������� ��������
                strTmp = "and ����= '" & mstr���� & "' and ���� like '" & gstrLike & strInput & "%'"
            Else
                strTmp = "and ����= '" & mstr���� & "' and ���� like '" & gstrLike & UCase(strInput) & "%'"
            End If
        Else
            strTmp = "and ����= '" & mstr���� & "'"
        End If
    End If

    strSql = "Select a.Id,a.����,a.����,���°汾 From �ٴ�·��Ŀ¼ A Where a.���°汾 >= 1 " & strTmp
    If InStr(mstrPrivs, "ȫԺ·��") = 0 Then
        'û��Ȩ��ʱ��ֻ�ܶ�ֻӦ���ڱ��Ƶ�·�����д���
        strSql = strSql & "And A.ͨ�� = 2 And Exists" & vbNewLine & _
                 "      (Select 1 From ������Ա C,�ٴ�·������ D " & vbNewLine & _
                 "       Where C.��Աid = [1] and D.����id = C.����id And D.·��id = A.ID  )"
    End If
    strSql = strSql & " order by  ����,����"

    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID)

    If lngCmd = 2 Then
        If rsTmp.RecordCount = 1 Then
            blnOK = True
        ElseIf rsTmp.RecordCount > 1 Then
            If zlDatabase.zlShowListSelect(Me, glngSys, glngModul, txtPathName, rsTmp, True, , "1", rsTmp) Then
                blnOK = True
            End If
        Else
            MsgBox "δ�ҵ�����������ƥ���·����ȱʡѡ��ԭ·��", vbInformation + vbOKOnly, Me.Caption
        End If
    End If

    If blnOK Then
        txtPathName.Text = rsTmp!����
        txtPathName.Tag = rsTmp!����
        mlng·��ID = rsTmp!ID
        mlng�汾�� = rsTmp!���°汾
        mstr���� = rsTmp!����
    Else
       
        txtPathName.Text = txtPathName.Tag 'δѡ��·��ʱ������ԭ��·������
    End If
    
    txtPathName.SelStart = Len(Trim(txtPathName.Text))
    If lngԭ·��ID <> mlng·��ID Or lngCmd = 0 Then
        Call LoadBranch
        If Not mrsStep Is Nothing Then
            Call ClearData
        End If
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadCategory()
'����:����·������
    Dim i As Long
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset

    strSql = "Select distinct a.���� From �ٴ�·��Ŀ¼ A Where a.���°汾 >= 1 "
    If InStr(mstrPrivs, "ȫԺ·��") = 0 Then
        'û��Ȩ��ʱ��ֻ�ܶ�ֻӦ���ڱ��Ƶ�·�����д���
        strSql = strSql & "And A.ͨ�� = 2 And Exists" & vbNewLine & _
                "      (Select 1 From ������Ա C,�ٴ�·������ D " & vbNewLine & _
                 "       Where C.��Աid = [1] and D.����id = C.����id And D.·��id = A.ID )"
    End If
    strSql = strSql & " order by  ����"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UserInfo.ID)

    cboCategory.Clear
    For i = 1 To rsTmp.RecordCount
        cboCategory.AddItem rsTmp!����
        rsTmp.MoveNext
    Next
    'ȱʡ����
    Call Cbo.Locate(cboCategory, mstr����, False)  '�ᴥ��cboCategory_Click�¼�
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtPathName_GotFocus()
    Call zlControl.TxtSelAll(txtPathName)
End Sub

Private Sub txtPathName_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        Call LoadPathName(2, txtPathName.Text)
    End If
End Sub

Private Sub txtPathName_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then KeyAscii = 0
End Sub

Private Sub InitStep()
'����:��ʼ���ٴ�·���׶α�
    Dim strcol As String, arrHead As Variant
    Dim i As Long
    
    strcol = "�׶�,2000,4;ѡ��,500,4;����,2500,4"
    arrHead = Split(strcol, ";")
    With vsStep
        .Redraw = flexRDNone
        .Clear
        .FixedRows = 1: .FixedCols = 1
        .Cols = UBound(arrHead) + 1
        .Rows = .FixedRows + 1    'ȱʡ��ʾһ�пհ�
        .Editable = flexEDNone

        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, i) = Split(arrHead(i), ",")(0)
            If Split(arrHead(i), ",")(0) = "ѡ��" Then .ColDataType(i) = flexDTBoolean
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(i) = False
                .ColWidth(i) = Val(Split(arrHead(i), ",")(1))
                .ColAlignment(i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(i) = True
                .ColWidth(i) = 0
            End If
        Next
        .Row = 0  '��ѡ���κ���
        .Redraw = True
    End With
End Sub

Private Sub InitVSItem(ByRef vsItem As VSFlexGrid, ByVal strHeads As String)
'����:��ʼ��vsItem�ٴ�·����Ŀ��
    Dim arrHead As Variant
    Dim arrHeads As Variant
    Dim lngRow As Long
    Dim i As Long
    Dim k As Long
    
    arrHeads = Split(strHeads, "|")
    If UBound(arrHeads) < 0 Then Exit Sub
    With vsItem
        .Redraw = flexRDNone
        .Clear
        .FixedRows = 2: .FixedCols = 1
        .Cols = UBound(Split(arrHeads(0), ";")) + 1
        .Rows = .FixedRows + 1    'ȱʡ��ʾһ�пհ�
        .Editable = flexEDNone  '������༭
        .AutoSizeMode = flexAutoSizeRowHeight
        .WordWrap = True
        lngRow = 0
        For k = LBound(arrHeads) To UBound(arrHeads)
            arrHead = Split(arrHeads(k), ";")
            For i = 0 To UBound(arrHead)
                .TextMatrix(lngRow, i) = Split(arrHead(i), ",")(0)
                If Split(arrHead(i), ",")(0) = "ѡ��" Then .ColDataType(i) = flexDTBoolean
                If UBound(Split(arrHead(i), ",")) > 0 Then
                    .ColHidden(i) = False
                    .ColWidth(i) = Val(Split(arrHead(i), ",")(1))
                    .ColAlignment(i) = Val(Split(arrHead(i), ",")(2))
                Else
                    .ColHidden(i) = True
                    .ColWidth(i) = 0
                End If
            Next
            lngRow = lngRow + 1
        Next
        .MergeRow(0) = True
        .MergeCol(0) = True

        .Row = 0
        .Redraw = True
    End With
End Sub

Private Sub InitVSAdvice(ByRef vsAdvice As VSFlexGrid, ByVal strHeads As String)
'����:��ʼ���ٴ�·��ҽ����
    Dim arrHead As Variant
    Dim arrHeads As Variant
    Dim lngRow As Long
    Dim i As Long
    Dim k As Long
    
    arrHeads = Split(strHeads, "|")
    If UBound(arrHeads) < 0 Then Exit Sub
    With vsAdvice
        .Redraw = flexRDNone
        .Clear
        .FixedRows = 2: .FixedCols = 2
        .Cols = UBound(Split(arrHeads(0), ";")) + 1
        .Rows = .FixedRows + 1    'ȱʡ��ʾһ�пհ�
        .MergeCells = flexMergeFree
        .Editable = flexEDKbdMouse
        .AutoSizeMode = flexAutoSizeRowHeight
        .WordWrap = True
        lngRow = 0
        For k = LBound(arrHeads) To UBound(arrHeads)
            arrHead = Split(arrHeads(k), ";")
            For i = 0 To UBound(arrHead)
                .TextMatrix(lngRow, i) = Split(arrHead(i), ",")(0)
                If Split(arrHead(i), ",")(0) = "ѡ��" Then .ColDataType(i) = flexDTBoolean
                If UBound(Split(arrHead(i), ",")) > 0 Then
                    .ColHidden(i) = False
                    .ColWidth(i) = Val(Split(arrHead(i), ",")(1))
                    .ColAlignment(i) = Val(Split(arrHead(i), ",")(2))
                Else
                    .ColHidden(i) = True
                    .ColWidth(i) = 0
                End If
            Next
            lngRow = lngRow + 1
        Next
 
        '�̶����кϲ�����
        .MergeRow(0) = True
        .MergeCol(0) = True

        .Editable = flexEDNone  '������༭
        .Row = 0
        .Redraw = True
    End With
End Sub

Private Sub txtRate_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txtRate(Index))
End Sub

Private Sub txtRate_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        If Index < txtRate.UBound Then
            txtRate(Index + 1).SetFocus
        ElseIf Index = txtRate.UBound Then
            cmdAnalyse.SetFocus
        End If
    End If
End Sub

Private Sub txtRate_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyBack Then Exit Sub

    If Not IsNumeric(Chr(KeyAscii)) Then
        KeyAscii = 0
    ElseIf IsNumeric(Chr(KeyAscii)) Then
        '��һλ����Ϊ0
        If txtRate(Index).SelStart = 0 And Chr(KeyAscii) = "0" Then
            KeyAscii = 0
        ElseIf txtRate(Index).SelStart = 2 And Chr(KeyAscii) <> "0" Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub SetSummaryInfo()
'����:���û�����Ϣ,�����ܲ��������ϸ����������������Ĳ��ˣ�����׼סԺ�գ�ƽ����׼סԺ��
    Dim rsTmp As ADODB.Recordset
    Dim strSql As String

    strSql = "  Select Count(1) As �ܲ�����, Sum(Decode(a.״̬, 2, 1, 0)) As �ϸ�����, b.��׼סԺ��," & _
             "  To_Char(Sum(Decode(a.״̬, 2, a.��ǰ����)) / Sum(Decode(a.״̬, 2, 1, 0)), '99999.0') As ƽ����׼סԺ��" & _
             "  From �����ٴ�·�� A, �ٴ�·���汾 B " & _
             "  Where a.·��id = b.·��id And a.�汾�� = b.�汾�� And b.·��id = [1] and b.�汾��= [2] " & _
             "  and a.����ʱ�� between [3] and [4] " & _
             "  Group By b.��׼סԺ��"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng·��ID, mlng�汾��, _
                                         CDate(Format(dtpTimeStart.Value, "yyyy-MM-dd 00:00:00")), CDate(Format(dtpTimeEnd.Value, "yyyy-MM-dd 23:59:59")))
    If rsTmp Is Nothing Then Exit Sub
    
    If rsTmp.RecordCount = 1 Then
        With rsTmp
            lblInfo(0).Tag = Val(!�ϸ����� & "")
            lblInfo(0).Caption = "���ٴ�·�����ܲ�������" & !�ܲ����� & " �ˣ���������������Ϊ��" & Val(!�ϸ����� & "") & " �ˡ�"
            '��׼סԺ�գ�<=N�죻M-N��
            lblInfo(1).Caption = "ԭ��׼סԺ��Ϊ��" & !��׼סԺ�� & " �죬��ƽ����׼סԺ��Ϊ��" & Val(!ƽ����׼סԺ�� & "") & " �졣"
        End With
    End If

    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadPhase()
'����:���ؽ׶���ǰ���Ӻ��������ָ��ֵ�Ľ׶�
    Dim rsStep As ADODB.Recordset
    Dim strSql As String
    Dim strIDs As String
    Dim i As Long
    
    For i = 1 To mrsStep.RecordCount
        strIDs = strIDs & "," & mrsStep!ID
        mrsStep.MoveNext
    Next
    strIDs = Mid(strIDs, 2)
    
    On Error GoTo errH:
    
    strSql = "Select a.�׶�id, To_Char((Sum(Decode(a.ʱ�����, 1, 1)) / Count(Distinct a.·����¼id)) * 100, '990.00') As ��ǰ��," & vbNewLine & _
            "       To_Char((Sum(Decode(a.ʱ�����, -1, 1)) / Count(Distinct a.·����¼id)) * 100, '999.00') As �Ӻ���" & vbNewLine & _
            "From (Select Distinct b.�׶�id, b.·����¼id, Decode(b.ʱ�����,2,1,b.ʱ�����) as ʱ�����" & vbNewLine & _
            "       From �����ٴ�·�� A, ����·������ B, Table(f_Str2list([5])) C" & vbNewLine & _
            "       Where a.Id = b.·����¼id And a.·��id = [1] And a.�汾�� = [2] And a.״̬ = 2  And b.�׶�id = c.Column_Value And" & vbNewLine & _
            "             a.����ʱ�� Between [3] And [4]) A" & vbNewLine & _
            "Group By a.�׶�id" & vbNewLine & _
            "Having(Sum(Decode(a.ʱ�����, 1, 1)) / Count(Distinct a.·����¼id)) * 100 >= [6] Or (Sum(Decode(a.ʱ�����, -1, 1)) / Count(Distinct a.·����¼id)) * 100 > [6]"
            
    Set rsStep = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng·��ID, mlng�汾��, CDate(Format(dtpTimeStart.Value, "yyyy-MM-dd 00:00:00")), CDate(Format(dtpTimeEnd.Value, "yyyy-MM-dd 23:59:59")), strIDs, Val(txtRate(RATE_STEP).Text))
    With vsStep
        .Rows = rsStep.RecordCount + .FixedRows
        For i = 1 To rsStep.RecordCount
            mrsStep.Filter = "ID =" & rsStep!�׶�id
            .TextMatrix(i, COL_�׶�) = mrsStep!���� & IIf(Nvl(mrsStep!��ID) = "", "", ",��֧:" & Nvl(mrsStep!˵��, mrsStep!���))
            .RowData(i) = IIf(IsNull(mrsStep!��֧ID), mlng·��ID, Nvl(mrsStep!��֧ID))
            .TextMatrix(i, COL_����) = IIf(IsNull(rsStep!��ǰ��), "������ǰ)", rsStep!��ǰ�� & "%��ǰ") & "/" & IIf(IsNull(rsStep!�Ӻ���), "(���Ӻ�)", rsStep!�Ӻ��� & "%�Ӻ�")
            rsStep.MoveNext
        Next
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog

End Sub

Private Sub LoadItem()
'����:�����ٴ�·��ȡ�������ӵ���Ŀ
    Dim strSql As String
    Dim rsItemRate As ADODB.Recordset
    Dim lngTmp As Long
    Dim i As Long, k As Long
    
    On Error GoTo errH
    
    '1�� ȡ����Ŀ����
    
    'δʹ�ñ���=����ʹ�øý׶ε�δ���ɷ�ҽ������Ŀ�Ĳ�����/����ʹ�øý׶εĲ�������
    'δʹ�ñ���=(1 - Nvl(d.������, 1) / Nvl(e.������, 1)) * 100
    'Ϊδ���ɵı���:Val(txtRate(0).Text)

    strSql = "Select d.�׶�id, d.Id, d.��Ŀ����, To_Char((1 - d.������ / Nvl(e.������, 1)) * 100, '990.00') As δʹ�ñ���" & vbNewLine & _
            "From (Select b.�׶�id, b.Id, b.��Ŀ����, Nvl(������, 0) As ������" & vbNewLine & _
            "       From (Select a.�׶�id, a.��Ŀ����, Count(Distinct c.����id) As ������" & vbNewLine & _
            "              From �ٴ�·����Ŀ A, ����·��ִ�� B, �����ٴ�·�� C" & vbNewLine & _
            "              Where a.Id = b.��Ŀid And b.·����¼id = c.Id And  a.·��id = [1] And a.�汾�� = [2] And c.״̬ = 2 And" & vbNewLine & _
            "                    c.����ʱ�� Between [3] And [4] And Not Exists" & vbNewLine & _
            "               (Select 1 From �ٴ�·��ҽ�� T Where t.·����Ŀid = a.Id)" & vbNewLine & _
            "              Group By a.�׶�id, a.��Ŀ����) A, �ٴ�·����Ŀ B" & vbNewLine & _
            "       Where b.·��id = [1] And b.�汾�� = [2] And b.�׶�id = a.�׶�id(+) And b.��Ŀ���� = a.��Ŀ����(+) And Not Exists" & vbNewLine & _
            "        (Select 1 From �ٴ�·��ҽ�� T Where t.·����Ŀid = b.Id)) D," & vbNewLine & _
            "     (Select b.�׶�id, Count(Distinct a.����id) As ������" & vbNewLine & _
            "       From �����ٴ�·�� A, ����·��ִ�� B" & vbNewLine & _
            "       Where a.Id = b.·����¼id And a.·��id = [1] And a.�汾�� = [2] And a.״̬ = 2 And" & vbNewLine & _
            "             a.����ʱ�� Between [3] And [4] " & vbNewLine & _
            "       Group By b.�׶�id) E" & vbNewLine & _
            "Where d.�׶�id = e.�׶�id(+) And (1 - d.������ / Nvl(e.������, 1)) * 100 >= [5]"


    Set rsItemRate = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng·��ID, mlng�汾��, CDate(Format(dtpTimeStart.Value, "yyyy-MM-dd 00:00:00")), CDate(Format(dtpTimeEnd.Value, "yyyy-MM-dd 23:59:59")), Val(txtRate(RATE_UNSEND).Text))
     '�����ݲ��뵽ȡ����Ŀ�б���
    With vsItem(Index_DEL)
        .Redraw = flexRDNone
        .Rows = .FixedRows '�������
        .Rows = .FixedRows + 1  '������ʱĬ�Ͽ�һ��
        .Rows = rsItemRate.RecordCount + .FixedRows
        lngTmp = .FixedRows  '��¼������
        mrsStep.Filter = ""  '�ָ����м�¼
        For i = 1 To mrsStep.RecordCount  '���׶�˳������������
            '��ǰ�׶�ȡ����Ŀ���
            rsItemRate.Filter = "�׶�id=" & mrsStep!ID
            For k = 1 To rsItemRate.RecordCount
                '���ز���
                .RowData(lngTmp) = IIf(IsNull(mrsStep!��֧ID), mlng·��ID, Nvl(mrsStep!��֧ID))
                .Cell(flexcpData, lngTmp, COL_Item_��Ŀ����) = CStr(rsItemRate!ID)
                .Cell(flexcpData, lngTmp, COL_Item_�׶�) = CStr(rsItemRate!�׶�id)
                
                '��ʾ����
                .TextMatrix(lngTmp, COL_Item_�׶�) = mrsStep!���� & IIf(IsNull(mrsStep!��ID), "", ",��֧:" & Nvl(mrsStep!˵��, mrsStep!���))
                .TextMatrix(lngTmp, COL_Item_��Ŀ����) = rsItemRate!��Ŀ���� & ""
                .TextMatrix(lngTmp, COL_Item_����) = rsItemRate!δʹ�ñ��� & ""
                lngTmp = lngTmp + 1
                
                rsItemRate.MoveNext
            Next
            mrsStep.MoveNext
        Next

        If .Rows > .FixedRows Then
            .Cell(flexcpAlignment, .FixedRows, COL_Item_��Ŀ����, .Rows - 1, COL_Item_��Ŀ����) = flexAlignLeftCenter
        End If
        .Redraw = flexRDDirect
        If .FixedRows > 1 Then .AutoSize .FixedCols, .Cols - 1, , 45    '��ҪDraw֮�����Ч
    End With
    
    '2��������Ŀ����
    'ʹ�ñ���=�ý׶�������Ӹ÷�ҽ����·������Ŀ�Ĳ�����/����ʹ�øý׶εĲ�������
    'ʹ�ñ���=(Nvl(d.������, 1) / Nvl(e.������, 1)) * 100
    strSql = " Select d.�׶�id, d.����,d.��Ŀ����, d.ִ��id, To_Char((Nvl(d.������, 1) / Nvl(e.������, 1)) * 100, '990.00') As ʹ�ñ���" & _
            "   From (Select b.�׶�id,b.����, b.��Ŀ����, Max(b.Id) As ִ��id, Count(Distinct c.����id) As ������" & _
            "       From ����·��ִ�� B, �����ٴ�·�� C" & _
            "       Where b.·����¼id = c.Id And c.·��id = [1] And c.�汾�� = [2] And c.״̬ = 2 And" & _
            "             c.����ʱ�� Between [3] And [4] And b.��Ŀid Is Null And b.��Ŀ���� <> 'δ�����κ���Ŀ' And" & _
            "             b.��Ŀ���� <> '·������Ŀ' And Not Exists (Select 1 From ����·��ҽ�� T Where t.·��ִ��id = b.Id)" & _
            "       Group By b.�׶�id, b.����,b.��Ŀ����) D," & _
            "     (Select b.�׶�id, Count(Distinct a.����id) As ������" & _
            "       From �����ٴ�·�� A, ����·��ִ�� B" & _
            "       Where a.Id = b.·����¼id And a.·��id = [1] And a.�汾�� = [2] And a.״̬ = 2 And" & _
            "             a.����ʱ�� Between [3] And [4] " & _
            "       Group By b.�׶�id) E" & _
            "   Where d.�׶�id = e.�׶�id And (Nvl(d.������, 1) / Nvl(e.������, 1)) * 100 >= [5]"

            
    Set rsItemRate = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng·��ID, mlng�汾��, CDate(Format(dtpTimeStart.Value, "yyyy-MM-dd 00:00:00")), CDate(Format(dtpTimeEnd.Value, "yyyy-MM-dd 23:59:59")), Val(txtRate(RATE_PATHOUT).Text))

    '�����ݲ��뵽���ӱ�����
    With vsItem(Index_Add)
        .Redraw = flexRDNone
        .Rows = .FixedRows '�������
        .Rows = .FixedRows + 1  '������ʱĬ�Ͽ�һ��
        .Rows = rsItemRate.RecordCount + .FixedRows
        
        lngTmp = .FixedRows  '��¼������
        mrsStep.Filter = ""  '�ָ����м�¼
        For i = 1 To mrsStep.RecordCount  '���׶�˳������������
         '��ǰ�׶�������Ŀ���
            rsItemRate.Filter = "�׶�id=" & mrsStep!ID
            For k = 1 To rsItemRate.RecordCount
                '���ز���
                .RowData(lngTmp) = IIf(IsNull(mrsStep!��֧ID), mlng·��ID, Nvl(mrsStep!��֧ID))
                .Cell(flexcpData, lngTmp, COL_Item_��Ŀ����) = CStr(rsItemRate!ִ��ID) '������Ŀ�е���ĿID���·��ִ��ID,����SelectVsItem��һ������
                .Cell(flexcpData, lngTmp, COL_Item_�׶�) = rsItemRate!�׶�id & ""
                .Cell(flexcpData, lngTmp, COL_Item_����) = rsItemRate!���� & ""
                '��ʾ����
                .TextMatrix(lngTmp, COL_Item_�׶�) = mrsStep!���� & IIf(IsNull(mrsStep!��ID), "", ",��֧:" & Nvl(mrsStep!˵��, mrsStep!���))
                .TextMatrix(lngTmp, COL_Item_��Ŀ����) = rsItemRate!��Ŀ���� & ""
                .TextMatrix(lngTmp, COL_Item_����) = rsItemRate!ʹ�ñ��� & ""
                lngTmp = lngTmp + 1
                rsItemRate.MoveNext
            Next
            mrsStep.MoveNext
        Next
     
        If .Rows > .FixedRows Then
            .Cell(flexcpAlignment, .FixedRows, COL_Item_��Ŀ����, .Rows - 1, COL_Item_��Ŀ����) = flexAlignLeftCenter
        End If
        .Redraw = flexRDDirect
        If .FixedRows > 1 Then .AutoSize .FixedCols, .Cols - 1, , 45    '��ҪDraw֮�����Ч
    End With
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

End Sub

Private Function MakeStepRS() As ADODB.Recordset
'����:�Զ����¼��,������װ�׶���Ϣ
'����:
    Set MakeStepRS = New ADODB.Recordset
    
    MakeStepRS.Fields.Append "ID", adBigInt
    MakeStepRS.Fields.Append "��ID", adBigInt, , adFldIsNullable
    MakeStepRS.Fields.Append "��֧ID", adBigInt, , adFldIsNullable
    MakeStepRS.Fields.Append "����", adVarChar, 100, adFldIsNullable
    MakeStepRS.Fields.Append "���", adBigInt, , adFldIsNullable
    MakeStepRS.Fields.Append "˵��", adVarChar, 100, adFldIsNullable
    
    MakeStepRS.CursorLocation = adUseClient
    MakeStepRS.LockType = adLockOptimistic
    MakeStepRS.CursorType = adOpenStatic
    MakeStepRS.Open
End Function

Private Sub LoadAdvice()
'����:�����ٴ�·��ȡ�������ӵ�ҽ��
    Dim strSql As String
    Dim rsAdviceRate As ADODB.Recordset
    Dim lngTmp As Long
    
    Dim i As Long, k As Long, j As Long
    
    On Error GoTo errH
    
    'ȡ��ҽ������
    
    'δʹ�ñ���=����ʹ�øý׶ε�δ���ɸ�·��ҽ���Ĳ�����/����ʹ�øý׶εĲ�������
    'δʹ�ñ���=(1 - d.������ / Nvl(e.������, 1)) * 100
            
    strSql = "Select d.�׶�id, d.��Ŀ����, d.ҽ��id, d.���id, d.ҽ������, d.��Ч, d.���, d.������Ŀid, d.�걾��λ, d.���, d.����," & vbNewLine & _
            "       To_Char((1 - d.������ / Nvl(e.������, 1)) * 100, '990.00') As δʹ�ñ���" & vbNewLine & _
            "From (Select a.�׶�id, a.��Ŀ����, e.Id As ҽ��id, e.���id, e.ҽ������, e.��Ч, e.���, e.������Ŀid, e.�걾��λ, f.���," & vbNewLine & _
            "              Nvl(g.���� || Decode(g.���, Null, Null, ' ' || g.���), f.����) As ����, Nvl(������, 0) As ������" & vbNewLine & _
            "       From (Select a.�׶�id, a.��Ŀ����, Count(Distinct c.����id) As ������" & vbNewLine & _
            "              From �ٴ�·����Ŀ A, ����·��ִ�� B, �����ٴ�·�� C" & vbNewLine & _
            "              Where a.Id = b.��Ŀid And b.·����¼id = c.Id And a.·��id = [1] And a.�汾�� = [2] And c.״̬ = 2 And" & vbNewLine & _
            "                    c.����ʱ�� Between [3] And [4] And Exists" & vbNewLine & _
            "               (Select 1 From �ٴ�·��ҽ�� T Where t.·����Ŀid = a.Id)" & vbNewLine & _
            "              Group By a.�׶�id, a.��Ŀ����) H, �ٴ�·����Ŀ A, �ٴ�·��ҽ�� D, ·��ҽ������ E, ������ĿĿ¼ F, �շ���ĿĿ¼ G" & vbNewLine & _
            "       Where a.�׶�id = h.�׶�id(+) And a.��Ŀ���� = h.��Ŀ����(+) And a.·��id = [1] And a.�汾�� = [2] And Exists" & vbNewLine & _
            "        (Select 1 From �ٴ�·��ҽ�� T Where t.·����Ŀid = a.Id) And a.Id = d.·����Ŀid And d.ҽ������id = e.Id And" & vbNewLine & _
            "             e.������Ŀid = f.Id(+) And Nvl(e.�շ�ϸĿid, -1) = g.Id(+) and Not (e.�����ĿID is not null and f.���='C')) D," & vbNewLine & _
            "     (Select b.�׶�id, Count(Distinct a.����id) As ������" & vbNewLine & _
            "       From �����ٴ�·�� A, ����·��ִ�� B" & vbNewLine & _
            "       Where a.Id = b.·����¼id And a.·��id = [1] And a.�汾�� = [2] And a.״̬ = 2 And" & vbNewLine & _
            "             a.����ʱ�� Between [3] And [4]" & vbNewLine & _
            "       Group By b.�׶�id) E" & vbNewLine & _
            "Where d.�׶�id = e.�׶�id(+) And (1 - d.������ / Nvl(e.������, 1)) * 100 >= [5]" & vbNewLine & _
            "Order By ҽ��id"


    Set rsAdviceRate = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng·��ID, mlng�汾��, CDate(Format(dtpTimeStart.Value, "yyyy-MM-dd 00:00:00")), CDate(Format(dtpTimeEnd.Value, "yyyy-MM-dd 23:59:59")), Val(txtRate(RATE_UNSEND).Text))

    '�����ݲ��뵽vsAdviceDel������
    With vsAdvice(Index_DEL)
        .Redraw = flexRDNone
        .Rows = .FixedRows  '�����һ�ε�����
        .Rows = .FixedRows + 1 'û���ݵ�ʱ��Ĭ����ʾһ�пհ�
        .Rows = .FixedRows + rsAdviceRate.RecordCount
        lngTmp = .FixedRows  '��¼������
        mrsStep.Filter = ""  '�ָ����м�¼
        For i = 1 To mrsStep.RecordCount  '���׶�˳������������
            rsAdviceRate.Filter = "�׶�id=" & mrsStep!ID
            For k = 1 To rsAdviceRate.RecordCount
                .RowData(lngTmp) = IIf(IsNull(mrsStep!��֧ID), mlng·��ID, Nvl(mrsStep!��֧ID))
                .Cell(flexcpData, lngTmp, COL_Advice_�׶�) = rsAdviceRate!�׶�id & ""
                .Cell(flexcpData, lngTmp, COL_Advice_��Ŀ����) = rsAdviceRate!��Ŀ���� & ""
                .Cell(flexcpData, lngTmp, COL_Advice_ҽ������) = rsAdviceRate!��� & "" 'ҽ�������������ݴ洢ҽ�����
                
                .TextMatrix(lngTmp, COL_Advice_�׶�) = mrsStep!���� & IIf(IsNull(mrsStep!��ID), "", ",��֧:" & Nvl(mrsStep!˵��, mrsStep!���))
                .TextMatrix(lngTmp, COL_Advice_ҽ��ID) = rsAdviceRate!ҽ��id
                .TextMatrix(lngTmp, COL_Advice_���ID) = IIf(IsNull(rsAdviceRate!���id), 0, rsAdviceRate!���id)
                .TextMatrix(lngTmp, COL_Advice_��Ч) = IIf(rsAdviceRate!��Ч = 1, "����", "����")
                .TextMatrix(lngTmp, COL_Advice_ҽ������) = IIf(rsAdviceRate!ҽ������ & "" = "", rsAdviceRate!����, rsAdviceRate!ҽ������)
                .TextMatrix(lngTmp, COL_Advice_������ĿID) = rsAdviceRate!������ĿID & ""
                .TextMatrix(lngTmp, COL_Advice_�걾��λ) = rsAdviceRate!�걾��λ & ""
                .TextMatrix(lngTmp, COL_Advice_����) = rsAdviceRate!δʹ�ñ��� & ""
                .TextMatrix(lngTmp, COL_Advice_�������) = rsAdviceRate!��� & ""
                lngTmp = lngTmp + 1
                rsAdviceRate.MoveNext
            Next
            mrsStep.MoveNext
        Next
        
        If .Rows > .FixedRows Then
            .Cell(flexcpAlignment, .FixedRows, COL_Advice_ҽ������, .Rows - 1, COL_Advice_ҽ������) = flexAlignLeftCenter
        End If
        .Redraw = flexRDDirect
        .AutoSize .FixedCols, .Cols - 1, , 45
    End With
    '����ҽ������
    'ʹ�ñ���=�ý׶�������Ӹ�ҽ����·������Ŀ�Ĳ�����/����ʹ�øý׶εĲ�������
    'ʹ�ñ���=(ʹ�ò�����/�ܲ�����)*100
    strSql = "Select a.�׶�id, a.����, a.������Ŀid,a.ִ��ID,a.ҽ��ID,c.���id,c.ҽ����Ч as ��Ч,c.�������,c.ҽ������,c.�걾��λ,To_Char(a.������ / b.������ * 100, '900.00') As ʹ�ñ���" & vbNewLine & _
            "From (Select b.�׶�id, b.����, d.������Ŀid,Max(b.ID) as ִ��ID, Count(Distinct a.����id) As ������,Max(d.ID) as ҽ��ID" & vbNewLine & _
            "       From �����ٴ�·�� A, ����·��ִ�� B, ����·��ҽ�� C, ����ҽ����¼ D" & vbNewLine & _
            "       Where a.Id = b.·����¼id And b.Id = c.·��ִ��id And c.����ҽ��id = d.Id And a.·��id = [1] And a.�汾�� = [2] And a.״̬ = 2 And" & vbNewLine & _
            "             a.����ʱ�� Between [3] And [4] And b.��Ŀid Is Null And d.������� <> 'E' And" & vbNewLine & _
            "             Not (d.���id Is Not Null And d.������� In ('F', 'G', 'D'))" & vbNewLine & _
            "       Group By b.�׶�id, b.����, d.������Ŀid) A," & vbNewLine & _
            "     (Select b.�׶�id, Count(Distinct a.����id) As ������" & vbNewLine & _
            "       From �����ٴ�·�� A, ����·��ִ�� B" & vbNewLine & _
            "       Where a.Id = b.·����¼id And a.·��id = [1] And a.�汾�� = [2] And a.״̬ = 2 And" & vbNewLine & _
            "             a.����ʱ�� Between [3] And [4]" & vbNewLine & _
            "       Group By b.�׶�id) B,����ҽ����¼ c" & vbNewLine & _
            "Where a.�׶�id = b.�׶�id and a.ҽ��ID=c.id And c.�����Ŀid Is  Null and a.������ / b.������ * 100>=[5] " & _
            " order by ҽ��ID "
 
    Set rsAdviceRate = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng·��ID, mlng�汾��, CDate(Format(dtpTimeStart.Value, "yyyy-MM-dd 00:00:00")), CDate(Format(dtpTimeEnd.Value, "yyyy-MM-dd 23:59:59")), Val(txtRate(RATE_PATHOUT).Text))

     '�����ݲ�������Ҫ������ݲ��뵽vsAdviceAdd������
    With vsAdvice(Index_Add)
        .Redraw = flexRDNone
        .Rows = .FixedRows  '�����һ�ε�����
        .Rows = .FixedRows + 1 'û���ݵ�ʱ��Ĭ����ʾһ�пհ�
        .Rows = .FixedRows + rsAdviceRate.RecordCount
        lngTmp = .FixedRows  '��¼������
        mrsStep.Filter = ""  '�ָ����м�¼
        For i = 1 To mrsStep.RecordCount  '���׶�˳������������
            rsAdviceRate.Filter = "�׶�id=" & mrsStep!ID
            For k = 1 To rsAdviceRate.RecordCount
                '����������
                .RowData(lngTmp) = IIf(IsNull(mrsStep!��֧ID), mlng·��ID, Nvl(mrsStep!��֧ID))
                .Cell(flexcpData, lngTmp, COL_Advice_�׶�) = rsAdviceRate!�׶�id & ""
                .Cell(flexcpData, lngTmp, COL_Advice_�������) = rsAdviceRate!���� & ""
                .Cell(flexcpData, lngTmp, COL_Advice_ҽ������) = rsAdviceRate!ִ��ID & ""
                
                .TextMatrix(lngTmp, COL_Advice_�׶�) = mrsStep!���� & IIf(IsNull(mrsStep!��ID), "", ",��֧:" & Nvl(mrsStep!˵��, mrsStep!���))
                .TextMatrix(lngTmp, COL_Advice_ҽ��ID) = rsAdviceRate!ҽ��id
                .TextMatrix(lngTmp, COL_Advice_���ID) = Nvl(rsAdviceRate!���id, 0)
                .TextMatrix(lngTmp, COL_Advice_��Ч) = IIf(rsAdviceRate!��Ч = 1, "����", "����")
                If rsAdviceRate!������� & "" = "C" Then
                    .TextMatrix(lngTmp, COL_Advice_ҽ������) = rsAdviceRate!ҽ������ & "��" & rsAdviceRate!�걾��λ & ")"
                Else
                    .TextMatrix(lngTmp, COL_Advice_ҽ������) = rsAdviceRate!ҽ������ & ""
                End If
                .TextMatrix(lngTmp, COL_Advice_����) = rsAdviceRate!ʹ�ñ���
                lngTmp = lngTmp + 1
                rsAdviceRate.MoveNext
            Next
            mrsStep.MoveNext
        Next
        
        If .Rows > .FixedRows Then
            .Cell(flexcpAlignment, .FixedRows, COL_Advice_ҽ������, .Rows - 1, COL_Advice_ҽ������) = flexAlignLeftCenter
        End If
        .Redraw = flexRDDirect
        .AutoSize .FixedCols, .Cols - 1, , 45
        'һ����ҩ������ߵĴ���
    End With
   
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub GetPathPhase()
'����:��ȡ��ǰ·���׶���Ϣ
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
        
    On Error GoTo errH
    strSql = "Select a.Id , a.��id,a.��֧id, a.����,Decode(b.���, Null, 0, a.���) As ���, a.˵��" & _
            "   From �ٴ�·���׶� A, �ٴ�·���׶� B" & _
            "   Where a.��id = b.Id(+)   And a.·��id = [1] And a.�汾�� =[2]" & _
            "   Order By Nvl(a.��֧ID,0), Nvl(b.���, a.���), Decode(b.���, Null, 0, a.���)"


    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng·��ID, mlng�汾��)
    Set mrsStep = MakeStepRS
    For i = 1 To rsTmp.RecordCount
        mrsStep.AddNew
        mrsStep!ID = rsTmp!ID
        mrsStep!��ID = rsTmp!��ID
        mrsStep!��֧ID = rsTmp!��֧ID
        mrsStep!���� = rsTmp!����
        mrsStep!��� = rsTmp!���
        mrsStep!˵�� = rsTmp!˵��
        rsTmp.MoveNext
    Next
    If mrsStep.RecordCount > 0 Then mrsStep.Update: mrsStep.MoveFirst
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function RowInһ����ҩ(ByVal lngRow As Long, lngBegin As Long, lngEnd As Long, ByVal blnIsHide As Boolean, ByVal vsfThis As VSFlexGrid) As Boolean
'���ܣ��ж�ָ�����Ƿ���һ����ҩ�ķ�Χ��,�����,ͬʱ�����кŷ�Χ
'������blnIsHide=��Χ�Ƿ�������ص���
    Dim i As Long, blnTmp As Boolean
    
    With vsfThis
        
        If .TextMatrix(lngRow, COL_Advice_�������) = "" Then Exit Function
        If InStr(",5,6,", .TextMatrix(lngRow, COL_Advice_�������)) = 0 Then Exit Function
        
        If Val(.TextMatrix(lngRow - 1, COL_Advice_���ID)) = Val(.TextMatrix(lngRow, COL_Advice_���ID)) Then
            blnTmp = True
        ElseIf lngRow + 1 <= .Rows - 1 Then
            If Val(.TextMatrix(lngRow + 1, COL_Advice_���ID)) = Val(.TextMatrix(lngRow, COL_Advice_���ID)) Then
                blnTmp = True
            End If
        End If
        If blnTmp Then
            lngBegin = lngRow
            For i = lngRow - 1 To .FixedRows Step -1
                If Val(.TextMatrix(i, COL_Advice_���ID)) = Val(.TextMatrix(lngRow, COL_Advice_���ID)) And Val(.TextMatrix(lngRow, COL_Advice_���ID)) <> 0 _
                    Or ((Val(.TextMatrix(lngRow, COL_Advice_���ID)) = Val(.TextMatrix(i, COL_Advice_ҽ��ID)) _
                    Or Val(.TextMatrix(i, COL_Advice_���ID)) = Val(.TextMatrix(lngRow, COL_Advice_ҽ��ID))) And blnIsHide) Then
                    lngBegin = i
                Else
                    Exit For
                End If
            Next
            lngEnd = lngRow
            For i = lngRow + 1 To .Rows - 1
                If Val(.TextMatrix(i, COL_Advice_���ID)) = Val(.TextMatrix(lngRow, COL_Advice_���ID)) And Val(.TextMatrix(lngRow, COL_Advice_���ID)) <> 0 _
                    Or ((Val(.TextMatrix(lngRow, COL_Advice_���ID)) = Val(.TextMatrix(i, COL_Advice_ҽ��ID)) _
                    Or Val(.TextMatrix(i, COL_Advice_���ID)) = Val(.TextMatrix(lngRow, COL_Advice_ҽ��ID))) And blnIsHide) Then
                    lngEnd = i
                Else
                    Exit For
                End If
            Next
        End If
        RowInһ����ҩ = blnTmp
    End With
End Function

Private Sub txtRate_LostFocus(Index As Integer)
    If txtRate(Index).Text = "" Then
        MsgBox "����ֵ����Ϊ�ա�", vbOKOnly + vbDefaultButton1, Me.Caption
        Call txtRate(Index).SetFocus
    End If
End Sub

Private Sub vsAdvice_DrawCell(Index As Integer, ByVal hDC As Long, ByVal Row As Long, ByVal Col As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, Done As Boolean)
'˵����1.OwnerDrawҪ����ΪOver(������Ԫ��������)
    Dim lngLeft As Long, lngRight As Long
    Dim lngBegin As Long, lngEnd As Long
    Dim vRect As RECT

    With vsAdvice(Index)
        '����һ����ҩ������еı��߼�����
        If Row < .FixedRows Then Exit Sub
        lngLeft = COL_Advice_��Ч: lngRight = COL_Advice_��Ч
        If Not Between(Col, lngLeft, lngRight) Then
            lngLeft = COL_Advice_����: lngRight = COL_Advice_����
        End If
        If Not Between(Col, lngLeft, lngRight) Then Exit Sub

        If Not RowInһ����ҩ(Row, lngBegin, lngEnd, False, vsAdvice(Index)) Then Exit Sub

        vRect.Left = Left '������߱����
        vRect.Right = Right - 1 '�����ұ߱����
        If Row = lngBegin Then
            vRect.Top = Bottom - 1 '���б�����������
            vRect.Bottom = Bottom
        Else
            If Row = lngEnd Then
                vRect.Top = Top
                vRect.Bottom = Bottom - 1 '���б����±���
            Else
                vRect.Top = Top
                vRect.Bottom = Bottom
            End If
        End If
        If Between(Row, .Row, .RowSel) Then
            SetBkColor hDC, OS.SysColor2RGB(.BackColorSel)
        Else
            SetBkColor hDC, OS.SysColor2RGB(.BackColor)
        End If
        ExtTextOut hDC, vRect.Left, vRect.Top, ETO_OPAQUE, vRect, " ", 1, 0
        Done = True
    End With
End Sub

Private Sub vsAdvice_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
         Call SelectVsAdvice(vsAdvice(Index), vsAdvice(Index).Row, COL_Advice_ѡ��)
    End If
End Sub

Private Sub vsAdvice_LostFocus(Index As Integer)
    vsAdvice(Index).Row = 0
End Sub

Private Sub vsAdvice_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        With vsAdvice(Index)
            If .MouseRow < .FixedRows Then Exit Sub
            If .MouseCol <> COL_Advice_ѡ�� Then Exit Sub
            Call SelectVsAdvice(vsAdvice(Index), .Row, .Col)
        End With
    End If
End Sub

Private Sub vsItem_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        Call SelectVsItem(vsItem(Index), vsItem(Index).Row, COL_Item_ѡ��)
    End If
End Sub

Private Sub vsItem_LostFocus(Index As Integer)
    vsItem(Index).Row = 0
End Sub

Private Sub vsItem_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        With vsItem(Index)
            If .MouseRow < .FixedRows Then Exit Sub
            If .MouseCol <> COL_Item_ѡ�� Then Exit Sub
            Call SelectVsItem(vsItem(Index), .Row, .Col)
        End With
    End If
End Sub

Private Sub SelectVsStep(ByVal lngRow As Long, ByVal lngCol As Long)
    With vsStep
        If COL_ѡ�� = .Col And lngRow >= .FixedRows Then
            If .Cell(flexcpChecked, lngRow, COL_ѡ��) = flexChecked Then
                .Cell(flexcpChecked, lngRow, COL_ѡ��) = Unchecked
                .TextMatrix(lngRow, COL_ѡ��) = "0"    'δѡ��
            Else
                .Cell(flexcpChecked, lngRow, COL_ѡ��) = Checked
                .TextMatrix(lngRow, COL_ѡ��) = "1"     'ѡ��
            End If
        End If
    End With
End Sub

Private Sub SelectVsItem(ByVal vsItem As VSFlexGrid, ByVal lngRow As Long, ByVal lngCol As Long)
'����:���ڹ�ѡ��Ŀ�����ѡ����Ŀ��δѡ����Ŀ����ͨ��FindRow����ȷ���Ƿ����δѡ����Ŀ�������
    With vsItem
            If mrsStep Is Nothing Then Exit Sub
            If lngCol = COL_Item_ѡ�� Then
                If .Cell(flexcpChecked, lngRow, lngCol) = flexChecked Then
                    .Cell(flexcpChecked, lngRow, lngCol) = flexUnchecked
                    .TextMatrix(lngRow, lngCol) = "0"       '���δѡ����
                Else
                    .Cell(flexcpChecked, lngRow, lngCol) = flexChecked
                    .TextMatrix(lngRow, lngCol) = "1"       '���ѡ����
                End If
            End If
        End With
End Sub

Private Sub SelectVsAdvice(ByVal vsAdvice As VSFlexGrid, ByVal lngRow As Long, ByVal lngCol As Long)
'����:���ڹ�ѡ��Ŀ�����ѡ����Ŀ��δѡ����Ŀ����ͨ��FindRow����ȷ���Ƿ����δѡ����Ŀ�������

    Dim lngBegin As Long
    Dim lngEnd As Long
    Dim i As Long
        
    With vsAdvice
        If Not mrsStep Is Nothing Then
            If lngRow < .FixedRows Then Exit Sub
        
            If lngCol = COL_Advice_ѡ�� Then   'ȡ����ҽ��
                If .Cell(flexcpChecked, lngRow, lngCol) = flexChecked Then
                  
                    Call RowInһ����ҩ(lngRow, lngBegin, lngEnd, False, vsAdvice)
                    If lngBegin = lngEnd Then lngBegin = lngRow: lngEnd = lngRow '��һ����ҩ
                    For i = lngBegin To lngEnd
                       .Cell(flexcpChecked, i, lngCol) = flexUnchecked   'δѡ��
                       .TextMatrix(i, lngCol) = "0"     '���δѡ����
                    Next
                Else
                    Call RowInһ����ҩ(lngRow, lngBegin, lngEnd, False, vsAdvice)
                    If lngBegin = lngEnd Then lngBegin = lngRow: lngEnd = lngRow '��һ����ҩ
                    For i = lngBegin To lngEnd
                        .Cell(flexcpChecked, i, lngCol) = flexChecked     'ѡ��
                        .TextMatrix(i, lngCol) = "1"    '���ѡ����
                    Next
                End If
            End If
        End If
    End With
End Sub

Private Sub vsStep_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        Call SelectVsStep(vsStep.Row, COL_ѡ��)
    End If
End Sub

Private Sub vsStep_LostFocus()
    vsStep.Row = 0
End Sub

Private Sub vsStep_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        If vsStep.MouseRow < vsStep.FixedRows Then Exit Sub
        If vsStep.MouseCol <> COL_ѡ�� Then Exit Sub
        Call SelectVsStep(vsStep.MouseRow, vsStep.MouseCol)
    End If
End Sub

Private Sub LoadBranch()
'����:���ط�֧·����Ϣ
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    Dim j As Integer
    
    On Error GoTo errH
    strSql = "Select a.id,a.���� as ��֧����  " & vbNewLine & _
                "From �ٴ�·����֧ A, �ٴ�·���׶� B, �ٴ�·���׶� C" & vbNewLine & _
                "Where a.ǰһ�׶�id = b.Id And b.��id = c.Id(+)" & vbNewLine & _
                "And a.·��id = [1] And a.�汾�� = [2]" & vbNewLine & _
                "Order By Nvl(c.���, b.���), a.����"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, mlng·��ID, mlng�汾��)
    If rsTmp Is Nothing Then Exit Sub
    cboBranch.Clear
    cboBranch.AddItem "��·��"
    cboBranch.ItemData(0) = mlng·��ID
    For i = 1 To rsTmp.RecordCount
        cboBranch.AddItem "��֧���ƣ�" & rsTmp!��֧����
        cboBranch.ItemData(i) = rsTmp!ID
        rsTmp.MoveNext
    Next
    Call Cbo.SetIndex(cboBranch.Hwnd, 0)
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub SetVSRowHidden(ByVal vsGrid As VSFlexGrid, ByVal lngId As Long)
'����:���ڷ�֧·��ʱ������·��ID��ʾ��Ӧ�Ľ׶���
'������vsGrid������
    Dim i As Integer
    Dim j As Integer
    Dim k As Integer
    Dim lngBegin As Long
    Dim lngEnd As Long
    Dim lngRow As Long
    Dim strTmp As String, str�걾 As String, str�巨 As String, str���� As String
    
    With vsGrid
        lngBegin = .FixedRows  '��ʼĬ�ϵ�һ��
        For i = .FixedRows To .Rows - 1
           
            If cboBranch.ListCount > 1 Then  '���ڷ�֧·��
                If .RowData(i) = lngId Then
                    .RowHidden(i) = False
                Else
                    .RowHidden(i) = True
                End If
            End If
            
             '�ٴ���һЩ�����е�����,��������ݵ���ʾ
            If vsGrid.Tag = "ȡ��ҽ��" And Not .RowHidden(i) Then
            '��ҩ;��
                If .TextMatrix(i, COL_Advice_�������) = "E" And Val(.TextMatrix(i, COL_Advice_���ID)) = 0 _
                   And Val(.TextMatrix(i - 1, COL_Advice_���ID)) = Val(.TextMatrix(i, COL_Advice_ҽ��ID)) _
                   And InStr(",5,6,", .TextMatrix(i - 1, COL_Advice_�������)) > 0 Then
                    .RowHidden(i) = True
                End If
                
    
                '��Ѫ;��
                If .TextMatrix(i, COL_Advice_�������) = "E" And .TextMatrix(i - 1, COL_Advice_�������) = "K" _
                   And Val(.TextMatrix(i, COL_Advice_���ID)) = Val(.TextMatrix(i - 1, COL_Advice_ҽ��ID)) Then
                    .RowHidden(i) = True
                End If
    
                '��ҩ�䷽�ͼ������
                If .TextMatrix(i, COL_Advice_�������) = "E" And Val(.TextMatrix(i, COL_Advice_���ID)) = 0 _
                   And Val(.TextMatrix(i - 1, COL_Advice_���ID)) = Val(.TextMatrix(i, COL_Advice_ҽ��ID)) _
                   And InStr(",7,E,C,", .TextMatrix(i - 1, COL_Advice_�������)) > 0 Then
    
                    str�巨 = "": str�걾 = "": strTmp = ""
                    j = .FindRow(CStr(Val(.TextMatrix(i, COL_Advice_ҽ��ID))), , COL_Advice_���ID)
    
                    'j--��ϼ�����Ŀ���к�
                    For k = j To i - 1
                        .RowHidden(k) = k <> i
                        If .TextMatrix(k, COL_Advice_�������) = "C" Then
                            strTmp = strTmp & "," & .TextMatrix(k, COL_Advice_ҽ������)
                            str�걾 = .TextMatrix(j, COL_Advice_�걾��λ)    'ȡ��һ��������Ŀ�ı걾
                        ElseIf .TextMatrix(k, COL_Advice_�������) = "E" And Val(.TextMatrix(k, COL_Advice_���ID)) <> 0 Then
                            str�巨 = .TextMatrix(k, COL_Advice_ҽ������)
                        End If
                    Next
    
                    If .TextMatrix(i - 1, COL_Advice_�������) = "C" Then
                        .TextMatrix(i, COL_Advice_ҽ������) = Mid(strTmp, 2) & IIf(str�걾 <> "", "(" & str�걾 & ")", "")
                    Else
                        .TextMatrix(i, COL_Advice_ҽ������) = "��ҩ�䷽," & str�巨 & "," & .TextMatrix(i, COL_Advice_ҽ������)
                    End If
                End If
    
                '������
                If .TextMatrix(i, COL_Advice_�������) = "D" And Val(.TextMatrix(i, COL_Advice_���ID)) = 0 Then
                    str�걾 = "": str�巨 = "": strTmp = ""
                    For j = i + 1 To .Rows - 1
                        If Val(.TextMatrix(j, COL_Advice_���ID)) = Val(.TextMatrix(i, COL_Advice_ҽ��ID)) Then
                            .RowHidden(j) = True
                            If .TextMatrix(j, COL_Advice_�걾��λ) <> "" _
                               And Val(.TextMatrix(j, COL_Advice_������ĿID)) = Val(.TextMatrix(i, COL_Advice_������ĿID)) Then    '��ͬ����ĿID�����·�ʽ
                                If .TextMatrix(j, COL_Advice_�걾��λ) <> strTmp And strTmp <> "" Then
                                    str�걾 = str�걾 & "," & strTmp & IIf(str�巨 <> "", "(" & Mid(str�巨, 2) & ")", "")
                                    str�巨 = ""
                                End If
                                strTmp = .TextMatrix(j, COL_Advice_�걾��λ)
                            End If
                        Else
                            Exit For
                        End If
                    Next
                    If strTmp <> "" Then
                        str�걾 = str�걾 & "," & strTmp & IIf(str�巨 <> "", "(" & Mid(str�巨, 2) & ")", "")
                    End If
                    If str�걾 <> "" Then    '��ǰ�ļ�鷽ʽʱ����ʾ��ϸҽ������
                        .TextMatrix(i, COL_Advice_ҽ������) = .TextMatrix(i, COL_Advice_ҽ������) & ":" & Mid(str�걾, 2)
                    End If
                End If
    
                '������Ŀ
                If .TextMatrix(i, COL_Advice_�������) = "F" And Val(.TextMatrix(i, COL_Advice_���ID)) = 0 Then
                    strTmp = "": str���� = ""
                    For j = i + 1 To .Rows - 1
                        If Val(.TextMatrix(j, COL_Advice_���ID)) = Val(.TextMatrix(i, COL_Advice_ҽ��ID)) Then
                            .RowHidden(j) = True
                            If .TextMatrix(j, COL_Advice_�������) = "F" Then
                                strTmp = strTmp & "," & .TextMatrix(j, COL_Advice_ҽ������)
                            ElseIf .TextMatrix(j, COL_Advice_�������) = "G" Then
                                str���� = .TextMatrix(j, COL_Advice_ҽ������)
                            End If
                        Else
                            Exit For
                        End If
                    Next
                    If strTmp <> "" Or str���� <> "" Then
                        If str���� <> "" Then
                            .TextMatrix(i, COL_Advice_ҽ������) = "�� " & str���� & " ���� " & .TextMatrix(i, COL_Advice_ҽ������)
                        Else
                            .TextMatrix(i, COL_Advice_ҽ������) = "�� " & .TextMatrix(i, COL_Advice_ҽ������)
                        End If
                        If strTmp <> "" Then
                            .TextMatrix(i, COL_Advice_ҽ������) = .TextMatrix(i, COL_Advice_ҽ������) & " �� " & Mid(strTmp, 2)
                        End If
                    End If
                End If
                     '��ͬ�׶δ���ҽ��������ͬʱ����ҽ�����ݴ���Ϊ��ҽ�����ƣ���Ŀ���ƣ��ķ�ʽ��������
                If i > .FixedRows And i <= .Rows - 1 Then
                    If .Cell(flexcpData, i, COL_Advice_�׶�) <> .Cell(flexcpData, i - 1, COL_Advice_�׶�) Or i = .Rows - 1 Then '��һ�׶�����һ�׶ν��Ӵ�
                        lngEnd = IIf(i = .Rows - 1, i, i - 1)
                        For j = lngBegin To lngEnd
                            If Not .RowHidden(j) Then
                                lngRow = .FindRow(.TextMatrix(j, COL_Advice_ҽ������), j + 1, COL_Advice_ҽ������) '����������
                                '��ͬ�׶��������ҵ���ͬҽ��
                                If lngRow <> -1 And lngRow > lngBegin And lngRow <= lngEnd Then
                                    .TextMatrix(j, COL_Advice_ҽ������) = .TextMatrix(j, COL_Advice_ҽ������) & "(" & .Cell(flexcpData, j, COL_Advice_��Ŀ����) & ")"
                                    .TextMatrix(lngRow, COL_Advice_ҽ������) = .TextMatrix(lngRow, COL_Advice_ҽ������) & "(" & .Cell(flexcpData, lngRow, COL_Advice_��Ŀ����) & ")"
                                End If
                            End If
                        Next
                        lngBegin = lngEnd + 1 '��һ�׶�����
                    End If
                End If
            End If
            
        Next
        
   
        .AutoSize .FixedCols, .Cols - 1, , 45
    End With

End Sub

Private Sub ClearData()
'����:�����������
    
    vsStep.Rows = vsStep.FixedRows
    vsStep.Rows = vsStep.FixedRows + 1
     
    vsItem(Index_DEL).Rows = vsItem(Index_DEL).FixedRows
    vsItem(Index_DEL).Rows = vsItem(Index_DEL).FixedRows + 1
    
    vsItem(Index_Add).Rows = vsItem(Index_Add).FixedRows
    vsItem(Index_Add).Rows = vsItem(Index_Add).FixedRows + 1
    
    vsAdvice(Index_DEL).Rows = vsAdvice(Index_DEL).FixedRows
    vsAdvice(Index_DEL).Rows = vsAdvice(Index_DEL).FixedRows + 1
    
    vsAdvice(Index_Add).Rows = vsAdvice(Index_Add).FixedRows
    vsAdvice(Index_Add).Rows = vsAdvice(Index_Add).FixedRows + 1

    If Not mrsStep Is Nothing Then
        Set mrsStep = Nothing
    End If

End Sub

