VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCISManage 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   Caption         =   "���Ӳ���������Ȩ"
   ClientHeight    =   10920
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   18780
   Icon            =   "frmCISManage.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   10920
   ScaleWidth      =   18780
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picApply 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   5160
      ScaleHeight     =   3495
      ScaleWidth      =   3255
      TabIndex        =   8
      Top             =   3480
      Width           =   3255
      Begin VB.Frame fraFillter 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "��ѯ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   735
         Left            =   0
         TabIndex        =   31
         Top             =   0
         Width           =   17055
         Begin VB.CommandButton cmdFind 
            Caption         =   "��ѯ(&F)"
            Height          =   375
            Left            =   12120
            TabIndex        =   17
            Top             =   240
            Width           =   1215
         End
         Begin VB.CheckBox chkFilter 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "������"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   0
            Left            =   7320
            MaskColor       =   &H00FFC0C0&
            TabIndex        =   13
            Top             =   240
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkFilter 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "������"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   1
            Left            =   8565
            MaskColor       =   &H00FFC0C0&
            TabIndex        =   14
            Top             =   240
            Value           =   1  'Checked
            Width           =   855
         End
         Begin VB.CheckBox chkFilter 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "������"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   2
            Left            =   9795
            MaskColor       =   &H00FFC0C0&
            TabIndex        =   15
            Top             =   240
            Width           =   855
         End
         Begin VB.CheckBox chkFilter 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "�Ѿܾ�"
            ForeColor       =   &H00000000&
            Height          =   375
            Index           =   3
            Left            =   11040
            MaskColor       =   &H00FFC0C0&
            TabIndex        =   16
            Top             =   240
            Width           =   855
         End
         Begin VB.ComboBox cboTime 
            Height          =   300
            Left            =   1290
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   277
            Width           =   1365
         End
         Begin MSComCtl2.DTPicker dtpTime 
            Height          =   300
            Index           =   0
            Left            =   2715
            TabIndex        =   11
            Top             =   270
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   215285763
            CurrentDate     =   40976
         End
         Begin MSComCtl2.DTPicker dtpTime 
            Height          =   300
            Index           =   1
            Left            =   5010
            TabIndex        =   26
            Top             =   270
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   215285763
            CurrentDate     =   40976
         End
         Begin VB.Image imgTime 
            Height          =   240
            Index           =   2
            Left            =   225
            Picture         =   "frmCISManage.frx":6852
            Top             =   300
            Width           =   240
         End
         Begin VB.Label lbltime 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "����ʱ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   32
            Top             =   330
            Width           =   855
         End
         Begin VB.Line Line 
            BorderColor     =   &H80000000&
            BorderWidth     =   3
            X1              =   4725
            X2              =   4925
            Y1              =   420
            Y2              =   420
         End
         Begin VB.Image imgFilter 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   0
            Left            =   7020
            Picture         =   "frmCISManage.frx":6DDC
            Stretch         =   -1  'True
            Top             =   300
            Width           =   240
         End
         Begin VB.Image imgFilter 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   1
            Left            =   8250
            Picture         =   "frmCISManage.frx":D62E
            Stretch         =   -1  'True
            Top             =   300
            Width           =   240
         End
         Begin VB.Image imgFilter 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   2
            Left            =   9495
            Picture         =   "frmCISManage.frx":13E80
            Stretch         =   -1  'True
            Top             =   300
            Width           =   240
         End
         Begin VB.Image imgFilter 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   3
            Left            =   10725
            Picture         =   "frmCISManage.frx":1A6D2
            Stretch         =   -1  'True
            Top             =   300
            Width           =   240
         End
      End
      Begin VB.Frame picInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "��Ȩ��Ϣ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   7335
         Left            =   12840
         TabIndex        =   25
         Top             =   840
         Width           =   4095
         Begin VSFlex8Ctl.VSFlexGrid vsInfo 
            Height          =   6915
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   4750
            _cx             =   8378
            _cy             =   12197
            Appearance      =   0
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
            MouseIcon       =   "frmCISManage.frx":20F24
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16444122
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   16777215
            GridColorFixed  =   16777215
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   0
            FocusRect       =   0
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   0
            GridLinesFixed  =   0
            GridLineWidth   =   1
            Rows            =   8
            Cols            =   1
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   400
            RowHeightMax    =   10000
            ColWidthMin     =   4650
            ColWidthMax     =   10000
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmCISManage.frx":217FE
            ScrollTrack     =   -1  'True
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
            BackColorFrozen =   16777215
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsList 
         Height          =   7275
         Left            =   0
         TabIndex        =   18
         Top             =   840
         Width           =   12645
         _cx             =   22304
         _cy             =   12832
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
         MouseIcon       =   "frmCISManage.frx":21827
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772554
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   16777215
         GridColor       =   16119285
         GridColorFixed  =   16777215
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   3
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   400
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   10000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCISManage.frx":22101
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
         Begin VB.PictureBox picTmp 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   0
            Left            =   1920
            ScaleHeight     =   240
            ScaleWidth      =   480
            TabIndex        =   33
            Top             =   1680
            Visible         =   0   'False
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox picLog 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3375
      Left            =   1800
      ScaleHeight     =   3375
      ScaleWidth      =   3255
      TabIndex        =   27
      Top             =   3480
      Width           =   3255
      Begin VB.Frame fraLog 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "��ѯ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   735
         Left            =   120
         TabIndex        =   35
         Top             =   120
         Width           =   17055
         Begin VB.ComboBox cboLogTime 
            Height          =   300
            Left            =   1290
            Style           =   2  'Dropdown List
            TabIndex        =   20
            Top             =   277
            Width           =   1365
         End
         Begin VB.CommandButton cmdLogFind 
            Caption         =   "��ѯ(&F)"
            Height          =   375
            Left            =   7200
            TabIndex        =   23
            Top             =   240
            Width           =   1215
         End
         Begin MSComCtl2.DTPicker dtpLogTime 
            Height          =   300
            Index           =   0
            Left            =   2715
            TabIndex        =   21
            Top             =   270
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   215285763
            CurrentDate     =   40976
         End
         Begin MSComCtl2.DTPicker dtpLogTime 
            Height          =   300
            Index           =   1
            Left            =   5040
            TabIndex        =   22
            Top             =   270
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   215285763
            CurrentDate     =   40976
         End
         Begin VB.Image imgTime 
            Height          =   240
            Index           =   0
            Left            =   240
            Picture         =   "frmCISManage.frx":2219C
            Top             =   300
            Width           =   240
         End
         Begin VB.Line LineTmp 
            BorderColor     =   &H80000000&
            BorderWidth     =   3
            Index           =   1
            X1              =   4725
            X2              =   4925
            Y1              =   420
            Y2              =   420
         End
         Begin VB.Label lbltime 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "����ʱ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   255
            Index           =   2
            Left            =   480
            TabIndex        =   36
            Top             =   330
            Width           =   855
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsLog 
         Height          =   7275
         Left            =   120
         TabIndex        =   24
         Top             =   960
         Width           =   12645
         _cx             =   22304
         _cy             =   12832
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
         MouseIcon       =   "frmCISManage.frx":22726
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772554
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   16777215
         GridColor       =   16119285
         GridColorFixed  =   16777215
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   3
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   400
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   10000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCISManage.frx":23000
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
         Begin VB.PictureBox picTmp 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   2
            Left            =   1920
            ScaleHeight     =   240
            ScaleWidth      =   480
            TabIndex        =   37
            Top             =   1680
            Visible         =   0   'False
            Width           =   480
         End
      End
   End
   Begin VB.PictureBox picManage 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   4215
      Left            =   9120
      ScaleHeight     =   4215
      ScaleWidth      =   4455
      TabIndex        =   9
      Top             =   3600
      Width           =   4455
      Begin VB.Frame fraManageInfo 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "��Ȩ��Ϣ"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   7335
         Left            =   13080
         TabIndex        =   30
         Top             =   960
         Width           =   4095
         Begin VSFlex8Ctl.VSFlexGrid vsManageInfo 
            Height          =   6915
            Left            =   120
            TabIndex        =   6
            Top             =   240
            Width           =   4750
            _cx             =   8378
            _cy             =   12197
            Appearance      =   0
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
            MouseIcon       =   "frmCISManage.frx":2309B
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16444122
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   16777215
            GridColorFixed  =   16777215
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   0
            FocusRect       =   0
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   1
            GridLines       =   0
            GridLinesFixed  =   0
            GridLineWidth   =   1
            Rows            =   11
            Cols            =   1
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   400
            RowHeightMax    =   10000
            ColWidthMin     =   4650
            ColWidthMax     =   10000
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmCISManage.frx":23975
            ScrollTrack     =   -1  'True
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
            BackColorFrozen =   16777215
            ForeColorFrozen =   0
            WallPaperAlignment=   9
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
      End
      Begin VB.Frame fraManageFilter 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Caption         =   "��ѯ����"
         BeginProperty Font 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000002&
         Height          =   735
         Left            =   120
         TabIndex        =   28
         Top             =   120
         Width           =   17055
         Begin VB.CheckBox chk������ 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "��ʾ�����ϵ���Ȩ"
            ForeColor       =   &H00000000&
            Height          =   375
            Left            =   7515
            MaskColor       =   &H00FFC0C0&
            TabIndex        =   3
            Top             =   240
            Width           =   1815
         End
         Begin VB.CommandButton cmdManageFind 
            Caption         =   "��ѯ(&F)"
            Height          =   375
            Left            =   9480
            TabIndex        =   4
            Top             =   240
            Width           =   1215
         End
         Begin VB.ComboBox cboManageTime 
            Height          =   300
            Left            =   1290
            Style           =   2  'Dropdown List
            TabIndex        =   0
            Top             =   277
            Width           =   1365
         End
         Begin MSComCtl2.DTPicker dtpManageTime 
            Height          =   300
            Index           =   0
            Left            =   2715
            TabIndex        =   1
            Top             =   270
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   215285763
            CurrentDate     =   40976
         End
         Begin MSComCtl2.DTPicker dtpManageTime 
            Height          =   300
            Index           =   1
            Left            =   5040
            TabIndex        =   2
            Top             =   270
            Width           =   1935
            _ExtentX        =   3413
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy-MM-dd HH:mm"
            Format          =   215285763
            CurrentDate     =   40976
         End
         Begin VB.Image imgTime 
            Height          =   240
            Index           =   1
            Left            =   200
            Picture         =   "frmCISManage.frx":2399E
            Top             =   300
            Width           =   240
         End
         Begin VB.Image imgFilter 
            Appearance      =   0  'Flat
            Height          =   240
            Index           =   4
            Left            =   7200
            Picture         =   "frmCISManage.frx":23F28
            Stretch         =   -1  'True
            Top             =   300
            Width           =   240
         End
         Begin VB.Label lbltime 
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFFF&
            Caption         =   "��Ȩʱ��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H80000002&
            Height          =   255
            Index           =   1
            Left            =   480
            TabIndex        =   29
            Top             =   330
            Width           =   855
         End
         Begin VB.Line LineTmp 
            BorderColor     =   &H80000000&
            BorderWidth     =   3
            Index           =   0
            X1              =   4725
            X2              =   4925
            Y1              =   420
            Y2              =   420
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsManage 
         Height          =   7275
         Left            =   120
         TabIndex        =   5
         Top             =   960
         Width           =   12645
         _cx             =   22304
         _cy             =   12832
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
         MouseIcon       =   "frmCISManage.frx":2A77A
         BackColor       =   16777215
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772554
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   16777215
         GridColor       =   16119285
         GridColorFixed  =   16777215
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   0
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   3
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   7
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   400
         RowHeightMax    =   2000
         ColWidthMin     =   0
         ColWidthMax     =   10000
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmCISManage.frx":2B054
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
         Begin VB.PictureBox picTmp 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000005&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   240
            Index           =   1
            Left            =   1920
            ScaleHeight     =   240
            ScaleWidth      =   480
            TabIndex        =   34
            Top             =   1680
            Visible         =   0   'False
            Width           =   480
         End
      End
   End
   Begin XtremeSuiteControls.TabControl tbcSub 
      Height          =   5580
      Left            =   240
      TabIndex        =   7
      Top             =   840
      Width           =   8130
      _Version        =   589884
      _ExtentX        =   14340
      _ExtentY        =   9842
      _StockProps     =   64
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   12
      Top             =   10560
      Width           =   18780
      _ExtentX        =   33126
      _ExtentY        =   635
      SimpleText      =   $"frmCISManage.frx":2B0EF
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmCISManage.frx":2B136
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   28046
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1058
            MinWidth        =   1058
            Text            =   "����"
            TextSave        =   "����"
            Key             =   "STANUM"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   1
            Alignment       =   1
            Enabled         =   0   'False
            Object.Width           =   1058
            MinWidth        =   1058
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
   Begin MSComctlLib.ImageList img16 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISManage.frx":2B9CA
            Key             =   "girl"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISManage.frx":3222C
            Key             =   "boy"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISManage.frx":38A8E
            Key             =   "����ʱ��"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISManage.frx":39028
            Key             =   "��������"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISManage.frx":395C2
            Key             =   "����ҽ��"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISManage.frx":39B5C
            Key             =   "���ʲ���"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISManage.frx":3A0F6
            Key             =   "AllCheck"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmCISManage.frx":3A250
            Key             =   "unCheck"
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   600
      Top             =   240
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
End
Attribute VB_Name = "frmCISManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private Enum colList
    COL_����ID = 1
    COL_�������� = 2
    COL_����ʱ�� = 3
    COL_����ʱ�� = 4
    COL_������ = 5

    COL_����ʱ�� = 6
    COL_������ = 7
    COL_������ʲ��� = 8
    COL_���ʿ�ʼʱ�� = 9
    COL_���ʽ���ʱ�� = 10
    COL_����ԭ�� = 11
    COL_����״̬ = 12
End Enum

Private Enum colManage
    '������
    COLM_ID = 1
    COLM_�������� = 2
    COLM_����ʱ�� = 3 '���סԺ������0-�����ƣ�1-δ�鵵������2-�ѹ鵵���������ﲡ��������'
    COLM_��Ȩ���� = 4 '0-��������,1-������Ȩ
    COLM_���ʲ��� = 5 '0-ȫԺ���ˣ�1-���Ʋ��ˣ�2-ָ�����Ҳ��ˣ�3-ָ�����ˣ�4-���Ϊָ�������Ĳ��ˣ�5-ָ�������Ĳ���
    COLM_���˷�Χ���� = 6
    '��ʾ��
    COLM_������ = 7
    COLM_��ע = 8
    COLM_���ʿ�ʼʱ�� = 9
    COLM_���ʽ���ʱ�� = 10
    COLM_��Ȩ�� = 11
    COLM_��Ȩʱ�� = 12
    COLM_������ = 13
    COLM_����ʱ�� = 14
    COLM_������ = 15
End Enum


Private Enum colLog
    '������
    COLG_ID = 1
    COLG_����ID = 2
    COLG_����ID = 3 '����Ϊ�Һ�ID��סԺΪ��ҳID';
    COLG_������Դ = 4 '1-���ﲡ�ˣ�2-סԺ����
    COLG_����ID = 5  '����ID�м�¼��Ӧ��ҵ���ļ���ʶID
    
    '��ʾ��
    COLG_����ʱ�� = 6
    COLG_������ = 7
    COLG_�������� = 8
    COLG_�����Ա� = 9
    COLG_�������� = 10
    COLG_���˱�ʶ�� = 11
    COLG_���˿��� = 12
    COLG_�������� = 13
    COLG_�������� = 14
End Enum



Private Enum RowInfo
    Row_���ʲ��˱��� = 0
    Row_���ʲ��� = 1
    Row_����ʱ�ޱ��� = 3
    Row_����ʱ�� = 4
    Row_�������ݱ��� = 6
    Row_�������� = 7
End Enum

Private Enum RowMInfo
    RowM_�����߱��� = 0
    RowM_������ = 1
    RowM_���ʲ��˱��� = 3
    RowM_���ʲ��� = 4
    RowM_����ʱ�ޱ��� = 6
    RowM_����ʱ�� = 7
    RowM_�������ݱ��� = 9
    RowM_�������� = 10
End Enum



Private Sub cboManageTime_Click()
    Dim curDate As Date
    
    dtpManageTime(0).Enabled = cboManageTime.ListIndex = cboManageTime.ListCount - 1
    dtpManageTime(1).Enabled = cboManageTime.ListIndex = cboManageTime.ListCount - 1
    
    curDate = zldatabase.Currentdate

    dtpManageTime(0).MaxDate = curDate + 1
    dtpManageTime(1).MaxDate = curDate + 1

    
    Select Case cboManageTime.ListIndex
    Case 0 '����
        dtpManageTime(0).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
        dtpManageTime(1).Value = Format(curDate, "yyyy-MM-dd 23:59:59")
    Case 1 '�������
        dtpManageTime(0).Value = Format(DateAdd("d", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpManageTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm")
    Case 2 '�������
        dtpManageTime(0).Value = Format(DateAdd("d", -2, curDate), "yyyy-MM-dd 00:00:00")
        dtpManageTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm")
    Case 3 '���һ��
        dtpManageTime(0).Value = Format(DateAdd("ww", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpManageTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm")
    Case 4 '���һ��
        dtpManageTime(0).Value = Format(DateAdd("m", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpManageTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm")
    Case 5 'ָ  ��
        If Me.Visible Then
            dtpManageTime(0).SetFocus
        End If
    End Select
End Sub

Private Sub cboLogTime_Click()
    Dim curDate As Date
    
    dtpLogTime(0).Enabled = cboLogTime.ListIndex = cboLogTime.ListCount - 1
    dtpLogTime(1).Enabled = cboLogTime.ListIndex = cboLogTime.ListCount - 1
    
    curDate = zldatabase.Currentdate

    dtpLogTime(0).MaxDate = curDate + 1
    dtpLogTime(1).MaxDate = curDate + 1

    
    Select Case cboLogTime.ListIndex
    Case 0 '����
        dtpLogTime(0).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
        dtpLogTime(1).Value = Format(curDate, "yyyy-MM-dd 23:59:59")
    Case 1 '�������
        dtpLogTime(0).Value = Format(DateAdd("d", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpLogTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm")
    Case 2 '�������
        dtpLogTime(0).Value = Format(DateAdd("d", -2, curDate), "yyyy-MM-dd 00:00:00")
        dtpLogTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm")
    Case 3 '���һ��
        dtpLogTime(0).Value = Format(DateAdd("ww", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpLogTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm")
    Case 4 '���һ��
        dtpLogTime(0).Value = Format(DateAdd("m", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpLogTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm")
    Case 5 'ָ  ��
        If Me.Visible Then
            dtpLogTime(0).SetFocus
        End If
    End Select
End Sub

Private Sub cboTime_Click()
    Dim curDate As Date
    
    dtpTime(0).Enabled = cboTime.ListIndex = cboTime.ListCount - 1
    dtpTime(1).Enabled = cboTime.ListIndex = cboTime.ListCount - 1
    
    curDate = zldatabase.Currentdate

    dtpTime(0).MaxDate = curDate + 1
    dtpTime(1).MaxDate = curDate + 1

    
    Select Case cboTime.ListIndex
    Case 0 '����
        dtpTime(0).Value = Format(curDate, "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd 23:59:59")
    Case 1 '�������
        dtpTime(0).Value = Format(DateAdd("d", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm")
    Case 2 '�������
        dtpTime(0).Value = Format(DateAdd("d", -2, curDate), "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm")
    Case 3 '���һ��
        dtpTime(0).Value = Format(DateAdd("ww", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm")
    Case 4 '���һ��
        dtpTime(0).Value = Format(DateAdd("m", -1, curDate), "yyyy-MM-dd 00:00:00")
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm")
    Case 5 'ָ  ��
        If Me.Visible Then
            dtpTime(0).SetFocus
        End If
    End Select
End Sub


Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim lngApplyID As Long
    Select Case Control.ID
        Case conMenu_Edit_ApplyAdd
            If frmCISManageEdit.ShowEdit(Me, 0, lngApplyID) Then
                   Call LoadManage(lngApplyID)
            End If
        Case conMenu_Edit_ApplyEdit
            If Val(vsManage.TextMatrix(vsManage.Row, COLM_ID)) = 0 Then Exit Sub
            lngApplyID = Val(vsManage.TextMatrix(vsManage.Row, COLM_ID))
            If frmCISManageEdit.ShowEdit(Me, 1, lngApplyID) Then
                Call LoadManage(lngApplyID)
            End If
        Case conMenu_Edit_Delete
            If tbcSub.Selected.Tag = "��Ȩ��¼" Then
                If Val(vsManage.TextMatrix(vsManage.Row, COLM_ID)) = 0 Or vsManage.TextMatrix(vsManage.Row, COLM_����ʱ��) <> "" Then Exit Sub
                lngApplyID = Val(vsManage.TextMatrix(vsManage.Row, COLM_ID))
                If ManageDelete(lngApplyID) Then
                    Call LoadManage(lngApplyID)
                End If
            ElseIf tbcSub.Selected.Tag = "������¼" Then
                If Val(vsList.TextMatrix(vsList.Row, COL_����ID)) = 0 Or vsList.TextMatrix(vsList.Row, COL_����״̬) <> "������" Then Exit Sub
                lngApplyID = Val(vsList.TextMatrix(vsList.Row, COL_����ID))
                If ApplyUpdate(lngApplyID, 2) Then
                    Call LoadList(lngApplyID)
                End If
            End If
        Case conMenu_Manage_Complete
            If tbcSub.Selected.Tag = "������¼" Then
                If Val(vsList.TextMatrix(vsList.Row, COL_����ID)) = 0 Or vsList.TextMatrix(vsList.Row, COL_����״̬) <> "������" Then Exit Sub
                lngApplyID = Val(vsList.TextMatrix(vsList.Row, COL_����ID))
                If ApplyUpdate(lngApplyID, 1) Then
                    Call LoadList(lngApplyID)
                End If
            End If
        Case conMenu_Manage_Undone
            If tbcSub.Selected.Tag = "������¼" Then
                If Val(vsList.TextMatrix(vsList.Row, COL_����ID)) = 0 Or vsList.TextMatrix(vsList.Row, COL_����״̬) <> "������" Then Exit Sub
                lngApplyID = Val(vsList.TextMatrix(vsList.Row, COL_����ID))
                If ApplyUpdate(lngApplyID, 3) Then
                    Call LoadList(lngApplyID)
                End If
            End If
        Case conMenu_Edit_Untread
            If tbcSub.Selected.Tag = "������¼" Then
                If Val(vsList.TextMatrix(vsList.Row, COL_����ID)) = 0 Or vsList.TextMatrix(vsList.Row, COL_����״̬) <> "�Ѿܾ�" Then Exit Sub
                lngApplyID = Val(vsList.TextMatrix(vsList.Row, COL_����ID))
                If ApplyUpdate(lngApplyID, 5) Then
                    Call LoadList(lngApplyID)
                End If
            End If
        Case conMenu_View_Refresh
            If tbcSub.Selected.Tag = "��Ȩ��¼" Then
                Call LoadManage
            ElseIf tbcSub.Selected.Tag = "������¼" Then
                Call LoadList
            End If
        Case conMenu_File_Excel
            Call zlRptPrint(3)
        Case conMenu_Help_Web_Home 'Web�ϵ�����
            Call zlHomePage(Me.hwnd)
        Case conMenu_Help_Web_Forum '������̳
            Call zlWebForum(Me.hwnd)
        Case conMenu_Help_Web_Mail '���ͷ���
            Call zlMailTo(Me.hwnd)
        Case conMenu_Help_About '����
            Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
        Case conMenu_Help_Help '����
            Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100))
        Case conMenu_File_Exit '�˳�
            Unload Me
    End Select
End Sub

Private Function ApplyUpdate(lngApplyID As Long, ByVal lngType As Long) As Boolean
    'lngType '1-������2-���ϣ�3-�ܾ�'��5-ȡ���ܾ�'
    Dim strSQL As String
    Dim curDate As Date
    Dim blnTran As Boolean
    
    On Error GoTo errH
    
    If MsgBox("ȷ��Ҫ" & Decode(lngType, 1, "����", 2, "����", 3, "�ܾ�", 5, "��") & "ѡ�е���Ȩ�����¼" & IIf(lngType = 5, "ȡ���ܾ�", "") & "��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
    curDate = zldatabase.Currentdate
    strSQL = "Zl_���Ӳ�����������_����״̬(" & lngApplyID & "," & lngType & ",'" & UserInfo.���� & "',To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
    Screen.MousePointer = 11
    gcnOracle.BeginTrans: blnTran = True
    Call zldatabase.ExecuteProcedure(strSQL, Me.Caption)
    gcnOracle.CommitTrans: blnTran = False
    Screen.MousePointer = 0
    ApplyUpdate = True
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans: blnTran = False
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
End Function

Private Function ManageDelete(ByVal lng��ȨID As Long) As Boolean
    '��Ȩ����
    Dim strSQL As String
    Dim curDate As Date
    Dim blnTran As Boolean
    
    On Error GoTo errH
    If MsgBox("ȷ��Ҫ����ѡ�е���Ȩ��¼��", vbQuestion + vbYesNo + vbDefaultButton1, gstrSysName) = vbNo Then Exit Function
    curDate = zldatabase.Currentdate
    strSQL = "Zl_���Ӳ���������Ȩ_����(" & lng��ȨID & ",'" & UserInfo.���� & "',To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'))"
    Screen.MousePointer = 11
    gcnOracle.BeginTrans: blnTran = True
    Call zldatabase.ExecuteProcedure(strSQL, Me.Caption)
    gcnOracle.CommitTrans: blnTran = False
    Screen.MousePointer = 0
    ManageDelete = True
    Exit Function
errH:
    If blnTran Then gcnOracle.RollbackTrans: blnTran = False
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Screen.MousePointer = 11
        Resume
    End If
    Call SaveErrLog
End Function




Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    '���ÿɼ�
    Select Case Control.ID
    Case conMenu_Edit_ApplyAdd
        Control.Visible = tbcSub.Selected.Tag = "��Ȩ��¼"
    Case conMenu_Edit_ApplyEdit
        If tbcSub.Selected.Tag = "��Ȩ��¼" Then
            Control.Visible = Val(vsManage.TextMatrix(vsManage.Row, COLM_ID)) <> 0 And vsManage.TextMatrix(vsManage.Row, COLM_������) = ""
        Else
            Control.Visible = False
        End If
    Case conMenu_Manage_Complete
        If tbcSub.Selected.Tag = "������¼" Then
            Control.Visible = vsList.TextMatrix(vsList.Row, COL_����״̬) = "������"
        Else
            Control.Visible = False
        End If
    Case conMenu_Manage_Undone
        If tbcSub.Selected.Tag = "������¼" Then
            Control.Visible = vsList.TextMatrix(vsList.Row, COL_����״̬) = "������"
        Else
            Control.Visible = False
        End If
    Case conMenu_Edit_Delete
        If tbcSub.Selected.Tag = "��Ȩ��¼" Then
            Control.Visible = Val(vsManage.TextMatrix(vsManage.Row, COLM_ID)) <> 0
        ElseIf tbcSub.Selected.Tag = "������¼" Then
            Control.Visible = vsList.TextMatrix(vsList.Row, COL_����״̬) = "������"
        Else
            Control.Visible = False
        End If
    Case conMenu_File_Excel
        Control.Visible = tbcSub.Selected.Tag = "������־"
    Case conMenu_Edit_Untread
        If tbcSub.Selected.Tag = "������¼" Then
            Control.Visible = vsList.TextMatrix(vsList.Row, COL_����״̬) = "�Ѿܾ�"
        Else
            Control.Visible = False
        End If
    End Select
End Sub


Private Sub chkFilter_Click(Index As Integer)
    Dim i As Long
    Dim blnCheck As Boolean
    
    For i = 0 To 3
        If chkFilter(i).Value = 1 Then
            blnCheck = True
            Exit For
        End If
    Next
    If Not blnCheck Then
        MsgBox "������ѡ��һ�ַ������ڹ��ˡ�", vbInformation, gstrSysName
        chkFilter(Index).Value = 1
        Exit Sub
    End If
End Sub


Private Sub cmdFind_Click()
    Call LoadList
End Sub


Private Sub LoadManage(Optional lng��ȨID As Long)
    Dim strSQL As String
    Dim i As Long
    Dim rsTmp As ADODB.Recordset
    Dim curDate As Date

    On Error GoTo errH
    If cboManageTime.ListIndex <> 5 Then
        curDate = zldatabase.Currentdate
        dtpManageTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm:ss")
    End If
    
    strSQL = "Select a.Id, a.��Ȩ����, a.����id, a.������, a.���ʲ���, a.���ʿ�ʼʱ��, a.���ʽ���ʱ��, a.����ʱ��, a.��Ȩ��, a.��Ȩʱ��, a.������, a.����ʱ��,a.��ע" & vbNewLine & _
                "From ���Ӳ���������Ȩ A Where A.��Ȩ���� = 1  And a.��Ȩʱ�� Between [1] And [2]" & IIf(chk������.Value = 0, " And A.����ʱ�� is null", "") & vbNewLine & _
                "Order by A.id"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(Format(dtpManageTime(0).Value, "yyyy-MM-dd hh:mm")), CDate(Format(IIf(cboManageTime.ListIndex <> 5, dtpManageTime(1).Value + 1, dtpManageTime(1).Value), "yyyy-MM-dd hh:mm")))
    With vsManage
        .Rows = .FixedRows
        If Not rsTmp.EOF Then
             .Redraw = flexRDNone
             .Rows = .FixedRows + rsTmp.RecordCount
             For i = 1 To rsTmp.RecordCount
                '������
                .TextMatrix(i, COLM_ID) = Val(rsTmp!ID & "")
                .TextMatrix(i, COLM_����ʱ��) = Val(rsTmp!����ʱ�� & "")
                .TextMatrix(i, COLM_��Ȩ����) = Val(rsTmp!��Ȩ���� & "")
                .TextMatrix(i, COLM_���ʲ���) = Val(rsTmp!���ʲ��� & "")
                
                    
                 '��ʾ��
                .TextMatrix(i, COLM_������) = rsTmp!������ & ""
                .TextMatrix(i, COLM_��ע) = rsTmp!��ע & ""
                .TextMatrix(i, COLM_��Ȩ��) = rsTmp!��Ȩ�� & ""
                .TextMatrix(i, COLM_��Ȩʱ��) = Format(rsTmp!��Ȩʱ�� & "", "yyyy-mm-dd hh:mm")
                .TextMatrix(i, COLM_���ʿ�ʼʱ��) = Format(rsTmp!���ʿ�ʼʱ�� & "", "yyyy-mm-dd hh:mm")
                .TextMatrix(i, COLM_���ʽ���ʱ��) = Format(rsTmp!���ʽ���ʱ�� & "", "yyyy-mm-dd hh:mm")
                .TextMatrix(i, COLM_������) = rsTmp!������ & ""
                .TextMatrix(i, COLM_����ʱ��) = Format(rsTmp!����ʱ�� & "", "yyyy-mm-dd hh:mm")
                
                If rsTmp!����ʱ�� & "" <> "" Then
                    .Cell(flexcpForeColor, i, 0, i, .Cols - 1) = &H808080
                    Set .Cell(flexcpPicture, i, 0) = imgFilter(4).Picture
                Else
                    Set .Cell(flexcpPicture, i, 0) = imgFilter(1).Picture
                End If
                
                

                If Val(rsTmp!ID & "") = lng��ȨID Then
                    .Row = i
                End If
                rsTmp.MoveNext
             Next
             .ColHidden(COLM_������) = chk������.Value = 0
             .ColHidden(COLM_����ʱ��) = chk������.Value = 0
             .Redraw = flexRDDirect
             stbThis.Panels(2).Text = "��ǰ���˲��ҵ� " & rsTmp.RecordCount & " ����Ȩ��Ϣ"
        Else
            .Rows = .FixedRows + 1
            stbThis.Panels(2).Text = "��ǰ����û�в��ҵ���Ȩ��Ϣ"
        End If
        
        If .Row <= 0 Then .Row = .Rows - 1

        .WordWrap = True
        '�Զ������и�
        .AutoSize COLM_��ע, COLM_������
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub



Public Function GetRs��������(rsTmp As ADODB.Recordset) As Boolean
    Dim str����IDs As String
    Dim arrTmp As Variant
    Dim colPati As Collection
    Dim i As Long, j As Long
    Dim str���� As String, colValue As Collection
    
    If rsTmp Is Nothing Then Exit Function
    If rsTmp.EOF Then Exit Function
    
    
    '���ز�����Ϣ
    str����IDs = ""
    If Not rsTmp Is Nothing Then
        For i = 1 To rsTmp.RecordCount
             If rsTmp!����ids & "" <> "" Then
                arrTmp = Split(rsTmp!����ids & "", ",")
                For j = LBound(arrTmp) To UBound(arrTmp)
                    If InStr("," & str����IDs & ",", "," & Val(arrTmp(j)) & ",") = 0 Then
                       str����IDs = str����IDs & "," & Val(arrTmp(j))
                    End If
                Next
             End If
             rsTmp.MoveNext
        Next
        If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
    End If

    
    If str����IDs <> "" Then
        str����IDs = Mid(str����IDs, 2)
        Set colPati = PatiSvrGetpatiinfo(1, 0, 1240, 0, 2, "", "", "", "", str����IDs)
        
        If Not colPati Is Nothing Then
            Set rsTmp = zldatabase.CopyNewRec(rsTmp)
            Do While Not rsTmp.EOF
               If rsTmp!����ids & "" <> "" Then
                    arrTmp = Split(rsTmp!����ids & "", ",")
                    str���� = ""
                    For j = LBound(arrTmp) To UBound(arrTmp)
                        If Val(arrTmp(j)) <> 0 Then
                            Set colValue = GetColObj(colPati, "_" & arrTmp(j))
                            If Not colValue Is Nothing Then
                                If GetColVal(colValue, "_pati_name") <> "" Then
                                    str���� = str���� & "," & GetColVal(colValue, "_pati_name")
                                End If
                            End If
                        End If
                    Next
                End If
                
                If str���� <> "" Then
                    str���� = Mid(str����, 2)
                    rsTmp!�������� = str����
                End If
                
                rsTmp.MoveNext
            Loop
            rsTmp.MoveFirst
        End If
    End If
End Function


Private Sub LoadList(Optional lng����id As Long)
    Dim strSQL As String
    Dim strFilter As String
    Dim i As Long
    Dim rsTmp As ADODB.Recordset
    Dim curDate As Date
    
    For i = 0 To 3
        If chkFilter(i).Value = 1 Then strFilter = strFilter & "," & i
    Next
    strFilter = Mid(strFilter, 2)
    
    On Error GoTo errH
    If cboTime.ListIndex <> 5 Then
        curDate = zldatabase.Currentdate
        dtpTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm:ss")
    End If
    
    strSQL = "Select a.Id, a.���ʿ�ʼʱ��, a.���ʽ���ʱ��, a.����ʱ��, a.����ԭ��, a.����״̬, a.������, a.����ʱ��,A.����ʱ��,A.������," & vbNewLine & _
                "       f_List2str(Cast(Collect(b.����id || '') As t_Strlist)) As ����ids,null as ��������" & vbNewLine & _
                "From ���Ӳ����������� A, ���Ӳ���������ʲ��� B" & vbNewLine & _
                "Where a.Id = b.����id And a.����ʱ�� Between [1] And [2]" & vbNewLine & _
                " And Instr([3], a.����״̬) > 0 and A.����ʱ�� is null" & vbNewLine & _
                "Group By a.Id, a.���ʿ�ʼʱ��, a.���ʽ���ʱ��, a.����ʱ��, a.����ԭ��, a.����״̬, a.������, a.����ʱ��,A.����ʱ��,A.������" & vbNewLine & _
                "Order by a.����״̬,A.id"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(Format(dtpTime(0).Value, "yyyy-MM-dd hh:mm")), CDate(Format(IIf(cboTime.ListIndex <> 5, dtpTime(1).Value + 1, dtpTime(1).Value), "yyyy-MM-dd hh:mm")), strFilter)
    Call GetRs��������(rsTmp)
    With vsList
        .Rows = .FixedRows
        If Not rsTmp.EOF Then
             .Redraw = flexRDNone
             .Rows = .FixedRows + rsTmp.RecordCount
             For i = 1 To rsTmp.RecordCount
                '������
                .TextMatrix(i, COL_����ID) = Val(rsTmp!ID & "")
                .TextMatrix(i, COL_����ʱ��) = Val(rsTmp!����ʱ�� & "")
                .TextMatrix(i, COL_����ʱ��) = Format(rsTmp!����ʱ�� & "", "yyyy-mm-dd hh:mm")
                .TextMatrix(i, COL_������) = rsTmp!������ & ""
                '��ʾ��
                .TextMatrix(i, COL_������) = rsTmp!������ & ""
                .TextMatrix(i, COL_����ʱ��) = Format(rsTmp!����ʱ�� & "", "yyyy-mm-dd hh:mm")
                .TextMatrix(i, COL_������ʲ���) = rsTmp!�������� & ""
                .TextMatrix(i, COL_���ʿ�ʼʱ��) = Format(rsTmp!���ʿ�ʼʱ�� & "", "yyyy-mm-dd hh:mm")
                .TextMatrix(i, COL_���ʽ���ʱ��) = Format(rsTmp!���ʽ���ʱ�� & "", "yyyy-mm-dd hh:mm")
                .TextMatrix(i, COL_����ԭ��) = rsTmp!����ԭ�� & ""
                
                .TextMatrix(i, COL_����״̬) = Decode(Val(rsTmp!����״̬ & ""), 0, "������", 1, "������", 2, "������", 3, "�Ѿܾ�")
                Set .Cell(flexcpPicture, i, 0) = imgFilter(Val(rsTmp!����״̬ & "")).Picture

                If Val(rsTmp!ID & "") = lng����id Then
                    .Row = i
                End If
                rsTmp.MoveNext
             Next
             .Redraw = flexRDDirect
             stbThis.Panels(2).Text = "��ǰ���˲��ҵ� " & rsTmp.RecordCount & " ��������Ϣ"
        Else
            .Rows = .FixedRows + 1
            stbThis.Panels(2).Text = "��ǰ����û�в��ҵ�������Ϣ"
        End If
        
        If .Row <= 0 Then .Row = .Rows - 1
        .WordWrap = True
        '�Զ������и�
        .AutoSize COL_������ʲ���, COL_����ԭ��
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function Get������(lngRow As Long) As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    If lngRow <= 0 Then Exit Function
    If Val(vsManage.TextMatrix(lngRow, COLM_ID)) = 0 Then Exit Function
    
    If vsManage.TextMatrix(lngRow, COLM_������) <> "" Then Get������ = vsManage.TextMatrix(lngRow, COLM_������): Exit Function
    
    strSQL = "Select f_List2str(Cast(Collect(b.���� || '') As t_Strlist)) As ��Ȩ��Ա" & vbNewLine & _
                "From ���Ӳ�����Ȩ������Ա A, ��Ա�� B" & vbNewLine & _
                "Where a.��Աid = b.Id And a.��Ȩid =[1]"
    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsManage.TextMatrix(lngRow, COLM_ID)))
    If Not rsTmp Is Nothing Then
        If Not rsTmp.EOF Then
            Get������ = Replace(rsTmp!��Ȩ��Ա & "", ",", "��")
            vsManage.TextMatrix(lngRow, COLM_������) = Replace(rsTmp!��Ȩ��Ա & "", ",", "��")
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Get���ʷ�Χ(lngRow As Long) As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strOut As String
    On Error GoTo errH
    If lngRow <= 0 Then Exit Function
    If Val(vsManage.TextMatrix(lngRow, COLM_ID)) = 0 Then Exit Function
    If vsManage.TextMatrix(lngRow, COLM_���˷�Χ����) <> "" Then Get���ʷ�Χ = vsManage.TextMatrix(lngRow, COLM_���˷�Χ����): Exit Function
    
    Select Case Val(vsManage.TextMatrix(lngRow, COLM_���ʲ���))
        Case 0 'ȫԺ����
            strOut = "������Աӵ�в鿴ȫԺ���˵��Ӳ�����Ȩ��"
        Case 1 '���Ʋ���
            strOut = "������Աӵ�в鿴���ڿ��ҵĲ��˵��Ӳ���Ȩ��"
        Case 2 'ָ�����Ҳ���
            strSQL = "Select f_List2str(Cast(Collect(b.���� || '') As t_Strlist)) As ���ʿ���" & vbNewLine & _
                        "From ���Ӳ�����Ȩ���ʲ��� A, ���ű� B" & vbNewLine & _
                        "Where a.��Ȩ���� = b.Id And a.��Ȩid = [1]"
            Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsManage.TextMatrix(lngRow, COLM_ID)))
            If Not rsTmp Is Nothing Then
                If Not rsTmp.EOF Then
                    strOut = Replace(rsTmp!���ʿ��� & "", ",", "��")
                End If
            End If
            strOut = "���ʿ��ң�" & strOut
        Case 3 'ָ������
            strSQL = "Select f_List2str(Cast(Collect(a.��Ȩ���� || '') As t_Strlist)) As ����ids,null as ��������" & vbNewLine & _
                        "From ���Ӳ�����Ȩ���ʲ��� A" & vbNewLine & _
                        "Where a.��Ȩid = [1]"
            Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsManage.TextMatrix(lngRow, COLM_ID)))
            Call GetRs��������(rsTmp)
            
            If Not rsTmp Is Nothing Then
                If Not rsTmp.EOF Then
                    strOut = Replace(rsTmp!�������� & "", ",", "��")
                End If
            End If
            strOut = "���ʲ��ˣ�" & strOut
    End Select
    
    vsManage.TextMatrix(lngRow, COLM_���˷�Χ����) = strOut
    Get���ʷ�Χ = strOut
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdLogFind_Click()
    Call LoadLog
End Sub

Private Sub cmdManageFind_Click()
    Call LoadManage
End Sub

Private Sub tbcSub_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    stbThis.Panels(2).Text = "�鿴" & tbcSub.Selected.Tag
End Sub

Private Sub vsList_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow <= 0 Or NewCol < 0 Then Exit Sub
    If vsList.Col >= vsList.FixedCols Then
        vsList.ForeColorSel = vsList.Cell(flexcpForeColor, NewRow, NewCol)
    End If
    With vsInfo
        If Val(vsList.TextMatrix(NewRow, COL_����ID)) <> 0 Then
            '���ʲ���
            .TextMatrix(Row_���ʲ���, 0) = vsList.TextMatrix(NewRow, COL_������ʲ���) & ""
            
            '����ʱ��
            .TextMatrix(Row_����ʱ��, 0) = "�� " & Format(vsList.TextMatrix(NewRow, COL_���ʿ�ʼʱ��), "yyyy-mm-dd hh:mm") & vbCrLf & "�� " & _
                                        Format(vsList.TextMatrix(NewRow, COL_���ʽ���ʱ��), "yyyy-mm-dd hh:mm") & "�ڼ�" & vbCrLf & "���ʲ���" & Decode(Val(vsList.TextMatrix(NewRow, COL_����ʱ��)), 0, "���в�������", 1, "δ�鵵�Ĳ���", "�ѹ鵵�Ĳ���")
                             
            '��������
            .TextMatrix(Row_��������, 0) = GetXmlInfo(1, NewRow)
        Else
            .TextMatrix(Row_���ʲ���, 0) = ""
            .TextMatrix(Row_����ʱ��, 0) = ""
            .TextMatrix(Row_��������, 0) = ""
        End If
        .WordWrap = True
        '�Զ������и�
        .AutoSize 0
    End With
End Sub

Private Function GetXmlInfo(ByVal intType As Integer, ByVal lngRow As Long) As String
    '��ȡ�������ݵ�Xml������
    'lngType =0 ��Ȩ���� =1 ��������
    Dim objXML As New zl9ComLib.clsXML
    Dim strXML As String
    Dim strErr As String
    Dim strValue As String
    Dim strOut As String
    Dim strTmp As String
    
    On Error GoTo errH
    If lngRow <= 0 Then Exit Function

    If intType = 0 Then
        If Val(vsManage.TextMatrix(lngRow, COLM_ID)) = 0 Then Exit Function
        '��ȡ����
        If vsManage.TextMatrix(lngRow, COLM_��������) <> "" Then GetXmlInfo = vsManage.TextMatrix(lngRow, COLM_��������): Exit Function

        strXML = Sys.ReadXML("���Ӳ���������Ȩ", "��������", "ID=[1]", strErr, Val(vsManage.TextMatrix(lngRow, COLM_ID)))
    Else
        If Val(vsList.TextMatrix(lngRow, COL_����ID)) = 0 Then Exit Function
        '��ȡ����
        If vsList.TextMatrix(lngRow, COL_��������) <> "" Then GetXmlInfo = vsList.TextMatrix(lngRow, COL_��������): Exit Function

        strXML = Sys.ReadXML("���Ӳ�����������", "��������", "ID=[1]", strErr, Val(vsList.TextMatrix(lngRow, COL_����ID)))
    End If

    If Err.Number = 0 And strErr <> "" Then
        MsgBox strErr, vbInformation, gstrSysName
        Exit Function
    End If
    
    If objXML.OpenXMLDocument(strXML) = False Then Exit Function

    '��������
    strValue = "": Call objXML.GetSingleNodeValue("all_files", strValue, xsNumber)
    If Val(strValue) = 1 Then
        strOut = "�����Ʒ�����������"
    Else
        '������ҳ��ҽ�����ٴ�·��
        strValue = "": Call objXML.GetSingleNodeValue("medical_record", strValue, xsNumber): If Val(strValue) = 1 Then strOut = "������ҳ��" & vbCrLf & vbCrLf
        strValue = "": Call objXML.GetSingleNodeValue("advice", strValue, xsNumber): If Val(strValue) = 1 Then strOut = strOut & "����ҽ����" & vbCrLf & vbCrLf
        strValue = "": Call objXML.GetSingleNodeValue("cispath", strValue, xsNumber): If Val(strValue) = 1 Then strOut = strOut & "�ٴ�·����" & vbCrLf & vbCrLf
        strValue = "": Call objXML.GetSingleNodeValue("patipeis", strValue, xsNumber): If Val(strValue) = 1 Then strOut = strOut & "��챨�桢" & vbCrLf & vbCrLf
        
        '�����¼
        strValue = "": Call objXML.GetSingleNodeValue("nursing_record", strValue, xsNumber)
        If Val(strValue) = 1 Then
            strValue = "": Call objXML.GetSingleNodeValue("nursing_info/nursing_all", strValue, xsNumber)
            If Val(strValue) = 1 Then
                strOut = strOut & "�����¼(���л����¼)" & vbCrLf & vbCrLf
            Else
                strValue = "": Call objXML.GetSingleNodeValue("nursing_info/thermometer", strValue, xsNumber): If Val(strValue) = 1 Then strTmp = "���µ���"
                strValue = "": Call objXML.GetSingleNodeValue("nursing_info/record_file", strValue, xsNumber)
                If Val(strValue) = 1 Then
                    Call GetXmlString(objXML, "nursing_info/file_name", strValue)
                    strValue = Replace(strValue, ",", "��")
                    strTmp = strTmp & strValue
                Else
                    strTmp = Replace(strTmp, "��", "")
                End If
                strOut = strOut & "�����¼" & vbCrLf & "(��¼��Χ��" & strTmp & ")" & vbCrLf & vbCrLf
            End If
        End If
        
        '��鱨��
        strValue = "": Call objXML.GetSingleNodeValue("pacs_report", strValue, xsNumber)
        If Val(strValue) = 1 Then
            strValue = "": Call objXML.GetSingleNodeValue("pacs_info/pacs_type", strValue, xsNumber)
            'pacs_type =0���м�鱨�� =1ָ�����͵ļ�鱨��
            If Val(strValue) = 0 Then
                strOut = strOut & "��鱨��(���м�鱨��)" & vbCrLf & vbCrLf
            Else
                Call GetXmlString(objXML, "pacs_info/pacs_report_type/type_name", strValue)
                strValue = Replace(strValue, ",", "��")
                strOut = strOut & "��鱨��" & vbCrLf & "(���ͷ�Χ��" & strValue & ")" & vbCrLf & vbCrLf
            End If
        End If
        
        '���鱨��
        strValue = "": Call objXML.GetSingleNodeValue("lis_report", strValue, xsNumber)
        If Val(strValue) = 1 Then
            strValue = "": Call objXML.GetSingleNodeValue("lis_info/lis_type", strValue, xsNumber)
            'lis_type =0 ���м��鱨�� =1ָ�����͵ļ��鱨��
            If Val(strValue) = 0 Then
                strOut = strOut & "���鱨��(���м��鱨��)" & vbCrLf & vbCrLf
            Else
                Call GetXmlString(objXML, "lis_info/lis_report_type/type_name", strValue)
                strValue = Replace(strValue, ",", "��")
                strOut = strOut & "���鱨��" & vbCrLf & "(���ͷ�Χ��" & strValue & ")" & vbCrLf & vbCrLf
            End If
        End If
        
        '���Ӳ���
        strValue = "": Call objXML.GetSingleNodeValue("emr", strValue, xsNumber)
        If Val(strValue) = 1 Then
            strValue = "": Call objXML.GetSingleNodeValue("emr_info/emr_type", strValue, xsNumber)
            'emr_type =0 ���е��Ӳ���  =1ָ�����͵ĵ��Ӳ���  =1ָ������ĵ��Ӳ���
            If Val(strValue) = 0 Then
                strOut = strOut & "���Ӳ���(���е��Ӳ���)" & vbCrLf & vbCrLf
            ElseIf Val(strValue) = 1 Then
                Call GetXmlString(objXML, "emr_info/standard_class/class_name", strValue)
                strValue = Replace(strValue, ",", "��")
                strOut = strOut & "���Ӳ���" & vbCrLf & "(�������ͷ�Χ��" & strValue & ")" & vbCrLf & vbCrLf
            Else
                Call GetXmlString(objXML, "emr_info/antetype_class/class_name", strValue)
                strValue = Replace(strValue, ",", "��")
                strOut = strOut & "���Ӳ���" & vbCrLf & "(������Χ��" & strValue & ")" & vbCrLf & vbCrLf
            End If
        End If
    End If
    
    If Right(strOut, 5) = "��" & vbCrLf & vbCrLf Then strOut = Left(strOut, Len(strOut) - 5)
    '������������
    If intType = 0 Then
        vsManage.TextMatrix(lngRow, COLM_��������) = strOut
    Else
        vsList.TextMatrix(lngRow, COL_��������) = strOut
    End If
    GetXmlInfo = strOut
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Function GetXmlString(objXML As Object, ByVal strNode As String, ByRef strValue As String) As Boolean
    Dim rsTmp As ADODB.Recordset
    On Error GoTo errH
    strValue = ""
    If objXML.GetMultiNodeRecord(strNode, rsTmp) Then
        If Not rsTmp Is Nothing Then
            Do While Not rsTmp.EOF
                strValue = strValue & "," & rsTmp!node_value
                rsTmp.MoveNext
            Loop
            strValue = Mid(strValue, 2)
        End If
    End If
    GetXmlString = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub Form_Load()
    'CommandBars
    '-----------------------------------------------------
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        '.UseFadedIcons = True '����VisualTheme����Ч
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    Set cbsMain.Icons = zlCommFun.GetPubIcons
    Call MainDefCommandBar
    'tabControl
    '-----------------------------------------------------
    With Me.tbcSub
        With .PaintManager
            .Appearance = xtpTabAppearancePropertyPage2003
            .ClientFrame = xtpTabFrameSingleLine
            .BoldSelected = True
            .OneNoteColors = True
            .ShowIcons = True
        End With
        '���Ӵ���ʱ��Form_Load�����Զ�ѡ�е�һ������Ŀ�Ƭ
        '������õ�ǰ��Ƭ����,�򲻻��Զ��л�ѡ��,����ʾ����δ��
        '����ָ����������Ч�����ձ�Ϊ0-N��ֻ�ǿ��ܸı����˳��
        .InsertItem(0, "��Ȩ��¼", picManage.hwnd, 0).Tag = "��Ȩ��¼"
        .InsertItem(1, "������¼", picApply.hwnd, 0).Tag = "������¼"
        .InsertItem(2, "������־", picLog.hwnd, 0).Tag = "������־"
        
        .Item(2).Selected = True
        .Item(1).Selected = True
        .Item(0).Selected = True
    End With
    
    Call InitListTable
    Call InitManageTable
    Call InitLogTable
    
    '��ʼ��������
    With vsInfo
        '���ʲ���
        .TextMatrix(Row_���ʲ��˱���, 0) = "���ʲ��ˣ�"
        .Cell(flexcpForeColor, Row_���ʲ��˱���, 0) = &H80000002
        Set .Cell(flexcpPicture, Row_���ʲ��˱���, 0) = img16.ListImages("���ʲ���").Picture
        .Cell(flexcpFontBold, Row_���ʲ��˱���, 0) = True
        
        '����ʱ��
        .TextMatrix(Row_����ʱ�ޱ���, 0) = "����ʱ�ޣ�"
        .Cell(flexcpForeColor, Row_����ʱ�ޱ���, 0) = &H80000002
        Set .Cell(flexcpPicture, Row_����ʱ�ޱ���, 0) = img16.ListImages("����ʱ��").Picture
        .Cell(flexcpFontBold, Row_����ʱ�ޱ���, 0) = True

        '��������
        .TextMatrix(Row_�������ݱ���, 0) = "�������ݣ�"
        .Cell(flexcpForeColor, Row_�������ݱ���, 0) = &H80000002
        Set .Cell(flexcpPicture, Row_�������ݱ���, 0) = img16.ListImages("��������").Picture
        .Cell(flexcpFontBold, Row_�������ݱ���, 0) = True

        .WordWrap = True
        '�Զ������и�
        .AutoSize 0
    End With

    '��ʼ��������
    With vsManageInfo
        '������
        .TextMatrix(RowM_�����߱���, 0) = "�����ߣ�"
        .Cell(flexcpForeColor, RowM_�����߱���, 0) = &H80000002
        Set .Cell(flexcpPicture, RowM_�����߱���, 0) = img16.ListImages("����ҽ��").Picture
        .Cell(flexcpFontBold, RowM_�����߱���, 0) = True
    
        '���ʲ���
        .TextMatrix(RowM_���ʲ��˱���, 0) = "���ʷ�Χ��"
        .Cell(flexcpForeColor, RowM_���ʲ��˱���, 0) = &H80000002
        Set .Cell(flexcpPicture, RowM_���ʲ��˱���, 0) = img16.ListImages("���ʲ���").Picture
        .Cell(flexcpFontBold, RowM_���ʲ��˱���, 0) = True
        
        '����ʱ��
        .TextMatrix(RowM_����ʱ�ޱ���, 0) = "����ʱ�ޣ�"
        .Cell(flexcpForeColor, RowM_����ʱ�ޱ���, 0) = &H80000002
        Set .Cell(flexcpPicture, RowM_����ʱ�ޱ���, 0) = img16.ListImages("����ʱ��").Picture
        .Cell(flexcpFontBold, RowM_����ʱ�ޱ���, 0) = True

        '��������
        .TextMatrix(RowM_�������ݱ���, 0) = "�������ݣ�"
        .Cell(flexcpForeColor, RowM_�������ݱ���, 0) = &H80000002
        Set .Cell(flexcpPicture, RowM_�������ݱ���, 0) = img16.ListImages("��������").Picture
        .Cell(flexcpFontBold, RowM_�������ݱ���, 0) = True

        .WordWrap = True
        '�Զ������и�
        .AutoSize 0
    End With
    
    '---cboTime
    cboTime.AddItem "��    ��"
    cboTime.AddItem "�������"
    cboTime.AddItem "�������"
    cboTime.AddItem "���һ��"
    cboTime.AddItem "���һ��"
    cboTime.AddItem "[ָ  ��]"
    cboTime.ListIndex = 3
    
    '---cboManageTime
    cboManageTime.AddItem "��    ��"
    cboManageTime.AddItem "�������"
    cboManageTime.AddItem "�������"
    cboManageTime.AddItem "���һ��"
    cboManageTime.AddItem "���һ��"
    cboManageTime.AddItem "[ָ  ��]"
    cboManageTime.ListIndex = 3
    
    '---cboLogTime
    cboLogTime.AddItem "��    ��"
    cboLogTime.AddItem "�������"
    cboLogTime.AddItem "�������"
    cboLogTime.AddItem "���һ��"
    cboLogTime.AddItem "���һ��"
    cboLogTime.AddItem "[ָ  ��]"
    cboLogTime.ListIndex = 3
    
    Call RestoreWinState(Me, App.ProductName, , True)
    Call LoadList
    Call LoadManage
    Me.Tag = "1"
End Sub
'


Private Sub MainDefCommandBar()
'���ܣ������ڲ˵����岿��
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim lngCount As Long
    
    '�˵�����
    '-----------------------------------------------------
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_FilePopup, "�ļ�(&F)", -1, False)
    objMenu.ID = conMenu_FilePopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�(&X)")
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�༭(&E)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyAdd, "������Ȩ(&A)")
            objControl.IconId = 3001
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyEdit, "������Ȩ(&E)")
            objControl.IconId = 3003
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Complete, "��������(&O)")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Undone, "�ܾ�����(&N)")
            objControl.IconId = 4114
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "ȡ���ܾ�(&U)")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "������Ȩ(&Q)")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "������&Excel")
            objControl.IconId = 30134
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_ViewPopup, "�鿴(&V)", -1, False)
    objMenu.ID = conMenu_ViewPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��(&R)")
    End With
    
    Set objMenu = cbsMain.ActiveMenuBar.Controls.Add(xtpControlPopup, conMenu_HelpPopup, "����(&H)", -1, False)
    objMenu.ID = conMenu_HelpPopup
    With objMenu.CommandBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "��������(&H)")
        Set objPopup = .Add(xtpControlButtonPopup, conMenu_Help_Web, "&WEB�ϵ�")
        With objPopup.CommandBar.Controls
            .Add xtpControlButton, conMenu_Help_Web_Home, "��ҳ(&H)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Forum, "��̳(&F)", -1, False
            .Add xtpControlButton, conMenu_Help_Web_Mail, "���ͷ���(&M)", -1, False
        End With
        Set objControl = .Add(xtpControlButton, conMenu_Help_About, "����(&A)��")
            objControl.BeginGroup = True
    End With

    '����������:������������
    '-----------------------------------------------------
    Set objBar = cbsMain.Add("������", xtpBarTop)
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyAdd, "������Ȩ")
            objControl.IconId = 3001
        Set objControl = .Add(xtpControlButton, conMenu_Edit_ApplyEdit, "������Ȩ")
            objControl.IconId = 3003
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Complete, "��������")
        Set objControl = .Add(xtpControlButton, conMenu_Manage_Undone, "�ܾ�����")
            objControl.IconId = 4114
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Untread, "ȡ���ܾ�")
        Set objControl = .Add(xtpControlButton, conMenu_Edit_Delete, "������Ȩ")
        Set objControl = .Add(xtpControlButton, conMenu_File_Excel, "������&Excel")
            objControl.IconId = 3013
        Set objControl = .Add(xtpControlButton, conMenu_View_Refresh, "ˢ��")
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "����")
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�")
    End With

    '����һЩ�������ȼ���
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add 0, vbKeyF5, conMenu_View_Refresh 'ˢ��
        .Add 0, vbKeyF1, conMenu_Help_Help '����
    End With

    '�ָ����̶���һЩ�˵�����
    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.SetIconSize 16, 16
    For lngCount = 2 To cbsMain.Count
        cbsMain(lngCount).ContextMenuPresent = False
        cbsMain(lngCount).ShowTextBelowIcons = False
        cbsMain(lngCount).EnableDocking xtpFlagHideWrap Or xtpFlagStretched
        For Each objControl In cbsMain(lngCount).Controls
            objControl.Style = xtpButtonIconAndCaption
        Next
    Next
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Call cbsMain_Resize
End Sub

Private Sub cbsMain_Resize()
    Dim lngLeft As Long, lngTop  As Long, lngRight  As Long, lngBottom  As Long
    
    Call Me.cbsMain.GetClientRect(lngLeft, lngTop, lngRight, lngBottom)
    
    On Error Resume Next
    With Me.tbcSub
        .Left = lngLeft: .Top = lngTop
        .Width = lngRight - lngLeft
        .Height = Me.Height - stbThis.Height - 1500
    End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SaveWinState(Me, App.ProductName)
End Sub



Private Sub picApply_Resize()
    On Error Resume Next
    '�̶���ϸ��Ϣ4000����
    picInfo.Width = 5000

    fraFillter.Top = 100: fraFillter.Left = 30
    fraFillter.Width = picApply.Width - 60
    
    vsList.Top = fraFillter.Top + fraFillter.Height + 150: vsList.Height = picApply.Height - fraFillter.Height - 260

    
    vsList.Left = fraFillter.Left
    vsList.Width = fraFillter.Width - 5000 - 30
    
    picInfo.Top = vsList.Top - 70: picInfo.Left = vsList.Left + vsList.Width + 50
    picInfo.Height = vsList.Height + 70
    vsInfo.Height = picInfo.Height - 300
End Sub


Private Sub picManage_Resize()
    On Error Resume Next
    '�̶���ϸ��Ϣ4000����
    fraManageInfo.Width = 5000

    fraManageFilter.Top = 100: fraManageFilter.Left = 30
    fraManageFilter.Width = picManage.Width - 60
    
    vsManage.Top = fraManageFilter.Top + fraManageFilter.Height + 150: vsManage.Height = picManage.Height - fraManageFilter.Height - 260

    
    vsManage.Left = fraManageFilter.Left
    vsManage.Width = fraManageFilter.Width - 5000 - 30
    
    fraManageInfo.Top = vsManage.Top - 70: fraManageInfo.Left = vsManage.Left + vsManage.Width + 50
    fraManageInfo.Height = vsManage.Height + 70
    vsManageInfo.Height = fraManageInfo.Height - 300
End Sub


Private Sub picLog_Resize()
    On Error Resume Next
    fraLog.Top = 100: fraLog.Left = 30
    fraLog.Width = picLog.Width - 60
    
    vsLog.Top = fraLog.Top + fraLog.Height + 150: vsLog.Height = picLog.Height - fraLog.Height - 260
    vsLog.Left = fraLog.Left
    vsLog.Width = fraLog.Width
End Sub


Private Sub InitListTable()
'���ܣ���ʼ���б��嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long, lngWidth As Long
              
    strHead = "����id;��������;����ʱ��;����ʱ��;������;" & _
                "����ʱ��,2000,1;������,800,4;������ʲ���,3200,1;���ʿ�ʼʱ��,2000,1;���ʽ���ʱ��,2000,1;����ԭ��,3800,1;����״̬,1050,4"
    arrHead = Split(strHead, ";")
    With vsList
        .Clear
        .FixedRows = 1
        .FixedCols = 1
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        .SelectionMode = flexSelectionByRow
        .FocusRect = flexFocusNone
        .HighLight = flexHighlightAlways
'        .BackColorSel = &HFAEADA


        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            .FixedAlignment(.FixedCols + i) = 4
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                lngWidth = Val(Split(arrHead(i), ",")(1))
                .ColWidth(.FixedCols + i) = lngWidth
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                'Ϊ��֧��zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0    'Ϊ��֧��zl9PrintMode
            End If
            .colData(.FixedCols + i) = .ColWidth(.FixedCols + i)    '��¼ԭʼ�п�������ѡ����
        Next
        .Editable = flexEDNone
    End With
End Sub


Private Sub InitManageTable()
'���ܣ���ʼ����Ȩ�б��嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long, lngWidth As Long
              
    strHead = "ID;��������;����ʱ��;��Ȩ����;���ʲ���;���˷�Χ����;" & _
                "������,2500,1;��ע,4000,1;���ʿ�ʼʱ��,1700,1;���ʽ���ʱ��,1700,1;��Ȩ��,1050,4;��Ȩʱ��,1700,1;������,1050,4;����ʱ��,1700,1;������"
    arrHead = Split(strHead, ";")
    With vsManage
        .Clear
        .FixedRows = 1
        .FixedCols = 1
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        .SelectionMode = flexSelectionByRow
        .FocusRect = flexFocusNone
        .HighLight = flexHighlightAlways
'        .BackColorSel = &HFAEADA


        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            .FixedAlignment(.FixedCols + i) = 4
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                lngWidth = Val(Split(arrHead(i), ",")(1))
                .ColWidth(.FixedCols + i) = lngWidth
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                'Ϊ��֧��zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0    'Ϊ��֧��zl9PrintMode
            End If
            .colData(.FixedCols + i) = .ColWidth(.FixedCols + i)    '��¼ԭʼ�п�������ѡ����
        Next
        .Editable = flexEDNone
    End With
End Sub

Private Sub vsManage_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If NewRow <= 0 Or NewCol < 0 Then Exit Sub
    If vsManage.Col >= vsManage.FixedCols Then
        vsManage.ForeColorSel = vsManage.Cell(flexcpForeColor, NewRow, NewCol)
    End If
    With vsManageInfo
        If Val(vsManage.TextMatrix(NewRow, COLM_ID)) <> 0 Then
            .TextMatrix(RowM_������, 0) = Get������(NewRow)
        
            '���ʲ���
            .TextMatrix(RowM_���ʲ���, 0) = Get���ʷ�Χ(NewRow)
            
            '����ʱ��
            .TextMatrix(RowM_����ʱ��, 0) = "�� " & Format(vsManage.TextMatrix(NewRow, COLM_���ʿ�ʼʱ��), "yyyy-mm-dd hh:mm") & vbCrLf & "�� " & _
                                        Format(vsManage.TextMatrix(NewRow, COLM_���ʽ���ʱ��), "yyyy-mm-dd hh:mm") & "�ڼ�" & vbCrLf & "���ʲ���" & Decode(Val(vsManage.TextMatrix(NewRow, COLM_����ʱ��)), 0, "���в�������", 1, "δ�鵵�Ĳ���", "�ѹ鵵�Ĳ���")
                                      
            '��������
            .TextMatrix(RowM_��������, 0) = GetXmlInfo(0, NewRow)
        Else
            .TextMatrix(RowM_������, 0) = ""
            .TextMatrix(RowM_���ʲ���, 0) = ""
            .TextMatrix(RowM_����ʱ��, 0) = ""
            .TextMatrix(RowM_��������, 0) = ""
        End If
        .WordWrap = True
        '�Զ������и�
        .AutoSize 0
    End With
End Sub


Private Sub InitLogTable()
'���ܣ���ʼ���б��嵥��ʽ
    Dim arrHead As Variant, strHead As String, i As Long, lngWidth As Long
              
    strHead = "ID;����ID;����ID;������Դ;����ID;" & _
                "����ʱ��,2000,1;������,1500,4;���ʲ���,1400,4;�Ա�,700,4;����,700,4;��ʶ��,950,4;����,1700,4;��������,4000,1;��������,5000,1"
    arrHead = Split(strHead, ";")
    With vsLog
        .Clear
        .FixedRows = 1
        .FixedCols = 1
        .Cols = .FixedCols + UBound(arrHead) + 1
        .Rows = .FixedRows + 1
        .SelectionMode = flexSelectionByRow
        .FocusRect = flexFocusNone
        .HighLight = flexHighlightAlways
'        .BackColorSel = &HFAEADA


        For i = 0 To UBound(arrHead)
            .TextMatrix(.FixedRows - 1, .FixedCols + i) = Split(arrHead(i), ",")(0)
            .FixedAlignment(.FixedCols + i) = 4
            If UBound(Split(arrHead(i), ",")) > 0 Then
                .ColHidden(.FixedCols + i) = False
                lngWidth = Val(Split(arrHead(i), ",")(1))
                .ColWidth(.FixedCols + i) = lngWidth
                .ColAlignment(.FixedCols + i) = Val(Split(arrHead(i), ",")(2))
                'Ϊ��֧��zl9PrintMode
                .Cell(flexcpAlignment, .FixedRows, .FixedCols + i, .Rows - 1, .FixedCols + i) = Val(Split(arrHead(i), ",")(2))
            Else
                .ColHidden(.FixedCols + i) = True
                .ColWidth(.FixedCols + i) = 0    'Ϊ��֧��zl9PrintMode
            End If
            .colData(.FixedCols + i) = .ColWidth(.FixedCols + i)    '��¼ԭʼ�п�������ѡ����
        Next
        .Editable = flexEDNone
    End With
End Sub


Public Sub zlRptPrint(ByVal bytMode As Byte)
    '����:�����ݸ��Ƶ��ɴ�ӡ�Ķ��󣬵��ô�ӡ
    '����:  bytMode��1-��ӡ;2-Ԥ��;3-�����EXCEL
    If Val(vsLog.TextMatrix(1, COLG_ID)) = 0 Then Exit Sub
    
    '-------------------------------------------------
    '���ô�ӡ��������
    Dim objPrint As New zlPrint1Grd
    Dim objAppRow As zlTabAppRow
    
    Set objPrint.Body = vsLog
    
    objPrint.Title.Text = "���Ӳ������ʼ�¼"
    Set objAppRow = New zlTabAppRow
    Call objAppRow.Add("")
    Call objAppRow.Add("��ӡʱ��:" & Now())
    Call objPrint.BelowAppRows.Add(objAppRow)
    
    If bytMode = 1 Then
        bytMode = zlPrintAsk(objPrint)
        If bytMode <> 0 Then zlPrintOrView1Grd objPrint, bytMode
    Else
        zlPrintOrView1Grd objPrint, bytMode
    End If
End Sub

Private Sub LoadLog()
    Dim strSQL As String
    Dim i As Long
    Dim rsTmp As ADODB.Recordset
    Dim curDate As Date

    On Error GoTo errH
    If cboLogTime.ListIndex <> 5 Then
        curDate = zldatabase.Currentdate
        dtpLogTime(1).Value = Format(curDate, "yyyy-MM-dd hh:mm:ss")
    End If

    strSQL = "Select g.*, f.���� As ��������" & vbNewLine & _
                "From (Select b.Id, b.����id, b.����id, b.������Դ, b.����id, b.����ʱ��, b.������, a.����, a.�Ա�, a.����, a.����� As ��ʶ��, a.ִ�в���id As ����," & vbNewLine & _
                "              a.����ʱ�� As ��ʼʱ��, Null As ����ʱ��, b.��������, -1 As ��������" & vbNewLine & _
                "       From ���˹Һż�¼ A, ���Ӳ���������־ B" & vbNewLine & _
                "       Where a.����id = b.����id And a.Id = b.����id And b.������Դ = 1 And b.����ʱ�� Between [1] And [2]" & vbNewLine & _
                "       Union All" & vbNewLine & _
                "       Select d.Id, d.����id, d.����id, d.������Դ, d.����id, d.����ʱ��, d.������, c.����, c.�Ա�, c.����, c.סԺ�� As ��ʶ��, c.��Ժ����id As ����," & vbNewLine & _
                "              c.��Ժ���� As ��ʼʱ��, c.��Ժ���� As ����ʱ��, d.��������, Nvl(��������, 0) As ��������" & vbNewLine & _
                "       From ������ҳ C, ���Ӳ���������־ D" & vbNewLine & _
                "       Where c.����id = d.����id And c.��ҳid = d.����id And d.������Դ = 2 And d.����ʱ�� Between [1] And [2]) G, ���ű� F" & vbNewLine & _
                "Where g.���� = f.Id" & vbNewLine & _
                "Order By ����ʱ�� Desc"

    Set rsTmp = zldatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(Format(dtpLogTime(0).Value, "yyyy-MM-dd hh:mm")), CDate(Format(IIf(cboLogTime.ListIndex <> 5, dtpLogTime(1).Value + 1, dtpLogTime(1).Value), "yyyy-MM-dd hh:mm")))
    With vsLog
        .Rows = .FixedRows
        If Not rsTmp.EOF Then
             .Redraw = flexRDNone
             .Rows = .FixedRows + rsTmp.RecordCount
             For i = 1 To rsTmp.RecordCount
                '������
                .TextMatrix(i, COLG_ID) = Val(rsTmp!ID & "")
                .TextMatrix(i, COLG_����ID) = Val(rsTmp!����ID & "")
                .TextMatrix(i, COLG_����ID) = Val(rsTmp!����id & "")
                .TextMatrix(i, COLG_������Դ) = Val(rsTmp!������Դ & "")
                .TextMatrix(i, COLG_����ID) = rsTmp!����ID & ""
                 '��ʾ��

                .TextMatrix(i, COLG_����ʱ��) = Format(rsTmp!����ʱ�� & "", "yyyy-mm-dd hh:mm")
                Set .Cell(flexcpPicture, i, COLG_����ʱ��) = img16.ListImages("����ʱ��").Picture
                .TextMatrix(i, COLG_������) = rsTmp!������ & ""
                .TextMatrix(i, COLG_��������) = rsTmp!���� & ""
                Set .Cell(flexcpPicture, i, COLG_��������) = img16.ListImages(IIf(rsTmp!�Ա� & "" = "Ů", "girl", "boy")).Picture
                .TextMatrix(i, COLG_�����Ա�) = rsTmp!�Ա� & ""
                .TextMatrix(i, COLG_��������) = rsTmp!���� & ""
                .TextMatrix(i, COLG_���˱�ʶ��) = rsTmp!��ʶ�� & ""
                .TextMatrix(i, COLG_���˿���) = rsTmp!�������� & ""
                .TextMatrix(i, COLG_��������) = rsTmp!�������� & ""
                .TextMatrix(i, COLG_��������) = IIf(Val(rsTmp!������Դ) = 2, "��" & rsTmp!����id & "��" & IIf(rsTmp!�������� = 1, "��������", IIf(rsTmp!�������� = 2, "סԺ����", "סԺ")), "�������") & " " & Format(rsTmp!��ʼʱ��, "yyyy-MM-dd HH:mm") & _
                    IIf(Not IsNull(rsTmp!����ʱ��), "��" & Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm"), "")
                rsTmp.MoveNext
             Next
             .Redraw = flexRDDirect
             stbThis.Panels(2).Text = "��ǰ���˲��ҵ� " & rsTmp.RecordCount & " ��������Ϣ"
        Else
            .Rows = .FixedRows + 1
            stbThis.Panels(2).Text = "��ǰ����û�в��ҷ�����Ϣ"
        End If

        If .Row <= 0 Then .Row = .Rows - 1

        .WordWrap = True
        '�Զ������и�
        .AutoSize COLG_��������, COLG_��������
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


