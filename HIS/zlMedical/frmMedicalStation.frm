VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "comct332.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmMedicalStation 
   Caption         =   "��칤������"
   ClientHeight    =   7395
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   11400
   Icon            =   "frmMedicalStation.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7395
   ScaleWidth      =   11400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin ComCtl3.CoolBar cbrThis 
      Align           =   1  'Align Top
      Height          =   780
      Left            =   0
      TabIndex        =   29
      Top             =   0
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   1376
      BandCount       =   2
      _CBWidth        =   11400
      _CBHeight       =   780
      _Version        =   "6.7.9782"
      Child1          =   "tbrThis"
      MinWidth1       =   4995
      MinHeight1      =   720
      Width1          =   8370
      Key1            =   "only"
      NewRow1         =   0   'False
      Caption2        =   "��첿��"
      Child2          =   "cboDept"
      MinWidth2       =   2100
      MinHeight2      =   300
      Width2          =   2100
      NewRow2         =   0   'False
      Begin VB.ComboBox cboDept 
         Height          =   300
         Left            =   9210
         Style           =   2  'Dropdown List
         TabIndex        =   31
         Top             =   240
         Width           =   2100
      End
      Begin MSComctlLib.Toolbar tbrThis 
         Height          =   720
         Left            =   165
         TabIndex        =   30
         Top             =   30
         Width           =   8040
         _ExtentX        =   14182
         _ExtentY        =   1270
         ButtonWidth     =   820
         ButtonHeight    =   1270
         AllowCustomize  =   0   'False
         Wrappable       =   0   'False
         Style           =   1
         ImageList       =   "ilsMenu"
         HotImageList    =   "ilsHotMenu"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   16
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Ԥ��"
               Key             =   "Ԥ��"
               Object.ToolTipText     =   "Ԥ��"
               Object.Tag             =   "Ԥ��"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��ӡ"
               Key             =   "��ӡ"
               Object.ToolTipText     =   "��ӡ"
               Object.Tag             =   "��ӡ"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_1"
               Style           =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "�������"
               Object.Tag             =   "����"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "���"
               Key             =   "���"
               Object.ToolTipText     =   "���"
               Object.Tag             =   "���"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_2"
               Style           =   3
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "��д"
               Key             =   "��д"
               Object.ToolTipText     =   "��д"
               Object.Tag             =   "��д"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�ܼ�"
               Key             =   "�ܼ�"
               Object.ToolTipText     =   "�ܼ�"
               Object.Tag             =   "�ܼ�"
               ImageIndex      =   7
               Style           =   5
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_3"
               Style           =   3
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   9
               Style           =   5
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_4"
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Key             =   "Split_5"
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "����"
               Key             =   "����"
               Object.ToolTipText     =   "����"
               Object.Tag             =   "����"
               ImageIndex      =   11
            EndProperty
            BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "�˳�"
               Key             =   "�˳�"
               Object.ToolTipText     =   "�˳�"
               Object.Tag             =   "�˳�"
               ImageIndex      =   12
            EndProperty
         EndProperty
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid vsfPrint 
      Height          =   1605
      Left            =   3510
      TabIndex        =   34
      Top             =   3150
      Visible         =   0   'False
      Width           =   2760
      _ExtentX        =   4868
      _ExtentY        =   2831
      _Version        =   393216
      FixedCols       =   0
      RowHeightMin    =   270
      MergeCells      =   1
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Timer tmr 
      Enabled         =   0   'False
      Interval        =   60000
      Left            =   6330
      Top             =   4335
   End
   Begin VB.PictureBox picClass 
      Height          =   5220
      Left            =   135
      ScaleHeight     =   5160
      ScaleWidth      =   2730
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   810
      Width           =   2790
      Begin VB.CommandButton cmdKind 
         Caption         =   "&Z.�Զ����ѯ"
         Height          =   300
         Index           =   3
         Left            =   135
         Style           =   1  'Graphical
         TabIndex        =   3
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   1005
         Width           =   1785
      End
      Begin VB.CommandButton cmdKind 
         Caption         =   "&C.������"
         Height          =   300
         Index           =   2
         Left            =   135
         TabIndex        =   4
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   690
         Width           =   1785
      End
      Begin VB.PictureBox picShow 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   4545
         Left            =   105
         ScaleHeight     =   4515
         ScaleWidth      =   2640
         TabIndex        =   7
         Top             =   1335
         Width           =   2670
         Begin VSFlex8Ctl.VSFlexGrid vsf 
            Height          =   1020
            Index           =   3
            Left            =   4875
            TabIndex        =   8
            Top             =   255
            Visible         =   0   'False
            Width           =   1620
            _cx             =   2857
            _cy             =   1799
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
            BackColorSel    =   16772055
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   12698049
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   2
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
            RowHeightMin    =   255
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   1
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
            Begin VB.Line lnX3 
               Index           =   0
               Visible         =   0   'False
               X1              =   -555
               X2              =   1230
               Y1              =   555
               Y2              =   555
            End
            Begin VB.Line lnY3 
               Index           =   0
               Visible         =   0   'False
               X1              =   270
               X2              =   270
               Y1              =   420
               Y2              =   1635
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsf 
            Height          =   1020
            Index           =   2
            Left            =   4305
            TabIndex        =   9
            Top             =   -60
            Visible         =   0   'False
            Width           =   1620
            _cx             =   2857
            _cy             =   1799
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
            BackColorSel    =   16772055
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   12698049
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   2
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
            RowHeightMin    =   255
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   1
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
            Begin VB.Line lnY2 
               Index           =   0
               Visible         =   0   'False
               X1              =   270
               X2              =   270
               Y1              =   420
               Y2              =   1635
            End
            Begin VB.Line lnX2 
               Index           =   0
               Visible         =   0   'False
               X1              =   -555
               X2              =   1230
               Y1              =   555
               Y2              =   555
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsf 
            Height          =   1020
            Index           =   1
            Left            =   5085
            TabIndex        =   10
            Top             =   570
            Visible         =   0   'False
            Width           =   1620
            _cx             =   2857
            _cy             =   1799
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
            BackColorSel    =   16772055
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   12698049
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   2
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
            RowHeightMin    =   255
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   1
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
            Begin VB.Line lnX1 
               Index           =   0
               Visible         =   0   'False
               X1              =   -555
               X2              =   1230
               Y1              =   555
               Y2              =   555
            End
            Begin VB.Line lnY1 
               Index           =   0
               Visible         =   0   'False
               X1              =   270
               X2              =   270
               Y1              =   420
               Y2              =   1635
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsf 
            Height          =   1395
            Index           =   0
            Left            =   225
            TabIndex        =   11
            Top             =   120
            Visible         =   0   'False
            Width           =   1830
            _cx             =   3228
            _cy             =   2461
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
            BackColorSel    =   16772055
            ForeColorSel    =   0
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   12698049
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   2
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
            RowHeightMin    =   255
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   ""
            ScrollTrack     =   0   'False
            ScrollBars      =   3
            ScrollTips      =   0   'False
            MergeCells      =   1
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
            Begin VB.Line lnY0 
               Index           =   0
               Visible         =   0   'False
               X1              =   270
               X2              =   270
               Y1              =   420
               Y2              =   1635
            End
            Begin VB.Line lnX0 
               Index           =   0
               Visible         =   0   'False
               X1              =   -555
               X2              =   1230
               Y1              =   555
               Y2              =   555
            End
         End
      End
      Begin VB.CommandButton cmdKind 
         Caption         =   "&B.�������"
         Height          =   300
         Index           =   1
         Left            =   135
         TabIndex        =   5
         TabStop         =   0   'False
         Tag             =   "1"
         Top             =   390
         Width           =   1785
      End
      Begin VB.CommandButton cmdKind 
         Caption         =   "&A.�ȴ����"
         Height          =   300
         Index           =   0
         Left            =   135
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   90
         Width           =   1785
      End
   End
   Begin VB.PictureBox picContainer 
      AutoRedraw      =   -1  'True
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   1725
      Left            =   4485
      ScaleHeight     =   1725
      ScaleWidth      =   3975
      TabIndex        =   1
      Top             =   5085
      Width           =   3975
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   7035
      Width           =   11400
      _ExtentX        =   20108
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmMedicalStation.frx":1CFA
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   15028
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   2
            Alignment       =   1
            AutoSize        =   2
            Enabled         =   0   'False
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
   Begin MSComctlLib.TabStrip tbs 
      Height          =   300
      Left            =   3435
      TabIndex        =   12
      Top             =   2595
      Width           =   5520
      _ExtentX        =   9737
      _ExtentY        =   529
      Style           =   2
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&1.����"
            Key             =   "����"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&2.�ܼ�"
            Key             =   "�ܼ�"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&3.����"
            Key             =   "����"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "&4.����"
            Key             =   "����"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsMenu 
      Left            =   9405
      Top             =   3180
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":258E
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":27AE
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":29CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":3148
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":38C2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":403C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":4256
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":49D0
            Key             =   "����"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":514A
            Key             =   "����"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":58C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":5ADE
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":5CFE
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ilsHotMenu 
      Left            =   8790
      Top             =   3180
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   24
      ImageHeight     =   24
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   12
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":5F1E
            Key             =   "Preview"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":613E
            Key             =   "Print"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":635E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":6AD8
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":7252
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":79CC
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":7BE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":8360
            Key             =   "����"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":8ADA
            Key             =   "����"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":9254
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":946E
            Key             =   "Help"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":968E
            Key             =   "Quit"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraBack 
      Height          =   1605
      Left            =   3360
      TabIndex        =   13
      Top             =   690
      Width           =   7485
      Begin VB.PictureBox pic 
         BorderStyle     =   0  'None
         Height          =   1380
         Index           =   1
         Left            =   6240
         ScaleHeight     =   1380
         ScaleWidth      =   6090
         TabIndex        =   47
         Top             =   90
         Visible         =   0   'False
         Width           =   6090
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   12
            Left            =   45
            TabIndex        =   55
            Top             =   60
            Width           =   810
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����������Ϣ��ҵ��˾"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   2
            Left            =   870
            TabIndex        =   54
            Top             =   60
            Width           =   1800
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�� ϵ ��:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   13
            Left            =   45
            TabIndex        =   53
            Top             =   330
            Width           =   810
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ϵ�绰:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   14
            Left            =   45
            TabIndex        =   52
            Top             =   630
            Width           =   810
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�����ʼ�:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   15
            Left            =   45
            TabIndex        =   51
            Top             =   915
            Width           =   810
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����������Ϣ��ҵ��˾"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   11
            Left            =   870
            TabIndex        =   50
            Top             =   330
            Width           =   1800
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����������Ϣ��ҵ��˾"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   12
            Left            =   870
            TabIndex        =   49
            Top             =   630
            Width           =   1800
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����������Ϣ��ҵ��˾"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   13
            Left            =   870
            TabIndex        =   48
            Top             =   915
            Width           =   1800
         End
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1425
         Index           =   0
         Left            =   30
         ScaleHeight     =   1425
         ScaleWidth      =   7050
         TabIndex        =   14
         Top             =   135
         Width           =   7050
         Begin VB.PictureBox picPhoto 
            Appearance      =   0  'Flat
            AutoRedraw      =   -1  'True
            BackColor       =   &H80000000&
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   1425
            Left            =   4965
            ScaleHeight     =   1425
            ScaleWidth      =   1020
            TabIndex        =   42
            Top             =   0
            Width           =   1018
         End
         Begin VB.PictureBox picState 
            Appearance      =   0  'Flat
            BackColor       =   &H80000000&
            ForeColor       =   &H80000008&
            Height          =   480
            Left            =   6030
            ScaleHeight     =   450
            ScaleWidth      =   540
            TabIndex        =   15
            Top             =   45
            Visible         =   0   'False
            Width           =   570
            Begin VB.Shape shpState 
               BorderColor     =   &H000000FF&
               BorderWidth     =   2
               Height          =   360
               Left            =   60
               Top             =   45
               Width           =   450
            End
            Begin VB.Label lblState 
               Appearance      =   0  'Flat
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��"
               BeginProperty Font 
                  Name            =   "����"
                  Size            =   18
                  Charset         =   134
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               ForeColor       =   &H000000FF&
               Height          =   360
               Left            =   90
               TabIndex        =   16
               Top             =   30
               Width           =   375
            End
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "12345678"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   1
            Left            =   3750
            TabIndex        =   46
            Top             =   330
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   11
            Left            =   3105
            TabIndex        =   45
            Top             =   330
            Width           =   630
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "139"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   10
            Left            =   2475
            TabIndex        =   44
            Top             =   345
            Width           =   270
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ϵ�绰:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   9
            Left            =   1650
            TabIndex        =   43
            Top             =   345
            Width           =   810
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ƭ:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   6
            Left            =   4470
            TabIndex        =   41
            Top             =   45
            Width           =   450
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   9
            Left            =   870
            TabIndex        =   40
            Top             =   915
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ײ�:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   2
            Left            =   75
            TabIndex        =   39
            Top             =   915
            Width           =   810
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ѻ�"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   0
            Left            =   870
            TabIndex        =   33
            Top             =   330
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����״��:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   45
            TabIndex        =   32
            Top             =   330
            Width           =   810
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������λ:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   7
            Left            =   75
            TabIndex        =   28
            Top             =   1170
            Width           =   810
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�� �� ��:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   8
            Left            =   45
            TabIndex        =   27
            Top             =   630
            Width           =   810
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�������:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   10
            Left            =   1650
            TabIndex        =   26
            Top             =   600
            Width           =   810
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ܼ���Ա:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   3
            Left            =   45
            TabIndex        =   25
            Top             =   60
            Width           =   810
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��    ��:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   4
            Left            =   1650
            TabIndex        =   24
            Top             =   60
            Width           =   810
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��  ��:"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   5
            Left            =   3105
            TabIndex        =   23
            Top             =   45
            Width           =   630
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   3
            Left            =   870
            TabIndex        =   22
            Top             =   60
            Width           =   540
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   4
            Left            =   2475
            TabIndex        =   21
            Top             =   60
            Width           =   180
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "30"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   5
            Left            =   3750
            TabIndex        =   20
            Top             =   45
            Width           =   180
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "2004-12-20"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   6
            Left            =   2475
            TabIndex        =   19
            Top             =   600
            Width           =   900
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "666666"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   7
            Left            =   870
            TabIndex        =   18
            Top             =   630
            Width           =   540
         End
         Begin VB.Label lblValue 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����������Ϣ��ҵ��˾"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   8
            Left            =   870
            TabIndex        =   17
            Top             =   1170
            Width           =   1800
         End
      End
   End
   Begin MSComctlLib.ImageList ils13 
      Left            =   3540
      Top             =   4995
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":98AE
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":9C48
            Key             =   "״̬"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":9FE2
            Key             =   "סԺ"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":A37C
            Key             =   "����"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":A716
            Key             =   "����"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":A9AC
            Key             =   "���"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":AF46
            Key             =   "�¿�"
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":B4E0
            Key             =   "��ʼ"
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":BA7A
            Key             =   "ȡ��"
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":C014
            Key             =   "ȷ��"
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":C5AE
            Key             =   "up"
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":C770
            Key             =   "down"
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":C932
            Key             =   "Close"
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMedicalStation.frx":CBC8
            Key             =   "Open"
         EndProperty
      EndProperty
   End
   Begin VB.Frame fraInfo 
      Height          =   630
      Left            =   225
      TabIndex        =   35
      Top             =   6270
      Width           =   2925
      Begin VB.CommandButton cmdMenu 
         Height          =   270
         Left            =   75
         Picture         =   "frmMedicalStation.frx":CE5E
         Style           =   1  'Graphical
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   240
         Width           =   285
      End
      Begin VB.TextBox txt 
         Height          =   300
         Index           =   1
         Left            =   1380
         TabIndex        =   37
         Top             =   225
         Width           =   1140
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "&6.����"
         Height          =   180
         Index           =   1
         Left            =   420
         TabIndex        =   36
         Tag             =   "����"
         Top             =   285
         Width           =   540
      End
   End
   Begin VB.Image imgY_S 
      Height          =   4395
      Left            =   3255
      MousePointer    =   9  'Size W E
      Top             =   1410
      Width           =   45
   End
   Begin VB.Menu mnuFile 
      Caption         =   "�ļ�(&F)"
      Begin VB.Menu mnuFilePrintSet 
         Caption         =   "��ӡ����(&S)"
      End
      Begin VB.Menu mnuFilePrintView 
         Caption         =   "��ӡԤ��(&V)"
      End
      Begin VB.Menu mnuFilePrint 
         Caption         =   "��ӡ(&P)"
         Shortcut        =   ^P
      End
      Begin VB.Menu mnuFileOutExcel 
         Caption         =   "�����&Excel"
      End
      Begin VB.Menu mnuFile_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFilePrintList 
         Caption         =   "���ָ����(&L)"
      End
      Begin VB.Menu mnuFileRequest 
         Caption         =   "������뵥(&H)"
      End
      Begin VB.Menu mnuFilePrintRequest 
         Caption         =   "��Ŀ���뵥(&R)"
      End
      Begin VB.Menu mnuFilePrintBook 
         Caption         =   "��챨����(&B)"
         Shortcut        =   {F4}
      End
      Begin VB.Menu mnuFileSendMail 
         Caption         =   "���ͱ�����(&I)"
      End
      Begin VB.Menu mnuFile_11 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileRptDesign 
         Caption         =   "�������(&D)"
         Begin VB.Menu mnuReportDesign 
            Caption         =   "�������(&1)"
            Index           =   0
         End
         Begin VB.Menu mnuReportDesign 
            Caption         =   "������(&2)"
            Index           =   1
         End
         Begin VB.Menu mnuReportDesign 
            Caption         =   "�����ܼ�(&3)"
            Index           =   2
         End
      End
      Begin VB.Menu mnuFilePara 
         Caption         =   "��������(&M)"
         Shortcut        =   {F12}
      End
      Begin VB.Menu mnuFile_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuFileExit 
         Caption         =   "�˳�(&X)"
      End
   End
   Begin VB.Menu mnuMedical 
      Caption         =   "���(&T)"
      Begin VB.Menu mnuMedicalNew 
         Caption         =   "���Ǽ�(&R)"
         Begin VB.Menu mnuMedicalNewType 
            Caption         =   "��������(&1)"
            Index           =   0
         End
         Begin VB.Menu mnuMedicalNewType 
            Caption         =   "��������(&2)"
            Index           =   1
         End
         Begin VB.Menu mnuMedicalNewType 
            Caption         =   "-"
            Index           =   2
         End
         Begin VB.Menu mnuMedicalNewType 
            Caption         =   "�޸ĵǼ�(&3)"
            Index           =   3
         End
         Begin VB.Menu mnuMedicalNewType 
            Caption         =   "ɾ���Ǽ�(&4)"
            Index           =   4
         End
      End
      Begin VB.Menu mnuMedicalPhoto 
         Caption         =   "��Ƭ�ɼ�(&P)"
      End
      Begin VB.Menu mnuMedical_0 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMedicalBegin 
         Caption         =   "�������(&B)"
         Shortcut        =   ^B
      End
      Begin VB.Menu mnuMedicalBeginCancel 
         Caption         =   "ȡ������(&C)"
      End
      Begin VB.Menu mnuMedicalGroupIn 
         Caption         =   "��Ա����(&I)"
      End
      Begin VB.Menu mnuMedicalGroupOut 
         Caption         =   "ȡ������(&X)"
      End
      Begin VB.Menu mnuMedical_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMedicalComplete 
         Caption         =   "������(&E)"
      End
      Begin VB.Menu mnuMedicalCompleteCancel 
         Caption         =   "ȡ�����(&R)"
      End
      Begin VB.Menu mnuMedical_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMedicalItems 
         Caption         =   "�����Ŀ(&D)"
      End
      Begin VB.Menu mnuMedicalItemsAddtion 
         Caption         =   "��Ա��Ŀ(&A)"
      End
      Begin VB.Menu mnuMedical_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMedicalGroupAdd 
         Caption         =   "�����Ա(&N)"
      End
      Begin VB.Menu mnuMedicalGroupDelete 
         Caption         =   "�Ƴ���Ա(&D)"
      End
      Begin VB.Menu mnuMedical_4 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMedicalDept 
         Caption         =   "ִ�е���(&M)"
      End
      Begin VB.Menu mnuMedicalCallBack 
         Caption         =   "������Ŀ(&K)"
      End
   End
   Begin VB.Menu mnuReport 
      Caption         =   "����(&E)"
      Begin VB.Menu mnuReportWrite 
         Caption         =   "��д����(&W)"
         Shortcut        =   {F2}
      End
      Begin VB.Menu mnuReportView 
         Caption         =   "�鿴����(&V)"
      End
      Begin VB.Menu mnuReportWriteMuli 
         Caption         =   "��������(&B)"
      End
      Begin VB.Menu mnuReport_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuReportAddOutLine 
         Caption         =   "����ܼ�(&A)"
         Begin VB.Menu mnuReportAddOutLineCase 
            Caption         =   "<�޿����ܼ�>"
            Index           =   0
         End
      End
      Begin VB.Menu mnuReportModifyOutLine 
         Caption         =   "�޸��ܼ�(&M)"
      End
      Begin VB.Menu mnuReportDelOutLine 
         Caption         =   "ɾ���ܼ�(&D)"
      End
   End
   Begin VB.Menu mnuCharge 
      Caption         =   "����(&C)"
      Begin VB.Menu mnuChargeMain 
         Caption         =   "����������(&G)"
      End
      Begin VB.Menu mnuCharge_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuChargeAdd 
         Caption         =   "���Ӹ��ӷ�(&A)"
         Begin VB.Menu mnuChargeAddType 
            Caption         =   "�շѵ���(&1)"
            Index           =   0
         End
         Begin VB.Menu mnuChargeAddType 
            Caption         =   "���ʵ���(&2)"
            Index           =   1
         End
         Begin VB.Menu mnuChargeAddType 
            Caption         =   "��Ѻ��õǼ�(&3)"
            Index           =   2
         End
      End
      Begin VB.Menu mnuChargeModify 
         Caption         =   "�޸ĸ��ӷ�(&M)"
      End
      Begin VB.Menu mnuChargeDelete 
         Caption         =   "ɾ�����ӷ�(&D)"
      End
   End
   Begin VB.Menu mnuView 
      Caption         =   "�鿴(&V)"
      Begin VB.Menu mnuViewTool 
         Caption         =   "������(&T)"
         Begin VB.Menu mnuViewToolButton 
            Caption         =   "��׼��ť(&S)"
            Checked         =   -1  'True
         End
         Begin VB.Menu mnuViewToolText 
            Caption         =   "�ı���ǩ(&T)"
            Checked         =   -1  'True
         End
      End
      Begin VB.Menu mnuViewStatus 
         Caption         =   "״̬��(&S)"
         Checked         =   -1  'True
      End
      Begin VB.Menu mnuView_1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewShowResult 
         Caption         =   "��ʾ����(&H)"
      End
      Begin VB.Menu mnuView_2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewPatientBrowse 
         Caption         =   "��Ա��Ϣ(&B)"
      End
      Begin VB.Menu mnuViewFind 
         Caption         =   "�ۺϲ�ѯ(&F)"
         Begin VB.Menu mnuViewFilterItem 
            Caption         =   "�Զ���..."
            Index           =   0
         End
         Begin VB.Menu mnuViewFilterItem 
            Caption         =   "-"
            Index           =   1
            Visible         =   0   'False
         End
      End
      Begin VB.Menu mnuView_3 
         Caption         =   "-"
      End
      Begin VB.Menu mnuViewFilter 
         Caption         =   "����(&F)"
         Shortcut        =   ^F
      End
      Begin VB.Menu mnuViewRefresh 
         Caption         =   "ˢ��(&R)"
         Shortcut        =   {F5}
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "����(&H)"
      Begin VB.Menu mnuHelpTopic 
         Caption         =   "��������(&H)"
         Shortcut        =   {F1}
      End
      Begin VB.Menu mnuHelpWeb 
         Caption         =   "&Web�ϵ�����"
         Begin VB.Menu mnuHelpWebHome 
            Caption         =   "������ҳ(&H)"
         End
         Begin VB.Menu mnuHelpWebForum 
            Caption         =   "������̳(&F)"
         End
         Begin VB.Menu mnuHelpWebMail 
            Caption         =   "���ͷ���(&K)..."
         End
      End
      Begin VB.Menu h1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "����(&A)..."
      End
   End
End
Attribute VB_Name = "frmMedicalStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'���������弶��������**************************************************************************************************
Private mblnStartUp As Boolean                          '����������־
Private mlngLoop As Long
Private mintIndex As Integer                                '��ǰ������
Private mlngSvrKey(0 To 3)  As Long                     '���ڱ����������ѡ�е��йؼ���
Private mfrmActive As Object                            '�Ӵ������
Private mlngDept As Long
Private mlngHideRows As Long
Private mblnNoAllowChange As Boolean
Private mobjCls As New clsCISWork
Private mclsCore As New clsCISCore
Private mlngCountTmr As Long
Private mstrPrivilege As String
Private mblnDataMoved As Boolean
Private mintSort As Integer

Public WithEvents mobjPopMenu As clsPopMenu                '�Զ��嵯���˵�����
Attribute mobjPopMenu.VB_VarHelpID = -1
Public mbytPopMenu As Byte

Private mlng��첡��id As Long
Private mint������ѯ���� As Integer
Private mstr���������ʱ�䷶Χ As String

Private Type usrSaveInfo
    lng�Ǽ�id As Long
    lng����id As Long
    str��� As String
End Type

Private usrSave As usrSaveInfo
Private mrsFind As New ADODB.Recordset
Private mstrSvrFind As String

'�������Զ�����̻���************************************************************************************************

Private Function SelectPerson(ByVal blnSingle As Boolean) As Boolean
    'ѡ��ĳһ���ܼ���Ա
    Dim lngLoop As Long
    Dim lngTotal As Long
    Dim lngStart As Long
    Dim lngEnd As Long
    Dim blnFirst As Boolean
    
    On Error Resume Next
    
    blnFirst = True
    lngStart = 1
    lngEnd = vsf(mintIndex).Rows - 1
    
ReStart:

    For lngLoop = lngStart To lngEnd
    
        If blnSingle = True And Val(vsf(mintIndex).TextMatrix(lngLoop, GetCol(vsf(mintIndex), "��־"))) = 0 Then
            vsf(mintIndex).IsCollapsed(lngLoop) = flexOutlineExpanded
        End If
        
        If Val(vsf(mintIndex).TextMatrix(lngLoop, GetCol(vsf(mintIndex), "�Ǽ�id"))) = usrSave.lng�Ǽ�id Then
            
            If Val(vsf(mintIndex).TextMatrix(lngLoop, GetCol(vsf(mintIndex), "��־"))) = 1 Then
                vsf(mintIndex).IsCollapsed(lngLoop) = flexOutlineExpanded
            End If
            
            If usrSave.str��� = "" Then
                vsf(mintIndex).IsCollapsed(lngLoop) = flexOutlineExpanded
                vsf(mintIndex).Row = lngLoop
                Exit For
            End If
            
            'չ��
            If Val(vsf(mintIndex).TextMatrix(lngLoop, GetCol(vsf(mintIndex), "��־"))) = 2 Then
                
                '���
                If vsf(mintIndex).TextMatrix(lngLoop, GetCol(vsf(mintIndex), "����")) = usrSave.str��� Then
                    vsf(mintIndex).IsCollapsed(lngLoop) = flexOutlineExpanded
                    
                    lngTotal = vsf(mintIndex).Rows - 1
                    If usrSave.lng����id = 0 Then
                        vsf(mintIndex).Row = lngLoop
'                        Exit For
                    End If
                End If
            End If
        
        
            If Val(vsf(mintIndex).TextMatrix(lngLoop, GetCol(vsf(mintIndex), "��־"))) >= 98 Then
                If Val(vsf(mintIndex).TextMatrix(lngLoop, GetCol(vsf(mintIndex), "����id"))) = usrSave.lng����id Then
                    vsf(mintIndex).Row = lngLoop
                    Exit For
                End If
            End If
        
        End If
    Next
    
    If lngEnd <> lngTotal And blnFirst Then
        lngStart = lngEnd + 1
        lngEnd = lngTotal
        blnFirst = False
        GoTo ReStart
    End If
    
    vsf(mintIndex).ShowCell vsf(mintIndex).Row, vsf(mintIndex).Col
    
    SelectPerson = True
    
End Function

Private Function Collapsed(ByVal intIndex As Integer, ByVal bytMode As CollapsedSettings)
    Dim lngLoop As Long

    With vsf(intIndex)
        
        For lngLoop = 1 To .Rows - 1
            
            If bytMode = flexOutlineCollapsed Then
                If .IsCollapsed(lngLoop) = flexOutlineExpanded Then
                    .IsCollapsed(lngLoop) = flexOutlineCollapsed
                End If
            Else
                If .IsCollapsed(lngLoop) = flexOutlineCollapsed Then
                    .IsCollapsed(lngLoop) = flexOutlineExpanded
                End If
            End If
        
        Next
        
    End With
    
    Call InheritAppendSpaceRows(intIndex)
End Function

Public Sub FindLocation(ByVal str���� As String)
    '--------------------------------------------------------------------------------------------------------
    '����:
    '--------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    
    lngRow = vsf(mintIndex).FindRow(str����, , GetCol(vsf(mintIndex), "����"), , False)
    If lngRow <= 0 Then
        ShowSimpleMsg "û���ҵ�����Ҫ�����Ϣ��"
    Else
        vsf(mintIndex).Row = lngRow
        vsf(mintIndex).ShowCell vsf(mintIndex).Row, vsf(mintIndex).Col
        
    End If
End Sub

Public Sub ActiveFormEnabled()
    '--------------------------------------------------------------------------------------------------------
    '
    '--------------------------------------------------------------------------------------------------------
    On Error GoTo errHand
    
    Call AdjustEnableState
    
errHand:
    
End Sub

Private Property Let AutoRefresh(vData As Boolean)
    '
    '����:�Զ�ˢ��
    '
    tmr.Enabled = vData
    
    If vData = True Then
        mlngCountTmr = 0
        tmr.Tag = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "�Զ�ˢ�¼��", 5))
        tmr.Enabled = (Val(tmr.Tag) > 0)
    End If
End Property

Private Function InitLoad() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:��ʼ�����ݣ������ڴ����Load�¼�
    '------------------------------------------------------------------------------------------------------------------
    Dim strVsf As String
    
    On Error GoTo errHand
    
    picState.BorderStyle = 0
    
    mintIndex = 0
    mlngDept = 0
    mlngSvrKey(0) = 0
    mlngSvrKey(1) = 0
    mlngSvrKey(2) = 0
    mlngSvrKey(3) = 0
    mlng��첡��id = 0
    mstrSvrFind = ""
    
    Call ResetActiveForm
        
    picShow.BorderStyle = 0
    
    For mlngLoop = 0 To cmdKind.UBound
        cmdKind(mlngLoop).Left = 15
        cmdKind(mlngLoop).Height = 300
    Next
    
    strVsf = ",450,1,1,1,[����];,255,4,1,1,[״̬];,255,4,1,1,[����];����,900,1,1,1,;�����,900,7,1,1,;������,810,1,1,1,;���￨��,900,1,1,1,;�����,990,1,1,1,;�Ա�,450,1,1,1,;����,450,1,1,1,;����״��,0,1,1,0,;��쵥��,0,1,1,0,;����,0,7,1,0,;�ϼ�id,0,1,1,0,;����id,0,1,1,0,;�Ǽ�id,0,1,1,0,;��־,0,1,1,0,;�Ƿ�װ��,0,1,1,0,"
    
    Call CreateVsf(vsf(0), strVsf)
    vsf(0).Cols = vsf(0).Cols + 1
    vsf(0).ColWidth(vsf(0).Cols - 1) = 15
    Set vsf(0).Cell(flexcpPicture, 0, 1) = ils13.ListImages("״̬").Picture
    Set vsf(0).Cell(flexcpPicture, 0, 2) = ils13.ListImages("����").Picture
    
    With vsf(0)
        .ExtendLastCol = True
        .Rows = 1
        .OutlineBar = flexOutlineBarComplete
    End With
    
    Call CreateVsf(vsf(1), strVsf)
    vsf(1).Cols = vsf(1).Cols + 1
    vsf(1).ColWidth(vsf(1).Cols - 1) = 15
    Set vsf(1).Cell(flexcpPicture, 0, 1) = ils13.ListImages("״̬").Picture
    Set vsf(1).Cell(flexcpPicture, 0, 2) = ils13.ListImages("����").Picture
    With vsf(1)
        .ExtendLastCol = True
        .Rows = 1
        .OutlineBar = flexOutlineBarComplete
    End With
        
    Call CreateVsf(vsf(2), strVsf)
    vsf(2).Cols = vsf(2).Cols + 1
    vsf(2).ColWidth(vsf(2).Cols - 1) = 15
    Set vsf(2).Cell(flexcpPicture, 0, 1) = ils13.ListImages("״̬").Picture
    Set vsf(2).Cell(flexcpPicture, 0, 2) = ils13.ListImages("����").Picture
    With vsf(2)
        .ExtendLastCol = True
        .Rows = 1
        .OutlineBar = flexOutlineBarComplete
    End With
    
    strVsf = ",255,4,1,1,[״̬];�Ǽ�id,0,1,1,0,;��쵥��,990,1,1,1,;����,900,1,1,1,;�Ա�,450,1,1,1,;����,450,1,1,1,;����״��,0,1,1,0,;�����,900,7,1,1,;������,810,1,1,1,;��������,1500,1,1,0,;����,1500,1,1,1,;����id,0,1,1,0,"
    Call CreateVsf(vsf(3), strVsf)
    vsf(3).Cols = vsf(3).Cols + 1
    vsf(3).ColWidth(vsf(3).Cols - 1) = 15
    Set vsf(3).Cell(flexcpPicture, 0, 0) = ils13.ListImages("״̬").Picture

    Dim strStart As String
    Dim strEnd As String
    
    strStart = GetDateTime(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "�����ʱ�䷶Χ", "��  ��"), 1)
    strEnd = GetDateTime(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "�����ʱ�䷶Χ", "��  ��"), 2)
    If strStart = "" Then strStart = GetDateTime("��  ��", 1)
    If strEnd = "" Then strEnd = GetDateTime("��  ��", 2)
    vsf(0).Tag = strStart & "|" & strEnd
    
    strStart = GetDateTime(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "�����ʱ�䷶Χ", "��  ��"), 1)
    strEnd = GetDateTime(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "�����ʱ�䷶Χ", "��  ��"), 2)
    If strStart = "" Then strStart = GetDateTime("��  ��", 1)
    If strEnd = "" Then strEnd = GetDateTime("��  ��", 2)
    vsf(1).Tag = strStart & "|" & strEnd
    
    '����ȱʡʱ�䷶Χ
    strStart = GetDateTime(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "���������ʱ�䷶Χ", "��  ��"), 1)
    strEnd = GetDateTime(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "���������ʱ�䷶Χ", "��  ��"), 2)
    If strStart = "" Then strStart = GetDateTime("��  ��", 1)
    If strEnd = "" Then strEnd = GetDateTime("��  ��", 2)
    mstr���������ʱ�䷶Χ = strStart & "|" & strEnd
    
    '�����������
    mint������ѯ���� = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "������ѯ����", "0"))
    
    strStart = GetDateTime(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "�������ʱ�䷶Χ", "��  ��"), 1)
    strEnd = GetDateTime(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "�������ʱ�䷶Χ", "��  ��"), 2)
    If strStart = "" Then strStart = GetDateTime("��  ��", 1)
    If strEnd = "" Then strEnd = GetDateTime("��  ��", 2)
    vsf(2).Tag = strStart & "|" & strEnd
    
    InitLoad = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Function InitActive() As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:��ʼ�����ݣ������ڴ����Active�¼�
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    On Error GoTo errHand
    
    tmr.Tag = "5"
    mlngCountTmr = 0
    
    gstrSQL = GetPublicSQL(SQL.��첿���嵥, IIf(InStr(gstrPrivs, "���п���") > 0, "����", ""))
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, UserInfo.ID)
    If rs.BOF Then
        ShowSimpleMsg "û��������ʵĲ��ţ����ڲ��Ź��������ã�"
        Exit Function
    End If
    
    '�����ݵ��ؼ���
    Call AddComboData(cboDept, rs)
    zlControl.CboLocate cboDept, UserInfo.����ID, True
    If cboDept.ListCount > 0 And cboDept.ListIndex = -1 Then cboDept.ListIndex = 0
    
    '3.��ȡע����������
    If Val(GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser, "ʹ�ø��Ի����", "0")) = 1 Then
        'ʹ�ø��Ի�����
        
        mnuViewShowResult.Checked = (Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��ʾ����", 0)) = 1)
        
        tmr.Tag = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "�Զ�ˢ�¼��", 5))
        lbl(1).Caption = "&6." & (GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "������Ϣ", "����"))
        lbl(1).Tag = Mid(lbl(1).Caption, 4)

    End If
    
    '4.�����¿ؼ�����
    tmr.Enabled = (Val(tmr.Tag) > 0)
    
    Call RefreshQueryMenu
    
    InitActive = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub InheritResetVsf(ByVal intIndex As Integer)
    '--------------------------------------------------------------------------------------------------------
    '�̳�ResetVsf����
    '--------------------------------------------------------------------------------------------------------
    On Error Resume Next

    Call ResetVsf(vsf(intIndex))
    vsf(intIndex).Cell(flexcpFontBold, 1, 0, 1, vsf(intIndex).Cols - 1) = False
    
    Call InheritAppendSpaceRows(intIndex)
End Sub

Private Function ClearData(Optional ByVal strMenuItem As String = "") As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '------------------------------------------------------------------------------------------------------------------
    
    strMenuItem = ";" & strMenuItem & ";"
    
    If InStr(strMenuItem, ";�����Ϣ;") > 0 Then
        Call InheritResetVsf(0)
        Call InheritResetVsf(1)
        Call InheritResetVsf(2)
    End If
    
    If InStr(strMenuItem, ";���;") > 0 Then
        
        On Error Resume Next
        
        For mlngLoop = 0 To lblValue.UBound
            lblValue(mlngLoop).Caption = ""
        Next
        
        picPhoto.Cls
        
'        mlng����id = 0
'        mlng��ҳid = 0
'        mlngҽ��id = 0
'        mlng���ͺ� = 0
        
        picState.Visible = False
        
        On Error Resume Next
 
        Call mfrmActive.zlClearData
    End If
        
End Function

Private Sub ApplyPrivilege(ByVal strPrivilege As String)
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ� Ӧ��Ȩ�޴���
    '������ strPrivilege                    Ȩ��
    '------------------------------------------------------------------------------------------------------------------
'    strPrivilege = "���п���;���Ǽ�;��ʼ���;ȡ����ʼ;������;ȡ�����;�����Ŀ;������Ŀ;��ӳ�Ա;�Ƴ���Ա;��д����;����С��;��д�ܽ�;��ӡ����;�ۺϲ�ѯ;���ô���;δ�շ����"
    
    mstrPrivilege = strPrivilege
    
    If InStr(strPrivilege, "��ӡ����") = 0 Then mnuFilePrintBook.Visible = False
    
    If InStr(strPrivilege, "�������") = 0 Then mnuFileRptDesign.Visible = False
    
    If InStr(strPrivilege, "���Ǽ�") = 0 And _
        InStr(strPrivilege, "��ʼ���") = 0 And _
        InStr(strPrivilege, "ȡ����ʼ") = 0 And _
        InStr(strPrivilege, "������") = 0 And _
        InStr(strPrivilege, "ȡ�����") = 0 And _
        InStr(strPrivilege, "�����Ŀ") = 0 And _
        InStr(strPrivilege, "������Ŀ") = 0 And _
        InStr(strPrivilege, "��ӳ�Ա") = 0 And _
        InStr(strPrivilege, "�Ƴ���Ա") = 0 Then
        
        mnuMedical.Visible = False
    Else
        
        If InStr(strPrivilege, "���Ǽ�") = 0 Then mnuMedicalNew.Visible = False
                
        If InStr(strPrivilege, "��ʼ���") = 0 Then
            mnuMedicalBegin.Visible = False
            mnuMedicalGroupIn.Visible = False
            mnuMedicalGroupOut.Visible = False
        End If
        
        If InStr(strPrivilege, "ȡ����ʼ") = 0 Then mnuMedicalBeginCancel.Visible = False
        If InStr(strPrivilege, "������") = 0 Then mnuMedicalComplete.Visible = False
        If InStr(strPrivilege, "ȡ�����") = 0 Then mnuMedicalCompleteCancel.Visible = False
        If InStr(strPrivilege, "�����Ŀ") = 0 Then mnuMedicalItems.Visible = False
        If InStr(strPrivilege, "������Ŀ") = 0 Then mnuMedicalItemsAddtion.Visible = False
        If InStr(strPrivilege, "��ӳ�Ա") = 0 Then mnuMedicalGroupAdd.Visible = False
        If InStr(strPrivilege, "�Ƴ���Ա") = 0 Then mnuMedicalGroupDelete.Visible = False
        
        Dim aryMenu As Variant
        
        aryMenu = Array(mnuMedicalNew, mnuMedical_0, mnuMedicalBegin, mnuMedicalBeginCancel, mnuMedicalGroupIn, mnuMedicalGroupOut, mnuMedical_1, mnuMedicalComplete, mnuMedicalCompleteCancel, mnuMedical_2, mnuMedicalItems, mnuMedicalItemsAddtion, mnuMedical_3, mnuMedicalGroupAdd, mnuMedicalGroupDelete)
        
        Call AdjustSplit(aryMenu)
        
    End If
    
    If InStr(strPrivilege, "��д����") = 0 Then
        mnuReportWrite.Visible = False
        mnuReportWriteMuli.Visible = False
    End If
    
    If InStr(strPrivilege, "��д�ܽ�") = 0 Then
        mnuReportAddOutLine.Visible = False
'        mnuReportAgain.Visible = False
    End If
    
    mnuReport_1.Visible = mnuReportAddOutLine.Visible
    'mnuReport_3.Visible = mnuReportAddOutLine.Visible
    
    If InStr(strPrivilege, "���ô���") = 0 Then mnuCharge.Visible = False
    
    If InStr(strPrivilege, "�ۺϲ�ѯ") = 0 Then mnuViewFind.Visible = False
    
    mnuReportModifyOutLine.Visible = mnuReportAddOutLine.Visible
    mnuReportDelOutLine.Visible = mnuReportAddOutLine.Visible
            
    tbrThis.Buttons("����").Visible = mnuMedicalBegin.Visible And mnuMedical.Visible
    tbrThis.Buttons("���").Visible = mnuMedicalComplete.Visible And mnuMedical.Visible
    tbrThis.Buttons("��д").Visible = mnuReportWrite.Visible And mnuReport.Visible
    tbrThis.Buttons("�ܼ�").Visible = mnuReportAddOutLine.Visible And mnuReport.Visible
    tbrThis.Buttons("����").Visible = mnuCharge.Visible
    tbrThis.Buttons("����").Visible = mnuCharge.Visible
    
    tbrThis.Buttons("Split_2").Visible = tbrThis.Buttons("����").Visible Or tbrThis.Buttons("���").Visible
    tbrThis.Buttons("Split_3").Visible = tbrThis.Buttons("�ܼ�").Visible Or tbrThis.Buttons("��д").Visible
    tbrThis.Buttons("Split_4").Visible = tbrThis.Buttons("����").Visible Or tbrThis.Buttons("����").Visible
    tbrThis.Buttons("Split_5").Visible = tbrThis.Buttons("����").Visible

End Sub

Private Sub AdjustSplit(ByVal aryMenu As Variant)
    
    Dim lngLoop As Long
    Dim lngPos As Long
    Dim lngSvrPos As Long
        
    For lngLoop = 0 To UBound(aryMenu)
        
        If aryMenu(lngLoop).Visible Then
            
            lngPos = lngPos + 1
            
            If aryMenu(lngLoop).Caption = "-" Then
                If lngPos = 1 Then
                    aryMenu(lngLoop).Visible = False
                    lngPos = 0
                Else
                    
                    If lngSvrPos + 1 = lngPos Then
                        aryMenu(lngLoop).Visible = False
                    End If
                    
                    lngSvrPos = lngPos
                End If
                
            End If
        End If
    Next
    
    If lngSvrPos = lngPos And lngPos > 0 Then
        aryMenu(lngSvrPos).Visible = False
    End If
    
        
End Sub


Private Sub AdjustEnableState()
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ� ���������ܲ˵��Ŀ���״̬
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    mnuFilePrintView.Enabled = True
    mnuFilePrint.Enabled = True
    mnuFileOutExcel.Enabled = True
    mnuFilePrintRequest.Enabled = True
    mnuFilePrintList.Enabled = True
    mnuFilePrintBook.Enabled = True
    mnuFileSendMail.Enabled = True
    
    mnuMedicalNew.Enabled = True
    
    mnuMedicalBegin.Enabled = True
    mnuMedicalBeginCancel.Enabled = True
    mnuMedicalComplete.Enabled = True
    mnuMedicalCompleteCancel.Enabled = True
    
    mnuMedicalGroupIn.Enabled = True
    mnuMedicalGroupOut.Enabled = True
    mnuMedicalGroupAdd.Enabled = True
    mnuMedicalGroupDelete.Enabled = True
    
    mnuMedicalItems.Enabled = True
    mnuMedicalItemsAddtion.Enabled = True
    
    mnuReportWrite.Enabled = True
    
    mnuReportAddOutLine.Enabled = True
    mnuReportModifyOutLine.Enabled = True
    mnuReportDelOutLine.Enabled = True
    
'    mnuReportAgain.Enabled = True
    mnuReportView.Enabled = True
    
    mnuReportWriteMuli.Enabled = True
        
    
    mnuChargeMain.Enabled = True
    mnuChargeAdd.Enabled = True
    mnuChargeModify.Enabled = True
    mnuChargeDelete.Enabled = True
    
    mnuViewPatientBrowse.Enabled = True
       
    Select Case mintIndex
    Case 0
        mnuFilePrintBook.Enabled = False
        mnuFilePrintRequest.Enabled = False
'        mnuFilePrintList.Enabled = False
        
        mnuMedicalBeginCancel.Enabled = False
        mnuMedicalComplete.Enabled = False
        mnuMedicalCompleteCancel.Enabled = False
        
        mnuReportWrite.Enabled = False
        mnuReportAddOutLine.Enabled = False
        mnuReportModifyOutLine.Enabled = False
        mnuReportDelOutLine.Enabled = False
        mnuReportView.Enabled = False
            
        mnuReportWriteMuli.Enabled = False
'        mnuReportAgain.Enabled = False
            
        mnuMedicalGroupIn.Enabled = False
        mnuMedicalGroupAdd.Enabled = False
        mnuMedicalGroupDelete.Enabled = False
        mnuMedicalItems.Enabled = False
        mnuMedicalItemsAddtion.Enabled = False
        
        mnuChargeMain.Enabled = False
        mnuChargeAdd.Enabled = False
        mnuChargeModify.Enabled = False
        mnuChargeDelete.Enabled = False
    
        If Val(vsf(mintIndex).RowData(vsf(mintIndex).Row)) = 0 Then
            mnuMedicalBegin.Enabled = False
            mnuFilePrintList.Enabled = False
        Else
            '0-���˷�����;1-����������;2-���������;99-�ܼ���Ա��
            Select Case Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "��־")))
            Case 0
                mnuMedicalBegin.Enabled = False
                mnuFilePrintList.Enabled = False
            Case 2
                mnuMedicalBegin.Enabled = False
            Case 99
                If Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "�ϼ�id"))) > 0 Then
                    mnuMedicalBegin.Enabled = False
                End If
            End Select
        End If
        
    Case 1
        mnuMedicalNew.Enabled = False
        mnuMedicalBegin.Enabled = False
        
        If vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "[״̬]")) = "���" Then
            mnuReportWrite.Enabled = False
            mnuReportAddOutLine.Enabled = False
            mnuReportModifyOutLine.Enabled = False
            mnuReportDelOutLine.Enabled = False
            mnuReportView.Enabled = False
            mnuReportWriteMuli.Enabled = False
            'mnuReportAgain.Enabled = False
            
            mnuMedicalGroupIn.Enabled = False
            mnuMedicalGroupAdd.Enabled = False
            mnuMedicalGroupDelete.Enabled = False
            mnuMedicalItems.Enabled = False
            mnuMedicalItemsAddtion.Enabled = False
            
            mnuChargeMain.Enabled = False
            mnuChargeAdd.Enabled = False
            mnuChargeModify.Enabled = False
            mnuChargeDelete.Enabled = False
            
            mnuMedicalComplete.Enabled = False
        Else
            mnuMedicalCompleteCancel.Enabled = False
        End If
        
        If Val(vsf(mintIndex).RowData(vsf(mintIndex).Row)) = 0 Then
            mnuMedicalBeginCancel.Enabled = False
            mnuMedicalComplete.Enabled = False
            
            mnuReportWrite.Enabled = False
            mnuReportAddOutLine.Enabled = False
            mnuReportModifyOutLine.Enabled = False
            mnuReportDelOutLine.Enabled = False
            mnuReportView.Enabled = False
            mnuReportWriteMuli.Enabled = False
'            mnuReportAgain.Enabled = False
            
            mnuFilePrintBook.Enabled = False
            mnuFilePrintRequest.Enabled = False
            mnuFilePrintList.Enabled = False
            
            mnuViewPatientBrowse.Enabled = False
                
            mnuMedicalGroupIn.Enabled = False
            mnuMedicalGroupAdd.Enabled = False
            mnuMedicalGroupDelete.Enabled = False
            mnuMedicalItems.Enabled = False
            mnuMedicalItemsAddtion.Enabled = False
            
            mnuChargeMain.Enabled = False
            mnuChargeAdd.Enabled = False
            mnuChargeModify.Enabled = False
            mnuChargeDelete.Enabled = False
        End If
        
        Select Case Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "��־")))
            Case 0               '0-���˷�����;1-����������;2-���������;98-�������ܼ���Ա��;99-�����ܼ���Ա
                mnuMedicalBeginCancel.Enabled = False
                mnuMedicalComplete.Enabled = False
                
                mnuReportWrite.Enabled = False
                mnuReportAddOutLine.Enabled = False
                mnuReportModifyOutLine.Enabled = False
                mnuReportDelOutLine.Enabled = False
                mnuReportView.Enabled = False
                mnuReportWriteMuli.Enabled = False
'                mnuReportAgain.Enabled = False
                
                mnuFilePrintBook.Enabled = False
                mnuFilePrintRequest.Enabled = False
                mnuFilePrintList.Enabled = False
                
                mnuViewPatientBrowse.Enabled = False
                    
                mnuMedicalGroupIn.Enabled = False
                mnuMedicalGroupAdd.Enabled = False
                mnuMedicalGroupDelete.Enabled = False
                mnuMedicalItems.Enabled = False
                mnuMedicalItemsAddtion.Enabled = False
                
                mnuChargeMain.Enabled = False
                mnuChargeAdd.Enabled = False
                mnuChargeModify.Enabled = False
                mnuChargeDelete.Enabled = False
            Case 1
                                
                mnuReportWrite.Enabled = False
                mnuReportAddOutLine.Enabled = False
                mnuReportModifyOutLine.Enabled = False
                mnuReportDelOutLine.Enabled = False
                mnuReportView.Enabled = False
                
                mnuViewPatientBrowse.Enabled = False
                    
                mnuMedicalGroupDelete.Enabled = False
'                mnuMedicalItemsAddtion.Enabled = False
                
                mnuChargeMain.Enabled = False
                mnuChargeAdd.Enabled = False
                mnuChargeModify.Enabled = False
                mnuChargeDelete.Enabled = False
            Case 2
                mnuMedicalBeginCancel.Enabled = False
                mnuMedicalComplete.Enabled = False
                
                mnuReportWrite.Enabled = False
                mnuReportAddOutLine.Enabled = False
                mnuReportModifyOutLine.Enabled = False
                mnuReportDelOutLine.Enabled = False
                mnuReportView.Enabled = False
                mnuReportWriteMuli.Enabled = False
                
                mnuFilePrintBook.Enabled = False
                mnuFilePrintRequest.Enabled = False
                mnuFilePrintList.Enabled = False
                
                mnuViewPatientBrowse.Enabled = False
                    
                mnuMedicalGroupIn.Enabled = False
                mnuMedicalGroupAdd.Enabled = False
                mnuMedicalGroupDelete.Enabled = False
                mnuMedicalItems.Enabled = False
'                mnuMedicalItemsAddtion.Enabled = False
                
                mnuChargeMain.Enabled = False
                mnuChargeAdd.Enabled = False
                mnuChargeModify.Enabled = False
                mnuChargeDelete.Enabled = False
            Case 98
            
                mnuMedicalGroupIn.Enabled = False
                mnuMedicalGroupAdd.Enabled = False
                'mnuMedicalItemsAddtion.Enabled = False
                mnuMedicalItems.Enabled = False
                mnuMedicalGroupDelete.Enabled = False
                
                If InStr(mstrPrivilege, "δ�շ����") = 0 Then
                    If picState.Visible = False Then
                        mnuReportWrite.Enabled = False
                        mnuReportWriteMuli.Enabled = False
                        
                        mnuReportAddOutLine.Enabled = False
                        mnuReportModifyOutLine.Enabled = False
                        mnuReportDelOutLine.Enabled = False
                        
'                        mnuReportAgain.Enabled = False
                    End If
                End If
                
            Case 99
                mnuMedicalBeginCancel.Enabled = False
                
                mnuMedicalGroupIn.Enabled = False
                mnuMedicalGroupAdd.Enabled = False
                
                mnuMedicalItems.Enabled = False
        End Select
                
        Select Case tbs.SelectedItem.Key
            Case "����"
                mnuChargeMain.Enabled = False
                mnuChargeAdd.Enabled = False
                mnuChargeModify.Enabled = False
                mnuChargeDelete.Enabled = False
                
                On Error Resume Next
                Select Case Val(mfrmActive.vsf.TextMatrix(mfrmActive.vsf.Row, GetCol(mfrmActive.vsf, "������Դ")))
                Case 1, 2
                    If mfrmActive.vsf.TextMatrix(mfrmActive.vsf.Row, GetCol(mfrmActive.vsf, "ִ��״̬")) = "����ִ��" Then
                        mnuReportWrite.Enabled = False
                        mnuReportView.Enabled = False
                    End If
                End Select
                
                If mfrmActive.vsf.TextMatrix(mfrmActive.vsf.Row, GetCol(mfrmActive.vsf, "[״̬]")) = "" Then
                    mnuReportWrite.Enabled = False
                    mnuReportView.Enabled = False
                End If
                
                If Val(mfrmActive.vsf.TextMatrix(mfrmActive.vsf.Row, GetCol(mfrmActive.vsf, "����id"))) = 0 Then
                    mnuReportView.Enabled = False
                End If
                
                On Error GoTo 0
                
            Case "�ܼ�"
                
                mnuReportWrite.Enabled = False
                mnuReportView.Enabled = False
                mnuChargeMain.Enabled = False
                mnuChargeAdd.Enabled = False
                mnuChargeModify.Enabled = False
                mnuChargeDelete.Enabled = False
                                
            Case "����"
                mnuReportView.Enabled = False
                mnuReportWrite.Enabled = False
                                
            Case "�ſ�"
                
                mnuReportWrite.Enabled = False
                mnuReportView.Enabled = False
                mnuChargeMain.Enabled = False
                mnuChargeAdd.Enabled = False
                mnuChargeModify.Enabled = False
                mnuChargeDelete.Enabled = False
                mnuReportAddOutLine.Enabled = False
'                mnuReportAgain.Enabled = False
                
                For mlngLoop = 1 To mfrmActive.vsf.Rows - 1
                    If Val(mfrmActive.vsf.RowData(mlngLoop)) > 0 Then
                        If Val(mfrmActive.vsf.TextMatrix(mlngLoop, 4)) = 0 Then
                            Exit For
                        End If
                    End If
                Next
                
                If mlngLoop = mfrmActive.vsf.Rows Then mnuMedicalGroupIn.Enabled = False
                
        End Select
        
        If Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "����id"))) > 0 Then
            mnuReportAddOutLine.Enabled = (mlng��첡��id = 0)
        End If
        
        If mnuReportModifyOutLine.Enabled Then
            mnuReportModifyOutLine.Enabled = (mlng��첡��id > 0)
            mnuReportDelOutLine.Enabled = mnuReportModifyOutLine.Enabled
        End If
        
    Case 2
        mnuFilePrintRequest.Enabled = False
        mnuFilePrintList.Enabled = False
        
        mnuMedicalNew.Enabled = False
        mnuMedicalBegin.Enabled = False
        mnuMedicalBeginCancel.Enabled = False
        mnuMedicalComplete.Enabled = False
        mnuReportWrite.Enabled = False
        
        mnuReportAddOutLine.Enabled = False
        mnuReportModifyOutLine.Enabled = False
        mnuReportDelOutLine.Enabled = False
        
        mnuReportWriteMuli.Enabled = False
'        mnuReportAgain.Enabled = False
        
        mnuMedicalGroupIn.Enabled = False
        mnuMedicalGroupAdd.Enabled = False
        mnuMedicalGroupDelete.Enabled = False
        mnuMedicalItems.Enabled = False
        mnuMedicalItemsAddtion.Enabled = False
                        
        mnuChargeMain.Enabled = False
        mnuChargeAdd.Enabled = False
        mnuChargeModify.Enabled = False
        mnuChargeDelete.Enabled = False
        
        If Val(vsf(mintIndex).RowData(vsf(mintIndex).Row)) = 0 Then
            mnuMedicalCompleteCancel.Enabled = False
            mnuFilePrintBook.Enabled = False
            mnuViewPatientBrowse.Enabled = False
        End If
        
        Select Case Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "��־")))
        Case 0, 2
            mnuMedicalCompleteCancel.Enabled = False
            mnuFilePrintBook.Enabled = False
            mnuViewPatientBrowse.Enabled = False
        End Select
        
    Case 3
        mnuFilePrintRequest.Enabled = False
        mnuFilePrintList.Enabled = False
        
        mnuMedicalNew.Enabled = False
        mnuMedicalBegin.Enabled = False
        mnuMedicalBeginCancel.Enabled = False
        mnuMedicalComplete.Enabled = False
        mnuMedicalCompleteCancel.Enabled = False
        mnuReportWrite.Enabled = False
        mnuReportAddOutLine.Enabled = False
        mnuReportModifyOutLine.Enabled = False
        mnuReportDelOutLine.Enabled = False
        mnuReportWriteMuli.Enabled = False
'        mnuReportAgain.Enabled = False
        
        mnuMedicalGroupIn.Enabled = False
        mnuMedicalGroupAdd.Enabled = False
        mnuMedicalGroupDelete.Enabled = False
        mnuMedicalItems.Enabled = False
        mnuMedicalItemsAddtion.Enabled = False
                        
        mnuChargeMain.Enabled = False
        mnuChargeAdd.Enabled = False
        mnuChargeModify.Enabled = False
        mnuChargeDelete.Enabled = False
        
        If Val(vsf(mintIndex).RowData(vsf(mintIndex).Row)) = 0 Then
            mnuFilePrintBook.Enabled = False
            mnuViewPatientBrowse.Enabled = False
        End If
    End Select
    
    If Val(vsf(mintIndex).RowData(vsf(mintIndex).Row)) = 0 Then
        mnuFilePrint.Enabled = False
        mnuFilePrintView.Enabled = False
        mnuFileOutExcel.Enabled = False
    End If
    
    If mintIndex <> 3 Then
        If Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "����id"))) = 0 Then
            mnuViewPatientBrowse.Enabled = False
        End If
    End If
    Select Case tbs.SelectedItem.Key
        Case "����"
                        
            On Error Resume Next
            
            If Val(mfrmActive.vsf.TextMatrix(mfrmActive.vsf.Row, GetCol(mfrmActive.vsf, "����id"))) = 0 Then
                mnuReportView.Enabled = False
            End If
            
            On Error GoTo 0
    Case Else
        mnuReportView.Enabled = False
    End Select
    
    mnuFileSendMail.Enabled = mnuFilePrintBook.Enabled
    mnuMedicalGroupOut.Enabled = mnuMedicalGroupDelete.Enabled
    
    tbrThis.Buttons("Ԥ��").Enabled = mnuFilePrintView.Enabled
    tbrThis.Buttons("��ӡ").Enabled = mnuFilePrint.Enabled
    
    tbrThis.Buttons("����").Enabled = mnuMedicalBegin.Enabled
    tbrThis.Buttons("���").Enabled = mnuMedicalComplete.Enabled
    tbrThis.Buttons("��д").Enabled = mnuReportWrite.Enabled
        
    tbrThis.Buttons("�ܼ�").Enabled = mnuReportAddOutLine.Enabled Or mnuReportModifyOutLine.Enabled Or mnuReportDelOutLine.Enabled
    
    tbrThis.Buttons("����").Enabled = mnuChargeMain.Enabled
    tbrThis.Buttons("����").Enabled = mnuChargeAdd.Enabled Or mnuChargeModify.Enabled Or mnuChargeDelete.Enabled

End Sub

Private Function ShowDeptCase(ByVal lngDept As Long) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����;
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim intIndex As Integer
    
    On Error GoTo errHand
    
    
    ShowDeptCase = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then Resume
End Function

Private Sub RefreshStateInfo()
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�ˢ��״̬����ʾ��Ϣ
    '------------------------------------------------------------------------------------------------------------------
    Dim strInfo As String
    Dim lngIndex As Long
    Dim lngLoop As Long
    Dim lngCount(0 To 3) As Long
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error Resume Next
    
    
    
    strSQL = GetPublicSQL(SQL.�������ͳ��, "2'" & vsf(0).Tag & "'0")
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, 2, CDate(Split(vsf(0).Tag, "|")(0)), CDate(Split(vsf(0).Tag, "|")(1)), 0)
    If rs.BOF = False Then lngCount(0) = rs.Fields(0).Value

    
    strSQL = GetPublicSQL(SQL.�������ͳ��, "4'" & vsf(1).Tag & "'1")
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, 4, CDate(Split(vsf(1).Tag, "|")(0)), CDate(Split(vsf(1).Tag, "|")(1)), 1)
    If rs.BOF = False Then lngCount(1) = rs.Fields(0).Value
    
    strSQL = GetPublicSQL(SQL.�������ͳ��, "5'" & vsf(2).Tag & "'1")
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, 5, CDate(Split(vsf(2).Tag, "|")(0)), CDate(Split(vsf(2).Tag, "|")(1)), 1)
    If rs.BOF = False Then lngCount(2) = rs.Fields(0).Value

    
    Select Case mintIndex
    Case 0
        
        If lngCount(mintIndex) = 0 Then
            strInfo = strInfo & "û�еȴ�����졣"
        Else
            strInfo = strInfo & "��" & lngCount(mintIndex) & "���ȴ�����졣"
        End If
    Case 1
        If lngCount(mintIndex) = 0 Then
            strInfo = strInfo & "û����������졣"
        Else
            strInfo = strInfo & "��" & lngCount(mintIndex) & "����������졣"
        End If
    Case 2
        If lngCount(mintIndex) = 0 Then
            strInfo = strInfo & "û������ɵ���졣"
        Else
            strInfo = strInfo & "��" & lngCount(mintIndex) & "������ɵ���졣"
        End If
    Case 3
        If vsf(mintIndex).Rows = 2 And Val(vsf(mintIndex).RowData(vsf(mintIndex).Row)) = 0 Then
            strInfo = "û�в�ѯ������Ҫ�����졣"
        Else
            strInfo = "����ѯ��" & vsf(mintIndex).Rows - 1 & "������Ҫ�����졣"
        End If
    End Select
 
    cmdKind(0).Caption = "&A.�ȴ����(" & Lpad(lngCount(0), 4, " ") & " ��)"
    cmdKind(1).Caption = "&B.�������(" & Lpad(lngCount(1), 4, " ") & " ��)"
    cmdKind(2).Caption = "&C.������(" & Lpad(lngCount(2), 4, " ") & " ��)"
    
    stbThis.Panels(2).Text = strInfo
End Sub

Private Function SaveRow(ByVal objVsf As Object) As String
    SaveRow = objVsf.RowData(objVsf.Row)
End Function

Private Sub InheritRestoreRow(ByVal objVsf As Object, ByVal strKey As String)
    '--------------------------------------------------------------------------------------------------------
    '����:�̳�RestoreRow����
    '����:
    '����:
    '--------------------------------------------------------------------------------------------------------
    Dim int����idCol As Integer
    Dim intRow As Integer
    
    On Error Resume Next
        
    Call RestoreRow(objVsf, Val(strKey))
    
    int����idCol = GetCol(objVsf, "����id")
    
    If int����idCol > 0 Then
        For intRow = objVsf.Row To 1 Step -1
            If Val(objVsf.TextMatrix(intRow, int����idCol)) <= 0 Then
                objVsf.IsCollapsed(objVsf.Row - 1) = flexOutlineExpanded
                Exit For
            End If
        Next
    End If

End Sub

Private Function CancelMedical(ByVal str���� As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�ȡ����ʼ���
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim strSQL() As String
    Dim blnTran As Boolean
    Dim lngLoop As Long
    
    On Error GoTo errHand
    
    ReDim Preserve strSQL(1 To 1)
    
    gstrSQL = "SELECT ID FROM ����ҽ����¼ WHERE ���ID IS NULL AND ������Դ=4 AND �Һŵ�='" & str���� & "'"
    Call zlDatabase.OpenRecordset(rs, gstrSQL, Me.Caption)
    If rs.BOF Then Exit Function
    
    Do While Not rs.EOF
        
        '���ﲡ�����ϵ�ͬʱҲ����
        strSQL(ReDimArray(strSQL)) = "ZL_���ǼǼ�¼_Cancel(" & rs("ID").Value & ")"
                
        rs.MoveNext
    Loop
    
    blnTran = True
    gcnOracle.BeginTrans
    
    For lngLoop = 1 To UBound(strSQL)
        If strSQL(lngLoop) <> "" Then Call zlDatabase.ExecuteProcedure(strSQL(lngLoop), Me.Caption)
    Next
    
    gcnOracle.CommitTrans
    blnTran = False
    
    CancelMedical = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then Resume
    If blnTran Then gcnOracle.RollbackTrans
    
End Function

Private Function CheckExecuteState(ByVal str���� As String, ByVal lng����id As Long) As Byte
    
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ����ִ��״̬
    '����:  1                   ��ʾ������ִ�е���Ŀ
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    
    If lng����id = 0 Then
        '���岡��
        gstrSQL = " IN (SELECT A.����id FROM �����Ա���� A,���ǼǼ�¼ B WHERE A.�Ǽ�id=B.ID AND B.����=[1])"
    Else
        '��������
        gstrSQL = "=[2]"
    End If
    
    gstrSQL = "SELECT 1 FROM ����ҽ������ WHERE ִ��״̬ = 3 AND  ҽ��id IN (select ID from ����ҽ����¼ where ����id " & gstrSQL & " and ������Դ = 4 and �Һŵ� = [1] and ҽ��״̬ <> 4)"
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str����, lng����id)
    
    CheckExecuteState = IIf(rs.BOF, 0, 1)
    
End Function

Private Function MenuClick(ByVal strMenuItem As String, Optional ByVal lng�ļ�����id As Long = 0) As Boolean
    '******************************************************************************************************************
    '���ܣ����ݱ༭/����
    '******************************************************************************************************************
    Dim lngKey As Long
    
    Dim blnTran As Boolean
    Dim lngLoop As Long
    Dim strSQL As String
    Dim lng�Ǽ�id As Long
    Dim str���� As String
    Dim lng����id As Long
    Dim strPrompt As String
    Dim blnGroup As Boolean
    Dim strGroup As String
    Dim lngStop As Long
    Dim rsItems As New ADODB.Recordset
    Dim rs As New ADODB.Recordset
    Dim rsNo As New ADODB.Recordset
    Dim rsSQL As New ADODB.Recordset
    Dim lngTmp As Long
    Dim strTmp As String
    
    On Error GoTo errHand
    
    AutoRefresh = False

    Call SQLRecord(rsSQL)
        
    '�˳�����(һ�����˵�����״̬Ӧ���������˵�)
    Select Case strMenuItem
    '------------------------------------------------------------------------------------------------------------------
    Case "��������", "������"
        
        '�޴���
    '------------------------------------------------------------------------------------------------------------------
    Case Else
        lngKey = Val(vsf(mintIndex).RowData(vsf(mintIndex).Row))
        If lngKey = 0 Then GoTo pointEnd
        
        '��ȡ����Ҫ�õĽ��ֵ
        lng�Ǽ�id = Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "�Ǽ�id")))
        str���� = vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "��쵥��"))
        If mintIndex <> 3 Then
            If vsf(mintIndex).Cell(flexcpData, vsf(mintIndex).Row, 0) = "" Then
                '���������Ա
                lng����id = Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "����id")))
            Else
                
                '������ʾ��
                If mintIndex = 0 Then
                    If vsf(mintIndex).Row + 1 > vsf(mintIndex).Rows - 1 Then
                        ShowSimpleMsg "��ǰ����û�����������Ա��"
                        GoTo pointEnd
                    End If
                End If
            End If
        
            blnGroup = False
            blnGroup = (Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "����id"))) = 0)
            If blnGroup = False Then
                blnGroup = (Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "��־"))) = 1)
            End If
        End If
    End Select
    
    '��һ������
    Select Case strMenuItem
    '------------------------------------------------------------------------------------------------------------------
    Case "���Ǽ�"
        
        Select Case lng�ļ�����id
        Case 0      '��������
            
            If Not frmScheduleEdit.ShowEdit(Me, 0, mlngDept, , 2) Then GoTo pointEnd
            
        Case 1      '��������
            
            If Not frmScheduleEdit.ShowEdit(Me, 0, mlngDept, True, 2) Then GoTo pointEnd
            
        Case 3      '�޸ĵǼ�
            
            If lng�Ǽ�id = 0 Then Exit Function
            
            If blnGroup = False Then
                blnGroup = (Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "�ϼ�id"))) > 0)
            End If
            
            If Not frmScheduleEdit.ShowEdit(Me, lng�Ǽ�id, mlngDept, blnGroup, 2) Then GoTo pointEnd
            
        Case 4      'ɾ���Ǽ�
            
            If lng�Ǽ�id = 0 Then Exit Function
            
            If MsgBox("�����Ҫɾ����ǰ���Ǽ���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                    
            strSQL = "ZL_���ǼǼ�¼_DELETE(" & lng�Ǽ�id & ")"
            Call SQLRecordAdd(rsSQL, strSQL)
            
        End Select
    '------------------------------------------------------------------------------------------------------------------
    Case "��Ƭ�ɼ�"
        
        If lng�Ǽ�id > 0 Then
            Call frmPersonPhoto.ShowEdit(Me, lng�Ǽ�id, lng����id)
            GoTo pointEnd
        End If
     '------------------------------------------------------------------------------------------------------------------
    Case "���ָ����", "��Ŀ���뵥"
        
        If lng�Ǽ�id > 0 Then
            Call frmMedicalStationPrintRpt.ShowEdit(Me, lng�Ǽ�id, lng����id, strMenuItem)
        End If
        
        GoTo pointEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "������뵥"
    
        If lng�Ǽ�id > 0 Then
            Call frmMedicalStationRequest.ShowEdit(Me, lng�Ǽ�id, lng����id)
        End If
        
        GoTo pointEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "��챨�浥"
    
        If lng�Ǽ�id > 0 Then
            Call frmMedicalStationRptBook.ShowEdit(Me, lng�Ǽ�id, lng����id)
        End If
        
        GoTo pointEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "�����ʼ�"
        
        If lng�Ǽ�id = 0 Then GoTo pointEnd
        
        If vsf(mintIndex).Cell(flexcpData, vsf(mintIndex).Row, 0) <> "" Then
            Call frmMedicalStationSendMail.ShowEdit(Me, lng�Ǽ�id)
        Else
            Call frmMedicalStationSendMail.ShowEdit(Me, lng�Ǽ�id, lng����id)
        End If
        
        GoTo pointEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "��������"
    
        If Not frmMedicalStationPara.ShowPara(Me) Then GoTo pointEnd
    
            Dim strStart As String
        Dim strEnd As String
        
        strStart = GetDateTime(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "�����ʱ�䷶Χ", "��  ��"), 1)
        strEnd = GetDateTime(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "�����ʱ�䷶Χ", "��  ��"), 2)
        If strStart = "" Then strStart = GetDateTime("��  ��", 1)
        If strEnd = "" Then strEnd = GetDateTime("��  ��", 2)
        vsf(0).Tag = strStart & "|" & strEnd
        
        strStart = GetDateTime(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "�����ʱ�䷶Χ", "��  ��"), 1)
        strEnd = GetDateTime(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "�����ʱ�䷶Χ", "��  ��"), 2)
        If strStart = "" Then strStart = GetDateTime("��  ��", 1)
        If strEnd = "" Then strEnd = GetDateTime("��  ��", 2)
        vsf(1).Tag = strStart & "|" & strEnd
        
        '����ȱʡʱ�䷶Χ
        strStart = GetDateTime(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "���������ʱ�䷶Χ", "��  ��"), 1)
        strEnd = GetDateTime(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "���������ʱ�䷶Χ", "��  ��"), 2)
        If strStart = "" Then strStart = GetDateTime("��  ��", 1)
        If strEnd = "" Then strEnd = GetDateTime("��  ��", 2)
        mstr���������ʱ�䷶Χ = strStart & "|" & strEnd
        
        '�����������
        mint������ѯ���� = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "������ѯ����", "0"))
                
        strStart = GetDateTime(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "�������ʱ�䷶Χ", "��  ��"), 1)
        strEnd = GetDateTime(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "�������ʱ�䷶Χ", "��  ��"), 2)
        If strStart = "" Then strStart = GetDateTime("��  ��", 1)
        If strEnd = "" Then strEnd = GetDateTime("��  ��", 2)
        vsf(2).Tag = strStart & "|" & strEnd
        
    '------------------------------------------------------------------------------------------------------------------
    Case "������"
        
        strTmp = vsf(0).Tag & "'" & vsf(1).Tag & "'" & vsf(2).Tag & "'" & mstr���������ʱ�䷶Χ & "'" & mint������ѯ����
        If frmMedicalStationSearch.ShowFilter(Me, strTmp) = False Then GoTo pointEnd
        
        vsf(0).Tag = Split(strTmp, "'")(0)
        vsf(1).Tag = Split(strTmp, "'")(1)
        vsf(2).Tag = Split(strTmp, "'")(2)
        
        mstr���������ʱ�䷶Χ = Split(strTmp, "'")(3)
        mint������ѯ���� = Val(Split(strTmp, "'")(4))
        
    '------------------------------------------------------------------------------------------------------------------
    Case "�������"             '�������
                        
        Select Case CheckAllowMedical(lng�Ǽ�id)
        Case 1
            strPrompt = "��ǰ��컹û������������壡"
        Case 2
            strPrompt = "��ǰ��컹û�����������Ա��"
        Case 3
            strPrompt = "��ǰ���������Ŀ��������ÿ���������������Ŀ����"
        Case 4
            strPrompt = "����û�з���������Ա��������ԤԼ�����н�����Ա��𻮷֣�"
        End Select
        
        If strPrompt <> "" Then
            ShowSimpleMsg strPrompt
            GoTo pointEnd
        End If
        
        If vsf(mintIndex).Cell(flexcpData, vsf(mintIndex).Row, 0) <> "" Then
            '����
            If Not frmMedicalStationBegin.ShowEdit(Me, lng�Ǽ�id) Then GoTo pointEnd
        Else
            If Not frmMedicalStationBegin.ShowEdit(Me, lng�Ǽ�id, lng����id) Then GoTo pointEnd
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "��Ա����"
        
        If str���� = "" Then GoTo pointEnd
        
        If blnGroup Then
            '����
            If Not frmMedicalStationBegin.ShowEdit(Me, lng�Ǽ�id, , True) Then GoTo pointEnd
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "ȡ����ʼ"
        If str���� = "" Then GoTo pointEnd
        
        If MsgBox("ȷ��Ҫȡ����ǰ���ڵ������" & vbCrLf & "����и�����Ŀ�������¿�ʼ������Ҫ������ӡ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then GoTo pointEnd
        strSQL = "ZL_���ǼǼ�¼_Cancel('" & str���� & "')"
        Call SQLRecordAdd(rsSQL, strSQL)
    '------------------------------------------------------------------------------------------------------------------
    Case "��д����"
        
        Call mfrmActive.zlMenuClick(Me, "��д����", CStr(lngKey) & "'1")
        GoTo pointEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "�鿴����"
        
        Call mfrmActive.zlMenuClick(Me, "�鿴����", CStr(lngKey) & "'1")
        GoTo pointEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "��д�ܽ�"             '
        If lng����id = 0 Or lng�Ǽ�id = 0 Then GoTo pointEnd
        
        If mlng��첡��id > 0 Then GoTo pointEnd
        
        mlng��첡��id = EditPatientFile("", lng����id, str����, 0, lng�ļ�����id, False, Me, , True, 2, 1)
        If mlng��첡��id = 0 Then GoTo pointEnd
        
        strSQL = "ZL_�����Ա����_�ܽ�(" & lng�Ǽ�id & "," & lng����id & "," & mlng��첡��id & ")"
        Call SQLRecordAdd(rsSQL, strSQL)
    '------------------------------------------------------------------------------------------------------------------
    Case "�޸��ܽ�"
        
        If mlng��첡��id = 0 Then GoTo pointEnd
        
        Call EditPatientFile(mlng��첡��id, lng����id, str����, 0, , False, Me, , True, 2, 1)
    '------------------------------------------------------------------------------------------------------------------
    Case "ɾ���ܽ�"
                
        If mlng��첡��id = 0 Then GoTo pointEnd
        
        If MsgBox("�Ƿ�ɾ������Ա������ܽ᣿", vbDefaultButton2 + vbQuestion + vbYesNo, gstrSysName) = vbNo Then GoTo pointEnd
        strSQL = "ZL_�����Ա����_�ܽ�(" & lng�Ǽ�id & "," & lng����id & ",null)"
        Call SQLRecordAdd(rsSQL, strSQL)
        
        strSQL = "zl_���˲���_DELETE(" & mlng��첡��id & ")"
        Call SQLRecordAdd(rsSQL, strSQL)
    '------------------------------------------------------------------------------------------------------------------
    Case "������д"
        
        If Not frmMedicalStationAdjust.ShowEdit(Me, mstrPrivilege) Then GoTo pointEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "������"
        
        If blnGroup Then

            If str���� = "" Then GoTo pointEnd
            
            '����Ƿ���������ɵĻ�û����д�����
            If CheckExecuteState(str����, 0) = 1 Then
                ShowSimpleMsg "������������Ա��������ִ�е���Ŀ��"
                GoTo pointEnd
            End If
            
            If MsgBox("��ǰ�����Ա����춼�������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then GoTo pointEnd
            
            strSQL = "ZL_���ǼǼ�¼_Finish('" & str���� & "',0)"
            Call SQLRecordAdd(rsSQL, strSQL)
        Else
            If str���� = "" Or lng����id = 0 Then GoTo pointEnd
                                    
            '����Ƿ���������ɵĻ�û����д�����
            If CheckExecuteState(str����, lng����id) = 1 Then
                ShowSimpleMsg "�������Ա��������ִ�е���Ŀ��"
                GoTo pointEnd
            End If
            
            If MsgBox("��ǰ�����Ա����춼�������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then GoTo pointEnd
            
            strSQL = "ZL_���ǼǼ�¼_Finish('" & str���� & "'," & lng����id & ")"
            Call SQLRecordAdd(rsSQL, strSQL)
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "ȡ�����"
                
        If blnGroup Then
            If str���� = "" Then GoTo pointEnd
                        
            If MsgBox("���Ҫȡ����ǰ��������ɵ���죿", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then GoTo pointEnd
            
            strSQL = "ZL_���ǼǼ�¼_CancelFinish('" & str���� & "',0)"
            Call SQLRecordAdd(rsSQL, strSQL)
        Else
            If str���� = "" Or lng����id = 0 Then GoTo pointEnd
   
            If MsgBox("���Ҫȡ����ǰ��Ա����ɵ���죿", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then GoTo pointEnd
            
            strSQL = "ZL_���ǼǼ�¼_CancelFinish('" & str���� & "'," & lng����id & ")"
            Call SQLRecordAdd(rsSQL, strSQL)
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "������Ŀ"
        lngTmp = Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "��־")))
        
        Select Case lngTmp
        Case 1, 2
            With mfrmActive.Body
                lng����id = Val(.RowData(.Row))
                lngTmp = Abs(Val(.TextMatrix(.Row, 4)))
            End With
        Case Else
            lngTmp = 1
        End Select
        If str���� = "" Or lng����id = 0 Or lng�Ǽ�id = 0 Then GoTo pointEnd
        
        Dim rsData As New ADODB.Recordset
        
        gstrSQL = GetPublicSQL(SQL.��Աԭʼ��Ŀ)
        Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng�Ǽ�id, lng����id)
        If ShowGrdFilter(Me, vsf(mintIndex), "����,2700,0,0;���,900,0,1;ִ�п���,1500,0,0;�ɼ���ʽ,1200,0,0;����걾,1200,0,0;��鲿λ,1200,0,0", Me.Name & "\������Ŀѡ��", "����б���ѡ��Ҫ����������Ŀ��", rsData, rs, 8790, 4500, , , True) Then
            If rs.RecordCount > 0 Then
                
                '������Ŀ��������Ŀ
                Call InsertItems(rsSQL, rs, lng�Ǽ�id, lng����id, True)

            End If
        End If
            
    '------------------------------------------------------------------------------------------------------------------
    Case "������Ŀ"
        
        lngTmp = Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "��־")))
        
        Select Case lngTmp
        Case 1, 2
            With mfrmActive.Body
                lng����id = Val(.RowData(.Row))
                lngTmp = Abs(Val(.TextMatrix(.Row, 4)))
            End With
        Case Else
            lngTmp = 1
        End Select
        If str���� = "" Or lng����id = 0 Or lng�Ǽ�id = 0 Then GoTo pointEnd

        Call MedicalItemsRecord(rsItems)
        
        gstrSQL = GetPublicSQL(SQL.��Ա�����Ŀ)
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng�Ǽ�id, lng����id)
        Call WriteItems(rs, rsItems, 2)
       
        Select Case lngTmp
        Case 0
            If Not frmItemsEdit.ShowEdit(Me, lng�Ǽ�id, rsItems, mlngDept, False, 1, lng����id) Then GoTo pointEnd
            '�����Ѿ�ɾ���������Ŀ
            Call FilterRecord(rsItems, "ɾ��='1'")
            Call DeleteItem(rsSQL, rsItems, str����, lng�Ǽ�id, lng����id)
    
            '��������ӵ������Ŀ
            Call FilterRecord(rsItems, "�¼�<>'1'")
            Call NewItem(rsSQL, rsItems, lng�Ǽ�id, lng����id)

        Case Else
            If Not frmItemsEdit.ShowEdit(Me, lng�Ǽ�id, rsItems, mlngDept, False, 2, lng����id) Then GoTo pointEnd
        
            '�����Ѿ�ɾ���������Ŀ
            Call FilterRecord(rsItems, "ɾ��='1'")
            Call DeleteItems(rsSQL, rsItems, str����, lng�Ǽ�id, lng����id)
            
            '��������ӵ������Ŀ
            Call FilterRecord(rsItems, "�¼�<>'1'")
            Call InsertItems(rsSQL, rsItems, lng�Ǽ�id, lng����id)
        End Select
        
    '------------------------------------------------------------------------------------------------------------------
    Case "�����Ŀ"

        If str���� = "" Or lng�Ǽ�id = 0 Then GoTo pointEnd
        
        Call MedicalItemsRecord(rsItems)
        
        '��ȡ�����Ŀ
        gstrSQL = GetPublicSQL(SQL.���������Ŀ)
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng�Ǽ�id)
        
        Call WriteItems(rs, rsItems, 1)
        If Not frmItemsEdit.ShowEdit(Me, lng�Ǽ�id, rsItems, mlngDept, blnGroup, 2) Then GoTo pointEnd
            
        '�����Ѿ�ɾ���������Ŀ
        Call FilterRecord(rsItems, "ɾ��='1'")
        Call DeleteItems(rsSQL, rsItems, str����, lng�Ǽ�id)
                
        '��������ӵ������Ŀ
        Call FilterRecord(rsItems, "�¼�<>'1'")
        Call InsertItems(rsSQL, rsItems, lng�Ǽ�id)
    '------------------------------------------------------------------------------------------------------------------
    Case "��ӳ�Ա"
    
        If str���� = "" Or lng�Ǽ�id = 0 Then GoTo pointEnd
        
        Dim intCount2 As Integer
        Dim str����� As String
        
        Call MedicalItemsRecord(rsItems, 2)
        
        gstrSQL = GetPublicSQL(SQL.�����Ա����)
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng�Ǽ�id)
        If WriteItems(rs, rsItems, 1, 2) = False Then Exit Function
        
        gstrSQL = "Select ��Լ��λid From ���ǼǼ�¼ Where ID=[1]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng�Ǽ�id)
        If rs.BOF Then Exit Function
        
        If Not frmPersonEdit.ShowEdit(Me, lng�Ǽ�id, rsItems, True, 2, zlCommFun.NVL(rs("��Լ��λid"), 0)) Then Exit Function
        
        '��������ӵ������Ա
        Call FilterRecord(rsItems, "�¼�<>'1'")
        If rsItems.RecordCount > 0 Then rsItems.MoveFirst
        
        Dim intCount As Integer
        Dim intCount1 As Integer
        Dim bytNew As Byte
        Dim lngCount As Long
        
        intCount = -1
        Do While Not rsItems.EOF
            
            Call SQLRecord(rsSQL)
            
            '����������
            If rsItems("��������") <> "" Then
                
                If CheckStrValid(rsItems("��������"), CHECKFORMAT.����) = False Then
                    ShowSimpleMsg rsItems("����").Value & "�ĳ���������Ч��"
                    Exit Function
                End If
            End If
            bytNew = 0
            lng����id = rsItems("����ID").Value
            If lng����id = 0 Then
                bytNew = 1
                intCount = intCount + 1
                lng����id = GetNextNo(1) + intCount
                
                rsItems("����ID").Value = lng����id
            End If
            
            intCount1 = intCount1 + 1
            
            If zlCommFun.NVL(rsItems("�����").Value, 0) < 1 Then
                'lng����� = GetNextNo(3) + intCount2
                str����� = CStr(GetNextNo(3) + intCount2)
                intCount2 = intCount2 + 1
            Else
                str����� = CStr(zlCommFun.NVL(rsItems("�����").Value, 0))
            End If
                        
            strSQL = "ZL_�����Ա����_INSERT(" & lng�Ǽ�id & "," & _
                                                                IIf(lng����id = 0, "NULL", lng����id) & ",'" & _
                                                                rsItems("���").Value & "','" & _
                                                                rsItems("����").Value & "','" & _
                                                                rsItems("���֤").Value & "','" & _
                                                                rsItems("�Ա�").Value & "'," & _
                                                                IIf(rsItems("��������").Value = "", "NULL", "TO_DATE('" & rsItems("��������").Value & "','yyyy-mm-dd hh24:mi:ss')") & ",'" & _
                                                                rsItems("����״��").Value & "','" & _
                                                                rsItems("����").Value & "','" & _
                                                                rsItems("����").Value & "','" & _
                                                                rsItems("ѧ��").Value & "','" & _
                                                                rsItems("ְҵ").Value & "','" & _
                                                                rsItems("��ϵ������").Value & "','" & _
                                                                rsItems("��ϵ�˵绰").Value & "','" & _
                                                                rsItems("�����ʼ�").Value & "','" & _
                                                                rsItems("��ϵ�˵�ַ").Value & "','" & _
                                                                rsItems("������λ").Value & "','" & _
                                                                rsItems("����").Value & "'," & _
                                                                Val(str�����) & ",'" & _
                                                                rsItems("IC����").Value & "','" & _
                                                                rsItems("������").Value & "','" & _
                                                                rsItems("���￨��").Value & "',0,0,0," & bytNew & _
                                                                ",Null)"
            
            Call SQLRecordAdd(rsSQL, strSQL)

            Dim lngSendNo As Long
            Dim str�ɼ�No As String
            Dim strNO As String
            
            lngSendNo = GetNextNo(10)
                        
            '�������õ��ݺ�
            strSQL = "Select b.ID,b.����;��,b.�ɼ���ʽid From �����Ŀ�嵥 b Where b.�������=[1] and b.�Ǽ�id=[2]"
            Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, rsItems("���").Value, lng�Ǽ�id)
            If rs.BOF = False Then
                Do While Not rs.EOF
                    
                    str�ɼ�No = ""
                    strNO = ""
                    
                    If zlCommFun.NVL(rs("����;��").Value, 1) = 1 Then
                        '����
                        strNO = GetNextNo(14)
                    Else
                        strNO = GetNextNo(13)
                    End If
                    
                    If zlCommFun.NVL(rs("�ɼ���ʽid").Value, 0) > 0 Then
                        '�ɼ�
                        If zlCommFun.NVL(rs("����;��").Value, 1) = 1 Then
                            '����
                            str�ɼ�No = GetNextNo(14)
                        Else
                            str�ɼ�No = GetNextNo(13)
                        End If
                    End If
                    
                    strSQL = "ZL_�����Ŀҽ��_NO(" & zlCommFun.NVL(rs("ID").Value, 0) & "," & lng����id & ",'" & strNO & "','" & str�ɼ�No & "')"
                    Call SQLRecordAdd(rsSQL, strSQL)
                    
                    rs.MoveNext
                Loop
            End If
            
            blnTran = True
            gcnOracle.BeginTrans
            If rsSQL.RecordCount > 0 Then rsSQL.MoveFirst
            For lngCount = 1 To rsSQL.RecordCount
                Call zlDatabase.ExecuteProcedure(CStr(rsSQL("SQL").Value), Me.Caption)
                rsSQL.MoveNext
            Next
            Call SQLRecord(rsSQL)
    
            strSQL = "zl_�����Ա����_Accept(" & lng�Ǽ�id & "," & lngSendNo & "," & lng����id & "," & mlngDept & ",NULL,1)"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            
            '������ط���
            If MakeMedicalCharge(rsSQL, lng�Ǽ�id) = False Then
                gcnOracle.RollbackTrans
                blnTran = False
                Exit Function
            End If
            
            strSQL = "zl_�����Ա����_Accept(" & lng�Ǽ�id & "," & lngSendNo & "," & lng����id & "," & mlngDept & ",NULL,2)"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            
            gcnOracle.CommitTrans
            blnTran = False
            
            
            rsItems.MoveNext
        Loop
    '------------------------------------------------------------------------------------------------------------------
    Case "�Ƴ���Ա"
        
        If lng����id = 0 Or lng�Ǽ�id = 0 Then GoTo pointEnd

        If MsgBox("�Ƴ��������Ա��ͬʱҲ��������챨�棬ȷ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then GoTo pointEnd

        strSQL = "ZL_�����Ա����_DELETE(" & lng�Ǽ�id & "," & lng����id & ",1)"
        Call SQLRecordAdd(rsSQL, strSQL)
    '------------------------------------------------------------------------------------------------------------------
    Case "ȡ������"
    
        If lng����id = 0 Or lng�Ǽ�id = 0 Then GoTo pointEnd

        If MsgBox("ȡ����Ա������ͬʱ������챨�棬ȷ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then GoTo pointEnd

        strSQL = "ZL_�����Ա����_DELETE(" & lng�Ǽ�id & "," & lng����id & ",1,1)"
        Call SQLRecordAdd(rsSQL, strSQL)
    '------------------------------------------------------------------------------------------------------------------
    Case "ִ�е���"
        '
        If Not frmMedicalStationDept.ShowEdit(Me, lng�Ǽ�id) Then GoTo pointEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "��Ա��Ϣ"
        If lng����id = 0 Then GoTo pointEnd
        
        Dim strParam As String
        Dim varParam As Variant
        
        gstrSQL = GetPublicSQL(SQL.�����Ա����_����)
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng�Ǽ�id, lng����id)
        
        If rs.BOF = False Then
            strParam = zlCommFun.NVL(rs("����id").Value) & "'"
            strParam = strParam & zlCommFun.NVL(rs("����").Value) & "'"
            strParam = strParam & zlCommFun.NVL(rs("���֤").Value) & "'"
            strParam = strParam & zlCommFun.NVL(rs("�Ա�").Value) & "'"
            strParam = strParam & zlCommFun.NVL(rs("��������").Value) & "'"
            strParam = strParam & zlCommFun.NVL(rs("����״��").Value) & "'"
            strParam = strParam & zlCommFun.NVL(rs("����").Value) & "'"
            strParam = strParam & zlCommFun.NVL(rs("����").Value) & "'"
            strParam = strParam & zlCommFun.NVL(rs("ѧ��").Value) & "'"
            strParam = strParam & zlCommFun.NVL(rs("ְҵ").Value) & "'"
            strParam = strParam & zlCommFun.NVL(rs("���").Value) & "'"
            strParam = strParam & zlCommFun.NVL(rs("��ϵ������").Value) & "'"
            strParam = strParam & zlCommFun.NVL(rs("��ϵ�˵绰").Value) & "'"
            strParam = strParam & zlCommFun.NVL(rs("�����ʼ�").Value) & "'"
            strParam = strParam & zlCommFun.NVL(rs("��ϵ�˵�ַ").Value) & "'"
            strParam = strParam & zlCommFun.NVL(rs("������λ").Value) & "'"
            strParam = strParam & zlCommFun.NVL(rs("����").Value) & "'"
            strParam = strParam & zlCommFun.NVL(rs("������").Value)
                        
            If frmPatientEdit.ShowEdit(Me, strParam, (mintIndex = 1)) Then
                
                If mintIndex = 1 Then
                    varParam = Split(strParam, "'")
                    
                    strSQL = "ZL_�����Ա����_INSERT(" & lng�Ǽ�id & "," & _
                                                    Val(varParam(0)) & "," & _
                                                    "NULL,'" & _
                                                    varParam(1) & "','" & _
                                                    varParam(2) & "','" & _
                                                    varParam(3) & "'," & _
                                                    IIf(varParam(4) = "", "NULL", "TO_DATE('" & varParam(4) & "','yyyy-mm-dd hh24:mi:ss')") & ",'" & _
                                                    varParam(5) & "','" & _
                                                    varParam(6) & "','" & _
                                                    varParam(7) & "','" & _
                                                    varParam(8) & "','" & _
                                                    varParam(9) & "','" & _
                                                    varParam(11) & "','" & _
                                                    varParam(12) & "','" & _
                                                    varParam(13) & "','" & _
                                                    varParam(14) & "','" & _
                                                    varParam(15) & "','" & _
                                                    varParam(16) & "'," & _
                                                    "NULL," & _
                                                    "NULL,'" & _
                                                    varParam(17) & "'," & _
                                                    "NULL," & _
                                                    "1,0,1,0,Null)"
                    Call SQLRecordAdd(rsSQL, strSQL)
                End If
            End If
            
        End If
        
    '------------------------------------------------------------------------------------------------------------------
    Case "��������"
                
        Call mfrmActive.zlMenuClick(Me, lngKey, "����������")
        GoTo pointEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "�����շѵ���"
                
        Call mfrmActive.zlMenuClick(Me, lngKey, "�շѵ���")
        GoTo pointEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "���Ӽ��ʵ���"
            
        Call mfrmActive.zlMenuClick(Me, lngKey, "���ʵ���")
        GoTo pointEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "������Ѻ��õǼ�"
        
        Call mfrmActive.zlMenuClick(Me, lngKey, "��Ѻ��õǼ�")
        GoTo pointEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "�޸ĸ���"
        
        Call mfrmActive.zlMenuClick(Me, lngKey, "�޸ĸ��ӷ���")
        GoTo pointEnd
    '------------------------------------------------------------------------------------------------------------------
    Case "ɾ������"
                
        Call mfrmActive.zlMenuClick(Me, lngKey, "ɾ�����ӷ���")
        GoTo pointEnd
    End Select
    
    '�ڶ�������
    
    blnTran = True
    
    gcnOracle.BeginTrans
    
    If rsSQL.RecordCount > 0 Then
        zlCommFun.ShowFlash "���ڴ���...", Me
        DoEvents
        rsSQL.MoveFirst
    End If
    
    For lngLoop = 1 To rsSQL.RecordCount
                    
        If lngLoop > lngStop And lngStop > 0 Then
            lngStop = 0
            '����ͣ��һ����,����ᵼ�²���ҽ��״̬�����ظ�(ҽ��ID,����ʱ��)
            Sleep 1006
        End If
        
        Call zlDatabase.ExecuteProcedure(CStr(rsSQL("SQL").Value), Me.Caption)
        rsSQL.MoveNext
    Next
    gcnOracle.CommitTrans
    
    zlCommFun.StopFlash
    DoEvents
        
    blnTran = False
    
    If strMenuItem = "ɾ���ܽ�" Then mlng��첡��id = 0
    
    'ˢ�´���
    Select Case strMenuItem
    '------------------------------------------------------------------------------------------------------------------
    Case "ȡ�����"
        
        If CheckState(lng�Ǽ�id, 2) And mintIndex = 2 Then
            Call mnuViewRefresh_Click
        Else
            Set vsf(mintIndex).Cell(flexcpPicture, vsf(mintIndex).Row, GetCol(vsf(mintIndex), "[״̬]")) = ils13.ListImages("��ʼ").Picture
            vsf(mintIndex).Cell(flexcpText, vsf(mintIndex).Row, GetCol(vsf(mintIndex), "[״̬]")) = "��ʼ"
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "������"
        
        '����Ƿ�ȫ�����
        If CheckState(lng�Ǽ�id) Then
            Call mnuViewRefresh_Click
        Else
            Set vsf(mintIndex).Cell(flexcpPicture, vsf(mintIndex).Row, GetCol(vsf(mintIndex), "[״̬]")) = ils13.ListImages("���").Picture
            vsf(mintIndex).Cell(flexcpText, vsf(mintIndex).Row, GetCol(vsf(mintIndex), "[״̬]")) = "���"
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case "ȡ������", "�Ƴ���Ա"
        
        vsf(mintIndex).RemoveItem vsf(mintIndex).Row
        Call vsf_AfterRowColChange(mintIndex, 0, 0, vsf(mintIndex).Row, vsf(mintIndex).Col)
    '------------------------------------------------------------------------------------------------------------------
    Case "�����Ŀ", "������Ŀ", "������д", "��д����"
    
        Call vsf_AfterRowColChange(mintIndex, 0, 0, vsf(mintIndex).Row, vsf(mintIndex).Col)
    '------------------------------------------------------------------------------------------------------------------
    Case "��츴��"
    
        Call RefreshData("����")
    '------------------------------------------------------------------------------------------------------------------
    Case "ɾ���Ǽ�", "�޸ĵǼ�"
        
        Call mnuViewRefresh_Click
    '------------------------------------------------------------------------------------------------------------------
    Case "��Ա��Ϣ"
        
        If mintIndex = 1 Then
            Call RefreshData("���")
            Call RefreshData("����")
        End If
    '------------------------------------------------------------------------------------------------------------------
    Case Else
    
        Call mnuViewRefresh_Click
        
    End Select
    
    Call AdjustEnableState
    Call RefreshStateInfo
    
    MenuClick = True
    
    AutoRefresh = True
    Exit Function
    
pointEnd:
    AutoRefresh = True
    
    
    Exit Function
    
errHand:
    zlCommFun.StopFlash
    DoEvents
    
    AutoRefresh = True
    If ErrCenter = 1 Then
        Resume
    End If
    If blnTran Then gcnOracle.RollbackTrans
End Function

Private Function CheckState(ByVal lng�Ǽ�id As Long, Optional ByVal bytMode As Byte = 1) As Boolean
    Dim strSQL As String
    Dim rs As New ADODB.Recordset
    
    Select Case bytMode
    Case 1
        strSQL = "Select 1 From ���ǼǼ�¼ Where ���״̬<>5 And ID=[1]"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�Ǽ�id)
        CheckState = rs.BOF
    Case 2
        strSQL = "Select 1  From ���ǼǼ�¼ Where ���״̬=4 And ID=[1]"
        Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�Ǽ�id)
        CheckState = (rs.BOF = False)
    End Select
    
    
    
    
End Function

Private Function DeleteItems(ByRef rsSQL As ADODB.Recordset, ByVal rs As ADODB.Recordset, ByVal str���� As String, ByVal lng�Ǽ�id As Long, Optional ByVal lng����id As Long = 0) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        
        Do While Not rs.EOF
            
            '���ϴ������Ŀ��������ҽ��

            If lng����id > 0 Then
            
                strSQL = "ZL_���ǼǼ�¼_ItemCancel('" & str���� & "'," & Val(rs("�嵥id").Value) & ",NULL," & lng����id & ")"
                Call SQLRecordAdd(rsSQL, strSQL)
                strSQL = "ZL_�����Ŀ�嵥_DELETE(" & lng�Ǽ�id & ",NULL," & Val(rs("�嵥id").Value) & "," & lng����id & ")"
                Call SQLRecordAdd(rsSQL, strSQL)
                
            Else
                
                strSQL = "ZL_���ǼǼ�¼_ItemCancel('" & str���� & "'," & Val(rs("�嵥id").Value) & ",'" & rs("���").Value & "')"
                Call SQLRecordAdd(rsSQL, strSQL)
                strSQL = "ZL_�����Ŀ�嵥_DELETE(" & lng�Ǽ�id & ",'" & rs("���").Value & "'," & Val(rs("�嵥id").Value) & ")"
                Call SQLRecordAdd(rsSQL, strSQL)
                
            End If
            
            rs.MoveNext
        Loop
    End If
    
    DeleteItems = True
    
End Function

Private Function InsertItems(ByRef rsSQL As ADODB.Recordset, _
                            ByVal rs As ADODB.Recordset, _
                            ByVal lng�Ǽ�id As Long, _
                            Optional ByVal lng����id As Long = 0, _
                            Optional ByVal blnCallBack As Boolean = False) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '����:  �¼��������Ŀ
    '------------------------------------------------------------------------------------------------------------------
    Dim strTmp As String
    Dim varRow As Variant
    Dim varCol As Variant
    Dim lngLoop As Long
    Dim strSQL As String
    Dim lngSendNo As Long
    Dim lngKey As Long
    Dim str�ɼ�No As String
    Dim strNO  As String
    Dim rsTmp As New ADODB.Recordset
    Dim blnTran As Boolean
    Dim lngCount As Long
    Dim rsSQLTmp As New ADODB.Recordset
    Dim lng�嵥id As Long
    
    If rs.RecordCount > 0 Then
        rs.MoveFirst
        
        lngSendNo = GetNextNo(10)
        
        Do While Not rs.EOF
            
            Call SQLRecord(rsSQLTmp)
            
            
            strTmp = ""
            
            If blnCallBack = False Then
                varRow = Split(rs("�Ʒ���ϸ").Value, ";")
                For lngLoop = 0 To UBound(varRow)
                    
                    varCol = Split(varRow(lngLoop), ":")
                    
                    If strTmp <> "" Then strTmp = strTmp & ";"
                    strTmp = strTmp & varCol(5) & ":" & varCol(2) & ":" & varCol(3) & ":" & varCol(4) & ":" & Val(varCol(8)) & ":" & Val(varCol(6))
                    
                Next
            End If
            
            '���������Ŀ����Ϊҽ��
            If lng����id > 0 Then
                
                lngKey = zlDatabase.GetNextId("�����Ŀ�嵥")
                
                If blnCallBack Then
                    lng�嵥id = Val(rs("�嵥id").Value)
                Else
                    lng�嵥id = 0
                End If
                
                strSQL = "ZL_�����Ŀ�嵥_INSERT(" & lng�Ǽ�id & "," & _
                                                    "NULL," & _
                                                    rs("ID").Value & ",'" & _
                                                    zlCommFun.NVL(rs("�������").Value) & "'," & _
                                                    Val(rs("�����۸�").Value) & "," & _
                                                    Val(rs("���۸�").Value) & "," & _
                                                    Val(rs("ִ�п���id").Value) & "," & _
                                                    IIf(zlCommFun.NVL(rs("�ɼ���ʽid")) = "", "NULL", rs("�ɼ���ʽid")) & "," & _
                                                    IIf(zlCommFun.NVL(rs("�ɼ�����id")) = "", "NULL", rs("�ɼ�����id")) & ",'" & _
                                                    zlCommFun.NVL(rs("����걾").Value) & "','" & _
                                                    zlCommFun.NVL(rs("��鲿λ").Value) & "','" & _
                                                    zlCommFun.NVL(rs("��鲿λid").Value) & "'," & lng����id & "," & IIf(rs("���㷽ʽ").Value = "����", "1", "2") & ",'" & strTmp & "'," & _
                                                    lngKey & "," & _
                                                    lng�嵥id & ")"
                Call SQLRecordAdd(rsSQLTmp, strSQL)
                
                '�������õ��ݺ�
                str�ɼ�No = ""
                strNO = ""
                If zlCommFun.NVL(rs("���㷽ʽ").Value) = "����" Then
                    '����
                    strNO = GetNextNo(14)
                Else
                    strNO = GetNextNo(13)
                End If
                
                If Val(zlCommFun.NVL(rs("�ɼ���ʽid").Value, 0)) > 0 Then
                    '�ɼ�
                    If zlCommFun.NVL(rs("���㷽ʽ").Value) = "����" Then
                        '����
                        str�ɼ�No = GetNextNo(14)
                    Else
                        str�ɼ�No = GetNextNo(13)
                    End If
                End If
                
                strSQL = "ZL_�����Ŀҽ��_NO(" & lngKey & "," & lng����id & ",'" & strNO & "','" & str�ɼ�No & "')"
                Call SQLRecordAdd(rsSQLTmp, strSQL)
                
                strSQL = "zl_�����Ա����_Accept(" & lng�Ǽ�id & "," & lngSendNo & "," & lng����id & "," & mlngDept & "," & lngKey & ",1)"
                Call SQLRecordAdd(rsSQLTmp, strSQL)

                blnTran = True
                gcnOracle.BeginTrans
                
                If rsSQLTmp.RecordCount > 0 Then rsSQLTmp.MoveFirst
                For lngCount = 1 To rsSQLTmp.RecordCount
                    Call zlDatabase.ExecuteProcedure(CStr(rsSQLTmp("SQL").Value), Me.Caption)
                    rsSQLTmp.MoveNext
                Next
                
                '������ط���
                If MakeMedicalCharge(rsSQL, lng�Ǽ�id) = False Then
                    gcnOracle.RollbackTrans
                    blnTran = False
                    Exit Function
                End If
                
                strSQL = "zl_�����Ա����_Accept(" & lng�Ǽ�id & "," & lngSendNo & "," & lng����id & "," & mlngDept & "," & lngKey & ",2)"
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                
                gcnOracle.CommitTrans
                blnTran = False
                
            Else
                
                lngKey = zlDatabase.GetNextId("�����Ŀ�嵥")
                
                strSQL = "ZL_�����Ŀ�嵥_INSERT(" & lng�Ǽ�id & ",'" & _
                                                    rs("���").Value & "'," & _
                                                    rs("ID").Value & ",'" & _
                                                    rs("�������").Value & "'," & _
                                                    Val(rs("�����۸�").Value) & "," & _
                                                    Val(rs("���۸�").Value) & "," & _
                                                    Val(rs("ִ�п���id").Value) & "," & _
                                                    IIf(rs("�ɼ���ʽid") = "", "NULL", rs("�ɼ���ʽid")) & "," & _
                                                    IIf(rs("�ɼ�����id") = "", "NULL", rs("�ɼ�����id")) & ",'" & _
                                                    zlCommFun.NVL(rs("����걾").Value) & "','" & _
                                                    rs("��鲿λ").Value & "','" & _
                                                    rs("��鲿λid").Value & "',0," & IIf(rs("���㷽ʽ").Value = "����", "1", "2") & ",'" & strTmp & "'," & lngKey & ")"
                Call SQLRecordAdd(rsSQLTmp, strSQL)
                
                
                '�������õ��ݺ�
                strSQL = "Select a.����id From �����Ա���� a Where a.�������=[1] and a.�Ǽ�id=[2]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, rs("���").Value, lng�Ǽ�id)
                If rsTmp.BOF = False Then
                    Do While Not rsTmp.EOF
                        
                        str�ɼ�No = ""
                        strNO = ""
                        
                        If zlCommFun.NVL(rs("���㷽ʽ").Value) = "����" Then
                            '����
                            strNO = GetNextNo(14)
                        Else
                            strNO = GetNextNo(13)
                        End If
                        
                        If Val(zlCommFun.NVL(rs("�ɼ���ʽid").Value, 0)) > 0 Then
                            '�ɼ�
                            If zlCommFun.NVL(rs("���㷽ʽ").Value) = "����" Then
                                '����
                                str�ɼ�No = GetNextNo(14)
                            Else
                                str�ɼ�No = GetNextNo(13)
                            End If
                        End If
                        
                        strSQL = "ZL_�����Ŀҽ��_NO(" & lngKey & "," & rsTmp("����id").Value & ",'" & strNO & "','" & str�ɼ�No & "')"
                        Call SQLRecordAdd(rsSQLTmp, strSQL)
                        
                        rsTmp.MoveNext
                    Loop
                End If
                
                strSQL = "zl_�����Ա����_Accept(" & lng�Ǽ�id & "," & lngSendNo & ",NULL," & mlngDept & "," & lngKey & ",1)"
                Call SQLRecordAdd(rsSQLTmp, strSQL)
                
                blnTran = True
                gcnOracle.BeginTrans
                If rsSQLTmp.RecordCount > 0 Then rsSQLTmp.MoveFirst
                For lngCount = 1 To rsSQLTmp.RecordCount
                    
                    Call zlDatabase.ExecuteProcedure(CStr(rsSQLTmp("SQL").Value), Me.Caption)
                    rsSQLTmp.MoveNext
                Next
                
                '������ط���
                If MakeMedicalCharge(rsSQL, lng�Ǽ�id) = False Then
                    gcnOracle.RollbackTrans
                    blnTran = False
                    Exit Function
                End If

                strSQL = "zl_�����Ա����_Accept(" & lng�Ǽ�id & "," & lngSendNo & ",NULL," & mlngDept & "," & lngKey & ",2)"
                Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                
                gcnOracle.CommitTrans
                blnTran = False
                
            End If

            rs.MoveNext
        Loop
    End If
    
    InsertItems = True
    
End Function

Private Function RefreshData(Optional ByVal strMenuItem As String = "", Optional ByVal strParam As String) As Boolean
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ�ˢ��/װ������
    '------------------------------------------------------------------------------------------------------------------
    Dim rs As New ADODB.Recordset
    Dim rsData As New ADODB.Recordset
    Dim lngKey As Long
    Dim lngLoop As Long
    Dim strStart As String
    Dim strEnd As String
    Dim varParam As Variant
    Dim strSQL As String
    Dim intRow As Integer
    Dim intCol As Integer
    Dim strField As String
    Dim objStd As New StdPicture
    Dim strTmpFile As String
    Dim strTmp As String
        
    If strParam = "" Then strParam = "'"
    varParam = Split(strParam, "'")
    
    On Error GoTo errHand
    
    Select Case strMenuItem
        Case "����"             '��ȡ�����Ա�Ļ�����Ϣ
            lngKey = Val(vsf(mintIndex).RowData(vsf(mintIndex).Row))
            If lngKey <= 0 Then Exit Function

            
            If mintIndex <> 3 Then
                If Trim(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "[����]"))) <> "" Then
                                                
                    Select Case Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "��־")))
                    Case 1, 2              '0-���˷�����;1-����������;2-���������;99-�ܼ���Ա��
                        pic(1).Visible = True
                        pic(0).Visible = False
                        
                        strSQL = "Select a.* From ��Լ��λ a,���ǼǼ�¼ b Where a.ID=b.��Լ��λID And b.ID=[1]"
                        
                        '����ת������
                        '----------------------------------------------------------------------------------------------
                        mblnDataMoved = False
                        If mintIndex = 2 Then mblnDataMoved = DataMove(lngKey)
                        If mblnDataMoved Then
                            strSQL = Replace(strSQL, "���ǼǼ�¼", "H���ǼǼ�¼")
                        End If
                        
                        Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "�Ǽ�id"))))
                        If rs.BOF = False Then
                            lblValue(2).Caption = zlCommFun.NVL(rs("����").Value)
                            lblValue(11).Caption = zlCommFun.NVL(rs("��ϵ��").Value)
                            lblValue(12).Caption = zlCommFun.NVL(rs("�绰").Value)
                            lblValue(13).Caption = zlCommFun.NVL(rs("�����ʼ�").Value)
                        End If
                    End Select
                    Exit Function
                End If
            End If
            
            pic(0).Visible = True
            pic(1).Visible = False

            If mintIndex <> 3 Then
                If Trim(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "[����]"))) <> "" Then Exit Function
            End If
            
            strSQL = GetPublicSQL(SQL.���˻�����Ϣ)
            
            '����ת������
            '----------------------------------------------------------------------------------------------------------
            mblnDataMoved = False
            If mintIndex >= 2 Then mblnDataMoved = DataMove(Val(vsf(mintIndex).RowData(vsf(mintIndex).Row)), 2)
            If mblnDataMoved Then
                strSQL = Replace(strSQL, "���ǼǼ�¼", "H���ǼǼ�¼")
                strSQL = Replace(strSQL, "�����Ա����", "H�����Ա����")
            End If
            
            Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsf(mintIndex).RowData(vsf(mintIndex).Row)))
            If rs.BOF = False Then
                lblValue(3).Caption = zlCommFun.NVL(rs("����").Value)
                lblValue(4).Caption = zlCommFun.NVL(rs("�Ա�").Value)
                lblValue(5).Caption = zlCommFun.NVL(rs("����").Value)
                lblValue(0).Caption = zlCommFun.NVL(rs("����״��").Value)
                lblValue(6).Caption = Format(zlCommFun.NVL(rs("���ʱ��").Value), "yyyy-MM-dd")
                lblValue(7).Caption = zlCommFun.NVL(rs("�����").Value)
                lblValue(8).Caption = zlCommFun.NVL(rs("������λ").Value)
                lblValue(1).Caption = zlCommFun.NVL(rs("������").Value)
                lblValue(9).Caption = zlCommFun.NVL(rs("�������").Value)
                lblValue(10).Caption = zlCommFun.NVL(rs("��ϵ�˵绰").Value)
                mlng��첡��id = zlCommFun.NVL(rs("��첡��id").Value, 0)
                                
            End If
                                            
            picState.Visible = True
            strSQL = GetPublicSQL(SQL.���˷��øſ�)
                        
            
            '����ת������
            '----------------------------------------------------------------------------------------------------------
            If mblnDataMoved Then
                '��ʱ����Ӧ����ȫת�������ڳ�صĶ������ǣ�
                strSQL = Replace(strSQL, "���˷��ü�¼", "H���˷��ü�¼")
                strSQL = Replace(strSQL, "�����Ա����", "H�����Ա����")
                strSQL = Replace(strSQL, "���ǼǼ�¼", "H���ǼǼ�¼")
                strSQL = Replace(strSQL, "����ҽ����¼", "H����ҽ����¼")
            Else
                '��ʱ���ܷ����Ѳ��ݻ���ȫת��
                gstrSQL = "Select a.���ʱ�� From ���ǼǼ�¼ a,�����Ա���� b Where a.ID=b.�Ǽ�id And b.ID=[1]"
                Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(vsf(mintIndex).RowData(vsf(mintIndex).Row)))
                If rs.BOF = False Then
                    If zlDatabase.DateMoved(Format(rs("���ʱ��").Value, "yyyy-MM-dd HH:mm:ss"), , glngSys, Me.Caption) Then
                        strTmp = strSQL
                        strTmp = Replace(strTmp, "���˷��ü�¼", "H���˷��ü�¼")
                        strSQL = strSQL & " Union All " & strTmp
                    End If
                End If
            End If
            
            
            Set rsData = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(vsf(mintIndex).RowData(vsf(mintIndex).Row)))
            If CalcCharge(rsData, rs) Then
                picState.Visible = (Val(Format(zlCommFun.NVL(rs("δ�ս��").Value, 0), "0.00")) = 0)
            End If
                                
            '������Ƭ
            picPhoto.Cls
            strSQL = "Select B.* From �����Ա���� A,������Ƭ B Where A.����id=B.����id AND A.ID=[1]"
            
            '����ת������
            '----------------------------------------------------------------------------------------------------------
            If mblnDataMoved Then
                strSQL = Replace(strSQL, "�����Ա����", "H�����Ա����")
            End If
            
            Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngKey)
            If rs.BOF = False Then
                strTmpFile = ""
                strTmpFile = ReadPicture(rs, "��Ƭ", strTmpFile)
                
                If strTmpFile <> "" Then
                    Set objStd = VB.LoadPicture(strTmpFile)
                    Call DrawPicture(picPhoto, objStd, objStd.Width, objStd.Height)
                End If
            End If
            
            
        Case "ԤԼ"             '��ȡȷ�ϵ�ԤԼ������Ա��Ϣ
                                    
            strStart = Split(vsf(0).Tag, "|")(0)
            strEnd = Split(vsf(0).Tag, "|")(1)
            
            strSQL = GetPublicSQL(SQL.���Ǽǵ���)
            Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDept, CDate(strStart), CDate(strEnd), 2)
            If rs.BOF = False Then Call LoadOutLineGrid(0, rs, , , ils13)
                        
        Case "���"             '��ȡ����������Ա��Ϣ
             
            strStart = Split(mstr���������ʱ�䷶Χ, "|")(0)
            strEnd = Split(mstr���������ʱ�䷶Χ, "|")(1)
            
'            strStart = Split(vsf(1).Tag, "|")(0)
'            strEnd = Split(vsf(1).Tag, "|")(1)
            
            strSQL = GetPublicSQL(SQL.���Ǽǵ���, mint������ѯ����)
            Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDept, CDate(strStart), CDate(strEnd), 4)
            If rs.BOF = False Then Call LoadOutLineGrid(1, rs, , , ils13)
            
        Case "���"             '��ȡ�����ɵ���Ա��Ϣ(��ʱ��)
            
            strStart = Split(vsf(2).Tag, "|")(0)
            strEnd = Split(vsf(2).Tag, "|")(1)
            
            '����ת������
            '----------------------------------------------------------------------------------------------------------
            mblnDataMoved = False
            mblnDataMoved = zlDatabase.DateMoved(Format(strStart, "yyyy-MM-dd HH:mm:ss"), , glngSys, Me.Caption)
            mblnDataMoved = True
            
            strSQL = GetPublicSQL(SQL.���Ǽǵ���, , mblnDataMoved)
            Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngDept, CDate(strStart), CDate(strEnd), 5)
            If rs.BOF = False Then Call LoadOutLineGrid(2, rs, , , ils13)
                        
        Case "��ѯ"             '����������ȡ������Ա��Ϣ
            
            Call InheritResetVsf(3)
            DoEvents
            
            Select Case Split(vsf(3).Tag, "^")(1)
            Case "ָ  ��"
                strStart = Split(vsf(3).Tag, "^")(2)
            Case "��  ��"
                strStart = ""
            Case Else
                strStart = GetDateTime(Split(vsf(3).Tag, "^")(1), 1)
            End Select
            
            '����ת������
            '----------------------------------------------------------------------------------------------------------
            If strStart = "" Then
                mblnDataMoved = True
            Else
                mblnDataMoved = False
                mblnDataMoved = zlDatabase.DateMoved(Format(strStart, "yyyy-MM-dd HH:mm:ss"), , glngSys, Me.Caption)
            End If
            
            strSQL = "SELECT B.ID," & _
                            "DECODE(B.���״̬,4,'��ʼ',5,'���') AS ״̬," & _
                            "A.����id," & _
                            "C.���� AS ��쵥��," & _
                            "A.�����," & _
                            "A.����," & _
                            "A.�Ա�," & _
                            "A.����," & _
                            "D.���� AS ����," & _
                            "to_char(A.��������,'yyyy-mm-dd') AS ��������," & _
                            "A.����״��,B.�Ǽ�id " & _
                        "FROM ������Ϣ A,�����Ա���� B,���ǼǼ�¼ C,��Լ��λ D  " & _
                        "WHERE C.��Լ��λid=D.ID(+) AND B.���״̬=5 AND B.��챨��=1 AND A.����ID=B.����ID AND C.ID=B.�Ǽ�id "
            
            strSQL = strSQL & GetQueryCondition(vsf(3).Tag)
            
            If mblnDataMoved Then
                strTmp = strSQL
                strTmp = Replace(strTmp, "���ǼǼ�¼", "H���ǼǼ�¼")
                strTmp = Replace(strTmp, "�����Ա����", "H�����Ա����")
                strTmp = Replace(strTmp, "���˲�������", "H���˲�������")
                strTmp = Replace(strTmp, "����ҽ����¼", "H����ҽ����¼")
                strTmp = Replace(strTmp, "����ҽ������", "H����ҽ������")
                strTmp = Replace(strTmp, "���˲���������", "H���˲���������")
                strSQL = strSQL & " Union All " & strTmp
            End If
            
            Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            If rs.BOF = False Then Call LoadGrid(vsf(3), rs, , , ils13)
            
        Case "�����Ա"
                                    
            Dim blnField As Boolean
            Dim strIcon As String
            Dim intTmp As Integer
            
            '����ת������
            '----------------------------------------------------------------------------------------------------------
            mblnDataMoved = False
            If Val(varParam(3)) = 5 Then
                '����ɵ����ҵ��
                
                If Val(varParam(1)) = 0 Then
                    '���˵����ҵ��
                    mblnDataMoved = zlDatabase.DateMoved(Format(Split(varParam(5), "|")(0), "yyyy-MM-dd HH:mm:ss"), , glngSys, Me.Caption)
                Else
                    '�����ĳ�����ҵ��
                    mblnDataMoved = DataMove(Val(varParam(1)))
                End If
            End If
            
            intTmp = 0
            If mintIndex = 0 Then
                strSQL = GetPublicSQL(SQL.��������Ա1, Val(varParam(1)) & "'" & intTmp, mblnDataMoved)
            Else
                If mintIndex = 1 Then
                    intTmp = mint������ѯ����
                End If
                strSQL = GetPublicSQL(SQL.��������Ա, Val(varParam(1)) & "'" & intTmp, mblnDataMoved)
            End If
            
            Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(varParam(1)), CStr(varParam(2)), Val(varParam(3)), Val(varParam(4)), CDate(Split(varParam(5), "|")(0)), CDate(Split(varParam(5), "|")(1)))
            If rs.BOF = False Then
                
                intRow = Val(varParam(0))
                
                Do While Not rs.EOF
                    
                    intRow = intRow + 1
                    
                    vsf(mintIndex).AddItem "", intRow
                    
                    vsf(mintIndex).RowData(intRow) = rs("ID").Value
                    For intCol = 0 To vsf(mintIndex).Cols - 1
                    
                        strField = vsf(mintIndex).Cell(flexcpData, 0, intCol)
                        If strField <> "" And strField <> "�Ƿ�װ��" Then
                            If Left(strField, 1) <> "[" Then
                                vsf(mintIndex).TextMatrix(intRow, intCol) = zlCommFun.NVL(rs(strField))
                            Else
                                strField = Mid(strField, 2, Len(strField) - 2)
                                strIcon = ""
                                
                                On Error Resume Next
                                blnField = False
                                blnField = (UCase(rs(strField).Name) = UCase(strField))
                                If blnField Then
                                
                                    On Error GoTo errHand
                            
                                    strIcon = zlCommFun.NVL(rs(strField))
                                    If strIcon <> "" Then
                                        Set vsf(mintIndex).Cell(flexcpPicture, intRow, intCol) = ils13.ListImages(strIcon).Picture
                                    End If
                                    
                                    
                                    vsf(mintIndex).Cell(flexcpData, intRow, intCol) = strIcon
                                    vsf(mintIndex).TextMatrix(intRow, intCol) = strIcon
                                End If
                            End If
                        End If
                    Next
                    rs.MoveNext
                Loop
            End If
            
    End Select
    
    RefreshData = True
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function GetQueryCondition(ByVal strCondition As String, Optional ByVal bytMode As Byte = 1) As String
    '--------------------------------------------------------------------------------------------------------
    '���ܣ�
    '--------------------------------------------------------------------------------------------------------
    Dim varTmp As Variant
    Dim varTmp2 As Variant
    Dim strTmp As String
    Dim strResult As String
     
   
    '�����Ǹ��������������ɵ��������
    
    If strCondition = "" Then Exit Function
    
    varTmp = Split(strCondition, "^")
    
    '��첿��
    If Val(varTmp(0)) > 0 Then strResult = strResult & " AND C.��첿��id + 0 = " & Val(varTmp(0))

    '���ʱ��
    If Trim(varTmp(1)) <> "��  ��" Then
        Select Case Trim(varTmp(1))
        Case "ָ  ��"
            strResult = strResult & " AND C.���ʱ�� BETWEEN TO_DATE('" & Format(varTmp(2), "yyyy-mm-dd") & " 00:00:00', 'yyyy-mm-dd hh24:mi:ss') AND TO_DATE('" & Format(varTmp(3), "yyyy-mm-dd") & " 23:59:59', 'yyyy-mm-dd hh24:mi:ss')"
        Case Else
            strResult = strResult & " AND C.���ʱ�� BETWEEN TO_DATE('" & GetDateTime(varTmp(1), 1) & "', 'yyyy-mm-dd hh24:mi:ss') AND TO_DATE('" & GetDateTime(varTmp(1), 2) & "', 'yyyy-mm-dd hh24:mi:ss')"
        End Select
    End If
    
    
    varTmp2 = Split(Trim(varTmp(4)), ",")
    strTmp = ""
    For mlngLoop = 0 To UBound(varTmp2)
        If InStr(varTmp2(mlngLoop), "-") = 0 Then
            strTmp = strTmp & "  OR C.����='" & varTmp2(mlngLoop) & "'"
        Else
            strTmp = strTmp & "  OR C.���� BETWEEN '" & Mid(varTmp2(mlngLoop), 1, InStr(varTmp2(mlngLoop), "-") - 1) & "' AND '" & Mid(varTmp2(mlngLoop), InStr(varTmp2(mlngLoop), "-") + 1) & "'"
        End If
    Next
    If strTmp <> "" Then strResult = strResult & " AND (1=2 " & strTmp & ")"

    If Trim(varTmp(5)) <> "��  ��" Then
        
        Select Case Trim(varTmp(5))
        Case "ָ  ��"
            strResult = strResult & " AND B.���ʱ�� BETWEEN TO_DATE('" & Format(varTmp(6), "yyyy-mm-dd") & " 00:00:00', 'yyyy-mm-dd hh24:mi:ss') AND TO_DATE('" & Format(varTmp(7), "yyyy-mm-dd") & " 23:59:59', 'yyyy-mm-dd hh24:mi:ss')"
        Case Else
            strResult = strResult & " AND B.���ʱ�� BETWEEN TO_DATE('" & GetDateTime(varTmp(5), 1) & "', 'yyyy-mm-dd hh24:mi:ss') AND TO_DATE('" & GetDateTime(varTmp(5), 2) & "', 'yyyy-mm-dd hh24:mi:ss')"
        End Select
        
    End If
    
    '��������
    If Trim(varTmp(8)) <> "" Then strResult = strResult & " AND A.���� LIKE '%" & varTmp(8) & "%'"
    
    '�������
    If Val(varTmp(9)) > 0 Then strResult = strResult & " AND C.��Լ��λid = " & Val(varTmp(9))
    
    '�����Ŀ���ԱȽ��
    If Val(varTmp(11)) > 0 Then
        strResult = strResult & _
                    " AND (C.����,B.����id) IN (SELECT E.�Һŵ�,E.����id " & _
                        "FROM ���˲��������� A, " & _
                             "����������Ŀ B, " & _
                             "���˲������� C, " & _
                             "����ҽ������ D, " & _
                             "����ҽ����¼ E  " & _
                        "Where A.������ID = B.ID " & _
                              "AND A.����id=C.ID " & _
                              "AND D.����id=C.������¼id " & _
                              "AND E.ID=D.ҽ��ID " & _
                              "AND E.������Դ=4 " & _
                              "AND B.ID=" & Val(varTmp(11))
                
        If Val(varTmp(12)) = 0 Then
            strResult = strResult & " AND A.��ֵ����=0 AND DECODE(A.��ֵ����,0,TO_NUMBER(A.��������),0)"
            strTmp = Val(varTmp(15))
        Else
            strResult = strResult & " AND A.��������"
            strTmp = "'" & varTmp(15) & "'"
        End If
        
        Select Case varTmp(14)
        Case "����"
            strResult = strResult & ">" & strTmp
        Case "С��"
            strResult = strResult & "<" & strTmp
        Case "���ڵ���"
            strResult = strResult & ">=" & strTmp
        Case "С�ڵ���"
            strResult = strResult & "<=" & strTmp
        Case "������"
            strResult = strResult & "<>" & strTmp
        Case "����"
            strResult = strResult & " LIKE '%" & varTmp(15) & "%'"
        Case "�ڷ�Χ��"
            If Val(varTmp(12)) = 0 Then
                strResult = strResult & " BETWEEN " & strTmp & " AND " & Val(varTmp(16))
            Else
                strResult = strResult & " BETWEEN " & strTmp & " AND '" & varTmp(16) & "'"
            End If
        Case Else
            strResult = strResult & "=" & strTmp
        End Select
        strResult = strResult & ")"
    End If
    
    GetQueryCondition = strResult
    
End Function


Private Sub InheritAppendSpaceRows(ByVal intIndex As Integer)
    '------------------------------------------------------------------------------------------------------------------
    '���ܣ����������
    '------------------------------------------------------------------------------------------------------------------
    Select Case intIndex
    Case 0
        Call AppendRows(vsf(intIndex), lnX0, lnY0)
    Case 1
        Call AppendRows(vsf(intIndex), lnX1, lnY1, mlngHideRows)
    Case 2
        Call AppendRows(vsf(intIndex), lnX2, lnY2)
    Case 3
        Call AppendRows(vsf(intIndex), lnX3, lnY3)
    End Select
End Sub

Private Sub ResetActiveForm()
    '--------------------------------------------------------------------------------------------------------
    '���ܣ�
    '--------------------------------------------------------------------------------------------------------
    
    If Not (mfrmActive Is Nothing) Then
        Unload mfrmActive
        Set mfrmActive = Nothing
    End If
        
End Sub

Private Sub PrintData(ByVal bytMode As Byte)
    '--------------------------------------------------------------------------------------------------------
    '���ܣ� ��ӡ����
    '������ bytMode                         ��ӡ��ʽ��1-��ӡ��2-Ԥ����3-�����Excel��
    '--------------------------------------------------------------------------------------------------------
    Dim objPrint As New zlPrint1Grd
    Dim objRow As zlTabAppRow
            
    mblnNoAllowChange = True
    
    If UserInfo.���� = "" Then Call GetUserInfo
        
    Select Case mintIndex
    Case 0
        Call CopyGrid(vsf(mintIndex), vsfPrint, 2)
        objPrint.Title = "ԤԼ��쵥"
    Case 1
        Call CopyGrid(vsf(mintIndex), vsfPrint, 2)
        objPrint.Title = "����������쵥"
    Case 2
        Call CopyGrid(vsf(mintIndex), vsfPrint, 2)
        objPrint.Title = "����ɵ���쵥"
    Case 3
        Call CopyGrid(vsf(mintIndex), vsfPrint, 1)
        objPrint.Title = "��ѯ��쵥"
    End Select
    
    If mintIndex <> 3 Then
        Set objRow = New zlTabAppRow
        objRow.Add "��첿��:" & zlCommFun.GetNeedName(cboDept.Text)
        objRow.Add ""
        objPrint.UnderAppRows.Add objRow
    End If
    
    Set objPrint.Body = vsfPrint
        
    If bytMode = 1 Then bytMode = zlPrintAsk(objPrint)

    If bytMode >= 1 And bytMode <= 3 Then Call zlPrintOrView1Grd(objPrint, bytMode)

    mblnNoAllowChange = False
End Sub

Private Sub RefreshQueryMenu()
    '--------------------------------------------------------------------------------------------------------
    '���ܣ�
    '--------------------------------------------------------------------------------------------------------
    Dim strSectoin  As String
    Dim strTmp As String
    Dim lngLoop As Long
    
    For lngLoop = mnuViewFilterItem.UBound To 2 Step -1
        Unload mnuViewFilterItem(lngLoop)
    Next
    mnuViewFilterItem(1).Visible = False
    
    strSectoin = "˽��ģ��\" & App.ProductName & "\���˲���"
    
    For lngLoop = 1 To CLng(Val(GetSetting("ZLSOFT", strSectoin, "��������", "0")))
        
        strTmp = GetSetting("ZLSOFT", strSectoin, "���˲���" & lngLoop, "")
        
        If Trim(strTmp) <> "" And InStr(strTmp, "|") > 0 Then
            mnuViewFilterItem(1).Visible = True
            Load mnuViewFilterItem(lngLoop + 1)
        
            mnuViewFilterItem(lngLoop + 1).Caption = Mid(strTmp, 1, InStr(strTmp, "|") - 1) & "(&" & lngLoop & ")"
            mnuViewFilterItem(lngLoop + 1).Tag = Mid(strTmp, InStr(strTmp, "|") + 1)
                        
        End If
    Next
End Sub

'���������弰��ؼ����¼�����******************************************************************************************
Private Sub cboDept_Click()
    Dim intIndex As Integer
    Dim rs As New ADODB.Recordset
    
    If mblnStartUp Then Exit Sub
    If mlngDept = cboDept.ItemData(cboDept.ListIndex) Then Exit Sub
    
    mlngDept = cboDept.ItemData(cboDept.ListIndex)
    
    gstrSQL = "SELECT A.*,ROWNUM AS ��� from " & _
                "(SELECT ID, ����,���� " & _
                "From �����ļ�Ŀ¼ " & _
                "Where ���� = 1 " & _
                    "AND Ӧ�� = 2 and ',' || ����ID || ',' like '%," & mlngDept & ",%' " & _
            ") A " & _
            "ORDER BY A.ID"
    
    For intIndex = 1 To mnuReportAddOutLineCase.UBound
        Unload mnuReportAddOutLineCase(intIndex)
    Next
    mnuReportAddOutLineCase(0).Caption = "<�޿��ò���>"
    mnuReportAddOutLineCase(0).Tag = ""
    
    Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    
    If rs.BOF = False Then
        intIndex = 0
        Do While Not rs.EOF
            If intIndex > 0 Then
                Load mnuReportAddOutLineCase(intIndex)
                mnuReportAddOutLineCase(intIndex).Visible = True
            End If
            
            mnuReportAddOutLineCase(intIndex).Caption = zlCommFun.NVL(rs("����").Value) & "(&" & rs("���").Value & ")"
            mnuReportAddOutLineCase(intIndex).Tag = zlCommFun.NVL(rs("ID").Value)
            
            intIndex = intIndex + 1
            rs.MoveNext
        Loop
    End If
    
    Call mnuViewRefresh_Click

End Sub

Private Sub cmdKind_Click(Index As Integer)

    If mintIndex = Index Then
        vsf(Index).SetFocus
        Exit Sub
    End If
    
    mstrSvrFind = ""
    picShow.Tag = Index
    
    mintIndex = Index
    
    '1.�������沼��
    For mlngLoop = cmdKind.LBound To cmdKind.UBound
        cmdKind(mlngLoop).Tag = IIf(mlngLoop <= Index, 0, 1)
    Next

    Call picClass_Resize
    Call picShow_Resize
    
    vsf(Index).SetFocus
    
    DoEvents
    
    mlngSvrKey(Index) = 0
    
    '����ұ�����
    Call ClearData("���")
    Call vsf_AfterRowColChange(Index, 0, 0, vsf(Index).Row, vsf(Index).Col)
    
    Call AdjustEnableState
    Call RefreshStateInfo
End Sub

Private Sub cmdMenu_Click()
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(cmdMenu.hWnd, objPoint)
    
    mbytPopMenu = 3
    Set mobjPopMenu = New clsPopMenu
    Call mobjPopMenu.ShowPopupMenu(objPoint.X * Screen.TwipsPerPixelX, objPoint.Y * Screen.TwipsPerPixelY - 255 * 11 - 300)
    
    txt(1).Text = ""
    LocationObj txt(1)
    
End Sub

Private Sub mnuFilePrintBook_Click()
    Call MenuClick("��챨�浥")
End Sub

Private Sub mnuFilePrintList_Click()
    Call MenuClick("���ָ����")
End Sub


Private Sub mnuFilePrintRequest_Click()
    Call MenuClick("��Ŀ���뵥")
End Sub

Private Sub mnuFileRequest_Click()
    Call MenuClick("������뵥")
End Sub

Private Sub mnuFileSendMail_Click()
    Call MenuClick("�����ʼ�")
End Sub

Private Sub mnuMedicalCallBack_Click()
    Call MenuClick("������Ŀ")
End Sub

Private Sub mnuMedicalDept_Click()
    Call MenuClick("ִ�е���")
End Sub

Private Sub mnuMedicalGroupDelete_Click()
    Call MenuClick("�Ƴ���Ա")
End Sub

Private Sub mnuMedicalGroupIn_Click()
    Call MenuClick("��Ա����")
End Sub

Private Sub mnuMedicalGroupOut_Click()
    Call MenuClick("ȡ������")
End Sub

Private Sub mnuMedicalNewType_Click(Index As Integer)
    Call MenuClick("���Ǽ�", Index)
End Sub

Private Sub mnuMedicalPhoto_Click()
    Call MenuClick("��Ƭ�ɼ�")
End Sub

Private Sub mnuReportAddOutLineCase_Click(Index As Integer)
    Call MenuClick("��д�ܽ�", Val(mnuReportAddOutLineCase(Index).Tag))
End Sub

'Private Sub mnuReportAgain_Click()
'    Call MenuClick("��츴��")
'End Sub


Private Sub mnuReportDelOutLine_Click()
    Call MenuClick("ɾ���ܽ�")
End Sub

Private Sub mnuReportDesign_Click(Index As Integer)
        
    Select Case Index
    Case 0
        Call ReportDesign(gcnOracle, glngSys, "ZL1_BILL_1861_2_1", Me, True)
    Case 1
        Call ReportDesign(gcnOracle, glngSys, "ZL1_BILL_1861_2", Me, True)
    Case 2
        Call ReportDesign(gcnOracle, glngSys, "ZL1_BILL_1861_2_2", Me, True)
    End Select
End Sub

Private Sub mnuReportModifyOutLine_Click()
    Call MenuClick("�޸��ܽ�")
End Sub

Private Sub mnuReportView_Click()
    Call MenuClick("�鿴����")
End Sub


Private Sub mnuViewFilter_Click()
    Call MenuClick("������")
End Sub

Private Sub mnuViewFilterItem_Click(Index As Integer)
    Dim strCondition As String
        
    AutoRefresh = False
    If Index = 0 Then
        '�����ѯ
            
        strCondition = vsf(3).Tag
        If frmMedicalStationFilter.ShowEdit(Me, strCondition) Then
            cmdKind(3).Caption = "&Z.�Զ����ѯ"
            vsf(3).Tag = strCondition
            
            Call cmdKind_Click(3)
            
            zlCommFun.ShowFlash "���Ժ����ڲ�ѯ...", Me
            DoEvents
            
            Call RefreshData("��ѯ")
            
            zlCommFun.StopFlash
            
            mintIndex = 1
            Call cmdKind_Click(3)
        End If
        
        Call RefreshQueryMenu
    Else
        '��ѯ����
        
        If mnuViewFilterItem(Index).Tag <> "" Then
            cmdKind(3).Caption = "&Z." & Mid(mnuViewFilterItem(Index).Caption, 1, Len(mnuViewFilterItem(Index).Caption) - 4)
            vsf(3).Tag = mnuViewFilterItem(Index).Tag
            
            Call cmdKind_Click(3)
            
            zlCommFun.ShowFlash "���Ժ����ڲ�ѯ...", Me
            DoEvents
            
            Call RefreshData("��ѯ")
            
            zlCommFun.StopFlash
            
            mintIndex = 1
            Call cmdKind_Click(3)
        End If
        
    End If
    AutoRefresh = True
End Sub

Private Sub Form_Activate()
    
    If mblnStartUp = False Then Exit Sub
    
    Call InitLoad
    Call InitSysPara
    Call ApplyPrivilege(gstrPrivs)
    
    Call InheritAppendSpaceRows(0)
    Call InheritAppendSpaceRows(1)
    Call InheritAppendSpaceRows(2)
    Call InheritAppendSpaceRows(3)
    
    DoEvents
    
    If InitActive = False Then
        mblnStartUp = False
        Unload Me
        Exit Sub
    End If
    
    mblnStartUp = False
    
    Call cboDept_Click
    Call tbs_Click          '�˵�����Ϊ��ˢ������
    
End Sub

Private Sub Form_Load()
    
    mblnStartUp = True
    Call RestoreWinState(Me, App.ProductName)
    
End Sub

Private Sub Form_Resize()
    
    On Error Resume Next
    
    If Me.WindowState = 1 Then Exit Sub
    
    If imgY_S.Left > Me.ScaleWidth - 1000 Then imgY_S.Left = Me.ScaleWidth - 1000
        
    With picClass
        .Left = 0
        .Top = IIf(cbrThis.Visible, cbrThis.Height, 0)
        .Width = imgY_S.Left
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0) - fraInfo.Height + 120
    End With
    
    With fraInfo
        .Left = 0
        .Top = picClass.Top + picClass.Height - 120
        .Width = picClass.Width
    End With
    
    With txt(1)
        .Width = fraInfo.Width - .Left - 75
    End With
    
    
    With imgY_S
        .Top = picClass.Top
        .Height = picClass.Height
    End With
    
    With fraBack
        .Left = imgY_S.Left + imgY_S.Width
        .Top = picClass.Top - 90
        .Width = Me.ScaleWidth - .Left
    End With
    
    With tbs
        .Left = fraBack.Left
        .Top = fraBack.Top + fraBack.Height + 30
        .Width = fraBack.Width
    End With
    
    With picContainer
        .Left = tbs.Left
        .Top = tbs.Top + tbs.Height + 15
        .Width = tbs.Width
        .Height = Me.ScaleHeight - .Top - IIf(stbThis.Visible, stbThis.Height, 0)
    End With
    
    With pic(0)
        .Width = fraBack.Width - .Left - 30
    End With
    
    With picState
        .Left = pic(0).Width - .Width - 30
    End With
    
    pic(1).Move pic(0).Left, pic(0).Top, pic(0).Width, pic(0).Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
        
    If mblnStartUp Then
        Cancel = True
        Exit Sub
    End If
    
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "������Ϣ", lbl(1).Tag)
    Call SaveSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "��ʾ����", IIf(mnuViewShowResult.Checked, 1, 0))
        
    Call SaveWinState(Me, App.ProductName)
    
    If mrsFind.State = adStateOpen Then mrsFind.Close
    Set mrsFind = Nothing
    
'    Set mobjCls = Nothing
'    Set mclsCore = Nothing
    
End Sub

Private Sub imgY_S_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 1 Then Exit Sub
    
    imgY_S.Left = imgY_S.Left + X
    
    If imgY_S.Left < 3000 Then imgY_S.Left = 3000
    If Me.Width - imgY_S.Left - imgY_S.Width < 1000 Then imgY_S.Left = Me.Width - imgY_S.Width - 1000

    Form_Resize
End Sub


Private Sub mnuChargeDelete_Click()
    Call MenuClick("ɾ������")
End Sub

Private Sub mnuChargeMain_Click()
    Call MenuClick("��������")
End Sub

Private Sub mnuChargeModify_Click()
    Call MenuClick("�޸ĸ���")
End Sub

Private Sub mnuChargeAddType_Click(Index As Integer)
    Select Case Index
    Case 0
        Call MenuClick("�����շѵ���")
    Case 1
        Call MenuClick("���Ӽ��ʵ���")
    Case 2
        Call MenuClick("������Ѻ��õǼ�")
    End Select
End Sub


Private Sub mnuMedicalCompleteCancel_Click()
    Call MenuClick("ȡ�����")
End Sub

Private Sub mnuMedicalComplete_Click()
    Call MenuClick("������")
End Sub

Private Sub mnuMedicalBeginCancel_Click()
    Call MenuClick("ȡ����ʼ")
End Sub

Private Sub mnuFileExit_Click()
    Unload Me
End Sub

Private Sub mnuFileOutExcel_Click()
    Call PrintData(3)
End Sub

Private Sub mnuFilePrint_Click()
    Call PrintData(1)
End Sub

Private Sub mnuFilePrintSet_Click()
    Call zlPrintSet
End Sub

Private Sub mnuFilePrintView_Click()
    Call PrintData(2)
End Sub

Private Sub mnuFilePara_Click()
    Call MenuClick("��������")
End Sub

Private Sub mnuHelpAbout_Click()
    Call ShowAbout(Me, App.Title, App.ProductName, App.Major & "." & App.Minor & "." & App.Revision)
End Sub

Private Sub mnuHelpTopic_Click()
   Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
End Sub

Private Sub mnuHelpWebHome_Click()
    Call zlHomePage(Me.hWnd)
End Sub

Private Sub mnuHelpWebMail_Click()
    Call zlMailTo(Me.hWnd)
End Sub

Private Sub mnuMedicalBegin_Click()
    Call MenuClick("�������")
End Sub


Private Sub mnuMedicalGroupAdd_Click()
    Call MenuClick("��ӳ�Ա")
End Sub

Private Sub mnuMedicalItems_Click()
    Call MenuClick("�����Ŀ")
End Sub

Private Sub mnuMedicalItemsAddtion_Click()
    
    If mintIndex <> 1 Then Exit Sub
    
    If Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "��־"))) = 98 Then
        Call MenuClick("�����Ŀ")
    Else
        Call MenuClick("������Ŀ")
    End If
    
End Sub


Private Sub mnuViewPatientBrowse_Click()
    Call MenuClick("��Ա��Ϣ")
End Sub


Private Sub mnuReportWrite_Click()
    Call MenuClick("��д����")
End Sub

Private Sub mnuReportWriteMuli_Click()
    Call MenuClick("������д")
End Sub

Private Sub mnuViewRefresh_Click()
    
    Dim intRow As Integer
    Dim lngLoop As Long
    Dim rs As New ADODB.Recordset
    Dim blnSingle As Boolean
    
    If mintIndex >= 3 Then Exit Sub
    
    zlCommFun.ShowFlash "���Ժ�����ˢ������...", Me
    DoEvents
    
    mblnNoAllowChange = True
    
    intRow = vsf(mintIndex).Row
    
    usrSave.lng�Ǽ�id = Val(vsf(mintIndex).TextMatrix(intRow, GetCol(vsf(mintIndex), "�Ǽ�id")))
    usrSave.lng����id = Val(vsf(mintIndex).TextMatrix(intRow, GetCol(vsf(mintIndex), "����id")))
    usrSave.str��� = ""
    
    Select Case Val(vsf(mintIndex).TextMatrix(intRow, GetCol(vsf(mintIndex), "��־")))
    Case 0, 98
        blnSingle = True
    End Select
    
    If Val(vsf(mintIndex).TextMatrix(intRow, GetCol(vsf(mintIndex), "��־"))) > 1 Then
        
        gstrSQL = "Select ������� From �����Ա���� Where �Ǽ�id=[1] And ����id=[2]"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, usrSave.lng�Ǽ�id, usrSave.lng����id)
        
        If rs.BOF = False Then
            usrSave.str��� = rs("�������").Value
        End If
        
    End If
    
    LockWindowUpdate vsf(mintIndex).hWnd
    
    Call ClearData("�����Ϣ;���")
    
    Call RefreshData("ԤԼ")
    Call RefreshData("���")
    Call RefreshData("���")
    
    'ѡ��ĳһ���ܼ���Ա
    Call SelectPerson(blnSingle)
    
    Call InheritAppendSpaceRows(mintIndex)
    
    LockWindowUpdate 0
    
    zlCommFun.StopFlash
    
    mblnNoAllowChange = False
    
    mlngSvrKey(mintIndex) = -1
    Call vsf_AfterRowColChange(mintIndex, 0, 0, vsf(mintIndex).Row, vsf(mintIndex).Col)
    
    Call AdjustEnableState
    Call RefreshStateInfo
    
End Sub


Private Sub mnuViewShowResult_Click()
    
    On Error Resume Next
    
    mnuViewShowResult.Checked = Not mnuViewShowResult.Checked
    
    If tbs.SelectedItem.Key = "����" Then
        mfrmActive.ShowResult = mnuViewShowResult.Checked
    End If
    
End Sub

Private Sub mnuViewStatus_Click()
    mnuViewStatus.Checked = Not mnuViewStatus.Checked
    stbThis.Visible = mnuViewStatus.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolButton_Click()
    mnuViewToolButton.Checked = Not mnuViewToolButton.Checked
    mnuViewToolText.Enabled = mnuViewToolButton.Checked
    cbrThis.Visible = mnuViewToolButton.Checked
    Call Form_Resize
End Sub

Private Sub mnuViewToolText_Click()
    Dim intLoop As Integer
    
    mnuViewToolText.Checked = Not mnuViewToolText.Checked
    For intLoop = 1 To tbrThis.Buttons.Count
        tbrThis.Buttons(intLoop).Caption = IIf(mnuViewToolText.Checked, tbrThis.Buttons(intLoop).Tag, "")
    Next
    cbrThis.Bands(1).MinHeight = tbrThis.Height
    Call Form_Resize
    
End Sub

Private Sub mobjPopMenu_MenuBeforeShow(Cancel As Boolean)
    
    Select Case mbytPopMenu
    Case 1
        If mnuChargeAddType(0).Visible Then mobjPopMenu.Add 1, "����" & mnuChargeAddType(0).Caption, , , mnuChargeAddType(0).Enabled
        If mnuChargeAddType(1).Visible Then mobjPopMenu.Add 2, "����" & mnuChargeAddType(1).Caption, , , mnuChargeAddType(1).Enabled
        If mnuChargeAddType(2).Visible Then mobjPopMenu.Add 3, "����" & mnuChargeAddType(2).Caption, , , mnuChargeAddType(2).Enabled

        mobjPopMenu.Add 4, "-", , 2, True
        
        mobjPopMenu.Add 5, mnuChargeModify.Caption, , , mnuChargeModify.Enabled
        mobjPopMenu.Add 6, mnuChargeDelete.Caption, , , mnuChargeDelete.Enabled
    Case 2
        
        For mlngLoop = 0 To mnuReportAddOutLineCase.UBound
            If mnuReportAddOutLineCase(mlngLoop).Caption = "<�޿����ܼ�>" Then
                If mnuReportAddOutLineCase(mlngLoop).Visible Then mobjPopMenu.Add mlngLoop + 1, mnuReportAddOutLineCase(mlngLoop).Caption, , , mnuReportAddOutLineCase(mlngLoop).Enabled And mnuReportAddOutLine.Enabled
            Else
                If mnuReportAddOutLineCase(mlngLoop).Visible Then mobjPopMenu.Add mlngLoop + 1, "����" & mnuReportAddOutLineCase(mlngLoop).Caption, , , mnuReportAddOutLineCase(mlngLoop).Enabled And mnuReportAddOutLine.Enabled
            End If
        Next
        
        mobjPopMenu.Add mnuReportAddOutLineCase.UBound + 1, "-", , 2, True
        
        If mnuReportModifyOutLine.Visible Then mobjPopMenu.Add 101, mnuReportModifyOutLine.Caption, , , mnuReportModifyOutLine.Enabled
        If mnuReportDelOutLine.Visible Then mobjPopMenu.Add 102, mnuReportDelOutLine.Caption, , , mnuReportDelOutLine.Enabled
        
    Case 3
        
        mobjPopMenu.Add 1, "&1.����", , , True, , (lbl(1).Tag = "����")
        mobjPopMenu.Add 2, "&2.�����", , , True, , (lbl(1).Tag = "�����")
        mobjPopMenu.Add 3, "&3.������", , , True, , (lbl(1).Tag = "������")
        mobjPopMenu.Add 4, "&4.���￨��", , , True, , (lbl(1).Tag = "���￨��")
        mobjPopMenu.Add 5, "&5.����ƴ��", , , True, , (lbl(1).Tag = "����ƴ��")
        mobjPopMenu.Add 6, "&6.�������", , , True, , (lbl(1).Tag = "�������")
        mobjPopMenu.Add 7, "&7.���֤��", , , True, , (lbl(1).Tag = "���֤��")
            
        mobjPopMenu.Add 8, "-", , 2, True
        mobjPopMenu.Add 9, "&8.��쵥��", , , True, , (lbl(1).Tag = "��쵥��")
        mobjPopMenu.Add 10, "&9.�����", , , True, , (lbl(1).Tag = "�����")
        mobjPopMenu.Add 11, "&A.�������", , , True, , (lbl(1).Tag = "�������")
        
    Case 4          '����
        If mnuCharge.Visible Then
            If mnuChargeMain.Visible Then mobjPopMenu.Add 1, mnuChargeMain.Caption, , , mnuChargeMain.Enabled
            
            If mnuChargeAddType(0).Visible Or mnuChargeAddType(1).Visible Or mnuChargeAddType(2).Visible Then
                mobjPopMenu.Add 2, "-", , 2, True
            End If
            
            If mnuChargeAddType(0).Visible Then mobjPopMenu.Add 3, "����" & mnuChargeAddType(0).Caption, , , mnuChargeAddType(0).Enabled
            If mnuChargeAddType(1).Visible Then mobjPopMenu.Add 4, "����" & mnuChargeAddType(1).Caption, , , mnuChargeAddType(1).Enabled
            If mnuChargeAddType(2).Visible Then mobjPopMenu.Add 5, "����" & mnuChargeAddType(2).Caption, , , mnuChargeAddType(2).Enabled
            
        End If
    Case 5
        
    End Select
    
End Sub

Private Sub mobjPopMenu_MenuClick(ByVal Key As Long, ByVal Caption As String)
    Select Case mbytPopMenu
    Case 1
        Select Case Key
        Case 1
            Call mnuChargeAddType_Click(0)
        Case 2
            Call mnuChargeAddType_Click(1)
        Case 3
            Call mnuChargeAddType_Click(2)
        Case 5
            Call mnuChargeModify_Click
        Case 6
            Call mnuChargeDelete_Click
        End Select
    Case 2
        If Key <= mnuReportAddOutLineCase.UBound + 1 Then
            Call mnuReportAddOutLineCase_Click(Key - 1)
            Exit Sub
        End If
        
        Select Case Key
        Case 101
            Call mnuReportModifyOutLine_Click
        Case 102
            Call mnuReportDelOutLine_Click
        End Select
    Case 3
    
        Caption = Mid(Caption, 4)
        
        lbl(1).Caption = "&6." & Left(Trim(Caption), Len(Trim(Caption)) - 1)
        lbl(1).Tag = Left(Trim(Caption), Len(Trim(Caption)) - 1)
        
    Case 4          '����
        Select Case Key
        Case 1
            Call mnuChargeMain_Click
        Case 3
            Call mnuChargeAddType_Click(0)
        Case 4
            Call mnuChargeAddType_Click(1)
        Case 5
            Call mnuChargeAddType_Click(2)
        End Select
        
    Case 5
        
    End Select
End Sub

Private Sub picClass_Resize()
    Dim lngCount As Long
    Dim lngLoop As Long
    
    On Error Resume Next
    
    LockWindowUpdate picClass.hWnd
    
    lngCount = cmdKind.UBound - 1
    If cmdKind(3).Visible Then lngCount = cmdKind.UBound
    
    For lngLoop = cmdKind.LBound To lngCount
        cmdKind(lngLoop).Width = picClass.ScaleWidth
        If Val(cmdKind(lngLoop).Tag) = 0 Then
            cmdKind(lngLoop).Top = picClass.ScaleTop + 285 * lngLoop
            picShow.Top = picClass.ScaleTop + 285 * (lngLoop + 1)
        Else
            cmdKind(lngLoop).Top = picClass.ScaleHeight - 285 * (lngCount - lngLoop + 1)
        End If
    Next
    
    picShow.Left = picClass.ScaleLeft - 30
    picShow.Width = picClass.ScaleWidth + 60
    picShow.Height = picClass.ScaleHeight - 285 * (lngCount + 1) + 15
    
    LockWindowUpdate 0
End Sub

Private Sub picContainer_Resize()
    On Error Resume Next
    
    If Not (mfrmActive Is Nothing) Then
        mfrmActive.Width = picContainer.Width
        mfrmActive.Height = picContainer.Height
    End If
End Sub

Private Sub picShow_Resize()
    
    On Error Resume Next
    
    vsf(0).Visible = False
    vsf(1).Visible = False
    vsf(2).Visible = False
    vsf(3).Visible = False
           
    vsf(Val(picShow.Tag)).Visible = True
    
    With vsf(Val(picShow.Tag))
        
        .Left = 0
        .Top = -15
        .Width = picShow.Width
        .Height = picShow.Height + 15
        
        Call InheritAppendSpaceRows(Val(picShow.Tag))
        
    End With
End Sub

Private Sub tbrThis_ButtonClick(ByVal Button As MSComctlLib.Button)
    Dim objPoint As POINTAPI
    
    Call ClientToScreen(tbrThis.hWnd, objPoint)
    
    Select Case Button.Key
    Case "Ԥ��"
        Call mnuFilePrintView_Click
    Case "��ӡ"
        Call mnuFilePrint_Click
    Case "����"
        Call mnuMedicalBegin_Click
    Case "���"
        Call mnuMedicalComplete_Click
    Case "��д"
        Call mnuReportWrite_Click
 
    Case "�ܼ�"
        
        mbytPopMenu = 2
        
        Set mobjPopMenu = New clsPopMenu
        Call mobjPopMenu.ShowPopupMenu(objPoint.X * 15 + Button.Left - 15, objPoint.Y * 15 + Button.Top + Button.Height + 15)
        
     Case "����"
     
        Call mnuChargeMain_Click
        
    Case "����"
    
        mbytPopMenu = 1
        Set mobjPopMenu = New clsPopMenu
                
        Call mobjPopMenu.ShowPopupMenu(objPoint.X * 15 + Button.Left - 15, objPoint.Y * 15 + Button.Top + Button.Height + 15)
        
    Case "����"
        
        If mnuViewFilter.Visible And mnuViewFilter.Enabled Then Call mnuViewFilter_Click
        
    Case "����"
        Call mnuHelpTopic_Click
    Case "�˳�"
        Call mnuFileExit_Click
    End Select
End Sub

Private Sub tbrThis_ButtonDropDown(ByVal Button As MSComctlLib.Button)
    
    Call tbrThis_ButtonClick(Button)
    
End Sub

Private Sub tbrThis_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu mnuViewTool
End Sub

Private Sub tbs_Click()
    Dim lngKey As Long
    Dim lngStyle As Long
    Dim blnShowed As Boolean
    Dim lng�Ǽ�id As Long
    Dim str��� As String
    
    blnShowed = False
    picContainer.BorderStyle = 0
    
    Select Case tbs.SelectedItem.Key
    Case "����"
        If TypeName(mfrmActive) = "frmMedicalStationReport" Then
            blnShowed = True
        Else
            Call ResetActiveForm
            Set mfrmActive = frmMedicalStationReport
        End If
        
        mfrmActive.ShowResult = mnuViewShowResult.Checked
    Case "�ܼ�"
        
        Call ResetActiveForm
        
        picContainer.BorderStyle = 1
        
        Set mfrmActive = mclsCore.ShowFileObject(Me, picContainer, 0, 0, gcnOracle, "", glngSys, "", "")
        
        Call mfrmActive.zlMenuClick(Me, mlng��첡��id, "ˢ��")
        
        Call AdjustEnableState
        Exit Sub
    Case "����"

        If TypeName(mfrmActive) = "frmMedicalStationCharge" Then
            blnShowed = True
        Else
            Call ResetActiveForm
            Set mfrmActive = frmMedicalStationCharge
        End If
    Case "����"
        If TypeName(mfrmActive) = "frmMedicalStationHistory" Then
            blnShowed = True
        Else
            Call ResetActiveForm
            Set mfrmActive = frmMedicalStationHistory
        End If
    Case "�ſ�"
        If TypeName(mfrmActive) = "frmMedicalStationGroup" Then
            blnShowed = True
        Else
            Call ResetActiveForm
            Set mfrmActive = frmMedicalStationGroup
        End If
    End Select
    
    '���ص�ǰ�����
    
    If Not mfrmActive Is Nothing Then
        If blnShowed = False Then
            
            Load mfrmActive
            
            lngStyle = GetWindowLong(mfrmActive.hWnd, GWL_STYLE)
            Call SetWindowLong(mfrmActive.hWnd, GWL_STYLE, lngStyle Or WS_CHILD)
            Call SetParent(mfrmActive.hWnd, picContainer.hWnd)
            Call MoveWindow(mfrmActive.hWnd, 0, 0, picContainer.ScaleWidth / Screen.TwipsPerPixelX, picContainer.ScaleHeight / Screen.TwipsPerPixelY, 1)
            
            mfrmActive.Show
            DoEvents
            
        End If
        
        'ˢ������
        On Error Resume Next
        
        str��� = ""
        lngKey = Val(vsf(mintIndex).RowData(vsf(mintIndex).Row))
        
        If mintIndex <> 3 Then
            If Trim(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "[����]"))) <> "" Then
                lngKey = 0
                lng�Ǽ�id = Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "�Ǽ�id")))
            End If
            
            Select Case Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "��־")))
            Case 2               '0-���˷�����;1-����������;2-���������;99-�ܼ���Ա��
            
                str��� = vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "����"))
            End Select
            
        End If
        
        Select Case UCase(tbs.SelectedItem.Key)
        Case "����"
            Call mfrmActive.zlMenuClick(Me, "ˢ��", CStr(lngKey) & "'" & mintIndex)
        Case "����"
            
            Dim strStart As String
            Dim strEnd As String
            
            strStart = GetDateTime(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "������췶Χ", "��  ��"), 1)
            strEnd = GetDateTime(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & Me.Name, "������췶Χ", "��  ��"), 2)
            If strStart = "" Then strStart = GetDateTime("��  ��", 1)
            If strEnd = "" Then strEnd = GetDateTime("��  ��", 2)
            
            Call mfrmActive.zlMenuClick(Me, "ˢ��", CStr(Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "����id")))) & "'" & strStart & "'" & strEnd)
            
        Case "�ܼ�", "����"
            Call mfrmActive.zlMenuClick(Me, lngKey, "ˢ��")
        Case "�ſ�"
            Call mfrmActive.zlMenuClick(Me, "ˢ��", CStr(lng�Ǽ�id) & "'" & str���)
        Case Else
            
        End Select
                
    End If
    
    Call AdjustEnableState
End Sub

Private Sub tmr_Timer()
    Dim strSvrKey As String
    
    mlngCountTmr = mlngCountTmr + 1
    
    If mlngCountTmr >= Val(tmr.Tag) Then
    
        'ʱ�䵽�ˣ���ʼ����
        mlngCountTmr = 0
        
        mblnNoAllowChange = True
        strSvrKey = SaveRow(vsf(mintIndex))
        
        LockWindowUpdate vsf(mintIndex).hWnd
        
        If mintIndex < 2 Then Call ClearData("�����Ϣ;���")
                
        Call RefreshData("ԤԼ")
        Call RefreshData("���")
        
        Call InheritAppendSpaceRows(mintIndex)
                
        LockWindowUpdate 0
        
        mblnNoAllowChange = False
        
        If mintIndex < 2 Then
            
            Call InheritRestoreRow(vsf(mintIndex), Val(strSvrKey))
            
            mlngSvrKey(mintIndex) = -1
            Call vsf_AfterRowColChange(mintIndex, 0, 0, vsf(mintIndex).Row, vsf(mintIndex).Col)
            
        End If
                
    End If
End Sub

Private Sub txt_GotFocus(Index As Integer)
    zlControl.TxtSelAll txt(Index)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim strCol As String
    Dim lngCol As Long
    Dim strSQL As String
    Dim lngRow As Long
    Dim blnCard As Boolean
    Dim strStart As String
    Dim strEnd As String
    
    If Chr(KeyAscii) = "'" Then KeyAscii = 0
    strCol = Mid(lbl(1).Caption, 4)
    lngCol = GetCol(vsf(mintIndex), strCol)
            
    If strCol = "���￨��" And mintIndex <> 3 And KeyAscii <> vbKeyReturn Then
        '���￨�ţ��Զ�ʶ��

        blnCard = InputIsCard(txt(Index).Text, KeyAscii)

        If blnCard And Len(txt(Index).Text) = ParamInfo.���￨���볤�� - 1 And KeyAscii <> 8 And txt(Index).Text <> "" Then
            If KeyAscii <> 13 Then
                txt(Index).Text = txt(Index).Text & Chr(KeyAscii)
                txt(Index).SelStart = Len(txt(Index).Text)
            End If
            KeyAscii = vbKeyReturn

        End If

    End If
            
    If KeyAscii = vbKeyReturn And mintIndex <> 3 Then
        
        If mintIndex = 1 And mint������ѯ���� = 1 Then
            
            Select Case strCol
            Case "�������"
                strSQL = "Select * From (Select a.�����,A.ID,A.�������,A.�Ǽ�id,A.����id,B.�Ƿ�����,b.����,c.����,c.�����,c.������,c.���֤��,c.���￨��,D.���� As ������� From �����Ա���� A,���ǼǼ�¼ B,������Ϣ C,��Լ��λ D " & _
                                    "Where B.���״̬=Decode([4],0,2,1,4,2,5) AND C.����id=A.����id AND A.�Ǽ�id=B.ID AND Nvl(b.�Ƿ�����,0)=1 AND a.���ʱ�� BETWEEN [5] AND [6] And b.��Լ��λid=D.ID(+) " & _
                                    ")"
            Case Else
                strSQL = "Select * From (Select a.�����,A.ID,A.�������,A.�Ǽ�id,A.����id,B.�Ƿ�����,b.����,c.����,c.�����,c.������,c.���֤��,c.���￨�� From �����Ա���� A,���ǼǼ�¼ B,������Ϣ C " & _
                                    "Where B.���״̬=Decode([4],0,2,1,4,2,5) AND C.����id=A.����id AND A.�Ǽ�id=B.ID AND Nvl(b.�Ƿ�����,0)=1 AND a.���ʱ�� BETWEEN [5] AND [6] " & _
                                    "Union All " & _
                                    "Select a.�����,A.ID,A.�������,A.�Ǽ�id,A.����id,B.�Ƿ�����,b.����,c.����,c.�����,c.������,c.���֤��,c.���￨�� From �����Ա���� A,���ǼǼ�¼ B,������Ϣ C " & _
                                    "Where B.���״̬=Decode([4],0,2,1,4,2,5) AND C.����id=A.����id AND A.�Ǽ�id=B.ID AND Nvl(b.�Ƿ�����,0)=0 AND a.���ʱ�� BETWEEN [1] AND [2])"
            End Select
        Else
            Select Case strCol
            Case "�������"
                strSQL = "Select * From (Select a.�����,A.ID,A.�������,A.�Ǽ�id,A.����id,B.�Ƿ�����,b.����,c.����,c.�����,c.������,c.���֤��,c.���￨��,D.���� As ������� From �����Ա���� A,���ǼǼ�¼ B,������Ϣ C,��Լ��λ D " & _
                                        "Where B.���״̬=Decode([4],0,2,1,4,2,5) AND C.����id=A.����id AND A.�Ǽ�id=B.ID AND Nvl(b.�Ƿ�����,0)=1 And b.���ʱ�� BETWEEN [5] AND [6] And b.��Լ��λid=D.ID(+) " & _
                                        ") "
            Case Else
                strSQL = "Select * From (Select a.�����,A.ID,A.�������,A.�Ǽ�id,A.����id,B.�Ƿ�����,b.����,c.����,c.�����,c.������,c.���֤��,c.���￨�� From �����Ա���� A,���ǼǼ�¼ B,������Ϣ C " & _
                                        "Where B.���״̬=Decode([4],0,2,1,4,2,5) AND C.����id=A.����id AND A.�Ǽ�id=B.ID AND Nvl(b.�Ƿ�����,0)=1 And b.���ʱ�� BETWEEN [5] AND [6]  " & _
                                        "Union All " & _
                                        "Select a.�����,A.ID,A.�������,A.�Ǽ�id,A.����id,B.�Ƿ�����,b.����,c.����,c.�����,c.������,c.���֤��,c.���￨�� From �����Ա���� A,���ǼǼ�¼ B,������Ϣ C " & _
                                        "Where B.���״̬=Decode([4],0,2,1,4,2,5) AND C.����id=A.����id AND A.�Ǽ�id=B.ID AND Nvl(b.�Ƿ�����,0)=0 And b.���ʱ�� BETWEEN [1] AND [2]) "
            End Select
            
        End If
        
        If Index = 1 And Trim(txt(Index).Text) <> "" Then
            
            Call txt_LostFocus(Index)
            
            If mintIndex = 1 Then
                strStart = Split(mstr���������ʱ�䷶Χ, "|")(0)
                strEnd = Split(mstr���������ʱ�䷶Χ, "|")(1)
            Else
                strStart = Split(vsf(mintIndex).Tag, "|")(0)
                strEnd = Split(vsf(mintIndex).Tag, "|")(1)
            End If
            
            
            If mstrSvrFind <> txt(Index).Text Then
                
                mstrSvrFind = txt(Index).Text

                Select Case strCol
                    Case "��쵥��"
                    
                        strSQL = strSQL & " Where ���� Like [3] Order By �Ƿ�����,�Ǽ�id "
                        Set mrsFind = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(Split(vsf(mintIndex).Tag, "|")(0)), CDate(Split(vsf(mintIndex).Tag, "|")(1)), "%" & txt(Index).Text & "%", mintIndex, CDate(strStart), CDate(strEnd))
                                                
                    Case "�����", "������", "���￨��"
                        strSQL = strSQL & " Where " & strCol & " = [3] Order By �Ƿ�����,�Ǽ�id "
                        
                        Set mrsFind = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(Split(vsf(mintIndex).Tag, "|")(0)), CDate(Split(vsf(mintIndex).Tag, "|")(1)), Val(txt(Index).Text), mintIndex, CDate(strStart), CDate(strEnd))
            
                    Case "���֤��"
                        
                        strSQL = strSQL & " Where ���֤��=[3] Order By �Ƿ�����,�Ǽ�id "
                        Set mrsFind = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(Split(vsf(mintIndex).Tag, "|")(0)), CDate(Split(vsf(mintIndex).Tag, "|")(1)), txt(Index).Text, mintIndex, CDate(strStart), CDate(strEnd))
                        
                    Case "����ƴ��"
                        
                        strSQL = strSQL & " Where zlSpellCode(����) Like [3] Order By �Ƿ�����,�Ǽ�id "
                        Set mrsFind = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(Split(vsf(mintIndex).Tag, "|")(0)), CDate(Split(vsf(mintIndex).Tag, "|")(1)), "%" & UCase(txt(Index).Text) & "%", mintIndex, CDate(strStart), CDate(strEnd))
                        
                    Case "�������"
                        strSQL = strSQL & " Where zlWBCode(����) Like [3] Order By �Ƿ�����,�Ǽ�id "
                
                        Set mrsFind = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(Split(vsf(mintIndex).Tag, "|")(0)), CDate(Split(vsf(mintIndex).Tag, "|")(1)), "%" & UCase(txt(Index).Text) & "%", mintIndex, CDate(strStart), CDate(strEnd))
                        
                    
                    Case "�������"
                        
                        strSQL = strSQL & " Where ������� Like [3] Order By �Ƿ�����,�Ǽ�id "
                        Set mrsFind = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(Split(vsf(mintIndex).Tag, "|")(0)), CDate(Split(vsf(mintIndex).Tag, "|")(1)), "%" & UCase(txt(Index).Text) & "%", mintIndex, CDate(strStart), CDate(strEnd))
                                                           
                                                           
                    Case Else
                    
                        strSQL = strSQL & " Where " & strCol & " Like [3] Order By �Ƿ�����,�Ǽ�id "
                        Set mrsFind = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(Split(vsf(mintIndex).Tag, "|")(0)), CDate(Split(vsf(mintIndex).Tag, "|")(1)), "%" & txt(Index).Text & "%", mintIndex, CDate(strStart), CDate(strEnd))
                        
                End Select

                If mrsFind.BOF Then
                    ShowSimpleMsg "û���ҵ�����Ҫ�����Ϣ��"
                    txt(Index).Text = ""
                    Exit Sub
                End If
            End If
            
            If mrsFind.EOF And mrsFind.RecordCount > 0 Then mrsFind.MoveFirst
            If Not mrsFind.EOF Then
                
                usrSave.lng�Ǽ�id = mrsFind("�Ǽ�id").Value
                usrSave.lng����id = mrsFind("����id").Value
                usrSave.str��� = mrsFind("�������").Value
                
                Call SelectPerson(IIf(mrsFind("�Ƿ�����") = 1, False, True))
                
            End If
            
            On Error Resume Next
            Err = 0
            mrsFind.MoveNext
            If Err <> 0 Then ShowSimpleMsg "�Ѿ������꣬���ٲ��ҽ���������һ�Σ�"
            
        End If
        
        txt(Index).SetFocus
        zlControl.TxtSelAll txt(Index)
    Else
        If Index = 1 Then
            Select Case lbl(1).Tag
            Case "��쵥��", "���￨��"
                KeyAscii = Asc(UCase(Chr(KeyAscii)))
            End Select
            
        End If
        
        
    End If
End Sub

Private Sub txt_LostFocus(Index As Integer)
    
    If lbl(1).Tag = "��쵥��" Then
        Dim intYear As Integer
        Dim strYear As String
        '�Զ����뵥�ݺ�
        If (UCase(Left(txt(Index).Text, 1)) < "A" Or UCase(Left(txt(Index).Text, 1)) > "Z") And Trim(txt(Index).Text) <> "" Then
            intYear = Format(zlDatabase.Currentdate, "YYYY") - 1990
            strYear = IIf(intYear < 10, CStr(intYear), Chr(55 + intYear))
            txt(Index).Text = strYear & Right("0000000" & txt(Index).Text, 7)
        End If
    End If
End Sub

Private Sub vsf_AfterCollapse(Index As Integer, ByVal Row As Long, ByVal State As Integer)
    Call InheritAppendSpaceRows(Index)
End Sub

Private Sub vsf_AfterRowColChange(Index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If mblnNoAllowChange Then Exit Sub
    
    If OldRow = NewRow Then Exit Sub
    
    Call ClearData("���")
        
    mlngSvrKey(Index) = Val(vsf(Index).RowData(NewRow))
    
    '��ȡϸ��
    Call RefreshData("����")
    
    If mintIndex = 1 Or mintIndex = 2 Then
        If Val(vsf(mintIndex).TextMatrix(vsf(mintIndex).Row, GetCol(vsf(mintIndex), "����id"))) > 0 Then
            If tbs.Tabs(1).Key <> "����" Then
                tbs.Tabs.Clear
                tbs.Tabs.Add , "����", "&1.����"
                tbs.Tabs.Add , "�ܼ�", "&2.�ܼ�"
                tbs.Tabs.Add , "����", "&3.����"
                tbs.Tabs.Add , "����", "&4.����"
            End If
        Else
            If tbs.Tabs(1).Key <> "�ſ�" Then
                tbs.Tabs.Clear
                tbs.Tabs.Add , "�ſ�", "&1.�ſ�"
            End If
        End If
    Else
        If tbs.Tabs(1).Key <> "����" Then
            tbs.Tabs.Clear
            tbs.Tabs.Add , "����", "&1.����"
            tbs.Tabs.Add , "�ܼ�", "&2.�ܼ�"
            tbs.Tabs.Add , "����", "&3.����"
            tbs.Tabs.Add , "����", "&4.����"
        End If
    End If
    
    Call tbs_Click
    
    Call AdjustEnableState
    
    On Error Resume Next
    vsf(Index).SetFocus
End Sub

Private Sub vsf_AfterScroll(Index As Integer, ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    Call InheritAppendSpaceRows(Index)
End Sub

Private Sub vsf_AfterUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    Call InheritAppendSpaceRows(Index)
End Sub

Private Sub vsf_BeforeCollapse(Index As Integer, ByVal Row As Long, ByVal State As Integer, Cancel As Boolean)
    Dim lng�Ǽ�id As Long
    Dim str��� As String
    Dim int��־ As Integer

    If Index > 2 Then Exit Sub
    If mblnStartUp Then Exit Sub

    On Error GoTo errHand

    If State = 0 Then
        'չ��,���û��װ��,��װ����Ա����
        int��־ = Val(vsf(Index).TextMatrix(Row, GetCol(vsf(Index), "��־")))
        Select Case int��־
            Case 0, 2              '0-���˷�����;1-����������;2-���������;99-�ܼ���Ա��
                'չ���������
                If Val(vsf(Index).TextMatrix(Row, GetCol(vsf(Index), "�Ƿ�װ��"))) = 0 Then
                    'û��װ�ع�
                    vsf(Index).TextMatrix(Row, GetCol(vsf(Index), "�Ƿ�װ��")) = 1

                    '1.ɾ������
                    vsf(Index).RemoveItem Row + 1

                    '2.װ�ش������Ա�嵥
                    lng�Ǽ�id = Val(vsf(Index).TextMatrix(Row, GetCol(vsf(Index), "�Ǽ�id")))

                    If int��־ = 0 Then
                        str��� = "ȱʡ"
                    Else
                        str��� = vsf(Index).TextMatrix(Row, GetCol(vsf(Index), "����"))
                    End If

                    Select Case Index
                    Case 0
                        Call RefreshData("�����Ա", Row & "'" & lng�Ǽ�id & "'" & str��� & "'2'0'" & vsf(Index).Tag)
                    Case 1
                        Call RefreshData("�����Ա", Row & "'" & lng�Ǽ�id & "'" & str��� & "'4'1'" & vsf(Index).Tag)
                    Case 2
                        Call RefreshData("�����Ա", Row & "'" & lng�Ǽ�id & "'" & str��� & "'5'1'" & vsf(Index).Tag)
                    End Select

                End If
        End Select

    End If

errHand:

End Sub

Private Sub vsf_BeforeUserResize(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Cancel = (Col < 3)
End Sub

Private Sub vsf_DblClick(Index As Integer)
    
    Dim r As Long
    
    On Error GoTo errHand
    
    With vsf(Index)

        r = .Row

        If .IsCollapsed(r) = flexOutlineCollapsed Then

            .IsCollapsed(r) = flexOutlineExpanded

        Else

            .IsCollapsed(r) = flexOutlineCollapsed

        End If

    End With
    Call InheritAppendSpaceRows(Index)
    
    Exit Sub
    
errHand:
    
End Sub

Private Sub vsf_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then vsf_DblClick (Index)
End Sub

Private Sub vsf_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If Index = 3 Then Exit Sub
        Call SendLMouseButton(vsf(Index).hWnd, X, Y)
        If mnuMedical.Visible Then Me.PopupMenu mnuMedical
    End If
    
    If Button = 1 Then
        If vsf(Index).MouseRow = 0 And vsf(Index).MouseCol > 2 Then
        
            mintSort = IIf(mintSort = flexSortGenericAscending, flexSortGenericDescending, flexSortGenericAscending)
            vsf(Index).Sort = mintSort
            
            Set vsf(Index).Cell(flexcpPicture, 0, 3, 0, vsf(Index).Cols - 1) = Nothing
            
            If mintSort = flexSortGenericAscending Then
                vsf(Index).Cell(flexcpPicture, 0, vsf(Index).Col) = ils13.ListImages("up").Picture
            Else
                vsf(Index).Cell(flexcpPicture, 0, vsf(Index).Col) = ils13.ListImages("down").Picture
            End If
        End If
    End If
    
End Sub
    
Public Function LoadOutLineGrid(ByVal intIndex As Integer, ByVal rsData As ADODB.Recordset, Optional ByVal MaskArray As Variant, Optional ByVal blnClear As Boolean = True, Optional ByVal objIls As Object, Optional ByVal blnCharge As Boolean = False) As Boolean
        '------------------------------------------------------------------------------------------------------------------
    '����:������ݵ�����
    '����:
    '����:
    '------------------------------------------------------------------------------------------------------------------
    Dim lngLoop As Long
    Dim strMask As String
    Dim lngRow As Long
    Dim strField As String
    Dim strIcon As String
    Dim blnField As Boolean
    Dim blnForeColor As Boolean
    Dim objMsf As Object

    Dim lngCol����id As Long
    Dim lngCol�ϼ�id As Long
    
    vsf(intIndex).Redraw = False
    
    On Error Resume Next
    
    blnForeColor = (rsData("ǰ��ɫ").Name = "ǰ��ɫ")
    
    On Error GoTo 0
    
    Set objMsf = vsf(intIndex)
    
    lngCol����id = GetCol(objMsf, "����id")
    lngCol�ϼ�id = GetCol(objMsf, "�ϼ�id")
    
    If blnClear Then
        objMsf.Rows = 2
        objMsf.RowData(1) = 0
        For lngLoop = 0 To objMsf.Cols - 1
            objMsf.TextMatrix(1, lngLoop) = ""
        Next
    End If
    
    lngRow = 0
    Do While Not rsData.EOF
        
        lngRow = lngRow + 1
        If objMsf.Rows < lngRow + 1 Then objMsf.Rows = lngRow + 1
        
        On Error Resume Next
        objMsf.RowData(lngRow) = CStr(zlCommFun.NVL(rsData("ID")))
        
        On Error GoTo errHand
        
        For lngLoop = 0 To objMsf.Cols - 1
            
            strField = objMsf.Cell(flexcpData, 0, lngLoop)
            
            If Trim(strField) <> "" Then
            
                On Error Resume Next
                
                strMask = ""
                strMask = MaskArray(lngLoop)
                                        
                On Error GoTo errHand
                
                If Left(strField, 1) = "[" Then
                
                    strField = Mid(strField, 2, Len(strField) - 2)
                    strIcon = ""
                    
                    On Error Resume Next
                    blnField = False
                    blnField = (UCase(rsData(strField).Name) = UCase(strField))
                    If blnField = False Then GoTo NextCol
                    On Error GoTo errHand
                    
                    If Not (objIls Is Nothing) Then
                        strIcon = zlCommFun.NVL(rsData(strField))
                        If strIcon <> "" Then
                            Set objMsf.Cell(flexcpPicture, lngRow, lngLoop) = objIls.ListImages(strIcon).Picture
                        End If
                    End If
                    
                    objMsf.Cell(flexcpData, lngRow, lngLoop) = strIcon
                    objMsf.TextMatrix(lngRow, lngLoop) = strIcon
                Else
                
                    On Error Resume Next
                    blnField = False
                    blnField = (UCase(rsData(strField).Name) = UCase(strField))
                    If blnField = False Then GoTo NextCol
                    On Error GoTo errHand
                    
                     If strMask <> "" Then
                        objMsf.TextMatrix(lngRow, lngLoop) = Format(zlCommFun.NVL(rsData(strField)), strMask)
                    Else
                        objMsf.TextMatrix(lngRow, lngLoop) = zlCommFun.NVL(rsData(strField))
                    End If
                
                    objMsf.Cell(flexcpData, lngRow, lngLoop, lngRow, lngLoop) = objMsf.TextMatrix(lngRow, lngLoop)
                End If
                
            End If
NextCol:
            '��һ��
        Next

        
        If blnForeColor Then objMsf.Cell(flexcpForeColor, lngRow, 0, lngRow, objMsf.Cols - 1) = Val(rsData("ǰ��ɫ").Value)
        
        If Val(vsf(intIndex).TextMatrix(lngRow, lngCol����id)) = 0 Then

            vsf(intIndex).MergeRow(lngRow) = True
            vsf(intIndex).IsSubtotal(lngRow) = True
            
            Select Case Val(vsf(intIndex).TextMatrix(lngRow, GetCol(vsf(intIndex), "��־")))
                Case 0               '0-���˷�����;1-����������;2-���������;99-�ܼ���Ա��
                    vsf(intIndex).Cell(flexcpFontBold, lngRow, 0, lngRow, vsf(intIndex).Cols - 1) = True
                    vsf(intIndex).Cell(flexcpBackColor, lngRow, 0, lngRow, vsf(intIndex).Cols - 1) = COLOR.��ɫ
                    vsf(intIndex).Cell(flexcpText, lngRow, 0, lngRow, 6) = vsf(intIndex).TextMatrix(lngRow, 3)
                    
                    vsf(intIndex).AddItem ""
                    lngRow = lngRow + 1
                Case 2
                    vsf(intIndex).RowOutlineLevel(lngRow) = 1
                    vsf(intIndex).Cell(flexcpBackColor, lngRow, 0, lngRow, vsf(intIndex).Cols - 1) = COLOR.ǳ��ɫ
                    vsf(intIndex).Cell(flexcpText, lngRow, 0, lngRow, 6) = vsf(intIndex).TextMatrix(lngRow, 3)
                    vsf(intIndex).AddItem ""
                    lngRow = lngRow + 1
                Case Else
                    vsf(intIndex).Cell(flexcpFontBold, lngRow, 0, lngRow, vsf(intIndex).Cols - 1) = True
                    vsf(intIndex).Cell(flexcpBackColor, lngRow, 0, lngRow, vsf(intIndex).Cols - 1) = COLOR.��ɫ
                    vsf(intIndex).Cell(flexcpText, lngRow, 0, lngRow, 6) = vsf(intIndex).TextMatrix(lngRow, 3)
            End Select
            

        End If
        
        rsData.MoveNext
    Loop

    vsf(intIndex).Redraw = True

    Call InheritAppendSpaceRows(intIndex)
    
    vsf(intIndex).Outline 1
    vsf(intIndex).Outline 0
    
    Exit Function
    
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub mnuHelpWebForum_Click()
    '-----------------------------------------------------------------------------
    '����:���ӵ�������̳
    '�޸���:���˺�
    '�޸�����:2006-12-11
    '-----------------------------------------------------------------------------
    Call zlWebForum(Me.hWnd)
End Sub

