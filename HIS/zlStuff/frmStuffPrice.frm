VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Begin VB.Form frmStuffPrice 
   Caption         =   "���ϵ��۵�"
   ClientHeight    =   7230
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11265
   Icon            =   "frmStuffPrice.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   ScaleHeight     =   7230
   ScaleWidth      =   11265
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picStoceBack 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   5085
      Left            =   4680
      ScaleHeight     =   5085
      ScaleWidth      =   8685
      TabIndex        =   32
      Top             =   2520
      Width           =   8685
      Begin VB.PictureBox picPay 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1890
         Left            =   240
         ScaleHeight     =   1890
         ScaleWidth      =   7755
         TabIndex        =   39
         Top             =   3120
         Width           =   7755
         Begin VB.CheckBox chk�Զ����� 
            Caption         =   "�Զ����ݿ�����Ӧ���䶯���"
            Height          =   195
            Left            =   0
            TabIndex        =   40
            Top             =   135
            Width           =   2985
         End
         Begin VSFlex8Ctl.VSFlexGrid vsPay 
            Height          =   1440
            Left            =   0
            TabIndex        =   41
            Top             =   480
            Width           =   6735
            _cx             =   11880
            _cy             =   2540
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
            BackColorBkg    =   -2147483634
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
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
            FormatString    =   $"frmStuffPrice.frx":058A
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
      Begin VB.PictureBox picStoce 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   2850
         Left            =   360
         ScaleHeight     =   2850
         ScaleWidth      =   8190
         TabIndex        =   33
         Top             =   720
         Width           =   8190
         Begin VB.CheckBox chk��ʾ���в��� 
            Caption         =   "��ʾ��ǰ������������"
            Height          =   270
            Left            =   3630
            TabIndex        =   36
            Top             =   0
            Width           =   2150
         End
         Begin VB.CheckBox chk���� 
            Caption         =   "���ⷿ���θ���"
            Height          =   210
            Left            =   135
            TabIndex        =   37
            Top             =   75
            Width           =   1620
         End
         Begin VB.CheckBox chkӦ�� 
            Caption         =   "Ӧ���ʿ����"
            Height          =   195
            Left            =   1845
            TabIndex        =   35
            Top             =   60
            Width           =   1635
         End
         Begin VB.CommandButton cmdPrintStoce 
            Caption         =   "��ӡ���䶯��(&S)��"
            Height          =   350
            Left            =   4800
            Picture         =   "frmStuffPrice.frx":0676
            TabIndex        =   34
            Top             =   360
            Width           =   1965
         End
         Begin VSFlex8Ctl.VSFlexGrid vsStoce 
            Height          =   2340
            Left            =   480
            TabIndex        =   38
            Top             =   840
            Width           =   6510
            _cx             =   11483
            _cy             =   4128
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
            BackColorBkg    =   -2147483634
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
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
            Cols            =   14
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmStuffPrice.frx":07C0
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
         Height          =   2445
         Left            =   240
         TabIndex        =   42
         Top             =   300
         Width           =   7950
         _Version        =   589884
         _ExtentX        =   14023
         _ExtentY        =   4313
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picSeach 
      Appearance      =   0  'Flat
      BackColor       =   &H80000004&
      ForeColor       =   &H80000008&
      Height          =   5040
      Left            =   150
      ScaleHeight     =   5010
      ScaleWidth      =   4395
      TabIndex        =   0
      Top             =   1725
      Visible         =   0   'False
      Width           =   4425
      Begin VB.Frame fraCost 
         Caption         =   "�ɱ��۵���"
         Height          =   1005
         Left            =   120
         TabIndex        =   11
         Top             =   3120
         Width           =   4125
         Begin VB.CommandButton cmdPriver 
            Caption         =   "��"
            Height          =   270
            Left            =   3800
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   225
            Width           =   255
         End
         Begin VB.TextBox txtPriver 
            Height          =   300
            Left            =   705
            TabIndex        =   13
            Top             =   210
            Width           =   3090
         End
         Begin VB.TextBox txt�ӳ��� 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   705
            TabIndex        =   16
            Top             =   600
            Width           =   735
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "��Ӧ��"
            Height          =   180
            Left            =   90
            TabIndex        =   12
            Top             =   255
            Width           =   540
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   1485
            TabIndex        =   17
            Top             =   645
            Width           =   225
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "�ӳ���"
            Height          =   180
            Index           =   1
            Left            =   105
            TabIndex        =   15
            Top             =   660
            Width           =   540
         End
      End
      Begin VB.CommandButton cmd���� 
         Caption         =   "���������˵���(&R)"
         Height          =   350
         Left            =   2040
         TabIndex        =   18
         Top             =   4440
         Width           =   2250
      End
      Begin VB.Frame fra������ 
         Caption         =   "������ʽ"
         Height          =   1155
         Left            =   120
         TabIndex        =   6
         Top             =   1800
         Width           =   4125
         Begin VB.TextBox txt������ 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   2760
            TabIndex        =   8
            Top             =   270
            Width           =   735
         End
         Begin VB.ComboBox cbo������ʽ 
            Height          =   300
            Left            =   150
            Style           =   2  'Dropdown List
            TabIndex        =   7
            Top             =   270
            Width           =   2580
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            Caption         =   "��"
            BeginProperty Font 
               Name            =   "����"
               Size            =   10.5
               Charset         =   134
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   210
            Left            =   3540
            TabIndex        =   9
            Top             =   315
            Width           =   225
         End
         Begin VB.Label lblInfor 
            Caption         =   "���ݳɱ��ۣ������µļӳ������¼ӳɵ���"
            Height          =   255
            Left            =   165
            TabIndex        =   10
            Top             =   690
            Width           =   3660
         End
      End
      Begin VB.Frame fra 
         Caption         =   "Ӧ�÷�Χ"
         Height          =   1005
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   4125
         Begin VB.OptionButton optӦ�� 
            Caption         =   "���ƶ�Ʒ������(&2)"
            Height          =   375
            Index           =   2
            Left            =   120
            TabIndex        =   43
            Top             =   600
            Width           =   1860
         End
         Begin VB.OptionButton optӦ�� 
            Caption         =   "��ǰ��������������(&1)"
            Height          =   375
            Index           =   1
            Left            =   1845
            TabIndex        =   5
            Top             =   255
            Width           =   2205
         End
         Begin VB.OptionButton optӦ�� 
            Caption         =   "��ǰ��������(&0)"
            Height          =   375
            Index           =   0
            Left            =   120
            TabIndex        =   4
            Top             =   255
            Width           =   1740
         End
      End
      Begin VB.CommandButton cmdType 
         Caption         =   "��"
         Height          =   270
         Left            =   3960
         TabIndex        =   2
         Top             =   150
         Width           =   255
      End
      Begin VB.TextBox txt���� 
         Height          =   300
         Left            =   645
         TabIndex        =   1
         Top             =   135
         Width           =   3420
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Index           =   0
         Left            =   120
         TabIndex        =   19
         Top             =   180
         Width           =   360
      End
   End
   Begin VB.PictureBox picPrice 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3450
      Left            =   4680
      ScaleHeight     =   3450
      ScaleWidth      =   10065
      TabIndex        =   21
      Top             =   -360
      Width           =   10065
      Begin VB.CheckBox chkAppAllColumn 
         Caption         =   "�޸ļ۸�Ӧ����������"
         Height          =   255
         Left            =   480
         TabIndex        =   44
         Top             =   360
         Width           =   2295
      End
      Begin VB.PictureBox picBakDown 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   810
         Left            =   0
         ScaleHeight     =   810
         ScaleWidth      =   8850
         TabIndex        =   23
         TabStop         =   0   'False
         Top             =   2400
         Width           =   8850
         Begin VB.CheckBox Chk���� 
            Caption         =   "ʱ�۲��ϸ�Ϊ��������(&D)"
            Enabled         =   0   'False
            Height          =   210
            Left            =   2505
            TabIndex        =   27
            Top             =   525
            Width           =   2370
         End
         Begin VB.TextBox txt������ 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   300
            Left            =   6285
            Locked          =   -1  'True
            TabIndex        =   26
            Top             =   90
            Width           =   2445
         End
         Begin VB.TextBox txt˵�� 
            Height          =   300
            Left            =   825
            TabIndex        =   25
            Top             =   90
            Width           =   4485
         End
         Begin VB.CheckBox chk����ִ�� 
            Caption         =   "���м۸�������Ч(&I)"
            Height          =   210
            Left            =   75
            TabIndex        =   24
            Top             =   525
            Width           =   2040
         End
         Begin MSComCtl2.DTPicker dtpִ������ 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "yyyy-MM-dd"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
            Height          =   300
            Left            =   6285
            TabIndex        =   28
            Top             =   465
            Width           =   2445
            _ExtentX        =   4313
            _ExtentY        =   529
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
            Format          =   184418307
            CurrentDate     =   36846.5833333333
         End
         Begin VB.Label lblValuer 
            AutoSize        =   -1  'True
            Caption         =   "������"
            Height          =   180
            Left            =   5655
            TabIndex        =   31
            Top             =   150
            Width           =   540
         End
         Begin VB.Label lblRunDate 
            AutoSize        =   -1  'True
            Caption         =   "ִ������"
            Height          =   180
            Left            =   5475
            TabIndex        =   30
            Top             =   525
            Width           =   720
         End
         Begin VB.Label lblSummary 
            AutoSize        =   -1  'True
            Caption         =   "����˵��"
            Height          =   180
            Left            =   30
            TabIndex        =   29
            Top             =   150
            Width           =   720
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsPrice 
         Height          =   2190
         Left            =   0
         TabIndex        =   22
         Top             =   600
         Width           =   10665
         _cx             =   18812
         _cy             =   3863
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
         BackColorBkg    =   -2147483634
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   18
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmStuffPrice.frx":09A2
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
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   20
      Top             =   6870
      Width           =   11265
      _ExtentX        =   19870
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmStuffPrice.frx":0C26
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ��������Ϣ��ҵ��˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   14790
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
   Begin MSComctlLib.ImageList imlPaneIcons 
      Left            =   10140
      Top             =   4155
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   65280
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPrice.frx":14BA
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmStuffPrice.frx":180E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin XtremeCommandBars.ImageManager imgPublic 
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      Icons           =   "frmStuffPrice.frx":1B62
   End
   Begin XtremeCommandBars.CommandBars cbsMain 
      Left            =   1530
      Top             =   150
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
   End
   Begin XtremeDockingPane.DockingPane DkPane 
      Bindings        =   "frmStuffPrice.frx":1AD48
      Left            =   0
      Top             =   0
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmStuffPrice"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mlngBillId As Long                '��������:0-���۴���;����-��ʾmlngBillIdȷ������ʷ���۵�
Private mlngStuffId As Long                '��������:0-δָ����������;����-����ʱֱ����ʾmlngStuffId��ԭ�۸����
Private Enum ���۷�ʽ
        T_�ۼ۵��� = 1
        T_�ɱ��۵��� = 2
        T_�ɱ����ۼ۵��� = 3
End Enum
Private m���۷�ʽ As ���۷�ʽ

Public Enum BillType
    B_��һ���� = 0
    B_�������� = 1              '�������ķ����������е���
    B_���� = 2
End Enum
'---------------------------------

Private mintUnit As Integer      '�Ƿ��Կⷿ��λ��ʾ
Private mblnModify As Boolean
Private mblnFirst  As Boolean
Private mBillType As BillType
Private mblnSucces As Boolean
Private mlngPreRow As Long
Private mlngPrice As Long

'----------------------------------------------------------------------------------------------------------
'���˺�:����С��λ���ĸ�ʽ��
'�޸�:2007/03/06
Private mFMT As g_FmtString
Private mOraFMT As g_FmtString
'----------------------------------------------------------------------------------------------------------
Private Const conMenu_Popup = 1           '������
Private Const conMenu_Preview = 102         'Ԥ��(&V)
Private Const conMenu_Print = 103           '��ӡ(&P)
Private Const conMenu_Excel = 104           '�����&Excel��
Private Const conMenu_Save = 305           '����
Private Const conMenu_Cancel = 304           'ȡ��
Private Const conMenu_Lable = 300           '�Ƽ۷�ʽ����
Private Const conMenu_Combo = 301           '�Ƽ۷�ʽCOMBOX


Private Const conMenu_Help_Help = 901        '��������(&H)


'CommandBar�����ȼ�
Private Const FSHIFT = 4
Private Const FCONTROL = 8
Private Const FALT = 16

Private Const ID_PANE_SEARCH = 1
Private Const ID_PANE_PRICE = 2
Private Const ID_PANE_STOCE = 3
Private mobjFindKey As CommandBarControl
 
Private mlngModule As Long
Private mstrPrivs As String
Private mdbl�ӳ��� As Double
Private mlng��Ӧ��ID As Long

'-----------------------------------------------------------------------------------------------------------------
Private Enum mPageNum
    Page_������ = 0
    Page_Ӧ������ = 1
End Enum

Private Sub InitCommandBar()
    '-------------------------------------------------------------------------------------------
    '����:��ʼ���˵�
    '����:
    '����:
    '����:���˺�
    '����:2007/08/07
    '-------------------------------------------------------------------------------------------
    Dim objMenu As CommandBarPopup
    Dim objBar As CommandBar
    Dim objDeptBar As CommandBar
    Dim objPopup As CommandBarPopup
    Dim objControl As CommandBarControl
    Dim cbrCustom As CommandBarControlCustom
    CommandBarsGlobalSettings.App = App
    CommandBarsGlobalSettings.ResourceFile = CommandBarsGlobalSettings.OcxPath & "\XTPResourceZhCn.dll"
    CommandBarsGlobalSettings.ColorManager.SystemTheme = xtpSystemThemeAuto
    
    cbsMain.VisualTheme = xtpThemeOffice2003
    With Me.cbsMain.Options
    
        .ShowExpandButtonAlways = False
        .ToolBarAccelTips = True
        .AlwaysShowFullMenus = False
        .IconsWithShadow = True '����VisualTheme����Ч
        .UseDisabledIcons = True
        .LargeIcons = True
        .SetIconSize True, 24, 24
        .SetIconSize False, 16, 16
    End With
    cbsMain.EnableCustomization False
    
    
    Set cbsMain.Icons = imgPublic.Icons
    
    '�˵�����:������������
    '    ���xtpControlPopup���͵�����ID���¸�ֵ
    '-----------------------------------------------------
    cbsMain.ActiveMenuBar.Title = "�˵�"
    cbsMain.ActiveMenuBar.EnableDocking xtpFlagStretched + xtpFlagHideWrap
    cbsMain.ActiveMenuBar.Visible = False
  
    Set objBar = cbsMain.Add("������", xtpBarTop)
    objBar.ContextMenuPresent = False
    objBar.ShowTextBelowIcons = False
    objBar.EnableDocking xtpFlagHideWrap Or xtpFlagStretched
    Dim objComBar As CommandBarComboBox
    
    With objBar.Controls
        Set objControl = .Add(xtpControlButton, conMenu_Print, "��ӡ")
        Set objControl = .Add(xtpControlButton, conMenu_Preview, "Ԥ��")
        
        Set objControl = .Add(xtpControlButton, conMenu_Save, "ȷ��"): objControl.IconId = conMenu_Save
            objControl.BeginGroup = True
        Set objControl = .Add(xtpControlButton, conMenu_Cancel, "ȡ��"):    objControl.IconId = conMenu_Cancel
        Set objControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): objControl.BeginGroup = True
        
        If mBillType = B_���� Then
            m���۷�ʽ = T_�ۼ۵���
        Else
        Set objControl = .Add(xtpControlLabel, conMenu_Lable, "���۷�ʽ")
        objControl.Flags = xtpFlagRightAlign
        Set objComBar = .Add(xtpControlComboBox, conMenu_Combo, "���۷�ʽ")
        objComBar.Flags = xtpFlagRightAlign
        Dim intIndex As Integer
        intIndex = 1
        If zlStr.IsHavePrivs(mstrPrivs, "�ۼ۹���") Then
            objComBar.AddItem "���ۼ۵���"
            objComBar.ItemData(intIndex) = 1
            intIndex = intIndex + 1
        End If
        If InStr(1, mstrPrivs, ";�ɱ��۹���;") <> 0 Then
            objComBar.AddItem "���ɱ��۵���"
            objComBar.ItemData(intIndex) = 2
            intIndex = intIndex + 1
        End If
        If intIndex = 3 Then
            objComBar.AddItem "���ۼۺͳɱ��۵���"
            objComBar.ItemData(intIndex) = 3
        End If
        objComBar.ListIndex = 1: objComBar.Width = 120
        m���۷�ʽ = objComBar.ItemData(1)
       End If
   End With

    For Each objControl In objBar.Controls
        If objControl.Type = xtpControlLabel Then
        Else
            objControl.Style = xtpButtonIconAndCaption
        End If
    Next
     
    '����Ŀ����:���������������Ѵ���
    '-----------------------------------------------------
    With cbsMain.KeyBindings
        .Add FCONTROL, vbKeyP, conMenu_Print   '��ӡ
    End With
End Sub

Private Function InitPanel()
    '-----------------------------------------------------------------------------------------------------------
    '����:����������Ϣ
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2008-11-06 12:19:20
    '-----------------------------------------------------------------------------------------------------------
    Dim objPane As Pane, objPaneFind As Pane
    
    With DkPane
        .ImageList = imlPaneIcons '
        
        Set objPaneFind = DkPane.CreatePane(ID_PANE_SEARCH, 400, 400, DockLeftOf, Nothing)
        objPaneFind.Title = "����������������"
        objPaneFind.Options = PaneNoCloseable
        objPaneFind.MinTrackSize.Width = 295
        objPaneFind.MaxTrackSize.Width = 495
        Set objPane = DkPane.CreatePane(ID_PANE_PRICE, 400, 400, DockRightOf, objPaneFind)
        objPane.Title = "������Ϣ"
        objPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picPrice.hwnd
        objPaneFind.Hide
        Set objPane = DkPane.CreatePane(ID_PANE_STOCE, 400, 400, DockBottomOf, objPane)
        objPane.Title = "���䶯��Ϣ"
        objPane.Options = PaneNoCloseable Or PaneNoCaption Or PaneNoFloatable Or PaneNoHideable
        objPane.Handle = picStoceBack.hwnd
        
        .SetCommandBars Me.cbsMain
        .Options.ThemedFloatingFrames = True
        .Options.UseSplitterTracker = False 'ʵʱ�϶�
        .Options.AlphaDockingContext = True
        .Options.HideClient = True
    End With
End Function

Private Function Get������Ŀ() As Boolean
    '--------------------------------------------------------------------------------------------------
    '����:���¸������������ȡ������Ŀ
    '����:��ȡ�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2007/09/18
    '--------------------------------------------------------------------------------------------------
    Dim lng����id As Long
    Dim rsTemp As New ADODB.Recordset
    Dim i As Long
    
    On Error GoTo ErrHandle
    If m���۷�ʽ = T_�ۼ۵��� Then
        mlng��Ӧ��ID = 0
        mdbl�ӳ��� = 0
    Else
        mlng��Ӧ��ID = Val(txtPriver.Tag)
        mdbl�ӳ��� = Val(txt�ӳ���.Text)
    End If
    lng����id = Val(txt����.Tag)
    If lng����id = 0 Then
        ShowMsgBox "δѡ�����,����!"
        zlControl.ControlSetFocus txt����, True
        Exit Function
    End If
   
    gstrSQL = "" & _
    "    Select I.ID, I.����, I.����, I.���, I.����, I.���㵥λ, P.��װ��λ, Decode(I.�Ƿ���, 1, 'ʱ��', '����') ����," & _
    "           P.ָ��������,P.ָ�����ۼ�, P.�ɱ���,  " & IIf(mintUnit = 0, "1", "nvl(p.����ϵ��,1)") & " As ����ϵ��,P.��������" & _
    "    From �շ���ĿĿ¼ I, �������� P, ������ĿĿ¼ M " & _
    "    Where   I.ID = P.����id And P.����id = M.ID  And "
    If optӦ��(2).Value = True Then
        gstrSQL = gstrSQL & " m.id=[1]"
    Else
        If optӦ��(0).Value Then
            gstrSQL = gstrSQL & _
            "          M.����id =[1]"
        Else
            gstrSQL = gstrSQL & _
            "          M.����id In (Select ID From ���Ʒ���Ŀ¼ Start With ID = [1] Connect By Prior id = �ϼ�ID)"
        End If
    End If
    If mlng��Ӧ��ID <> 0 Then
        gstrSQL = gstrSQL & " And exists(Select 1 From ҩƷ��� where I.id=ҩƷid and �ϴι�Ӧ��ID=[2])"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����id, mlng��Ӧ��ID)
    
    With vsPrice
         .Redraw = flexRDNone
         i = 1
         If rsTemp.RecordCount = 0 Then
            .Rows = 2
            For i = 0 To .Cols - 1
                .TextMatrix(1, i) = ""
                .Cell(flexcpData, 1, i) = ""
            Next
            Call InitControl
            .Redraw = flexRDBuffered
            Get������Ŀ = True
            Exit Function
         Else
            .Rows = rsTemp.RecordCount + 1
         End If
         Call InitControl
        .Col = .ColIndex("�ּ�")
         Do While Not rsTemp.EOF
            .TextMatrix(i, .ColIndex("Ʒ��")) = "[" & zlStr.Nvl(rsTemp!����) & "]" & zlStr.Nvl(rsTemp!����)
            .Cell(flexcpData, i, .ColIndex("Ʒ��")) = zlStr.Nvl(rsTemp!Id)
            .TextMatrix(i, .ColIndex("���")) = zlStr.Nvl(rsTemp!���)
            .TextMatrix(i, .ColIndex("����")) = zlStr.Nvl(rsTemp!����)
            .TextMatrix(i, .ColIndex("��λ")) = IIf(mintUnit = 0, zlStr.Nvl(rsTemp!���㵥λ), zlStr.Nvl(rsTemp!��װ��λ))
            .TextMatrix(i, .ColIndex("����")) = zlStr.Nvl(rsTemp!����)
            .Cell(flexcpData, i, .ColIndex("����")) = zlStr.Nvl(rsTemp!��������)

            .TextMatrix(i, .ColIndex("ϵ��")) = zlStr.Nvl(rsTemp!����ϵ��)
            .TextMatrix(i, .ColIndex("ԭ�ɱ���")) = Format(Val(zlStr.Nvl(rsTemp!�ɱ���)) * Val(zlStr.Nvl(rsTemp!����ϵ��)), mFMT.FM_�ɱ���)
            .Cell(flexcpData, i, .ColIndex("ԭ�ɱ���")) = Val(zlStr.Nvl(rsTemp!�ɱ���))
            .TextMatrix(i, .ColIndex("�ֳɱ���")) = Format(Val(zlStr.Nvl(rsTemp!�ɱ���)) * Val(zlStr.Nvl(rsTemp!����ϵ��)), mFMT.FM_�ɱ���)
            .Cell(flexcpData, i, .ColIndex("�ֳɱ���")) = Val(zlStr.Nvl(rsTemp!�ɱ���))
            
            .TextMatrix(i, .ColIndex("ԭ�ɹ��޼�")) = Format(Val(zlStr.Nvl(rsTemp!ָ��������)) * Val(rsTemp!����ϵ��), mFMT.FM_�ɱ���)
            .TextMatrix(i, .ColIndex("�ֲɹ��޼�")) = .TextMatrix(i, .ColIndex("ԭ�ɹ��޼�"))
            .Cell(flexcpData, i, .ColIndex("ԭ�ɹ��޼�")) = Val(zlStr.Nvl(rsTemp!ָ��������))
            .Cell(flexcpData, i, .ColIndex("�ֲɹ��޼�")) = Val(zlStr.Nvl(rsTemp!ָ��������))
            
            .TextMatrix(i, .ColIndex("ָ�����ۼ�")) = Format(Val(zlStr.Nvl(rsTemp!ָ�����ۼ�)) * Val(rsTemp!����ϵ��), mFMT.FM_���ۼ�)
            .TextMatrix(i, .ColIndex("ԭָ���ۼ�")) = .TextMatrix(i, .ColIndex("ָ�����ۼ�"))
            .TextMatrix(i, .ColIndex("��ָ���ۼ�")) = .TextMatrix(i, .ColIndex("ԭָ���ۼ�"))
            
            .Cell(flexcpData, i, .ColIndex("ָ�����ۼ�")) = Val(zlStr.Nvl(rsTemp!ָ�����ۼ�))
            .Cell(flexcpData, i, .ColIndex("ԭָ���ۼ�")) = Val(zlStr.Nvl(rsTemp!ָ�����ۼ�))
            .Cell(flexcpData, i, .ColIndex("��ָ���ۼ�")) = Val(zlStr.Nvl(rsTemp!ָ�����ۼ�))
            
            Call zlGetPrice(Val(zlStr.Nvl(rsTemp!Id)), IIf(.TextMatrix(i, .ColIndex("����")) = "ʱ��", True, False), True, i)
            Call LoadStockData(Val(zlStr.Nvl(rsTemp!Id)), Val(.Cell(flexcpData, i, .ColIndex("ԭ��"))), Val(.Cell(flexcpData, i, .ColIndex("�ּ�"))))
            i = i + 1
            rsTemp.MoveNext
        Loop
        
        '����Ӧ���䶯���
        If (m���۷�ʽ = T_�ɱ��۵��� Or m���۷�ʽ = T_�ɱ����ۼ۵���) Then
            Call RefreshPayData
        End If
        mlngPreRow = 0:
        Call vsPrice_RowColChange
        .Redraw = flexRDBuffered
     End With
     Get������Ŀ = True
     Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function
Private Sub InitOther()
    '------------------------------------------------------------
    '����:��ʼ������������������Ϣ
    '------------------------------------------------------------
    With cbo������ʽ
        .AddItem "���ݳɱ��۰��ӳɵ���"
        .ItemData(.NewIndex) = 1
        .ListIndex = .NewIndex
        .AddItem "�����ۼ۰���������"
        .ItemData(.NewIndex) = 2
        .AddItem "�����ۼ۰��̶�������"
        .ItemData(.NewIndex) = 3
    End With
End Sub

Public Function ShowBill(ByVal frmMain As Form, ByVal EditType As BillType, ByVal lngBillId As Long, ByVal lng����ID As Long) As Boolean
    '--------------------------------------------------------------------------------------------------------------
    '����:��ʾ���۵������
    '����:frmMain-���õĸ�����
    '     lngBillID-����ID
    '     lng����ID-����ID
    '����:���۳ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2007/09/18
    '--------------------------------------------------------------------------------------------------------------
    mlngBillId = lngBillId: mlngStuffId = lng����ID: mBillType = EditType
    
    Me.Show 1, frmMain
    ShowBill = mblnSucces
End Function
  
Private Sub cbo������ʽ_Click()
    If cbo������ʽ.ListIndex < 0 Then Exit Sub
    Select Case cbo������ʽ.ItemData(cbo������ʽ.ListIndex)
    Case 1
        lblInfor.Caption = "���ݳɱ��ۣ������µļӳ������¼ӳɵ���"
        lbl����.Caption = "��"
        txt������.MaxLength = 3
    Case 2
        lblInfor.Caption = "�ڵ�ǰ�ۼۻ����ϰ��ձ�������"
        lbl����.Caption = "��"
        txt������.MaxLength = 3
    Case 3
        lblInfor.Caption = "�ڵ�ǰ�ۼۻ����ϰ��̶����Ӽ�����"
        lbl����.Caption = "Ԫ"
        txt������.MaxLength = 10
    End Select
End Sub

Private Sub cbo������ʽ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub cbsMain_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim ctrCombox As CommandBarComboBox
    
    Select Case Control.Id
    Case conMenu_Preview
        Call printbill(2)
    Case conMenu_Print
        Call printbill(1)
    Case conMenu_Save   '����
         '����������Ϸ���
        If ISValied = False Then Exit Sub
        '�����ʱִ�У�����ù���zl_�����շ���¼_Adjust
        If SaveData() = False Then Exit Sub
        mblnModify = False
        mblnSucces = True
        Unload Me
        Exit Sub
    Case conMenu_Cancel  'ȡ��
        mlngBillId = 0
        mlngStuffId = 0
        Unload Me
    Case conMenu_Combo  '�Ƽ۷�ʽѡ��
        Set ctrCombox = Control
        Select Case ctrCombox.ItemData(ctrCombox.ListIndex)
        Case 1      '�ۼ۵���
            m���۷�ʽ = T_�ۼ۵���
        Case 2      '�ɱ��۵���
            m���۷�ʽ = T_�ɱ��۵���
        Case 3      '�ۼ���ɱ��۵���
            m���۷�ʽ = T_�ɱ����ۼ۵���
        Case Else   '�������Ԥ֪�Ļ������ۼ�Ϊ׼
             m���۷�ʽ = T_�ۼ۵���
        End Select
        Call SetControlVisble
        Call SetColor(m���۷�ʽ)
        Call picStoce_Resize
    Case conMenu_Help_Help  '����
        Call ShowHelp(App.ProductName, Me.hwnd, Me.Name, Int(glngSys / 100))
    End Select
End Sub

Private Sub SetColor(ByVal int��ʽ As Integer)
    Dim intRow As Integer
    Dim intCol As Integer
    
    With vsPrice
        .Cell(flexcpBackColor, 1, 1, .Rows - 1, .Cols - 1) = &H8000000F '��ɫ
        For intRow = 1 To .Rows - 1
            If int��ʽ = 2 Then
                .Cell(flexcpBackColor, 1, .ColIndex("Ʒ��"), .Rows - 1, .ColIndex("Ʒ��")) = &H80000005 ' ��ɫ
                .Cell(flexcpBackColor, 1, .ColIndex("�ֲɹ��޼�"), .Rows - 1, .ColIndex("�ֲɹ��޼�")) = &H80000005 ' ��ɫ
                .Cell(flexcpBackColor, 1, .ColIndex("��ָ���ۼ�"), .Rows - 1, .ColIndex("��ָ���ۼ�")) = &H80000005 ' ��ɫ
                .Cell(flexcpBackColor, 1, .ColIndex("�ֳɱ���"), .Rows - 1, .ColIndex("�ֳɱ���")) = &H80000005 ' ��ɫ
            ElseIf int��ʽ = 3 Then
                .Cell(flexcpBackColor, 1, .ColIndex("Ʒ��"), .Rows - 1, .ColIndex("Ʒ��")) = &H80000005 ' ��ɫ
                .Cell(flexcpBackColor, 1, .ColIndex("�ּ�"), .Rows - 1, .ColIndex("�ּ�")) = &H80000005 ' ��ɫ
                .Cell(flexcpBackColor, 1, .ColIndex("��������"), .Rows - 1, .ColIndex("��������")) = &H80000005 ' ��ɫ
                .Cell(flexcpBackColor, 1, .ColIndex("�ֲɹ��޼�"), .Rows - 1, .ColIndex("�ֲɹ��޼�")) = &H80000005 ' ��ɫ
                .Cell(flexcpBackColor, 1, .ColIndex("��ָ���ۼ�"), .Rows - 1, .ColIndex("��ָ���ۼ�")) = &H80000005 ' ��ɫ
                .Cell(flexcpBackColor, 1, .ColIndex("�ֳɱ���"), .Rows - 1, .ColIndex("�ֳɱ���")) = &H80000005 ' ��ɫ
            Else    '������ʽ�����ۼ۷�ʽ����
                .Cell(flexcpBackColor, 1, .ColIndex("Ʒ��"), .Rows - 1, .ColIndex("Ʒ��")) = &H80000005 ' ��ɫ
                .Cell(flexcpBackColor, 1, .ColIndex("�ּ�"), .Rows - 1, .ColIndex("�ּ�")) = &H80000005 ' ��ɫ
                .Cell(flexcpBackColor, 1, .ColIndex("��������"), .Rows - 1, .ColIndex("��������")) = &H80000005 ' ��ɫ
                .Cell(flexcpBackColor, 1, .ColIndex("�ֲɹ��޼�"), .Rows - 1, .ColIndex("�ֲɹ��޼�")) = &H80000005 ' ��ɫ
                .Cell(flexcpBackColor, 1, .ColIndex("��ָ���ۼ�"), .Rows - 1, .ColIndex("��ָ���ۼ�")) = &H80000005 ' ��ɫ
            End If
        Next
    End With
End Sub
Private Sub cbsMain_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then
        Bottom = stbThis.Height
    End If
End Sub
Private Sub cbsMain_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Dim blnData As Boolean, i As Long

    Select Case Control.Id
    Case conMenu_Preview, conMenu_Print
        With vsPrice
            blnData = False
            For i = 0 To .Rows - 1
                If Val(.Cell(flexcpData, i, .ColIndex("Ʒ��"))) <> 0 Then
                    blnData = True
                    Exit For
                End If
            Next
        End With
        Control.Enabled = blnData
    Case conMenu_Save   '����
        
        With vsPrice
            blnData = False
            For i = 0 To .Rows - 1
                If Val(.Cell(flexcpData, i, .ColIndex("Ʒ��"))) <> 0 Then
                    blnData = True
                    Exit For
                End If
            Next
        End With
        Control.Enabled = blnData
        If mBillType = B_���� Then Control.Visible = False
    Case conMenu_Cancel  'ȡ��
    End Select
End Sub

Private Sub chkAppAllColumn_Click()
    If chkAppAllColumn.Value = 1 Then
        chk����.Enabled = False
        chk����.Value = 0
        chk��ʾ���в���.Value = 1
    Else
        chk����.Enabled = True
    End If
End Sub

Private Sub chk��ʾ���в���_Click()
    Dim i As Long
    With vsStoce
        For i = 1 To .Rows - 1
            .RowHidden(i) = IIf(chk��ʾ���в���.Value = 1, False, True)
        Next
    End With
    mlngPreRow = 0
    Call vsPrice_RowColChange
End Sub

Private Sub chk�Զ�����_Click()
    '����Ӧ���䶯���
    If (m���۷�ʽ = T_�ɱ��۵��� Or m���۷�ʽ = T_�ɱ����ۼ۵���) And chkӦ��.Value = 1 And chk�Զ�����.Value = 1 Then
        Call RefreshPayData
    End If
End Sub

Private Sub cmdPriver_Click()
    If Select��Ӧ��(Me, txtPriver, "") = False Then Exit Sub
End Sub

Private Sub cmdType_Click()
   Call Select���Ʒ���("")
   If txt����.Enabled Then txt����.SetFocus
End Sub

Private Sub cmd����_Click()
    If Get������Ŀ = False Then Exit Sub
End Sub

Private Sub DkPane_Action(ByVal Action As XtremeDockingPane.DockingPaneAction, ByVal Pane As XtremeDockingPane.IPane, Cancel As Boolean)
    If Action = PaneActionDocking Or Action = PaneActionExpanding Then
        If Pane.Id = ID_PANE_SEARCH And Pane.Hidden = False Then
            Cancel = True
        End If
    ElseIf Action = PaneActionPinning Or Action = PaneActionCollapsing Then
    Else
        Cancel = True
    End If
End Sub

Private Sub optӦ��_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
End Sub

Private Sub picBakDown_Resize()
    err = 0: On Error Resume Next
    Me.txt������.Left = picBakDown.ScaleWidth - Me.txt������.Width
    Me.lblValuer.Left = txt������.Left - lblValuer.Width - 50
    Me.txt˵��.Width = lblValuer.Left - txt˵��.Left - 300
    Me.dtpִ������.Left = picBakDown.ScaleWidth - Me.dtpִ������.Width
    Me.lblRunDate.Left = dtpִ������.Left - lblRunDate.Width - 50
End Sub
 
Private Sub picPrice_Resize()
    err = 0: On Error Resume Next
    With vsPrice
        .Left = picPrice.ScaleLeft
        .Top = picPrice.ScaleTop + chkAppAllColumn.Height
         picBakDown.Top = picPrice.ScaleHeight - picBakDown.Height
        picBakDown.Left = .Left
        picBakDown.Width = picPrice.ScaleWidth
        .Height = picBakDown.Top - .Top
        .Width = picPrice.ScaleWidth
    End With
End Sub

Private Sub picSeach_Resize()
    err = 0: On Error Resume Next
    With cmdType
        .Left = picSeach.ScaleWidth - .Width - 50
        txt����.Width = .Left - txt����.Left
    End With
    With fra
        .Width = picSeach.ScaleWidth - .Left - 50
    End With
    With fra������
        .Width = picSeach.ScaleWidth - .Left - 50
        
    End With
    With fraCost
        .Width = picSeach.ScaleWidth - .Left - 50
        cmdPriver.Left = .Width - cmdPriver.Width - 100
        txtPriver.Width = cmdPriver.Left - txtPriver.Left
    End With
    cmd����.Left = picSeach.ScaleWidth - cmd����.Width - 50
    
End Sub

Private Sub picStoceBack_Resize()
    err = 0: On Error Resume Next
    With tbPage
        .Left = picStoceBack.ScaleLeft
        .Width = picStoceBack.ScaleWidth
        .Top = picStoceBack.ScaleTop
        .Height = picStoceBack.ScaleHeight
        
    End With
End Sub


Private Sub txtPriver_Change()
    txtPriver.Tag = ""
End Sub

Private Sub txtPriver_GotFocus()
    OS.OpenIme False
    zlControl.TxtSelAll txtPriver
End Sub

Private Sub txtPriver_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If txtPriver.Tag <> "" Then OS.PressKey vbKeyTab: Exit Sub
    If txtPriver.Tag = "" And Trim(txtPriver.Text) = "" Then OS.PressKey vbKeyTab: Exit Sub
    If Select��Ӧ��(Me, txtPriver, Trim(txtPriver.Text)) = False Then Exit Sub
End Sub

Private Sub txt������_KeyPress(KeyAscii As Integer)
    If cbo������ʽ.ItemData(cbo������ʽ.ListIndex) = 3 Then
        Call zlControl.TxtCheckKeyPress(txt������, KeyAscii, m�����ʽ)
    Else
        Call zlControl.TxtCheckKeyPress(txt������, KeyAscii, m���ʽ)
    End If
End Sub

Private Sub txt����_Change()
    txt����.Tag = ""
End Sub
Private Sub txt����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    If Trim(txt����.Tag) <> "" Then
        OS.PressKey vbKeyTab
        Exit Sub
    End If
    If Trim(txt����.Text) = "" Then
        OS.PressKey vbKeyTab
        Exit Sub
    End If
    
    If Select���Ʒ���(Trim(txt����.Text)) = False Then
        Exit Sub
    End If
    OS.PressKey vbKeyTab
End Sub
Private Sub FullStoce�ּ�(ByVal lng����ID As Long, ByVal dbl�ּ� As Double)
    '-----------------------------------------------------------------------------------------------------------
    '����:�����ּ�,�����䶯���ּۼ�������
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2008-11-07 10:32:13
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, dbl������ As Double
    With vsStoce
        For lngRow = 1 To .Rows - 1
            If Val(.Cell(flexcpData, lngRow, .ColIndex("������Ϣ"))) = lng����ID Then
                .Cell(flexcpData, lngRow, .ColIndex("�ּ�")) = dbl�ּ�
                .TextMatrix(lngRow, .ColIndex("�ּ�")) = Format(dbl�ּ� * Val(.Cell(flexcpData, lngRow, .ColIndex("��λ"))), mFMT.FM_���ۼ�)
                '������=����*(�ּ�-ԭ��)
                dbl������ = (dbl�ּ� - Val(.Cell(flexcpData, lngRow, .ColIndex("ԭ��")))) * Val(.Cell(flexcpData, lngRow, .ColIndex("����")))
                .TextMatrix(lngRow, .ColIndex("������")) = Format(dbl������, mFMT.FM_���)
                .Cell(flexcpData, lngRow, .ColIndex("������")) = dbl������
                '��Ҫ���ݼӳ������¼�������ĳɱ���
                 Call AutoCalcStoce(lngRow, .ColIndex("�ּ�"))
            End If
        Next
    End With
End Sub

Private Sub FullStoce�ɱ���(ByVal lng����ID, ByVal dbl�ɱ��� As Double)
    '�ɱ���
    Dim lngRow As Long, dbl������ As Double
    With vsStoce
        For lngRow = 1 To .Rows - 1
            If Val(.Cell(flexcpData, lngRow, .ColIndex("������Ϣ"))) = lng����ID Then
                .Cell(flexcpData, lngRow, .ColIndex("�ֳɱ���")) = dbl�ɱ���
                .TextMatrix(lngRow, .ColIndex("�ֳɱ���")) = dbl�ɱ���
                 Call AutoCalcStoce(lngRow, .ColIndex("�ֳɱ���"))
            End If
        Next
    End With
End Sub

Private Sub txt�ӳ���_GotFocus()
    OS.OpenIme False
    zlControl.TxtSelAll txt�ӳ���
    
End Sub

Private Sub txt�ӳ���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then OS.PressKey vbKeyTab
    
End Sub

Private Sub txt�ӳ���_KeyPress(KeyAscii As Integer)
    Call zlControl.TxtCheckKeyPress(txt�ӳ���, KeyAscii, m���ʽ)
    
End Sub

Private Sub vsPay_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'    zl_VsGridRowChange vsPay, OldRow, NewRow, OldCol, NewCol
    
End Sub

Private Sub vsPay_GotFocus()
'    zl_VsGridGotFocus vsPay
    
End Sub

Private Sub vsPay_LostFocus()
'    zl_VsGridLOSTFOCUS vsPay
End Sub

Private Sub vsPrice_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '--------------------------------------------------------------------------------
    '������صĸ�ʽ
    '���˺�:2007/09/17
    '--------------------------------------------------------------------------------
    Dim lngRow As Long, dbl�ּ� As Double
    
    With vsPrice
        Select Case Col
        Case .ColIndex("ԭ��"), .ColIndex("��ָ���ۼ�")
            .TextMatrix(Row, Col) = Format(Val(.TextMatrix(Row, Col)), mFMT.FM_���ۼ�)
        Case .ColIndex("�ּ�")
            If chkAppAllColumn.Value = 0 Then
                'Ҫ�������С��λ
                dbl�ּ� = Val(.TextMatrix(Row, Col)) / Val(.TextMatrix(Row, .ColIndex("ϵ��")))
                .TextMatrix(Row, Col) = Format(Val(.TextMatrix(Row, Col)), mFMT.FM_���ۼ�)
                Call FullStoce�ּ�(Val(.Cell(flexcpData, Row, .ColIndex("Ʒ��"))), dbl�ּ�)
            Else
                Call AutoCalc���п��۸�
            End If
        Case .ColIndex("��ָ������")
            .TextMatrix(Row, Col) = Format(Val(.TextMatrix(Row, Col)), mFMT.FM_�ɱ���)
        Case .ColIndex("Ʒ��")
            .ColComboList(Col) = "..."
        Case .ColIndex("��������")
            .ColComboList(Col) = "..."
        Case .ColIndex("�ֳɱ���")
            '�ɱ��۵����Ͱ��ɱ����ۼ�һ�����ʱ�����޸ĳɱ���
            If chkAppAllColumn.Value = 1 Then
                Call AutoCalc���п��۸�
            Else
                Call FullStoce�ɱ���(Val(.Cell(flexcpData, Row, .ColIndex("Ʒ��"))), Val(.TextMatrix(.Row, .Col)))
            End If
        End Select
    End With
End Sub

Private Sub vsPrice_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'    zl_VsGridRowChange vsPrice, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vsPrice_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mBillType = B_���� Then Cancel = True: Exit Sub
    mlngPrice = Val(vsPrice.TextMatrix(vsPrice.Row, vsPrice.Col))
    With vsPrice
        Select Case Col
        Case .ColIndex("Ʒ��"), .ColIndex("�ֲɹ��޼�"), .ColIndex("��ָ���ۼ�") '
            .FocusRect = flexFocusSolid
            .HighLight = flexHighlightNever
            If Val(.Cell(flexcpData, Row, .ColIndex("Ʒ��"))) = 0 And (Col = .ColIndex("�ֲɹ��޼�") Or Col = .ColIndex("��ָ���ۼ�")) Then Cancel = True
        Case .ColIndex("�ּ�"), .ColIndex("��������")
            If m���۷�ʽ = T_�ɱ��۵��� Then
                .FocusRect = flexFocusHeavy
                Cancel = True
                Exit Sub
            Else
                .FocusRect = flexFocusSolid
                .HighLight = flexHighlightNever
                If Val(.Cell(flexcpData, Row, .ColIndex("Ʒ��"))) = 0 Then Cancel = True
            End If
        Case .ColIndex("�ֳɱ���")
            If m���۷�ʽ = T_�ɱ����ۼ۵��� Or m���۷�ʽ = T_�ɱ��۵��� Then
                .FocusRect = flexFocusSolid
                .HighLight = flexHighlightNever
                If Val(.Cell(flexcpData, Row, .ColIndex("Ʒ��"))) = 0 Then Cancel = True
            Else
                .FocusRect = flexFocusHeavy
                Cancel = True
            End If
        Case Else
            .FocusRect = flexFocusHeavy
            Cancel = True
        End Select
    End With
End Sub

Private Sub vsPrice_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    '--------------------------------------------------------------------------
    '����:��ťѡ��
    '����:
    '
    '--------------------------------------------------------------------------
    With vsPrice
        Select Case Col
        Case .ColIndex("Ʒ��")
            If SelectStuff("") = False Then Exit Sub
        Case .ColIndex("��������")
            If Select������Ŀ("") = False Then Exit Sub
        Case Else
        End Select
    End With
End Sub

Private Sub vsPrice_ChangeEdit()
    mblnModify = True
End Sub
Private Sub vsPrice_EnterCell()
    If mBillType = B_���� Then Exit Sub
    
    With vsPrice
        Select Case .Col
        Case .ColIndex("Ʒ��")
             .ColComboList(.Col) = "..."
        Case .ColIndex("��������")
            .ColComboList(.Col) = "..."
        End Select
    End With
End Sub

Private Sub vsPrice_GotFocus()
'    zl_VsGridGotFocus vsPrice
End Sub

Private Sub vsPrice_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long, lngRow As Long
    Dim i As Integer
    With vsPrice
        If (.Col = .ColIndex("Ʒ��") Or .Col = .ColIndex("��������")) And KeyCode <> vbKeyReturn Then
            vsPrice.ColComboList(.Col) = ""
        End If
        
        If KeyCode = vbKeyDelete Then
            If MsgBox("���Ƿ����Ҫɾ�����еĵ�����Ŀ��?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
           Call MoveStockData(Val(.Cell(flexcpData, .Row, .ColIndex("Ʒ��"))))
            If .Row = .Rows - 1 And .Row = 1 Then
                .Clear 1
                .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
                Call InitControl
            Else
                .RemoveItem .Row
                Call RefreshPayData
            End If
        End If
        
        For i = 1 To vsStoce.Rows - 1
            If Val(vsStoce.Cell(flexcpData, i, vsStoce.ColIndex("������Ϣ"))) = Val(vsPrice.Cell(flexcpData, vsPrice.Row, vsPrice.ColIndex("Ʒ��"))) Then
                vsStoce.RowHidden(i) = False
            Else
                vsStoce.RowHidden(i) = True
            End If
        Next
        
    End With
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With vsPrice
        If Val(.Cell(flexcpData, vsPrice.Row, .ColIndex("Ʒ��"))) = 0 Then
            OS.PressKey vbKeyTab
            Exit Sub
        End If
        Call zlVsMoveGridCell(vsPrice, , , mBillType <> B_����, lngRow)
    End With
End Sub

Private Sub vsPrice_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim intCol As Integer
    Dim strKey As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsPrice
        Select Case Col
        Case .ColIndex("Ʒ��")
        
            strKey = Trim(vsPrice.EditText)
            strKey = Replace(strKey, Chr(vbKeyReturn), "")
            strKey = Replace(strKey, Chr(10), "")
            If strKey = "" Then Exit Sub
            If SelectStuff(strKey) = False Then Exit Sub
            vsPrice.EditText = vsPrice.TextMatrix(Row, Col)
        Case .ColIndex("��������")
            strKey = Trim(vsPrice.EditText)
            strKey = Replace(strKey, Chr(vbKeyReturn), "")
            strKey = Replace(strKey, Chr(10), "")
            If strKey = "" Then Exit Sub
            If strKey = "" Then Exit Sub
            If Select������Ŀ(strKey) = False Then
                vsPrice.TextMatrix(Row, Col) = vsPrice.EditText
                vsPrice.Cell(flexcpData, Row, Col) = ""
                Exit Sub
            End If
            vsPrice.EditText = vsPrice.TextMatrix(Row, Col)
        Case Else
            Call zlVsMoveGridCell(vsStoce, , , False)
        End Select
    End With
End Sub

Private Sub vsPrice_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub vsPrice_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0: Exit Sub
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Exit Sub
    With vsPrice
        Select Case Col
        Case .ColIndex("Ʒ��")
            Call VsFlxGridCheckKeyPress(vsPrice, Row, Col, KeyAscii, m�ı�ʽ)
        Case .ColIndex("�ּ�")
            Call VsFlxGridCheckKeyPress(vsPrice, Row, Col, KeyAscii, m���ʽ)
        Case .ColIndex("��ָ������")
            Call VsFlxGridCheckKeyPress(vsPrice, Row, Col, KeyAscii, m���ʽ)
        Case .ColIndex("��ָ�����")
            Call VsFlxGridCheckKeyPress(vsPrice, Row, Col, KeyAscii, m���ʽ)
        Case .ColIndex("��������")
            Call VsFlxGridCheckKeyPress(vsPrice, Row, Col, KeyAscii, m�ı�ʽ)
        Case Else
        End Select
    End With
End Sub

Private Sub vsPrice_LostFocus()
'    zl_VsGridLOSTFOCUS vsPrice
End Sub

Private Sub vsPrice_RowColChange()
    '�ҵ�ָ������������
    Dim lng����ID As Long
    With vsPrice
'        .FocusRect = IIf(.Editable = flexEDKbdMouse, flexFocusHeavy, flexFocusSolid)
        If mlngPreRow = .Row Then Exit Sub
        mlngPreRow = .Row
        lng����ID = Val(.Cell(flexcpData, .Row, .ColIndex("Ʒ��")))
        If lng����ID = 0 Then Exit Sub
        Call Find����(lng����ID)
    End With
End Sub
Private Sub Find����(ByVal lng����ID As Long, Optional FindNext As Boolean = False)
    '-----------------------------------------------------------------------------------------------------------
    '����:����ָ���Ĳ���
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-12-08 15:18:19
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim BlnFind As Boolean
    BlnFind = False
    With vsStoce
        For i = 1 To .Rows - 1
            If Val(.Cell(flexcpData, i, .ColIndex("������Ϣ"))) <> lng����ID Then
                If chk��ʾ���в���.Value = 0 Then
                    .RowHidden(i) = True
                Else
                    .RowHidden(i) = False
                End If
            Else
                If chk��ʾ���в���.Value = 1 Then
                    .RowHidden(i) = False
                    .Row = i
                    .TopRow = .Row
                    Exit Sub
                Else
                    .RowHidden(i) = False
                    If Not BlnFind Then
                    .Row = i
                    BlnFind = True
                    End If
                End If
            End If
        Next
    End With
    
End Sub


Private Sub vsPrice_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKey As String
    Dim intCol As Integer
    Dim strTemp As String
    Dim intRow As Integer
    If mBillType = B_���� Then Cancel = True: Exit Sub
    
    strKey = Trim(vsPrice.EditText)
    strKey = Replace(strKey, Chr(vbKeyReturn), "")
    strKey = Replace(strKey, Chr(10), "")
    With vsPrice
        Select Case Col
        Case .ColIndex("Ʒ��")
        Case .ColIndex("��������") '
        Case .ColIndex("�ּ�")
            If strKey <> "" Then
                If zlCommFun.DblIsValid(strKey, 12, , False, , "�ּ�") = False Then Cancel = True: Exit Sub
                If Val(.Cell(flexcpData, .Row, .ColIndex("Ʒ��"))) = 0 Then
                    vsPrice.EditText = Format(Val(strKey), mFMT.FM_���ۼ�)
                    Exit Sub
                End If
                If Val(strKey) > Val(.TextMatrix(.Row, .ColIndex("��ָ���ۼ�"))) And Val(.TextMatrix(.Row, .ColIndex("��ָ���ۼ�"))) <> 0 Then
                    MsgBox "�ּ۲��ܴ���ָ�����ۼۣ���" & Format(Val(.TextMatrix(.Row, .ColIndex("��ָ���ۼ�"))), mFMT.FM_���ۼ�) & "��", vbQuestion + vbDefaultButton1, gstrSysName
                    Cancel = True
                    Exit Sub
                End If
                vsPrice.EditText = Format(Val(strKey), mFMT.FM_���ۼ�)
                mblnModify = True
            End If
            If chkAppAllColumn.Value = 1 And mlngPrice <> vsPrice.EditText Then
                For intRow = 1 To .Rows - 1
                    .TextMatrix(intRow, .ColIndex("�ּ�")) = vsPrice.EditText
                Next
            End If
        Case .ColIndex("��ָ������")
            If strKey <> "" Then
                If zlCommFun.DblIsValid(strKey, 12, , False, , "��ָ������") = False Then Cancel = True: Exit Sub
                vsPrice.EditText = Format(Val(strKey), mFMT.FM_�ɱ���)
            End If
        Case .ColIndex("��ָ���ۼ�")
            If strKey <> "" Then
                If zlCommFun.DblIsValid(strKey, 12, , False, , "��ָ���ۼ�") = False Then Cancel = True: Exit Sub
                vsPrice.EditText = Format(Val(strKey), mFMT.FM_���ۼ�)
                
                If chkAppAllColumn.Value = 1 And mlngPrice <> vsPrice.EditText Then
                    For intRow = 1 To .Rows - 1
                        .TextMatrix(intRow, .ColIndex("��ָ���ۼ�")) = vsPrice.EditText
                    Next
                End If
            End If
        Case .ColIndex("�ֲɹ��޼�")
            If strKey <> "" Then
                vsPrice.EditText = Format(Val(strKey), mFMT.FM_�ɱ���)
                If chkAppAllColumn.Value = 1 And mlngPrice <> vsPrice.EditText Then
                    For intRow = 1 To .Rows - 1
                        .TextMatrix(intRow, .ColIndex("�ֲɹ��޼�")) = vsPrice.EditText
                    Next
                End If
            End If
        Case .ColIndex("�ֳɱ���")
            If strKey <> "" Then
                vsPrice.EditText = Format(Val(strKey), mFMT.FM_�ɱ���)
                If chkAppAllColumn.Value = 1 And mlngPrice <> vsPrice.EditText Then
                    For intRow = 1 To .Rows - 1
                        .TextMatrix(intRow, .ColIndex("�ֳɱ���")) = vsPrice.EditText
                    Next
                End If
            End If
        End Select
    End With
End Sub
Private Function Select���Ʒ���(ByVal strSeach As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:ѡ��ָ������������
    '����:strKey-��ѡ�������
    '����:ѡ��ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2007/09/17
    '-----------------------------------------------------------------------------------------------------------
    Dim blnCancel As Boolean, strKey As String, strTittle As String, lngH As Long
    Dim objCtl As Object: Dim vRect As RECT
    Dim rsTemp  As ADODB.Recordset
    'zlDatabase.ShowSelect
    '���ܣ��๦��ѡ����
    '������
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    
    
    Set objCtl = txt����
    vRect = zlControl.GetControlRect(txt����.hwnd)
    lngH = txt����.Height
    strKey = GetMatchingSting(strSeach)
      
    strTittle = "�������Ϸ���ѡ��"
    If strSeach = "" Then
'        gstrSQL = "" & _
'                "   Select ID,�ϼ�ID, ����,����,���� From ���Ʒ���Ŀ¼ a " & _
'                "   where  ����=7 start with �ϼ�id is null connect by prior id=�ϼ�id"
        
        gstrSQL = "Select ID, �ϼ�id, ����, ����, ���" & _
                " From (Select ID, �ϼ�id, ����, ����, '����' ���" & _
                       " From ���Ʒ���Ŀ¼" & _
                       " Where ���� = 7" & _
                       " Start With �ϼ�id Is Null" & _
                       " Connect By Prior ID = �ϼ�id" & _
                       " Union All" & _
                       " Select a.Id, a.����id As �ϼ�id, a.����, a.����, 'Ʒ��' ���" & _
                       " From ������ĿĿ¼ A," & _
                       "     (Select ID From ���Ʒ���Ŀ¼ Where ���� = 7 Start With �ϼ�id Is Null Connect By Prior ID = �ϼ�id) B" & _
                       " Where a.��� = '4' And a.����id = b.Id)" & _
                " Start With �ϼ�id Is Null" & _
                " Connect By Prior ID = �ϼ�id"
        
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 1, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False)
    Else
        If optӦ��(2).Value = False Then
            gstrSQL = "" & _
                    "   Select ID,�ϼ�ID, ����,����,����,'����' ��� From ���Ʒ���Ŀ¼ a " & _
                    "   Where (���� like [1] or  ����  like [1] or  ����  like  [1]) and ����=7  " & _
                    "   order by ����"
        Else
            gstrSQL = "select a.����id,a.id,a.����,a.����,'Ʒ��' ��� from ������ĿĿ¼ a,������Ŀ���� b " & _
            " where a.��� ='4' and a.id=b.������Ŀid and (a.���� like [1] or a.���� like [1] OR b.���� like [1])"
        End If
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, strTittle, False, "", "", False, False, True, vRect.Left - 15, vRect.Top, lngH, blnCancel, False, False, strKey)
    End If
    
    If blnCancel = True Then
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    If rsTemp Is Nothing Then
        ShowMsgBox "û�����������Ĳ��Ϸ���,����!"
        If objCtl.Enabled Then objCtl.SetFocus
        Exit Function
    End If
    
    objCtl.Text = zlStr.Nvl(rsTemp!����) & "-" & zlStr.Nvl(rsTemp!����)
    objCtl.Tag = zlStr.Nvl(rsTemp!Id)
    If InStr(1, rsTemp!���, "����") > 0 Then '����
        optӦ��(0).Enabled = True
        optӦ��(1).Enabled = True
        optӦ��(2).Enabled = False
        optӦ��(2).Value = False
    Else 'Ʒ��
        optӦ��(0).Enabled = False
        optӦ��(1).Enabled = False
        optӦ��(2).Enabled = True
        optӦ��(2).Value = True
    End If
    
    Select���Ʒ��� = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function SelectStuff(ByVal strKey As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:ѡ��ָ������������
    '����:strKey-��ѡ�������
    '����:ѡ��ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2007/09/17
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim blnCancel As Boolean, i As Long
    Dim vRect As RECT, sngX As Single, sngY As Single
    Dim intϵ�� As Integer
    'zlDatabase.ShowSelect
    '���ܣ��๦��ѡ����
    '������
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    err = 0: On Error GoTo ErrHand:
    Call CalcPosition(sngX, sngY, vsPrice)
              
    If strKey <> "" Then
        strKey = GetMatchingSting(strKey)
        gstrSQL = "" & _
            "   Select distinct I.ID,I.����,I.����,I.���,I.����,I.���㵥λ,P.����ϵ��,P.��װ��λ," & _
            "         decode(I.�Ƿ���,1,'ʱ��','����') ����," & _
            "         P.�ɱ��� as �ɱ���ID,P.ָ�������� as ָ��������ID,P.ָ�����ۼ� as ָ�����ۼ�ID," & _
            "         to_char(p.�ɱ���," & mOraFMT.FM_�ɱ��� & ") as �ɱ���," & _
            "         to_char(p.ָ��������," & mOraFMT.FM_�ɱ��� & ") ָ��������," & _
            "         to_char(p.ָ�����ۼ�," & mOraFMT.FM_���ۼ� & ") ָ�����ۼ�," & _
            "          P.��������" & _
            "   From �շ���ĿĿ¼ I,�շ���Ŀ���� N,�������� P" & _
            "   Where I.ID=N.�շ�ϸĿID and I.���='4' And I.ID=P.����ID " & _
            "       and (I.���� like [1] or N.���� Like [1] or N.���� Like [1])" & _
            "       and (I.����ʱ�� Is Null Or I.����ʱ��=To_Date('3000-01-01','yyyy-MM-dd'))"
     Else
        gstrSQL = "" & _
            "   Select  I.ID,I.����,I.����,I.���,I.����,I.���㵥λ,P.����ϵ��,P.��װ��λ, " & _
            "           decode(I.�Ƿ���,1,'ʱ��','����') ����," & _
            "           P.�ɱ��� as �ɱ���ID,P.ָ�������� as ָ��������ID,P.ָ�����ۼ� as ָ�����ۼ�ID," & _
            "           to_char(p.�ɱ���," & mOraFMT.FM_�ɱ��� & ") as �ɱ���," & _
            "           to_char(p.ָ��������," & mOraFMT.FM_�ɱ��� & ") ָ��������," & _
            "           to_char(p.ָ�����ۼ�," & mOraFMT.FM_���ۼ� & ") ָ�����ۼ�," & _
            "           P.��������" & _
            "   From �շ���ĿĿ¼ I,�������� P" & _
            "   Where I.���='4' And I.ID=P.����ID" & _
            "           and (I.����ʱ�� Is Null Or I.����ʱ��=To_Date('3000-01-01','yyyy-MM-dd'))"
            
    End If
    
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "��������ѡ��", False, "", "", False, False, True, sngX, sngY - vsPrice.CellHeight, vsPrice.CellHeight, blnCancel, False, False, strKey)
    If blnCancel = True Then Exit Function
    
    If rsTemp Is Nothing Then
        ShowMsgBox "������ָ������������,����!"
        Exit Function
    End If
    
    With Me.vsPrice
        '����Ƿ�ѡ����ͬһ��Ʒ�ֵ���������
        For i = 1 To .Rows - 1
            If Val(.Cell(flexcpData, i, .ColIndex("Ʒ��"))) <> 0 Then
                If Val(.Cell(flexcpData, i, .ColIndex("Ʒ��"))) = Val(zlStr.Nvl(rsTemp!Id)) And i <> .Row Then
                    ShowMsgBox "�����������Ѿ����ڣ����ܽ��е��ۣ�"
                    Exit Function
                End If
            End If
        Next
        
        '����Ƿ�ı���ԭ���Ѿ����ڵ���������
        If Val(.Cell(flexcpData, .Row, .ColIndex("Ʒ��"))) <> Val(zlStr.Nvl(rsTemp!Id)) And Val(.Cell(flexcpData, .Row, .ColIndex("Ʒ��"))) <> 0 Then
            '��Ҫ�Ƴ����������ϵĿⷿ�䶯���������ܸ���
             Call MoveStockData(Val(.Cell(flexcpData, .Row, .ColIndex("Ʒ��"))))
        End If
        
        .Redraw = flexRDNone
        .TextMatrix(.Row, .ColIndex("Ʒ��")) = "[" & zlStr.Nvl(rsTemp!����) & "]" & zlStr.Nvl(rsTemp!����)
        .Cell(flexcpData, .Row, .ColIndex("Ʒ��")) = zlStr.Nvl(rsTemp!Id)
        .TextMatrix(.Row, .ColIndex("���")) = zlStr.Nvl(rsTemp!���)
        .TextMatrix(.Row, .ColIndex("����")) = zlStr.Nvl(rsTemp!����)
        .TextMatrix(.Row, .ColIndex("��λ")) = IIf(mintUnit = 0, zlStr.Nvl(rsTemp!���㵥λ), zlStr.Nvl(rsTemp!��װ��λ))
        .TextMatrix(.Row, .ColIndex("����")) = zlStr.Nvl(rsTemp!����)
        .Cell(flexcpData, .Row, .ColIndex("����")) = zlStr.Nvl(rsTemp!��������)
        
        intϵ�� = IIf(mintUnit = 0, 1, zlStr.Nvl(rsTemp!����ϵ��))
        .TextMatrix(.Row, .ColIndex("ϵ��")) = intϵ��
        
        
        .TextMatrix(.Row, .ColIndex("�ֳɱ���")) = Format(Val(zlStr.Nvl(rsTemp!�ɱ���ID)) * intϵ��, mFMT.FM_�ɱ���)
        .Cell(flexcpData, .Row, .ColIndex("�ֳɱ���")) = Val(zlStr.Nvl(rsTemp!�ɱ���ID))
        
        .TextMatrix(.Row, .ColIndex("ԭ�ɹ��޼�")) = Format(Val(zlStr.Nvl(rsTemp!ָ��������ID)) * intϵ��, mFMT.FM_�ɱ���)
        .TextMatrix(.Row, .ColIndex("�ֲɹ��޼�")) = .TextMatrix(.Row, .ColIndex("ԭ�ɹ��޼�"))
        .Cell(flexcpData, .Row, .ColIndex("ԭ�ɹ��޼�")) = Val(zlStr.Nvl(rsTemp!ָ��������ID))
        .Cell(flexcpData, .Row, .ColIndex("�ֲɹ��޼�")) = Val(zlStr.Nvl(rsTemp!ָ��������ID))
        
        
        .TextMatrix(.Row, .ColIndex("ָ�����ۼ�")) = Format(Val(zlStr.Nvl(rsTemp!ָ�����ۼ�ID)) * intϵ��, mFMT.FM_���ۼ�)
        .TextMatrix(.Row, .ColIndex("ԭָ���ۼ�")) = .TextMatrix(.Row, .ColIndex("ָ�����ۼ�"))
        .TextMatrix(.Row, .ColIndex("��ָ���ۼ�")) = .TextMatrix(.Row, .ColIndex("ָ�����ۼ�"))
        
        .Cell(flexcpData, .Row, .ColIndex("ָ�����ۼ�")) = Val(zlStr.Nvl(rsTemp!ָ�����ۼ�ID))
        .Cell(flexcpData, .Row, .ColIndex("ԭָ���ۼ�")) = Val(zlStr.Nvl(rsTemp!ָ�����ۼ�ID))
        .Cell(flexcpData, .Row, .ColIndex("��ָ���ۼ�")) = Val(zlStr.Nvl(rsTemp!ָ�����ۼ�ID))
        Call zlGetPrice(Val(zlStr.Nvl(rsTemp!Id)), IIf(.TextMatrix(.Row, .ColIndex("����")) = "ʱ��", True, False))
        Call LoadStockData(Val(zlStr.Nvl(rsTemp!Id)), Val(.Cell(flexcpData, .Row, .ColIndex("ԭ��"))), Val(.Cell(flexcpData, .Row, .ColIndex("�ּ�"))))
        
        .Col = .ColIndex("�ּ�")
        .Redraw = flexRDBuffered
        zlControl.ControlSetFocus vsPrice, True
        mlngPreRow = 0:
        Call vsPrice_RowColChange
    End With
    SelectStuff = True
    Exit Function
ErrHand:
    vsPrice.Redraw = flexRDBuffered
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub zlGetPrice(lng����ID As Long, blnʵ�� As Boolean, Optional bln���� As Boolean = False, Optional lngRow As Long = -1)
    '----------------------------------------------------
    '���ܣ���дָ������id�Ķ�Ӧ�۸���Ϣ
    '��Σ�lng����ID-����ID
    '      blnʵ��:�Ƿ�ʱ������
    '      bln����-False�������������ּ�,true-�������������ּ�
    '����:���˺�
    '����:2007/09/17
    '----------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim bytType As Byte
    Dim dbl���� As Double
    
    On Error GoTo ErrHandle
    If bln���� Then
        bytType = cbo������ʽ.ItemData(cbo������ʽ.ListIndex)
    End If
    
    If blnʵ�� Then
        Me.Chk����.Enabled = True
        '��ʾʱ�����ĵ��ۣ�ȡ�����/���������Ϊ��۸�
        gstrSQL = "" & _
            "   Select  P.id,Decode(Nvl(K.�������,0),0,P.�ּ�,K.�����/Nvl(K.�������,1)) �ּ�," & _
            "           P.ִ������,P.������Ŀid,I.���� as ��������, " & IIf(mintUnit = 0, "1", " Nvl(M.����ϵ��,1)") & " as  ϵ��,nvl(m.�ɱ���,0) as �ɱ���,m.��������" & _
            "   From �շѼ�Ŀ P,������Ŀ I,�������� M," & _
            "       (   Select Sum(ʵ�ʽ��) �����,Sum(ʵ������) �������" & _
            "           From ҩƷ��� " & _
            "           Where  ����=1 and ҩƷID=[1] " & _
            "        ) K" & _
            " where p.�շ�ϸĿid=M.����id and P.������Ŀid=I.id and P.�շ�ϸĿid=[1] " & _
            "       and (P.��ֹ���� is null or P.��ֹ����=to_date('3000-01-01','YYYY-MM-DD'))" & _
            GetPriceClassString("P")
    Else
        '��ʱ�����ĵ��ۣ�ȡ��۸��¼�еļ۸�
        gstrSQL = "" & _
            "   Select P.id,P.�ּ�,P.ִ������,P.������Ŀid,I.���� as ��������," & IIf(mintUnit = 0, "1", " Nvl(M.����ϵ��,1)") & " as  ϵ��,nvl(m.�ɱ���,0) as �ɱ���,m.��������" & _
            "   From �շѼ�Ŀ P,������Ŀ I,�������� M" & _
            "   Where p.�շ�ϸĿid=M.����id and P.������Ŀid=I.id and P.�շ�ϸĿid=[1]  " & _
            "           and (P.��ֹ���� is null or P.��ֹ����=to_date('3000-01-01','YYYY-MM-DD'))" & _
            GetPriceClassString("P")
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID)
    With vsPrice
        If lngRow < 0 Then lngRow = .Row
        If rsTemp.RecordCount > 0 Then
            .RowData(lngRow) = Val(zlStr.Nvl(rsTemp!Id))
            .Cell(flexcpData, lngRow, .ColIndex("����")) = zlStr.Nvl(rsTemp!��������)
            
            .TextMatrix(lngRow, .ColIndex("�ϴ�����")) = Format(rsTemp!ִ������, "YYYY-MM-DD HH:MM:SS")
            .TextMatrix(lngRow, .ColIndex("ԭ��")) = Format(Val(zlStr.Nvl(rsTemp!�ּ�)) * Val(zlStr.Nvl(rsTemp!ϵ��)), mFMT.FM_���ۼ�)
            .Cell(flexcpData, lngRow, .ColIndex("ԭ��")) = Val(zlStr.Nvl(rsTemp!�ּ�))
            If bln���� = False Then
                .TextMatrix(lngRow, .ColIndex("�ּ�")) = Format(Val(zlStr.Nvl(rsTemp!�ּ�)) * Val(zlStr.Nvl(rsTemp!ϵ��)), mFMT.FM_���ۼ�)
                .Cell(flexcpData, lngRow, .ColIndex("�ּ�")) = Val(zlStr.Nvl(rsTemp!�ּ�))
            Else
                If Val(txt������.Text) = 0 Then
                    .TextMatrix(lngRow, .ColIndex("�ּ�")) = Format(Val(zlStr.Nvl(rsTemp!�ּ�)) * Val(zlStr.Nvl(rsTemp!ϵ��)), mFMT.FM_���ۼ�)
                    .Cell(flexcpData, lngRow, .ColIndex("�ּ�")) = Val(zlStr.Nvl(rsTemp!�ּ�)) * Val(zlStr.Nvl(rsTemp!ϵ��))
                Else
                    Select Case bytType
                    Case 1      '���ݳɱ��ۼӳ�
                        dbl���� = 1 + Val(txt������.Text) / 100
                        .TextMatrix(lngRow, .ColIndex("�ּ�")) = Format(Val(zlStr.Nvl(rsTemp!�ɱ���)) * dbl���� * Val(zlStr.Nvl(rsTemp!ϵ��)), mFMT.FM_���ۼ�)
                        .Cell(flexcpData, lngRow, .ColIndex("�ּ�")) = Val(zlStr.Nvl(rsTemp!�ɱ���)) * dbl����
                    Case 2      '�������ۼ۰�����
                        dbl���� = 1 + Val(txt������.Text) / 100
                        .TextMatrix(lngRow, .ColIndex("�ּ�")) = Format(Val(zlStr.Nvl(rsTemp!�ּ�)) * dbl���� * Val(zlStr.Nvl(rsTemp!ϵ��)), mFMT.FM_���ۼ�)
                        .Cell(flexcpData, lngRow, .ColIndex("�ּ�")) = Val(zlStr.Nvl(rsTemp!�ּ�)) * dbl����
                    Case 3      '�������ۼ۰��̶����Ӽ�
                        dbl���� = Val(txt������.Text)
                        .TextMatrix(lngRow, .ColIndex("�ּ�")) = Format((Val(zlStr.Nvl(rsTemp!�ּ�)) * Val(zlStr.Nvl(rsTemp!ϵ��))) + dbl����, mFMT.FM_���ۼ�)
                        .Cell(flexcpData, lngRow, .ColIndex("�ּ�")) = Val(zlStr.Nvl(rsTemp!�ּ�)) + dbl����
                    End Select
                End If
                If Val(.TextMatrix(lngRow, .ColIndex("�ּ�"))) > Val(.TextMatrix(lngRow, .ColIndex("ָ�����ۼ�"))) And Val(.TextMatrix(lngRow, .ColIndex("ָ�����ۼ�"))) <> 0 Then
                    .TextMatrix(lngRow, .ColIndex("�ּ�")) = Format(Val(.TextMatrix(lngRow, .ColIndex("ָ�����ۼ�"))), mFMT.FM_���ۼ�)
                    .Cell(flexcpData, lngRow, .ColIndex("�ּ�")) = Val(.Cell(flexcpData, lngRow, .ColIndex("ָ�����ۼ�")))
                End If
            End If
            .TextMatrix(lngRow, .ColIndex("�ֳɱ���")) = Format(rsTemp!�ɱ��� * Val(zlStr.Nvl(rsTemp!ϵ��)), mFMT.FM_�ɱ���)
            .Cell(flexcpData, lngRow, .ColIndex("�ֳɱ���")) = Val(zlStr.Nvl(rsTemp!�ɱ���))
            
            .TextMatrix(lngRow, .ColIndex("ԭ����id")) = Val(zlStr.Nvl(rsTemp!������Ŀid))
            .TextMatrix(lngRow, .ColIndex("��������")) = zlStr.Nvl(rsTemp!��������)
            .Cell(flexcpData, lngRow, .ColIndex("��������")) = Val(zlStr.Nvl(rsTemp!������Ŀid))
        Else
            .RowData(lngRow) = -1
            .TextMatrix(lngRow, .ColIndex("�ϴ�����")) = ""
            .TextMatrix(lngRow, .ColIndex("ԭ��")) = Format(0, mFMT.FM_���ۼ�)
            .TextMatrix(lngRow, .ColIndex("�ּ�")) = Format(0, mFMT.FM_���ۼ�)
            .TextMatrix(lngRow, .ColIndex("�ֳɱ���")) = Format(0, mFMT.FM_�ɱ���)
            .Cell(flexcpData, lngRow, .ColIndex("ԭ��")) = 0
            .Cell(flexcpData, lngRow, .ColIndex("�ּ�")) = 0
            .Cell(flexcpData, lngRow, .ColIndex("�ֳɱ���")) = 0
            If bln���� Then
                '��һ������:
                If Val(txt������.Text) = 0 Then
                    '���û���õ�����,��Ϊ0
                    .TextMatrix(lngRow, .ColIndex("�ּ�")) = Format(0, mFMT.FM_���ۼ�)
                    .Cell(flexcpData, lngRow, .ColIndex("�ּ�")) = 0
                Else
                    Select Case bytType
                    Case 1      '���ݳɱ��ۼӳ�
                        dbl���� = 1 + Val(txt������.Text) / 100
                        .TextMatrix(lngRow, .ColIndex("�ּ�")) = Format(0, mFMT.FM_���ۼ�)
                        .Cell(flexcpData, lngRow, .ColIndex("�ּ�")) = 0 * dbl����
                    Case 2      '�������ۼ۰�����
                        dbl���� = 1 + Val(txt������.Text) / 100
                        .TextMatrix(lngRow, .ColIndex("�ּ�")) = Format(0 * dbl���� * Val(.TextMatrix(lngRow, .ColIndex("ϵ��"))), mFMT.FM_���ۼ�)
                        .Cell(flexcpData, lngRow, .ColIndex("�ּ�")) = 0 * dbl����
                    Case 3      '�������ۼ۰��̶����Ӽ�
                        dbl���� = Val(txt������.Text)
                        .TextMatrix(lngRow, .ColIndex("�ּ�")) = Format(0 + dbl���� * Val(.TextMatrix(lngRow, .ColIndex("ϵ��"))), mFMT.FM_���ۼ�)
                        .Cell(flexcpData, lngRow, .ColIndex("�ּ�")) = 0 + dbl����
                    End Select
                End If
                If Val(.TextMatrix(lngRow, .ColIndex("�ּ�"))) > Val(.TextMatrix(lngRow, .ColIndex("ָ�����ۼ�"))) And Val(.TextMatrix(lngRow, .ColIndex("ָ�����ۼ�"))) <> 0 Then
                    .TextMatrix(lngRow, .ColIndex("�ּ�")) = Format(Val(.TextMatrix(lngRow, .ColIndex("ָ�����ۼ�"))), mFMT.FM_���ۼ�)
                    .Cell(flexcpData, lngRow, .ColIndex("�ּ�")) = Val(.Cell(flexcpData, lngRow, .ColIndex("ָ�����ۼ�")))
                End If
            End If
            If lngRow > 1 Then
                .TextMatrix(lngRow, .ColIndex("ԭ����id")) = .TextMatrix(lngRow - 1, .ColIndex("ԭ����id"))
                .TextMatrix(lngRow, .ColIndex("��������")) = .TextMatrix(lngRow - 1, .ColIndex("��������"))
                .Cell(flexcpData, lngRow, .ColIndex("��������")) = .Cell(flexcpData, lngRow - 1, .ColIndex("��������"))
            End If
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Select������Ŀ(ByVal strKey As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:ѡ��ָ����������Ŀ��Ϣ
    '����:strKey-��ѡ�������
    '����:ѡ��ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2007/09/17
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset
    Dim blnCancel As Boolean
    Dim vRect As RECT
    Dim sngX As Single, sngY As Single
    
    'zlDatabase.ShowSelect
    '���ܣ��๦��ѡ����
    '������
    '     frmParent=��ʾ�ĸ�����
    '     strSQL=������Դ,��ͬ����ѡ������SQL�е��ֶ��в�ͬҪ��
    '     bytStyle=ѡ�������
    '       Ϊ0ʱ:�б���:ID,��
    '       Ϊ1ʱ:���η��:ID,�ϼ�ID,����,����(���blnĩ��������Ҫĩ���ֶ�)
    '       Ϊ2ʱ:˫����:ID,�ϼ�ID,����,����,ĩ������ListViewֻ��ʾĩ��=1����Ŀ
    '     strTitle=ѡ������������,Ҳ���ڸ��Ի�����
    '     blnĩ��=������ѡ����(bytStyle=1)ʱ,�Ƿ�ֻ��ѡ��ĩ��Ϊ1����Ŀ
    '     strSeek=��bytStyle<>2ʱ��Ч,ȱʡ��λ����Ŀ��
    '             bytStyle=0ʱ,��ID���ϼ�ID֮��ĵ�һ���ֶ�Ϊ׼��
    '             bytStyle=1ʱ,�����Ǳ��������
    '     strNote=ѡ������˵������
    '     blnShowSub=��ѡ��һ���Ǹ����ʱ,�Ƿ���ʾ�����¼������е���Ŀ(��Ŀ��ʱ����)
    '     blnShowRoot=��ѡ������ʱ,�Ƿ���ʾ������Ŀ(��Ŀ��ʱ����)
    '     blnNoneWin,X,Y,txtH=����ɷǴ�����,X,Y,txtH��ʾ���ý�������������(�������Ļ)�͸߶�
    '     Cancel=���ز���,��ʾ�Ƿ�ȡ��,��Ҫ����blnNoneWin=Trueʱ
    '     blnMultiOne=��bytStyle=0ʱ,�Ƿ񽫶Զ�����ͬ��¼����һ���ж�
    '     blnSearch=�Ƿ���ʾ�к�,�����������кŶ�λ
    '���أ�ȡ��=Nothing,ѡ��=SQLԴ�ĵ��м�¼��
    '˵����
    '     1.ID���ϼ�ID����Ϊ�ַ�������
    '     2.ĩ�����ֶβ�Ҫ����ֵ
    'Ӧ�ã������ڸ������������������Ǻܴ��ѡ����,����ƥ���б�ȡ�
    err = 0: On Error GoTo ErrHand:
    Call CalcPosition(sngX, sngY, vsPrice)
    
    If strKey <> "" Then
        strKey = GetMatchingSting(strKey)
        gstrSQL = "" & _
            "   Select id,����,����,����,�վݷ�Ŀ,������Ŀ" & _
            "   From ������Ŀ" & _
            "   Where (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','yyyy-MM-dd')) and ĩ��=1 " & _
            "         and (���� like [1] or ���� Like [1] or ���� Like [1])"
     Else
        gstrSQL = "" & _
            "   Select id,����,����,����,�վݷ�Ŀ,������Ŀ" & _
            "   From ������Ŀ" & _
            "   Where (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','yyyy-MM-dd')) and ĩ��=1 "
    End If
    
    Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSQL, 0, "������Ŀѡ����", False, "", "", False, False, True, sngX, sngY - vsPrice.CellHeight, vsPrice.CellHeight, blnCancel, False, False, strKey)
    If blnCancel = True Then Exit Function
    
    If rsTemp Is Nothing Then
        ShowMsgBox "������ָ����������Ŀ,����!"
        Exit Function
    End If
    
    With Me.vsPrice
        .Redraw = flexRDNone
        .TextMatrix(.Row, .ColIndex("��������")) = zlStr.Nvl(rsTemp!����)
        .Cell(flexcpData, .Row, .ColIndex("��������")) = zlStr.Nvl(rsTemp!Id)
        .Redraw = flexRDBuffered
    End With
    
    Select������Ŀ = True
    Exit Function
ErrHand:
    vsPrice.Redraw = flexRDBuffered
    If ErrCenter = 1 Then
        Resume
    End If
End Function
  
 

Private Sub chk����ִ��_Click()
    Dim rsTemp As New ADODB.Recordset
    Dim i As Long, lng����ID As Long
    
    Dim mlngStuffIdThis As Long, IntCheck As Integer
    
    On Error GoTo ErrHandle
    If chk����ִ��.Value = 1 Then
        
        'ѭ���ж����в���
        With vsPrice
            For i = 1 To .Rows - 1
                lng����ID = Val(.Cell(flexcpData, i, .ColIndex("Ʒ��")))
                gstrSQL = "Select count(*) as δִ�� From �շѼ�Ŀ where �䶯ԭ��=0 and  �շ�ϸĿid=[1]" & _
                        GetPriceClassString("")
                
                Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID)
                If Not rsTemp.EOF Then
                    If Val(zlStr.Nvl(rsTemp!δִ��)) <> 0 Then
                        MsgBox "��������" & .TextMatrix(i, .ColIndex("Ʒ��")) & "����δִ�м۸񣬲�������Ϊ����ִ�У�", vbInformation, gstrSysName
                        chk����ִ��.Value = 0
                        Exit Sub
                    End If
                End If
            Next
        End With
    End If
    If Me.chk����ִ��.Value Then
        Me.dtpִ������.Enabled = False
    Else
        Me.dtpִ������.Enabled = True
    End If
    err = 0: On Error Resume Next
    Me.vsPrice.SetFocus
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

 
 
Private Function ISValied() As Boolean
    '-------------------------------------------------------------------------------------------
    '����:�������ĺϷ���
    '����:
    '����:���ݺϷ�,����true,���򷵻�False
    '����:���˺�
    '����:2007/09/15
    '-------------------------------------------------------------------------------------------
    '����ִ�м۸��Ƿ���ȷ
    '�Լ�������Ŀ��ͬ��������ּ��Ƿ���ԭ����ͬ
    
    Dim i As Long, blnZero As Boolean, lng����ID As Long
    Dim strOldID As String, strNewID As String, strTemp As String
    Dim blnHaving As Boolean
    
    ISValied = False
    
    strNewID = "": strOldID = ""
    With vsPrice
        blnZero = False
        For i = 1 To .Rows - 1
        
            lng����ID = Val(.Cell(flexcpData, i, .ColIndex("Ʒ��")))
            If lng����ID <> 0 Then
                blnHaving = True
                If Not IsNumeric(Trim(.TextMatrix(i, .ColIndex("�ּ�")))) Then
                    MsgBox "��" & i & "�е����������ּ��к��зǷ��ַ���", vbInformation, gstrSysName
                    Exit Function
                End If
                                
                If m���۷�ʽ <> T_�ɱ��۵��� Then
                    '���˺�:��Ҫ�ǽ������Ϊ������,���磺����.����ѵ�
                    '����:9569 2006-11-20
                    If Val(.TextMatrix(i, .ColIndex("�ּ�"))) = 0 And blnZero = False Then
                        If MsgBox("��" & i & "�е����������ּ�Ϊ����,�Ƿ����?", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbYes Then
                            blnZero = True
                        Else
                            Exit Function
                        End If
                    End If
                End If
                
                If Val(.TextMatrix(i, .ColIndex("ԭ����ID"))) = Val(.Cell(flexcpData, i, .ColIndex("��������"))) And _
                   Val(.TextMatrix(i, .ColIndex("�ּ�"))) = Val(.TextMatrix(i, .ColIndex("ԭ��"))) Then
                   '��Ҫ�����صĵ�����Ϣ
                   If m���۷�ʽ <> T_�ɱ��۵��� Then
                        '�϶���Ҫ���гɱ��۵���
                        'If .TextMatrix(i, .ColIndex("����")) = "ʱ��" And .Cell(flexcpData, i, .ColIndex("����")) <> "1" Then
                        '    '�Ƿ�Χ���Ǹ����������ϵ�ʵ������
                        'Else
                            MsgBox "��" & i & "�е����������ּ���ԭ����ͬ������ִ�е��ۣ�", vbInformation, gstrSysName
                            Exit Function
                        'End If
                   End If
                End If
                
                If m���۷�ʽ <> T_�ɱ��۵��� Then
                    If .TextMatrix(i, .ColIndex("����")) = "ʱ��" And Me.chk����ִ��.Value <> 1 Then
                        MsgBox "��" & i & "��Ϊʱ���������ϣ���������Ϊ����ִ�У�", vbInformation, gstrSysName
                        Exit Function
                    End If
                Else
                    If chk����ִ��.Value = 0 Then
                        ShowMsgBox "Ϊ�ɱ��۵���ʱ,��������ִ��,����!"
                        Exit Function
                    End If
                End If
                
                If .RowData(i) <> -1 Then
                    If InStr(1, strOldID & ",", "," & .RowData(i) & ",") > 0 Then
                        ShowMsgBox "�ڵ�" & i & "����,���ܶ���ͬƷ��(" & .TextMatrix(i, .ColIndex("Ʒ��")) & ")�ظ�����"
                        .Row = i: .Col = .ColIndex("Ʒ��")
                        .SetFocus
                        Exit Function
                    End If
                    strOldID = strOldID & "," & .RowData(i)
                Else
                    If InStr(1, strNewID & ",", "," & lng����ID & ",") > 0 Then
                        MsgBox "���ܶ���ͬƷ��(" & .TextMatrix(i, .ColIndex("Ʒ��")) & ")�ظ����ü۸�", vbExclamation, gstrSysName
                        .Row = i: .Col = .ColIndex("Ʒ��")
                        .SetFocus
                        Exit Function
                    End If
                    strNewID = strNewID & "," & lng����ID
                End If
                
                If Val(.TextMatrix(i, .ColIndex("�ּ�"))) > Val(.TextMatrix(i, .ColIndex("��ָ���ۼ�"))) And Val(.TextMatrix(i, .ColIndex("��ָ���ۼ�"))) <> 0 Then
                    ShowMsgBox "�ڵ�" & i & "����,Ʒ��(" & .TextMatrix(i, .ColIndex("Ʒ��")) & ")���ּ۳�����ָ�����ۼ�(" & Val(.TextMatrix(i, .ColIndex("ָ�����ۼ�"))) & ")"
                    .Row = i: .Col = .ColIndex("�ּ�")
                    .SetFocus
                    Exit Function
                End If
                If IsValied�ɱ���(lng����ID) = False Then
                    .Row = i: .Col = .ColIndex("�ּ�")
                    .SetFocus
                    Exit Function
                End If
            End If
        Next
        
        If blnHaving = False Then
            MsgBox "δ���õ�����Ŀ,����!", vbInformation, gstrSysName
            .Row = 1: .Col = .ColIndex("Ʒ��")
            .SetFocus
            Exit Function
        End If
    End With
    If IsValiedӦ����Ϣ = False Then Exit Function
    ISValied = True
End Function
Public Function IsValied�ɱ���(ByVal lng����ID As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:���ɱ��۵����Ƿ�Ϸ�
    '���:
    '����:
    '����:�Ϸ�����true,���򷵻�False
    '����:���˺�
    '����:2008-11-10 10:04:24
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset, blnHaveData As Boolean
    Dim i As Long
    
    On Error GoTo ErrHandle

    '����ɱ��۵��ۣ���ֱ�ӷ�����
    If m���۷�ʽ = T_�ۼ۵��� Then IsValied�ɱ��� = True: Exit Function
    
    '����Ƿ���δִ�еĳɱ��۵��ۼƻ�
    gstrSQL = "Select 1 From �ɱ��۵�����Ϣ Where ҩƷid = [1] And ִ������ Is Null And Rownum = 1 "
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID)
    If rsTemp.RecordCount = 0 Then
        '��Ҫ���ò����Ƿ����δ��˵���
        If zl����δ��˵���(lng����ID) = True Then
            gstrSQL = "Select ���� From �շ���ĿĿ¼ where id=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID)
            If rsTemp.EOF Then Exit Function
            If MsgBox(rsTemp!���� & "����δ��˵��ݣ������ɱ��ۿ��ܻ���ɲ����" & _
                vbCrLf & Space(4) & "�����ȴ���δ��˵��ݡ��Ƿ񻹼������ۣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
        IsValied�ɱ��� = True:
        Exit Function    '��ʾδ����ִ�м۸������ֱ���˳�(��Ϊ���ܵ����벻���ۣ���ûʲô����)
    End If
    
    '���Ƿ�����Ӧ�ĳɱ��۵���
    With vsStoce
            blnHaveData = False
            For i = 1 To .Rows - 1
                '���ڳɱ��۵�������˷���True
                If lng����ID = Val(.Cell(flexcpData, i, .ColIndex("������Ϣ"))) Then
                    If Val(.TextMatrix(i, .ColIndex("��۵�����"))) <> 0 Then
                        
                        gstrSQL = "Select ���� From �շ���ĿĿ¼ where id=[1]"
                        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID)
                        If rsTemp.EOF Then Exit Function
                        MsgBox "�������ϡ�" & zlStr.Nvl(rsTemp!����) & "������δִ�гɱ��ۣ��������ñ��ε��ۣ�", vbInformation, gstrSysName
                        Exit Function
                    End If
                    blnHaveData = True
                End If
            Next
    End With
    If blnHaveData Then
        '���ڸò���,����Ҫ���ò����Ƿ����δ��˵���
        If zl����δ��˵���(lng����ID) = True Then
            gstrSQL = "Select ���� From �շ���ĿĿ¼ where id=[1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID)
            If rsTemp.EOF Then Exit Function
            If MsgBox(rsTemp!���� & "����δ��˵��ݣ������ɱ��ۿ��ܻ���ɲ����" & _
                vbCrLf & Space(4) & "�����ȴ���δ��˵��ݡ��Ƿ񻹼������ۣ�", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Function
            End If
        End If
    End If
    IsValied�ɱ��� = True:
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveData() As Boolean
    '------------------------------------------------------------------------------
    '����:��������
    '����:
    '����:����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2007/09/17
    '------------------------------------------------------------------------------
    Dim dtToDay As Date, lng�շѼ�Ŀid As Long, lng����ID As Long, lngId As Long, strID As String
    Dim ArrayID As Variant, strTemp As String
    Dim strNo As Variant, i As Long, lng��� As Long
    Dim cllProc As Collection
    
    Set cllProc = New Collection
    
    err = 0: On Error GoTo ErrInfor:
    dtToDay = sys.Currentdate
    
    If m���۷�ʽ = T_�ɱ��۵��� Then
    Else
        lng�շѼ�Ŀid = sys.NextId("�շѼ�Ŀ")
        strNo = sys.GetNextNo(9)
        If IsNull(strNo) Then Exit Function
    End If
    With Me.vsPrice
        strID = ""
        lng��� = 1
        For i = 1 To .Rows - 1
            lng����ID = Val(.Cell(flexcpData, i, .ColIndex("Ʒ��")))
            If lng����ID <> 0 Then
                If Val(.TextMatrix(i, .ColIndex("ԭ����id"))) <> Val(.Cell(flexcpData, i, .ColIndex("��������"))) Or _
                    Val(.TextMatrix(i, .ColIndex("ԭ��"))) <> Val(.TextMatrix(i, .ColIndex("�ּ�"))) _
                    And m���۷�ʽ <> T_�ɱ��۵��� Then
                        lngId = sys.NextId("�շѼ�Ŀ")
                        If Me.chk����ִ��.Value = 1 Then
                            strID = strID & "," & lngId
                        ElseIf .RowData(i) = -1 Then
                            strID = strID & "," & lngId
                        End If
                        If .RowData(i) <> 0 Then
                            '���˺�:��Ҫ�ǽ������Ϊ������,���磺����.����ѵ�
                            '����:9569 2006-11-20
                            'If Val(.TextMatrix(i, col�ּ�)) <> 0 Then
                                '������һ�εļ۸��¼��ִֹ��
                                ' zl_�շѼ�Ŀ_stop (
                                gstrSQL = "zl_�շѼ�Ŀ_stop("
                                '    �շ�ϸĿID_IN IN �շѼ�Ŀ.�շ�ϸĿID%TYPE,
                                gstrSQL = gstrSQL & "" & lng����ID & ","
                                '    ��ֹ����_IN IN �շѼ�Ŀ.��ֹ����%TYPE := NULL
                                If Me.chk����ִ��.Value Then
                                    gstrSQL = gstrSQL & "to_date('" & Format(DateAdd("s", -1, dtToDay), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                                Else
                                    gstrSQL = gstrSQL & "to_date('" & Format(DateAdd("s", -1, Me.dtpִ������.Value), "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
                                End If
                                gstrSQL = gstrSQL & ")"
                                AddArray cllProc, gstrSQL
                                
                                'Zl_�շѼ�Ŀ_Insert
                                gstrSQL = "zl_�շѼ�Ŀ_Insert("
                                '  Id_In         In �շѼ�Ŀ.ID%Type,
                                gstrSQL = gstrSQL & "" & lngId & ","
                                '  ԭ��id_In     In �շѼ�Ŀ.ԭ��id%Type := Null,
                                gstrSQL = gstrSQL & "" & IIf(.RowData(i) = -1, "NUll", .RowData(i)) & ","
                                '  �շ�ϸĿid_In In �շѼ�Ŀ.�շ�ϸĿid%Type := Null,
                                gstrSQL = gstrSQL & "" & lng����ID & ","
                                '  ������Ŀid_In In �շѼ�Ŀ.������Ŀid%Type := Null,
                                gstrSQL = gstrSQL & "" & IIf(Val(.Cell(flexcpData, i, .ColIndex("��������"))) = 0, "NULL", Val(.Cell(flexcpData, i, .ColIndex("��������")))) & ","
                                '  ԭ��_In       In �շѼ�Ŀ.ԭ��%Type := Null,
                                If .TextMatrix(i, .ColIndex("����")) = "ʱ��" And Val(.Cell(flexcpData, i, .ColIndex("����"))) = 0 Then
                                    '�Ǹ����������ϵ�ʵ�����ģ����Է�Χ�����ģ�����Ҫ��ҽ��Ӧ��),ʼ����Ϊ��
                                    gstrSQL = gstrSQL & "" & 0 & ","
                                Else
                                    gstrSQL = gstrSQL & "" & Round(Val(.TextMatrix(i, .ColIndex("ԭ��"))) / Val(.TextMatrix(i, .ColIndex("ϵ��"))), g_С��λ��.obj_ɢװС��.���ۼ�С��) & ","
                                End If
                                
                                '  �ּ�_In       In �շѼ�Ŀ.�ּ�%Type := Null,
                                gstrSQL = gstrSQL & "" & Round(Val(.TextMatrix(i, .ColIndex("�ּ�"))) / Val(.TextMatrix(i, .ColIndex("ϵ��"))), g_С��λ��.obj_ɢװС��.���ۼ�С��) & ","
                                '  �����շ���_In In �շѼ�Ŀ.�����շ���%Type := Null,
                                gstrSQL = gstrSQL & "NULL,"
                                '  �Ӱ�Ӽ���_In In �շѼ�Ŀ.�Ӱ�Ӽ���%Type := Null,
                                gstrSQL = gstrSQL & "NULL,"
                                '  ����˵��_In   In �շѼ�Ŀ.����˵��%Type := Null,
                                gstrSQL = gstrSQL & "'" & Me.txt˵��.Text & "',"
                                '  ����id_In     In �շѼ�Ŀ.����id%Type := Null,
                                gstrSQL = gstrSQL & "" & lng�շѼ�Ŀid & ","
                                '  ������_In     In �շѼ�Ŀ.������%Type := Null,
                                gstrSQL = gstrSQL & "'" & Me.txt������.Text & "',"
                                '  ִ������_In   In �շѼ�Ŀ.ִ������%Type := Null,
                                If Me.chk����ִ��.Value Then
                                    gstrSQL = gstrSQL & "to_date('" & Format(dtToDay, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
                                Else
                                    gstrSQL = gstrSQL & "to_date('" & Format(Me.dtpִ������.Value, "YYYY-MM-DD HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),"
                                End If
                                '  �䶯ԭ��_In   In �շѼ�Ŀ.�䶯ԭ��%Type := 1,
                                gstrSQL = gstrSQL & "" & 0 & ","
                                '  No_In         In �շѼ�Ŀ.NO%Type := Null,
                                gstrSQL = gstrSQL & "'" & strNo & "',"
                                '  ���_In       In �շѼ�Ŀ.���%Type := 1
                                gstrSQL = gstrSQL & "" & lng��� & ","
                                'ȱʡ�۸�_In
                                If .TextMatrix(i, .ColIndex("����")) = "ʱ��" And Val(.Cell(flexcpData, i, .ColIndex("����"))) = 0 Then
                                        gstrSQL = gstrSQL & "" & Round(Val(.TextMatrix(i, .ColIndex("�ּ�"))) / Val(.TextMatrix(i, .ColIndex("ϵ��"))), g_С��λ��.obj_ɢװС��.���ۼ�С��) & ")"
                                Else
                                        gstrSQL = gstrSQL & "NULL)"
                                End If
                                AddArray cllProc, gstrSQL
                                lng��� = lng��� + 1
                        End If
                End If
                '�Ƿ����ָ���۸�ĵ�����������ڣ�����ָ���۸�
                '����ָ�����ۼ�
                If lng����ID <> 0 Then
                    If Val(.TextMatrix(i, .ColIndex("ԭָ���ۼ�"))) <> Val(.TextMatrix(i, .ColIndex("��ָ���ۼ�"))) Then
                        strTemp = Round(Val(.TextMatrix(i, .ColIndex("��ָ���ۼ�"))) / Val(.TextMatrix(i, .ColIndex("ϵ��"))), g_С��λ��.obj_ɢװС��.���ۼ�С��)
                        'zl_��������_UpdateCustom ( ����ID_IN ,SQL_IN)
                        gstrSQL = "zl_��������_UpdateCustom(" & lng����ID & ",'ָ�����ۼ�=" & strTemp & "')"
                        AddArray cllProc, gstrSQL
                    End If
                    '���²ɹ��޼�
                    If Val(.TextMatrix(i, .ColIndex("ԭ�ɹ��޼�"))) <> Val(.TextMatrix(i, .ColIndex("�ֲɹ��޼�"))) Then
                        strTemp = Round(Val(.TextMatrix(i, .ColIndex("�ֲɹ��޼�"))) / Val(.TextMatrix(i, .ColIndex("ϵ��"))), g_С��λ��.obj_ɢװС��.�ɱ���С��)
                        'zl_��������_UpdateCustom ( ����ID_IN ,SQL_IN)
                        gstrSQL = "zl_��������_UpdateCustom(" & lng����ID & ",'ָ��������=" & strTemp & "')"
                        AddArray cllProc, gstrSQL
                    End If
                End If
            End If
        Next
    End With
    
    Dim lng��Ӧ��ID As Long, lng���� As Long, lng�ⷿID As Long, dbl�ɱ��� As Double
    Dim str��Ʊ�� As String, str��Ʊ���� As String, dbl��Ʊ��� As Double, lngϵ�� As Long, j As Long
    
    Dim cllTemp As Collection
    
    '�ɱ��۵��۴���
    If m���۷�ʽ = T_�ɱ��۵��� Or m���۷�ʽ = T_�ɱ����ۼ۵��� Then
        With vsStoce
            For i = 1 To .Rows - 1
                lng�ⷿID = Val(.Cell(flexcpData, i, .ColIndex("�ⷿ")))
                lng��Ӧ��ID = Val(.Cell(flexcpData, i, .ColIndex("��Ӧ��")))
                lng����ID = Val(.Cell(flexcpData, i, .ColIndex("������Ϣ")))
                lng���� = Val(.Cell(flexcpData, i, .ColIndex("����")))
                lngϵ�� = Val(.Cell(flexcpData, i, .ColIndex("��λ")))
                If lng����ID <> 0 Then
                    str��Ʊ�� = "": str��Ʊ���� = "": dbl��Ʊ��� = 0
                    If chkӦ��.Value = 1 Then
                        With vsPay
                            For j = 1 To .Rows - 1
                                If Val(.Cell(flexcpData, j, .ColIndex("������Ϣ"))) = lng����ID And _
                                    Val(.Cell(flexcpData, j, .ColIndex("��Ӧ��"))) = lng��Ӧ��ID Then
                                    '���Ƿ��д��������Ͽ��䶯���
                                    str��Ʊ�� = Trim(.TextMatrix(j, .ColIndex("��Ʊ��")))
                                    str��Ʊ���� = Trim(.TextMatrix(j, .ColIndex("��Ʊ����")))
                                    dbl��Ʊ��� = Val(.TextMatrix(j, .ColIndex("��Ʊ���")))
                                    Exit For
                                End If
                            Next
                        End With
                    End If
                    
                    dbl�ɱ��� = Round(Val(.TextMatrix(i, .ColIndex("�ֳɱ���"))) / lngϵ��, g_С��λ��.obj_ɢװС��.�ɱ���С��)
                    
                    ' Zl_���ϳɱ�����_Insert
                    gstrSQL = "Zl_���ϳɱ�����_Insert("
                    '  ��ҩ��λid_In In �ɱ��۵�����Ϣ.��ҩ��λid%Type,
                    gstrSQL = gstrSQL & IIf(lng��Ӧ��ID = 0, "Null", lng��Ӧ��ID) & ","
                    '  �ⷿid_In     In �ɱ��۵�����Ϣ.�ⷿid%Type,
                    gstrSQL = gstrSQL & "" & lng�ⷿID & ","
                    '  ����id_In     In �ɱ��۵�����Ϣ.ҩƷid%Type,
                    gstrSQL = gstrSQL & "" & lng����ID & ","
                    '  ����_In       In �ɱ��۵�����Ϣ.����%Type := Null,
                    gstrSQL = gstrSQL & "" & lng���� & ","
                    '  ԭ�ɱ���_In   In �ɱ��۵�����Ϣ.ԭ�ɱ���%Type := Null,
                    gstrSQL = gstrSQL & "" & Round(Val(.Cell(flexcpData, i, .ColIndex("ԭ�ɱ���"))), g_С��λ��.obj_ɢװС��.�ɱ���С��) & ","
                    '  �³ɱ���_In   In �ɱ��۵�����Ϣ.�³ɱ���%Type := Null,
                    gstrSQL = gstrSQL & "" & dbl�ɱ��� & ","
                    '  ��Ʊ��_In     In �ɱ��۵�����Ϣ.��Ʊ��%Type := Null,
                    gstrSQL = gstrSQL & "'" & str��Ʊ�� & "',"
                    '  ��Ʊ����_In   In �ɱ��۵�����Ϣ.��Ʊ����%Type := Null,
                    gstrSQL = gstrSQL & "" & IIf(str��Ʊ���� = "", "NULL", "to_date('" & str��Ʊ���� & "','yyyy-mm-dd') ") & ","
                    '  ��Ʊ���_In   In �ɱ��۵�����Ϣ.��Ʊ���%Type := Null,
                    gstrSQL = gstrSQL & "" & dbl��Ʊ��� & ","
                    '  Ӧ����䶯_In In �ɱ��۵�����Ϣ.Ӧ����䶯%Type := 0
                    gstrSQL = gstrSQL & "" & IIf(chkӦ��.Value = 1 And lng��Ӧ��ID <> 0 And dbl��Ʊ��� <> 0, 1, 0) & ")"
                    AddArray cllProc, gstrSQL
                End If
            Next
        End With
    End If
    
    
    '����������¶Գɱ��۽��е���:
    '1.����Ϊ�ɱ��۵��ۼ�����ִ��ʱ�������Գɱ��۽��е���
    '2.��������ִ�кͷǳɱ���(���ɱ��۵��۷�ʽ)����ʱ�����������ϵ���ʱ����ִ�С�
     '�����ɱ��۵���ʱ
    If m���۷�ʽ = T_�ɱ��۵��� And Me.chk����ִ��.Value = 1 Then
        With vsPrice
            For i = 1 To .Rows - 1
                lng����ID = Val(.Cell(flexcpData, i, .ColIndex("Ʒ��")))
                If lng����ID <> 0 Then
                  ' Zl_�����շ���¼_Adjust
                  gstrSQL = "Zl_�����շ���¼_Adjust("
                  '  ����id_In In Number, --���ۼ�¼��ID
                  gstrSQL = gstrSQL & "" & 0 & ","
                  '  ����_In   In Number := 0, --�Ƿ�תΪ�������ۣ����²������ԡ��շ�ϸĿ�еı�ۣ�
                  gstrSQL = gstrSQL & "" & 0 & ","
                  '  ����id_In In Number := 0 --����Ϊ0ʱ��ʾ�ǳɱ��۵��ۣ��������ۼ��������
                    gstrSQL = gstrSQL & "" & lng����ID & ")"
                  AddArray cllProc, gstrSQL
                End If
            Next
        End With
    End If
    
    If strID <> "" Then strID = Mid(strID, 2)
    'ѭ��ִ�й���
    ArrayID = Split(strID, ",")
    For i = 0 To UBound(ArrayID)
        If Val(ArrayID(i)) <> 0 Then
            'Zl_�����շ���¼_Adjust
            gstrSQL = "zl_�����շ���¼_adjust("
            '  Adjustid In Number, --���ۼ�¼��ID
            gstrSQL = gstrSQL & "" & ArrayID(i) & ","
            '  Bln����  In Number := 0 --�Ƿ�תΪ�������ۣ�����ҩƷĿ¼���շ�ϸĿ�еı�ۣ�
            gstrSQL = gstrSQL & "" & Me.Chk����.Value & ")"
            AddArray cllProc, gstrSQL
        End If
    Next
    err = 0: On Error GoTo ErrHand:
    ExecuteProcedureArrAy cllProc, Me.Caption
    mlngBillId = 0
    mlngStuffId = 0
    SaveData = True
    Exit Function
ErrInfor:
    If ErrCenter = 1 Then
        Resume
    End If
    Exit Function
ErrHand:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Function
 

 Private Sub printbill(ByVal intPrintMode As Byte)
    '-------------------------------------------------------------------------------------
    '����:��ӡ
    '����:intPrintMode-1-��ӡ,2-Ԥ��,3-Excel
    '-------------------------------------------------------------------------------------
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    
    If Trim(Me.vsPrice.TextMatrix(1, 0)) = "" Then Exit Sub
    objPrint.Title.Text = "���ĵ���֪ͨ��"
    
    Set objRow = New zlTabAppRow
    objRow.Add "����˵��:" & Me.txt˵��.Text
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "ִ��ʱ��:" & Format(IIf(Me.chk����ִ��.Value, sys.Currentdate, Me.dtpִ������.Value), "yyyy��MM��DD�� HH:mm:ss")
    objRow.Add "������:" & Me.txt������.Text
    objPrint.UnderAppRows.Add objRow
    
    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & gstrUserName
    objRow.Add "��ӡʱ��:" & Format(sys.Currentdate, "yyyy��MM��DD�� HH:mm:ss")
    objPrint.BelowAppRows.Add objRow
    
    Set objPrint.Body = Me.vsPrice
    objPrint.PageFooter = 2
     
    If intPrintMode = 1 Then
        Select Case zlPrintAsk(objPrint)
        Case 1
             zlPrintOrView1Grd objPrint, 1
        Case 2
            zlPrintOrView1Grd objPrint, 2
        Case 3
            zlPrintOrView1Grd objPrint, 3
        End Select
    Else
        zlPrintOrView1Grd objPrint, intPrintMode
    End If
    Set objPrint = Nothing
End Sub

Private Sub cmdPrintStoce_Click()
    Dim objPrint As New zlPrint1Grd
    Dim objRow As New zlTabAppRow
    Dim i As Long
    
    
    
    If Trim(vsStoce.TextMatrix(1, vsStoce.ColIndex("������Ϣ"))) = "" Then Exit Sub

    objPrint.Title.Text = "���ۿ��䶯��"

    Set objRow = New zlTabAppRow
    objRow.Add "����˵��:" & Me.txt˵��.Text
    objPrint.UnderAppRows.Add objRow

    Set objRow = New zlTabAppRow
    objRow.Add "ִ��ʱ��:" & Format(IIf(Me.chk����ִ��.Value, sys.Currentdate, Me.dtpִ������.Value), "yyyy��MM��DD�� HH:mm:ss")
    objRow.Add "������:" & Me.txt������.Text
    objPrint.UnderAppRows.Add objRow

    Set objRow = New zlTabAppRow
    objRow.Add "��ӡ��:" & gstrUserName
    objRow.Add "��ӡʱ��:" & Format(sys.Currentdate, "yyyy��MM��DD�� HH:mm:ss")
    objPrint.BelowAppRows.Add objRow
    '��������������еĿ��
    With vsStoce
        For i = 0 To .Cols - 1
            If .ColHidden(i) Then
                .ColData(i) = .ColWidth(i)
                .ColWidth(i) = 0
            End If
        Next
    End With
    
    Set objPrint.Body = vsStoce
    objPrint.PageFooter = 2

    Select Case zlPrintAsk(objPrint)
    Case 1
         zlPrintOrView1Grd objPrint, 1
    Case 2
        zlPrintOrView1Grd objPrint, 2
    Case 3
        zlPrintOrView1Grd objPrint, 3
    End Select
    Set objPrint = Nothing
    '��ӡ��ɺ�,�ָ���������еĿ��
    With vsStoce
        For i = 0 To .Cols - 1
            If .ColHidden(i) Then
                .ColWidth(i) = Val(.ColData(i))
                .ColData(i) = ""
            End If
        Next
    End With

End Sub
  
Private Sub SetCtlEnabled()
    '---------------------------------------------------------------------------------------------
    '����:������ؿؼ���Enabled����
    '����:
    '����:���˺�
    '����:2007/07/17
    '---------------------------------------------------------------------------------------------
    If mBillType <> B_���� Then
        With vsPrice
            .Editable = flexEDKbdMouse
            
        End With
        Exit Sub
    End If
    vsPrice.Editable = flexEDNone
    DkPane.Panes(ID_PANE_SEARCH).Close
    
    'DkPane.CloseAll
    Me.txt˵��.Enabled = False
    Me.chk����ִ��.Value = 0
    Me.chk����ִ��.Enabled = False
    Me.dtpִ������.Enabled = False
End Sub
Private Sub InitBill()
    '---------------------------------------------------------------------------------------------
    '����:��ʼ��������Ϣ
    '����:
    '����:���˺�
    '����:2007/07/17
    '---------------------------------------------------------------------------------------------
    Dim dtDate As Date, rsTemp As New ADODB.Recordset, i As Long
    dtDate = sys.Currentdate
    
    On Error GoTo ErrHandle

    If mlngBillId = 0 Then
        '������۱༭״̬
        stbThis.Panels(2).Text = "���䶯��(���ڵ���δ���棬��ӳ�Ŀ����ܲ�׼ȷ)"
         Me.dtpִ������.MinDate = DateAdd("s", 1, dtDate)
        Me.dtpִ������.Value = DateAdd("d", 1, dtDate)
        Me.txt������.Text = gstrUserName
        
        If mlngStuffId = 0 Then Exit Sub
        '���ָ�����ȵ��۵����ģ���ֱ�ӽ������ĵ���
        gstrSQL = "" & _
            "   Select I.ID,I.����,I.����,I.���,I.����,I.���㵥λ,M.����ID,J.���� ||'-'||J.���� as ����," & _
            "           P.��װ��λ,decode(I.�Ƿ���,1,'ʱ��','����') ����,p.ָ�����ۼ�," & _
            "           P.ָ��������,P.ָ�����ۼ�,p.�ɱ���," & _
                        IIf(mintUnit = 0, "1", "nvl(p.����ϵ��,1)") & " as ����ϵ��" & _
            "   From �շ���ĿĿ¼ I,�������� P,������ĿĿ¼ M,���Ʒ���Ŀ¼ J" & _
            "   Where I.ID=[1] And I.ID=P.����ID And P.����ID=M.id and M.����ID=J.id(+)"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngStuffId)
        With vsPrice
            If rsTemp.EOF Then
                Exit Sub
            End If
            txt����.Text = zlStr.Nvl(rsTemp!����)
            txt����.Tag = zlStr.Nvl(rsTemp!����id)
            .Redraw = flexRDNone
            .TextMatrix(.Row, .ColIndex("Ʒ��")) = "[" & zlStr.Nvl(rsTemp!����) & "]" & zlStr.Nvl(rsTemp!����)
            .Cell(flexcpData, .Row, .ColIndex("Ʒ��")) = zlStr.Nvl(rsTemp!Id)
            .TextMatrix(.Row, .ColIndex("���")) = zlStr.Nvl(rsTemp!���)
            .TextMatrix(.Row, .ColIndex("����")) = zlStr.Nvl(rsTemp!����)
            .TextMatrix(.Row, .ColIndex("��λ")) = IIf(mintUnit = 0, zlStr.Nvl(rsTemp!���㵥λ), zlStr.Nvl(rsTemp!��װ��λ))
            .TextMatrix(.Row, .ColIndex("����")) = zlStr.Nvl(rsTemp!����)
            .TextMatrix(.Row, .ColIndex("ϵ��")) = zlStr.Nvl(rsTemp!����ϵ��)
            .TextMatrix(.Row, .ColIndex("ԭ�ɱ���")) = Format(Val(zlStr.Nvl(rsTemp!�ɱ���)) * Val(zlStr.Nvl(rsTemp!����ϵ��)), mFMT.FM_�ɱ���)
            .Cell(flexcpData, .Row, .ColIndex("ԭ�ɱ���")) = Val(zlStr.Nvl(rsTemp!�ɱ���))
            .TextMatrix(.Row, .ColIndex("�ֳɱ���")) = Format(Val(zlStr.Nvl(rsTemp!�ɱ���)) * Val(zlStr.Nvl(rsTemp!����ϵ��)), mFMT.FM_�ɱ���)
            .Cell(flexcpData, .Row, .ColIndex("�ֳɱ���")) = Val(zlStr.Nvl(rsTemp!�ɱ���))
            
            .TextMatrix(.Row, .ColIndex("ԭ�ɹ��޼�")) = Format(Val(zlStr.Nvl(rsTemp!ָ��������)) * Val(rsTemp!����ϵ��), mFMT.FM_�ɱ���)
            .TextMatrix(.Row, .ColIndex("�ֲɹ��޼�")) = .TextMatrix(.Row, .ColIndex("ԭ�ɹ��޼�"))
            .Cell(flexcpData, .Row, .ColIndex("ԭ�ɹ��޼�")) = Val(zlStr.Nvl(rsTemp!ָ��������))
            .Cell(flexcpData, .Row, .ColIndex("�ֲɹ��޼�")) = Val(zlStr.Nvl(rsTemp!ָ��������))
            
            .TextMatrix(.Row, .ColIndex("ָ�����ۼ�")) = Format(Val(zlStr.Nvl(rsTemp!ָ�����ۼ�)) * Val(rsTemp!����ϵ��), mFMT.FM_���ۼ�)
            .TextMatrix(.Row, .ColIndex("ԭָ���ۼ�")) = .TextMatrix(.Row, .ColIndex("ָ�����ۼ�"))
            .TextMatrix(.Row, .ColIndex("��ָ���ۼ�")) = .TextMatrix(.Row, .ColIndex("ָ�����ۼ�"))
            
            .Cell(flexcpData, .Row, .ColIndex("ָ�����ۼ�")) = Val(zlStr.Nvl(rsTemp!ָ�����ۼ�))
            .Cell(flexcpData, .Row, .ColIndex("ԭָ���ۼ�")) = Val(zlStr.Nvl(rsTemp!ָ�����ۼ�))
            .Cell(flexcpData, .Row, .ColIndex("��ָ���ۼ�")) = Val(zlStr.Nvl(rsTemp!ָ�����ۼ�))
            Call zlGetPrice(Val(zlStr.Nvl(rsTemp!Id)), IIf(.TextMatrix(.Row, .ColIndex("����")) = "ʱ��", True, False), False, .Row)
            Call LoadStockData(Val(zlStr.Nvl(rsTemp!Id)), Val(.Cell(flexcpData, .Row, .ColIndex("ԭ��"))), Val(.Cell(flexcpData, .Row, .ColIndex("�ּ�"))))
            .Col = .ColIndex("�ּ�")
            .Redraw = flexRDBuffered
            mlngPreRow = 0:
            Call vsPrice_RowColChange
            Exit Sub
        End With
    End If
    
    '���������ʾ״̬
    Dim strBills As String
    strBills = ""
    gstrSQL = "" & _
        "   Select P.ID,M.id as ����id,'['||M.����||']'||M.���� as Ʒ�� ,decode(M.�Ƿ���,1,'ʱ��','����') ����,M.���,M.����,M.���㵥λ as ��λ," & _
                IIf(mintUnit = 0, "1", " nvl(j.����ϵ��,1) ") & " as ����ϵ�� ,j.��װ��λ," & _
        "        P.ԭ��,P.�ּ�,P.������Ŀid,I.���� as ��������," & _
        "        To_Char(P.ִ������,'yyyy-MM-dd hh24:mi:ss') ִ������,P.�䶯ԭ��,P.����˵��,P.������,j.�ɱ���,j.ָ�����ۼ�" & _
        "   From �շѼ�Ŀ P,�շ���ĿĿ¼ M,������Ŀ I,�������� J" & _
        "   Where P.�շ�ϸĿid=M.id and P.������Ŀid=I.id And M.ID=J.����ID and P.ID=[1] " & _
        GetPriceClassString("P") & _
        "   Order by P.id"                            '�����IDȡ���Ǽ۸��¼ID����һ��ID
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngBillId)
    i = 1
    With vsPrice
        .Redraw = flexRDNone
        If rsTemp.EOF = False Then
            Me.txt˵�� = zlStr.Nvl(rsTemp!����˵��)
            Me.txt������.Text = zlStr.Nvl(rsTemp!������)
            Me.dtpִ������.Value = zlStr.Nvl(rsTemp!ִ������)
            .Rows = rsTemp.RecordCount + 1
        Else
            .Rows = 2
        End If
        Do While Not rsTemp.EOF
            strBills = strBills & "," & rsTemp!Id
            .RowData(i) = Val(zlStr.Nvl(rsTemp!Id))
            .TextMatrix(i, .ColIndex("Ʒ��")) = zlStr.Nvl(rsTemp!Ʒ��)
            .Cell(flexcpData, i, .ColIndex("Ʒ��")) = zlStr.Nvl(rsTemp!����ID)
            .TextMatrix(i, .ColIndex("���")) = zlStr.Nvl(rsTemp!���)
            .TextMatrix(i, .ColIndex("����")) = zlStr.Nvl(rsTemp!����)
            .TextMatrix(i, .ColIndex("��λ")) = IIf(mintUnit = 0, zlStr.Nvl(rsTemp!��λ), zlStr.Nvl(rsTemp!��װ��λ))
            .TextMatrix(i, .ColIndex("����")) = zlStr.Nvl(rsTemp!����)
            .TextMatrix(i, .ColIndex("ϵ��")) = zlStr.Nvl(rsTemp!����ϵ��)
            .TextMatrix(i, .ColIndex("�ֳɱ���")) = Format(Val(zlStr.Nvl(rsTemp!�ɱ���)) * Val(rsTemp!����ϵ��), mFMT.FM_�ɱ���)
            .TextMatrix(i, .ColIndex("ָ�����ۼ�")) = Format(Val(zlStr.Nvl(rsTemp!ָ�����ۼ�)) * Val(rsTemp!����ϵ��), mFMT.FM_���ۼ�)
            .TextMatrix(i, .ColIndex("�ϴ�����")) = Format(rsTemp!ִ������, "YYYY-MM-DD HH:MM:SS")
            .TextMatrix(i, .ColIndex("ԭ��")) = Format(Val(zlStr.Nvl(rsTemp!ԭ��)) * Val(zlStr.Nvl(rsTemp!����ϵ��)), mFMT.FM_���ۼ�)
            .TextMatrix(i, .ColIndex("�ּ�")) = Format(Val(zlStr.Nvl(rsTemp!�ּ�)) * Val(zlStr.Nvl(rsTemp!����ϵ��)), mFMT.FM_���ۼ�)
            .TextMatrix(i, .ColIndex("ԭ����id")) = Val(zlStr.Nvl(rsTemp!������Ŀid))
            .TextMatrix(i, .ColIndex("��������")) = zlStr.Nvl(rsTemp!��������)
            .Cell(flexcpData, i, .ColIndex("��������")) = Val(zlStr.Nvl(rsTemp!������Ŀid))
            
            If zlStr.Nvl(rsTemp!ִ������) <= Format(dtDate, "yyyy-mm-dd HH:MM:SS") And rsTemp!�䶯ԭ�� = 0 Then       'δ���е��ۼ���,��ִ�м���
                gstrSQL = "zl_�����շ���¼_Adjust(" & rsTemp!Id & ")"
                zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
            End If
            rsTemp.MoveNext
        Loop
        .Col = .ColIndex("�ּ�")
        .Redraw = flexRDBuffered
    End With
    If strBills <> "" Then strBills = Mid(strBills, 2)
    
    If rsTemp.RecordCount = 0 Then
        mlngPreRow = 0:
        Call vsPrice_RowColChange
        Exit Sub
    End If
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst

    If rsTemp!ִ������ > dtDate Then
        '���ִ��ʱ��δ������ֻ��ģ����ʾ���䶯
        Me.stbThis.Panels(2).Text = "���䶯��(����ִ��ʱ��δ������ӳ�Ŀ����ܲ�׼ȷ)"
    Else
        'ִ��ʱ���ѵ����϶�Ҳ�����˵��ۼ��㣬ֱ�Ӵ��շ���¼��ȡ���۱䶯���
        gstrSQL = "" & _
        "   Select S.ID,S.ҩƷID as ����ID,D.���� as �ⷿ,'['||M.����||']'||M.���� as ������Ϣ,M.���,M.����,M.���㵥λ as ��λ, " & _
        "       P.��װ��λ,P.����ϵ��,S.����,S.����,S.ԭ��,S.�ּ�,S.�������" & _
        "   From (  Select ID,�ⷿID,ҩƷID,����,��д���� as ����,�ɱ��� as ԭ��,���ۼ� as �ּ�,���۽�� as �������" & _
        "           From (  Select P.ID,N.�ⷿID,N.ҩƷID,N.����,N.��д����,N.�ɱ���,N.���ۼ�,N.���۽��" & _
        "                   From ҩƷ�շ���¼ N, (select ID,�շ�ϸĿID,ִ������,��ֹ���� from �շѼ�Ŀ where ID=[1]" & _
        GetPriceClassString("") & ") P" & _
        "                   where   N.ҩƷID=P.�շ�ϸĿID and N.����=13 and N.����ID is null " & _
        "                           and N.������� Between P.ִ������ and nvl(P.��ֹ����,sysdate))) S," & _
        "       ���ű� D,�շ���ĿĿ¼ M,�������� P" & _
        " where S.�ⷿid+0=D.id and S.ҩƷID=M.ID And M.ID=P.����ID" & _
        " order by M.����,S.����"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngBillId)
        With vsStoce
            .Rows = 2
            .Clear 1
            If rsTemp.RecordCount > 0 Then .Rows = rsTemp.RecordCount + 1
            i = 1
            Do While Not rsTemp.EOF
                .TextMatrix(i, .ColIndex("�ⷿ")) = zlStr.Nvl(rsTemp!�ⷿ)
                .TextMatrix(i, .ColIndex("������Ϣ")) = zlStr.Nvl(rsTemp!������Ϣ)
                .TextMatrix(i, .ColIndex("���|����")) = IIf(IsNull(rsTemp!���), "", rsTemp!���) & IIf(IsNull(rsTemp!����), "", "|" & rsTemp!����)
                If mintUnit = 0 Then
                    .TextMatrix(i, .ColIndex("��λ")) = IIf(IsNull(rsTemp!��λ), "", rsTemp!��λ)
                Else
                    .TextMatrix(i, .ColIndex("��λ")) = IIf(IsNull(rsTemp!��װ��λ), "", rsTemp!��װ��λ)
                End If
                .TextMatrix(i, .ColIndex("����")) = IIf(IsNull(rsTemp!����), "", rsTemp!����)
                .TextMatrix(i, .ColIndex("����")) = Format(Val(zlStr.Nvl(rsTemp!����)) / Val(zlStr.Nvl(rsTemp!����ϵ��)), mFMT.FM_����)
                .TextMatrix(i, .ColIndex("ԭ��")) = Format(Val(zlStr.Nvl(rsTemp!ԭ��)) * Val(zlStr.Nvl(rsTemp!����ϵ��)), mFMT.FM_���ۼ�)
                .TextMatrix(i, .ColIndex("�ּ�")) = Format(Val(zlStr.Nvl(rsTemp!�ּ�)) * Val(zlStr.Nvl(rsTemp!����ϵ��)), mFMT.FM_���ۼ�)
                .TextMatrix(i, .ColIndex("������")) = Format(Val(zlStr.Nvl(rsTemp!�������)), mFMT.FM_���)
                i = i + 1
                rsTemp.MoveNext
            Loop
        End With
        mlngPreRow = 0:
        Call vsPrice_RowColChange
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    '-----------------������ʾ����---------------------------------
    Me.Caption = "�������ϵ���"
    Call InitOther
    Call InitCommandBar
    Call InitPanel
    
    '��ʼҳ��
    Call InitPage
    Call InitControl
    Call InitBill
    Call SetControlVisble
    Call SetCtlEnabled
    Call SetColor(1)
    '-----------------------------------------------------------
     
    zl_vsGrid_Para_Restore mlngModule, vsPay, Me.Caption, "Ӧ���䶯"
    zl_vsGrid_Para_Restore mlngModule, vsPay, Me.Caption, "���䶯"
    DkPane.RecalcLayout
    With vsStoce
        .Cell(flexcpBackColor, 1, 1, .Rows - 1, .Cols - 1) = &H8000000F
    End With
    With vsPay
        .Cell(flexcpBackColor, 1, 1, .Rows - 1, .Cols - 1) = &H8000000F
    End With
    
    mblnSucces = False
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 39 Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    mblnFirst = True
    Call RestoreWinState(Me)
    
    mlng��Ӧ��ID = 0
    mdbl�ӳ��� = 0
    '�ж��Ƿ��Կⷿ��λ��ʾ
    mintUnit = Get���۵�λ
    mlngModule = 1711
    mstrPrivs = ";" & GetPrivFunc(glngSys, mlngModule) & ";"
    
    
    '���˺�:����С����ʽ����
    With mFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���)
        .FM_��� = GetFmtString(mintUnit, g_���)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�)
        .FM_���� = GetFmtString(mintUnit, g_����)
    End With
    With mOraFMT
        .FM_�ɱ��� = GetFmtString(mintUnit, g_�ɱ���, True)
        .FM_��� = GetFmtString(mintUnit, g_���, True)
        .FM_���ۼ� = GetFmtString(mintUnit, g_�ۼ�, True)
        .FM_���� = GetFmtString(mintUnit, g_����, True)
    End With
    Call vsPrice_LostFocus
    Call vsStoce_LostFocus
    Call vsPay_LostFocus
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = 1 Then Exit Sub
    chkAppAllColumn.Move 0, 0
    
    If Me.Height < 5000 Then
        Me.Height = 5000
    End If
    If Me.Width < 9720 Then
        Me.Width = 9720
    End If
    Dim panKind As Pane
    Set panKind = Me.DkPane.FindPane(ID_PANE_SEARCH)
    If Not panKind Is Nothing Then
        panKind.MinTrackSize.SetSize 295, Me.ScaleHeight / Screen.TwipsPerPixelY
        panKind.MaxTrackSize.SetSize 400, Me.ScaleHeight / Screen.TwipsPerPixelY
    End If
    Set panKind = Me.DkPane.FindPane(ID_PANE_STOCE)
    If Not panKind Is Nothing Then
        panKind.MinTrackSize.Height = 50
        panKind.MaxTrackSize.Height = (Me.ScaleHeight * 0.7) / Screen.TwipsPerPixelY
    End If
    Me.DkPane.RecalcLayout
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnModify Then If MsgBox("��ȷ��Ҫ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Cancel = 1: Exit Sub
    SaveWinState Me
    DkPane.SaveState Me.Caption & "_Search", App.Title, "Layout"
    
     
    zl_vsGrid_Para_Save mlngModule, vsPay, Me.Caption, "Ӧ���䶯"
    zl_vsGrid_Para_Save mlngModule, vsStoce, Me.Caption, "���䶯"
End Sub
Private Sub txt˵��_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 And Me.dtpִ������.Enabled Then Me.dtpִ������.SetFocus
End Sub

Private Sub DkPane_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.Id
    Case ID_PANE_SEARCH
        If Item.Handle = 0 Then Item.Handle = picSeach.hwnd
    Case ID_PANE_PRICE
        If Item.Handle = 0 Then Item.Handle = picPrice.hwnd
    Case ID_PANE_STOCE
        If Item.Handle = 0 Then Item.Handle = picStoceBack.hwnd
    End Select
End Sub
 
'***************************************************************************************************************
'**���䶯��Ӧ���䶯����
Private Sub SetControlVisble()
    '-----------------------------------------------------------------------------------------------------------
    '����:���ÿؼ���Eanbled��Visble����
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-11-05 15:40:21
    '-----------------------------------------------------------------------------------------------------------
    Dim bln���ڳɱ��۵��� As Boolean, blnָ���۹��� As Boolean
    
    bln���ڳɱ��۵��� = (m���۷�ʽ = T_�ɱ��۵��� Or m���۷�ʽ = T_�ɱ����ۼ۵���)
    If mBillType = B_���� Then
        bln���ڳɱ��۵��� = False
    End If
    
    If m���۷�ʽ = T_�ɱ��۵��� Then
        chk����ִ��.Value = 1
        chk����ִ��.Enabled = False
        Chk����.Visible = False
        fra������.Enabled = False
        txt������.Text = ""
    Else
        chk����ִ��.Enabled = True
        Chk����.Visible = True
        fra������.Enabled = True
    End If
        
    With vsStoce
        '����ǳɱ��۵���,��Ҫ����༭
        .Editable = IIf(bln���ڳɱ��۵��� And mBillType <> B_����, flexEDKbdMouse, flexEDNone)
        .ColHidden(.ColIndex("ԭ�ɱ���")) = Not bln���ڳɱ��۵���
        .ColHidden(.ColIndex("�ֳɱ���")) = Not bln���ڳɱ��۵���
        .ColHidden(.ColIndex("�ӳ���")) = Not bln���ڳɱ��۵���
        .ColHidden(.ColIndex("��۵�����")) = Not bln���ڳɱ��۵���
        
        .ColHidden(.ColIndex("������")) = m���۷�ʽ = T_�ɱ��۵���
        '.ColHidden(.ColIndex("ԭ��")) = m���۷�ʽ = T_�ɱ��۵���
        '.ColHidden(.ColIndex("�ּ�")) = m���۷�ʽ = T_�ɱ��۵���
    End With
    
    chkӦ��.Visible = bln���ڳɱ��۵���
    chk����.Visible = bln���ڳɱ��۵���
    '���Ƿ��д�ҳ��Ϣû��
    tbPage.Item(1).Visible = bln���ڳɱ��۵��� And chkӦ��.Value = 1
    With vsPay
        .Editable = IIf(bln���ڳɱ��۵���, flexEDKbdMouse, flexEDNone)
    End With
    '����Ƿ����ָ���۸����Ȩ��
    blnָ���۹��� = InStr(1, mstrPrivs, ";ָ���۸����;") > 0 And Not (mBillType = B_����)
    With vsPrice
        .ColHidden(.ColIndex("ԭ�ɹ��޼�")) = Not blnָ���۹���
        .ColHidden(.ColIndex("�ֲɹ��޼�")) = Not blnָ���۹���
        .ColHidden(.ColIndex("ԭָ���ۼ�")) = Not blnָ���۹���
        .ColHidden(.ColIndex("��ָ���ۼ�")) = Not blnָ���۹���
    End With
    fraCost.Enabled = bln���ڳɱ��۵���
    If fraCost.Enabled = False Then
        txtPriver.BackColor = &H8000000F
        cmdPriver.BackColor = &H8000000F
        txt�ӳ���.BackColor = &H8000000F
    End If
End Sub
Private Sub InitControl()
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʼ���ؼ���Ĭ������
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2008-11-05 15:33:54
    '-----------------------------------------------------------------------------------------------------------
    With vsPrice
        .GridLines = flexGridInset
    End With
    With vsPay
        .Clear 1
        .Rows = 2
        .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
        .SelectionMode = flexSelectionByRow
        .GridLines = flexGridInset
    End With
    With vsStoce
        .Clear 1
        .Rows = 2
        .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
        .GridLines = flexGridInset
    End With
    
End Sub

Private Function MoveStockData(ByVal lng����ID As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:�Ƴ�ָ�����ϵ�����
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-11-05 12:00:34
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Long
    Dim lngRow As Long
    err = 0: On Error GoTo ErrHand:
    
    With vsStoce
        lngRow = 1
ReDo:
        If .Rows > 2 Then
            For i = .FixedRows To .Rows - 1
                If Val(.Cell(flexcpData, i, .ColIndex("������Ϣ"))) = lng����ID Then
                    lngRow = i
                    .RemoveItem i
                    GoTo ReDo
                End If
            Next
        End If
        If .Rows <= 2 Then
            If Val(.Cell(flexcpData, 1, .ColIndex("������Ϣ"))) = lng����ID Then
                .Rows = 2
                .Clear 1
                .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
            End If
        End If
        .Row = 1
    End With
    MoveStockData = True
    Exit Function
ErrHand:
        If ErrCenter = 1 Then
            Resume
        End If
End Function
Private Function LoadStockData(ByVal lng����ID As Long, ByVal dblԭ�� As Double, ByVal dbl�ּ� As Double, Optional bln���� As Boolean = False) As Boolean
   '-----------------------------------------------------------------------------------------------------------
    '����:���ؿⷿ����
    '���:
    '����:
    '����: ����true,���򷵻�False
    '����:���˺�
    '����:2008-11-05 11:57:09
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long, lng��Ӧ��ID As Long
    Dim dblԭ�ɱ��� As Double, dbl�ֳɱ��� As Double, dbl�ӳ��� As Double, dbl��Ʊ��� As Double
    Dim dblTemp As Double
    
    err = 0: On Error GoTo ErrHand:
        
    '���Ƴ�������
    Call MoveStockData(lng����ID)
 
    gstrSQL = "" & _
    "   Select S.�ⷿID, S.ҩƷID as ����ID,S.����, " & _
    "           D.���� as �ⷿ,decode(L.����,NULL ,'','['||L.����||']') ||L.���� as ��Ӧ��, " & _
    "           '['||M.����||']'||M.���� as ����,M.���,M.����,M.���㵥λ," & _
    "           Nvl(M.�Ƿ���, 0) ���,S.����,S.����,S.ʱ���ۼ�,S.�ɱ���,S.�ϴι�Ӧ��ID," & _
    "           p.��װ��λ,P.ָ������� As �����," & IIf(mintUnit = 0, "1", "nvl(p.����ϵ��,1)") & " as ����ϵ��" & _
    "   From (  Select  S.�ⷿID,S.ҩƷID,S.�ϴι�Ӧ��ID,S.�ϴ����� ����,S.ʵ������ as ����,S.����, " & _
    "                 decode(nvl(���ۼ�,0),0,decode(nvl(ʵ������,0),0,0,S.ʵ�ʽ�� / S.ʵ������) ,���ۼ�) ʱ���ۼ�, " & _
    "                   s.ƽ���ɱ��� As �ɱ���" & _
    "           From ҩƷ��� S" & _
    "           Where S.ʵ������<>0 and S.����=1 and S.ҩƷid=[1] " & IIf(mlng��Ӧ��ID = 0, "", " And Nvl(S.�ϴι�Ӧ��ID,0)=[2]") & ") S, " & _
    "       ���ű� D,�շ���ĿĿ¼ M,�������� P,��Ӧ�� L" & _
    " where S.�ⷿid=D.id and S.ҩƷID=M.ID And M.ID=P.����ID and S.�ϴι�Ӧ��ID=L.ID(+)" & _
    " order by M.����,S.����"
    
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lng����ID, mlng��Ӧ��ID)
 
    With vsStoce
        lngRow = .Rows - 1
        If Val(.Cell(flexcpData, lngRow, .ColIndex("������Ϣ"))) <> 0 Then lngRow = lngRow + 1
        If lngRow = 1 Then
            .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        Else
            .Rows = .Rows + rsTemp.RecordCount
        End If
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("�ⷿ")) = zlStr.Nvl(rsTemp!�ⷿ)
            .Cell(flexcpData, lngRow, .ColIndex("�ⷿ")) = zlStr.Nvl(rsTemp!�ⷿID)
            .TextMatrix(lngRow, .ColIndex("��Ӧ��")) = zlStr.Nvl(rsTemp!��Ӧ��):
            .Cell(flexcpData, lngRow, .ColIndex("��Ӧ��")) = zlStr.Nvl(rsTemp!�ϴι�Ӧ��id)
            .TextMatrix(lngRow, .ColIndex("������Ϣ")) = zlStr.Nvl(rsTemp!����)
            .Cell(flexcpData, lngRow, .ColIndex("������Ϣ")) = zlStr.Nvl(rsTemp!����ID)
            .TextMatrix(lngRow, .ColIndex("���|����")) = zlStr.Nvl(rsTemp!���) & IIf(IsNull(rsTemp!����), "", "|" & rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("��λ")) = IIf(mintUnit = 0, zlStr.Nvl(rsTemp!���㵥λ), zlStr.Nvl(rsTemp!��װ��λ))
            
            .Cell(flexcpData, lngRow, .ColIndex("��λ")) = zlStr.Nvl(rsTemp!����ϵ��)
            .TextMatrix(lngRow, .ColIndex("����")) = zlStr.Nvl(rsTemp!����)
            .Cell(flexcpData, lngRow, .ColIndex("����")) = zlStr.Nvl(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("����")) = Format(Val(zlStr.Nvl(rsTemp!����)) / Val(zlStr.Nvl(rsTemp!����ϵ��)), mFMT.FM_����)
            .Cell(flexcpData, lngRow, .ColIndex("����")) = zlStr.Nvl(rsTemp!����)
            
            '����ԭ��
            dblTemp = IIf(Val(zlStr.Nvl(rsTemp!���)) = 1, Val(zlStr.Nvl(rsTemp!ʱ���ۼ�)), dblԭ��) * Val(zlStr.Nvl(rsTemp!����ϵ��))
            .TextMatrix(lngRow, .ColIndex("ԭ��")) = Format(dblTemp, mFMT.FM_���ۼ�)
            .Cell(flexcpData, lngRow, .ColIndex("ԭ��")) = IIf(Val(zlStr.Nvl(rsTemp!���)) = 1, Val(zlStr.Nvl(rsTemp!ʱ���ۼ�)), dblԭ��)
            
            .TextMatrix(lngRow, .ColIndex("�ּ�")) = Format(dbl�ּ� * Val(zlStr.Nvl(rsTemp!����ϵ��)), mFMT.FM_���ۼ�)
            .Cell(flexcpData, lngRow, .ColIndex("�ּ�")) = dbl�ּ�
            .TextMatrix(lngRow, .ColIndex("������")) = Format(Val(zlStr.Nvl(rsTemp!����)) * (dbl�ּ� - Val(.Cell(flexcpData, lngRow, .ColIndex("ԭ��")))), mFMT.FM_���)
            .Cell(flexcpData, lngRow, .ColIndex("������")) = Val(zlStr.Nvl(rsTemp!����)) * (dbl�ּ� - Val(.Cell(flexcpData, lngRow, .ColIndex("ԭ��"))))
             
             dblԭ�ɱ��� = Val(zlStr.Nvl(rsTemp!�ɱ���))
            
            If mdbl�ӳ��� > 0 Then
                dbl�ӳ��� = Round(mdbl�ӳ��� / 100, 7)
            ElseIf dblԭ�ɱ��� > 0 Then
                dbl�ӳ��� = Round(dblԭ�� / dblԭ�ɱ��� - 1, 7)
            Else
                dbl�ӳ��� = Round(1 / (1 - rsTemp!����� / 100) - 1, 7)
            End If
            
            If 1 + dbl�ӳ��� = 0 Then
                dbl�ֳɱ��� = 0
            Else
                dbl�ֳɱ��� = dbl�ּ� / (1 + dbl�ӳ���)
            End If
            If dbl�ӳ��� = -1 Then dbl�ӳ��� = 0
            
            .TextMatrix(lngRow, .ColIndex("ԭ�ɱ���")) = Format(dblԭ�ɱ��� * Val(zlStr.Nvl(rsTemp!����ϵ��)), mFMT.FM_�ɱ���)
            .Cell(flexcpData, lngRow, .ColIndex("ԭ�ɱ���")) = dblԭ�ɱ���
            
            .TextMatrix(lngRow, .ColIndex("�ӳ���")) = Format(dbl�ӳ��� * 100, GFM_VBJCL)
            .Cell(flexcpData, lngRow, .ColIndex("�ӳ���")) = dbl�ӳ��� * 100
            
            
            .TextMatrix(lngRow, .ColIndex("�ֳɱ���")) = Format(dbl�ֳɱ��� * Val(zlStr.Nvl(rsTemp!����ϵ��)), mFMT.FM_�ɱ���)
            .Cell(flexcpData, lngRow, .ColIndex("�ֳɱ���")) = dbl�ֳɱ���
            
            .TextMatrix(lngRow, .ColIndex("��۵�����")) = Format((dblԭ�ɱ��� - dbl�ֳɱ���) * Val(zlStr.Nvl(rsTemp!����)), mFMT.FM_���)
            .Cell(flexcpData, lngRow, .ColIndex("��۵�����")) = (dblԭ�ɱ��� - dbl�ֳɱ���) * Val(zlStr.Nvl(rsTemp!����))
            .RowHidden(lngRow) = IIf(chk��ʾ���в���.Value = 1, False, True)
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If Val(.Cell(flexcpData, .Rows - 1, .ColIndex("������Ϣ"))) = 0 And .Rows - 1 <> 1 Then
            .Rows = .Rows - 1
        End If
        
        '����Ӧ���䶯���
        If (m���۷�ʽ = T_�ɱ��۵��� Or m���۷�ʽ = T_�ɱ����ۼ۵���) And bln���� = False Then
            Call RefreshPayData
        End If
    End With
    LoadStockData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Function RefreshPayData() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:���»�ȡӦ������䶯����
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-11-05 15:03:46
    '-----------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, dbl��Ʊ��� As Double
    Dim lng��Ӧ��ID As Long, lng����ID As Long, blnData As Boolean
    
    err = 0: On Error GoTo ErrHand:
    If chk�Զ�����.Value <> 1 Then RefreshPayData = True: Exit Function
    
    With vsPay
        .Rows = 2
        .Clear 1
         .Cell(flexcpData, 1, .ColIndex("��Ʊ���"), .Rows - 1, .ColIndex("��Ʊ���")) = ""
    End With
    
    With vsStoce
        For i = 1 To .Rows - 1
            lng��Ӧ��ID = Val(.Cell(flexcpData, i, .ColIndex("��Ӧ��")))
            lng����ID = Val(.Cell(flexcpData, i, .ColIndex("������Ϣ")))
            If lng��Ӧ��ID <> 0 And lng����ID <> 0 Then
                dbl��Ʊ��� = Val(.Cell(flexcpData, i, .ColIndex("��۵�����")))
                If dbl��Ʊ��� <> 0 Then
                    '������صĹ�Ӧ���Ƿ����
                    With vsPay
                        blnData = False
                        For j = 1 To .Rows - 1
                            If lng����ID = Val(.Cell(flexcpData, j, .ColIndex("������Ϣ"))) And _
                               lng��Ӧ��ID = Val(.Cell(flexcpData, j, .ColIndex("��Ӧ��"))) Then
                               .Cell(flexcpData, j, .ColIndex("��Ʊ���")) = Val(.Cell(flexcpData, j, .ColIndex("��Ʊ���"))) + dbl��Ʊ���
                                .TextMatrix(j, .ColIndex("��Ʊ���")) = Format(Val(.Cell(flexcpData, j, .ColIndex("��Ʊ���"))), mFMT.FM_���)
                               blnData = True
                               Exit For
                            End If
                        Next
                        If blnData = False Then
                            'û�д˹�Ӧ�̻����,�����Ҫ��������
                            If Val(.Cell(flexcpData, .Rows - 1, .ColIndex("��Ӧ��"))) <> 0 Then
                                .Rows = .Rows + 1
                            End If
                            .TextMatrix(.Rows - 1, .ColIndex("��Ӧ��")) = vsStoce.TextMatrix(i, vsStoce.ColIndex("��Ӧ��"))
                             .Cell(flexcpData, .Rows - 1, .ColIndex("��Ӧ��")) = vsStoce.Cell(flexcpData, i, vsStoce.ColIndex("��Ӧ��"))
                            .TextMatrix(.Rows - 1, .ColIndex("������Ϣ")) = vsStoce.TextMatrix(i, vsStoce.ColIndex("������Ϣ"))
                             .Cell(flexcpData, .Rows - 1, .ColIndex("������Ϣ")) = vsStoce.Cell(flexcpData, i, vsStoce.ColIndex("������Ϣ"))
                            .TextMatrix(.Rows - 1, .ColIndex("���|����")) = vsStoce.TextMatrix(i, vsStoce.ColIndex("���|����"))
                            .Cell(flexcpData, .Rows - 1, .ColIndex("��Ʊ���")) = dbl��Ʊ���
                            .TextMatrix(.Rows - 1, .ColIndex("��Ʊ���")) = Format(Val(.Cell(flexcpData, .Rows - 1, .ColIndex("��Ʊ���"))), mFMT.FM_���)
                        End If
                    End With
                End If
            End If
        Next
    End With
    
    RefreshPayData = True
    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function
Private Sub InitPage()
    '------------------------------------------------------------------------------
    '����:��ʼ��ҳ��ؼ�
    '����:
    '����:���˺�
    '����:2007/08/18
    '------------------------------------------------------------------------------
    Dim i As Long
    Dim objItem As TabControlItem
    
    Set objItem = tbPage.InsertItem(mPageNum.Page_������, "���䶯", picStoce.hwnd, 0)
    objItem.Tag = mPageNum.Page_������
    Set objItem = tbPage.InsertItem(mPageNum.Page_Ӧ������, "Ӧ���䶯", picPay.hwnd, 0)
    objItem.Tag = mPageNum.Page_Ӧ������
    
    With tbPage
        tbPage.Item(0).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = True
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub chkӦ��_Click()
    Call SetControlVisble
End Sub

Private Sub picPay_Resize()
    err = 0: On Error Resume Next
    With picPay
        vsPay.Width = .ScaleWidth
        vsPay.Left = .ScaleLeft
        chk�Զ�����.Move .ScaleLeft, .ScaleTop + 100
        vsPay.Top = chk�Զ�����.Top + chk�Զ�����.Height + 100
        vsPay.Height = .ScaleHeight - .Top
    End With
End Sub
Private Sub picStoce_Resize()
    err = 0: On Error Resume Next
    With picStoce
        vsStoce.Width = .ScaleWidth
        vsStoce.Left = .ScaleLeft
        chk����.Move .ScaleLeft + 100, .ScaleTop + 100
        chkӦ��.Move chk����.Left + chk����.Width + 100, .ScaleTop + 100
        If chkӦ��.Visible = False And chk����.Visible = False Then
            chk��ʾ���в���.Move chk����.Left, .ScaleTop + 100
        ElseIf chkӦ��.Visible = False And chk����.Visible = True Then
            chk��ʾ���в���.Move chkӦ��.Left, .ScaleTop + 100
        ElseIf chkӦ��.Visible = True And chk����.Visible = False Then
            chkӦ��.Left = chk����.Left
            chk��ʾ���в���.Move chkӦ��.Left + chkӦ��.Width + 100, .ScaleTop + 100
        Else
            chk��ʾ���в���.Move chkӦ��.Left + chkӦ��.Width + 100, .ScaleTop + 100
        End If
        
        vsStoce.Top = chk����.Top + chk����.Height + 100
        
        cmdPrintStoce.Top = .ScaleTop
        cmdPrintStoce.Left = .Width - cmdPrintStoce.Width - 15
        vsStoce.Height = .ScaleHeight - vsStoce.Top
    End With

End Sub

Private Sub vsStoce_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsStoce
        Select Case Col
        Case .ColIndex("ԭ�ɱ���"), .ColIndex("�ֳɱ���")
            .TextMatrix(Row, Col) = Format(Val(.TextMatrix(Row, Col)), mFMT.FM_�ɱ���)
            '������ص�ֵ
            Call AutoCalcStoce(Row, Col)
        Case .ColIndex("�ӳ���")
            .TextMatrix(Row, Col) = Format(Val(.TextMatrix(Row, Col)), GFM_VBJCL)
            '������ص�ֵ
            Call AutoCalcStoce(Row, Col)
        Case .ColIndex("Ʒ��")
            .ColComboList(Col) = "..."
        Case .ColIndex("��������")
            .ColComboList(Col) = "..."
        End Select
    End With
End Sub

Private Sub vsStoce_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
'    zl_VsGridRowChange vsStoce, OldRow, NewRow, OldCol, NewCol
End Sub

Private Sub vsStoce_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mBillType = B_���� Then Cancel = True: Exit Sub
    If m���۷�ʽ = T_�ۼ۵��� Then Cancel = True: Exit Sub
    
    With vsStoce
        Select Case Col
        Case .ColIndex("�ֳɱ���"), .ColIndex("�ӳ���")
             If Val(.Cell(flexcpData, Row, .ColIndex("������Ϣ"))) = 0 Then Cancel = True
        Case Else
            Cancel = True
        End Select
    End With
    
End Sub


Private Sub vsStoce_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    With vsPrice
        Select Case Col
        Case .ColIndex("������Ϣ")
            '����
        Case Else
        End Select
    End With
End Sub

Private Sub vsStoce_CellChanged(ByVal Row As Long, ByVal Col As Long)
    mblnModify = True
End Sub
Private Sub vsStoce_EnterCell()
    If mBillType = B_���� Then Exit Sub
    With vsStoce
        Select Case .Col
        Case .ColIndex("������Ϣ")
        End Select
    End With
End Sub

Private Sub vsStoce_GotFocus()
'        zl_VsGridGotFocus vsStoce
End Sub

Private Sub vsStoce_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long, lngRow As Long
    
    With vsStoce
        If (.Col = .ColIndex("������Ϣ")) And KeyCode <> vbKeyReturn Then
            .ColComboList(.Col) = ""
        End If
        
        If KeyCode <> vbKeyReturn Then Exit Sub
        
        If Val(.Cell(flexcpData, .Row, .ColIndex("������Ϣ"))) = 0 Then
            OS.PressKey vbKeyTab
            Exit Sub
        End If
        Call zlVsMoveGridCell(vsStoce, , , False, lngRow)
    End With
End Sub

Private Sub vsStoce_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim intCol As Integer
    Dim strKey As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsStoce
        Select Case Col
        Case .ColIndex("�ֳɱ���")
            strKey = Trim(vsStoce.EditText)
            strKey = Replace(strKey, Chr(vbKeyReturn), "")
            strKey = Replace(strKey, Chr(10), "")
             
        Case Else
        End Select
    End With
    Call zlVsMoveGridCell(vsStoce, , , False)
End Sub

Private Sub vsStoce_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub vsStoce_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0: Exit Sub
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Exit Sub
    With vsStoce
        Select Case Col
        Case .ColIndex("������Ϣ")
            Call VsFlxGridCheckKeyPress(vsStoce, Row, Col, KeyAscii, m�ı�ʽ)
        Case .ColIndex("�ֳɱ���")
            Call VsFlxGridCheckKeyPress(vsStoce, Row, Col, KeyAscii, m���ʽ)
        Case Else
        End Select
    End With
End Sub

Private Sub vsStoce_LostFocus()
'    zl_VsGridLOSTFOCUS vsStoce
End Sub

Private Sub vsStoce_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKey As String
    Dim intCol As Integer
    Dim strTemp As String
    If mBillType = B_���� Then Cancel = True: Exit Sub
    
    strKey = Trim(vsStoce.EditText)
    strKey = Replace(strKey, Chr(vbKeyReturn), "")
    strKey = Replace(strKey, Chr(10), "")
    With vsStoce
        Select Case Col
        Case .ColIndex("�ֳɱ���")
            If strKey <> "" Then
                If zlCommFun.DblIsValid(strKey, 12, , False, , "��ָ������") = False Then Cancel = True: Exit Sub
                If Val(.Cell(flexcpData, .Row, .ColIndex("������Ϣ"))) <> 0 Then
                    If Check�ɱ���(Val(.Cell(flexcpData, .Row, .ColIndex("������Ϣ"))), Val(strKey)) = False Then
                        Cancel = True: Exit Sub
                    End If
                End If
                vsStoce.EditText = Format(Val(strKey), mFMT.FM_�ɱ���)
            End If
        Case .ColIndex("�ӳ���")
            If strKey <> "" Then
                If zlCommFun.DblIsValid(strKey, 5, , False, , "�ӳ���") = False Then Cancel = True: Exit Sub
                vsStoce.EditText = Format(Val(strKey), GFM_VBJCL)
            End If
        End Select
    End With
End Sub


'*****************************************************************************************************************
'**Ӧ���䶯����
Private Sub vsPay_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsPay
        Select Case Col
        Case .ColIndex("��Ʊ���")
            .TextMatrix(Row, Col) = Format(Val(.TextMatrix(Row, Col)), mFMT.FM_���)
        Case .ColIndex("��Ʊ����")
            .ColComboList(.Col) = "..."
        End Select
    End With
End Sub

Private Sub vsPay_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If mBillType = B_���� Then Cancel = True: Exit Sub
    If m���۷�ʽ = T_�ۼ۵��� Then Cancel = True: Exit Sub
    
    With vsPay
        Select Case Col
        Case .ColIndex("��Ʊ���"), .ColIndex("��Ʊ��"), .ColIndex("��Ʊ����")
             If Val(.Cell(flexcpData, Row, .ColIndex("������Ϣ"))) = 0 Then Cancel = True
        Case Else
            Cancel = True
        End Select
    End With
    
End Sub
 
Private Sub vsPay_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    With vsPay
        Select Case Col
        Case .ColIndex("��Ʊ����")
            Call SelDate

        Case Else
        End Select
    End With
End Sub
Private Function SelDate() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:ѡ��Ʊ����
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-11-07 11:59:54
    '-----------------------------------------------------------------------------------------------------------
    Dim strDate As String, blnreturn As Boolean
    Dim sngX As Single, sngY As Single, lngH As Long
    strDate = vsPay.TextMatrix(vsPay.Row, vsPay.ColIndex("��Ʊ����"))
    lngH = vsPay.CellHeight
    Call CalcPosition(sngX, sngY, vsPay)
      
    blnreturn = frmDateSel.SelectDate(Me, sngX, sngY, lngH, strDate)
    If blnreturn = False Then Exit Function
    With vsPay
        .TextMatrix(.Row, .ColIndex("��Ʊ����")) = strDate
    End With
    SelDate = True
End Function
Private Sub vsPay_CellChanged(ByVal Row As Long, ByVal Col As Long)
    mblnModify = True
End Sub
Private Sub vsPay_EnterCell()
    If mBillType = B_���� Then Exit Sub
    With vsPay
        Select Case .Col
        Case .ColIndex("��Ʊ����")
            .ColComboList(.Col) = "..."
        End Select
    End With
End Sub

Private Sub vsPay_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long, lngRow As Long
    
    With vsPay
        If (.Col = .ColIndex("��Ʊ����")) And KeyCode <> vbKeyReturn And KeyCode <> Asc("*") And KeyCode <> vbKeySpace And KeyCode <> vbKeyShift Then
            If Shift = 1 And KeyCode = 56 Then
                vsPay_CellButtonClick .Row, .Col
            Else
                .ColComboList(.Col) = ""
            End If
        End If
        If KeyCode = vbKeyDelete Then
            If MsgBox("���Ƿ����Ҫɾ�����е�Ӧ���䶯��¼��?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
            If .Row = .Rows - 1 And .Row = 1 Then
                .Clear 1
                .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""

            Else
                .RemoveItem .Row
            End If
        End If
        If KeyCode <> vbKeyReturn Then Exit Sub
        
        If Val(.Cell(flexcpData, .Row, .ColIndex("������Ϣ"))) = 0 Then
            OS.PressKey vbKeyTab
            Exit Sub
        End If
        Call zlVsMoveGridCell(vsPay, , , False, lngRow)
    End With
End Sub

Private Sub vsPay_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim intCol As Integer
    Dim strKey As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsPay
        Select Case Col
        Case .ColIndex("��Ʊ���")
            strKey = Trim(vsPay.EditText)
            strKey = Replace(strKey, Chr(vbKeyReturn), "")
            strKey = Replace(strKey, Chr(10), "")
            .EditText = .TextMatrix(Row, Col)
        Case Else
        End Select
    End With
    Call zlVsMoveGridCell(vsPay, , , False)
End Sub
Private Sub vsPay_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub
Private Sub vsPay_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0: Exit Sub
    If KeyAscii = vbKeyReturn Then KeyAscii = 0: Exit Sub
    With vsPay
        Select Case Col
        Case .ColIndex("��Ʊ��"), .ColIndex("��Ʊ����")
            Call VsFlxGridCheckKeyPress(vsPay, Row, Col, KeyAscii, m�ı�ʽ)
        Case .ColIndex("��Ʊ���")
            Call VsFlxGridCheckKeyPress(vsPay, Row, Col, KeyAscii, m�����ʽ)
        Case Else
        End Select
    End With
End Sub

Private Sub vsPay_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKey As String
    Dim intCol As Integer
    Dim strTemp As String
    If mBillType = B_���� Then Cancel = True: Exit Sub
    
    strKey = Trim(vsPay.EditText)
    strKey = Replace(strKey, Chr(vbKeyReturn), "")
    strKey = Replace(strKey, Chr(10), "")
    
    With vsPay
        Select Case Col
        Case .ColIndex("��Ʊ���")
            If strKey <> "" Then
                If zlCommFun.DblIsValid(strKey, 12, , False, , "��Ʊ���") = False Then Cancel = True: Exit Sub
                vsPay.EditText = Format(Val(strKey), mFMT.FM_�ɱ���)
            End If
        Case .ColIndex("��Ʊ����")
            If strKey = "" Then Exit Sub
            strKey = zlCheckIsDate(strKey, "��Ʊ����")
            If strKey = "" Then Cancel = True: Exit Sub
            .EditText = strKey
        Case .ColIndex("��Ʊ��")
            If strKey = "" Then Exit Sub
            If zlCommFun.StrIsValid(strKey, 200, 0, "��Ʊ��") = False Then Cancel = True: Exit Sub
        End Select
    End With
End Sub
'*************************************************************************************************************************

Private Sub AutoCalcStoce(ByVal lngEditRow As Long, ByVal lngEditCol As Long)
    '-----------------------------------------------------------------------------------------------------------
    '����:�Զ����������Ϣ(���ݼӳ��ʼ����ֳɱ��ۼ����,�����ֳɱ��ۼ�����ӳ���)
    '���:lngEditRow-��ǰ�༭����
    '     lngEditCol-��ǰ�༭����
    '����:
    '����:
    '����:���˺�
    '����:2008-11-06 17:03:02
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, dbl�ֳɱ��� As Double, dbl�ӳ��� As Double, dbl�ɱ���� As Double, dbl��۵����� As Double
    Dim lng����ID As Long, bln�ⷿ���� As Boolean, lng��Ӧ��ID As Long, lngTemp As Long, i As Long
    Dim blnHaveData As Boolean, lngStep As Long, lngSteps As Long
    
    err = 0: On Error GoTo ErrHand:
    With vsStoce
        bln�ⷿ���� = chk����.Value = 1
        lngStep = IIf(bln�ⷿ����, lngEditRow, 1)
        lngSteps = IIf(bln�ⷿ����, lngEditRow, .Rows - 1)
        Select Case lngEditCol
        Case .ColIndex("�ӳ���")
            dbl�ӳ��� = Val(.TextMatrix(lngEditRow, lngEditCol)) / 100
            If dbl�ӳ��� = -1 Then dbl�ӳ��� = 0
            '�ֳɱ���=�����ۼ�/(1+�ӳ���)
            dbl�ֳɱ��� = Round(Val(.Cell(flexcpData, lngEditRow, .ColIndex("�ּ�"))) / (1 + dbl�ӳ���), 7)
            '��۵�����=(ԭ�ɱ���-�ֳɱ���)
            dbl�ɱ���� = (Val(.Cell(flexcpData, lngEditRow, .ColIndex("ԭ�ɱ���"))) - dbl�ֳɱ���)
        Case .ColIndex("�ֳɱ���")
            '��Ϊ���ڰ�װ�������⣬��ˣ�Ŀǰ����С��λ�������õ���
            dbl�ֳɱ��� = Val(.TextMatrix(lngEditRow, lngEditCol)) / Val(.Cell(flexcpData, lngEditRow, .ColIndex("��λ")))
            '�ӳ���=�����ۼ�/�ֳɱ���-1
            If dbl�ֳɱ��� <> 0 Then
                dbl�ӳ��� = Round(Val(.Cell(flexcpData, lngEditRow, .ColIndex("�ּ�"))) / dbl�ֳɱ��� - 1, 7)
            Else
                dbl�ӳ��� = 0
            End If
            '��۵�����=(�ֳɱ���-ԭ�ɱ���)
            dbl�ɱ���� = Round((Val(.Cell(flexcpData, lngEditRow, .ColIndex("ԭ�ɱ���"))) - dbl�ֳɱ���), 7)
        Case .ColIndex("��۵�����")
            Exit Sub
        Case .ColIndex("�ּ�")
            '�ּ۷����ı�ʱ,��Ҫ���¸��ݼӳ��ʼ�����ص��ֳɱ���
            dbl�ӳ��� = Round(Val(.TextMatrix(lngEditRow, .ColIndex("�ӳ���"))) / 100, 7)
            If dbl�ӳ��� = -1 Then dbl�ӳ��� = 0
            '�ֳɱ���=�����ۼ�/(1+�ӳ���)
            dbl�ֳɱ��� = Round(Val(.Cell(flexcpData, lngEditRow, .ColIndex("�ּ�"))) / (1 + dbl�ӳ���), 7)
            '��۵�����=(�ֳɱ���-ԭ�ɱ���)
            dbl�ɱ���� = (dbl�ֳɱ��� - Val(.Cell(flexcpData, lngEditRow, .ColIndex("ԭ�ɱ���"))))
            lngStep = lngEditRow
            lngSteps = lngEditRow
        Case Else
            Exit Sub
        End Select

        lng����ID = Val(.Cell(flexcpData, lngEditRow, .ColIndex("������Ϣ")))
        lng��Ӧ��ID = Val(.Cell(flexcpData, lngEditRow, .ColIndex("��Ӧ��")))
        Dim cllData As New Collection
        For lngRow = lngStep To lngSteps
            If lng����ID = Val(.Cell(flexcpData, lngRow, .ColIndex("������Ϣ"))) Then
                If dbl�ӳ��� = -1 Then dbl�ӳ��� = 0
                .TextMatrix(lngRow, .ColIndex("�ӳ���")) = Format(dbl�ӳ��� * 100, GFM_VBJCL)
                '�óɱ���������С��λΪ׼�ģ����Ҫ��С����ϵ��.
                .TextMatrix(lngRow, .ColIndex("�ֳɱ���")) = Format(dbl�ֳɱ��� * Val(.Cell(flexcpData, lngRow, .ColIndex("��λ"))), mFMT.FM_�ɱ���)
                dbl�ɱ���� = (Val(.Cell(flexcpData, lngRow, .ColIndex("ԭ�ɱ���"))) - dbl�ֳɱ���)
                 '��۵�����=(�ֳɱ���-ԭ�ɱ���)*����
                 dbl��۵����� = Round(dbl�ɱ���� * Val(.Cell(flexcpData, lngRow, .ColIndex("����"))), 7)
                .TextMatrix(lngRow, .ColIndex("��۵�����")) = Format(dbl��۵�����, mFMT.FM_���)
                .Cell(flexcpData, lngRow, .ColIndex("��۵�����")) = dbl��۵�����
                lngTemp = Val(.Cell(flexcpData, lngRow, .ColIndex("������Ϣ")))
                lng��Ӧ��ID = Val(.Cell(flexcpData, lngRow, .ColIndex("��Ӧ��")))
                
                If lng��Ӧ��ID <> 0 Then
                    err = 0: On Error Resume Next
                    cllData.Add Array(lngTemp, lng��Ӧ��ID, dbl��۵�����, .TextMatrix(lngRow, .ColIndex("��Ӧ��")), .TextMatrix(lngRow, .ColIndex("������Ϣ")), .TextMatrix(lngRow, .ColIndex("���|����"))), "K" & lng��Ӧ��ID & "_" & lngTemp
                    If err <> 0 Then
                        '�ۼƲ�۵�����
                        dbl��۵����� = Val(cllData("K" & lng��Ӧ��ID & "_" & lngTemp)(2)) + dbl��۵�����
                        cllData.Remove "K" & lng��Ӧ��ID & "_" & lngTemp
                         err = 0: On Error GoTo ErrHand:
                        cllData.Add Array(lngTemp, lng��Ӧ��ID, dbl��۵�����, .TextMatrix(lngRow, .ColIndex("��Ӧ��")), .TextMatrix(lngRow, .ColIndex("������Ϣ")), .TextMatrix(lngRow, .ColIndex("���|����"))), "K" & lng��Ӧ��ID & "_" & lngTemp
                       
                    End If
                    On Error GoTo ErrHand:
                End If
            End If
        Next
        If chk�Զ�����.Value = 1 Then
            '��Ҫ�Զ�������ص�Ӧ���䶯��¼
            For i = 1 To cllData.Count
                With vsPay
                    blnHaveData = False
                    For lngRow = 1 To .Rows - 1
                        lngTemp = Val(.Cell(flexcpData, lngRow, .ColIndex("������Ϣ")))
                        lng��Ӧ��ID = Val(.Cell(flexcpData, lngRow, .ColIndex("��Ӧ��")))
                        If lngTemp = Val(cllData(i)(0)) _
                            And lng��Ӧ��ID = Val(cllData(i)(1)) Then
                            '���ļ���Ӧ����ͬ,�����ص�ֵ
                            .TextMatrix(lngRow, .ColIndex("��Ʊ���")) = Format(Val(cllData(i)(2)), mFMT.FM_���)
                            .Cell(flexcpData, lngRow, .ColIndex("��Ʊ���")) = Val(cllData(i)(2))
                             blnHaveData = True
                        End If
                    Next
                    If blnHaveData = False Then
                        '��Ҫ���Ӹ��Ӧ�̵�����
                        If Val(.Cell(flexcpData, .Rows - 1, .ColIndex("������Ϣ"))) <> 0 Then
                            .Rows = .Rows + 1
                        End If
                        lngRow = .Rows - 1
                        .TextMatrix(lngRow, .ColIndex("��Ӧ��")) = cllData(i)(3)
                        .Cell(flexcpData, lngRow, .ColIndex("��Ӧ��")) = cllData(i)(1)
                        .TextMatrix(lngRow, .ColIndex("������Ϣ")) = cllData(i)(4)
                        .Cell(flexcpData, lngRow, .ColIndex("������Ϣ")) = cllData(i)(0)
                        .TextMatrix(lngRow, .ColIndex("���|����")) = cllData(i)(5)
                        .TextMatrix(lngRow, .ColIndex("��Ʊ���")) = Format(Val(cllData(i)(2)), mFMT.FM_���)
                        .Cell(flexcpData, lngRow, .ColIndex("��Ʊ���")) = Val(cllData(i)(2))
                    End If
                End With
            Next
        End If
    End With
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub AutoCalc���п��۸�()
    '-----------------------------------------------------------------------------------------------------------
    '����:�Զ��������п��ļ۸�
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long, dbl�ֳɱ��� As Double, dbl�ּ� As Double, dbl�ӳ��� As Double, dbl�ɱ���� As Double, dbl��۵����� As Double, dbl������ As Double
    Dim lng����ID As Long, bln�ⷿ���� As Boolean, lng��Ӧ��ID As Long, lngTemp As Long, i As Long
    Dim blnHaveData As Boolean, lngStep As Long, lngSteps As Long
    Dim intCol As Integer
    Dim cllData As New Collection
    
    err = 0: On Error GoTo ErrHand:
    
    '��Ϊ���ڰ�װ�������⣬��ˣ�Ŀǰ����С��λ�������õ���
    dbl�ֳɱ��� = Val(vsPrice.TextMatrix(vsPrice.Row, vsPrice.ColIndex("�ֳɱ���")))
    dbl�ּ� = Val(vsPrice.TextMatrix(vsPrice.Row, vsPrice.ColIndex("�ּ�")))
    
    With vsStoce
        For lngRow = 1 To .Rows - 1
            If vsPrice.Col = vsPrice.ColIndex("�ֳɱ���") Then
                .TextMatrix(lngRow, .ColIndex("�ֳɱ���")) = dbl�ֳɱ���
                '�ӳ���=�����ۼ�/�ֳɱ���-1
                If dbl�ֳɱ��� <> 0 Then
                    dbl�ӳ��� = Round(Val(.Cell(flexcpData, lngRow, .ColIndex("�ּ�"))) / dbl�ֳɱ��� - 1, 7)
                Else
                    dbl�ӳ��� = 0
                End If
                '��۵�����=(�ֳɱ���-ԭ�ɱ���)
                dbl�ɱ���� = Round((Val(.Cell(flexcpData, lngRow, .ColIndex("ԭ�ɱ���"))) - dbl�ֳɱ���), 7)
            ElseIf vsPrice.Col = vsPrice.ColIndex("�ּ�") Then
                .TextMatrix(lngRow, .ColIndex("�ּ�")) = dbl�ּ�
                '�ּ۷����ı�ʱ,��Ҫ���¸��ݼӳ��ʼ�����ص��ֳɱ���
                dbl�ӳ��� = Round(Val(.TextMatrix(lngRow, .ColIndex("�ӳ���"))) / 100, 7)
                If dbl�ӳ��� = -1 Then dbl�ӳ��� = 0
                '�ֳɱ���=�����ۼ�/(1+�ӳ���)
                dbl�ֳɱ��� = Round(dbl�ּ� / (1 + dbl�ӳ���), 7)
                '��۵�����=(�ֳɱ���-ԭ�ɱ���)
                dbl�ɱ���� = (dbl�ֳɱ��� - Val(.Cell(flexcpData, lngRow, .ColIndex("ԭ�ɱ���"))))
                
                '������=����*(�ּ�-ԭ��)
                dbl������ = (dbl�ּ� - Val(.Cell(flexcpData, lngRow, .ColIndex("ԭ��")))) * Val(.Cell(flexcpData, lngRow, .ColIndex("����")))
                .TextMatrix(lngRow, .ColIndex("������")) = Format(dbl������, mFMT.FM_���)
                .Cell(flexcpData, lngRow, .ColIndex("������")) = dbl������
            End If
            
            lng����ID = Val(.Cell(flexcpData, lngRow, .ColIndex("������Ϣ")))
            lng��Ӧ��ID = Val(.Cell(flexcpData, lngRow, .ColIndex("��Ӧ��")))
            
            If dbl�ӳ��� = -1 Then dbl�ӳ��� = 0
            .TextMatrix(lngRow, .ColIndex("�ӳ���")) = Format(dbl�ӳ��� * 100, GFM_VBJCL)
            dbl�ɱ���� = (Val(.Cell(flexcpData, lngRow, .ColIndex("ԭ�ɱ���"))) - dbl�ֳɱ���)
             '��۵�����=(�ֳɱ���-ԭ�ɱ���)*����
             dbl��۵����� = Round(dbl�ɱ���� * Val(.Cell(flexcpData, lngRow, .ColIndex("����"))), 7)
            .TextMatrix(lngRow, .ColIndex("��۵�����")) = Format(dbl��۵�����, mFMT.FM_���)
            .Cell(flexcpData, lngRow, .ColIndex("��۵�����")) = dbl��۵�����
            lngTemp = Val(.Cell(flexcpData, lngRow, .ColIndex("������Ϣ")))
            lng��Ӧ��ID = Val(.Cell(flexcpData, lngRow, .ColIndex("��Ӧ��")))
            
            If lng��Ӧ��ID <> 0 Then
                err = 0: On Error Resume Next
                cllData.Add Array(lngTemp, lng��Ӧ��ID, dbl��۵�����, .TextMatrix(lngRow, .ColIndex("��Ӧ��")), .TextMatrix(lngRow, .ColIndex("������Ϣ")), .TextMatrix(lngRow, .ColIndex("���|����"))), "K" & lng��Ӧ��ID & "_" & lngTemp
                If err <> 0 Then
                    '�ۼƲ�۵�����
                    dbl��۵����� = Val(cllData("K" & lng��Ӧ��ID & "_" & lngTemp)(2)) + dbl��۵�����
                    cllData.Remove "K" & lng��Ӧ��ID & "_" & lngTemp
                     err = 0: On Error GoTo ErrHand:
                    cllData.Add Array(lngTemp, lng��Ӧ��ID, dbl��۵�����, .TextMatrix(lngRow, .ColIndex("��Ӧ��")), .TextMatrix(lngRow, .ColIndex("������Ϣ")), .TextMatrix(lngRow, .ColIndex("���|����"))), "K" & lng��Ӧ��ID & "_" & lngTemp
                   
                End If
                On Error GoTo ErrHand:
            End If
        Next
        
        If chk�Զ�����.Value = 1 Then
            '��Ҫ�Զ�������ص�Ӧ���䶯��¼
            For i = 1 To cllData.Count
                With vsPay
                    blnHaveData = False
                    For lngRow = 1 To .Rows - 1
                        lngTemp = Val(.Cell(flexcpData, lngRow, .ColIndex("������Ϣ")))
                        lng��Ӧ��ID = Val(.Cell(flexcpData, lngRow, .ColIndex("��Ӧ��")))
                        If lngTemp = Val(cllData(i)(0)) _
                            And lng��Ӧ��ID = Val(cllData(i)(1)) Then
                            '���ļ���Ӧ����ͬ,�����ص�ֵ
                            .TextMatrix(lngRow, .ColIndex("��Ʊ���")) = Format(Val(cllData(i)(2)), mFMT.FM_���)
                            .Cell(flexcpData, lngRow, .ColIndex("��Ʊ���")) = Val(cllData(i)(2))
                             blnHaveData = True
                        End If
                    Next
                    If blnHaveData = False Then
                        '��Ҫ���Ӹ��Ӧ�̵�����
                        If Val(.Cell(flexcpData, .Rows - 1, .ColIndex("������Ϣ"))) <> 0 Then
                            .Rows = .Rows + 1
                        End If
                        lngRow = .Rows - 1
                        .TextMatrix(lngRow, .ColIndex("��Ӧ��")) = cllData(i)(3)
                        .Cell(flexcpData, lngRow, .ColIndex("��Ӧ��")) = cllData(i)(1)
                        .TextMatrix(lngRow, .ColIndex("������Ϣ")) = cllData(i)(4)
                        .Cell(flexcpData, lngRow, .ColIndex("������Ϣ")) = cllData(i)(0)
                        .TextMatrix(lngRow, .ColIndex("���|����")) = cllData(i)(5)
                        .TextMatrix(lngRow, .ColIndex("��Ʊ���")) = Format(Val(cllData(i)(2)), mFMT.FM_���)
                        .Cell(flexcpData, lngRow, .ColIndex("��Ʊ���")) = Val(cllData(i)(2))
                    End If
                End With
            Next
        End If
    End With
    
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Function IsValiedӦ����Ϣ() As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:���Ӧ����Ϣ�Ƿ����ȷ
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-11-10 10:28:17
    '-----------------------------------------------------------------------------------------------------------
    Dim str��Ʊ�� As String, str��Ʊ���� As String, dbl��Ʊ��� As Double, lngϵ�� As Long, j As Long
 
    
    IsValiedӦ����Ϣ = False
    If m���۷�ʽ = T_�ۼ۵��� Then IsValiedӦ����Ϣ = True: Exit Function
    If chkӦ��.Value <> 1 Then IsValiedӦ����Ϣ = True: Exit Function
    
    With vsPay
        For j = 1 To .Rows - 1
            If Val(.Cell(flexcpData, j, .ColIndex("������Ϣ"))) <> 0 Then
                '���Ƿ��д��������Ͽ��䶯���
                str��Ʊ�� = Trim(.TextMatrix(j, .ColIndex("��Ʊ��")))
                str��Ʊ���� = Trim(.TextMatrix(j, .ColIndex("��Ʊ����")))
                dbl��Ʊ��� = Val(.TextMatrix(j, .ColIndex("��Ʊ���")))
                If str��Ʊ���� <> "" Then
                    str��Ʊ���� = zlCheckIsDate(str��Ʊ����, "��Ʊ����")
                    If str��Ʊ���� = "" Then
                        tbPage.Item(1).Selected = True
                        .Row = j: .Col = .ColIndex("��Ʊ����")
                        zlControl.ControlSetFocus vsPay, True
                        Exit Function
                    End If
                Else
                    ShowMsgBox "�ڵ�" & j & "���еķ�Ʊ����δ���룬����!"
                    If tbPage.Item(1).Visible Then tbPage.Item(1).Selected = True
                    .Row = j: .Col = .ColIndex("��Ʊ����")
                    zlControl.ControlSetFocus vsPay, True
                    Exit Function
                End If
                
                If zlCommFun.StrIsValid(str��Ʊ��, 100, 0, "��Ʊ��") = False Then
                        If tbPage.Item(1).Visible Then tbPage.Item(1).Selected = True
                        .Row = j: .Col = .ColIndex("��Ʊ��")
                        zlControl.ControlSetFocus vsPay, True
                        Exit Function
                End If
                If str��Ʊ�� = "" Then
                    ShowMsgBox "�ڵ�" & j & "���еķ�Ʊ��δ���룬����!"
                   If tbPage.Item(1).Visible Then tbPage.Item(1).Selected = True
                    .Row = j: .Col = .ColIndex("��Ʊ��")
                    zlControl.ControlSetFocus vsPay, True
                    Exit Function
                End If
            End If
        Next
    End With
    IsValiedӦ����Ϣ = True
End Function
Private Function Check�ɱ���(ByVal lng����ID As Long, ByVal dbl�ɱ��� As Double, Optional ByRef dblOut�ɱ��� As Double) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:���ɱ����Ƿ���Ч
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-11-10 14:27:56
    '-----------------------------------------------------------------------------------------------------------
    Dim lngRow As Long
    dblOut�ɱ��� = dbl�ɱ���
    With vsPrice
        For lngRow = 1 To .Rows - 1
            If Val(.Cell(flexcpData, lngRow, .ColIndex("Ʒ��"))) = lng����ID Then
                If Val(.TextMatrix(lngRow, .ColIndex("��ָ���ۼ�"))) < dbl�ɱ��� Then
                    dblOut�ɱ��� = Val(.TextMatrix(lngRow, .ColIndex("��ָ���ۼ�")))
                    If MsgBox("ע�⣺" & vbCrLf & "    �������ϡ�" & .TextMatrix(lngRow, .ColIndex("Ʒ��")) & "��" & vbCrLf & _
                        "�ĳɱ���(" & Format(dbl�ɱ���, mFMT.FM_�ɱ���) & ")������ָ�����ۼ�(" & Format(Val(.TextMatrix(lngRow, .ColIndex("��ָ���ۼ�"))), mFMT.FM_���ۼ�) & ")" & _
                        "���Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
                        Check�ɱ��� = True
                        Exit Function
                    Else
                        Check�ɱ��� = False
                        Exit Function
                    End If
                Else
                    Check�ɱ��� = True: Exit Function
                End If
            End If
        Next
    End With
    'δ�ҵ���صĵ�����Ϣ��Ҳ����true
    Check�ɱ��� = True
End Function





