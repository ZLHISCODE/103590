VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmSetExpence 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   5430
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8325
   ControlBox      =   0   'False
   Icon            =   "frmSetExpence.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5430
   ScaleWidth      =   8325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin TabDlg.SSTab stab 
      Height          =   5175
      Left            =   105
      TabIndex        =   33
      Top             =   120
      Width           =   6735
      _ExtentX        =   11880
      _ExtentY        =   9128
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   5
      TabHeight       =   520
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      TabCaption(0)   =   "���ʲ���"
      TabPicture(0)   =   "frmSetExpence.frx":000C
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraTY"
      Tab(0).Control(1)=   "txtת��"
      Tab(0).Control(2)=   "txtOutDay0"
      Tab(0).Control(3)=   "fraDoctor"
      Tab(0).Control(4)=   "lst�շ����"
      Tab(0).Control(5)=   "UDOutDay(0)"
      Tab(0).Control(6)=   "chkת��"
      Tab(0).Control(7)=   "fraҩ��"
      Tab(0).Control(8)=   "lblOutDate(0)"
      Tab(0).Control(9)=   "Label1"
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "���ʲ���(&1)"
      TabPicture(1)   =   "frmSetExpence.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "chkȱʡ���"
      Tab(1).Control(1)=   "chkRefundStyle"
      Tab(1).Control(2)=   "chk(14)"
      Tab(1).Control(3)=   "UDOutDay(1)"
      Tab(1).Control(4)=   "txtOutDay1"
      Tab(1).Control(5)=   "lblOutDate(1)"
      Tab(1).ControlCount=   6
      TabCaption(2)   =   "����Ʊ�ݿ���(&2)"
      TabPicture(2)   =   "frmSetExpence.frx":0044
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "lblInUse"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lblOutUse"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "fraTitle"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "cmdPrintSetup"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "cmdListPrintSet"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "cmd�˿��վ�"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "cmdBillZY"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "cboInvoiceKindZY"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "fraƱ�ݸ�ʽ"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "cmdOwnFee"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "cboInvoiceKindMZ"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "cmdBillMZ"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).ControlCount=   12
      TabCaption(3)   =   "����Ʊ�ݿ���(&3)"
      TabPicture(3)   =   "frmSetExpence.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "cmdRed"
      Tab(3).Control(1)=   "fraRed"
      Tab(3).Control(2)=   "cmdPrepayPrintSet"
      Tab(3).Control(3)=   "fraPrepay"
      Tab(3).ControlCount=   4
      Begin VB.CommandButton cmdBillMZ 
         Caption         =   "����Ʊ������(&P)"
         Height          =   350
         Left            =   5055
         TabIndex        =   53
         TabStop         =   0   'False
         Top             =   1920
         Width           =   1560
      End
      Begin VB.ComboBox cboInvoiceKindMZ 
         Height          =   300
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   1950
         Width           =   3270
      End
      Begin VB.CheckBox chkȱʡ��� 
         Caption         =   "�����˿�ʱѡ���ֽ����ȱʡ�˿���"
         Height          =   255
         Left            =   -74655
         TabIndex        =   26
         Top             =   1635
         Width           =   3525
      End
      Begin VB.CommandButton cmdRed 
         Caption         =   "���ʺ�Ʊ��ӡ����(&S)"
         Height          =   350
         Left            =   -71235
         TabIndex        =   50
         Top             =   4170
         Width           =   1965
      End
      Begin VB.Frame fraRed 
         Caption         =   "���Ϻ�Ʊ��ʽ"
         Height          =   1515
         Left            =   -74970
         TabIndex        =   48
         Top             =   2550
         Width           =   6600
         Begin VSFlex8Ctl.VSFlexGrid vsRedFormat 
            Height          =   1155
            Left            =   60
            TabIndex        =   49
            Top             =   225
            Width           =   6375
            _cx             =   11245
            _cy             =   2037
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
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   8421504
            GridColorFixed  =   8421504
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   2
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmSetExpence.frx":007C
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
            ExplorerBar     =   2
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
      Begin VB.CommandButton cmdPrepayPrintSet 
         Caption         =   "Ԥ��Ʊ�ݴ�ӡ����(&S)"
         Height          =   350
         Left            =   -74190
         TabIndex        =   43
         Top             =   4170
         Width           =   1965
      End
      Begin VB.CommandButton cmdOwnFee 
         Caption         =   "�Է��嵥����(&4)"
         Height          =   350
         Left            =   165
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   4740
         Width           =   1860
      End
      Begin VB.Frame fraƱ�ݸ�ʽ 
         Caption         =   "�շ�Ʊ�ݸ�ʽ"
         Height          =   1620
         Left            =   120
         TabIndex        =   32
         Top             =   2655
         Width           =   6540
         Begin VSFlex8Ctl.VSFlexGrid vsBillFormat 
            Height          =   1320
            Left            =   60
            TabIndex        =   39
            Top             =   225
            Width           =   6330
            _cx             =   11165
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
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   8421504
            GridColorFixed  =   8421504
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmSetExpence.frx":010E
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
            ExplorerBar     =   2
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
      Begin VB.ComboBox cboInvoiceKindZY 
         Height          =   300
         Left            =   1710
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   2325
         Width           =   3270
      End
      Begin VB.Frame fraTY 
         Height          =   1170
         Left            =   -74760
         TabIndex        =   0
         Top             =   690
         Width           =   2400
         Begin VB.CheckBox chk 
            Caption         =   "סԺ���۲��˼���"
            Height          =   195
            Index           =   5
            Left            =   195
            TabIndex        =   3
            Top             =   840
            Width           =   1740
         End
         Begin VB.CheckBox chk 
            Caption         =   "�������۲��˼���"
            Height          =   195
            Index           =   4
            Left            =   195
            TabIndex        =   2
            Top             =   570
            Width           =   1740
         End
         Begin VB.CheckBox chk 
            Caption         =   "�����˶���������"
            ForeColor       =   &H00000000&
            Height          =   195
            Index           =   3
            Left            =   180
            TabIndex        =   1
            Top             =   300
            Width           =   1740
         End
      End
      Begin VB.Frame fraPrepay 
         Caption         =   "���ع���Ԥ��Ʊ��"
         Height          =   1995
         Left            =   -74970
         TabIndex        =   41
         Top             =   450
         Width           =   6600
         Begin VSFlex8Ctl.VSFlexGrid vsPrepay 
            Height          =   1455
            Left            =   75
            TabIndex        =   42
            Top             =   270
            Width           =   6375
            _cx             =   11245
            _cy             =   2566
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
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   8421504
            GridColorFixed  =   8421504
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmSetExpence.frx":01B4
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
            ExplorerBar     =   2
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
      Begin VB.CommandButton cmdBillZY 
         Caption         =   "����Ʊ������(&P)"
         Height          =   350
         Left            =   5055
         TabIndex        =   29
         TabStop         =   0   'False
         Top             =   2295
         Width           =   1560
      End
      Begin VB.CommandButton cmd�˿��վ� 
         Caption         =   "�˿��վ�����(&3)"
         Height          =   350
         Left            =   4050
         TabIndex        =   38
         TabStop         =   0   'False
         Top             =   4305
         Width           =   1620
      End
      Begin VB.CommandButton cmdListPrintSet 
         Caption         =   "��ӡ������ϸ����(1)"
         Height          =   350
         Left            =   165
         TabIndex        =   37
         TabStop         =   0   'False
         Top             =   4305
         Width           =   1860
      End
      Begin VB.CommandButton cmdPrintSetup 
         Caption         =   "�ص�Ʊ�ݴ�ӡ����(&2)"
         Height          =   350
         Left            =   2115
         TabIndex        =   36
         TabStop         =   0   'False
         Top             =   4305
         Width           =   1860
      End
      Begin VB.Frame fraTitle 
         Caption         =   "���ع����շ�Ʊ��"
         Height          =   1470
         Left            =   90
         TabIndex        =   30
         Top             =   435
         Width           =   6540
         Begin VSFlex8Ctl.VSFlexGrid vsBill 
            Height          =   1125
            Left            =   75
            TabIndex        =   31
            Top             =   255
            Width           =   6330
            _cx             =   11165
            _cy             =   1984
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
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   8421504
            GridColorFixed  =   8421504
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   6
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmSetExpence.frx":0294
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
            ExplorerBar     =   2
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
      Begin VB.CheckBox chkRefundStyle 
         Caption         =   "�����˿�ȱʡ��Ԥ���ɿʽ"
         Height          =   255
         Left            =   -74655
         TabIndex        =   25
         Top             =   1335
         Width           =   3525
      End
      Begin VB.TextBox txtת�� 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   -73725
         MaxLength       =   2
         TabIndex        =   21
         Text            =   "3"
         Top             =   3690
         Width           =   255
      End
      Begin VB.TextBox txtOutDay0 
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   -73920
         MaxLength       =   3
         TabIndex        =   18
         Text            =   "0"
         ToolTipText     =   "����Ϊ 0 ��ʾֻ��ѡ����Ժ����"
         Top             =   3315
         Width           =   450
      End
      Begin VB.Frame fraDoctor 
         Caption         =   "��ʾ������"
         Height          =   1170
         Left            =   -72045
         TabIndex        =   4
         Top             =   720
         Width           =   1755
         Begin VB.OptionButton optDoctorKind 
            Caption         =   "������"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   0
            Left            =   210
            TabIndex        =   5
            Top             =   435
            Value           =   -1  'True
            Width           =   1020
         End
         Begin VB.OptionButton optDoctorKind 
            Caption         =   "������"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   1
            Left            =   210
            TabIndex        =   6
            Top             =   735
            Width           =   1020
         End
      End
      Begin VB.CheckBox chk 
         Caption         =   "LED��ʾ��ӭ��Ϣ"
         Height          =   225
         Index           =   14
         Left            =   -74655
         TabIndex        =   24
         ToolTipText     =   "�շѴ������벡�˺�,�Ƿ���ʾ��ӭ��Ϣ������"
         Top             =   1065
         Value           =   1  'Checked
         Width           =   1770
      End
      Begin MSComCtl2.UpDown UDOutDay 
         Height          =   270
         Index           =   1
         Left            =   -73410
         TabIndex        =   34
         Top             =   675
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   476
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtOutDay1"
         BuddyDispid     =   196637
         OrigLeft        =   1486
         OrigTop         =   3375
         OrigRight       =   1726
         OrigBottom      =   3645
         Max             =   999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtOutDay1 
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   -73860
         MaxLength       =   3
         TabIndex        =   23
         Text            =   "0"
         ToolTipText     =   "����Ϊ 0 ��ʾֻ��ѡ����Ժ����"
         Top             =   690
         Width           =   450
      End
      Begin VB.ListBox lst�շ���� 
         Height          =   3000
         Left            =   -70095
         Style           =   1  'Checkbox
         TabIndex        =   16
         ToolTipText     =   "�븴ѡ����ʹ�õ��շ����"
         Top             =   930
         Width           =   1545
      End
      Begin MSComCtl2.UpDown UDOutDay 
         Height          =   270
         Index           =   0
         Left            =   -73470
         TabIndex        =   19
         TabStop         =   0   'False
         Top             =   3315
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   476
         _Version        =   393216
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtOutDay0"
         BuddyDispid     =   196633
         OrigLeft        =   1486
         OrigTop         =   2760
         OrigRight       =   1726
         OrigBottom      =   3030
         Max             =   999
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.CheckBox chkת�� 
         Caption         =   "��ʾ���   ��ת���Ĳ���"
         Height          =   195
         Left            =   -74715
         TabIndex        =   20
         Top             =   3720
         Width           =   2370
      End
      Begin VB.Frame fraҩ�� 
         Caption         =   " ҩ���뷢�ϲ������� "
         Height          =   1185
         Left            =   -74745
         TabIndex        =   7
         Top             =   1965
         Width           =   4470
         Begin VB.ComboBox cbo���� 
            Height          =   300
            Left            =   2880
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   720
            Width           =   1305
         End
         Begin VB.ComboBox cbo��ҩ 
            Height          =   300
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   720
            Width           =   1305
         End
         Begin VB.ComboBox cbo��ҩ 
            Height          =   300
            Left            =   720
            Style           =   2  'Dropdown List
            TabIndex        =   9
            Top             =   360
            Width           =   1305
         End
         Begin VB.ComboBox cbo��ҩ 
            Height          =   300
            Left            =   2880
            Style           =   2  'Dropdown List
            TabIndex        =   11
            Top             =   360
            Width           =   1305
         End
         Begin VB.Label Label11 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���ϲ���"
            Height          =   180
            Left            =   2100
            TabIndex        =   14
            Top             =   780
            Width           =   720
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�в�ҩ"
            Height          =   180
            Left            =   120
            TabIndex        =   12
            Top             =   780
            Width           =   540
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ҩ"
            Height          =   180
            Left            =   120
            TabIndex        =   8
            Top             =   420
            Width           =   540
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�г�ҩ"
            Height          =   180
            Left            =   2280
            TabIndex        =   10
            Top             =   420
            Width           =   540
         End
      End
      Begin VB.Label lblOutUse 
         AutoSize        =   -1  'True
         Caption         =   "�������Ʊ��ʹ��"
         Height          =   180
         Left            =   150
         TabIndex        =   51
         Top             =   2010
         Width           =   1440
      End
      Begin VB.Label lblInUse 
         AutoSize        =   -1  'True
         Caption         =   "סԺ����Ʊ��ʹ��"
         Height          =   180
         Left            =   150
         TabIndex        =   27
         Top             =   2385
         Width           =   1440
      End
      Begin VB.Label lblOutDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ѡ��         ���ڳ�Ժ�Ĳ���"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   0
         Left            =   -74700
         TabIndex        =   17
         Top             =   3360
         Width           =   2790
      End
      Begin VB.Label lblOutDate 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����ѡ��         ���ڳ�Ժ�Ĳ���"
         ForeColor       =   &H00000000&
         Height          =   180
         Index           =   1
         Left            =   -74655
         TabIndex        =   22
         Top             =   750
         Width           =   2790
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�������:"
         Height          =   180
         Left            =   -70125
         TabIndex        =   35
         Top             =   660
         Width           =   810
      End
   End
   Begin VB.CommandButton cmdDeviceSetup 
      Caption         =   "�豸����(&S)"
      Height          =   350
      Left            =   7035
      TabIndex        =   46
      Top             =   1710
      Width           =   1110
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   7035
      TabIndex        =   44
      Top             =   645
      Width           =   1110
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   7035
      TabIndex        =   45
      Top             =   1185
      Width           =   1110
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   7035
      TabIndex        =   47
      Top             =   4275
      Width           =   1110
   End
End
Attribute VB_Name = "frmSetExpence"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Public mbytInFun As Byte '0=����,1=����
Public mbytUseType As Byte '0:��ͨ����,1-���ҷ�ɢ����,2-ҽ�����Ҽ���
Public mstrPrivs As String
Public mlngModul As Long
Public mblnOnlyDrugStock As Boolean  '����ʾҩ������
Private Enum chkBPS
    C0���� = 0
    C1���� = 1
    C2��� = 2
End Enum
Private Enum chks
    C03�����˶����� = 3
    C04�������ۼ��� = 4
    C05סԺ���ۼ��� = 5
    C09ҽ�����ʲ��� = 9
    C14LED��ӭ��Ϣ = 14
End Enum
Private Enum InvoiceKind
    C1�շ��վ� = 1
    C3�����վ� = 3
    C4�����վ� = 10
End Enum
Private Const CModule As Long = 1150    'סԺ���ʲ���
Private mstrOptPrivs As String


Private Sub zlOnlyDrugStrock()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ʾҩ�����������
    '����:���˺�
    '����:2010-01-25 15:24:21
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim ctl As Control
    Err = 0: On Error GoTo ErrHand:
    If Not (mblnOnlyDrugStock And mbytInFun = 0) Then Exit Sub
    
    For Each ctl In Me.Controls
       Select Case UCase(TypeName(ctl))
       Case UCase("ImageList")
       Case UCase("sstab")
            ctl.Visible = True
       Case Else
            If ctl Is fraҩ�� Or ctl.Container Is fraҩ�� Or ctl Is cmdOK Or ctl Is cmdCancel Then
                ctl.Visible = True
            Else
                 ctl.Visible = False
            End If
       End Select
    Next
    
    fraҩ��.Top = fraTY.Top
    Me.Height = 3525: Me.Width = 5470
    cmdCancel.Top = ScaleHeight - cmdCancel.Height - 100
    cmdCancel.Left = ScaleWidth - cmdCancel.Width - 100
    cmdOK.Top = cmdCancel.Top
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 100
    
    stab.Height = cmdOK.Top - stab.Top - 100
    stab.Width = ScaleWidth - stab.Left * 2
    stab.TabCaption(0) = "ҩ������"
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub cboInvoiceKindZY_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo��ҩ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

 
Private Sub cbo����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo��ҩ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

 
Private Sub cbo��ҩ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chk_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

 

Private Sub chkRefundStyle_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkȱʡ���_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

'����:27380
Private Sub chkת��_Click()
    txtת��.Enabled = chkת��.Value = 1
    If txtת��.Visible And txtת��.Enabled Then txtת��.SetFocus
End Sub
Private Sub chkת��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cmdListPrintSet_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1137_3", Me)
End Sub

Private Sub cmdOwnFee_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_INSIDE_1137_4", Me)
End Sub

Private Sub cmdPrintSetup_Click()
     Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1137_4", Me)
End Sub
Private Sub cmd�˿��վ�_Click()
    '���˺� ����:27776 ����:2010-02-04 16:44:39
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1137_3", Me)
End Sub

 
 
Private Sub lst�շ����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optDoctorKind_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtOutDay0_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtOutDay1_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub txtת��_GotFocus()
   zlControl.TxtSelAll txtת��
End Sub

Private Sub txtת��_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub


Private Sub cboInvoiceKindZY_Click()
    Dim bytKind As Byte
    If Visible Then '����ʱǿ�Ƶ���
        If cboInvoiceKindZY.ListIndex = 0 And cboInvoiceKindMZ.ListIndex = 0 Then
            bytKind = InvoiceKind.C3�����վ�
        ElseIf cboInvoiceKindZY.ListIndex = 1 And cboInvoiceKindMZ.ListIndex = 1 Then
            bytKind = InvoiceKind.C1�շ��վ�
        Else
            bytKind = InvoiceKind.C4�����վ�
        End If
        Call InitShareInvoice(bytKind)
        'Call SetShareInvoice(IIf(cboInvoiceKindZY.ListIndex = 0, InvoiceKind.C3�����վ�, InvoiceKind.C1�շ��վ�))
        'Call SetFactBillFormat
    End If
End Sub

Private Sub cboInvoiceKindMZ_Click()
    Dim bytKind As Byte
    If Visible Then '����ʱǿ�Ƶ���
        If cboInvoiceKindZY.ListIndex = 0 And cboInvoiceKindMZ.ListIndex = 0 Then
            bytKind = InvoiceKind.C3�����վ�
        ElseIf cboInvoiceKindZY.ListIndex = 1 And cboInvoiceKindMZ.ListIndex = 1 Then
            bytKind = InvoiceKind.C1�շ��վ�
        Else
            bytKind = InvoiceKind.C4�����վ�
        End If
        Call InitShareInvoice(bytKind)
        'Call SetShareInvoice(IIf(cboInvoiceKindZY.ListIndex = 0, InvoiceKind.C3�����վ�, InvoiceKind.C1�շ��վ�))
        'Call SetFactBillFormat
    End If
End Sub

Private Sub cmdBillZY_Click()
    If gblnBillPrint Then
        Call gobjBillPrint.zlConfigure
    Else
        Call ReportPrintSet(gcnOracle, glngSys, IIf(cboInvoiceKindZY.ListIndex = 0, "ZL" & glngSys \ 100 & "_BILL_1137", "ZL" & glngSys \ 100 & "_BILL_1137_2"), Me)
    End If
End Sub

Private Sub cmdBillMZ_Click()
    If gblnBillPrint Then
        Call gobjBillPrint.zlConfigure
    Else
        Call ReportPrintSet(gcnOracle, glngSys, IIf(cboInvoiceKindMZ.ListIndex = 0, "ZL" & glngSys \ 100 & "_BILL_1137", "ZL" & glngSys \ 100 & "_BILL_1137_2"), Me)
    End If
End Sub

Private Sub cmdRed_Click()
    Call ReportPrintSet(gcnOracle, glngSys, IIf(cboInvoiceKindZY.ListIndex = 0, "ZL" & glngSys \ 100 & "_BILL_1137_5", "ZL" & glngSys \ 100 & "_BILL_1137_6"), Me)
End Sub

Private Sub cmdCancel_Click()
    mblnOnlyDrugStock = False
    Unload Me
End Sub

Private Sub cmdDeviceSetup_Click()
    Call zlCommFun.DeviceSetup(Me, 100, 1137)
End Sub

Private Sub cmdHelp_Click()
    Select Case stab.Tab
        Case 0
            ShowHelp App.ProductName, Me.hWnd, "frmSetExpence1"
        Case 1
            ShowHelp App.ProductName, Me.hWnd, "frmSetExpence2"
    End Select
End Sub

Private Sub cmdPrepayPrintSet_Click()
    Call ReportPrintSet(gcnOracle, glngSys, "ZL" & glngSys \ 100 & "_BILL_1103", Me)
End Sub

Private Sub cmdOK_Click()
    Dim strValue As String, i As Long, lngShareID As Long
    Dim blnHavePrivs As Boolean, strTemp As String
    Dim blnBillOptSet As Boolean
    
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    
    If mbytInFun = 0 And cbo��ҩ.Visible Then
        If cbo��ҩ.ListIndex = -1 And cbo��ҩ.ListCount > 0 And cbo��ҩ.Enabled Then
            MsgBox "��ѡ����ҩ��.", vbInformation, gstrSysName
            stab.Tab = 0: cbo��ҩ.SetFocus: Exit Sub
        End If
        If cbo��ҩ.ListIndex = -1 And cbo��ҩ.ListCount > 0 And cbo��ҩ.Enabled Then
            MsgBox "��ѡ���ҩ��.", vbInformation, gstrSysName
            stab.Tab = 0: cbo��ҩ.SetFocus: Exit Sub
        End If
        If cbo��ҩ.ListIndex = -1 And cbo��ҩ.ListCount > 0 And cbo��ҩ.Enabled Then
            MsgBox "��ѡ����ҩ��.", vbInformation, gstrSysName
            stab.Tab = 0: cbo��ҩ.SetFocus: Exit Sub
        End If
        If cbo����.ListIndex = -1 And cbo����.ListCount > 0 And cbo����.Enabled Then
            MsgBox "��ѡ�����ķ��ϲ���.", vbInformation, gstrSysName
            stab.Tab = 0: cbo����.SetFocus: Exit Sub
        End If
    End If
    '�������ע����Ϣ
    '����ʹ���������ۼ���ʱ,����������ʾ��������Ƿ����������ü��ʿ���
    If mbytInFun = 0 And (mbytUseType = 0 Or mbytUseType = 1) And chk(chks.C04�������ۼ���).Value = 0 Then
        If Not CheckUnits Then
            MsgBox "����ʹ���������ۼ���ʱ,��û�п��Լ��ʵĿ���,�����޷������ã�", vbInformation, gstrSysName
            Exit Sub
        End If
    End If
    
    On Error Resume Next
    If mbytInFun = 0 Then
        blnBillOptSet = InStr(1, mstrOptPrivs, ";����ѡ������;") > 0
    
        'ҩ��
        zlDatabase.SetPara "ȱʡ��ҩ��", IIf(cbo��ҩ.ListIndex = 0, "0", cbo��ҩ.ItemData(cbo��ҩ.ListIndex)), glngSys, CModule, blnBillOptSet
        zlDatabase.SetPara "ȱʡ��ҩ��", IIf(cbo��ҩ.ListIndex = 0, "0", cbo��ҩ.ItemData(cbo��ҩ.ListIndex)), glngSys, CModule, blnBillOptSet
        zlDatabase.SetPara "ȱʡ��ҩ��", IIf(cbo��ҩ.ListIndex = 0, "0", cbo��ҩ.ItemData(cbo��ҩ.ListIndex)), glngSys, CModule, blnBillOptSet
        zlDatabase.SetPara "ȱʡ���ϲ���", IIf(cbo����.ListIndex = 0, "0", cbo����.ItemData(cbo����.ListIndex)), glngSys, CModule, blnBillOptSet
        If mblnOnlyDrugStock Then GoTo GoOver:
        
        '1150�Ĳ���
        '--------------------------------------------------------------------------------
        '�շ����
        For i = lst�շ����.ListCount - 1 To 0 Step -1
            If lst�շ����.Selected(i) Then strValue = strValue & "'" & Chr(lst�շ����.ItemData(i)) & "',"
        Next
        If strValue <> "" Then strValue = Left(strValue, Len(strValue) - 1)
        zlDatabase.SetPara "�շ����", strValue, glngSys, CModule, blnBillOptSet
    
           
        '���۲��˼���
        zlDatabase.SetPara "�������۲��˼���", chk(chks.C04�������ۼ���).Value, glngSys, CModule, blnBillOptSet
        zlDatabase.SetPara "סԺ���۲��˼���", chk(chks.C05סԺ���ۼ���).Value, glngSys, CModule, blnBillOptSet
        
        zlDatabase.SetPara "��Ժ��������", Val(txtOutDay0.Text), glngSys, CModule, blnBillOptSet
        zlDatabase.SetPara "��������ʾ��ʽ", IIf(optDoctorKind(0).Value, 1, 2), glngSys, CModule, blnBillOptSet
        zlDatabase.SetPara "����ҽ��", IIf(chk(chks.C03�����˶�����).Value = 1, 0, 1), glngSys, CModule, blnBillOptSet
        
        If mbytUseType = 1 Then
            '���˺� ����:27380 ����:2010-01-22 14:45:32
            zlDatabase.SetPara "���ת������", IIf(chkת��.Value = 1, "1", "0") & "|" & Val(txtת��.Text), glngSys, mlngModul, blnHavePrivs
        End If
    Else
        '���ع��ý���Ʊ��
        zlDatabase.SetPara "סԺ����Ʊ������", cboInvoiceKindZY.ListIndex, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "�������Ʊ������", cboInvoiceKindMZ.ListIndex, glngSys, mlngModul, blnHavePrivs
        Call SaveInvoice
        
'        lngShareID = 0
'        For i = 1 To lvwBill.ListItems.Count
'            If lvwBill.ListItems(i).Checked Then lngShareID = Val(Mid(lvwBill.ListItems(i).Key, 2))
'        Next
'        zlDatabase.SetPara "���ý���Ʊ������", lngShareID, glngSys, mlngModul, blnHavePrivs
        
        'LED�豸
        zlDatabase.SetPara "LED��ʾ��ӭ��Ϣ", chk(chks.C14LED��ӭ��Ϣ).Value, glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "��Ժ��������", Val(txtOutDay1.Text), glngSys, mlngModul, blnHavePrivs
        zlDatabase.SetPara "�����˿�ȱʡ��ʽ", chkRefundStyle.Value, glngSys, mlngModul, blnHavePrivs  '30036
        zlDatabase.SetPara "�˿��ֽ����ȱʡ���", chkȱʡ���.Value, glngSys, mlngModul, blnHavePrivs
    End If
GoOver:
    If mblnOnlyDrugStock Then
        Call zlInitҩ��
    Else
        Call InitLocPar(mlngModul)
    End If
    gblnOK = True
    mblnOnlyDrugStock = False
    Unload Me
End Sub

Private Sub Form_Activate()
    If stab.TabVisible(0) Then
        If chk(chks.C03�����˶�����).Visible And chk(chks.C03�����˶�����).Enabled Then chk(chks.C03�����˶�����).SetFocus
    Else
        If vsBill.Enabled And vsBill.Visible Then vsBill.SetFocus
    End If
End Sub


Private Sub Loadҩ��()
    Dim rsTmp As ADODB.Recordset
        
    On Error GoTo errH
    Set rsTmp = GetDepartments("'��ҩ��','��ҩ��','��ҩ��','���ϲ���'", "2,3")
        
    cbo��ҩ.AddItem "�˹�ѡ��"
    cbo��ҩ.AddItem "�˹�ѡ��"
    cbo��ҩ.AddItem "�˹�ѡ��"
    cbo����.AddItem "�˹�ѡ��"
    
    If Not rsTmp.EOF Then
        rsTmp.Filter = "��������='��ҩ��'"
        Do While Not rsTmp.EOF
            cbo��ҩ.AddItem rsTmp!����
            cbo��ҩ.ItemData(cbo��ҩ.ListCount - 1) = rsTmp!ID
            rsTmp.MoveNext
        Loop
        rsTmp.Filter = "��������='��ҩ��'"
        Do While Not rsTmp.EOF
            cbo��ҩ.AddItem rsTmp!����
            cbo��ҩ.ItemData(cbo��ҩ.ListCount - 1) = rsTmp!ID
            rsTmp.MoveNext
        Loop
        rsTmp.Filter = "��������='��ҩ��'"
        Do While Not rsTmp.EOF
            cbo��ҩ.AddItem rsTmp!����
            cbo��ҩ.ItemData(cbo��ҩ.ListCount - 1) = rsTmp!ID
            rsTmp.MoveNext
        Loop
        
        rsTmp.Filter = "��������='���ϲ���'"
        Do While Not rsTmp.EOF
            cbo����.AddItem rsTmp!����
            cbo����.ItemData(cbo����.ListCount - 1) = rsTmp!ID
                            
            rsTmp.MoveNext
        Loop
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Load()
    Dim rsTmp As New ADODB.Recordset, strSQL As String
    Dim i As Long, strValue As String, blnParSet As Boolean, blnBillOptSet As Boolean
    Dim strDefault As String
    Dim varData As Variant
    Dim bytKind As Byte
    
    gblnOK = False
    On Error GoTo errH
    blnParSet = InStr(1, mstrPrivs, ";��������;") > 0
    
    If mbytInFun = 0 Then
        mstrOptPrivs = ";" & GetInsidePrivs(Enum_Inside_Program.p���ʲ���) & ";"
        blnBillOptSet = InStr(1, mstrOptPrivs, "����ѡ������") > 0
        '����1150�Ĳ���
        '--------------------------------------------------------------------------------------
        
        '1150�Ĳ���
        '------------------------------------------------------------------
        '�շ����(�Һų���)
        strSQL = "Select ����,���� as ��� From �շ���Ŀ��� Where ����<>'1' Order by ���"
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        Do While Not rsTmp.EOF
            lst�շ����.AddItem rsTmp!���
            lst�շ����.ItemData(lst�շ����.NewIndex) = Asc(rsTmp!����)
            rsTmp.MoveNext
        Loop
        strValue = zlDatabase.GetPara("�շ����", glngSys, CModule, , Array(lst�շ����), blnBillOptSet)
        If strValue = "" Then
            For i = 0 To lst�շ����.ListCount - 1
                lst�շ����.Selected(i) = True
            Next
        Else
            For i = 0 To lst�շ����.ListCount - 1
                If InStr(strValue, Chr(lst�շ����.ItemData(i))) Then lst�շ����.Selected(i) = True
            Next
        End If
        If lst�շ����.ListCount > 0 Then lst�շ����.TopIndex = 0: lst�շ����.ListIndex = 0
        
        '���۲��˼���
        chk(chks.C04�������ۼ���).Value = IIf(zlDatabase.GetPara("�������۲��˼���", glngSys, CModule, , Array(chk(chks.C04�������ۼ���)), blnBillOptSet) = "1", 1, 0)
        chk(chks.C05סԺ���ۼ���).Value = IIf(zlDatabase.GetPara("סԺ���۲��˼���", glngSys, CModule, , Array(chk(chks.C05סԺ���ۼ���)), blnBillOptSet) = "1", 1, 0)
                      
        txtOutDay0.Text = Val(zlDatabase.GetPara("��Ժ��������", glngSys, CModule, 0, Array(txtOutDay0, lblOutDate(0), UDOutDay(0)), blnBillOptSet))
        If Val(zlDatabase.GetPara("��������ʾ��ʽ", glngSys, CModule, 0, Array(optDoctorKind(0), optDoctorKind(1)), blnBillOptSet)) = 1 Then
            optDoctorKind(0).Value = True
        Else
            optDoctorKind(1).Value = True
        End If
        
        
        chk(chks.C03�����˶�����).Value = IIf(zlDatabase.GetPara("����ҽ��", glngSys, CModule, , Array(chk(chks.C03�����˶�����)), blnBillOptSet) = "1", 0, 1)
        
                
        
       
        '--------------------------
        Call Loadҩ��
        
        strValue = zlDatabase.GetPara("ȱʡ��ҩ��", glngSys, CModule, , Array(cbo��ҩ), blnBillOptSet)
        If IsNumeric(strValue) Then Call zlControl.CboLocate(cbo��ҩ, strValue, True)
        If cbo��ҩ.ListIndex = -1 And Val(strValue) = 0 Then cbo��ҩ.ListIndex = 0
        
        strValue = zlDatabase.GetPara("ȱʡ��ҩ��", glngSys, CModule, , Array(cbo��ҩ), blnBillOptSet)
        If IsNumeric(strValue) Then Call zlControl.CboLocate(cbo��ҩ, strValue, True)
        If cbo��ҩ.ListIndex = -1 And Val(strValue) = 0 Then cbo��ҩ.ListIndex = 0
        
        strValue = zlDatabase.GetPara("ȱʡ��ҩ��", glngSys, CModule, , Array(cbo��ҩ), blnBillOptSet)
        If IsNumeric(strValue) Then Call zlControl.CboLocate(cbo��ҩ, strValue, True)
        If cbo��ҩ.ListIndex = -1 And Val(strValue) = 0 Then cbo��ҩ.ListIndex = 0
        
        strValue = zlDatabase.GetPara("ȱʡ���ϲ���", glngSys, CModule, , Array(cbo����), blnBillOptSet)
        If IsNumeric(strValue) Then Call zlControl.CboLocate(cbo����, strValue, True)
        If cbo����.ListIndex = -1 And Val(strValue) = 0 Then cbo����.ListIndex = 0
        
        
        chkת��.Visible = False: txtת��.Visible = False
        If mbytUseType = 1 Then
            '���˺� ����:27380 ����:2010-01-22 14:45:32
            chkת��.Visible = True: txtת��.Visible = True
            Dim strת�� As String
            'CModule
            strת�� = zlDatabase.GetPara("���ת������", glngSys, mlngModul, "0|3", Array(chkת��, txtת��), InStr(1, mstrPrivs, ";��������;") > 0)
            txtת��.Text = Val(Split(strת�� & "|", "|")(1))
            chkת��.Value = IIf(Val(Split(strת�� & "|", "|")(0)) = 1, 1, 0)
        End If
        
    ElseIf mbytInFun = 1 Then
        chkRefundStyle.Value = IIf(Val(zlDatabase.GetPara("�����˿�ȱʡ��ʽ", glngSys, mlngModul, , Array(chkRefundStyle), blnParSet)) = 1, 1, 0)
        chkȱʡ���.Value = IIf(Val(zlDatabase.GetPara("�˿��ֽ����ȱʡ���", glngSys, mlngModul, , Array(chkȱʡ���), blnParSet)) = 1, 1, 0)
        
        cboInvoiceKindZY.AddItem "סԺҽ�Ʒ��վ�"
        cboInvoiceKindZY.AddItem "����ҽ�Ʒ��վ�"
        i = Val(zlDatabase.GetPara("סԺ����Ʊ������", glngSys, mlngModul, 0, Array(cboInvoiceKindZY), blnParSet))
        If i <> 0 Then i = 1
        cboInvoiceKindZY.ListIndex = i
        
        cboInvoiceKindMZ.AddItem "סԺҽ�Ʒ��վ�"
        cboInvoiceKindMZ.AddItem "����ҽ�Ʒ��վ�"
        i = Val(zlDatabase.GetPara("�������Ʊ������", glngSys, mlngModul, 0, Array(cboInvoiceKindMZ), blnParSet))
        If i <> 0 Then i = 1
        cboInvoiceKindMZ.ListIndex = i
        
        If InStr(1, mstrPrivs, ";������ý���;") = 0 Then '�������������ý���ʱ,ֻ��ʹ��סԺҽ�Ʒ��վ�
            cboInvoiceKindZY.ListIndex = 0
            cboInvoiceKindZY.Enabled = False
            cboInvoiceKindMZ.Enabled = False
        End If
        
        If cboInvoiceKindZY.ListIndex = 0 And cboInvoiceKindMZ.ListIndex = 0 Then
            bytKind = InvoiceKind.C3�����վ�
        ElseIf cboInvoiceKindZY.ListIndex = 1 And cboInvoiceKindMZ.ListIndex = 1 Then
            bytKind = InvoiceKind.C1�շ��վ�
        Else
            bytKind = InvoiceKind.C4�����վ�
        End If
        Call InitShareInvoice(bytKind)
        'Call SetShareInvoice(IIf(cboInvoiceKindZY.ListIndex = 0, InvoiceKind.C3�����վ�, InvoiceKind.C1�շ��վ�))
        '����:35142
        'Call SetFactBillFormat '������ͨ��ҽ�����˽��ʷ�Ʊ��ʽ
        'LED�豸
        chk(chks.C14LED��ӭ��Ϣ).Value = IIf(zlDatabase.GetPara("LED��ʾ��ӭ��Ϣ", glngSys, mlngModul, "1", Array(chk(chks.C14LED��ӭ��Ϣ)), blnParSet) = "1", 1, 0)
        txtOutDay1.Text = Val(zlDatabase.GetPara("��Ժ��������", glngSys, mlngModul, 0, Array(txtOutDay1, lblOutDate(1), UDOutDay(1)), blnParSet))
    End If
    If mbytInFun = 0 Then
        stab.TabVisible(1) = False
        stab.TabVisible(2) = False
        stab.TabVisible(3) = False
        '����:27380
        txtת��.Visible = mbytUseType = 1 '���ҷ�ɢ����
        chkת��.Visible = mbytUseType = 1 '���ҷ�ɢ����

    ElseIf mbytInFun = 1 Then
        stab.TabVisible(0) = False
    End If
    Call zlOnlyDrugStrock
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

'
'Private Sub SetShareInvoice(ByVal bytKind As Byte)
'    Dim rstmp As New ADODB.Recordset, strSQL As String
'    Dim i As Long, lngShareID As Long
'    Dim objItem As ListItem
'
'    '��ȡ���ù��ý�������
'    Set rstmp = GetShareInvoiceGroupID(bytKind)
'    lngShareID = Val(zlDatabase.GetPara("���ý���Ʊ������", glngSys, mlngModul, 0, Array(lvwBill), InStr(1, mstrPrivs, ";��������;") > 0))
'    lvwBill.ListItems.Clear
'    For i = 1 To rstmp.RecordCount
'        Set objItem = lvwBill.ListItems.Add(, "_" & rstmp!ID, rstmp!������, , 1)
'        objItem.SubItems(1) = Format(rstmp!�Ǽ�ʱ��, "yyyy-MM-dd")
'        objItem.SubItems(2) = rstmp!��ʼ���� & "," & rstmp!��ֹ����
'        objItem.SubItems(3) = rstmp!ʣ������
'        If rstmp!ID = lngShareID Then
'            objItem.Checked = True
'            objItem.Selected = True
'            lngShareID = 0
'        End If
'        rstmp.MoveNext
'    Next
'    If lngShareID <> 0 Then zlDatabase.SetPara "���ý���Ʊ������", 0, glngSys, mlngModul, InStr(1, mstrPrivs, ";��������;") > 0
'
'    Exit Sub
'errH:
'    If ErrCenter() = 1 Then Resume
'    Call SaveErrLog
'End Sub

Private Sub Form_Unload(Cancel As Integer)
    mbytInFun = 0
    mbytUseType = 0
    zl_vsGrid_Para_Save mlngModul, vsPrepay, Me.Name, "����Ԥ��Ʊ���б�", False, False
End Sub

Private Sub vsPrepay_AfterMoveColumn(ByVal Col As Long, Position As Long)
   zl_vsGrid_Para_Save mlngModul, vsPrepay, Me.Name, "����Ԥ��Ʊ���б�", False, False
End Sub

Private Sub vsPrepay_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
   zl_vsGrid_Para_Save mlngModul, vsPrepay, Me.Name, "����Ԥ��Ʊ���б�", False, False
End Sub

Private Sub vsPrepay_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        Dim i As Long
        With vsPrepay
            Select Case Col
            Case .ColIndex("ѡ��")
                If Val(.TextMatrix(Row, Col)) <> 0 And Val(.RowData(Row)) <> 0 Then
                    For i = 1 To .Rows - 1
                        If Trim(.Cell(flexcpData, Row, .ColIndex("Ԥ������"))) = Trim(.Cell(flexcpData, i, .ColIndex("Ԥ������"))) _
                            And i <> Row Then
                            .TextMatrix(i, Col) = 0
                        End If
                    Next
                End If
            End Select
        End With
End Sub

Private Sub vsPrepay_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        With vsPrepay
            If Val(.Tag) = 1 Then
                If InStr(1, mstrPrivs, ";��������;") = 0 Then Cancel = True: Exit Sub
            End If
            Select Case Col
            Case .ColIndex("ѡ��")
                If Val(.RowData(Row)) = 0 Then Cancel = True
            Case Else
                Cancel = True
            End Select
        End With
End Sub

Private Sub lst�շ����_ItemCheck(Item As Integer)
    If lst�շ����.SelCount = 0 And Not lst�շ����.Selected(Item) Then
        lst�շ����.Selected(Item) = True
    End If
End Sub
'
'Private Sub lvwBill_ItemCheck(ByVal Item As MSComctlLib.ListItem)
'    Dim i As Long
'    For i = 1 To lvwBill.ListItems.Count
'        If lvwBill.ListItems(i).Key <> Item.Key Then lvwBill.ListItems(i).Checked = False
'    Next
'    Item.Selected = True
'End Sub

Private Sub txtOutDay0_GotFocus()
    zlControl.TxtSelAll txtOutDay0
End Sub

Private Sub txtOutDay0_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txtOutDay1_GotFocus()
    zlControl.TxtSelAll txtOutDay1
End Sub

Private Sub txtOutDay1_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub
Private Function CheckUnits() As Boolean
'���ܣ���鰴��������֮��,�Ƿ��п��ü����ٴ�����
'˵��������ʹ���������ۼ���֮��,������ʾ�����ٴ�����
    Dim rsTmp As ADODB.Recordset
    Dim i As Long, lng����ID As Long
    Dim strSQL As String
    
    On Error GoTo errH
    
    '��Ȩ����ʾ����۲��Ҷ�Ӧ���ٴ�����,סԺ������סԺ��ͬ
    If InStr(mstrPrivs, ";�������ۼ���;") And (chk(chks.C04�������ۼ���).Value = 1) Then
        strSQL = "1,2,3"
    Else
        strSQL = "2,3"
    End If
    If InStr(";" & mstrPrivs, ";���в���;") > 0 Then
        strSQL = _
             " Select Distinct A.ID,A.����,A.����" & _
             " From ���ű� A,��������˵�� B" & _
             " Where B.����ID = A.ID And B.������� IN(" & strSQL & ") And B.�������� IN('�ٴ�','����')" & _
             " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
             " Order by A.����"
    Else
        '����Ȩ�޵Ŀ��ң��������ڿ���+�������������Ŀ���
        '#������Ա��������۲���ʱ����ʹû���������ۼ��ʵ�Ȩ��,Ҳ��ʾ��Ӧ�������ٴ�����,���޷�����
        strSQL = _
            " Select A.ID,A.����,A.����,Nvl(C.ȱʡ,0) as ȱʡ" & _
            " From ���ű� A,��������˵�� B,������Ա C" & _
            " Where A.ID=B.����ID And A.ID=C.����ID And C.��ԱID=[1]" & _
            " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
            " And B.������� IN(" & strSQL & ") And B.�������� IN('�ٴ�','����')" & _
            " Order by A.����"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UserInfo.ID)
    CheckUnits = Not rsTmp.EOF
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
'
'Private Sub SetFactBillFormat()
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    '����:���÷�Ʊ��ʽ
'    '����:���˺�
'    '����:2010-12-31 19:29:48
'    '����:35142
'    '---------------------------------------------------------------------------------------------------------------------------------------------
'    Dim strRptName As String, rstmp As ADODB.Recordset, i As Long, blnParSet As Boolean, strSQL As String
'    blnParSet = zlStr.IsHavePrivs(mstrPrivs, ";��������;")
'    strRptName = IIf(cboInvoiceKindZY.ListIndex = 0, "ZL" & glngSys \ 100 & "_BILL_1137", "ZL" & glngSys \ 100 & "_BILL_1137_2")
'    cboFactNormal.Clear: cboFactMediCare.Clear
'
'    cboFactNormal.AddItem "ʹ�ñ���ȱʡ��ʽ"
'    cboFactMediCare.AddItem "ʹ�ñ���ȱʡ��ʽ"
'    '    Call ReportPrintSet(gcnOracle, glngSys, IIf(cboInvoiceKindZY.ListIndex = 0, "ZL" & glngSys \ 100 & "_BILL_1137", "ZL" & glngSys \ 100 & "_BILL_1137_2"), Me)
'    strSQL = "" & _
'    "   Select B.˵��,B.��� From zlReports A,zlRptFmts B" & _
'    "    Where A.ID=B.����ID And A.���=[1] " & _
'    "   Order by b.���"
'    Set rstmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strRptName)
'    For i = 1 To rstmp.RecordCount
'        cboFactNormal.AddItem rstmp!˵��
'        cboFactNormal.ItemData(cboFactNormal.NewIndex) = rstmp!���
'        cboFactMediCare.AddItem rstmp!˵��
'        cboFactMediCare.ItemData(cboFactMediCare.NewIndex) = rstmp!���
'        rstmp.MoveNext
'    Next
'    cboFactNormal.ListIndex = 0: cboFactMediCare.ListIndex = 0
'    i = Val(zlDatabase.GetPara("��ͨ��Ʊ��ʽ", glngSys, mlngModul, , Array(lblFactNormal, cboFactNormal), blnParSet))
'    Call zlControl.CboLocate(cboFactNormal, i, True)
'    i = Val(zlDatabase.GetPara("ҽ����Ʊ��ʽ", glngSys, mlngModul, , Array(lblFactMediCare, cboFactMediCare), blnParSet))
'    Call zlControl.CboLocate(cboFactMediCare, i, True)
'End Sub

Private Sub vsBill_AfterEdit(ByVal Row As Long, ByVal Col As Long)
        Dim i As Long
        With vsBill
            Select Case Col
            Case .ColIndex("ѡ��")
                If Val(.TextMatrix(Row, Col)) <> 0 And Val(.RowData(Row)) <> 0 Then
                    For i = 1 To .Rows - 1
                        If Trim(.TextMatrix(Row, .ColIndex("ʹ�����"))) = Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) _
                            And i <> Row Then
                            .TextMatrix(i, Col) = 0
                        End If
                    Next
                End If
            End Select
        End With
End Sub
 
Private Sub vsBill_AfterMoveColumn(ByVal Col As Long, Position As Long)
    zl_vsGrid_Para_Save mlngModul, vsBill, Me.Name, "����Ʊ��������", False, False
End Sub

Private Sub vsBill_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    zl_vsGrid_Para_Save mlngModul, vsBill, Me.Name, "����Ʊ��������", False, False
End Sub

Private Sub vsBill_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
        With vsBill
            If Val(.Tag) = 1 Then
                If InStr(1, mstrPrivs, ";��������;") = 0 Then Cancel = True: Exit Sub
            End If
            Select Case Col
            Case .ColIndex("ѡ��")
                If Val(.RowData(Row)) = 0 Then Cancel = True
            Case Else
                Cancel = True
            End Select
        End With
End Sub

Private Sub vsBillFormat_AfterMoveColumn(ByVal Col As Long, Position As Long)
        zl_vsGrid_Para_Save mlngModul, vsBillFormat, Me.Name, "����Ʊ�ݸ�ʽ", False, False
End Sub
Private Sub vsBillFormat_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
        zl_vsGrid_Para_Save mlngModul, vsBillFormat, Me.Name, "����Ʊ�ݸ�ʽ", False, False
End Sub

Private Sub vsBillFormat_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsBillFormat
        Select Case Col
        Case .ColIndex("סԺ����Ʊ�ݸ�ʽ")
            If Val(.ColData(Col)) = 1 And InStr(1, mstrPrivs, ";��������;") = 0 Then Cancel = True: Exit Sub
        Case .ColIndex("�������Ʊ�ݸ�ʽ")
            If Val(.ColData(Col)) = 1 And InStr(1, mstrPrivs, ";��������;") = 0 Then Cancel = True: Exit Sub
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub vsRedFormat_AfterMoveColumn(ByVal Col As Long, Position As Long)
        zl_vsGrid_Para_Save mlngModul, vsRedFormat, Me.Name, "���ʺ�Ʊ��ʽ", False, False
End Sub
Private Sub vsRedFormat_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
        zl_vsGrid_Para_Save mlngModul, vsRedFormat, Me.Name, "���ʺ�Ʊ��ʽ", False, False
End Sub

Private Sub vsRedFormat_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsRedFormat
        Select Case Col
        Case .ColIndex("Ʊ�ݸ�ʽ")
            If Val(.ColData(Col)) = 1 And InStr(1, mstrPrivs, ";��������;") = 0 Then Cancel = True: Exit Sub
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub SaveInvoice()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���淢Ʊ���Ʊ��
    '����:���˺�
    '����:2011-04-28 18:16:48
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim blnHavePrivs As Boolean, strValue As String
    Dim i As Long
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    '���湲��Ʊ��
    strValue = ""
    With vsBill
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("ѡ��"))) <> 0 And Val(.RowData(i)) <> 0 Then
                strValue = strValue & "|" & Val(.RowData(i)) & "," & Trim(.TextMatrix(i, .ColIndex("ʹ�����")))
            End If
        Next
    End With
    If strValue <> "" Then strValue = Mid(strValue, 2)
    zlDatabase.SetPara "���ý���Ʊ������", strValue, glngSys, mlngModul, blnHavePrivs
    
    '����Ԥ��Ʊ��
    strValue = ""
    With vsPrepay
        For i = 1 To .Rows - 1
            If Val(.TextMatrix(i, .ColIndex("ѡ��"))) <> 0 And Val(.RowData(i)) <> 0 Then
                strValue = strValue & "|" & Val(.RowData(i)) & "," & Trim(.Cell(flexcpData, i, .ColIndex("Ԥ������")))
            End If
        Next
    End With
    If strValue <> "" Then strValue = Mid(strValue, 2)
    zlDatabase.SetPara "����Ԥ��Ʊ������", strValue, glngSys, mlngModul, blnHavePrivs
    
    Dim strPrintMode As String
    '���������ʽ
    strValue = "": strPrintMode = ""
    With vsBillFormat
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) <> "" Then
                strValue = strValue & "|" & Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) & "," & Val(.TextMatrix(i, .ColIndex("�������Ʊ�ݸ�ʽ")))
            End If
        Next
        If strValue <> "" Then strValue = Mid(strValue, 2)
        zlDatabase.SetPara "������ʷ�Ʊ��ʽ", strValue, glngSys, mlngModul, blnHavePrivs
    End With
    
    '����סԺ��ʽ
    strValue = "": strPrintMode = ""
    With vsBillFormat
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) <> "" Then
                strValue = strValue & "|" & Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) & "," & Val(.TextMatrix(i, .ColIndex("סԺ����Ʊ�ݸ�ʽ")))
            End If
        Next
        If strValue <> "" Then strValue = Mid(strValue, 2)
        zlDatabase.SetPara "סԺ���ʷ�Ʊ��ʽ", strValue, glngSys, mlngModul, blnHavePrivs
    End With
    
    strValue = "": strPrintMode = ""
    With vsRedFormat
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) <> "" Then
                strValue = strValue & "|" & Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) & "," & Val(.TextMatrix(i, .ColIndex("Ʊ�ݸ�ʽ")))
            End If
        Next
        If strValue <> "" Then strValue = Mid(strValue, 2)
        zlDatabase.SetPara "���Ϸ�Ʊ��ʽ", strValue, glngSys, mlngModul, blnHavePrivs
    End With
End Sub
Private Function isValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������Ч�Լ��
    '����:���Ϸ�,����true,���򷵻�False
    '����:���˺�
    '����:2011-04-28 18:24:16
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, j As Long, lngSelCount As Long, str��� As String
     If mbytInFun <> 0 Then isValied = True: Exit Function
     
    isValied = False
    On Error GoTo errHandle
    '���ÿ��ʹ����ʽֻ��һ��ѡ��
    With vsBill
        str��� = "-"
        For i = 1 To vsBill.Rows - 1
            If str��� <> Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) Then
               str��� = Trim(.TextMatrix(i, .ColIndex("ʹ�����")))
               lngSelCount = 0
                For j = 1 To vsBill.Rows - 1
                    If Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) = Trim(.TextMatrix(j, .ColIndex("ʹ�����"))) Then
                        If Val(.TextMatrix(j, .ColIndex("ѡ��"))) <> 0 Then
                            lngSelCount = lngSelCount + 1
                        End If
                    End If
                Next
                If lngSelCount > 1 Then
                    MsgBox "ע��:" & vbCrLf & "    ʹ�����Ϊ��" & str��� & "����ֻ��ѡ��һ��Ʊ��,����!", vbInformation + vbOKOnly
                    Exit Function
                End If
            End If
        Next
    End With
    
      '���ÿ��ʹ��Ԥ��ֻ��һ��ѡ��
    With vsPrepay
        str��� = "-"
        For i = 1 To .Rows - 1
            If str��� <> Trim(.TextMatrix(i, .ColIndex("Ԥ������"))) Then
               str��� = Trim(.TextMatrix(i, .ColIndex("Ԥ������")))
               lngSelCount = 0
                For j = 1 To .Rows - 1
                    If Trim(.TextMatrix(i, .ColIndex("Ԥ������"))) = Trim(.TextMatrix(j, .ColIndex("Ԥ������"))) Then
                        If Val(.TextMatrix(j, .ColIndex("ѡ��"))) <> 0 Then
                            lngSelCount = lngSelCount + 1
                        End If
                    End If
                Next
                If lngSelCount > 1 Then
                    MsgBox "ע��:" & vbCrLf & "    Ԥ������Ϊ��" & str��� & "����ֻ��ѡ��һ��Ʊ��,����!", vbInformation + vbOKOnly
                    Exit Function
                End If
            End If
        Next
    End With
    

    isValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub InitShareInvoice(ByVal intKind As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ù���Ʊ
    '����:���˺�
    '     intKind:
    '����:2011-04-28 15:09:10
    '����:27559
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, lngRow As Long
    Dim strShareInvoice As String '����Ʊ������,��ʽ:����,����
    Dim varData As Variant, varTemp As Variant
    Dim varType As Variant, varTemp1 As Variant
    Dim intTYPE As Integer, intType1 As Integer, intType2 As Integer   '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    Dim lngTemp As Long, i As Long, strSQL As String
    Dim strRptName As String, blnHavePrivs As Boolean
    Dim strPrintMode As String, varDataMZ As Variant
    Dim str��Լ��λ���� As String, strShareInvoiceMZ As String
    
    blnHavePrivs = InStr(1, mstrPrivs, ";��������;") > 0
    
    On Error GoTo errHandle
    
    '�ָ��п��
    zl_vsGrid_Para_Restore mlngModul, vsBill, Me.Name, "����Ʊ��������", False, False
    zl_vsGrid_Para_Restore mlngModul, vsBillFormat, Me.Name, "����Ʊ�ݸ�ʽ", False, False
    zl_vsGrid_Para_Restore mlngModul, vsRedFormat, Me.Name, "���ʺ�Ʊ��ʽ", False, False
    strShareInvoice = zlDatabase.GetPara("���ý���Ʊ������", glngSys, mlngModul, , , True, intTYPE)
    '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    vsBill.Tag = ""
    Select Case intTYPE
    Case 1, 3, 5, 15
        vsBill.ForeColor = vbBlue: vsBill.ForeColorFixed = vbBlue
        fraTitle.ForeColor = vbBlue: vsBill.Tag = 1
        If intTYPE = 5 Then vsBill.Tag = ""
    Case Else
        vsBill.ForeColor = &H80000008: vsBill.ForeColorFixed = &H80000008
        fraTitle.ForeColor = &H80000008
    End Select
    With vsBill
        .Editable = flexEDKbdMouse
        If Val(.Tag) = 1 And Not blnHavePrivs Then .Editable = flexEDNone
    End With
    
    
    '��ʽ:����ID1,ʹ�����1|����IDn,ʹ�����n|...
    varData = Split(strShareInvoice, "|")
    '1.���ù���Ʊ��
    Set rsTemp = GetShareInvoiceGroupID(intKind)
    With vsBill
        .Clear 1: .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFixedOnly
        .MergeCol(0) = True
        Do While Not rsTemp.EOF
            .RowData(lngRow) = Val(NVL(rsTemp!ID))
            .TextMatrix(lngRow, .ColIndex("ʹ�����")) = NVL(rsTemp!ʹ�����, " ")
            .TextMatrix(lngRow, .ColIndex("������")) = NVL(rsTemp!������)
            .TextMatrix(lngRow, .ColIndex("��������")) = Format(rsTemp!�Ǽ�ʱ��, "yyyy-MM-dd")
            .TextMatrix(lngRow, .ColIndex("���뷶Χ")) = rsTemp!��ʼ���� & "," & rsTemp!��ֹ����
            .TextMatrix(lngRow, .ColIndex("ʣ��")) = Format(Val(NVL(rsTemp!ʣ������)), "##0;-##0;;")
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & ",", ",")
                lngTemp = Val(varTemp(0))
                If Val(.RowData(lngRow)) = lngTemp _
                    And varTemp(1) = Trim(.TextMatrix(lngRow, .ColIndex("ʹ�����"))) Then
                    .TextMatrix(lngRow, .ColIndex("ѡ��")) = -1: Exit For
                End If
            Next
            .MergeRow(lngRow) = True
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    
    strRptName = IIf(cboInvoiceKindZY.ListIndex = 0, "ZL" & glngSys \ 100 & "_BILL_1137", "ZL" & glngSys \ 100 & "_BILL_1137_2")
    'סԺƱ�ݸ�ʽ����
    strSQL = "" & _
    "   Select 'ʹ�ñ���ȱʡ��ʽ' as ˵��,0 as ���  From Dual Union ALL " & _
    "   Select B.˵��,B.���  " & _
    "   From zlReports A,zlRptFmts B" & _
    "   Where A.ID=B.����ID And A.���=[1]" & _
    "   Order by  ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strRptName)
    With vsBillFormat
        .Clear 1
        .ColComboList(.ColIndex("סԺ����Ʊ�ݸ�ʽ")) = .BuildComboList(rsTemp, "���,*˵��", "���")
    End With
    
    strRptName = IIf(cboInvoiceKindMZ.ListIndex = 0, "ZL" & glngSys \ 100 & "_BILL_1137", "ZL" & glngSys \ 100 & "_BILL_1137_2")
    '����Ʊ�ݸ�ʽ����
    strSQL = "" & _
    "   Select 'ʹ�ñ���ȱʡ��ʽ' as ˵��,0 as ���  From Dual Union ALL " & _
    "   Select B.˵��,B.���  " & _
    "   From zlReports A,zlRptFmts B" & _
    "   Where A.ID=B.����ID And A.���=[1]" & _
    "   Order by  ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strRptName)
    With vsBillFormat
        .ColComboList(.ColIndex("�������Ʊ�ݸ�ʽ")) = .BuildComboList(rsTemp, "���,*˵��", "���")
    End With
    
    '��ȡ����ֵ
    strShareInvoice = zlDatabase.GetPara("סԺ���ʷ�Ʊ��ʽ", glngSys, mlngModul, , , True, intTYPE)
    strShareInvoiceMZ = zlDatabase.GetPara("������ʷ�Ʊ��ʽ", glngSys, mlngModul, , , True, intTYPE)
    '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    With vsBillFormat
         .ColData(.ColIndex("סԺ����Ʊ�ݸ�ʽ")) = "0"
        .ForeColor = &H80000008:  .ForeColorFixed = &H80000008
        Select Case intTYPE
        Case 1, 3, 5, 15
             .ColData(.ColIndex("סԺ����Ʊ�ݸ�ʽ")) = IIf(intTYPE = 5, 0, 1)
        End Select
        If Val(.ColData(.ColIndex("סԺ����Ʊ�ݸ�ʽ"))) = 1 Then
            .Editable = IIf(Not blnHavePrivs, flexEDNone, flexEDKbdMouse)
        Else
            .Editable = flexEDKbdMouse
        End If
        '.ColComboList(.ColIndex("���ʺ��ӡ��ʽ")) = "0-����ӡƱ��|1-�Զ���ӡƱ��|2-ѡ���Ƿ��ӡƱ��"
    End With
    With vsBillFormat
         .ColData(.ColIndex("�������Ʊ�ݸ�ʽ")) = "0"
        .ForeColor = &H80000008:  .ForeColorFixed = &H80000008
        Select Case intTYPE
        Case 1, 3, 5, 15
             .ColData(.ColIndex("�������Ʊ�ݸ�ʽ")) = IIf(intTYPE = 5, 0, 1)
        End Select
        If Val(.ColData(.ColIndex("�������Ʊ�ݸ�ʽ"))) = 1 Then
            .Editable = IIf(Not blnHavePrivs, flexEDNone, flexEDKbdMouse)
        Else
            .Editable = flexEDKbdMouse
        End If
        '.ColComboList(.ColIndex("���ʺ��ӡ��ʽ")) = "0-����ӡƱ��|1-�Զ���ӡƱ��|2-ѡ���Ƿ��ӡƱ��"
    End With
    
    varData = Split(strShareInvoice, "|")
    varDataMZ = Split(strShareInvoiceMZ, "|")
    strSQL = "" & _
    "   Select ���� ,����" & _
    "   From  Ʊ��ʹ�����" & _
    "   Order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With vsBillFormat
        .Clear 1
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("ʹ�����")) = NVL(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("סԺ����Ʊ�ݸ�ʽ")) = "0"
            .TextMatrix(lngRow, .ColIndex("�������Ʊ�ݸ�ʽ")) = "0"
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(NVL(rsTemp!����)) Then
                    .TextMatrix(lngRow, .ColIndex("סԺ����Ʊ�ݸ�ʽ")) = Val(varTemp(1)): Exit For
                End If
            Next
            For i = 0 To UBound(varDataMZ)
                varTemp = Split(varDataMZ(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(NVL(rsTemp!����)) Then
                    .TextMatrix(lngRow, .ColIndex("�������Ʊ�ݸ�ʽ")) = Val(varTemp(1)): Exit For
                End If
            Next
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If Val(.ColData(.ColIndex("סԺ����Ʊ�ݸ�ʽ"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("סԺ����Ʊ�ݸ�ʽ"), .Rows - 1, .ColIndex("סԺ����Ʊ�ݸ�ʽ")) = vbBlue
        End If
        If Val(.ColData(.ColIndex("�������Ʊ�ݸ�ʽ"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("�������Ʊ�ݸ�ʽ"), .Rows - 1, .ColIndex("�������Ʊ�ݸ�ʽ")) = vbBlue
        End If
    End With
    
    strRptName = IIf(cboInvoiceKindZY.ListIndex = 0, "ZL" & glngSys \ 100 & "_BILL_1137_5", "ZL" & glngSys \ 100 & "_BILL_1137_6")
    'Ʊ�ݸ�ʽ����
    strSQL = "" & _
    "   Select 'ʹ�ñ���ȱʡ��ʽ' as ˵��,0 as ���  From Dual Union ALL " & _
    "   Select B.˵��,B.���  " & _
    "   From zlReports A,zlRptFmts B" & _
    "   Where A.ID=B.����ID And A.���=[1]" & _
    "   Order by  ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strRptName)
    With vsRedFormat
        .Clear 1
        .ColComboList(.ColIndex("Ʊ�ݸ�ʽ")) = .BuildComboList(rsTemp, "���,*˵��", "���")
    End With
    
    '��ȡ����ֵ
    strShareInvoice = zlDatabase.GetPara("���Ϸ�Ʊ��ʽ", glngSys, mlngModul, , , True, intTYPE)
    '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    With vsRedFormat
         .ColData(.ColIndex("Ʊ�ݸ�ʽ")) = "0"
        .ForeColor = &H80000008:  .ForeColorFixed = &H80000008
        Select Case intTYPE
        Case 1, 3, 5, 15
             .ColData(.ColIndex("Ʊ�ݸ�ʽ")) = IIf(intTYPE = 5, 0, 1)
        End Select
        If Val(.ColData(.ColIndex("Ʊ�ݸ�ʽ"))) = 1 Then
            .Editable = IIf(Not blnHavePrivs, flexEDNone, flexEDKbdMouse)
        Else
            .Editable = flexEDKbdMouse
        End If
        '.ColComboList(.ColIndex("���ʺ��ӡ��ʽ")) = "0-����ӡƱ��|1-�Զ���ӡƱ��|2-ѡ���Ƿ��ӡƱ��"
    End With
    
    varData = Split(strShareInvoice, "|")
    strSQL = "" & _
    "   Select ���� ,����" & _
    "   From  Ʊ��ʹ�����" & _
    "   order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With vsRedFormat
        .Clear 1
        .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("ʹ�����")) = NVL(rsTemp!����)
            .TextMatrix(lngRow, .ColIndex("Ʊ�ݸ�ʽ")) = "0"
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(NVL(rsTemp!����)) Then
                    .TextMatrix(lngRow, .ColIndex("Ʊ�ݸ�ʽ")) = Val(varTemp(1)): Exit For
                End If
            Next
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
        If Val(.ColData(.ColIndex("Ʊ�ݸ�ʽ"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("Ʊ�ݸ�ʽ"), .Rows - 1, .ColIndex("Ʊ�ݸ�ʽ")) = vbBlue
        End If
    End With
    
    '����Ԥ��Ʊ������
    '�ָ��п��
    zl_vsGrid_Para_Restore mlngModul, vsPrepay, Me.Name, "����Ԥ��Ʊ���б�", False, False
    
    strShareInvoice = zlDatabase.GetPara("����Ԥ��Ʊ������", glngSys, mlngModul, , , True, intTYPE)
    '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    vsBill.Tag = ""
    Select Case intTYPE
    Case 1, 3, 5, 15
        vsPrepay.ForeColor = vbBlue: vsPrepay.ForeColorFixed = vbBlue
        fraPrepay.ForeColor = vbBlue: vsBill.Tag = 1
        If intTYPE = 5 Then vsBill.Tag = ""
    Case Else
        vsPrepay.ForeColor = &H80000008: vsPrepay.ForeColorFixed = &H80000008
        fraPrepay.ForeColor = &H80000008
    End Select
    With vsPrepay
        .Editable = flexEDKbdMouse
        If Val(.Tag) = 1 And InStr(1, mstrPrivs, ";��������;") = 0 Then .Editable = flexEDNone
    End With
    
    '��ʽ:����ID1,Ԥ�����ID1|����IDn,Ԥ�����IDn|...
    varData = Split(strShareInvoice, "|")
    '1.���ù���Ʊ��
    Set rsTemp = GetShareInvoiceGroupID(2)
    With vsPrepay
        .Clear 1: .Rows = IIf(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount) + 1
        lngRow = 1
        .MergeCells = flexMergeRestrictRows
        .MergeCellsFixed = flexMergeFixedOnly
        .MergeCol(0) = True
        Do While Not rsTemp.EOF
            .RowData(lngRow) = Val(NVL(rsTemp!ID))
            If Val(NVL(rsTemp!ʹ�����, "")) = 0 Then
                .TextMatrix(lngRow, .ColIndex("Ԥ������")) = "�����סԺ����"
            ElseIf Val(NVL(rsTemp!ʹ�����, "")) = 1 Then
                .TextMatrix(lngRow, .ColIndex("Ԥ������")) = "Ԥ������Ʊ��"
            Else
                .TextMatrix(lngRow, .ColIndex("Ԥ������")) = "Ԥ��סԺƱ��"
            End If
            .Cell(flexcpData, lngRow, .ColIndex("Ԥ������")) = Val(NVL(rsTemp!ʹ�����))
            
            .TextMatrix(lngRow, .ColIndex("������")) = NVL(rsTemp!������)
            .TextMatrix(lngRow, .ColIndex("��������")) = Format(rsTemp!�Ǽ�ʱ��, "yyyy-MM-dd")
            .TextMatrix(lngRow, .ColIndex("���뷶Χ")) = rsTemp!��ʼ���� & "," & rsTemp!��ֹ����
            .TextMatrix(lngRow, .ColIndex("ʣ��")) = Format(Val(NVL(rsTemp!ʣ������)), "##0;-##0;;")
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & ",", ",")
                lngTemp = Val(varTemp(0))
                If Val(.RowData(lngRow)) = lngTemp _
                    And varTemp(1) = Val(.Cell(flexcpData, lngRow, .ColIndex("Ԥ������"))) Then
                    .TextMatrix(lngRow, .ColIndex("ѡ��")) = -1: Exit For
                End If
            Next
            .MergeRow(lngRow) = True
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub


