VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{945E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.DockingPane.9600.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTendFileStyle 
   Appearance      =   0  'Flat
   Caption         =   "�����ļ���ʽ"
   ClientHeight    =   7590
   ClientLeft      =   120
   ClientTop       =   405
   ClientWidth     =   12225
   Icon            =   "frmTendStyle.frx":0000
   LinkTopic       =   "Form1"
   MinButton       =   0   'False
   ScaleHeight     =   7590
   ScaleWidth      =   12225
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picTable 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   3405
      Left            =   960
      ScaleHeight     =   3405
      ScaleWidth      =   6885
      TabIndex        =   0
      Top             =   885
      Width           =   6885
      Begin VB.OptionButton optTabTiers 
         Caption         =   "��(&3)"
         Height          =   180
         Index           =   2
         Left            =   2790
         TabIndex        =   12
         Top             =   750
         Width           =   780
      End
      Begin VB.Frame fraSplit 
         Height          =   30
         Index           =   3
         Left            =   4380
         TabIndex        =   126
         Top             =   1905
         Width           =   2235
      End
      Begin VB.Frame fraSplit 
         Height          =   30
         Index           =   2
         Left            =   4380
         TabIndex        =   124
         Top             =   255
         Width           =   2235
      End
      Begin VB.Frame fraSplit 
         Height          =   30
         Index           =   1
         Left            =   1020
         TabIndex        =   121
         Top             =   2355
         Width           =   2235
      End
      Begin VB.Frame fraSplit 
         Height          =   30
         Index           =   0
         Left            =   1020
         TabIndex        =   120
         Top             =   225
         Width           =   2235
      End
      Begin VB.TextBox txtHeadText 
         Height          =   885
         Left            =   4980
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   41
         Top             =   2070
         Width           =   1710
      End
      Begin MSComCtl2.UpDown udHeadCol 
         Height          =   300
         Left            =   4335
         TabIndex        =   39
         Top             =   2400
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtHeadCol"
         BuddyDispid     =   196635
         OrigLeft        =   5985
         OrigTop         =   2085
         OrigRight       =   6225
         OrigBottom      =   2370
         Max             =   5
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtHeadCol 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4020
         MaxLength       =   2
         TabIndex        =   38
         Text            =   "1"
         Top             =   2430
         Width           =   330
      End
      Begin VB.TextBox txtHeadRow 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4035
         MaxLength       =   1
         TabIndex        =   35
         Text            =   "2"
         Top             =   2070
         Width           =   330
      End
      Begin MSComCtl2.UpDown udRecordTo 
         Height          =   285
         Left            =   5730
         TabIndex        =   28
         Top             =   465
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Value           =   8
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtRecordTo"
         BuddyDispid     =   196642
         OrigLeft        =   5985
         OrigTop         =   405
         OrigRight       =   6225
         OrigBottom      =   705
         Max             =   23
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udRecordFrom 
         Height          =   285
         Left            =   4365
         TabIndex        =   25
         Top             =   465
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Value           =   18
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtRecordFrom"
         BuddyDispid     =   196643
         OrigLeft        =   4440
         OrigTop         =   405
         OrigRight       =   4680
         OrigBottom      =   705
         Max             =   23
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udTabCols 
         Height          =   285
         Left            =   1530
         TabIndex        =   6
         Top             =   330
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Value           =   5
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtTabCols"
         BuddyDispid     =   196650
         OrigLeft        =   1530
         OrigTop         =   105
         OrigRight       =   1770
         OrigBottom      =   360
         Max             =   60
         Min             =   3
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.CommandButton cmdRecordFont 
         Caption         =   "��������(&N)"
         Height          =   300
         Left            =   3810
         TabIndex        =   30
         Top             =   840
         Width           =   1185
      End
      Begin VB.CommandButton cmdRecordColor 
         Caption         =   "������ɫ(&L)"
         Height          =   300
         Left            =   3810
         TabIndex        =   29
         Top             =   1200
         Width           =   1185
      End
      Begin VB.TextBox txtRecordTo 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5385
         MaxLength       =   2
         TabIndex        =   27
         Text            =   "8"
         Top             =   465
         Width           =   345
      End
      Begin VB.TextBox txtRecordFrom 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4020
         MaxLength       =   2
         TabIndex        =   24
         Text            =   "18"
         Top             =   465
         Width           =   585
      End
      Begin VB.CommandButton cmdTitleFont 
         Caption         =   "��������(&T)"
         Height          =   300
         Left            =   435
         TabIndex        =   20
         Top             =   2910
         Width           =   1185
      End
      Begin VB.TextBox txtTitleText 
         Height          =   300
         Left            =   435
         TabIndex        =   19
         Text            =   "�ر����¼��"
         Top             =   2535
         Width           =   2790
      End
      Begin VB.CommandButton cmdTabGridColor 
         Caption         =   "�����ɫ(&G)"
         Height          =   300
         Left            =   435
         TabIndex        =   17
         Top             =   1785
         Width           =   1185
      End
      Begin VB.CommandButton cmdTabTextColor 
         Caption         =   "�ı���ɫ(&R)"
         Height          =   300
         Left            =   435
         TabIndex        =   15
         Top             =   1425
         Width           =   1185
      End
      Begin VB.CommandButton cmdTabFont 
         Caption         =   "�ı�����(&F)"
         Height          =   300
         Left            =   435
         TabIndex        =   13
         Top             =   1065
         Width           =   1185
      End
      Begin VB.TextBox txtTabRowHeight 
         Height          =   300
         Left            =   2730
         MaxLength       =   3
         TabIndex        =   8
         Text            =   "300"
         Top             =   345
         Width           =   510
      End
      Begin VB.OptionButton optTabTiers 
         Caption         =   "˫(&2)"
         Height          =   180
         Index           =   1
         Left            =   1995
         TabIndex        =   11
         Top             =   750
         Width           =   780
      End
      Begin VB.OptionButton optTabTiers 
         Caption         =   "��(&1)"
         Height          =   180
         Index           =   0
         Left            =   1200
         TabIndex        =   10
         Top             =   750
         Value           =   -1  'True
         Width           =   780
      End
      Begin VB.TextBox txtTabCols 
         Enabled         =   0   'False
         Height          =   300
         Left            =   1230
         MaxLength       =   2
         TabIndex        =   5
         Text            =   "5"
         Top             =   345
         Width           =   300
      End
      Begin MSComCtl2.UpDown udHeadRow 
         Height          =   285
         Left            =   4335
         TabIndex        =   36
         Top             =   2070
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   503
         _Version        =   393216
         Value           =   2
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtHeadRow"
         BuddyDispid     =   196636
         OrigLeft        =   4920
         OrigTop         =   2085
         OrigRight       =   5160
         OrigBottom      =   2385
         Max             =   3
         Min             =   2
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Label lblTabHint 
         AutoSize        =   -1  'True
         Caption         =   "��ͬ���ݵ�ͬ�����ڵ�Ԫ�Զ��ϲ���"
         Height          =   180
         Index           =   1
         Left            =   3810
         TabIndex        =   125
         Top             =   3030
         Width           =   2880
      End
      Begin VB.Label lblTabHint 
         AutoSize        =   -1  'True
         Caption         =   "��¼����ʱ������ʱʹ�øø�ʽ��"
         Height          =   180
         Index           =   0
         Left            =   3810
         TabIndex        =   123
         Top             =   1605
         Width           =   2700
      End
      Begin VB.Label lblBasic 
         AutoSize        =   -1  'True
         Caption         =   "������̬"
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
         Left            =   225
         TabIndex        =   122
         Top             =   150
         Width           =   780
      End
      Begin VB.Label lblHeadText 
         AutoSize        =   -1  'True
         Caption         =   "�ı�"
         Height          =   180
         Left            =   4620
         TabIndex        =   40
         Top             =   2115
         Width           =   360
      End
      Begin VB.Label lblHeadCol 
         AutoSize        =   -1  'True
         Caption         =   "�к�"
         Height          =   180
         Left            =   3630
         TabIndex        =   37
         Top             =   2490
         Width           =   360
      End
      Begin VB.Label lblHeadRow 
         AutoSize        =   -1  'True
         Caption         =   "���"
         Height          =   180
         Left            =   3630
         TabIndex        =   34
         Top             =   2130
         Width           =   360
      End
      Begin VB.Label lblHeadSet 
         AutoSize        =   -1  'True
         Caption         =   "��ͷ��Ԫ"
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
         Left            =   3600
         TabIndex        =   33
         Top             =   1845
         Width           =   780
      End
      Begin VB.Label lblRecordFont 
         Caption         =   "����,9"
         Height          =   180
         Left            =   5010
         TabIndex        =   32
         Top             =   960
         Width           =   1605
      End
      Begin VB.Label lblRecordColor 
         Caption         =   "������ɫ"
         ForeColor       =   &H000000FF&
         Height          =   180
         Left            =   5010
         TabIndex        =   31
         Top             =   1320
         Width           =   1605
      End
      Begin VB.Label lblRecordTo 
         AutoSize        =   -1  'True
         Caption         =   "����       ��"
         Height          =   180
         Left            =   4995
         TabIndex        =   26
         Top             =   525
         Width           =   1170
      End
      Begin VB.Label lblRecordFrom 
         AutoSize        =   -1  'True
         Caption         =   "��       ����"
         Height          =   180
         Left            =   3810
         TabIndex        =   23
         Top             =   525
         Width           =   1170
      End
      Begin VB.Label lblRecordStyle 
         AutoSize        =   -1  'True
         Caption         =   "������ʽ"
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
         Left            =   3600
         TabIndex        =   22
         Top             =   180
         Width           =   780
      End
      Begin VB.Label lblTitleFont 
         Caption         =   "����,20"
         Height          =   180
         Left            =   1635
         TabIndex        =   21
         Top             =   3015
         Width           =   1605
      End
      Begin VB.Label lblTitleText 
         AutoSize        =   -1  'True
         Caption         =   "�����ı�"
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
         Left            =   225
         TabIndex        =   18
         Top             =   2265
         Width           =   780
      End
      Begin VB.Shape shpTabGridColor 
         Height          =   180
         Left            =   1635
         Top             =   1890
         Width           =   1605
      End
      Begin VB.Label lblTabTextColor 
         Caption         =   "�ı���ɫ"
         Height          =   180
         Left            =   1635
         TabIndex        =   16
         Top             =   1530
         Width           =   1605
      End
      Begin VB.Label lblTabFont 
         Caption         =   "����,9"
         Height          =   180
         Left            =   1635
         TabIndex        =   14
         Top             =   1185
         Width           =   1605
      End
      Begin VB.Label lblTabRowHeight 
         AutoSize        =   -1  'True
         Caption         =   "��С�и�"
         Height          =   180
         Left            =   1935
         TabIndex        =   7
         Top             =   405
         Width           =   720
      End
      Begin VB.Label lblTabTiers 
         AutoSize        =   -1  'True
         Caption         =   "��ͷ����"
         Height          =   180
         Left            =   435
         TabIndex        =   9
         Top             =   750
         Width           =   720
      End
      Begin VB.Label lblTabCols 
         AutoSize        =   -1  'True
         Caption         =   "�������"
         Height          =   180
         Left            =   435
         TabIndex        =   4
         Top             =   405
         Width           =   720
      End
   End
   Begin VB.PictureBox picCloumn 
      BackColor       =   &H00C0E0FF&
      BorderStyle     =   0  'None
      Height          =   3705
      Left            =   540
      ScaleHeight     =   3705
      ScaleWidth      =   10290
      TabIndex        =   2
      Top             =   150
      Width           =   10290
      Begin TabDlg.SSTab sstCOl 
         Height          =   3075
         Left            =   6945
         TabIndex        =   144
         Top             =   435
         Width           =   3195
         _ExtentX        =   5636
         _ExtentY        =   5424
         _Version        =   393216
         Style           =   1
         Tabs            =   2
         TabHeight       =   520
         TabCaption(0)   =   "���÷������  "
         TabPicture(0)   =   "frmTendStyle.frx":058A
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "vfgColRelation"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).Control(1)=   "picImg(0)"
         Tab(0).Control(1).Enabled=   0   'False
         Tab(0).ControlCount=   2
         TabCaption(1)   =   "����ҽ����Ӧ��  "
         TabPicture(1)   =   "frmTendStyle.frx":05A6
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "picImg(1)"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "vsfColDoctor"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).ControlCount=   2
         Begin VB.PictureBox picImg 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   1
            Left            =   -72135
            Picture         =   "frmTendStyle.frx":05C2
            ScaleHeight     =   225
            ScaleWidth      =   285
            TabIndex        =   146
            Top             =   30
            Width           =   285
         End
         Begin VB.PictureBox picImg 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   225
            Index           =   0
            Left            =   1215
            Picture         =   "frmTendStyle.frx":6E14
            ScaleHeight     =   225
            ScaleWidth      =   285
            TabIndex        =   145
            Top             =   45
            Width           =   285
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfColDoctor 
            Height          =   2655
            Left            =   -74910
            TabIndex        =   68
            Top             =   360
            Width           =   3015
            _cx             =   5318
            _cy             =   4683
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
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   0
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   3
            FixedRows       =   2
            FixedCols       =   2
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmTendStyle.frx":D666
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
         Begin VSFlex8Ctl.VSFlexGrid vfgColRelation 
            Height          =   2595
            Left            =   120
            TabIndex        =   67
            Top             =   390
            Width           =   2970
            _cx             =   5239
            _cy             =   4577
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
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   0
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   4
            FixedRows       =   2
            FixedCols       =   2
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmTendStyle.frx":D6F6
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
      Begin VB.Frame fraSplit 
         Height          =   30
         Index           =   5
         Left            =   7740
         TabIndex        =   142
         Top             =   255
         Width           =   2235
      End
      Begin VB.Frame fraSplit 
         Height          =   30
         Index           =   4
         Left            =   1035
         TabIndex        =   140
         Top             =   255
         Width           =   2235
      End
      Begin VB.PictureBox picAlign 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   315
         Left            =   4425
         ScaleHeight     =   315
         ScaleWidth      =   1005
         TabIndex        =   65
         TabStop         =   0   'False
         Top             =   3180
         Width           =   1005
         Begin VB.OptionButton optAlign 
            Height          =   315
            Index           =   0
            Left            =   0
            Picture         =   "frmTendStyle.frx":D7B2
            Style           =   1  'Graphical
            TabIndex        =   135
            TabStop         =   0   'False
            Top             =   0
            Width           =   345
         End
         Begin VB.OptionButton optAlign 
            Height          =   315
            Index           =   1
            Left            =   330
            Picture         =   "frmTendStyle.frx":DB38
            Style           =   1  'Graphical
            TabIndex        =   134
            TabStop         =   0   'False
            Top             =   0
            Width           =   345
         End
         Begin VB.OptionButton optAlign 
            Height          =   315
            Index           =   2
            Left            =   660
            Picture         =   "frmTendStyle.frx":DEC8
            Style           =   1  'Graphical
            TabIndex        =   133
            TabStop         =   0   'False
            Top             =   0
            Width           =   345
         End
      End
      Begin VB.CheckBox chk 
         Caption         =   "�Խ���"
         Height          =   210
         Left            =   5610
         TabIndex        =   66
         Top             =   3240
         Width           =   1020
      End
      Begin VB.TextBox txtColumnPrefix 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4425
         TabIndex        =   61
         Top             =   2430
         Width           =   2295
      End
      Begin VB.CommandButton cmdColumn 
         Caption         =   "Ӧ��(&Y)"
         Height          =   300
         Index           =   2
         Left            =   2760
         TabIndex        =   100
         Top             =   1875
         Width           =   1100
      End
      Begin VB.TextBox txtColumnPostfix 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4425
         TabIndex        =   63
         Top             =   2805
         Width           =   2295
      End
      Begin VB.ListBox lstColumnUsed 
         Height          =   1680
         Left            =   3990
         TabIndex        =   59
         Top             =   690
         Width           =   2730
      End
      Begin MSComCtl2.UpDown udColumnNo 
         Height          =   300
         Left            =   4605
         TabIndex        =   58
         Top             =   360
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         AutoBuddy       =   -1  'True
         BuddyControl    =   "txtColumnNo"
         BuddyDispid     =   196621
         OrigLeft        =   5400
         OrigTop         =   75
         OrigRight       =   5640
         OrigBottom      =   375
         Max             =   5
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtColumnNo 
         Enabled         =   0   'False
         Height          =   300
         Left            =   4215
         MaxLength       =   2
         TabIndex        =   57
         Text            =   "1"
         Top             =   360
         Width           =   405
      End
      Begin VB.CommandButton cmdColumn 
         Caption         =   "ɾ��(&E)"
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   2760
         TabIndex        =   55
         Top             =   1425
         Width           =   1100
      End
      Begin VB.CommandButton cmdColumn 
         Caption         =   "ѡ��(&S)"
         Height          =   300
         Index           =   0
         Left            =   2760
         TabIndex        =   54
         Top             =   1125
         Width           =   1100
      End
      Begin VB.ListBox lstColumnItems 
         Height          =   2760
         Left            =   240
         TabIndex        =   53
         Top             =   705
         Width           =   2370
      End
      Begin VB.Label lblColRelation 
         AutoSize        =   -1  'True
         Caption         =   "���й�ϵ"
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
         Left            =   6945
         TabIndex        =   143
         Top             =   180
         Width           =   780
      End
      Begin VB.Label lblColText 
         AutoSize        =   -1  'True
         Caption         =   "��������"
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
         Left            =   240
         TabIndex        =   141
         Top             =   180
         Width           =   780
      End
      Begin VB.Label lbl�ж��� 
         AutoSize        =   -1  'True
         Caption         =   "����"
         Height          =   180
         Left            =   3990
         TabIndex        =   64
         Top             =   3240
         Width           =   360
      End
      Begin VB.Label lblColumnPostfix 
         AutoSize        =   -1  'True
         Caption         =   "��׺"
         Height          =   180
         Left            =   3990
         TabIndex        =   62
         Top             =   2850
         Width           =   360
      End
      Begin VB.Label lblColumnPrefix 
         AutoSize        =   -1  'True
         Caption         =   "ǰ׺"
         Height          =   180
         Left            =   3990
         TabIndex        =   60
         Top             =   2475
         Width           =   360
      End
      Begin VB.Label lblColumnNo 
         AutoSize        =   -1  'True
         Caption         =   "��        ��������Ŀ:"
         Height          =   180
         Left            =   4005
         TabIndex        =   56
         Top             =   420
         Width           =   1890
      End
      Begin VB.Label lblColumnItems 
         AutoSize        =   -1  'True
         Caption         =   "��ѡ�����¼��Ŀ:"
         Height          =   180
         Left            =   240
         TabIndex        =   52
         Top             =   420
         Width           =   1530
      End
   End
   Begin VB.PictureBox picPaper 
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   0  'None
      Height          =   3405
      Left            =   660
      ScaleHeight     =   3405
      ScaleWidth      =   6885
      TabIndex        =   3
      Top             =   1005
      Width           =   6885
      Begin VB.CheckBox chkHideTime 
         Caption         =   "ʱ�������ش�ӡ"
         Height          =   195
         Left            =   3195
         TabIndex        =   99
         Top             =   2640
         Width           =   1635
      End
      Begin VB.CheckBox chk���кϲ� 
         Caption         =   "����ʱ���кϲ���ӡ"
         Height          =   195
         Left            =   1110
         TabIndex        =   98
         Top             =   2640
         Width           =   1980
      End
      Begin VB.Frame fraPaper 
         Height          =   30
         Index           =   0
         Left            =   210
         TabIndex        =   101
         Top             =   2475
         Width           =   6390
      End
      Begin VB.ComboBox cboPaperKind 
         Height          =   300
         Left            =   1050
         Style           =   2  'Dropdown List
         TabIndex        =   70
         Top             =   435
         Width           =   5355
      End
      Begin VB.TextBox txtHeight 
         Height          =   300
         Left            =   3600
         MaxLength       =   6
         TabIndex        =   76
         Text            =   "297.08"
         Top             =   825
         Width           =   975
      End
      Begin VB.TextBox txtWidth 
         Height          =   300
         Left            =   1290
         MaxLength       =   6
         TabIndex        =   72
         Text            =   "210.05"
         Top             =   825
         Width           =   945
      End
      Begin VB.TextBox txtMarjin 
         Height          =   300
         Index           =   3
         Left            =   3000
         MaxLength       =   6
         TabIndex        =   93
         Text            =   "19"
         Top             =   1665
         Width           =   615
      End
      Begin VB.TextBox txtMarjin 
         Height          =   300
         Index           =   2
         Left            =   1290
         MaxLength       =   6
         TabIndex        =   89
         Text            =   "19"
         Top             =   1665
         Width           =   615
      End
      Begin VB.TextBox txtMarjin 
         Height          =   300
         Index           =   1
         Left            =   3000
         MaxLength       =   6
         TabIndex        =   85
         Text            =   "25"
         Top             =   1260
         Width           =   615
      End
      Begin VB.OptionButton optOrient 
         Caption         =   "����"
         Enabled         =   0   'False
         Height          =   270
         Index           =   0
         Left            =   5745
         TabIndex        =   96
         Top             =   1275
         Value           =   -1  'True
         Width           =   750
      End
      Begin VB.OptionButton optOrient 
         Caption         =   "����"
         Enabled         =   0   'False
         Height          =   270
         Index           =   1
         Left            =   5745
         TabIndex        =   97
         Top             =   1680
         Width           =   750
      End
      Begin MSComCtl2.UpDown udHeight 
         Height          =   300
         Left            =   4575
         TabIndex        =   77
         TabStop         =   0   'False
         Top             =   840
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtHeight"
         BuddyDispid     =   196676
         OrigLeft        =   4170
         OrigTop         =   900
         OrigRight       =   4425
         OrigBottom      =   1185
         Max             =   765
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udWidth 
         Height          =   300
         Left            =   2235
         TabIndex        =   73
         TabStop         =   0   'False
         Top             =   840
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtWidth"
         BuddyDispid     =   196677
         OrigLeft        =   1830
         OrigTop         =   893
         OrigRight       =   2070
         OrigBottom      =   1178
         Max             =   765
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udMarjin 
         Height          =   300
         Index           =   0
         Left            =   1905
         TabIndex        =   82
         TabStop         =   0   'False
         Top             =   1275
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtMarjin(0)"
         BuddyDispid     =   196678
         BuddyIndex      =   0
         OrigLeft        =   1785
         OrigTop         =   1410
         OrigRight       =   2025
         OrigBottom      =   1710
         Max             =   210
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udMarjin 
         Height          =   300
         Index           =   1
         Left            =   3615
         TabIndex        =   86
         TabStop         =   0   'False
         Top             =   1275
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtMarjin(1)"
         BuddyDispid     =   196678
         BuddyIndex      =   1
         OrigLeft        =   3780
         OrigTop         =   1410
         OrigRight       =   4020
         OrigBottom      =   1710
         Max             =   210
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udMarjin 
         Height          =   300
         Index           =   2
         Left            =   1905
         TabIndex        =   90
         TabStop         =   0   'False
         Top             =   1680
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtMarjin(2)"
         BuddyDispid     =   196678
         BuddyIndex      =   2
         OrigLeft        =   1785
         OrigTop         =   1815
         OrigRight       =   2025
         OrigBottom      =   2115
         Max             =   210
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComCtl2.UpDown udMarjin 
         Height          =   300
         Index           =   3
         Left            =   3615
         TabIndex        =   94
         TabStop         =   0   'False
         Top             =   1680
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         BuddyControl    =   "txtMarjin(3)"
         BuddyDispid     =   196678
         BuddyIndex      =   3
         OrigLeft        =   3780
         OrigTop         =   1815
         OrigRight       =   4020
         OrigBottom      =   2115
         Max             =   210
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtMarjin 
         Height          =   300
         Index           =   0
         Left            =   1290
         MaxLength       =   6
         TabIndex        =   81
         Text            =   "25"
         Top             =   1260
         Width           =   615
      End
      Begin VB.Label lblPrinter 
         AutoSize        =   -1  'True
         Caption         =   "��ӡ��:"
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
         Left            =   270
         TabIndex        =   118
         Top             =   135
         Width           =   690
      End
      Begin VB.Label lblOrient 
         AutoSize        =   -1  'True
         Caption         =   "ֽ�ŷ���:"
         Height          =   180
         Left            =   4830
         TabIndex        =   127
         Top             =   1320
         Width           =   810
      End
      Begin VB.Label lblPaperHint 
         AutoSize        =   -1  'True
         Caption         =   "ע��:  ���ʵ�ʴ�ӡ���͵�ǰ��ӡ�����������ܵ���ֽ������ʧЧ��"
         Height          =   180
         Left            =   270
         TabIndex        =   119
         Top             =   2115
         Width           =   5490
      End
      Begin VB.Label lblPaper 
         AutoSize        =   -1  'True
         Caption         =   "ֽ������"
         Height          =   180
         Left            =   270
         TabIndex        =   69
         Top             =   480
         Width           =   720
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Index           =   0
         Left            =   2535
         TabIndex        =   74
         Top             =   885
         Width           =   360
      End
      Begin VB.Label lblHeight 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   3375
         TabIndex        =   75
         Top             =   885
         Width           =   180
      End
      Begin VB.Label lblWidth 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Left            =   1065
         TabIndex        =   71
         Top             =   885
         Width           =   180
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Index           =   1
         Left            =   4890
         TabIndex        =   78
         Top             =   885
         Width           =   360
      End
      Begin VB.Label lblRound 
         AutoSize        =   -1  'True
         Caption         =   "ҳ�߾�:"
         Height          =   180
         Left            =   270
         TabIndex        =   79
         Top             =   1335
         Width           =   630
      End
      Begin VB.Label lblMarjin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Index           =   3
         Left            =   2790
         TabIndex        =   92
         Top             =   1725
         Width           =   180
      End
      Begin VB.Label lblMarjin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Index           =   2
         Left            =   1065
         TabIndex        =   88
         Top             =   1725
         Width           =   180
      End
      Begin VB.Label lblMarjin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Index           =   1
         Left            =   2790
         TabIndex        =   84
         Top             =   1320
         Width           =   180
      End
      Begin VB.Label lblMarjin 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��"
         Height          =   180
         Index           =   0
         Left            =   1065
         TabIndex        =   80
         Top             =   1320
         Width           =   180
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Index           =   2
         Left            =   2190
         TabIndex        =   83
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Index           =   3
         Left            =   3915
         TabIndex        =   87
         Top             =   1320
         Width           =   360
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Index           =   4
         Left            =   2190
         TabIndex        =   91
         Top             =   1725
         Width           =   360
      End
      Begin VB.Label lblUnit 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Index           =   5
         Left            =   3915
         TabIndex        =   95
         Top             =   1725
         Width           =   360
      End
   End
   Begin VB.PictureBox picFoot 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0FF&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3405
      Left            =   870
      ScaleHeight     =   3405
      ScaleWidth      =   6885
      TabIndex        =   136
      Top             =   1500
      Width           =   6885
      Begin VB.ComboBox cboҳ�� 
         Enabled         =   0   'False
         Height          =   300
         Left            =   3585
         Style           =   2  'Dropdown List
         TabIndex        =   116
         Top             =   3030
         Width           =   1320
      End
      Begin VB.CheckBox chk���������� 
         Caption         =   "����������"
         Height          =   195
         Left            =   5280
         TabIndex        =   117
         Top             =   3090
         Width           =   1395
      End
      Begin VB.CheckBox chkҳ�� 
         Caption         =   "��ӡҳ��"
         Height          =   195
         Left            =   270
         TabIndex        =   111
         Top             =   3075
         Width           =   1155
      End
      Begin VB.OptionButton optPageAlign 
         Enabled         =   0   'False
         Height          =   315
         Index           =   2
         Left            =   2160
         Picture         =   "frmTendStyle.frx":E221
         Style           =   1  'Graphical
         TabIndex        =   114
         TabStop         =   0   'False
         Top             =   3030
         Width           =   345
      End
      Begin VB.OptionButton optPageAlign 
         Enabled         =   0   'False
         Height          =   315
         Index           =   1
         Left            =   1830
         Picture         =   "frmTendStyle.frx":E57A
         Style           =   1  'Graphical
         TabIndex        =   113
         TabStop         =   0   'False
         Top             =   3030
         Width           =   345
      End
      Begin VB.OptionButton optPageAlign 
         Enabled         =   0   'False
         Height          =   315
         Index           =   0
         Left            =   1500
         Picture         =   "frmTendStyle.frx":E90A
         Style           =   1  'Graphical
         TabIndex        =   112
         TabStop         =   0   'False
         Top             =   3030
         Width           =   345
      End
      Begin VB.CommandButton cmdͬ�� 
         Caption         =   "ͬ��(&G)"
         Height          =   350
         Left            =   5730
         TabIndex        =   109
         ToolTipText     =   "���л����ļ���ҳüҳ���뵱ǰ�ļ���ҳüҳ�Ÿ�ʽһ��"
         Top             =   1350
         Width           =   1100
      End
      Begin VB.CommandButton cmdOpen 
         Caption         =   "��ͼ(&I)"
         Height          =   350
         Left            =   150
         TabIndex        =   103
         Top             =   1350
         Width           =   1200
      End
      Begin VB.CheckBox chkI 
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   5250
         Picture         =   "frmTendStyle.frx":EC90
         Style           =   1  'Graphical
         TabIndex        =   108
         ToolTipText     =   "б��(Alt+I)"
         Top             =   1350
         Width           =   345
      End
      Begin VB.CheckBox chkU 
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   4920
         Picture         =   "frmTendStyle.frx":154E2
         Style           =   1  'Graphical
         TabIndex        =   107
         ToolTipText     =   "�»���(Alt+U)"
         Top             =   1350
         Width           =   345
      End
      Begin VB.CheckBox chkB 
         BeginProperty Font 
            Name            =   "����"
            Size            =   14.25
            Charset         =   134
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   350
         Left            =   4590
         Picture         =   "frmTendStyle.frx":1BD34
         Style           =   1  'Graphical
         TabIndex        =   106
         ToolTipText     =   "����(Alt+B)"
         Top             =   1350
         Width           =   345
      End
      Begin VB.ComboBox cboFSize 
         Height          =   300
         Left            =   3810
         Style           =   2  'Dropdown List
         TabIndex        =   105
         Top             =   1380
         Width           =   750
      End
      Begin RichTextLib.RichTextBox rtbHead 
         Height          =   1200
         Left            =   30
         TabIndex        =   102
         TabStop         =   0   'False
         Top             =   30
         Width           =   6810
         _ExtentX        =   12012
         _ExtentY        =   2117
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         ScrollBars      =   2
         OLEDragMode     =   0
         OLEDropMode     =   0
         TextRTF         =   $"frmTendStyle.frx":22586
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin RichTextLib.RichTextBox rtbFoot 
         Height          =   1200
         Left            =   30
         TabIndex        =   110
         TabStop         =   0   'False
         Top             =   1770
         Width           =   6810
         _ExtentX        =   12012
         _ExtentY        =   2117
         _Version        =   393217
         BorderStyle     =   0
         Enabled         =   -1  'True
         ScrollBars      =   2
         OLEDragMode     =   0
         OLEDropMode     =   0
         TextRTF         =   $"frmTendStyle.frx":22623
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   10.5
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin VB.ComboBox cboFont 
         Height          =   300
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   104
         Top             =   1380
         Width           =   1905
      End
      Begin VB.Label lblҳ�� 
         AutoSize        =   -1  'True
         Caption         =   "ҳ��λ��"
         Enabled         =   0   'False
         Height          =   180
         Left            =   2790
         TabIndex        =   115
         Top             =   3090
         Width           =   720
      End
      Begin VB.Label lbl���� 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   1515
         TabIndex        =   137
         Top             =   1440
         Width           =   360
      End
   End
   Begin VB.PictureBox picLabel 
      BackColor       =   &H00C0C0FF&
      BorderStyle     =   0  'None
      Height          =   3405
      Left            =   90
      ScaleHeight     =   3405
      ScaleWidth      =   6885
      TabIndex        =   1
      Top             =   840
      Width           =   6885
      Begin VB.CheckBox chkLabelCrLf 
         Caption         =   "����"
         Height          =   195
         Left            =   3870
         TabIndex        =   48
         Top             =   3000
         Width           =   780
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "ɾ��(&D)"
         Enabled         =   0   'False
         Height          =   300
         Index           =   1
         Left            =   2700
         TabIndex        =   45
         Top             =   1185
         Width           =   1100
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "ѡ��(&U)"
         Height          =   300
         Index           =   0
         Left            =   2700
         TabIndex        =   44
         Top             =   885
         Width           =   1100
      End
      Begin VB.TextBox txtLabelPrefix 
         Enabled         =   0   'False
         Height          =   300
         Left            =   5670
         TabIndex        =   50
         Top             =   2940
         Width           =   975
      End
      Begin VB.ListBox lstLabelUsed 
         Height          =   2400
         Left            =   3870
         TabIndex        =   47
         Top             =   465
         Width           =   2775
      End
      Begin VB.ListBox lstLabelItems 
         Height          =   2760
         Left            =   240
         TabIndex        =   43
         Top             =   465
         Width           =   2370
      End
      Begin VB.CommandButton cmdLabel 
         Caption         =   "Ӧ��(&A)"
         Height          =   300
         Index           =   2
         Left            =   2700
         TabIndex        =   51
         Top             =   1635
         Width           =   1100
      End
      Begin VB.Label lblLabelPrefix 
         AutoSize        =   -1  'True
         Caption         =   "ǰ׺�ı�"
         Height          =   180
         Left            =   4905
         TabIndex        =   49
         Top             =   3000
         Width           =   720
      End
      Begin VB.Label lblLabelUsed 
         AutoSize        =   -1  'True
         Caption         =   "�����õı�ǩ:"
         Height          =   180
         Left            =   3885
         TabIndex        =   46
         Top             =   180
         Width           =   1170
      End
      Begin VB.Label lblLabelItems 
         AutoSize        =   -1  'True
         Caption         =   "��ѡ��ǩ��Ŀ:"
         Height          =   180
         Left            =   240
         TabIndex        =   42
         Top             =   180
         Width           =   1170
      End
   End
   Begin MSComDlg.CommonDialog dlgThis 
      Left            =   8265
      Top             =   6345
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin MSComctlLib.StatusBar stbThis 
      Align           =   2  'Align Bottom
      Height          =   360
      Left            =   0
      TabIndex        =   132
      Top             =   7230
      Width           =   12225
      _ExtentX        =   21564
      _ExtentY        =   635
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   3
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Bevel           =   2
            Object.Width           =   2355
            MinWidth        =   882
            Picture         =   "frmTendStyle.frx":226C0
            Text            =   "�������"
            TextSave        =   "�������"
            Key             =   "ZLFLAG"
            Object.ToolTipText     =   "��ӭʹ���������޹�˾���"
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   17674
            Text            =   "���Ը���ҽԺʵ����������õ��������¼�Ĳ鿴�������ʽ��"
            TextSave        =   "���Ը���ҽԺʵ����������õ��������¼�Ĳ鿴�������ʽ��"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
            Enabled         =   0   'False
            Object.Width           =   953
            MinWidth        =   529
            Text            =   "�༭"
            TextSave        =   "�༭"
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
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   4005
      Index           =   0
      Left            =   8640
      ScaleHeight     =   4005
      ScaleWidth      =   6645
      TabIndex        =   128
      Top             =   360
      Width           =   6645
      Begin XtremeSuiteControls.TabControl tbcStyle 
         Height          =   3930
         Left            =   600
         TabIndex        =   131
         Top             =   210
         Width           =   5460
         _Version        =   589884
         _ExtentX        =   9631
         _ExtentY        =   6932
         _StockProps     =   64
      End
   End
   Begin VB.PictureBox picPane 
      BackColor       =   &H8000000D&
      BorderStyle     =   0  'None
      Height          =   1665
      Index           =   1
      Left            =   8700
      ScaleHeight     =   1665
      ScaleWidth      =   7860
      TabIndex        =   129
      Top             =   4560
      Width           =   7860
      Begin VSFlex8Ctl.VSFlexGrid vfgThis 
         Height          =   1425
         Left            =   0
         TabIndex        =   130
         Top             =   -15
         Width           =   3810
         _cx             =   6720
         _cy             =   2514
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
         BackColorFixed  =   8421504
         ForeColorFixed  =   12632256
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483644
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   0
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   8
         Cols            =   6
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmTendStyle.frx":22F54
         ScrollTrack     =   -1  'True
         ScrollBars      =   3
         ScrollTips      =   0   'False
         MergeCells      =   1
         MergeCompare    =   0
         AutoResize      =   -1  'True
         AutoSizeMode    =   1
         AutoSearch      =   0
         AutoSearchDelay =   2
         MultiTotals     =   0   'False
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
         ForeColorFrozen =   0
         WallPaperAlignment=   9
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
   End
   Begin VB.PictureBox picSum 
      BorderStyle     =   0  'None
      Height          =   1875
      Left            =   585
      ScaleHeight     =   1875
      ScaleWidth      =   6885
      TabIndex        =   138
      Top             =   3255
      Width           =   6885
      Begin VSFlex8Ctl.VSFlexGrid vsf 
         Height          =   1020
         Left            =   435
         TabIndex        =   139
         Top             =   390
         Width           =   1995
         _cx             =   3519
         _cy             =   1799
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
      End
   End
   Begin XtremeCommandBars.CommandBars cbsThis 
      Left            =   525
      Top             =   150
      _Version        =   589884
      _ExtentX        =   635
      _ExtentY        =   635
      _StockProps     =   0
      VisualTheme     =   2
   End
   Begin XtremeDockingPane.DockingPane dkpMan 
      Bindings        =   "frmTendStyle.frx":2303C
      Left            =   75
      Top             =   60
      _Version        =   589884
      _ExtentX        =   450
      _ExtentY        =   423
      _StockProps     =   0
      VisualTheme     =   5
   End
End
Attribute VB_Name = "frmTendFileStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Const conRowHeight = 300        '��׼�и߶�

'��ӡֽ�ų���(256=�Զ���)
Const PageSize1 = "�ż㣬 8 1/2��11 Ӣ��"
Const PageSize2 = "+A611 С���ż㣬 8 1/2��11 Ӣ��"
Const PageSize3 = "С�ͱ��� 11��17 Ӣ��"
Const PageSize4 = "�����ʣ� 17��11 Ӣ��"
Const PageSize5 = "�����ļ��� 8 1/2��14 Ӣ��"
Const PageSize6 = "�����飬5 1/2��8 1/2 Ӣ��"
Const PageSize7 = "�����ļ���7 1/2��10 1/2 Ӣ��"
Const PageSize8 = "A3, 297��420 ����"
Const PageSize9 = "A4, 210��297 ����"
Const PageSize10 = "A4С�ţ� 210��297 ����"
Const PageSize11 = "A5, 148��210 ����"
Const PageSize12 = "B4, 250��354 ����"
Const PageSize13 = "B5, 182��257 ����"
Const PageSize14 = "�Կ����� 8 1/2��13 Ӣ��"
Const PageSize15 = "�Ŀ����� 215��275 ����"
Const PageSize16 = "10��14 Ӣ��"
Const PageSize17 = "11��17 Ӣ��"
Const PageSize18 = "������8 1/2��11 Ӣ��"
Const PageSize19 = "#9 �ŷ⣬ 3 7/8��8 7/8 Ӣ��"
Const PageSize20 = "#10 �ŷ⣬ 4 1/8��9 1/2 Ӣ��"
Const PageSize21 = "#11 �ŷ⣬ 4 1/2��10 3/8 Ӣ��"
Const PageSize22 = "#12 �ŷ⣬ 4 1/2��11 Ӣ��"
Const PageSize23 = "#14 �ŷ⣬ 5��11 1/2 Ӣ��"
Const PageSize24 = "C �ߴ繤����"
Const PageSize25 = "D �ߴ繤����"
Const PageSize26 = "E �ߴ繤����"
Const PageSize27 = "DL ���ŷ⣬ 110��220 ����"
Const PageSize28 = "C5 ���ŷ⣬ 162��229 ����"
Const PageSize29 = "C3 ���ŷ⣬ 324��458 ����"
Const PageSize30 = "C4 ���ŷ⣬ 229��324 ����"
Const PageSize31 = "C6 ���ŷ⣬ 114��162 ����"
Const PageSize32 = "C65 ���ŷ⣬114��229 ����"
Const PageSize33 = "B4 ���ŷ⣬ 250��353 ����"
Const PageSize34 = "B5 ���ŷ⣬176��250 ����"
Const PageSize35 = "B6 ���ŷ⣬ 176��125 ����"
Const PageSize36 = "�ŷ⣬ 110��230 ����"
Const PageSize37 = "�ŷ������ 3 7/8��7 1/2 Ӣ��"
Const PageSize38 = "�ŷ⣬ 3 5/8��6 1/2 Ӣ��"
Const PageSize39 = "U.S. ��׼��д���� 14 7/8��11 Ӣ��"
Const PageSize40 = "�¹���׼��д���� 8 1/2��12 Ӣ��"
Const PageSize41 = "�¹����ɸ�д���� 8 1/2��13 Ӣ��"

Private mlngFileID As Long       '���༭�ļ�¼ID���޸ġ�����ʱ���ϼ�����ͨ��ShowMe���ݽ���,����ʱΪ0��
Private mblnRTBFoot As Boolean
Private mblnOk As Boolean        '�Ƿ���ɱ༭�˳�
Private mlngPageRow As Long     '��ǰҳ���ʽ�������Ч������

Private WithEvents mclsVsf As clsVsf
Attribute mclsVsf.VB_VarHelpID = -1

'ҳüҳ�����
'######################################################################################################
Private Type CHARRANGE
    cpMin As Long
    cpMax As Long
End Type
'����
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'�������ڸ�ʽ��ָ���豸�������Ϣ
Private Type FORMATRANGE
    hDC As Long             '��Ⱦ�豸
    hdcTarget As Long       'Ŀ���豸
    rc As RECT              '��Ⱦ���򣬵�λ��羡�
    rcPage As RECT          '��Ⱦ�豸���������򣬵�λ��羡�
    chrg As CHARRANGE       '���ڸ�ʽ�����ı���Χ��
End Type

Private Type PageInfo
    PageNumber As Long      'ҳ��
    Start As Long           '�ַ���ʼλ��
    End As Long             '�ַ���ֹλ��
    ActualHeight As Long    '��ҳʵ�ʴ�ӡ�߶�
End Type
Private AllPages() As PageInfo   'ҳ��Ϣ
Private Const WM_PASTE = &H302&              'ճ��
Private Const WM_USER = &H400                'ͨ���� WM_USER + X ���Զ�����Ϣ
Private Const EM_FORMATRANGE = (WM_USER + 57)    'Ϊĳһ�豸��ʽ��ָ����Χ���ı���
Private Const EM_SETTARGETDEVICE = (WM_USER + 72) '�����������������õ�Ŀ���豸���п�
Private Const EM_HIDESELECTION = (WM_USER + 63)  '��ʾ/�����ı���
Private Const PHYSICALOFFSETX = 112  '���ڴ�ӡ�豸���ԣ���ʾ������ҳ�����Ե���ɴ�ӡ��������Ե�ľ��룬�����豸��λ��
Private Const PHYSICALOFFSETY = 113  '���ڴ�ӡ�豸���ԣ���ʾ������ҳ���ϱ�Ե���ɴ�ӡ������ϱ�Ե�ľ��룬�����豸��λ��
Private Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hDC As Long, ByVal nIndex As Long) As Long
Private Declare Function lstrlen Lib "kernel32" Alias "lstrlenA" (ByVal lpString As String) As Long '��ȡ��Ӣ�Ļ���ַ�������
'######################################################################################################

'��ʱ����
Private rsItems As New ADODB.Recordset
Private rsTemp As New ADODB.Recordset
Private lngCount As Long
Private strTemp As String
Private lngCurColor As Long
Private strCurFont As String
Private objFont As StdFont
Private mblnChanged As Boolean
Private rtbThis As Object
Private mstrColDoctor As String     '��Ŀ����;�ܸ�����|��Ŀ����;ҽ������ (��ҽ����ʹ��)

'��Ŀ���ʣ�¼���ѡ�񣬵�ѡ����ѡ
'һ��һ����Ŀ�ģ����򲻿���
'һ����������Ŀ�ģ���Ŀ֮������ǰ׺���׺����ʶ������Ŀ���ʱ�����ͬ��ֻ����¼�����ѡ����
'һ��������������Ŀ�ģ�ֻ����¼����
'�����޸�:�����ͷ�ı���/��ֻ��������Ŀ,��Ŀ�ķָ���Ҳ��/,�����ͱ�����ͬ,ֻ����¼���ѡ����;���򲻿���
'
'�����ṩһ�ָ�ʽ����/�£���8/6


Private Property Let DataChanged(vData As Boolean)
    
    mblnChanged = vData
        
    If mblnChanged Then
        stbThis.Panels(3).Enabled = True
    Else
        stbThis.Panels(3).Enabled = False
    End If
    
End Property

Private Property Get DataChanged() As Boolean
    
    DataChanged = mblnChanged

End Property

Private Function GetPaperName(ByVal intSize As Integer) As String
    '���ܣ� ���ݵ�ǰ��ӡ�������ã���ȡֽ������
    '���أ� ֽ������
    If intSize = 256 Then
        GetPaperName = "�û��Զ��� ..."
    ElseIf intSize >= 1 And intSize <= 41 Then
        GetPaperName = Switch( _
            intSize = 1, PageSize1, intSize = 2, PageSize2, intSize = 3, PageSize3, intSize = 4, PageSize4, intSize = 5, PageSize5, _
            intSize = 6, PageSize6, intSize = 7, PageSize7, intSize = 8, PageSize8, intSize = 9, PageSize9, intSize = 10, PageSize10, _
            intSize = 11, PageSize11, intSize = 12, PageSize12, intSize = 13, PageSize13, intSize = 14, PageSize14, intSize = 15, PageSize15, _
            intSize = 16, PageSize16, intSize = 17, PageSize17, intSize = 18, PageSize18, intSize = 19, PageSize19, intSize = 20, PageSize20, _
            intSize = 21, PageSize21, intSize = 22, PageSize22, intSize = 23, PageSize23, intSize = 24, PageSize24, intSize = 25, PageSize25, _
            intSize = 26, PageSize26, intSize = 27, PageSize27, intSize = 28, PageSize28, intSize = 29, PageSize29, intSize = 30, PageSize30, _
            intSize = 31, PageSize31, intSize = 32, PageSize32, intSize = 33, PageSize33, intSize = 34, PageSize34, intSize = 35, PageSize35, _
            intSize = 36, PageSize36, intSize = 37, PageSize37, intSize = 38, PageSize38, intSize = 39, PageSize39, intSize = 40, PageSize40, _
            intSize = 41, PageSize41)
    Else
        GetPaperName = "���ɲ��ֽ�� ..."
    End If
End Function

Private Sub LoadPaper()
    '---------------------------------------------------
    '���ܣ�װ�뵱ǰ��ӡ�����õ�ֽ��
    '---------------------------------------------------
    Dim intCurPaper As Integer
    Dim strDevice As String
        
    With Me.cboPaperKind
        .AddItem GetPaperName(256)
        .ItemData(.NewIndex) = 256
        If Not ExistsPrinter Then
            .ListIndex = 0
            .Enabled = False
            Exit Sub
        End If
        
        strDevice = GetSetting("ZLSOFT", "����ģ��\zl9PrintMode\Default", "DeviceName", Printer.DeviceName)
        For lngCount = 0 To Printers.Count - 1
            If Printers(lngCount).DeviceName = strDevice Then
                Set Printer = Printers(lngCount)
                Exit For
            End If
        Next
        Me.lblPrinter.Caption = "��ǰ��ӡ��: " & Printer.DeviceName
        
        intCurPaper = Printer.PaperSize
        .Enabled = True
        For lngCount = 1 To 41
            Err = 0: On Error Resume Next
            Printer.PaperSize = lngCount
            Err = 0: On Error GoTo 0
            If Printer.PaperSize = lngCount Then
                .AddItem GetPaperName(lngCount)
                .ItemData(.NewIndex) = lngCount
                If lngCount = intCurPaper Then .ListIndex = .NewIndex
            End If
        Next
        If .ListIndex < 0 Then .ListIndex = 0
    End With
End Sub

Private Sub cboҳ��_Click()
    If Val(cboҳ��.Tag) = Val(cboҳ��.ItemData(cboҳ��.ListIndex)) Then Exit Sub
    cboҳ��.Tag = Val(cboҳ��.ItemData(cboҳ��.ListIndex))
    mblnChanged = True
End Sub

Private Sub cboҳ��_KeyPress(KeyAscii As Integer)
    Call zlControl.CboMatchIndex(cboҳ��.hWnd, KeyAscii)
End Sub

Private Sub chkHideTime_Click()
    mblnChanged = True
    If chk���кϲ�.Value > 0 And chkHideTime.Value > 0 Then
        chk���кϲ�.Value = 0
    End If
End Sub

Private Sub chk���кϲ�_Click()
    mblnChanged = True
    If chk���кϲ�.Value > 0 And chkHideTime.Value > 0 Then
        chkHideTime.Value = 0
    End If
End Sub

Private Sub mclsVsf_BeforeNewRow(ByVal Row As Long, Col As Long, Cancel As Boolean)
    With vsf
        Cancel = (.TextMatrix(Row, 1) = "" Or .TextMatrix(Row, 2) = "" Or .TextMatrix(Row, 3) = "")
    End With
End Sub

Public Function ShowMe(ByVal frmParent As Object, Optional ByVal lngFileID As Long) As Boolean
    '---------------------------------------------------
    '���ܣ��ϼ�������ñ�����ģ����ݲ���������ʾ����
    '���أ�ȷ�������������޸ĵ�ID��ȡ������0
    '---------------------------------------------------
    Dim strTitle As String
    
    mlngFileID = lngFileID

    Err = 0: On Error GoTo errHand
    
    If RefreshData = False Then
        DataChanged = False
        Unload Me
        Exit Function
    End If
    
    DataChanged = False
    
    '---------------------------------------------------
    '������ʾ
    If frmParent Is Nothing Then
        Me.Show vbModal
    Else
        Me.Show vbModal, frmParent
    End If
'    DataChanged = False
    
    ShowMe = mblnOk
    Exit Function

errHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    ShowMe = False
End Function

Private Function RefreshData() As Boolean


    '---------------------------------------------------
    '���ܣ��ϼ�������ñ�����ģ����ݲ���������ʾ����
    '���أ�ȷ�������������޸ĵ�ID��ȡ������0
    '---------------------------------------------------
    Dim strTitle As String
    Dim strColRelation As String
    Dim arrColName, arrTmp, lngRow As Long, i As Integer, j As Integer
    Dim strColName  As String, strColNumber As String
    '
    With vfgThis
        .Cols = 6
        .Cell(flexcpText, 1, 1, .Rows - 1, .Cols - 1) = ""
        .Cell(flexcpData, 1, 1, .Rows - 1, .Cols - 1) = ""
    End With
    
    With vsf
        .Cell(flexcpText, 1, 1, .Rows - 1, .Cols - 1) = ""
        .Cell(flexcpData, 1, 1, .Rows - 1, .Cols - 1) = ""
    End With
    
    With vfgColRelation
        .Rows = 3
        .Tag = ""
        .Cell(flexcpText, .FixedRows, 0, .Rows - 1, .Cols - 1) = ""
        .Cell(flexcpData, .FixedRows, 0, .Rows - 1, .Cols - 1) = ""
    End With
    
    With vsfColDoctor
        .Rows = 2
        .Tag = ""
    End With
    '---------------------------------------------------
    '���ݵ�ǰ��ӡ����װ���ѡֽ��
    '---------------------------------------------------
    Call LoadPaper
    Call PreparePage
    
    '46251,������,2012-09-11,װ��ҳ�����λ��
    With cboҳ��
        .Clear
        .AddItem "ҳü�Ϸ�": .ItemData(.NewIndex) = 1
        .AddItem "ҳü�·�": .ItemData(.NewIndex) = 2
        .AddItem "ҳ���Ϸ�": .ItemData(.NewIndex) = 3
        .AddItem "ҳ���·�": .ItemData(.NewIndex) = 4
        cboҳ��.Tag = 3
        Call zlControl.CboSetIndex(cboҳ��.hWnd, 2)
    End With
    
    '---------------------------------------------------
    '�������ݻ�ȡ
    '---------------------------------------------------
    Err = 0: On Error GoTo errHand
    
    gstrSQL = "Select l.���, l.����, l.˵�� From �����ļ��б� l Where l.Id = [1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    If rsTemp.RecordCount <= 0 Then Exit Function
    
    Me.Caption = "�����¼��ʽ - " & rsTemp!����
    strTitle = rsTemp!����
    
    '68336
'    gstrSQL = "Select i.������" & _
'            " From ����������Ŀ i, ������������ k" & _
'            " Where k.Id = i.����id And k.���� In ('02', '05', '06') And i.�滻�� = 1" & _
'            " Order By i.����"
    gstrSQL = _
        " Select ������,�滻��" & vbNewLine & _
        " From ����������Ŀ i, ������������ k" & vbNewLine & _
        " Where k.Id = i.����id And ((k.���� In ('02', '05', '06') And i.�滻�� = 1) Or (k.���� = 2 And k.���� = '06' And NVL(i.�滻��,0) = 0))" & vbNewLine & _
        " Order By k.����, k.����, i.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    With rsTemp
        Me.lstLabelItems.Clear
        Do While Not .EOF
            Me.lstLabelItems.AddItem "" & !������
            .MoveNext
        Loop
        If Me.lstLabelItems.ListCount > 0 Then Me.lstLabelItems.ListIndex = 0
    End With
    
    gstrSQL = "Select ��Ŀ���,��Ŀ����,��Ŀ����,��Ŀ��ʾ,��Ŀ���� From �����¼��Ŀ Order By ��Ŀ���"
    Set rsItems = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    With rsItems
        Me.lstColumnItems.Clear
        Me.lstColumnItems.AddItem "����"
        Me.lstColumnItems.AddItem "ʱ��"
        Do While Not .EOF
            Me.lstColumnItems.AddItem "" & !��Ŀ����
            .MoveNext
        Loop
        Me.lstColumnItems.AddItem "��ʿ"
        Me.lstColumnItems.AddItem "ǩ����"
        Me.lstColumnItems.AddItem "ǩ��ʱ��"
        Me.lstColumnItems.ListIndex = 0
        .MoveFirst
    End With
    
    '---------------------------------------------------
    '������ʽ��ȡ
    '---------------------------------------------------
    '�ձ��ʱδ���������ͷ���,����ȱʡ��ͷ�����ǵ���,����ͷ������������СֵΪ2,�ڵ���ñ�ͷʱ����
    Me.optTabTiers(0).Value = True
    Call optTabTiers_Click(0)
    
    gstrSQL = "Select d.�������, d.�����ı�, d.Ҫ������,d.Ҫ�ر�ʾ " & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '�����ʽ'" & _
        " Order By d.�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    mlngPageRow = 0
    With rsTemp
        Do While Not .EOF
            Select Case "" & !Ҫ������
            Case "��ͷ����"
                If Val("" & !�����ı�) = 1 Then
                    Me.optTabTiers(0).Value = True
                    Call optTabTiers_Click(0)
                ElseIf Val("" & !�����ı�) = 2 Then
                    Me.optTabTiers(1).Value = True
                    Call optTabTiers_Click(1)
                Else
                    Me.optTabTiers(2).Value = True
                    Call optTabTiers_Click(2)
                End If
            Case "������":  Me.udTabCols.Value = Val("" & !�����ı�)
            Case "��С�и�"
                Me.txtTabRowHeight.Text = Val("" & !�����ı�)
                Call txtTabRowHeight_Change
            Case "�ı�����"
                Me.lblTabFont.Caption = "" & !�����ı�
                strCurFont = Me.lblTabFont.Caption
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "��") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "б") > 0 Then .Italic = True
                End With
                Set Me.vfgThis.Font = objFont
            Case "�ı���ɫ"
                Me.lblTabTextColor.ForeColor = Val("" & !�����ı�)
                Me.vfgThis.ForeColor = Me.lblTabTextColor.ForeColor
            Case "�����ɫ"
                Me.shpTabGridColor.BorderColor = Val("" & !�����ı�)
                With Me.vfgThis
                    .GridColor = Me.shpTabGridColor.BorderColor
                    .CellBorderRange 2, .FixedCols, 2, .Cols - 1, .GridColor, 0, 0, 0, 1, 0, 0
                    .CellBorderRange 3, .FixedCols, 7, .Cols - 1, .GridColor, 1, 1, 1, 1, 1, 1
                End With
            
            Case "�����ı�"
                Me.txtTitleText.Text = "" & !�����ı�
                Call txtTitleText_Change
            Case "��������"
                Me.lblTitleFont.Caption = "" & !�����ı�
                strCurFont = Me.lblTitleFont.Caption
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "��") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "б") > 0 Then .Italic = True
                End With
                With Me.vfgThis
                    Set .Cell(flexcpFont, 1, .FixedCols, 1, .Cols - 1) = objFont
                    .ROWHEIGHT(1) = objFont.Size * 20 + 150
                End With
            
            Case "��ʼʱ��": Me.udRecordFrom.Value = Val("" & !�����ı�)
            Case "��ֹʱ��": Me.udRecordTo.Value = Val("" & !�����ı�)
            Case "��������"
                Me.lblRecordFont.Caption = "" & !�����ı�
                strCurFont = Me.lblRecordFont.Caption
                Set objFont = New StdFont
                With objFont
                    .Name = Split(strCurFont, ",")(0)
                    .Size = Val(Split(strCurFont, ",")(1))
                    .Bold = False: .Italic = False
                    If InStr(1, strCurFont, "��") > 0 Then .Bold = True
                    If InStr(1, strCurFont, "б") > 0 Then .Italic = True
                End With
                With Me.vfgThis
                    Set .Cell(flexcpFont, 7, .FixedCols, 7, .Cols - 1) = objFont
                End With
            Case "������ɫ"
                Me.lblRecordColor.ForeColor = Val("" & !�����ı�)
                With Me.vfgThis
                    .Cell(flexcpForeColor, 7, 1, 7, .Cols - 1) = Me.lblRecordColor.ForeColor
                End With
            Case "��Ч������"
                mlngPageRow = Val("" & !�����ı�)
            Case "����ʱ��ϲ�"
                chk���кϲ�.Value = Val("" & !�����ı�)
            Case "ʱ��������"
                chkHideTime.Value = Val("" & !�����ı�)
            End Select
            .MoveNext
        Loop
    End With
    chk����������.Value = (IIf(mlngPageRow = 0, 1, 0))
    
    Dim strPaper As String, blnHead As Boolean, blnFoot As Boolean
    gstrSQL = "Select ����||'-'||��� AS KEY,��ʽ,ҳü,ҳ�� From ����ҳ���ʽ Where ���� = 3 And ��� In (Select ҳ�� From �����ļ��б� Where Id = [1])"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    If Not rsTemp.EOF Then
        strPaper = "" & rsTemp!��ʽ:
        blnHead = ReadPageHead(rtbHead, rsTemp!Key)
        blnFoot = ReadPageFoot(rtbFoot, rsTemp!Key)
        cmdͬ��.Enabled = blnHead Or blnFoot
        picFoot.Tag = rsTemp!Key
        
        chkҳ��.Value = IIf(Val(NVL(rsTemp!ҳ��, 0)) > 0, 1, 0)
        If chkҳ��.Value = 1 Then optPageAlign(Val(NVL(rsTemp!ҳ��, 0)) - 1).Value = True
        '46251,������,2012-09-11,װ��ҳ�����λ��
        If CInt(Val(NVL(rsTemp!ҳü, 0))) > 0 And CInt(Val(NVL(rsTemp!ҳü, 0))) < 5 Then
            cboҳ��.ListIndex = CInt(Val(NVL(rsTemp!ҳü, 0))) - 1
        End If
    End If
    
    If UBound(Split(strPaper, ";")) >= 0 Then
        For lngCount = 0 To Me.cboPaperKind.ListCount - 1
            If Me.cboPaperKind.ItemData(lngCount) = Val(Split(strPaper, ";")(0)) Then Me.cboPaperKind.ListIndex = lngCount: Exit For
        Next
        If Me.cboPaperKind.ListIndex = 0 Then
            If UBound(Split(strPaper, ";")) >= 2 Then Me.txtHeight.Text = Round(Me.ScaleY(Val(Split(strPaper, ";")(2)), vbTwips, vbMillimeters), 2)
            If UBound(Split(strPaper, ";")) >= 3 Then Me.txtWidth.Text = Round(Me.ScaleY(Val(Split(strPaper, ";")(3)), vbTwips, vbMillimeters), 2)
        End If
    End If
    If UBound(Split(strPaper, ";")) >= 1 Then
        If Val(Split(strPaper, ";")(1)) = 2 Then
            Me.optOrient(1).Value = True
        Else
            Me.optOrient(0).Value = True
        End If
    End If
    If UBound(Split(strPaper, ";")) >= 4 Then Me.txtMarjin(2).Text = Round(Me.ScaleY(Val(Split(strPaper, ";")(4)), vbTwips, vbMillimeters), 2)
    If UBound(Split(strPaper, ";")) >= 5 Then Me.txtMarjin(3).Text = Round(Me.ScaleY(Val(Split(strPaper, ";")(5)), vbTwips, vbMillimeters), 2)
    If UBound(Split(strPaper, ";")) >= 6 Then Me.txtMarjin(0).Text = Round(Me.ScaleX(Val(Split(strPaper, ";")(6)), vbTwips, vbMillimeters), 2)
    If UBound(Split(strPaper, ";")) >= 7 Then Me.txtMarjin(1).Text = Round(Me.ScaleX(Val(Split(strPaper, ";")(7)), vbTwips, vbMillimeters), 2)
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select d.�������, d.�����ı�, d.Ҫ������, Nvl(d.�Ƿ���, 0) As �Ƿ���,d.Ҫ�ر�ʾ " & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '���ϱ�ǩ'" & _
        " Order By d.�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    With rsTemp
        Me.lstLabelUsed.Clear
        Do While Not .EOF
            Me.lstLabelUsed.AddItem !�����ı� & "{" & !Ҫ������ & "}"
            Me.lstLabelUsed.ItemData(Me.lstLabelUsed.NewIndex) = !�Ƿ���
            .MoveNext
        Loop
        If Me.lstLabelUsed.ListCount > 0 Then
            Me.lstLabelUsed.ListIndex = 0
            Me.cmdLabel(1).Enabled = True
            Me.chkLabelCrLf.Enabled = True
            Me.txtLabelPrefix.Enabled = True
        Else
            Me.cmdLabel(1).Enabled = False
            Me.chkLabelCrLf.Enabled = False: Me.chkLabelCrLf.Value = vbUnchecked
            Me.txtLabelPrefix.Enabled = False: Me.txtLabelPrefix.Text = ""
        End If
        Call cmdLabel_Click(2)
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select d.�������, d.�����д�, d.�����ı�" & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '��ͷ��Ԫ'" & _
        " Order By d.�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    With rsTemp
        Do While Not .EOF
            Me.vfgThis.TextMatrix(!�����д� + 2, !�������) = "" & !�����ı�
            .MoveNext
        Loop
        Call udHeadCol_Change
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    gstrSQL = "Select d.�������,d.������, d.��������, d.�����д�, d.�����ı�, d.Ҫ������, d.Ҫ�ص�λ,d.Ҫ�ر�ʾ,d.Ҫ��ֵ�� " & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '���м���'" & _
        " Order By d.�������, d.�����д�"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    strColRelation = ""
    With rsTemp
        Me.lstColumnUsed.Clear
        Do While Not .EOF
            Me.vfgThis.ColWidth(!�������) = Val(Split("" & !��������, "`")(0))
            If InStr(1, "" & !��������, "`") <> 0 Then
                vfgThis.Cell(flexcpAlignment, 6, !�������, 7, !�������) = Val(Split("" & !��������, "`")(1))
            Else
                vfgThis.Cell(flexcpAlignment, 6, !�������, 7, !�������) = flexAlignLeftCenter
            End If
            If Me.udColumnNo.Value <> !������� Then Me.udColumnNo.Value = !�������
            Me.lstColumnUsed.AddItem !�����ı� & "{" & !Ҫ������ & "}" & !Ҫ�ص�λ
            Me.lstColumnUsed.ItemData(lstColumnUsed.NewIndex) = zlCommFun.NVL(!Ҫ�ر�ʾ, 0)
            
            If IsNumeric(NVL(!������)) And Val(NVL(!������)) > 0 And Val(NVL(!������)) <> Val(NVL(!�������)) Then
                strColRelation = strColRelation & "|" & Val(NVL(!�������)) & "," & Val(NVL(!������))
            End If
            
            .MoveNext
            If .EOF Then
                vfgColRelation.Tag = "OK"
                Call cmdColumn_Click(2)
                strColRelation = Mid(strColRelation, 2)
                '��ʼ��ʱ����֮ǰ���õĻ����ж��չ�ϵ
                If strColRelation <> "" Then
                    arrColName = Split(strColRelation, "|")
                    With vfgColRelation
                        If .ColComboList(.ColIndex("��������")) <> "" Then
                            arrTmp = Split(.ColComboList(.ColIndex("��������")), "|#")
                            For lngRow = .FixedRows To .Rows - 1
                                If Val(.TextMatrix(lngRow, .ColIndex("�����к�"))) > 0 Then
                                    For i = 0 To UBound(arrColName)
                                        If Val(Split(arrColName(i), ",")(0)) = Val(.TextMatrix(lngRow, .ColIndex("�����к�"))) Then
                                            For j = 0 To UBound(arrTmp)
                                                If InStr(1, arrTmp(j), ";") > 0 And Val(Split(arrColName(i), ",")(1)) > 0 Then
                                                    If Val(Split(arrColName(i), ",")(1)) = Val(Mid(arrTmp(j), 1, InStr(1, arrTmp(j), ";") - 1)) Then
                                                        .TextMatrix(lngRow, .ColIndex("�����к�")) = Mid(arrTmp(j), 1, InStr(1, arrTmp(j), ";") - 1)
                                                        .TextMatrix(lngRow, .ColIndex("��������")) = Mid(arrTmp(j), InStr(1, arrTmp(j), ";") + 1)
                                                        
                                                    End If
                                                End If
                                            Next j
                                        End If
                                    Next i
                                End If
                            Next lngRow
                        End If
                    End With
                End If
            ElseIf Me.udColumnNo.Value <> !������� Then
                Call cmdColumn_Click(2)
            End If
        Loop
        .Filter = " �����д� > 1"
        .Sort = "�������"
        strColNumber = ""
        Do While Not .EOF
            If Not InStr(1, strColNumber & ",", "," & !������� & ",") > 0 Then
            strColNumber = strColNumber & "," & !�������
            End If
            .MoveNext
        Loop
        .Filter = " �����д� = 1"
        .Sort = "�������"
        mstrColDoctor = ""
        vsfColDoctor.Tag = "OK"
        strColName = " |����|ҽ������|�ܸ�����|ִ��Ƶ��|��ҩĿ��|��ҩ����|ҽ������|��ʼִ��ʱ��|��ҩ;��"
        Do While Not .EOF
            If Not InStr(1, ",����,ʱ��,��ʿ,ǩ����,ǩ��ʱ��,", "," & !Ҫ������ & ",") > 0 And Not InStr(strColNumber & ",", "," & !������� & ",") > 0 Then
                vsfColDoctor.Rows = vsfColDoctor.Rows + 1
                vsfColDoctor.TextMatrix(vsfColDoctor.Rows - 1, vsfColDoctor.ColIndex("��Ŀ�к�")) = Val(NVL(!�������))
                vsfColDoctor.TextMatrix(vsfColDoctor.Rows - 1, vsfColDoctor.ColIndex("��Ŀ����")) = NVL(!Ҫ������)
                If Not NVL(!Ҫ��ֵ��) = "" Then mstrColDoctor = mstrColDoctor & "|" & NVL(!Ҫ������) & ";" & NVL(!Ҫ��ֵ��)
                vsfColDoctor.ComboList = strColName
            End If
            .MoveNext
        Loop
        Me.udColumnNo.Value = Me.vfgThis.Col
    End With
    
    mstrColDoctor = Mid(mstrColDoctor, 2)
    Call cmdColumn_Click(2)
    '����ʱ��
    '------------------------------------------------------------------------------------------------------------------
    Dim aryTmp As Variant
    
    
    gstrSQL = "Select d.�������, d.��������, d.�����д�, d.�����ı� " & _
        " From �����ļ��ṹ d, �����ļ��ṹ p" & _
        " Where p.Id = d.��id And p.�ļ�id = [1] And p.�������� = 1 And p.�����ı� = '����ʱ��'" & _
        " Order By d.�������, d.�����д�"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlngFileID)
    With vsf
        If rsTemp.BOF = False Then
            Do While Not rsTemp.EOF
                strTemp = zlCommFun.NVL(rsTemp!�����ı�)
                If strTemp <> "" Then
                    
                    aryTmp = Split(strTemp, ",")
                    
                    If UBound(aryTmp) >= 2 Then
                        If .TextMatrix(.Rows - 1, 1) <> "" And .TextMatrix(.Rows - 1, 2) <> "" And .TextMatrix(.Rows - 1, 3) <> "" Then .Rows = .Rows + 1
                        
                        .TextMatrix(.Rows - 1, 1) = Trim(aryTmp(0))
                        .TextMatrix(.Rows - 1, 2) = Trim(aryTmp(1))
                        .TextMatrix(.Rows - 1, 3) = Trim(aryTmp(2))
                    End If
                End If
                rsTemp.MoveNext
            Loop
            mclsVsf.AppendRows = True
        End If
    End With
    
    '�ٰ��кϲ�
    For lngCount = 1 To vfgThis.Cols - 1
        vfgThis.MergeCol(lngCount) = True
    Next
    vfgThis.AutoSize 0, vfgThis.Cols - 1
    Call cmdLabel_Click(2)
    RefreshData = True
    
    Exit Function

errHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog

    
End Function

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
        .SetIconSize True, 24, 24
        .LargeIcons = True
    End With

    '------------------------------------------------------------------------------------------------------------------
    '�˵�����
    cbsThis.ActiveMenuBar.Title = "�˵�"
    cbsThis.ActiveMenuBar.EnableDocking (xtpFlagAlignTop)
    
    
     '�����
    With cbsThis.KeyBindings

        .Add FCONTROL, Asc("S"), conMenu_Edit_Transf_Save
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_F2, conMenu_Edit_Transf_Save
    End With
    
    '------------------------------------------------------------------------------------------------------------------
    '����������
    Set cbrToolBar = cbsThis.Add("��׼", xtpBarTop)
    cbrToolBar.ShowTextBelowIcons = False
    cbrToolBar.EnableDocking xtpFlagStretched
    With cbrToolBar.Controls

        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_SaveExit, "���沢�˳�"): cbrControl.BeginGroup = True
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Transf_Save, "����"): cbrControl.ToolTipText = "�����Ѹ��ĵ�����(Ctrl+S,F2)"
        Set cbrControl = .Add(xtpControlButton, conMenu_Edit_Transf_Cancle, "�ָ�"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "�ָ����ϴα���ʱ������״̬"
        Set cbrControl = .Add(xtpControlButton, conMenu_Help_Help, "����"): cbrControl.BeginGroup = True: cbrControl.ToolTipText = "����(F1)"
        Set cbrControl = .Add(xtpControlButton, conMenu_File_Exit, "�˳�"): cbrControl.ToolTipText = "�˳���ǰ����ƴ���(Esc)"

    End With
        
    For Each cbrControl In cbrToolBar.Controls
        If cbrControl.Type <> xtpControlCustom And cbrControl.Type <> xtpControlLabel Then
            cbrControl.STYLE = xtpButtonIconAndCaption
        End If
    Next
    
     '�����
    With cbsThis.KeyBindings
        .Add FCONTROL, Asc("S"), conMenu_Edit_Transf_Save
        .Add 0, VK_F1, conMenu_Help_Help
        .Add 0, VK_ESCAPE, conMenu_File_Exit
        
        .Add 0, vbKeyF2, conMenu_Edit_Transf_Save
    End With
    
    InitMenuBar = True
    
    Exit Function
errHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub zlFontSet(strTitle As String, strFont As String)
    With Me.dlgThis
        .flags = &H3 Or &H400 Or &H200 Or &H10000
        .DialogTitle = strTitle
        .FontName = Split(strFont, ",")(0)
        .FontSize = Val(Split(strFont, ",")(1))
        If InStr(1, strFont, "��") > 0 Then
            .FontBold = True
        Else
            .FontBold = False
        End If
        If InStr(1, strFont, "б") > 0 Then
            .FontItalic = True
        Else
            .FontItalic = False
        End If
        Err = 0: On Error Resume Next
        .ShowFont
        .flags = 0
        If Err.Number <> 0 Then Exit Sub
        strFont = .FontName & "," & .FontSize
        If .FontBold Or .FontItalic Then
            strFont = strFont & "," & IIf(.FontBold, "��", "") & IIf(.FontItalic, "б", "")
        End If
    End With
End Sub

Private Sub zlColorSet(strTitle As String, lngColor As Long)
    With Me.dlgThis
        .DialogTitle = strTitle
        .COLOR = lngColor
        Err = 0: On Error Resume Next
        .ShowColor
        If Err.Number <> 0 Then Exit Sub
        lngColor = .COLOR
    End With
End Sub

Private Sub cboPaperKind_Click()
    Dim intOrientation As Integer
    If Me.optOrient(0).Value = True Then
        intOrientation = 1
    Else
        intOrientation = 2
    End If
    '���ô�ӡ��ֽ�ŷ���̶�Ϊ1������ʱ����������
    Printer.Orientation = 1
    If Me.cboPaperKind.ListIndex <= 0 Then
        Me.txtWidth.Enabled = True: Me.udWidth.Enabled = True
        Me.txtHeight.Enabled = True: Me.udHeight.Enabled = True
        Me.optOrient(0).Value = True
        Me.optOrient(0).Enabled = True: Me.optOrient(1).Enabled = True
        If intOrientation = 1 Then
            Me.optOrient(0).Value = True
        Else
            Me.optOrient(1).Value = True
        End If
    Else
        Me.txtWidth.Enabled = False: Me.udWidth.Enabled = False
        Me.txtHeight.Enabled = False: Me.udHeight.Enabled = False
        Me.optOrient(0).Enabled = True: Me.optOrient(1).Enabled = True
        Err = 0: On Error Resume Next
        Printer.PaperSize = Me.cboPaperKind.ItemData(Me.cboPaperKind.ListIndex)
        Me.txtWidth.Text = Me.ScaleX(Printer.Width, vbTwips, vbMillimeters)
        Me.txtHeight.Text = Me.ScaleY(Printer.Height, vbTwips, vbMillimeters)
        If intOrientation = 1 Then
            Me.optOrient(0).Value = True
        Else
            Me.optOrient(1).Value = True
        End If
    End If
    DataChanged = True
End Sub

Private Sub cbsThis_Execute(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Edit_SaveExit
        
        If SaveData Then
            DataChanged = False
            Unload Me
        End If
        
    Case conMenu_Edit_Transf_Save
        
        If SaveData Then
            DataChanged = False
        End If
        
    Case conMenu_Edit_Transf_Cancle
                
        Call RefreshData
        DataChanged = False
        
    Case conMenu_File_Exit
        
        mblnOk = False
        Unload Me
        
    Case conMenu_Help_Help
        
        Call ShowHelp(App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100))
        
    End Select
End Sub

Private Sub cbsThis_GetClientBordersWidth(Left As Long, Top As Long, Right As Long, Bottom As Long)
    If stbThis.Visible Then Bottom = stbThis.Height
End Sub

Private Sub cbsThis_Update(ByVal Control As XtremeCommandBars.ICommandBarControl)
    Select Case Control.ID
    Case conMenu_Edit_SaveExit
        
        Control.Enabled = DataChanged
        
    Case conMenu_Edit_Transf_Save
        
        Control.Enabled = DataChanged
        
    Case conMenu_Edit_Transf_Cancle
                
        Control.Enabled = DataChanged
        
    End Select
End Sub

Private Sub chk_Click()
    With Me.lstColumnUsed
        If .ListIndex = -1 Then Exit Sub
        .ItemData(.ListIndex) = chk.Value
    End With
    
    vfgThis.Cell(flexcpData, 6, udColumnNo.Value, 6, udColumnNo.Value) = Replace(vfgThis.Cell(flexcpData, 6, udColumnNo.Value, 6, udColumnNo.Value), IIf(chk.Value = 0, "`1", "`0"), "`" & lstColumnUsed.ItemData(lstColumnUsed.ListIndex))
    mblnChanged = True
End Sub

Private Sub chkLabelCrLf_Click()
    With Me.lstLabelUsed
        If .ListIndex = -1 Then Exit Sub
        .ItemData(.ListIndex) = Me.chkLabelCrLf.Value
    End With
End Sub

Private Sub chkҳ��_Click()
    optPageAlign(0).Enabled = chkҳ��.Value
    optPageAlign(1).Enabled = chkҳ��.Value
    optPageAlign(2).Enabled = chkҳ��.Value
    lblҳ��.Enabled = chkҳ��.Value
    cboҳ��.Enabled = chkҳ��.Value
    If chkҳ��.Value = 1 Then
        If Not optPageAlign(0).Value Then
            If Not optPageAlign(1).Value Then
                If Not optPageAlign(2).Value Then optPageAlign(0).Value = True
            End If
        End If
        If cboҳ��.ListIndex = -1 And cboҳ��.ListCount > 0 Then cboҳ��.ListIndex = cboҳ��.ListCount - 1
    End If
    mblnChanged = True
End Sub

Private Sub cmdColumn_Click(Index As Integer)
    Dim strTmp As String, strTmp1 As String
    Dim blnSplit As Boolean                         '�����Ŀʱ���,���ǰ�����Ŀ�޺�׺�Һ�һ����Ŀ��ǰ׺,��blnSplit=False,���������
    Dim intType As Integer, intFace As Integer, intLen As Integer      '��Ŀ����
    Dim strFaces As String                          '������Ŀ,ֻ����¼����Ŀ0�뵥ѡ��Ŀ4;������������Ŀ,ֻ����¼����Ŀ0
    Dim strName As String                           '��Ŀ����
    Dim intCount As Integer, arrTmp() As String
    Dim arrCol(), arrColValue()
    
    With Me.lstColumnUsed
        Select Case Index
        Case 0
            If Me.lstColumnItems.ListIndex = -1 Then Exit Sub
            .AddItem "{" & Me.lstColumnItems.List(Me.lstColumnItems.ListIndex) & "}"
            .ListIndex = .NewIndex
            Me.cmdColumn(1).Enabled = True
            Me.txtColumnPrefix.Enabled = True
            Me.txtColumnPostfix.Enabled = True
            chk.Enabled = True
        Case 1
            If .ListIndex = -1 Then Exit Sub
            .RemoveItem .ListIndex
            If .ListCount > 0 Then
                .ListIndex = 0
            Else
                .ListIndex = -1
                Me.cmdColumn(1).Enabled = False
                Me.txtColumnPrefix.Enabled = False: Me.txtColumnPrefix.Text = ""
                Me.txtColumnPostfix.Enabled = False: Me.txtColumnPostfix.Text = ""
            End If
            
            chk.Enabled = True
        Case 2
            '��һ�а�2����Ŀʱ����Ŀ֮��������ǰ׺/���׺���ż�������
            '��һ�а󶨶����Ŀʱ����Ŀ���ͱ�����¼������Ŀ
            '��ѡ���ѡ��Ŀ������������Ŀһ�����ĳ��
            'ϵͳ�̶�����Ŀ����ǩ���ˣ����ڣ�ʱ��ȣ�һ��ֻ�ܰ�һ��
            strTemp = ""
            strTmp = ""
            strTmp1 = ""
            
            If .ListCount = 1 Then
                strFaces = "0,1,2,3,4,5"
            ElseIf .ListCount = 2 Then
                strFaces = "0,4,5"
            Else
                strFaces = "0"
            End If
            
            '��ǰ��
            For lngCount = 0 To .ListCount - 1
                strName = Mid(.List(lngCount), InStr(1, .List(lngCount), "{"))
                strName = Mid(strName, 1, InStr(1, strName, "}"))
                strTmp1 = strTmp1 & "'" & strName
            Next lngCount
            
            '������
            arrCol = Array()
            arrColValue = Array()
            For lngCount = 1 To vfgThis.Cols - 1
                If lngCount <> udColumnNo.Value And Trim(vfgThis.Cell(flexcpData, 6, lngCount, 6, lngCount)) <> "" Then
                    arrTmp = Split(Trim(vfgThis.Cell(flexcpData, 6, lngCount, 6, lngCount)), " ")
                    For intCount = 0 To UBound(arrTmp)
                        If InStr(1, arrTmp(intCount), "`") > 0 Then
                            strName = CStr(Split(arrTmp(intCount), "`")(0))
                        Else
                            strName = CStr(arrTmp(intCount))
                        End If
                        strName = Mid(strName, InStr(1, strName, "{"))
                        strName = Mid(strName, 1, InStr(1, strName, "}"))
                        If strName <> "{}" Then
                            ReDim Preserve arrCol(UBound(arrCol) + 1)
                            arrCol(UBound(arrCol)) = lngCount
                            ReDim Preserve arrColValue(UBound(arrColValue) + 1)
                            arrColValue(UBound(arrColValue)) = strName
                        End If
                    Next intCount
                End If
            Next lngCount
            
            For lngCount = 0 To .ListCount - 1
                strTemp = strTemp & Space(1) & .List(lngCount) & "`" & .ItemData(lngCount)
                strTmp = strTmp & Space(1) & .List(lngCount)
                strName = Mid(.List(lngCount), InStr(1, .List(lngCount), "{") + 1)
                strName = Mid(strName, 1, InStr(1, strName, "}") - 1)
                
                'ÿһ�а󶨵���Ŀ�����ظ�
                If .ListCount > 1 Then
                    If UBound(Split(strTmp1, "'{" & strName & "}")) > 1 Then
                        MsgBox "��һ�а󶨶����Ŀʱ����Ŀ���Ʋ����ظ���", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
                
                '�����е���Ŀ�Ƿ�����������Ŀ�ظ�
                For intCount = 0 To UBound(arrCol)
                    If CStr(arrColValue(intCount)) = "{" & strName & "}" And Trim(strName) <> "" Then
                        MsgBox "{" & strName & "}�Ѿ��ڵ�" & CInt(arrCol(intCount)) & "����,���飡", vbInformation, gstrSysName
                        Exit Sub
                    End If
                Next intCount
                
                If lngCount > 0 Then
                    '�����Ŀ���Ƿ���ڷָ���
                    If Not blnSplit Then
                        If Trim(Split(.List(lngCount), "{")(0) = "") Then
                            MsgBox "��һ�а󶨶����Ŀʱ����Ŀ֮�����Ҫ����ǰ׺���׺���ż������֣�", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                    '�����Ŀ�����Ƿ�һ��
                    rsItems.Filter = "��Ŀ����='" & strName & "'"
                    If rsItems.RecordCount <> 0 Then
                        If Not (intType = rsItems!��Ŀ���� And intFace = rsItems!��Ŀ��ʾ) Then
                            MsgBox "��һ�а󶨶����Ŀʱ����Ŀ�����ͱ���һ�£�", vbInformation, gstrSysName
                            Exit Sub
                        End If
                        
                        If rsItems!��Ŀ���� = 1 And rsItems!��Ŀ��ʾ = 0 And NVL(rsItems!��Ŀ����, 1) > 3 Then
                            MsgBox "һ�����ֻ�ܰ���������С�ڻ����3���ı���Ŀ��", vbInformation, gstrSysName
                            Exit Sub
                        End If
                        
                    Else
                        If InStr(1, "����,ʱ��", strName) = 0 Then
                            MsgBox "�̶���Ŀ������������Ŀ����һ��", vbInformation, gstrSysName
                            Exit Sub
                        End If
                    End If
                Else
                    'ֻ��Ҫȡ��һ����Ŀ�����Լ���
                    rsItems.Filter = "��Ŀ����='" & strName & "'"
                    If rsItems.RecordCount <> 0 Then
                        intType = rsItems!��Ŀ����
                        intFace = rsItems!��Ŀ��ʾ
                        intLen = NVL(rsItems!��Ŀ����, 1)
                        If .ListCount > 1 Then
                            If intType = 1 And intFace = 0 And intLen > 3 Then
                                MsgBox "һ�����ֻ�ܰ���������С�ڻ����3���ı���Ŀ��", vbInformation, gstrSysName
                                Exit Sub
                            End If
                            If intFace = 3 Then
                                MsgBox "��ѡ��ֻ�ܵ����󶨣�", vbInformation, gstrSysName
                                Exit Sub
                            End If
                        End If
                        'һ��������������������<=3���ı���Ŀ
                        If .ListCount > 2 Then
                            If intType = 1 And intFace = 0 Then
                                MsgBox "һ�����ֻ�ܰ���������С�ڻ����3���ı���Ŀ��", vbInformation, gstrSysName
                                Exit Sub
                            End If
                        End If
                    Else
                        If .ListCount > 1 Then
                            If InStr(1, "����,ʱ��", strName) = 0 Then
                                MsgBox "�̶���Ŀ������������Ŀ����һ��", vbInformation, gstrSysName
                                Exit Sub
                            End If
                        End If
                    End If
                End If
                '�̶�ȡ��һ����Ŀ�ĺ�׺
                blnSplit = Trim(Split(.List(lngCount), "}")(1) <> "")
                '�����Ŀ��ʾ(������Ŀ,ֻ����¼����Ŀ0�뵥ѡ��Ŀ4;������������Ŀ,ֻ����¼����Ŀ0)
'                If rsItems.RecordCount <> 0 Then
'                    If InStr(1, strFaces, rsItems!��Ŀ��ʾ) = 0 Then
'                        If .ListCount > 2 Then
'                            MsgBox "��һ�а󶨶����Ŀʱ��ֻ��ѡ��¼���͵���Ŀ��", vbInformation, gstrSysName
'                        Else
'                            MsgBox "��һ�а�������Ŀʱ��ֻ��ѡ��¼���ͻ�ѡ���͵���Ŀ��", vbInformation, gstrSysName
'                        End If
'                        Exit Sub
'                    End If
'                End If
            Next
            strTemp = Trim(strTemp)
            strTmp = Trim(strTmp)
            rsItems.Filter = 0
            
            With vfgThis
                .TextMatrix(6, Me.udColumnNo.Value) = strTmp
                '���ݶ��뷽ʽ����������
                Select Case .Cell(flexcpAlignment, 6, Me.udColumnNo.Value)
                Case 4
                    .TextMatrix(7, Me.udColumnNo.Value) = " " & strTmp & " "
                Case 7
                    .TextMatrix(7, Me.udColumnNo.Value) = " " & strTmp
                Case Else
                    .TextMatrix(7, Me.udColumnNo.Value) = strTmp & " "
                End Select
                
                .Cell(flexcpData, 6, udColumnNo.Value, 6, udColumnNo.Value) = strTemp
                .Cell(flexcpData, 7, udColumnNo.Value, 7, udColumnNo.Value) = strTemp & " "
                
            End With
            
            '���й�ϵ���ݵ�����
            Call InstallColRelation
            DataChanged = True
        End Select
    End With
End Sub

Private Sub InstallColRelation()
'����:װ�ػ������������й�ϵ
'-------------------------------------
    Dim strTmp As String
    Dim intType As Integer, intFace As Integer, intLen As Integer      '��Ŀ����
    Dim strName As String  '��Ŀ����
    Dim intCount As Integer, i As Integer
    Dim arrColName, arrTmp
    Dim arrColNo, strColName As String, j As Integer
    
    If vfgColRelation.Tag <> "OK" Then Exit Sub
    
    With vfgColRelation
        '���Ȼ�ȡ֮ǰ�趨���еĹ�ϵ
        arrColName = Array()
        For lngCount = .FixedRows To .Rows - 1
            If Val(.TextMatrix(lngCount, .ColIndex("�����к�"))) > 0 And Val(.TextMatrix(lngCount, .ColIndex("�����к�"))) > 0 And Val(.TextMatrix(lngCount, .ColIndex("�����к�"))) <> Val(.TextMatrix(lngCount, .ColIndex("�����к�"))) Then
                ReDim Preserve arrColName(UBound(arrColName) + 1)
                arrColName(UBound(arrColName)) = .TextMatrix(lngCount, .ColIndex("��������")) & "'" & .TextMatrix(lngCount, .ColIndex("��������"))
            End If
        Next lngCount
        
        '������Ŀ�������еĹ�ϵ�趨���������㣺��ֻ����һ����Ŀ,�����п�Ϊ�ı���ѡ���ı���Ŀ����С��100(��Ϊ��¼���������ݴ����ı�����>=100��Ϊ�Ǵ��ı���Ŀ)
        .Rows = .FixedRows + 1
        .Cell(flexcpText, .FixedRows, 0, .Rows - 1, .Cols - 1) = ""
        .Cell(flexcpData, .FixedRows, 0, .Rows - 1, .Cols - 1) = ""
        intCount = .Rows
        strTmp = "": .ColComboList(.ColIndex("��������")) = ""
        For lngCount = 1 To vfgThis.Cols - 1
            If Trim(vfgThis.Cell(flexcpData, 6, lngCount, 6, lngCount)) <> "" Then
                arrTmp = Split(Trim(vfgThis.Cell(flexcpData, 6, lngCount, 6, lngCount)), " ")
                If UBound(arrTmp) = 0 Then
                    If InStr(1, arrTmp(0), "`") > 0 Then
                        strName = CStr(Split(arrTmp(0), "`")(0))
                    Else
                        strName = CStr(arrTmp(0))
                    End If
                    strName = Mid(strName, InStr(1, strName, "{") + 1)
                    strName = Mid(strName, 1, InStr(1, strName, "}") - 1)
                    If strName <> "" Then
                        rsItems.Filter = "��Ŀ����='" & strName & "'"
                        If rsItems.RecordCount > 0 Then
                            intType = rsItems!��Ŀ����
                            intFace = rsItems!��Ŀ��ʾ
                            intLen = NVL(rsItems!��Ŀ����, 1)
                            If intType = 0 And intFace = 4 Then '������Ŀȷ��
                                If intCount > .Rows Then .Rows = .Rows + 1
                                .TextMatrix(.Rows - 1, .ColIndex("�����к�")) = lngCount
                                .TextMatrix(.Rows - 1, .ColIndex("��������")) = strName
                                intCount = intCount + 1
                            '��������Ŀȷ��
                            ElseIf intType = 1 And intFace = 0 And intLen < 100 Then '�ı���Ŀ
                                strTmp = strTmp & "|#" & lngCount & ";" & strName
                            ElseIf intType = 1 And intFace = 2 Then '��ѡ��Ŀ
                                strTmp = strTmp & "|#" & lngCount & ";" & strName
                            End If
                        End If
                    End If
                End If
            End If
        Next lngCount
        If strTmp <> "" Then
            arrTmp = Split(Mid(strTmp, 3), "|#")
            .ColComboList(.ColIndex("��������")) = "#; " & strTmp
            '�ָ�ԭ�а󶨵�������
            If UBound(arrColName) >= 0 Then
                For lngCount = .FixedRows To .Rows - 1
                    If Val(.TextMatrix(lngCount, .ColIndex("�����к�"))) > 0 Then
                        For intCount = 0 To UBound(arrColName)
                            If .TextMatrix(lngCount, .ColIndex("��������")) = Split(arrColName(intCount), "'")(0) Then
                                For i = 0 To UBound(arrTmp)
                                    If Split(arrColName(intCount), "'")(1) = Mid(arrTmp(i), InStr(1, arrTmp(i), ";") + 1) Then
                                        .TextMatrix(lngCount, .ColIndex("�����к�")) = Mid(arrTmp(i), 1, InStr(1, arrTmp(i), ";") - 1)
                                        .TextMatrix(lngCount, .ColIndex("��������")) = Split(arrColName(intCount), "'")(1)
                                    End If
                                Next i
                            End If
                        Next intCount
                    End If
                Next lngCount
            End If
        End If
    End With
    
    '����ҽ���󶨱��
    With vsfColDoctor
        .Rows = .FixedRows + 1
        .Cell(flexcpText, .FixedRows, 0, .Rows - 1, .Cols - 1) = ""
        .Cell(flexcpData, .FixedRows, 0, .Rows - 1, .Cols - 1) = ""
        If vsfColDoctor.Tag = "OK" Then
            arrColNo = Split(mstrColDoctor, "|")
            
            intCount = .Rows
            For lngCount = 1 To vfgThis.Cols - 1
                If Trim(vfgThis.Cell(flexcpData, 6, lngCount, 6, lngCount)) <> "" Then
                    arrTmp = Split(Trim(vfgThis.Cell(flexcpData, 6, lngCount, 6, lngCount)), " ")
                    If UBound(arrTmp) = 0 Then
                        If InStr(1, arrTmp(0), "`") > 0 Then
                            strName = CStr(Split(arrTmp(0), "`")(0))
                        Else
                            strName = CStr(arrTmp(0))
                        End If
                        strName = Mid(strName, InStr(1, strName, "{") + 1)
                        strName = Mid(strName, 1, InStr(1, strName, "}") - 1)
                        If strName <> "" Then
                            rsItems.Filter = "��Ŀ����='" & strName & "'"
                            If rsItems.RecordCount > 0 Then
                                intType = rsItems!��Ŀ����
                                intFace = rsItems!��Ŀ��ʾ
                                If intCount > .Rows Then .Rows = .Rows + 1
                                vsfColDoctor.TextMatrix(vsfColDoctor.Rows - 1, vsfColDoctor.ColIndex("��Ŀ�к�")) = lngCount
                                vsfColDoctor.TextMatrix(vsfColDoctor.Rows - 1, vsfColDoctor.ColIndex("��Ŀ����")) = strName
                                vsfColDoctor.ComboList = " |����|ҽ������|�ܸ�����|ִ��Ƶ��|��ҩĿ��|��ҩ����|ҽ������|��ʼִ��ʱ��|��ҩ;��"
                                intCount = intCount + 1
                                
                            End If
                        End If
                    End If
                End If
            Next lngCount
            '�����Ѿ��󶨵���
            For i = 0 To UBound(arrColNo)
                strColName = Split(arrColNo(i), ";")(0)
                strName = NVL(Split(arrColNo(i), ";")(1))
                For j = vsfColDoctor.FixedRows To vsfColDoctor.Rows - 1
                    If vsfColDoctor.TextMatrix(j, .ColIndex("��Ŀ����")) = strColName Then vsfColDoctor.TextMatrix(j, .ColIndex("��������")) = strName: Exit For
                Next
            Next
        End If
    End With
End Sub

Private Sub cmdLabel_Click(Index As Integer)
    With Me.lstLabelUsed
        Select Case Index
        Case 0
            If Me.lstLabelItems.ListIndex = -1 Then Exit Sub
            .AddItem Me.lstLabelItems.List(Me.lstLabelItems.ListIndex) & "��{" & Me.lstLabelItems.List(Me.lstLabelItems.ListIndex) & "}"
            .ListIndex = .NewIndex
            Me.cmdLabel(1).Enabled = True
            Me.chkLabelCrLf.Enabled = True
            Me.txtLabelPrefix.Enabled = True
        Case 1
            If .ListIndex = -1 Then Exit Sub
            .RemoveItem .ListIndex
            If .ListCount > 0 Then
                .ListIndex = 0
            Else
                .ListIndex = -1
                Me.cmdLabel(1).Enabled = False
                Me.chkLabelCrLf.Enabled = False: Me.chkLabelCrLf.Value = vbUnchecked
                Me.txtLabelPrefix.Enabled = False: Me.txtLabelPrefix.Text = ""
            End If
        Case 2
            Dim intCrLf As Integer
            intCrLf = 0
            strTemp = ""
            For lngCount = 0 To .ListCount - 1
                If .ItemData(lngCount) <> 0 Then intCrLf = intCrLf + 1
                strTemp = strTemp & Space(1) & IIf(.ItemData(lngCount) = 0, "", vbCrLf) & .List(lngCount)
            Next
            strTemp = Trim(strTemp)
            For lngCount = Me.vfgThis.FixedCols To Me.vfgThis.Cols - 1
                Me.vfgThis.TextMatrix(2, lngCount) = strTemp
            Next
            Me.vfgThis.ROWHEIGHT(2) = Me.vfgThis.FontSize * 20 * (intCrLf + 1) + 150
            
            DataChanged = True
        End Select
    End With
End Sub

Private Function CheckData() As Boolean
    Dim arrData
    Dim intType As Integer, intFace As Integer, intLen As Integer                  '������Ŀ�ı�ʾ��ʽһ��������
    Dim bln��ʿ As Boolean, bln���� As Boolean, blnʱ�� As Boolean
    Dim StrText As String, strItem As String
    Dim lngCol As Long, lngCount As Long
    Dim intDo As Integer, intHead As Integer
    Dim intRow As Integer
    
    'ÿ�ֻ����¼��������Ҫ��һ�а󶨻�ʿ����
    lngCount = vfgThis.Cols - 1
    For lngCol = 1 To lngCount
        If InStr(1, ",{��ʿ},{ǩ����},", "," & vfgThis.TextMatrix(6, lngCol) & ",") <> 0 Then
            bln��ʿ = True
        End If
    Next
    If Not bln��ʿ Then
        MsgBox "������һ�а󶨻�ʿ��Ŀ��ǩ������Ŀ��", vbInformation, gstrSysName
        Exit Function
    End If
    If vfgThis.TextMatrix(6, 1) <> "{����}" Then
        MsgBox "��һ�б����������Ŀ��", vbInformation, gstrSysName
        Exit Function
    End If
    If vfgThis.TextMatrix(6, 2) <> "{ʱ��}" Then
        MsgBox "�ڶ��б����ʱ����Ŀ��", vbInformation, gstrSysName
        Exit Function
    End If
    
    'ֻ�а�������Ŀ���У���Ŀ�ķָ���Ҳ��/������Ŀ���ͱ�����ͬ,ֻ����¼�롢ѡ��������������������и�ʽΪ�Խ��ߣ��������������öԽ��߱�ʾ���ڼ�д��
    'ֻȡ���������ͷ��3�б�ͷȡ5��2�б�ͷȡ4��1�б�ͷȡ3)
    If optTabTiers(0).Value Then
        intHead = 3
    ElseIf optTabTiers(1).Value Then
        intHead = 4
    Else
        intHead = 5
    End If
    For lngCol = 1 To lngCount
        If vfgThis.Cell(flexcpData, 6, lngCol) <> "" Then
            StrText = Val(Split(vfgThis.Cell(flexcpData, 6, lngCol), "`")(1))
            
            If StrText = 1 Then
                '��ʽ��{A}{B}����}�ֽ⣬2�зֽ�����������=2
                StrText = vfgThis.TextMatrix(6, lngCol)
                If UBound(Split(StrText, "}")) <> 2 Then
                    If StrText <> "{����}" Then
                        MsgBox "��" & lngCol & "�� " & vfgThis.TextMatrix(intHead, lngCol) & " �����������жԽ��ߣ�" & vbCrLf & "[ֻ�������С���������Ŀ���в����������жԽ���]", vbInformation, gstrSysName
                        Exit Function
                    Else
                        GoTo ntloop
                    End If
                End If
                
                '������Ŀ�ķָ���Ҳ������/
                If Trim(Mid(StrText, InStr(1, StrText, "}") + 1, InStr(InStr(1, StrText, "{") + 1, StrText, "{") - InStr(1, StrText, "}") - 1)) <> "/" Then
                    MsgBox "��" & lngCol & "�� " & vfgThis.TextMatrix(intHead, lngCol) & " �������жԽ��ߣ�Ҫ��󶨵���Ŀ��ʽӦ����:A/B", vbInformation, gstrSysName
                    Exit Function
                End If
                
                '������Ŀ����Ŀ���ͱ���һ��
                For intDo = 0 To 1
                    strItem = Trim(GetItemName(StrText, intDo))
                    rsItems.Filter = "��Ŀ����='" & strItem & "'"
                    '¼���ʱ�������,ϵͳ��������,ʱ��Ȳ�������̶����,����,��ʱ�������Ҳ��������
                    If rsItems.RecordCount = 0 Then
                        MsgBox "��Ŀ:" & strItem & "�Ѹ������Ѿ���ɾ����", vbInformation, gstrSysName
                        Exit Function
                    End If
                    If intDo > 0 Then
                        If Not (intFace = rsItems!��Ŀ��ʾ And intType = rsItems!��Ŀ����) Then
                            MsgBox "��" & lngCol & "�� " & vfgThis.TextMatrix(intHead, lngCol) & " �󶨵�������Ŀ�ı༭��ʽ����һ�£�", vbInformation, gstrSysName
                            Exit Function
                        End If
                        If rsItems!��Ŀ���� = 1 And rsItems!��Ŀ��ʾ = 0 And NVL(rsItems!��Ŀ����, 1) > 3 Then
                            MsgBox "��" & lngCol & "�� " & vfgThis.TextMatrix(intHead, lngCol) & " �󶨵������ı���Ŀ���Ȳ��ܴ���3��", vbInformation, gstrSysName
                            Exit Function
                        End If
                    Else
                        intFace = rsItems!��Ŀ��ʾ
                        intType = rsItems!��Ŀ����
                        intLen = NVL(rsItems!��Ŀ����, 1)
                        If InStr(1, "0,2,4,5", intFace) = 0 Then
                            MsgBox "��" & lngCol & "�� " & vfgThis.TextMatrix(intHead, lngCol) & " �󶨵�������Ŀ���붼����ֵ�͡�ѡ�����ѡ�������򳤶�С�ڵ���3���ı���Ŀ��", vbInformation, gstrSysName
                            Exit Function
                        End If
                        If intType = 1 And intFace = 0 And intLen > 3 Then
                            MsgBox "��" & lngCol & "�� " & vfgThis.TextMatrix(intHead, lngCol) & " �󶨵������ı���Ŀ���Ȳ��ܴ���3��", vbInformation, gstrSysName
                            Exit Function
                        End If
                    End If
                Next
            End If
        End If
ntloop:
    Next
    
    '�������й�ϵ
    With vfgColRelation
        For intRow = .FixedRows To .Rows - 1
            If Val(.TextMatrix(intRow, .ColIndex("�����к�"))) > 0 And Val(.TextMatrix(intRow, .ColIndex("�����к�"))) > 0 Then
                If Val(.TextMatrix(intRow, .ColIndex("�����к�"))) = Val(.TextMatrix(intRow, .ColIndex("�����к�"))) Then
                    MsgBox "���й�ϵ�е�" & intRow & "�л����кźͶ����к�����ͬһ�У����飡", vbInformation, gstrSysName
                    Exit Function
                End If
                '������õĶ�����Ŀ�����Ƿ��ظ�
                For lngCount = intRow + 1 To .Rows - 1
                    If Val(.TextMatrix(intRow, .ColIndex("�����к�"))) = Val(.TextMatrix(lngCount, .ColIndex("�����к�"))) Then
                        MsgBox "���й�ϵ�е�" & intRow & "�к͵�" & lngCount & "�����õĶ�����Ŀ�ظ�,���飡", vbInformation, gstrSysName
                        Exit Function
                    End If
                Next lngCount
            End If
        Next intRow
    End With
    
    '������ʱ��
    With vsf
        For intRow = 1 To .Rows - 1
            If .TextMatrix(intRow, 1) <> "" And .TextMatrix(intRow, 2) <> "" And .TextMatrix(intRow, 3) <> "" Then
                
                If CheckTime(intRow, 2, .TextMatrix(intRow, 2)) = False Then Exit Function
                If CheckTime(intRow, 3, .TextMatrix(intRow, 3)) = False Then Exit Function
                
            End If
        Next
    End With
    
    rsItems.Filter = 0
    CheckData = True
End Function

Private Function GetItemName(ByVal strData As String, ByVal intOrder As Integer) As String
    Dim intDo As Integer, intPos As Integer
    '��ȡָ����ʽ����ָ����ŵ���Ŀ���ƣ���ʽ�磺{����ѹ}/ {����ѹ}mmHg
    
    intPos = InStr(1, strData, "{")
    If intOrder > 0 Then intPos = InStr(intPos + 1, strData, "{")
    strData = Mid(strData, intPos + 1)
    strData = Mid(strData, 1, InStr(1, strData, "}") - 1)
    GetItemName = strData
End Function

Private Function CheckTime(ByVal lngRow As Long, ByVal lngCol As Long, StrText As String) As Boolean

    Dim intRow As Integer
    Dim intPos As Integer
    Dim strTmp As String
    Dim lngHour As Long, lngMin As Long
    
    intPos = InStr(StrText, ":")
    
    If StrText = "" Then
        CheckTime = True
        Exit Function
    End If
    
    If intPos > 0 Then
        
        strTmp = Mid(StrText, 1, intPos - 1)
        lngHour = Val(strTmp)
        
        If Val(strTmp) < 0 Or Val(strTmp) > 23 Then
            MsgBox "Сʱֻ����0-23֮�䣡", vbInformation, gstrSysName
            Exit Function
        End If
        
        StrText = Mid(StrText, intPos + 1)
        intPos = InStr(StrText, ":")
        If intPos > 0 Then
            strTmp = Mid(StrText, 1, intPos - 1)
            lngMin = Val(strTmp)
            If Val(strTmp) < 0 Or Val(strTmp) > 59 Then
                MsgBox "����ֻ����0-59֮�䣡", vbInformation, gstrSysName
                Exit Function
            End If
            
            If InStr(Mid(StrText, intPos + 1), ":") > 0 Then
                
                MsgBox "ʱ���ʽ����ȷ��", vbInformation, gstrSysName
                Exit Function
            Else
                strTmp = Mid(StrText, intPos + 1)
                If Val(strTmp) < 0 Or Val(strTmp) > 59 Then
                    MsgBox "����ֻ����0-59֮�䣡", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        Else
            strTmp = StrText
            lngMin = Val(strTmp)
            If Val(strTmp) < 0 Or Val(strTmp) > 59 Then
                MsgBox "����ֻ����0-59֮�䣡", vbInformation, gstrSysName
                Exit Function
            End If
            
        End If
    Else
        If Val(StrText) < 0 Or Val(StrText) > 23 Then
            MsgBox "Сʱֻ����0-23֮�䣡", vbInformation, gstrSysName
            Exit Function
        End If
        lngHour = Val(StrText)
    End If
    
    StrText = String(2 - Len(CStr(lngHour)), "0") & CStr(lngHour) & ":" & String(2 - Len(CStr(lngMin)), "0") & CStr(lngMin)
    vsf.TextMatrix(lngRow, lngCol) = StrText
    
    CheckTime = True
    
End Function

Private Function SaveData() As Boolean
    Dim blnTrans As Boolean
    Dim blnReCalc As Boolean    '��ÿҳ�Ŀɴ�ӡ�����з����仯��������δ��ӡ����ļ���������
    Dim strCOL As String
    Dim strInput As String
    Dim intPageAlign As Integer, intPageOption As Integer
    Dim lngRows As Long, lngFixedRows As Long
    Dim lngRow As Long, strRelation As String, blnRelation As Boolean '�����ж��չ�ϵ�趨����
    Dim strHead As String
    Dim strColItemRelation As String
    
    If CheckData = False Then Exit Function
        
    '��������
    If Me.optOrient(0).Value = True Then
        If Val(Me.txtMarjin(0).Text) > Val(Me.txtHeight.Text) / 3 Then MsgBox "�ϱ߾�̫��", vbExclamation, gstrSysName: Exit Function
        If Val(Me.txtMarjin(1).Text) > Val(Me.txtHeight.Text) / 3 Then MsgBox "�±߾�̫��", vbExclamation, gstrSysName: Exit Function
        If Val(Me.txtMarjin(2).Text) > Val(Me.txtWidth.Text) / 3 Then MsgBox "��߾�̫��", vbExclamation, gstrSysName: Exit Function
        If Val(Me.txtMarjin(3).Text) > Val(Me.txtWidth.Text) / 3 Then MsgBox "�ұ߾�̫��", vbExclamation, gstrSysName: Exit Function
    Else
        If Val(Me.txtMarjin(0).Text) > Val(Me.txtWidth.Text) / 3 Then MsgBox "�ϱ߾�̫��", vbExclamation, gstrSysName: Exit Function
        If Val(Me.txtMarjin(1).Text) > Val(Me.txtWidth.Text) / 3 Then MsgBox "�±߾�̫��", vbExclamation, gstrSysName: Exit Function
        If Val(Me.txtMarjin(2).Text) > Val(Me.txtHeight.Text) / 3 Then MsgBox "��߾�̫��", vbExclamation, gstrSysName: Exit Function
        If Val(Me.txtMarjin(3).Text) > Val(Me.txtHeight.Text) / 3 Then MsgBox "�ұ߾�̫��", vbExclamation, gstrSysName: Exit Function
    End If
    
    '������Ч������(ֽ��|ֽ��|��|��|�ϱ߾�|�±߾�|��߾�|�ұ߾�|�и�|�̶�����|������������|�����������С|�����ı�|������������|�����������С|�������ı�)
    If optTabTiers(0).Value Then
        lngFixedRows = 1
    ElseIf optTabTiers(1).Value Then
        lngFixedRows = 2
    Else
        lngFixedRows = 3
    End If
    
    strHead = ""
    '��ȡ��ͷ��Ϣ
    For lngCount = vfgThis.FixedCols To vfgThis.Cols - 1
        If vfgThis.RowHidden(3) = False Then strHead = strHead & "'" & lngCount - 1 & ",0," & Trim(vfgThis.TextMatrix(3, lngCount)) & "," & vfgThis.ColWidth(lngCount)
        If vfgThis.RowHidden(4) = False Then strHead = strHead & "'" & lngCount - 1 & ",1," & Trim(vfgThis.TextMatrix(4, lngCount)) & "," & vfgThis.ColWidth(lngCount)
        If vfgThis.RowHidden(5) = False Then strHead = strHead & "'" & lngCount - 1 & ",2," & Trim(vfgThis.TextMatrix(5, lngCount)) & "," & vfgThis.ColWidth(lngCount)
    Next
    strHead = Mid(strHead, 2)
    
    strInput = Me.cboPaperKind.ItemData(Me.cboPaperKind.ListIndex) & "|" & IIf(Me.optOrient(0).Value, 1, 2) & "|" & _
               Int(Me.ScaleY(Val(Me.txtHeight.Text), vbMillimeters, vbTwips)) & "|" & Int(Me.ScaleX(Val(Me.txtWidth.Text), vbMillimeters, vbTwips)) & "|" & _
               Int(Me.ScaleY(Val(Me.txtMarjin(0).Text), vbMillimeters, vbTwips)) & "|" & Int(Me.ScaleY(Val(Me.txtMarjin(1).Text), vbMillimeters, vbTwips)) & "|" & _
               Int(Me.ScaleX(Val(Me.txtMarjin(2).Text), vbMillimeters, vbTwips)) & "|" & Int(Me.ScaleX(Val(Me.txtMarjin(3).Text), vbMillimeters, vbTwips)) & "|" & _
               Val(txtTabRowHeight.Text) & "|" & lngFixedRows & "|" & Split(lblTitleFont.Caption, ",")(0) & "|" & Split(lblTitleFont.Caption, ",")(1) & "|" & _
               txtTitleText.Text & "|" & Split(lblTabFont.Caption, ",")(0) & "|" & Split(lblTabFont.Caption, ",")(1) & "|" & strTemp & "|" & strHead
    lngRows = frmTendFilePreview.ShowMe(Me, strInput)
    If chk����������.Value = 1 Then
        If lngRows <> mlngPageRow And mlngPageRow > 0 Then
            '�����з����仯�����Ѵ�ӡ������Ӱ�죬��ʾ
            If MsgBox("    �����޸ĵ���ÿҳ�ɴ�ӡ�������з����仯���Ѵ�ӡ���������޸Ļ��ش򽫻ᵼ�´�ӡ���ң��Ƿ������" & vbCrLf & "    ԭ��ÿҳ�ɴ�ӡ" & mlngPageRow & "�У�����ÿҳ�ɴ�ӡ" & lngRows & "��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            blnReCalc = True
            '87990,����,2015-09-02
        Else
            If MsgBox("    �����޸���������ѡ�����������У�Ĭ�Ͻ����ʹ�øø�ʽ���ļ����¼��������������Ƿ������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
            blnReCalc = True
        End If
    Else
        lngRows = mlngPageRow
    End If
    
    '�������ҳ�������ֹ����
    If OverRun Then
        MsgBox "���Ŀ�ȳ�����ֽ����Ч��ӡ��Χ��", vbInformation, gstrSysName
        Exit Function
    End If
    If Not PageHeadTest Then Exit Function
    
    'ƴ��
    If chkҳ��.Value = 1 Then
        If optPageAlign(0).Value Then
            intPageAlign = 1
        ElseIf optPageAlign(1).Value Then
            intPageAlign = 2
        Else
            intPageAlign = 3
        End If
        '46251,������,2012-09-11
        intPageOption = Val(cboҳ��.ItemData(cboҳ��.ListIndex))
    End If
    If Me.optTabTiers(0).Value Then
        gstrSQL = mlngFileID & ",1," & Me.udTabCols.Value & "," & Val(Me.txtTabRowHeight.Text)
    ElseIf Me.optTabTiers(1).Value Then
        gstrSQL = mlngFileID & ",2," & Me.udTabCols.Value & "," & Val(Me.txtTabRowHeight.Text)
    Else
        gstrSQL = mlngFileID & ",3," & Me.udTabCols.Value & "," & Val(Me.txtTabRowHeight.Text)
    End If

    gstrSQL = gstrSQL & ",'" & Me.lblTabFont.Caption & "'," & Me.lblTabTextColor.ForeColor & "," & Me.shpTabGridColor.BorderColor
    gstrSQL = gstrSQL & ",'" & Trim(Me.txtTitleText.Text) & "','" & Me.lblTitleFont.Caption & "'"
    gstrSQL = gstrSQL & "," & Me.udRecordFrom.Value & "," & Me.udRecordTo.Value & ",'" & Me.lblRecordFont.Caption & "'," & Me.lblRecordColor.ForeColor & "," & lngRows & "," & chk���кϲ�.Value & "," & chkHideTime.Value
    
    gstrSQL = gstrSQL & ",'" & Me.cboPaperKind.ItemData(Me.cboPaperKind.ListIndex)
    gstrSQL = gstrSQL & ";" & IIf(Me.optOrient(0).Value, 1, 2)
    gstrSQL = gstrSQL & ";" & Int(Me.ScaleY(Val(Me.txtHeight.Text), vbMillimeters, vbTwips))
    gstrSQL = gstrSQL & ";" & Int(Me.ScaleX(Val(Me.txtWidth.Text), vbMillimeters, vbTwips))
    gstrSQL = gstrSQL & ";" & Int(Me.ScaleX(Val(Me.txtMarjin(2).Text), vbMillimeters, vbTwips))
    gstrSQL = gstrSQL & ";" & Int(Me.ScaleX(Val(Me.txtMarjin(3).Text), vbMillimeters, vbTwips))
    gstrSQL = gstrSQL & ";" & Int(Me.ScaleY(Val(Me.txtMarjin(0).Text), vbMillimeters, vbTwips))
    gstrSQL = gstrSQL & ";" & Int(Me.ScaleY(Val(Me.txtMarjin(1).Text), vbMillimeters, vbTwips)) & "'"
    gstrSQL = gstrSQL & "," & intPageOption & "," & intPageAlign
    
    With Me.vfgThis
        gstrSQL = gstrSQL & ",'" & Replace(Replace(.TextMatrix(2, .FixedCols), " ", "|"), vbCrLf, "'||Chr(13)||Chr(10)||'") & "'"
        strTemp = ""
        For lngCount = .FixedCols To .Cols - 1
            If .RowHidden(3) = False Then strTemp = strTemp & "|" & lngCount & ",1," & Trim(.TextMatrix(3, lngCount))
            If .RowHidden(4) = False Then strTemp = strTemp & "|" & lngCount & ",2," & Trim(.TextMatrix(4, lngCount))
            If .RowHidden(5) = False Then strTemp = strTemp & "|" & lngCount & ",3," & Trim(.TextMatrix(5, lngCount))
'            strTemp = strTemp & "|" & lngCount & ",3," & Trim(.TextMatrix(5, lngCount))
        Next
        gstrSQL = gstrSQL & ",'" & Mid(strTemp, 2) & "'"
        
        strCOL = ""
        blnRelation = False
        For lngCount = .FixedCols To .Cols - 1
            '�������������ж��չ�ϵ���趨
            strRelation = ""
            If Trim(.Cell(flexcpData, 6, lngCount, 6, lngCount)) <> "" Then
                For lngRow = vfgColRelation.FixedRows To vfgColRelation.Rows - 1
                    If Val(vfgColRelation.TextMatrix(lngRow, vfgColRelation.ColIndex("�����к�"))) = lngCount And Val(vfgColRelation.TextMatrix(lngRow, vfgColRelation.ColIndex("�����к�"))) > 0 Then
                        strRelation = Val(vfgColRelation.TextMatrix(lngRow, vfgColRelation.ColIndex("�����к�")))
                        Exit For
                    End If
                Next lngRow
            End If
            
            strColItemRelation = ""
            If Trim(.Cell(flexcpData, 6, lngCount, 6, lngCount)) <> "" Then
                For lngRow = vsfColDoctor.FixedRows To vsfColDoctor.Rows - 1
                    If Val(vsfColDoctor.TextMatrix(lngRow, vsfColDoctor.ColIndex("��Ŀ�к�"))) = lngCount And Trim(vsfColDoctor.TextMatrix(lngRow, vsfColDoctor.ColIndex("��������"))) <> "" Then
                        strColItemRelation = vsfColDoctor.TextMatrix(lngRow, vsfColDoctor.ColIndex("��������"))
                        Exit For
                    End If
                Next lngRow
            End If
            
            If strRelation = "" Then
                strRelation = lngCount
            Else
                strRelation = lngCount & "`" & Val(strRelation)
                blnRelation = True
            End If
            strCOL = strCOL & "|" & strRelation & "," & .ColWidth(lngCount) & "`" & .Cell(flexcpAlignment, 6, lngCount) & "," & Trim(.Cell(flexcpData, 6, lngCount, 6, lngCount)) & "," & strColItemRelation
'            strTemp = strTemp & "|" & lngCount & "," & .ColWidth(lngCount) & "," & Trim(.Cell(flexcpData, 6, lngCount, 6, lngCount))
        Next
        gstrSQL = gstrSQL & ",'" & Mid(strCOL, 2) & "'"
    End With
    
    '��д����ʱ��
    '------------------------------------------------------------------------------------------------------------------
    With vsf
        strTemp = ""
        
        For lngCount = 1 To .Rows - 1
            If .TextMatrix(lngCount, 1) <> "" And .TextMatrix(lngCount, 2) <> "" And .TextMatrix(lngCount, 3) <> "" Then
                strTemp = strTemp & "|" & Trim(.TextMatrix(lngCount, 1)) & "," & Trim(.TextMatrix(lngCount, 2)) & "," & Trim(.TextMatrix(lngCount, 3))
            Else
                .TextMatrix(lngCount, 1) = ""
                .TextMatrix(lngCount, 2) = ""
                .TextMatrix(lngCount, 3) = ""
            End If
        Next
        
        If strTemp <> "" Then strTemp = Mid(strTemp, 2)
        gstrSQL = gstrSQL & ",'" & strTemp & "'"
    End With
    
    gstrSQL = gstrSQL & ",NULL"
    gstrSQL = "Zl_�����ļ���ʽ_Update(" & gstrSQL & ")"
    
    Err = 0: On Error GoTo errHand
    gcnOracle.BeginTrans
    blnTrans = True
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    If Not SavePageHead(picFoot.Tag) Then
        gcnOracle.RollbackTrans
        Exit Function
    End If
    If Not SavePageFoot(picFoot.Tag) Then
        gcnOracle.RollbackTrans
        Exit Function
    End If
    gcnOracle.CommitTrans
    blnTrans = False
    
    SaveData = True
    mblnOk = True
    mlngPageRow = lngRows
    cmdͬ��.Enabled = True
    
    '�������з����仯��������������
    If blnReCalc Then
        If MsgBox("�ɴ�ӡ�����з����仯���Ƿ��ʹ�øø�ʽ�������ļ��������㣿" & vbCrLf & "�����´�ӡ���ݿ��ܻ�ռ���㼸����ʱ�䣩", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            If frmTendFilePreview.AnaliseData(Me, mlngFileID, strInput) Then MsgBox "��ӡ�����Ѹ��£�", vbInformation, gstrSysName
        End If
    End If
    
    Exit Function

errHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SavePageHead(ByVal StrKey As String, Optional ByVal strZipFile As String = "") As Boolean
    'blnBuild=False:�����ļ���ѹ��;True:�Ѳ���ѹ���ļ�
    Dim strFile As String, strZip As String
    If strZipFile = "" Then
        strFile = App.Path & "\Head_S.rtf"
        rtbHead.SaveFile strFile
        strZip = zlFileZip(strFile)
    Else
        strZip = strZipFile
    End If
    SavePageHead = zlBlobSave(12, StrKey, strZip)
    If strZipFile = "" Then
        gobjFSO.DeleteFile strFile, True
        gobjFSO.DeleteFile strZip, True
    End If
End Function

Private Function SavePageFoot(ByVal StrKey As String, Optional ByVal strZipFile As String = "") As Boolean
    'blnBuild=False:�����ļ���ѹ��;True:�Ѳ���ѹ���ļ�
    Dim strFile As String, strZip As String
    If strZipFile = "" Then
        strFile = App.Path & "\Foot_S.rtf"
        rtbFoot.SaveFile strFile
        strZip = zlFileZip(strFile)
    Else
        strZip = strZipFile
    End If
    SavePageFoot = zlBlobSave(13, StrKey, strZip)
    If strZipFile = "" Then
        gobjFSO.DeleteFile strFile, True
        gobjFSO.DeleteFile strZip, True
    End If
End Function

Private Function PreparePage()
    Dim sFont As String, i As Integer
    
    For i = 0 To Screen.FontCount - 1
       sFont = Screen.Fonts(i)
       cboFont.AddItem sFont
       If sFont = "����" Then cboFont.ListIndex = i
    Next i
    With cboFSize
        .AddItem "����"
        .AddItem "С��"
        .AddItem "һ��"
        .AddItem "Сһ"
        .AddItem "����"
        .AddItem "С��"
        .AddItem "����"
        .AddItem "С��"
        .AddItem "�ĺ�"
        .AddItem "С��"
        .AddItem "���"
        .AddItem "С��"
        .AddItem "����"
        .AddItem "С��"
        .AddItem "�ߺ�"
        .AddItem "�˺�"
        .AddItem 5
        .AddItem 5.5
        .AddItem 6.5
        .AddItem 7.5
        .AddItem 8
        .AddItem 9
        .AddItem 10
        .AddItem 10.5
        .AddItem 11
        .AddItem 12
        .AddItem 14
        .AddItem 16
        .AddItem 18
        .AddItem 20
        .AddItem 22
        .AddItem 24
        .AddItem 26
        .AddItem 28
        .AddItem 36
        .AddItem 48
        .AddItem 72
        .ListIndex = 10
    End With
End Function

Private Function OverRun() As Boolean
    Dim lngPageMargin As Long    '�߾�
    Dim lngPageWidth As Long     'ֽ�ſ��
    Dim lngTrimSize As Long      'ֽ��ʵ�ʿ��
    Dim dblTableWidth As Double     '�����
    Dim lngCol As Long, lngCols As Long
    '�������������Ƿ񳬹����ұ߾�
    
    lngCols = Me.vfgThis.Cols - 1
    For lngCol = 1 To lngCols
        dblTableWidth = dblTableWidth + Val(vfgThis.ColWidth(lngCol))
    Next
    
    '��������Ƿ񳬳�ҳ����Ч��ӡ��Χ
    Printer.Orientation = IIf(optOrient(0).Value, 1, 2)
    If cboPaperKind.ItemData(cboPaperKind.ListIndex) = 256 Then
        Call SetCustonPager(Me.hWnd, Int(Me.ScaleY(Val(Me.txtWidth.Text), vbMillimeters, vbTwips)), Int(Me.ScaleY(Val(Me.txtHeight.Text), vbMillimeters, vbTwips)))
    Else
        Printer.PaperSize = cboPaperKind.ItemData(cboPaperKind.ListIndex)
    End If
    lngPageWidth = Printer.ScaleWidth
    lngPageMargin = Int(Me.ScaleX(Val(Me.txtMarjin(2).Text), vbMillimeters, vbTwips)) + Int(Me.ScaleX(Val(Me.txtMarjin(3).Text), vbMillimeters, vbTwips))
    lngTrimSize = lngPageWidth - lngPageMargin - 100
    If dblTableWidth > lngTrimSize Then
        OverRun = True
        Exit Function
    End If
End Function

Private Sub cmdRecordColor_Click()
    lngCurColor = Me.lblRecordColor.ForeColor
    Call zlColorSet("������ʽ��ɫ", lngCurColor)
    If lngCurColor = Me.lblRecordColor.ForeColor Then Exit Sub
    Me.lblRecordColor.ForeColor = lngCurColor
    With Me.vfgThis
        .Cell(flexcpForeColor, 7, 1, 7, .Cols - 1) = Me.lblRecordColor.ForeColor
    End With
    DataChanged = True
End Sub

Private Sub cmdRecordFont_Click()
    strCurFont = Me.lblRecordFont.Caption
    Call zlFontSet("������ʽ����", strCurFont)
    If strCurFont = Me.lblRecordFont.Caption Then Exit Sub
    Me.lblRecordFont.Caption = strCurFont
    Set objFont = New StdFont
    With objFont
        .Name = Split(strCurFont, ",")(0)
        .Size = Val(Split(strCurFont, ",")(1))
        .Bold = False: .Italic = False
        If InStr(1, strCurFont, "��") > 0 Then .Bold = True
        If InStr(1, strCurFont, "б") > 0 Then .Italic = True
    End With
    With Me.vfgThis
        Set .Cell(flexcpFont, 7, .FixedCols, 7, .Cols - 1) = objFont
    End With
    DataChanged = True
End Sub

Private Sub cmdTabFont_Click()
    strCurFont = Me.lblTabFont.Caption
    Call zlFontSet("�ı�����", strCurFont)
    If strCurFont = Me.lblTabFont.Caption Then Exit Sub
    Me.lblTabFont.Caption = strCurFont
'    Me.lblRecordFont.Caption = strCurFont
    Set objFont = New StdFont
    With objFont
        .Name = Split(strCurFont, ",")(0)
        .Size = Val(Split(strCurFont, ",")(1))
         .Bold = False: .Italic = False
        If InStr(1, strCurFont, "��") > 0 Then .Bold = True
        If InStr(1, strCurFont, "б") > 0 Then .Italic = True
    End With
    Set Me.vfgThis.Font = objFont
    DataChanged = True
End Sub

Private Sub cmdTabGridColor_Click()
    lngCurColor = Me.shpTabGridColor.BorderColor
    Call zlColorSet("�����ɫ", lngCurColor)
    If lngCurColor = Me.shpTabGridColor.BorderColor Then Exit Sub
    Me.shpTabGridColor.BorderColor = lngCurColor
    With Me.vfgThis
        .GridColor = Me.shpTabGridColor.BorderColor
        .CellBorderRange 2, .FixedCols, 2, .Cols - 1, .GridColor, 0, 0, 0, 1, 0, 0
        .CellBorderRange 3, .FixedCols, 7, .Cols - 1, .GridColor, 1, 1, 1, 1, 1, 1
    End With
    DataChanged = True
End Sub

Private Sub cmdTabTextColor_Click()
    lngCurColor = Me.lblTabTextColor.ForeColor
    Call zlColorSet("�ı���ɫ", lngCurColor)
    If lngCurColor = Me.lblTabTextColor.ForeColor Then Exit Sub
    Me.lblTabTextColor.ForeColor = lngCurColor
    Me.vfgThis.ForeColor = Me.lblTabTextColor.ForeColor
    DataChanged = True
End Sub

Private Sub cmdTitleFont_Click()
    strCurFont = Me.lblTitleFont.Caption
    Call zlFontSet("��������", strCurFont)
    If strCurFont = Me.lblTitleFont.Caption Then Exit Sub
    Me.lblTitleFont.Caption = strCurFont
    Set objFont = New StdFont
    With objFont
        .Name = Split(strCurFont, ",")(0)
        .Size = Val(Split(strCurFont, ",")(1))
        .Bold = False: .Italic = False
        If InStr(1, strCurFont, "��") > 0 Then .Bold = True
        If InStr(1, strCurFont, "б") > 0 Then .Italic = True
    End With
    With Me.vfgThis
        Set .Cell(flexcpFont, 1, .FixedCols, 1, .Cols - 1) = objFont
        .ROWHEIGHT(1) = objFont.Size * 20 + 150
    End With
    DataChanged = True
End Sub

Private Sub cmdͬ��_Click()
    Dim intPageAlign As Integer
    Dim strZIPHead As String, strZIPFoot As String
    Dim rsTemp As New ADODB.Recordset
    Dim blnTrans As Boolean
    On Error GoTo errHand
    '����ǰ��ʽӦ�õ����л����ļ�
    
    gstrSQL = " Select ����||'-'||��� AS KEY From �����ļ��б� Where ����=3 and ID<>[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, "��ȡ�����ļ�", mlngFileID)
    If rsTemp.RecordCount = 0 Then
        MsgBox "��ǰֻ��һ�ݻ����ļ�������Ҫִ��ͬ�����ܣ�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    If MsgBox("���ٴ�ȷ�ϣ�" & vbCrLf & "        ִ�иù��ܺ����л����ļ���ҳüҳ�Ÿ�ʽ��ͳһ�뵱ǰ�ļ����ñ���һ�£�", vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Sub
    
    '�Ƚ���ǰ�ļ���ҳüҳ��ѹ����������
    If chkҳ��.Value = 1 Then
        If optPageAlign(0).Value Then
            intPageAlign = 1
        ElseIf optPageAlign(1).Value Then
            intPageAlign = 2
        Else
            intPageAlign = 3
        End If
    End If
    strZIPHead = ReadPageHeadFile(picFoot.Tag)
    strZIPFoot = ReadPageFootFile(picFoot.Tag)
    
    gcnOracle.BeginTrans
    blnTrans = True
    'ѭ��д�����ݿ�
    With rsTemp
        Do While Not .EOF
            If Not SavePageHead(!Key, strZIPHead) Then
                gcnOracle.RollbackTrans
                Exit Sub
            End If
            If Not SavePageFoot(!Key, strZIPFoot) Then
                gcnOracle.RollbackTrans
                Exit Sub
            End If
            .MoveNext
        Loop
    End With
    Call zlDatabase.ExecuteProcedure("ZL_����ҳ���ʽ_ҳ��(" & intPageAlign & ",'" & Split(picFoot.Tag, "-")(1) & "')", "����ҳ��")
    gcnOracle.CommitTrans
    blnTrans = False
    'ɾ����ʱ�ļ�
    
    gobjFSO.DeleteFile strZIPHead, True
    gobjFSO.DeleteFile strZIPFoot, True
    
    MsgBox "ͬ���ɹ���", vbInformation, gstrSysName
    Exit Sub
errHand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub dkpMan_AttachPane(ByVal Item As XtremeDockingPane.IPane)
    Select Case Item.ID
    Case 1
        Item.Handle = picPane(0).hWnd
    Case 2
        Item.Handle = picPane(1).hWnd
    End Select
End Sub

Private Sub Form_Load()
    Me.picTable.BackColor = Me.BackColor
    Me.picLabel.BackColor = Me.BackColor
    Me.picCloumn.BackColor = Me.BackColor
    Me.picPaper.BackColor = Me.BackColor
    Me.picFoot.BackColor = Me.BackColor

    '---------------------------------------------------
    '����ҳ����֯
    With Me.tbcStyle
        .Left = Me.ScaleLeft: .Top = Me.ScaleTop
        With .PaintManager
'            .Appearance = xtpTabAppearancePropertyPage
            .Appearance = xtpTabAppearancePropertyPage2003
            .BoldSelected = True
            .ClientFrame = xtpTabFrameBorder
        End With
        .InsertItem 0, "��������", Me.picTable.hWnd, 0
        .InsertItem 1, "���ϱ�ǩ", Me.picLabel.hWnd, 0
        .InsertItem 2, "��������", Me.picCloumn.hWnd, 0
        .InsertItem 3, "��ӡ����", Me.picPaper.hWnd, 0
        .InsertItem 4, "ҳ���ʽ", Me.picFoot.hWnd, 0
        .InsertItem 5, "����ʱ��", Me.picSum.hWnd, 0
        .Item(0).Selected = True
    End With
    With Me.vfgThis
        .Left = Me.ScaleLeft: .Width = Me.ScaleWidth
        .Top = Me.tbcStyle.Top + Me.tbcStyle.Height + 45: .Height = Me.ScaleHeight - .Top
    End With
'    With Me.picVBar
'        .BackColor = Me.BackColor
'        .Left = Me.ScaleWidth - .Width
'        .Top = 0: .Height = Me.Height
'    End With
    
    '---------------------------------------------------
    'Ĭ����ʽ����
    Err = 0: On Error GoTo 0
    With Me.vfgThis
        .Rows = 9
        .TextMatrix(1, 0) = "�����ı�"
        .TextMatrix(2, 0) = "���ϱ�ǩ"
        
        .TextMatrix(3, 0) = "��ͷ��Ԫ"
        .TextMatrix(4, 0) = "��ͷ��Ԫ"
        .TextMatrix(5, 0) = "��ͷ��Ԫ"
        
        .TextMatrix(6, 0) = "��������"
        .TextMatrix(7, 0) = "������ʽ"
        .TextMatrix(8, 0) = "���±�ǩ"
        
        .ColWidth(0) = 1200
        .MergeCol(0) = True
        .RowHidden(3) = True
        .RowHidden(4) = True
        
        For lngCount = .FixedCols To .Cols - 1
            .TextMatrix(0, lngCount) = lngCount
            .ColAlignment(lngCount) = flexAlignCenterCenter
            .FixedAlignment(lngCount) = flexAlignCenterCenter
'            .TextMatrix(.Rows - 1, lngCount) = "��ӡʱ�䣺[��ӡʱ��]                            ��[ҳ��]ҳ"
        Next
        .MergeRow(1) = True
        .MergeRow(2) = True: .Cell(flexcpAlignment, 2, 1, 2, .Cols - 1) = flexAlignGeneralCenter
        .MergeRow(.Rows - 1) = True: .Cell(flexcpAlignment, .Rows - 1, 1, .Rows - 1, .Cols - 1) = flexAlignGeneralCenter
        .MergeRow(3) = True
        .MergeRow(4) = True
        .MergeRow(5) = True
        
        .Cell(flexcpAlignment, 6, 1, 7, .Cols - 1) = flexAlignGeneralCenter
        
        Call udTabCols_Change
        Call txtTabRowHeight_Change
        strCurFont = Me.lblTabFont.Caption
        Set objFont = New StdFont
        With objFont
            .Name = Split(strCurFont, ",")(0)
            .Size = Val(Split(strCurFont, ",")(1))
            .Bold = False: .Italic = False
            If InStr(1, strCurFont, "��") > 0 Then .Bold = True
            If InStr(1, strCurFont, "б") > 0 Then .Italic = True
        End With
        
        .ForeColor = Me.lblTabTextColor.ForeColor
        
        .GridColor = Me.shpTabGridColor.BorderColor
        .CellBorderRange 2, .FixedCols, 2, .Cols - 1, .GridColor, 0, 0, 0, 1, 0, 0
        .CellBorderRange 3, .FixedCols, 7, .Cols - 1, .GridColor, 1, 1, 1, 1, 1, 1
            
        Call txtTitleText_Change
        strCurFont = Me.lblTitleFont.Caption
        Set objFont = New StdFont
        With objFont
            .Name = Split(strCurFont, ",")(0)
            .Size = Val(Split(strCurFont, ",")(1))
            .Bold = False: .Italic = False
            If InStr(1, strCurFont, "��") > 0 Then .Bold = True
            If InStr(1, strCurFont, "б") > 0 Then .Italic = True
        End With
        Set .Cell(flexcpFont, 1, .FixedCols, 1, .Cols - 1) = objFont
        .ROWHEIGHT(1) = objFont.Size * 20 + 150
            
        strCurFont = Me.lblRecordFont.Caption
        Set objFont = New StdFont
        With objFont
            .Name = Split(strCurFont, ",")(0)
            .Size = Val(Split(strCurFont, ",")(1))
            .Bold = False: .Italic = False
            If InStr(1, strCurFont, "��") > 0 Then .Bold = True
            If InStr(1, strCurFont, "б") > 0 Then .Italic = True
        End With
        Set .Cell(flexcpFont, 7, .FixedCols, 7, .Cols - 1) = objFont
        .Cell(flexcpForeColor, 7, 1, 7, .Cols - 1) = Me.lblRecordColor.ForeColor
        
        .RowHidden(3) = True
        .RowHidden(4) = True
        
    End With
    vfgThis.AutoSize 0, vfgThis.Cols - 1
    
    With Me.vfgColRelation
        .Rows = 3
        .Cols = 4
        .FixedRows = 2
        .FixedCols = 2
        .MergeCells = flexMergeNever
        .MergeCellsFixed = flexMergeRestrictRows
        .MergeRow(0) = True
        .TextMatrix(0, 0) = "������Ŀ"
        .TextMatrix(0, 1) = "������Ŀ"
        .TextMatrix(0, 2) = "������Ŀ"
        .TextMatrix(0, 3) = "������Ŀ"
        .TextMatrix(1, 0) = "�к�"
        .TextMatrix(1, 1) = "����"
        .TextMatrix(1, 2) = "�к�"
        .TextMatrix(1, 3) = "����"
        .ColKey(0) = "�����к�"
        .ColKey(1) = "��������"
        .ColKey(2) = "�����к�"
        .ColKey(3) = "��������"
        .ExtendLastCol = True
        .Editable = flexEDKbdMouse
    End With
    
    
    With Me.vsfColDoctor
        .Rows = 2
        .Cols = 3
        .FixedRows = 2
        .FixedCols = 2
        .MergeCells = flexMergeNever
        .MergeCellsFixed = flexMergeRestrictRows
        .MergeRow(0) = True
        
        .TextMatrix(0, 0) = "��Ŀ"
        .TextMatrix(0, 1) = "��Ŀ"
        .TextMatrix(0, 2) = "������Ŀ"
        .TextMatrix(1, 0) = "�к�"
        .TextMatrix(1, 1) = "����"
        .TextMatrix(1, 2) = "����"
        .ColKey(0) = "��Ŀ�к�"
        .ColKey(1) = "��Ŀ����"
        .ColKey(2) = "��������"
        .ExtendLastCol = True
        .Editable = flexEDKbdMouse
    End With
    
    Set mclsVsf = New clsVsf
    With mclsVsf
        
        Call .Initialize(Me.Controls, vsf, True, True)
        Call .ClearColumn

        Call .AppendColumn("", 255, flexAlignCenterCenter, flexDTString, "", "[ͼ��]", False)
        
        Call .AppendColumn("ʱ������", 2100, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("��ʼʱ��", 900, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("����ʱ��", 900, flexAlignLeftCenter, flexDTString, "", , True)
        Call .AppendColumn("", 15, flexAlignLeftCenter, flexDTString, "", , True)
        
        Call .InitializeEdit(True, True, True)
        Call .InitializeEditColumn(mclsVsf.ColIndex("ʱ������"), True, vbVsfEditText)
        Call .InitializeEditColumn(mclsVsf.ColIndex("��ʼʱ��"), True, vbVsfEditText)
        Call .InitializeEditColumn(mclsVsf.ColIndex("����ʱ��"), True, vbVsfEditText)
                
        .AppendRows = True
    End With
    
    vsf.ColHidden(0) = True
    
    Dim objPane As Pane
    
    Call RestoreWinState(Me, App.ProductName)
    
    Call InitMenuBar
    
    dkpMan.Options.ThemedFloatingFrames = True
    dkpMan.Options.UseSplitterTracker = False 'ʵʱ�϶�
    dkpMan.Options.AlphaDockingContext = True
    dkpMan.Options.CloseGroupOnButtonClick = True
    dkpMan.Options.HideClient = True
    dkpMan.SetCommandBars cbsThis

    Set objPane = dkpMan.CreatePane(1, 100, 100, DockTopOf, Nothing): objPane.Title = "ʾ��": objPane.Options = PaneNoCaption
    Set objPane = dkpMan.CreatePane(2, 100, 200, DockBottomOf, objPane): objPane.Title = "���": objPane.Options = PaneNoCaption
                
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    
    rtbHead.Width = picFoot.Width - 500
    rtbFoot.Width = rtbHead.Width
    
    Call SetPaneRange(dkpMan, 1, 15, 260, Me.ScaleWidth, 260)
    dkpMan.RecalcLayout
    
    rtbHead.Width = picFoot.Width - 500
    rtbFoot.Width = rtbHead.Width
End Sub

Private Sub Form_Unload(Cancel As Integer)

    If DataChanged Then
        Cancel = (MsgBox("���ĺ����Ʊ��뱣������Ч���Ƿ�������棿", vbYesNo + vbQuestion + vbDefaultButton2, gstrSysName) = vbNo)
    End If
    
    If Cancel Then Exit Sub
    
    DataChanged = False
    
    Call SaveWinState(Me, App.ProductName)
End Sub

Private Sub lstColumnItems_DblClick()
    Call cmdColumn_Click(0)
End Sub

Private Sub lstColumnUsed_Click()
    With Me.lstColumnUsed
        If .ListIndex = -1 Then Exit Sub
        Me.txtColumnPrefix.Text = Left(.List(.ListIndex), InStr(1, .List(.ListIndex), "{") - 1)
        Me.txtColumnPostfix.Text = Mid(.List(.ListIndex), InStr(1, .List(.ListIndex), "}") + 1)
        chk.Value = .ItemData(.ListIndex)
    End With
End Sub

Private Sub lstLabelItems_DblClick()
    Call cmdLabel_Click(0)
End Sub

Private Sub lstLabelUsed_Click()
    With Me.lstLabelUsed
        If .ListIndex = -1 Then Exit Sub
        Me.chkLabelCrLf.Value = IIf(.ItemData(.ListIndex) = 0, vbUnchecked, vbChecked)
        Me.txtLabelPrefix.Text = Left(.List(.ListIndex), InStr(1, .List(.ListIndex), "{") - 1)
    End With
End Sub

Private Sub optAlign_Click(Index As Integer)
    Dim intAlign As Integer
    
    Select Case Index
    Case 0
        intAlign = flexAlignLeftCenter
        vfgThis.TextMatrix(7, vfgThis.Col) = vfgThis.TextMatrix(6, vfgThis.Col) & " "
    Case 1
        intAlign = flexAlignCenterCenter
        vfgThis.TextMatrix(7, vfgThis.Col) = " " & vfgThis.TextMatrix(6, vfgThis.Col) & " "
    Case 2
        intAlign = flexAlignRightCenter
        vfgThis.TextMatrix(7, vfgThis.Col) = " " & vfgThis.TextMatrix(6, vfgThis.Col)
    End Select
    vfgThis.Cell(flexcpAlignment, 6, vfgThis.Col, 7, vfgThis.Col) = intAlign
    DataChanged = True
    
    On Error Resume Next
    chk.SetFocus
End Sub

Private Sub optOrient_Click(Index As Integer)
    DataChanged = True
End Sub

Private Sub optPageAlign_Click(Index As Integer)
    mblnChanged = True
End Sub

Private Sub optTabTiers_Click(Index As Integer)
    
    With vfgThis
        If optTabTiers(0).Value Then
            If .Row = 5 Or .Row = 4 Then .Row = 3
            .RowHidden(3) = False
            .RowHidden(4) = True
            .RowHidden(5) = True
            udHeadRow.Min = 1
            udHeadRow.Max = 1
        ElseIf optTabTiers(1).Value Then
            If .Row = 5 Then .Row = 4
            .RowHidden(3) = False
            .RowHidden(4) = False
            .RowHidden(5) = True
            udHeadRow.Min = 1
            udHeadRow.Max = 2
        Else
            .RowHidden(3) = False
            .RowHidden(4) = False
            .RowHidden(5) = False
            udHeadRow.Min = 1
            udHeadRow.Max = 3
        End If

    End With
    DataChanged = True
End Sub

Private Sub optTabTiers_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub picImg_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strInfo As String
    If Index = 0 Then
        strInfo = "��ܰ��ʾ��"
        strInfo = strInfo & vbCrLf & "1.ֻ��԰���һ��������Ŀ���С�"
        strInfo = strInfo & vbCrLf & "2.������Ŀֻ��ѡ�����һ����Ŀ���У����Ұ󶨵���ĿΪ����С��100���ı���Ŀ��ѡ��Ŀ��"
        strInfo = strInfo & vbCrLf & "3.���������Ŀ�������п�ѡ����Ŀ��"
    ElseIf Index = 1 Then
        strInfo = "��ܰ��ʾ��"
        strInfo = strInfo & vbCrLf & "1.ֻ��԰�һ���ǹ̶���Ŀ����"
        strInfo = strInfo & vbCrLf & "2.������Ŀ��֮��ͨ����������,�Զ�������а�ҽ������Ŀ,�󶨵�ҽ����ĿҪ���趨����Ŀ������ͬ(�������ܸ����������ֵ��Ŀ)"
        strInfo = strInfo & vbCrLf & "3.���������Ŀ�������п�ѡ����Ŀ��"
    End If
    Call zlCommFun.ShowTipInfo(picImg(Index).hWnd, strInfo, True)
End Sub

Private Sub picPane_Resize(Index As Integer)
    
    On Error Resume Next
    
    Select Case Index
    Case 0
        tbcStyle.Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
'        cmdOK.Left = tbcStyle.Left + tbcStyle.Width + 90
'        cmdCancel.Left = tbcStyle.Left + tbcStyle.Width + 90
'        imgNote.Left = tbcStyle.Left + tbcStyle.Width + 90
'        lblNote.Left = tbcStyle.Left + tbcStyle.Width + 90
    Case 1
        vfgThis.Move 15, 15, picPane(Index).Width - 30, picPane(Index).Height - 30
    End Select
End Sub

Private Sub picSum_Resize()
    On Error Resume Next
    
    vsf.Move 15, 15, picSum.Width - 30, picSum.Height - 30
    mclsVsf.AppendRows = True
End Sub

Private Sub rtbFoot_Change()
    mblnChanged = True
End Sub

Private Sub rtbFoot_GotFocus()
    mblnRTBFoot = True
End Sub

Private Sub rtbHead_Change()
    mblnChanged = True
End Sub

Private Sub rtbHead_GotFocus()
    mblnRTBFoot = False
End Sub

Private Sub tbcStyle_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If Item.Caption = "ҳ���ʽ" Then
        On Error Resume Next
        
        Printer.PaperSize = cboPaperKind.ItemData(cboPaperKind.ListIndex)
        Printer.Orientation = IIf(optOrient(0).Value, 1, 2)
        If Printer.PaperSize = 256 Then
            Call SetCustonPager(Me.hWnd, Int(Me.ScaleY(Val(Me.txtWidth.Text), vbMillimeters, vbTwips)), Int(Me.ScaleY(Val(Me.txtHeight.Text), vbMillimeters, vbTwips)))
        End If
        Call SendMessage(rtbHead.hWnd, EM_SETTARGETDEVICE, Me.hDC, ByVal CLng(Printer.ScaleWidth))
        SendMessageLong rtbHead.hWnd, EM_HIDESELECTION, 0, 0
        Call SendMessage(rtbFoot.hWnd, EM_SETTARGETDEVICE, Me.hDC, ByVal CLng(Printer.ScaleWidth))
        SendMessageLong rtbFoot.hWnd, EM_HIDESELECTION, 0, 0
    
        rtbHead.Width = picFoot.Width - 500
        rtbFoot.Width = rtbHead.Width
    End If
End Sub

Private Sub txtColumnPostfix_Change()
    With Me.lstColumnUsed
        If .ListIndex = -1 Then Exit Sub
        .List(.ListIndex) = Left(.List(.ListIndex), InStr(1, .List(.ListIndex), "}")) & Me.txtColumnPostfix.Text
    End With
End Sub

Private Sub txtColumnPostfix_GotFocus()
    Me.txtColumnPostfix.SelStart = 0: Me.txtColumnPostfix.SelLength = 4000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtColumnPostfix_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" &[]{}+'""|", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtColumnPrefix_Change()
    With Me.lstColumnUsed
        If .ListIndex = -1 Then Exit Sub
        .List(.ListIndex) = Me.txtColumnPrefix.Text & Mid(.List(.ListIndex), InStr(1, .List(.ListIndex), "{"))
    End With
End Sub

Private Sub txtColumnPrefix_GotFocus()
    Me.txtColumnPrefix.SelStart = 0: Me.txtColumnPrefix.SelLength = 4000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtColumnPrefix_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" &[]{}+'""|", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtHeadRow_Change()
    DataChanged = True
End Sub

Private Sub txtHeadText_Change()
    Dim strInput As String
    Dim lngStart As Long, lngRow As Long, lngCol As Long
    Dim blnExist As Boolean
    
    strInput = Trim(Me.txtHeadText.Text)
    '���,��������ڵ��ĸ���Ԫ���ֵ��ͬ,����������(�п�����Ҫ������,����,����,���½��м��)
    lngRow = udHeadRow.Value + 2
    lngCol = Me.udHeadCol.Value
    Me.vfgThis.TextMatrix(Me.udHeadRow.Value + 2, lngCol) = strInput
    
    If lngRow <= 4 Then
        If (vfgThis.TextMatrix(3, lngCol) = vfgThis.TextMatrix(4, lngCol) And vfgThis.TextMatrix(3, lngCol) <> "") Then
            If lngCol > 1 Then
                '����
                If vfgThis.TextMatrix(lngRow, lngCol - 1) = vfgThis.TextMatrix(lngRow, lngCol) And vfgThis.TextMatrix(lngRow + IIf(lngRow = 3, 1, -1), lngCol - 1) = vfgThis.TextMatrix(lngRow, lngCol) Then
                    blnExist = True
                    GoTo Limit
                End If
            End If
            If lngCol < vfgThis.Cols - 1 Then
                '����
                If vfgThis.TextMatrix(lngRow, lngCol) = vfgThis.TextMatrix(lngRow, lngCol + 1) And vfgThis.TextMatrix(lngRow + IIf(lngRow = 3, 1, -1), lngCol + 1) = vfgThis.TextMatrix(lngRow, lngCol) Then
                    blnExist = True
                    GoTo Limit
                End If
            End If
        End If
    End If
    If lngRow >= 4 Then
        If (vfgThis.TextMatrix(4, lngCol) = vfgThis.TextMatrix(5, lngCol) And vfgThis.TextMatrix(4, lngCol) <> "") Then
            If lngCol > 1 Then
                '����
                If vfgThis.TextMatrix(lngRow, lngCol - 1) = vfgThis.TextMatrix(lngRow, lngCol) And vfgThis.TextMatrix(lngRow + IIf(lngRow = 4, 1, -1), lngCol - 1) = vfgThis.TextMatrix(lngRow, lngCol) Then
                    blnExist = True
                    GoTo Limit
                End If
            End If
            If lngCol < vfgThis.Cols - 1 Then
                '����
                If vfgThis.TextMatrix(lngRow, lngCol) = vfgThis.TextMatrix(lngRow, lngCol + 1) And vfgThis.TextMatrix(lngRow + IIf(lngRow = 4, 1, -1), lngCol + 1) = vfgThis.TextMatrix(lngRow, lngCol) Then
                    blnExist = True
                    GoTo Limit
                End If
            End If
        End If
    End If
Limit:
    If blnExist Then strInput = strInput & "_1"

WriteIt:
    Me.txtHeadText.Text = strInput
    Me.vfgThis.TextMatrix(Me.udHeadRow.Value + 2, lngCol) = strInput
    vfgThis.AutoSize 0, vfgThis.Cols - 1
    Call cmdLabel_Click(2)
    DataChanged = True
End Sub

Private Sub txtHeadText_GotFocus()
    Me.txtHeadText.SelStart = 0: Me.txtHeadText.SelLength = 4000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtHeadText_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*()[]{}_+|=-`;'"":\.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtHeight_Change()
    DataChanged = True
End Sub

Private Sub txtLabelPrefix_Change()
    With Me.lstLabelUsed
        If .ListIndex = -1 Then Exit Sub
        .List(.ListIndex) = Me.txtLabelPrefix.Text & Mid(.List(.ListIndex), InStr(1, .List(.ListIndex), "{"))
    End With
End Sub

Private Sub txtLabelPrefix_GotFocus()
    Me.txtLabelPrefix.SelStart = 0: Me.txtLabelPrefix.SelLength = 4000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtLabelPrefix_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" &[]{}+'""|", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtMarjin_Change(Index As Integer)
    DataChanged = True
End Sub

Private Sub txtRecordFrom_Change()
    DataChanged = True
End Sub

Private Sub txtRecordTo_Change()
    DataChanged = True
End Sub

Private Sub txtTabCols_Change()
    DataChanged = True
End Sub

Private Sub txtTabRowHeight_Change()
    Me.vfgThis.RowHeightMin = Val(Me.txtTabRowHeight.Text)
    DataChanged = True
End Sub

Private Sub txtTabRowHeight_GotFocus()
    Me.txtTabRowHeight.SelStart = 0: Me.txtTabRowHeight.SelLength = 100
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtTabRowHeight_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtTitleText_Change()
    With Me.vfgThis
        For lngCount = .FixedCols To .Cols - 1
            .TextMatrix(1, lngCount) = Trim(Me.txtTitleText.Text)
        Next
    End With
    
    DataChanged = True
End Sub

Private Sub txtTitleText_GotFocus()
    Me.txtTitleText.SelStart = 0: Me.txtTitleText.SelLength = 4000
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txtTitleText_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr("~!@#$%^&*()[]{}_+|=-`;'"":/\.,<>?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txtWidth_Change()
    DataChanged = True
End Sub

Private Sub udColumnNo_Change()
    Me.lstColumnUsed.Clear
'    strTemp = Trim(Me.vfgThis.TextMatrix(5, Me.udColumnNo.Value))
'
    strTemp = vfgThis.Cell(flexcpData, 6, udColumnNo.Value, 6, udColumnNo.Value)
    
    If strTemp = "" Then
        Me.lstColumnUsed.ListIndex = -1
        Me.cmdColumn(1).Enabled = False
        Me.txtColumnPrefix.Enabled = False: Me.txtColumnPrefix.Text = ""
        Me.txtColumnPostfix.Enabled = False: Me.txtColumnPostfix.Text = ""
    Else
        Dim aryCol() As String
        aryCol = Split(strTemp, Space(1))
        For lngCount = 0 To UBound(aryCol)
            If InStr(aryCol(lngCount), "`") > 0 Then
                Me.lstColumnUsed.AddItem Mid(aryCol(lngCount), 1, InStr(aryCol(lngCount), "`") - 1)
                Me.lstColumnUsed.ItemData(lstColumnUsed.NewIndex) = Val(Mid(aryCol(lngCount), InStr(aryCol(lngCount), "`") + 1))
            Else
                Me.lstColumnUsed.AddItem aryCol(lngCount)
                Me.lstColumnUsed.ItemData(lstColumnUsed.NewIndex) = 0
            End If
        Next
        Me.lstColumnUsed.ListIndex = 0
        Me.cmdColumn(1).Enabled = True
        Me.txtColumnPrefix.Enabled = True
        Me.txtColumnPostfix.Enabled = True
        
        chk.Enabled = True
        
    End If
End Sub

Private Sub udHeadCol_Change()
    
    Dim blnSvrChanged As Boolean
    
    blnSvrChanged = DataChanged
    
    txtHeadText.Text = vfgThis.TextMatrix(udHeadRow.Value + 2, udHeadCol.Value)

    DataChanged = blnSvrChanged
End Sub

Private Sub udHeadRow_Change()

    Call udHeadCol_Change
    
End Sub

Private Sub udRecordFrom_Change()
    If Me.udRecordFrom.Value > Me.udRecordTo.Value Then
        Me.lblRecordTo.Caption = "����" & Space(7) & "��"
    Else
        Me.lblRecordTo.Caption = "����" & Space(7) & "��"
    End If
End Sub

Private Sub udRecordTo_Change()
    Call udRecordFrom_Change
End Sub

Private Sub udTabCols_Change()
    Me.vfgThis.Cols = Me.udTabCols.Value + 1
    Me.vfgThis.MergeCol(Me.vfgThis.Cols - 1) = True
    Me.udHeadCol.Max = Me.udTabCols.Value
    If Val(Me.txtHeadCol.Text) > Me.udHeadCol.Max Then Me.txtHeadCol.Text = Me.udHeadCol.Max
    Me.udColumnNo.Max = Me.udTabCols.Value
    If Val(Me.txtColumnNo.Text) > Me.udColumnNo.Max Then Me.txtColumnNo.Text = Me.udColumnNo.Max
    
    With Me.vfgThis
        For lngCount = .FixedCols To .Cols - 1
            .TextMatrix(0, lngCount) = lngCount
            .ColAlignment(lngCount) = flexAlignCenterCenter
            .FixedAlignment(lngCount) = flexAlignCenterCenter
            .TextMatrix(1, lngCount) = .TextMatrix(1, .FixedCols)
            .TextMatrix(2, lngCount) = .TextMatrix(2, .FixedCols)
            .TextMatrix(.Rows - 1, lngCount) = .TextMatrix(.Rows - 1, .FixedCols)
        Next
        '.Cell(flexcpAlignment, 6, 1, 7, .Cols - 1) = flexAlignGeneralCenter
        .CellBorderRange 2, .FixedCols, 2, .Cols - 1, Me.shpTabGridColor.BorderColor, 0, 0, 0, 1, 0, 0
        .CellBorderRange 3, .FixedCols, 7, .Cols - 1, Me.shpTabGridColor.BorderColor, 1, 1, 1, 1, 1, 1
        .Cell(flexcpForeColor, 7, 1, 7, .Cols - 1) = Me.lblRecordColor.ForeColor
    End With
End Sub

Private Sub vfgColRelation_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    vfgColRelation.TextMatrix(Row, Col) = vfgColRelation.EditText
End Sub

Private Sub vfgColRelation_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
    With vfgColRelation
        If Row >= .FixedRows And Col = .ColIndex("��������") And .ComboIndex >= 0 Then
            FinishEdit = True
        End If
    End With
End Sub

Private Sub vfgColRelation_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vfgColRelation
        If Row >= .FixedRows And Col = .ColIndex("��������") Then
            If .ColComboList(Col) = "" Then
                Cancel = True
            End If
        Else
            Cancel = True
        End If
    End With
End Sub

Private Sub vfgColRelation_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim lngRow As Long
    Dim blnCancle As Boolean
    With vfgColRelation
        If Row >= .FixedRows And Col = .ColIndex("��������") And .ColComboList(.ColIndex("��������")) <> "" Then
            If Trim(vfgColRelation.EditText) = Trim(.TextMatrix(Row, .ColIndex("��������"))) Then Exit Sub
            If .ComboIndex < 0 Or Trim(vfgColRelation.EditText) = "" Then
                .TextMatrix(Row, .ColIndex("�����к�")) = ""
                DataChanged = True
                Exit Sub
            End If
            For lngRow = .FixedRows To .Rows - 1
                If lngRow <> Row Then
                    If .TextMatrix(lngRow, .ColIndex("��������")) = vfgColRelation.EditText Then
                        blnCancle = True
                        Exit For
                    End If
                End If
            Next lngRow
            If blnCancle = True Then
                MsgBox "��ѡ��Ķ�����Ŀ�Ѿ�������������Ŀ�����˶��չ�ϵ��������ѡ��!", vbInformation, gstrSignName
                Cancel = True
            Else
                .TextMatrix(Row, .ColIndex("�����к�")) = .ComboData(.ComboIndex)
                .TextMatrix(Row, .ColIndex("��������")) = vfgColRelation.EditText
                DataChanged = True
            End If
        End If
    End With
End Sub

Private Sub vfgThis_AfterMoveColumn(ByVal Col As Long, Position As Long)
    Dim lngCol As Long
    If Col >= Position Then
        For lngCol = Position To Col
            vfgThis.TextMatrix(0, lngCol) = lngCol
        Next lngCol
    Else
        For lngCol = Col To Position
            vfgThis.TextMatrix(0, lngCol) = lngCol
        Next lngCol
    End If
    Call InstallColRelation
    DataChanged = True
End Sub

Private Sub vfgThis_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    Dim blnSvrChanged As Boolean
    
    blnSvrChanged = DataChanged
    
    Me.udHeadCol.Value = NewCol
    Me.udColumnNo.Value = NewCol
    If NewRow >= 3 And NewRow <= 5 Then udHeadRow.Value = NewRow - 2
    
    Dim intAlign As Integer
    Select Case vfgThis.Cell(flexcpAlignment, 6, NewCol)
    Case Is >= 6    '��
        intAlign = 2
    Case Is >= 2    '��
        intAlign = 1
    Case Is <= 2    '��
        intAlign = 0
    End Select
    optAlign(intAlign).Value = True
    
    DataChanged = blnSvrChanged
End Sub

Private Sub vfgThis_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    vfgThis.AutoSize 0, vfgThis.Cols - 1
    Call cmdLabel_Click(2)
    DataChanged = True
End Sub

Private Sub GetrtbObject()
    If mblnRTBFoot Then
        Set rtbThis = rtbFoot
    Else
        Set rtbThis = rtbHead
    End If
End Sub

Private Sub cmdOpen_Click()
    Dim picTemp As StdPicture
    
    With Me.dlgThis
        .DialogTitle = "��־ͼѡ��"
        .Filename = ""
        .Filter = "ͼ��|*.jpg;*.bmp;*.ico;*.gif"
        .CancelError = True
        On Error Resume Next
        .ShowOpen
        If Err.Number <> 0 Then
            Err.Clear
            Exit Sub
        End If
    End With
    Set picTemp = Nothing
    Set picTemp = LoadPicture(Me.dlgThis.Filename)
    If picTemp Is Nothing Then MsgBox "������Ч��ͼƬ�ļ���", vbExclamation, Me.Caption: Exit Sub
    
    Clipboard.Clear
    Clipboard.SetData picTemp
    
    Call GetrtbObject
    SendMessageLong rtbThis.hWnd, WM_PASTE, 0, 0
    mblnChanged = True
End Sub

Private Sub cboFont_Click()
    Call GetrtbObject
    rtbThis.SelFontName = cboFont.List(cboFont.ListIndex)
    mblnChanged = True
End Sub

Private Sub cboFSize_Click()
    Call GetrtbObject
    rtbThis.SelFontSize = GetFontSizeNumber(cboFSize.List(cboFSize.ListIndex))
    mblnChanged = True
End Sub

Private Sub chkB_Click()
    Call GetrtbObject
    If chkB.Value = vbChecked Then
        rtbThis.SelBold = True
    Else
        rtbThis.SelBold = True
    End If
    mblnChanged = True
End Sub

Private Sub chkI_Click()
    Call GetrtbObject
    If chkI.Value = vbChecked Then
        rtbThis.SelItalic = True
    Else
        rtbThis.SelItalic = False
    End If
    mblnChanged = True
End Sub

Private Sub chkU_Click()
    Call GetrtbObject
    If chkU.Value = vbChecked Then
        rtbThis.SelUnderline = True
    Else
        rtbThis.SelUnderline = False
    End If
    mblnChanged = True
End Sub

Private Function PageHeadTest() As Boolean
    '�����ϱ߾෵�ؼ�
    Dim fr As FORMATRANGE           '��ʽ�����ı���Χ
    Dim rcDrawTo As RECT            'Ŀ����������
    Dim rcPage As RECT              'Ŀ��ҳ������
    Dim gTargetDC As Long
    Dim lngOffsetLeft As Long
    Dim lngOffsetTop As Long
'    Dim lngOffsetWidth As Long
'    Dim lngOffsetHeight As Long
    Dim lngNextPos As Long, lngLen As Long, lngTMP As Long, lngPageCount As Long
    Dim blnTrue As Boolean
    
    lngLen = lstrlen(rtbHead.Text)
    'printer.Duplex = vbPRDPHorizontal
    'printer.ScaleMode = vbTwips
    lngOffsetLeft = Printer.ScaleX(GetDeviceCaps(Printer.hDC, PHYSICALOFFSETX), vbPixels, vbTwips)
    lngOffsetTop = Printer.ScaleY(GetDeviceCaps(Printer.hDC, PHYSICALOFFSETY), vbPixels, vbTwips)
'    lngOffsetWidth = Printer.ScaleWidth
'    lngOffsetHeight = Printer.ScaleHeight
    
    gTargetDC = hDC
    With rcPage
        .Left = 0
        .Top = 0
        .Right = Printer.Width
        .Bottom = Printer.Height
    End With
    With rcDrawTo
        .Left = lngOffsetLeft
        .Top = lngOffsetTop
        .Right = Printer.Width - lngOffsetLeft
        .Bottom = Printer.ScaleX(txtMarjin(0).Text, vbMillimeters, vbTwips)
        '46251,������,2012-09-11
        If Val(cboҳ��.ItemData(cboҳ��.ListIndex)) = 1 Or Val(cboҳ��.ItemData(cboҳ��.ListIndex)) = 2 Then
            .Bottom = .Bottom - Printer.TextHeight("1")
            blnTrue = True
        End If
    End With
    With fr
        .hDC = Printer.hDC
        .hdcTarget = gTargetDC
        .rc = rcDrawTo
        .rcPage = rcPage
        .chrg.cpMin = 0
        .chrg.cpMax = -1
    End With
    
    Do
        lngNextPos = SendMessage(rtbHead.hWnd, EM_FORMATRANGE, 0, fr)
        
        lngPageCount = lngPageCount + 1             ' ҳ����1
        '��¼��ҳ��Ϣ
        ReDim Preserve AllPages(1 To lngPageCount) As PageInfo
        AllPages(lngPageCount).PageNumber = lngPageCount
        AllPages(lngPageCount).ActualHeight = fr.rc.Bottom - fr.rc.Top          'ʵ�ʴ�ӡ�߶�
        AllPages(lngPageCount).Start = lngTMP
        AllPages(lngPageCount).End = lngNextPos
        
        fr.chrg.cpMin = lngNextPos
        If lngNextPos <= lngTMP Or lngNextPos >= lngLen Then Exit Do      ' �������ҳ��ķ�ҳ
        lngTMP = lngNextPos
    Loop
    Call SendMessage(rtbHead.hWnd, EM_FORMATRANGE, 0, ByVal CLng(0))
    
    If fr.rc.Bottom > rcDrawTo.Bottom Or lngPageCount > 1 Then
        MsgBox "��Ƶ�ҳü����" & IIf(blnTrue = True, "��ҳ��߶�", "") & "�������ϱ߾࣡", vbInformation, gstrSysName
        Exit Function
    End If
    PageHeadTest = True
End Function

Private Function ReadPageHead(objHead As RichTextBox, ByVal StrKey As String) As Boolean
'################################################################################################################
'## ���ܣ�  ��ȡҳ��ͼƬ
'## ������  ��������-ҳ����
'## ���أ�  ���ػ�õ�ͼƬ������
'################################################################################################################
    Dim strFile As String, strZip As String
    strZip = zlBlobRead(12, StrKey, App.Path & "\Head_L.zip")
    If gobjFSO.FileExists(strZip) Then
        strFile = UnzipTendPage(strZip, "Head_S.RTF")
        objHead.LoadFile strFile, rtfRTF           '��ȡ�ļ�
        gobjFSO.DeleteFile strFile, True      'ɾ����ʱ�ļ�
        ReadPageHead = True
    Else
        objHead.Text = ""
    End If
End Function

Private Function ReadPageFoot(objFoot As RichTextBox, ByVal StrKey As String) As Boolean
'################################################################################################################
'## ���ܣ�  ��ȡҳ��ͼƬ
'## ������  ��������-ҳ����
'## ���أ�  ���ػ�õ�ͼƬ������
'################################################################################################################
    Dim strFile As String, strZip As String
    strZip = zlBlobRead(13, StrKey, App.Path & "\Foot_L.zip")
    If gobjFSO.FileExists(strZip) Then
        strFile = UnzipTendPage(strZip, "Foot_S.RTF")
        objFoot.LoadFile strFile, rtfRTF           '��ȡ�ļ�
        gobjFSO.DeleteFile strFile, True      'ɾ����ʱ�ļ�
        ReadPageFoot = True
    Else
        objFoot.Text = ""
    End If
End Function

Private Function ReadPageHeadFile(ByVal StrKey As String) As String
'################################################################################################################
'## ���ܣ�  ��ȡҳ��ͼƬ
'## ������  ��������-ҳ����
'## ���أ�  ���ػ�õ�ͼƬ������
'################################################################################################################
    Dim strZip As String
    strZip = zlBlobRead(12, StrKey, App.Path & "\Head_L.zip")
    If gobjFSO.FileExists(strZip) Then
        ReadPageHeadFile = strZip
    End If
End Function

Private Function ReadPageFootFile(ByVal StrKey As String) As String
'################################################################################################################
'## ���ܣ�  ��ȡҳ��ͼƬ
'## ������  ��������-ҳ����
'## ���أ�  ���ػ�õ�ͼƬ������
'################################################################################################################
    Dim strZip As String
    strZip = zlBlobRead(13, StrKey, App.Path & "\Foot_L.zip")
    If gobjFSO.FileExists(strZip) Then
        ReadPageFootFile = strZip
    End If
End Function

'################################################################################################################
'## ���ܣ�  ��ѹ���ļ���ͬĿ¼�ͷŲ�����ѹ�ļ�
'## ������  strZipFile     :ѹ���ļ�
'## ���أ�  ��ѹ�ļ�����ʧ���򷵻��㳤��""
'################################################################################################################
Private Function UnzipTendPage(ByVal strZipFile As String, ByVal strTarFile As String) As String
    Dim strZipPathTmp As String
    Dim strZipPath As String
    Dim strZipFileTmp As String
    Dim strZipFileName As String
    Dim mclsUnzip As New cUnzip
    
    On Error GoTo errHand
    
    If Not gobjFSO.FileExists(strZipFile) Then UnzipTendPage = "": Exit Function
    strZipPath = Left(strZipFile, Len(strZipFile) - Len(Dir(strZipFile)))
    
    strZipPath = gobjFSO.GetSpecialFolder(2)
    strZipPathTmp = strZipPath & Format(Now, "yyMMddHHmmss") & CStr(100 * Timer)
    Call gobjFSO.CreateFolder(strZipPathTmp)
    
    strZipFileTmp = strZipPathTmp ' & "\TMP.RTF"
    
    With mclsUnzip
        .ZipFile = strZipFile
        .UnzipFolder = strZipPathTmp
        .Unzip
    End With
    If gobjFSO.FolderExists(strZipFileTmp) Then
        
        strZipFileName = gobjFSO.GetFile(strZipFileTmp & "\" & strTarFile)
        Call gobjFSO.CopyFile(strZipFileName, "C:\" & strTarFile)
        
        On Error Resume Next
        gobjFSO.DeleteFolder strZipPathTmp, True
        gobjFSO.DeleteFile strZipFile, True
        
        UnzipTendPage = "C:\" & strTarFile
    Else
        UnzipTendPage = ""
    End If
    
    Exit Function
    
errHand:
    Call SaveErrLog
End Function

Private Sub vsf_AfterScroll(ByVal OldTopRow As Long, ByVal OldLeftCol As Long, ByVal NewTopRow As Long, ByVal NewLeftCol As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    mclsVsf.AppendRows = True
End Sub

Private Sub vsf_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    
    '�༭����
    Call mclsVsf.AfterEdit(Row, Col)
    DataChanged = True
End Sub

Private Sub vsf_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    
    '�༭����
    Call mclsVsf.AfterRowColChange(OldRow, OldCol, NewRow, NewCol)
End Sub

Private Sub vsf_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call mclsVsf.BeforeResizeColumn(Col, Cancel)
End Sub

Private Sub vsf_DblClick()
    '�༭����
    Call mclsVsf.DbClick
End Sub

Private Sub vsf_KeyDown(KeyCode As Integer, Shift As Integer)
    '�༭����
    Call mclsVsf.KeyDown(KeyCode, Shift)
End Sub

Private Sub vsf_KeyPress(KeyAscii As Integer)
    
    '�༭����,������
    If KeyAscii <> vbKeyReturn Then
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
        
        Select Case vsf.Col
        Case vsf.ColIndex("��ʼʱ��"), vsf.ColIndex("����ʱ��")
            If FilterKeyAscii(KeyAscii, 99, "0123456789:") = 0 Then KeyAscii = 0
        End Select
        
        
    End If
    Call mclsVsf.KeyPress(KeyAscii)

End Sub

Private Sub vsf_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    '�༭����
    
    If KeyAscii <> vbKeyReturn Then
        If Chr(KeyAscii) = "'" Then KeyAscii = 0
        
        Select Case vsf.Col
        Case vsf.ColIndex("��ʼʱ��"), vsf.ColIndex("����ʱ��")
            If FilterKeyAscii(KeyAscii, 99, "0123456789:") = 0 Then KeyAscii = 0
        End Select

    End If
    Call mclsVsf.KeyPressEdit(KeyAscii)
    
End Sub

Private Sub vsf_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Select Case Button
    Case 1
        Call mclsVsf.AutoAddRow(vsf.MouseRow, vsf.MouseCol)
    End Select
End Sub

Private Sub vsf_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    '�༭����
    Call mclsVsf.EditSelAll
End Sub

Private Sub vsf_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    '�༭����
    Call mclsVsf.BeforeEdit(Row, Col, Cancel)
End Sub

Private Sub vsf_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim StrText As String
    Call mclsVsf.ValidateEdit(Col, Cancel)
    StrText = vsf.EditText
    If Col = 2 Or Col = 3 Then
        Cancel = Not CheckTime(Row, Col, StrText)
    End If
    vsf.EditText = StrText
End Sub

Private Sub vsfColDoctor_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    vsfColDoctor.TextMatrix(Row, Col) = vsfColDoctor.EditText
End Sub

Private Sub vsfColDoctor_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
    With vsfColDoctor
        If Row >= .FixedRows And Col = .ColIndex("��������") And .ComboIndex >= 0 Then
            FinishEdit = True
        End If
    End With
End Sub

Private Sub vsfColDoctor_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strName As String
    Dim lngRow As Long
    Dim blnCancle As Boolean
    
    strName = vsfColDoctor.TextMatrix(Row, Col - 1)
    rsItems.Filter = ""
    If Not rsItems.RecordCount > 0 Then Exit Sub
    With vsfColDoctor
        If .ColComboList(.ColIndex("��������")) <> "" Then
            If Trim(vsfColDoctor.EditText) = Trim(.TextMatrix(Row, .ColIndex("��������"))) Then Exit Sub
        End If
        If .ComboIndex < 0 Or Trim(vsfColDoctor.EditText) = "" Then
            DataChanged = True
            Exit Sub
        End If
        
        For lngRow = .FixedRows To .Rows - 1
            If lngRow <> Row Then
                If .TextMatrix(lngRow, .ColIndex("��������")) = vsfColDoctor.EditText Then
                    blnCancle = True
                    Exit For
                End If
            End If
        Next lngRow
        
        
        If blnCancle = True Then
                MsgBox "��ѡ��Ķ�����Ŀ�Ѿ�������������Ŀ�����˶��չ�ϵ��������ѡ��!", vbInformation, gstrSignName
                Cancel = True
                Exit Sub
        End If
            
        Select Case vsfColDoctor.EditText
            Case "����", "�ܸ�����"
                rsItems.Filter = "��Ŀ����='" & strName & "'"
                If Not rsItems.EOF Then
                    If rsItems!��Ŀ���� <> 0 Then
                        Cancel = True
                        MsgBox "��ҽ������Ŀʱ����Ŀ�����ͱ���һ�£�����Ŀ���ı���Ŀ��", vbInformation, gstrSysName
                        Cancel = True
                    Else
                        DataChanged = True
                    End If
                    
                Else
                    vsfColDoctor.EditText = " "
                End If
            Case Else
                rsItems.Filter = "��Ŀ����='" & strName & "'"
                If Not rsItems.EOF Then
                    If rsItems!��Ŀ���� <> 1 Then
                        Cancel = True
                        MsgBox "��ҽ������Ŀʱ����Ŀ�����ͱ���һ�£�����Ŀ����ֵ��Ŀ��", vbInformation, gstrSysName
                        Cancel = True
                    Else
                        DataChanged = True
                    End If
                Else
                    vsfColDoctor.EditText = " "
                End If
        End Select
    End With
End Sub



