VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Object = "{555E8FCC-830E-45CC-AF00-A012D5AE7451}#9.60#0"; "Codejock.CommandBars.9600.ocx"
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmParPublic 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "������������"
   ClientHeight    =   8235
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   12690
   Icon            =   "frmParPublic.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8235
   ScaleWidth      =   12690
   StartUpPosition =   1  '����������
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7575
      Index           =   0
      Left            =   2400
      ScaleHeight     =   7545
      ScaleWidth      =   10245
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   10275
      Begin VB.Frame fraDevSvr 
         Caption         =   "������������"
         Height          =   2730
         Left            =   345
         TabIndex        =   195
         Top             =   4755
         Width           =   8265
         Begin VB.CommandButton cmdSvrChk 
            Caption         =   "������֤"
            Height          =   300
            Left            =   6900
            TabIndex        =   196
            Top             =   2295
            Width           =   1260
         End
         Begin VSFlex8Ctl.VSFlexGrid vsThirdSvr 
            Height          =   1935
            Left            =   105
            TabIndex        =   197
            Top             =   285
            Width           =   8055
            _cx             =   14208
            _cy             =   3413
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
            BackColorSel    =   16777215
            ForeColorSel    =   0
            BackColorBkg    =   16777215
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483637
            GridColorFixed  =   16777215
            TreeColor       =   16777215
            FloodColor      =   192
            SheetBorder     =   16777215
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   4
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   0
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmParPublic.frx":6852
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
      Begin VB.CheckBox chk 
         Caption         =   "ҽ�ƻ�������������¼��"
         Height          =   255
         Index           =   25
         Left            =   390
         TabIndex        =   182
         Top             =   4470
         Width           =   2535
      End
      Begin VB.CheckBox chk 
         Caption         =   "��Ժҽ�������Ƚ���"
         Height          =   255
         Index           =   22
         Left            =   5895
         TabIndex        =   170
         Top             =   4455
         Width           =   2175
      End
      Begin VB.CheckBox chk 
         Caption         =   "���˵�ַ�ṹ��¼��"
         Height          =   255
         Index           =   50
         Left            =   480
         TabIndex        =   168
         Top             =   3795
         Width           =   2055
      End
      Begin VB.Frame fraSTAddress 
         Height          =   645
         Left            =   375
         TabIndex        =   167
         Top             =   3795
         Width           =   3975
         Begin VB.CheckBox chk 
            Caption         =   "���򼶵�ַ�ṹ��¼��"
            Enabled         =   0   'False
            Height          =   255
            Index           =   51
            Left            =   390
            TabIndex        =   169
            Top             =   300
            Width           =   2535
         End
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   15
         Left            =   1890
         Style           =   2  'Dropdown List
         TabIndex        =   2
         Top             =   1980
         Width           =   2445
      End
      Begin VB.Frame fra����¼�� 
         Caption         =   "����¼������"
         Height          =   1020
         Left            =   360
         TabIndex        =   11
         Top             =   2700
         Width           =   3975
         Begin VB.TextBox txtUD 
            Height          =   300
            Index           =   0
            Left            =   1550
            MaxLength       =   4
            TabIndex        =   13
            Top             =   300
            Width           =   650
         End
         Begin VB.CheckBox chk 
            Caption         =   "ת��������ֻ����¼����"
            Height          =   210
            Index           =   0
            Left            =   360
            TabIndex        =   15
            Top             =   720
            Value           =   1  'Checked
            Width           =   2520
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   300
            Index           =   0
            Left            =   2220
            TabIndex        =   14
            TabStop         =   0   'False
            Top             =   300
            Width           =   255
            _ExtentX        =   450
            _ExtentY        =   529
            _Version        =   393216
            Value           =   1
            BuddyControl    =   "txtUD(0)"
            BuddyDispid     =   196637
            BuddyIndex      =   0
            OrigLeft        =   2470
            OrigTop         =   315
            OrigRight       =   2725
            OrigBottom      =   585
            Max             =   9999
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.Label lblInputHours 
            AutoSize        =   -1  'True
            Caption         =   "Сʱ"
            Height          =   180
            Left            =   2520
            TabIndex        =   148
            Top             =   360
            Width           =   360
         End
         Begin VB.Label lbl����¼�� 
            AutoSize        =   -1  'True
            Caption         =   "ʱ��(0-9999)"
            Height          =   180
            Left            =   360
            TabIndex        =   12
            Top             =   360
            Width           =   1080
         End
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   10
         ItemData        =   "frmParPublic.frx":68E1
         Left            =   1890
         List            =   "frmParPublic.frx":68E3
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   2340
         Width           =   2445
      End
      Begin VB.TextBox txtUD 
         Alignment       =   2  'Center
         Height          =   270
         IMEMode         =   3  'DISABLE
         Index           =   9
         Left            =   1860
         Locked          =   -1  'True
         MaxLength       =   2
         TabIndex        =   9
         Text            =   "12"
         Top             =   480
         Width           =   380
      End
      Begin VB.Frame Fra 
         Caption         =   " ������� "
         Height          =   1860
         Index           =   12
         Left            =   5895
         TabIndex        =   25
         Top             =   2400
         Width           =   4095
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   16
            ItemData        =   "frmParPublic.frx":68E5
            Left            =   960
            List            =   "frmParPublic.frx":68E7
            Style           =   2  'Dropdown List
            TabIndex        =   31
            Top             =   1260
            Width           =   2535
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   8
            ItemData        =   "frmParPublic.frx":68E9
            Left            =   960
            List            =   "frmParPublic.frx":68EB
            Style           =   2  'Dropdown List
            TabIndex        =   29
            Top             =   795
            Width           =   2535
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   1
            ItemData        =   "frmParPublic.frx":68ED
            Left            =   960
            List            =   "frmParPublic.frx":68EF
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   330
            Width           =   2535
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "סԺ"
            Height          =   180
            Index           =   51
            Left            =   405
            TabIndex        =   30
            Top             =   1320
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            Height          =   180
            Index           =   27
            Left            =   405
            TabIndex        =   28
            Top             =   855
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Դ"
            Height          =   180
            Index           =   39
            Left            =   405
            TabIndex        =   26
            Top             =   390
            Width           =   360
         End
      End
      Begin VB.CheckBox chk 
         Caption         =   "����ȫ������ʱֻ���ұ���"
         Height          =   195
         Index           =   10
         Left            =   1860
         TabIndex        =   6
         Top             =   1245
         Value           =   1  'Checked
         Width           =   2460
      End
      Begin VB.CheckBox chk 
         Caption         =   $"frmParPublic.frx":68F1
         Height          =   195
         Index           =   11
         Left            =   1860
         TabIndex        =   7
         Top             =   1530
         Value           =   1  'Checked
         Width           =   2460
      End
      Begin VB.Frame Fra 
         Caption         =   " �������°�ʱ�� "
         Height          =   1635
         Index           =   1
         Left            =   5895
         TabIndex        =   16
         Top             =   360
         Width           =   4095
         Begin MSComCtl2.DTPicker dtp 
            Height          =   315
            Index           =   0
            Left            =   1005
            TabIndex        =   18
            Top             =   420
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   105906179
            UpDown          =   -1  'True
            CurrentDate     =   36526.3541666667
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   315
            Index           =   1
            Left            =   2475
            TabIndex        =   20
            Top             =   420
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   105906179
            UpDown          =   -1  'True
            CurrentDate     =   36526.5
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   315
            Index           =   2
            Left            =   1005
            TabIndex        =   22
            Top             =   915
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   105906179
            UpDown          =   -1  'True
            CurrentDate     =   36526.5625
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   315
            Index           =   3
            Left            =   2475
            TabIndex        =   24
            Top             =   915
            Width           =   855
            _ExtentX        =   1508
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   105906179
            UpDown          =   -1  'True
            CurrentDate     =   36526.75
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "-"
            Height          =   180
            Index           =   5
            Left            =   2100
            TabIndex        =   23
            Top             =   990
            Width           =   90
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   195
            Index           =   3
            Left            =   435
            TabIndex        =   21
            Top             =   975
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "-"
            Height          =   180
            Index           =   4
            Left            =   2100
            TabIndex        =   19
            Top             =   495
            Width           =   90
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Index           =   2
            Left            =   435
            TabIndex        =   17
            Top             =   480
            Width           =   360
         End
      End
      Begin MSComCtl2.UpDown ud 
         Height          =   270
         Index           =   9
         Left            =   2265
         TabIndex        =   10
         TabStop         =   0   'False
         Top             =   480
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   476
         _Version        =   393216
         Value           =   2
         BuddyControl    =   "txtUD(9)"
         BuddyDispid     =   196637
         BuddyIndex      =   9
         OrigLeft        =   2205
         OrigTop         =   1200
         OrigRight       =   2460
         OrigBottom      =   1470
         Max             =   20
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "ҽ��������"
         Height          =   180
         Index           =   42
         Left            =   720
         TabIndex        =   1
         Top             =   2040
         Width           =   1080
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "������Ŀ�������"
         Height          =   180
         Index           =   36
         Left            =   360
         TabIndex        =   3
         Top             =   2400
         Width           =   1440
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "��ͯ����綨����         ��"
         Height          =   180
         Index           =   47
         Left            =   360
         TabIndex        =   8
         Top             =   525
         Width           =   2430
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "�շ���Ŀ��������Ŀ����ƥ�䷽ʽ"
         Height          =   180
         Index           =   40
         Left            =   360
         TabIndex        =   5
         Top             =   960
         Width           =   2700
      End
   End
   Begin VB.PictureBox picPar 
      Height          =   7575
      Index           =   7
      Left            =   2400
      ScaleHeight     =   7515
      ScaleWidth      =   9675
      TabIndex        =   172
      Top             =   0
      Width           =   9735
      Begin VB.CheckBox chk 
         Caption         =   "����ҽѧӰ����Ϣϵͳרҵ��"
         Height          =   255
         Index           =   52
         Left            =   120
         TabIndex        =   173
         Top             =   240
         Width           =   2895
      End
      Begin TabDlg.SSTab sstRIS 
         Height          =   6495
         Left            =   120
         TabIndex        =   185
         Top             =   960
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   11456
         _Version        =   393216
         Tabs            =   2
         Tab             =   1
         TabHeight       =   520
         TabCaption(0)   =   "RIS�ֳ�������"
         TabPicture(0)   =   "frmParPublic.frx":690F
         Tab(0).ControlEnabled=   0   'False
         Tab(0).Control(0)=   "chkShowSel"
         Tab(0).Control(1)=   "vsfRisDepts"
         Tab(0).Control(2)=   "vsfRISEnables"
         Tab(0).ControlCount=   3
         TabCaption(1)   =   "HISҽԺ����"
         TabPicture(1)   =   "frmParPublic.frx":692B
         Tab(1).ControlEnabled=   -1  'True
         Tab(1).Control(0)=   "Label9"
         Tab(1).Control(0).Enabled=   0   'False
         Tab(1).Control(1)=   "Frame3"
         Tab(1).Control(1).Enabled=   0   'False
         Tab(1).Control(2)=   "Frame1"
         Tab(1).Control(2).Enabled=   0   'False
         Tab(1).ControlCount=   3
         Begin VB.CheckBox chkShowSel 
            Caption         =   "ֻ��ʾ�����õĳ���"
            Enabled         =   0   'False
            Height          =   255
            Left            =   -74880
            TabIndex        =   191
            Top             =   600
            Width           =   2775
         End
         Begin VB.Frame Frame1 
            Caption         =   "��Ժ����"
            Height          =   855
            Left            =   240
            TabIndex        =   188
            Top             =   2280
            Width           =   9135
            Begin VB.TextBox txtMainHosp 
               Height          =   375
               Left            =   1320
               MaxLength       =   20
               TabIndex        =   189
               ToolTipText     =   "ҽԺ���룬���20���ַ�"
               Top             =   300
               Width           =   2415
            End
            Begin VB.Label Label10 
               Caption         =   "ҽԺ����"
               Height          =   255
               Left            =   360
               TabIndex        =   190
               Top             =   360
               Width           =   1215
            End
         End
         Begin VB.Frame Frame3 
            Caption         =   "��Ժ����"
            Height          =   3015
            Left            =   240
            TabIndex        =   186
            Top             =   3240
            Width           =   9135
            Begin VSFlex8Ctl.VSFlexGrid vsfBranchHosp 
               Height          =   2640
               Left            =   120
               TabIndex        =   187
               Top             =   240
               Width           =   8895
               _cx             =   15690
               _cy             =   4657
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
               Rows            =   2
               Cols            =   5
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
         Begin VSFlex8Ctl.VSFlexGrid vsfRisDepts 
            Height          =   5400
            Left            =   -68400
            TabIndex        =   192
            Top             =   960
            Visible         =   0   'False
            Width           =   2775
            _cx             =   4895
            _cy             =   9525
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
            Rows            =   2
            Cols            =   5
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
         Begin VSFlex8Ctl.VSFlexGrid vsfRISEnables 
            Height          =   5400
            Left            =   -74880
            TabIndex        =   193
            Top             =   960
            Width           =   6375
            _cx             =   11245
            _cy             =   9525
            Appearance      =   1
            BorderStyle     =   1
            Enabled         =   0   'False
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
            Rows            =   2
            Cols            =   5
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
         Begin VB.Label Label9 
            Caption         =   $"frmParPublic.frx":6947
            ForeColor       =   &H000000C0&
            Height          =   1695
            Left            =   240
            TabIndex        =   194
            Top             =   480
            Width           =   9135
         End
      End
      Begin VB.Label Label8 
         Caption         =   "�����ѡ���κ�RIS��ԤԼ���ã���ʾȫԺ������ҽѧӰ����Ϣϵͳרҵ�桱�������ա��������+���ϡ����ơ�"
         ForeColor       =   &H000000C0&
         Height          =   255
         Left            =   120
         TabIndex        =   174
         Top             =   600
         Width           =   9495
      End
   End
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7575
      Index           =   6
      Left            =   2400
      ScaleHeight     =   7545
      ScaleWidth      =   9705
      TabIndex        =   162
      Top             =   0
      Visible         =   0   'False
      Width           =   9735
      Begin VB.CheckBox chk 
         Caption         =   "ת����ʱ��δִ�л򲿷�ִ�еķ���ת���²���"
         Height          =   195
         Index           =   24
         Left            =   240
         TabIndex        =   176
         Top             =   1560
         Width           =   4200
      End
      Begin VB.Frame fraBabyWristlet 
         Caption         =   "Ӥ�����"
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   240
         TabIndex        =   103
         Top             =   4005
         Width           =   4815
         Begin VB.OptionButton optBabyWristletPrint 
            Caption         =   "ѡ���Ƿ��ӡ"
            Height          =   180
            Index           =   2
            Left            =   2805
            TabIndex        =   106
            Top             =   285
            Width           =   1500
         End
         Begin VB.OptionButton optBabyWristletPrint 
            Caption         =   "�Զ���ӡ"
            Height          =   180
            Index           =   1
            Left            =   1425
            TabIndex        =   105
            Top             =   285
            Width           =   1020
         End
         Begin VB.OptionButton optBabyWristletPrint 
            Caption         =   "����ӡ"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   104
            Top             =   285
            Value           =   -1  'True
            Width           =   900
         End
      End
      Begin VB.Frame fraPatiWristlet 
         Caption         =   "�������"
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   240
         TabIndex        =   99
         Top             =   3165
         Width           =   4815
         Begin VB.OptionButton optPatiWristletPrint 
            Caption         =   "����ӡ"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   100
            Top             =   285
            Value           =   -1  'True
            Width           =   900
         End
         Begin VB.OptionButton optPatiWristletPrint 
            Caption         =   "�Զ���ӡ"
            Height          =   180
            Index           =   1
            Left            =   1425
            TabIndex        =   101
            Top             =   285
            Width           =   1020
         End
         Begin VB.OptionButton optPatiWristletPrint 
            Caption         =   "ѡ���Ƿ��ӡ"
            Height          =   180
            Index           =   2
            Left            =   2805
            TabIndex        =   102
            Top             =   285
            Width           =   1500
         End
      End
      Begin VB.CheckBox chk 
         Caption         =   "��Ժ��ס�����������Ժ����"
         Height          =   195
         Index           =   15
         Left            =   240
         TabIndex        =   92
         Top             =   390
         Width           =   4200
      End
      Begin VB.CheckBox chk 
         Caption         =   "ת����סʱ����ȼ�Ĭ��Ϊ��"
         Height          =   195
         Index           =   16
         Left            =   240
         TabIndex        =   93
         Top             =   675
         Width           =   4200
      End
      Begin VB.CheckBox chk 
         Caption         =   "��Ժʱ���´�������ҽ��������������Ժ"
         Height          =   195
         Index           =   18
         Left            =   240
         TabIndex        =   95
         Top             =   1230
         Width           =   4200
      End
      Begin VB.CheckBox chk 
         Caption         =   "��סʱ����ָ��ҽ��С��"
         Height          =   195
         Index           =   14
         Left            =   240
         TabIndex        =   91
         Top             =   120
         Width           =   4200
      End
      Begin VB.CheckBox chk 
         Caption         =   "��Ժʱ����ȡ��Ժ���ΪĬ�ϵĳ�Ժ���"
         Height          =   195
         Index           =   17
         Left            =   240
         TabIndex        =   94
         Top             =   960
         Width           =   4200
      End
      Begin VB.Frame fraInDeptTime 
         Caption         =   "ȱʡ��סʱ��"
         Height          =   615
         Left            =   240
         TabIndex        =   96
         Top             =   2310
         Width           =   4815
         Begin VB.OptionButton OptInDeptTime 
            Caption         =   "��Ժʱ��"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   97
            Top             =   285
            Width           =   1215
         End
         Begin VB.OptionButton OptInDeptTime 
            Caption         =   "ϵͳʱ��"
            Height          =   180
            Index           =   1
            Left            =   1470
            TabIndex        =   98
            Top             =   285
            Width           =   1215
         End
      End
   End
   Begin VB.PictureBox picFunc 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      FillColor       =   &H8000000A&
      ForeColor       =   &H80000008&
      Height          =   7650
      Left            =   0
      ScaleHeight     =   7650
      ScaleWidth      =   2415
      TabIndex        =   150
      Top             =   0
      Width           =   2415
      Begin VB.PictureBox picVbar 
         BackColor       =   &H80000003&
         BorderStyle     =   0  'None
         FillColor       =   &H8000000A&
         Height          =   5820
         Left            =   2280
         MousePointer    =   9  'Size W E
         ScaleHeight     =   5820
         ScaleWidth      =   45
         TabIndex        =   154
         Top             =   120
         Width           =   45
      End
      Begin VB.PictureBox picTPL 
         BorderStyle     =   0  'None
         Height          =   6135
         Left            =   0
         ScaleHeight     =   6135
         ScaleWidth      =   2250
         TabIndex        =   151
         Top             =   0
         Width           =   2250
         Begin XtremeSuiteControls.TaskPanel tplFunc 
            Height          =   5250
            Left            =   0
            TabIndex        =   152
            Top             =   720
            Width           =   2205
            _Version        =   589884
            _ExtentX        =   3889
            _ExtentY        =   9260
            _StockProps     =   64
            Behaviour       =   1
            ItemLayout      =   2
            HotTrackStyle   =   3
         End
         Begin XtremeCommandBars.ImageManager imgFunc 
            Left            =   1920
            Top             =   360
            _Version        =   589884
            _ExtentX        =   635
            _ExtentY        =   635
            _StockProps     =   0
            Icons           =   "frmParPublic.frx":6AA2
         End
         Begin XtremeSuiteControls.ShortcutCaption sccFunc 
            Height          =   300
            Left            =   0
            TabIndex        =   153
            Top             =   0
            Width           =   2200
            _Version        =   589884
            _ExtentX        =   3881
            _ExtentY        =   529
            _StockProps     =   6
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            SubItemCaption  =   -1  'True
            Alignment       =   1
         End
      End
      Begin XtremeSuiteControls.ShortcutBar scbFunc 
         Height          =   6765
         Left            =   0
         TabIndex        =   155
         Top             =   0
         Width           =   2400
         _Version        =   589884
         _ExtentX        =   4233
         _ExtentY        =   11933
         _StockProps     =   64
      End
      Begin XtremeCommandBars.ImageManager imgType 
         Left            =   0
         Top             =   0
         _Version        =   589884
         _ExtentX        =   635
         _ExtentY        =   635
         _StockProps     =   0
         Icons           =   "frmParPublic.frx":B758
      End
   End
   Begin VB.PictureBox PicBottom 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000003&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   590
      Left            =   0
      ScaleHeight     =   585
      ScaleWidth      =   12690
      TabIndex        =   141
      Top             =   7650
      Width           =   12690
      Begin VB.TextBox txtLocate 
         Height          =   300
         Index           =   1
         Left            =   4700
         TabIndex        =   158
         Top             =   120
         Width           =   1200
      End
      Begin VB.TextBox txtLocate 
         Height          =   300
         Index           =   0
         Left            =   2400
         TabIndex        =   146
         Top             =   120
         Width           =   1200
      End
      Begin VB.CommandButton cmdHelp 
         Caption         =   "����(&H)"
         CausesValidation=   0   'False
         Height          =   350
         Left            =   60
         TabIndex        =   144
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   11040
         TabIndex        =   143
         Top             =   120
         Width           =   1100
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   9885
         TabIndex        =   142
         Top             =   120
         Width           =   1100
      End
      Begin VB.Label lblPrompt 
         BackStyle       =   0  'Transparent
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   6000
         TabIndex        =   159
         Top             =   165
         Width           =   3855
      End
      Begin VB.Label lblLocate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���Ҳ���(&F)"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   3600
         TabIndex        =   157
         Top             =   165
         Width           =   1095
      End
      Begin VB.Label lblLocate 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������(&S)"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   1200
         TabIndex        =   145
         Top             =   165
         Width           =   1095
      End
   End
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7575
      Index           =   2
      Left            =   2400
      ScaleHeight     =   7545
      ScaleWidth      =   9705
      TabIndex        =   156
      Top             =   0
      Width           =   9735
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   3
         Left            =   1215
         Style           =   2  'Dropdown List
         TabIndex        =   112
         Top             =   840
         Width           =   2475
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   4
         Left            =   1215
         Style           =   2  'Dropdown List
         TabIndex        =   108
         Top             =   120
         Width           =   2475
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   2
         Left            =   1215
         Style           =   2  'Dropdown List
         TabIndex        =   110
         Top             =   480
         Width           =   2475
      End
      Begin MSComctlLib.ListView lvwNo 
         Height          =   6120
         Left            =   240
         TabIndex        =   114
         Top             =   1500
         Width           =   3465
         _ExtentX        =   6112
         _ExtentY        =   10795
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "iltC32"
         SmallIcons      =   "imgC16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "��������"
            Object.Width           =   2822
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "�������"
            Object.Width           =   2646
         EndProperty
      End
      Begin ZL9BillEdit.BillEdit BillҩƷ���ұ�� 
         Height          =   3960
         Left            =   4200
         TabIndex        =   116
         Top             =   360
         Width           =   4050
         _ExtentX        =   7144
         _ExtentY        =   6985
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
      Begin ZL9BillEdit.BillEdit Bill���Ŀ��ұ�� 
         Height          =   2400
         Left            =   4200
         TabIndex        =   118
         Top             =   4740
         Width           =   4065
         _ExtentX        =   7170
         _ExtentY        =   4233
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   2
         RowHeight0      =   315
         RowHeightMin    =   315
         ColWidth0       =   1005
         BackColor       =   -2147483643
         BackColorBkg    =   -2147483643
         BackColorSel    =   10249818
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         ForeColorSel    =   -2147483634
         GridColor       =   -2147483630
         ColAlignment0   =   9
         ListIndex       =   -1
         CellBackColor   =   -2147483643
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "סԺ�Ź���"
         Height          =   180
         Index           =   10
         Left            =   240
         TabIndex        =   109
         Top             =   555
         Width           =   900
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "���ۺŹ���"
         Height          =   180
         Index           =   0
         Left            =   240
         TabIndex        =   111
         Top             =   915
         Width           =   900
      End
      Begin VB.Label Label2 
         Caption         =   "���Ŀ��Ҷ�Ӧ�ĵ��ݱ��"
         Height          =   285
         Left            =   4200
         TabIndex        =   117
         Top             =   4440
         Width           =   3975
      End
      Begin VB.Label Label5 
         AutoSize        =   -1  'True
         Caption         =   "ע�⣺���ұ�ſ�ѡ��ΧA-Z��1-9��ͬ���п��ұ�Ų����ظ���"
         Height          =   285
         Left            =   4200
         TabIndex        =   119
         Top             =   7215
         Width           =   5040
      End
      Begin VB.Label Label4 
         Caption         =   "ҩƷ���Ҷ�Ӧ�ĵ��ݱ��"
         Height          =   285
         Left            =   4200
         TabIndex        =   115
         Top             =   120
         Width           =   3975
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "����Ź���"
         Height          =   180
         Index           =   22
         Left            =   240
         TabIndex        =   107
         Top             =   180
         Width           =   900
      End
      Begin VB.Label Label3 
         Caption         =   "���ݺŵı���������˫���ɸı����ã�"
         Height          =   285
         Left            =   240
         TabIndex        =   113
         Top             =   1260
         Width           =   3675
      End
   End
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7575
      Index           =   1
      Left            =   2400
      ScaleHeight     =   7545
      ScaleWidth      =   9705
      TabIndex        =   147
      Top             =   0
      Visible         =   0   'False
      Width           =   9735
      Begin VB.Frame fra������� 
         Caption         =   "�������"
         Height          =   1065
         Left            =   5160
         TabIndex        =   177
         Top             =   6180
         Width           =   4305
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   5
            ItemData        =   "frmParPublic.frx":14DCC
            Left            =   1980
            List            =   "frmParPublic.frx":14DCE
            Style           =   2  'Dropdown List
            TabIndex        =   181
            Top             =   660
            Width           =   2100
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   0
            ItemData        =   "frmParPublic.frx":14DD0
            Left            =   1980
            List            =   "frmParPublic.frx":14DD2
            Style           =   2  'Dropdown List
            TabIndex        =   180
            Top             =   300
            Width           =   2100
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "δִ��������Ŀ���"
            Height          =   180
            Index           =   53
            Left            =   270
            TabIndex        =   179
            Top             =   720
            Width           =   1620
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "δ��ҩƷ���"
            Height          =   180
            Index           =   18
            Left            =   810
            TabIndex        =   178
            Top             =   360
            Width           =   1080
         End
      End
      Begin VB.CheckBox chk 
         Caption         =   "����ÿ��סԺʹ���µ�סԺ��"
         Height          =   195
         Index           =   1
         Left            =   5160
         TabIndex        =   46
         Top             =   2760
         Width           =   2640
      End
      Begin VB.CommandButton cmd�������� 
         Caption         =   "����(&S)"
         Enabled         =   0   'False
         Height          =   350
         Left            =   8400
         TabIndex        =   34
         Top             =   480
         Width           =   1100
      End
      Begin VB.Frame fra��Ժ��� 
         Caption         =   "����ת�ƻ��Ժ(δ��ҩƷ)"
         Height          =   1185
         Left            =   5160
         TabIndex        =   53
         Top             =   4815
         Width           =   4305
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   22
            ItemData        =   "frmParPublic.frx":14DD4
            Left            =   1080
            List            =   "frmParPublic.frx":14DD6
            Style           =   2  'Dropdown List
            TabIndex        =   57
            Top             =   660
            Width           =   3015
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   23
            ItemData        =   "frmParPublic.frx":14DD8
            Left            =   1080
            List            =   "frmParPublic.frx":14DDA
            Style           =   2  'Dropdown List
            TabIndex        =   55
            Top             =   300
            Width           =   3015
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "��Ժʱ"
            Height          =   180
            Index           =   46
            Left            =   375
            TabIndex        =   56
            Top             =   720
            Width           =   540
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "ת��ʱ"
            Height          =   180
            Index           =   48
            Left            =   390
            TabIndex        =   54
            Top             =   360
            Width           =   540
         End
      End
      Begin VB.CheckBox chk 
         Caption         =   "���ʱ����ȷ������ȼ�"
         Height          =   180
         Index           =   2
         Left            =   5160
         TabIndex        =   47
         Top             =   3120
         Width           =   2280
      End
      Begin VB.Frame FraChangeDept 
         Caption         =   "����ת�ƻ��Ժ"
         Height          =   1080
         Left            =   5160
         TabIndex        =   48
         Top             =   3480
         Width           =   4305
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   28
            ItemData        =   "frmParPublic.frx":14DDC
            Left            =   1905
            List            =   "frmParPublic.frx":14DDE
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   255
            Width           =   2205
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   29
            ItemData        =   "frmParPublic.frx":14DE0
            Left            =   1905
            List            =   "frmParPublic.frx":14DE2
            Style           =   2  'Dropdown List
            TabIndex        =   52
            Top             =   630
            Width           =   2205
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "(ת��)δ�����ʵ���"
            Height          =   180
            Index           =   57
            Left            =   210
            TabIndex        =   49
            Top             =   315
            Width           =   1620
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "(��Ժ)���ڻ�������"
            Height          =   180
            Index           =   8
            Left            =   210
            TabIndex        =   51
            Top             =   690
            Width           =   1620
         End
      End
      Begin VB.Frame fra��Ժ��鸱 
         Caption         =   "����ת�ƻ��Ժ(δִ��������Ŀ)"
         Height          =   4200
         Left            =   240
         TabIndex        =   35
         Top             =   1800
         Width           =   4425
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   6
            ItemData        =   "frmParPublic.frx":14DE4
            Left            =   1080
            List            =   "frmParPublic.frx":14DE6
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   675
            Width           =   3015
         End
         Begin VB.ComboBox cbo 
            Height          =   300
            Index           =   19
            ItemData        =   "frmParPublic.frx":14DE8
            Left            =   1080
            List            =   "frmParPublic.frx":14DEA
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   315
            Width           =   3015
         End
         Begin VSFlex8Ctl.VSFlexGrid vsUnCheckItem 
            Height          =   2565
            Left            =   240
            TabIndex        =   41
            Top             =   1440
            Width           =   3900
            _cx             =   6879
            _cy             =   4524
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
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483636
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   3
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   8
            Cols            =   2
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParPublic.frx":14DEC
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
            Editable        =   2
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
         Begin VB.Label Label17 
            Caption         =   "���������δִ��������Ŀ��"
            Height          =   255
            Left            =   240
            TabIndex        =   40
            Top             =   1200
            Width           =   2415
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "��Ժʱ"
            Height          =   180
            Index           =   50
            Left            =   255
            TabIndex        =   38
            Top             =   705
            Width           =   540
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "ת��ʱ"
            Height          =   180
            Index           =   17
            Left            =   255
            TabIndex        =   36
            Top             =   375
            Width           =   540
         End
      End
      Begin VB.Frame Fra 
         Caption         =   " ��Ժʱ���� "
         Height          =   765
         Index           =   5
         Left            =   5160
         TabIndex        =   42
         Top             =   1800
         Width           =   4335
         Begin VB.CheckBox chk 
            Caption         =   "������￨"
            Height          =   195
            Index           =   5
            Left            =   2880
            TabIndex        =   45
            Top             =   285
            Width           =   1200
         End
         Begin VB.CheckBox chk 
            Caption         =   "��ȡԤ����"
            Height          =   195
            Index           =   4
            Left            =   240
            TabIndex        =   43
            Top             =   285
            Width           =   1200
         End
         Begin VB.CheckBox chk 
            Caption         =   "���䴲λ��"
            Height          =   195
            Index           =   6
            Left            =   1560
            TabIndex        =   44
            Top             =   285
            Width           =   1200
         End
      End
      Begin MSComctlLib.ListView lvw���� 
         Height          =   1065
         Left            =   240
         TabIndex        =   33
         Top             =   480
         Width           =   8055
         _ExtentX        =   14208
         _ExtentY        =   1879
         View            =   3
         Arrange         =   2
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "���"
            Object.Width           =   1147
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "����"
            Object.Width           =   5645
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "˵��"
            Object.Width           =   3351
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "����"
            Object.Width           =   2857
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   4
            Text            =   "����"
            Object.Width           =   952
         EndProperty
      End
      Begin VB.Label lbl���������ӿ� 
         AutoSize        =   -1  'True
         Caption         =   "���������ӿ�(�ڹҺź�ҽ��վʹ��)"
         Height          =   180
         Left            =   240
         TabIndex        =   32
         Top             =   240
         Width           =   2880
      End
   End
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7575
      Index           =   3
      Left            =   2400
      ScaleHeight     =   7545
      ScaleWidth      =   10245
      TabIndex        =   149
      Top             =   0
      Visible         =   0   'False
      Width           =   10275
      Begin VB.CheckBox chk 
         Caption         =   "Ѫ��"
         Enabled         =   0   'False
         Height          =   195
         Index           =   26
         Left            =   4320
         TabIndex        =   183
         Top             =   1080
         Width           =   660
      End
      Begin VB.CommandButton cmd 
         Caption         =   "����"
         Height          =   350
         Index           =   0
         Left            =   7080
         TabIndex        =   171
         Top             =   102
         Width           =   1100
      End
      Begin VB.CheckBox chk 
         Caption         =   "�¿�ҽ��ǩ��ʱһ��ҽ��ǩ��һ��"
         Height          =   195
         Index           =   49
         Left            =   960
         TabIndex        =   122
         Top             =   540
         Width           =   3540
      End
      Begin VB.ComboBox cbo 
         Height          =   300
         Index           =   11
         ItemData        =   "frmParPublic.frx":14E2A
         Left            =   960
         List            =   "frmParPublic.frx":14E2C
         Style           =   2  'Dropdown List
         TabIndex        =   121
         Top             =   127
         Width           =   5940
      End
      Begin VB.CheckBox chk 
         Caption         =   "�����¼,������"
         Enabled         =   0   'False
         Height          =   195
         Index           =   47
         Left            =   6000
         TabIndex        =   127
         Top             =   825
         Width           =   1860
      End
      Begin VB.CheckBox chk 
         Caption         =   "ҽ��ҽ��,����"
         Enabled         =   0   'False
         Height          =   195
         Index           =   46
         Left            =   4320
         TabIndex        =   126
         Top             =   833
         Width           =   1500
      End
      Begin VB.CheckBox chk 
         Caption         =   "סԺҽ��,����"
         Enabled         =   0   'False
         Height          =   195
         Index           =   45
         Left            =   2640
         TabIndex        =   125
         Top             =   833
         Width           =   1500
      End
      Begin VB.CheckBox chk 
         Caption         =   "����ҽ��,����"
         Enabled         =   0   'False
         Height          =   195
         Index           =   44
         Left            =   960
         TabIndex        =   124
         Top             =   833
         Width           =   1620
      End
      Begin VB.CheckBox chk 
         Caption         =   "ҩƷ��ҩ"
         Enabled         =   0   'False
         Height          =   195
         Index           =   48
         Left            =   8040
         TabIndex        =   128
         Top             =   833
         Width           =   1020
      End
      Begin VB.CheckBox chk 
         Caption         =   "LIS"
         Enabled         =   0   'False
         Height          =   195
         Index           =   43
         Left            =   960
         TabIndex        =   129
         Top             =   1080
         Width           =   660
      End
      Begin VB.CheckBox chk 
         Caption         =   "PACS"
         Enabled         =   0   'False
         Height          =   195
         Index           =   42
         Left            =   2640
         TabIndex        =   130
         Top             =   1080
         Width           =   660
      End
      Begin TabDlg.SSTab sstSign 
         Height          =   5895
         Left            =   120
         TabIndex        =   132
         Top             =   1560
         Width           =   10140
         _ExtentX        =   17886
         _ExtentY        =   10398
         _Version        =   393216
         Style           =   1
         Tabs            =   9
         TabsPerRow      =   9
         TabHeight       =   520
         TabCaption(0)   =   "����ҽ��,����"
         TabPicture(0)   =   "frmParPublic.frx":14E2E
         Tab(0).ControlEnabled=   -1  'True
         Tab(0).Control(0)=   "vsDept(0)"
         Tab(0).Control(0).Enabled=   0   'False
         Tab(0).ControlCount=   1
         TabCaption(1)   =   "סԺҽ��ҽ��,����"
         TabPicture(1)   =   "frmParPublic.frx":14E4A
         Tab(1).ControlEnabled=   0   'False
         Tab(1).Control(0)=   "vsDept(1)"
         Tab(1).ControlCount=   1
         TabCaption(2)   =   "סԺ��ʿҽ��"
         TabPicture(2)   =   "frmParPublic.frx":14E66
         Tab(2).ControlEnabled=   0   'False
         Tab(2).Control(0)=   "vsDept(2)"
         Tab(2).ControlCount=   1
         TabCaption(3)   =   "ҽ��ҽ��,����"
         TabPicture(3)   =   "frmParPublic.frx":14E82
         Tab(3).ControlEnabled=   0   'False
         Tab(3).Control(0)=   "vsDept(3)"
         Tab(3).ControlCount=   1
         TabCaption(4)   =   "�����¼,������"
         TabPicture(4)   =   "frmParPublic.frx":14E9E
         Tab(4).ControlEnabled=   0   'False
         Tab(4).Control(0)=   "vsDept(4)"
         Tab(4).ControlCount=   1
         TabCaption(5)   =   "ҩƷ��ҩ"
         TabPicture(5)   =   "frmParPublic.frx":14EBA
         Tab(5).ControlEnabled=   0   'False
         Tab(5).Control(0)=   "vsDept(5)"
         Tab(5).ControlCount=   1
         TabCaption(6)   =   "LIS"
         TabPicture(6)   =   "frmParPublic.frx":14ED6
         Tab(6).ControlEnabled=   0   'False
         Tab(6).Control(0)=   "vsDept(6)"
         Tab(6).ControlCount=   1
         TabCaption(7)   =   "PACS"
         TabPicture(7)   =   "frmParPublic.frx":14EF2
         Tab(7).ControlEnabled=   0   'False
         Tab(7).Control(0)=   "vsDept(7)"
         Tab(7).ControlCount=   1
         TabCaption(8)   =   "Ѫ��"
         TabPicture(8)   =   "frmParPublic.frx":14F0E
         Tab(8).ControlEnabled=   0   'False
         Tab(8).Control(0)=   "vsDept(8)"
         Tab(8).ControlCount=   1
         Begin VSFlex8Ctl.VSFlexGrid vsDept 
            Height          =   5175
            Index           =   5
            Left            =   -74880
            TabIndex        =   138
            Top             =   480
            Width           =   9870
            _cx             =   17410
            _cy             =   9128
            Appearance      =   1
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
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483633
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
            Rows            =   2
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParPublic.frx":14F2A
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
         Begin VSFlex8Ctl.VSFlexGrid vsDept 
            Height          =   5175
            Index           =   4
            Left            =   -74880
            TabIndex        =   137
            Top             =   480
            Width           =   9870
            _cx             =   17410
            _cy             =   9128
            Appearance      =   1
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
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483633
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
            Rows            =   2
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParPublic.frx":14FBD
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
         Begin VSFlex8Ctl.VSFlexGrid vsDept 
            Height          =   5175
            Index           =   2
            Left            =   -74880
            TabIndex        =   135
            Top             =   480
            Width           =   9870
            _cx             =   17410
            _cy             =   9128
            Appearance      =   1
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
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483633
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
            Rows            =   2
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParPublic.frx":15050
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
         Begin VSFlex8Ctl.VSFlexGrid vsDept 
            Height          =   5175
            Index           =   7
            Left            =   -74880
            TabIndex        =   140
            Top             =   480
            Width           =   9870
            _cx             =   17410
            _cy             =   9128
            Appearance      =   1
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
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483633
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
            Rows            =   2
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParPublic.frx":150E3
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
            Begin ComctlLib.ImageList imgCheck 
               Left            =   0
               Top             =   720
               _ExtentX        =   1005
               _ExtentY        =   1005
               BackColor       =   -2147483643
               ImageWidth      =   16
               ImageHeight     =   16
               MaskColor       =   12632256
               _Version        =   327682
               BeginProperty Images {0713E8C2-850A-101B-AFC0-4210102A8DA7} 
                  NumListImages   =   2
                  BeginProperty ListImage1 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "frmParPublic.frx":15176
                     Key             =   "Checked"
                  EndProperty
                  BeginProperty ListImage2 {0713E8C3-850A-101B-AFC0-4210102A8DA7} 
                     Picture         =   "frmParPublic.frx":15350
                     Key             =   "UnChecked"
                  EndProperty
               EndProperty
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsDept 
            Height          =   5175
            Index           =   3
            Left            =   -74880
            TabIndex        =   136
            Top             =   480
            Width           =   9870
            _cx             =   17410
            _cy             =   9128
            Appearance      =   1
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
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483633
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
            Rows            =   2
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParPublic.frx":1552A
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
         Begin VSFlex8Ctl.VSFlexGrid vsDept 
            Height          =   5175
            Index           =   1
            Left            =   -74880
            TabIndex        =   134
            Top             =   480
            Width           =   9870
            _cx             =   17410
            _cy             =   9128
            Appearance      =   1
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
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483633
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
            Rows            =   2
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParPublic.frx":155BD
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
         Begin VSFlex8Ctl.VSFlexGrid vsDept 
            Height          =   5145
            Index           =   0
            Left            =   120
            TabIndex        =   133
            Top             =   495
            Width           =   9870
            _cx             =   17410
            _cy             =   9075
            Appearance      =   1
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
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483633
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
            Rows            =   2
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParPublic.frx":15650
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
         Begin VSFlex8Ctl.VSFlexGrid vsDept 
            Height          =   5175
            Index           =   6
            Left            =   -74880
            TabIndex        =   139
            Top             =   480
            Width           =   9870
            _cx             =   17410
            _cy             =   9128
            Appearance      =   1
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
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483633
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
            Rows            =   2
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParPublic.frx":156E3
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
         Begin VSFlex8Ctl.VSFlexGrid vsDept 
            Height          =   5175
            Index           =   8
            Left            =   -74880
            TabIndex        =   184
            Top             =   480
            Width           =   9870
            _cx             =   17410
            _cy             =   9128
            Appearance      =   1
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
            BackColor       =   -2147483643
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483633
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483633
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
            Rows            =   2
            Cols            =   5
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   280
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmParPublic.frx":15776
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
      Begin VB.Label Label6 
         Caption         =   "��֤����"
         Height          =   255
         Left            =   200
         TabIndex        =   120
         Top             =   150
         Width           =   735
      End
      Begin VB.Label Label7 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "����"
         Height          =   180
         Left            =   575
         TabIndex        =   123
         Top             =   840
         Width           =   360
      End
      Begin VB.Label Label15 
         Caption         =   "����ĳ�����Ϻ�δ��ѡ�κο��ң���ʾ�������ҿ��ơ�"
         Height          =   255
         Left            =   200
         TabIndex        =   131
         Top             =   1320
         Width           =   4815
      End
   End
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7575
      Index           =   4
      Left            =   2400
      ScaleHeight     =   7545
      ScaleWidth      =   9705
      TabIndex        =   160
      Top             =   0
      Visible         =   0   'False
      Width           =   9735
      Begin VSFlex8Ctl.VSFlexGrid vsgInput 
         Height          =   4560
         Index           =   0
         Left            =   240
         TabIndex        =   163
         Top             =   2775
         Width           =   5940
         _cx             =   10477
         _cy             =   8043
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
         BackColor       =   -2147483634
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483634
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483643
         FloodColor      =   192
         SheetBorder     =   -2147483637
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmParPublic.frx":15809
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
      Begin VB.Frame fraƱ�ݸ�ʽ 
         Caption         =   "Ԥ��Ʊ�ݸ�ʽ"
         Height          =   1395
         Left            =   240
         TabIndex        =   61
         Top             =   975
         Width           =   5955
         Begin VSFlex8Ctl.VSFlexGrid vfgBillFormat 
            Height          =   1095
            Left            =   120
            TabIndex        =   62
            Top             =   225
            Width           =   5775
            _cx             =   10186
            _cy             =   1931
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
            FormatString    =   $"frmParPublic.frx":158AD
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
      Begin VB.CheckBox chk 
         Caption         =   "ɨ�����֤ǩԼ"
         Height          =   180
         Index           =   19
         Left            =   240
         TabIndex        =   58
         Top             =   120
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.CheckBox chk 
         Caption         =   "����ͬʱ���뷢��"
         Height          =   255
         Index           =   20
         Left            =   240
         TabIndex        =   59
         Top             =   360
         Width           =   1935
      End
      Begin VB.CheckBox chk 
         Caption         =   "���￨�����Լ��˷�ʽ��ȡ"
         Height          =   180
         Index           =   21
         Left            =   240
         TabIndex        =   60
         Top             =   660
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VB.Label lblInput 
         AutoSize        =   -1  'True
         Caption         =   "���������"
         Height          =   180
         Left            =   240
         TabIndex        =   164
         Top             =   2520
         Width           =   900
      End
   End
   Begin VB.PictureBox picPar 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   7575
      Index           =   5
      Left            =   2400
      ScaleHeight     =   7545
      ScaleWidth      =   9705
      TabIndex        =   161
      Top             =   0
      Visible         =   0   'False
      Width           =   9735
      Begin VB.CheckBox chk 
         Caption         =   "�������޿մ����ܵǼ�"
         Height          =   255
         Index           =   23
         Left            =   240
         TabIndex        =   175
         Top             =   1680
         Width           =   3375
      End
      Begin VB.Frame fraDeptFirst 
         Caption         =   "���ҡ��������ȼ�"
         ForeColor       =   &H00000000&
         Height          =   615
         Left            =   240
         TabIndex        =   71
         Top             =   2400
         Width           =   4095
         Begin VB.OptionButton optDeptFirst 
            Caption         =   "��ѡ����"
            Height          =   255
            Index           =   1
            Left            =   1485
            MaskColor       =   &H00000000&
            TabIndex        =   73
            Top             =   285
            Width           =   1215
         End
         Begin VB.OptionButton optDeptFirst 
            Caption         =   "��ѡ����"
            Height          =   255
            Index           =   0
            Left            =   135
            MaskColor       =   &H00000000&
            TabIndex        =   72
            Top             =   285
            Value           =   -1  'True
            Width           =   1215
         End
      End
      Begin VB.CheckBox chk 
         Caption         =   "����ͨ������������ģ�����Ҳ�����Ϣ"
         Height          =   195
         Index           =   12
         Left            =   240
         MaskColor       =   &H00000000&
         TabIndex        =   67
         Top             =   1185
         Width           =   3540
      End
      Begin VB.CheckBox chk 
         Caption         =   "ɨ�����֤ǩԼ"
         Height          =   180
         Index           =   7
         Left            =   240
         TabIndex        =   64
         Top             =   390
         Value           =   1  'Checked
         Width           =   2520
      End
      Begin VB.CheckBox chk 
         Caption         =   "��Ժʱ�Զ�����һ�η���"
         Height          =   180
         Index           =   8
         Left            =   240
         MaskColor       =   &H00000000&
         TabIndex        =   65
         Top             =   660
         Value           =   1  'Checked
         Width           =   2295
      End
      Begin VB.TextBox txt 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H8000000F&
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
         ForeColor       =   &H80000012&
         Height          =   180
         Index           =   0
         Left            =   1875
         MaxLength       =   3
         TabIndex        =   69
         Text            =   "3"
         Top             =   1455
         Width           =   285
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00000000&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Left            =   1875
         TabIndex        =   70
         Top             =   1650
         Width           =   285
      End
      Begin VB.CheckBox chk 
         Caption         =   "�������תסԺ���ú������˷ѻ�����"
         Height          =   180
         Index           =   13
         Left            =   240
         TabIndex        =   86
         Top             =   5640
         Width           =   3615
      End
      Begin VB.Frame FraDepositMtoZ 
         Caption         =   "����תסԺԤ����Ʊ��"
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   240
         TabIndex        =   87
         Top             =   6000
         Width           =   4095
         Begin VB.OptionButton optDepositMtoZ 
            Caption         =   "ѡ���Ƿ��ӡ"
            Height          =   180
            Index           =   2
            Left            =   2640
            TabIndex        =   90
            Top             =   285
            Width           =   1380
         End
         Begin VB.OptionButton optDepositMtoZ 
            Caption         =   "�Զ���ӡ"
            Height          =   180
            Index           =   1
            Left            =   1305
            TabIndex        =   89
            Top             =   285
            Width           =   1020
         End
         Begin VB.OptionButton optDepositMtoZ 
            Caption         =   "����ӡ"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   88
            Top             =   285
            Value           =   -1  'True
            Width           =   900
         End
      End
      Begin VB.Frame fraWristlet 
         Caption         =   "�������"
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   240
         TabIndex        =   82
         Top             =   4560
         Width           =   4095
         Begin VB.OptionButton optWristletPrint 
            Caption         =   "����ӡ"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   83
            Top             =   285
            Value           =   -1  'True
            Width           =   900
         End
         Begin VB.OptionButton optWristletPrint 
            Caption         =   "�Զ���ӡ"
            Height          =   180
            Index           =   1
            Left            =   1305
            TabIndex        =   84
            Top             =   285
            Width           =   1020
         End
         Begin VB.OptionButton optWristletPrint 
            Caption         =   "ѡ���Ƿ��ӡ"
            Height          =   180
            Index           =   2
            Left            =   2640
            TabIndex        =   85
            Top             =   285
            Width           =   1380
         End
      End
      Begin VB.Frame fraPatientPage 
         Caption         =   "������ҳ"
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   240
         TabIndex        =   78
         Top             =   3840
         Width           =   4095
         Begin VB.OptionButton optFpagePrint 
            Caption         =   "ѡ���Ƿ��ӡ"
            Height          =   180
            Index           =   2
            Left            =   2655
            TabIndex        =   81
            Top             =   285
            Width           =   1380
         End
         Begin VB.OptionButton optFpagePrint 
            Caption         =   "�Զ���ӡ"
            Height          =   180
            Index           =   1
            Left            =   1305
            TabIndex        =   80
            Top             =   285
            Width           =   1020
         End
         Begin VB.OptionButton optFpagePrint 
            Caption         =   "����ӡ"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   79
            Top             =   285
            Value           =   -1  'True
            Width           =   900
         End
      End
      Begin VB.Frame fraDeposit 
         Caption         =   "Ԥ����Ʊ��"
         ForeColor       =   &H00000000&
         Height          =   645
         Left            =   240
         TabIndex        =   74
         Top             =   3120
         Width           =   4095
         Begin VB.OptionButton optPrepayPrint 
            Caption         =   "����ӡ"
            Height          =   180
            Index           =   0
            Left            =   135
            TabIndex        =   75
            Top             =   300
            Value           =   -1  'True
            Width           =   900
         End
         Begin VB.OptionButton optPrepayPrint 
            Caption         =   "�Զ���ӡ"
            Height          =   180
            Index           =   1
            Left            =   1305
            TabIndex        =   76
            Top             =   285
            Width           =   1020
         End
         Begin VB.OptionButton optPrepayPrint 
            Caption         =   "ѡ���Ƿ��ӡ"
            Height          =   180
            Index           =   2
            Left            =   2640
            TabIndex        =   77
            Top             =   285
            Width           =   1380
         End
      End
      Begin VB.CheckBox chk 
         Caption         =   "���벡�˵�����Ϣ"
         Height          =   210
         Index           =   3
         Left            =   240
         MaskColor       =   &H00000000&
         TabIndex        =   63
         Top             =   120
         Width           =   1740
      End
      Begin VB.CheckBox chk 
         Caption         =   "ҽ�ƿ������Լ��˷�ʽ��ȡ"
         Height          =   180
         Index           =   9
         Left            =   240
         MaskColor       =   &H00000000&
         TabIndex        =   66
         Top             =   933
         Value           =   1  'Checked
         Width           =   2535
      End
      Begin VSFlex8Ctl.VSFlexGrid vsgInput 
         Height          =   6915
         Index           =   1
         Left            =   4560
         TabIndex        =   165
         Top             =   360
         Width           =   4860
         _cx             =   8572
         _cy             =   12197
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
         BackColor       =   -2147483634
         ForeColor       =   -2147483640
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483634
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483643
         FloodColor      =   192
         SheetBorder     =   -2147483637
         FocusRect       =   3
         HighLight       =   2
         AllowSelection  =   0   'False
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   5
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmParPublic.frx":1593B
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
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "���������"
         Height          =   180
         Index           =   1
         Left            =   4560
         TabIndex        =   166
         Top             =   120
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "ԤԼ����ʱ��ȡ����    ���ڵ������Ϣ"
         Height          =   180
         Left            =   240
         TabIndex        =   68
         Top             =   1455
         Width           =   3855
      End
   End
End
Attribute VB_Name = "frmParPublic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mrsPar As ADODB.Recordset   '������ؼ���Ӧ��¼����ͬһ���������ܶ�Ӧһ�����ؼ���
Private marrFunc(2) As String
Private mlngPreFind As Long
Private mblnOk As Boolean
Private mobjESign As Object                 '����ǩ���ӿ�

Private Enum constTxtLocate
    txt_Par = 0
    txt_Dept = 1
End Enum

Private Enum constChk '������55
    chkֻ����¼���� = 0
    chk_ÿ��סԺʹ����סԺ�� = 1
    
    chk_���ȷ������ȼ� = 2
    chk_��ȡԤ���� = 4
    chk_ʱ������￨ = 5
    chk_���䴲λ�� = 6
    
    chk_ȫ����ֻ����� = 10
    chk_ȫ��ĸֻ����� = 11
    
    chk_���˵�ַ�ṹ��¼�� = 50
    chk_�����ַ�ṹ��¼�� = 51
    chk_��Ժҽ�������Ƚ��� = 22
    chk_����ҽѧӰ����Ϣϵͳרҵ��ӿ� = 52
    chk_ҽ�ƻ�������������¼�� = 25
    
    '������Ժ����
    chk_���벡�˵�����Ϣ = 3
    chk_��Ժʱɨ�����֤ǩԼ = 7
    chk_��Ժʱ�Զ�����һ�η��� = 8
    chk_��Ժʱ���Ѽ��� = 9
    chk_��Ժʱ����ģ������ = 12
    chk_����ת�������˷� = 13
    chk_�������޿մ����ܵǼ� = 23
    
    '�����������
    chk_��סָ��ҽ��С�� = 14
    chk_��ס���������Ժ���� = 15
    chk_ת����סʱ����ȼ�Ĭ��Ϊ�� = 16
    chk_��ԺĬ����� = 17
    chk_�´�����ҽ��������������Ժ = 18
    chk_ת����ת���� = 24
    
    '������Ϣ����
    chk_ɨ�����֤ǩԼ = 19
    chk_����ͬʱ���뷢�� = 20
    chk_���Ѽ��� = 21
    
    chk_Sign_pacs = 42
    chk_Sign_lis = 43
    chk_Sign_���� = 44
    chk_Sign_סԺ = 45
    chk_Sign_ҽ�� = 46
    chk_Sign_���� = 47
    chk_Sign_ҩƷ = 48
    chk_�¿�һ��ҽ��ǩ��һ�� = 49
    chk_sign_Ѫ�� = 26
End Enum

Private Enum constCbo
    cbo_����Ź��� = 4
    cbo_���ۺŹ��� = 3 'סԺ����
    cbo_סԺ�Ź��� = 2
    
    cbo_���������Դ = 1
    cbo_����������� = 8
    cbo_סԺ������� = 16
    
    cbo_���Ʊ���ģʽ = 10
    cbo_ҽ�������� = 15
    cbo_����ǩ����֤���� = 11
        
    cbo_ת��ʱδִ����Ŀ��� = 19
    cbo_��Ժʱδִ����Ŀ��� = 6
    cbo_��Ժʱδ��ҩ��Ŀ��� = 22
    cbo_ת��ʱδ��ҩ��Ŀ��� = 23
    cmd_ת��ʱδ������ʵ��� = 28
    cmd_��Ժʱ���ڻ������� = 29
    
    cbo_����_����δ��ҩƷ��� = 0
    cbo_����_����δִ����Ŀ��� = 5
    
End Enum

Private Enum constUpDown
    ud_��¼ʱ�� = 0
    ud_��ͯ����綨���� = 9
End Enum

Private Enum const����
    dtp_�����ϰ� = 0
    dtp_�����°� = 1
    dtp_�����ϰ� = 2
    dtp_�����°� = 3
End Enum

'�����õ���ǩ���Ĳ���
Private Enum constDeptCol
    col_ѡ�� = 0
    col_վ�� = 1
    col_���� = 2
    col_���� = 3
    col_���� = 4
End Enum
'����ǩ������
Private Enum constSign
    sst_���� = 0
    sst_סԺҽ�� = 1
    sst_סԺ��ʿ = 2
    sst_ҽ�� = 3
    sst_���� = 4
    sst_ҩƷ = 5
    sst_lis = 6
    sst_Pacs = 7
    sst_Ѫ�� = 8
End Enum

'����RIS����
Private Enum constRisEnables
    col_RIS���ü������ = 0
    col_RIS���ó��� = 1
    col_RIS���ÿ��� = 2
    col_RIS���ÿ���ID = 3
    col_RIS����ԤԼ����ȫ = 4
    col_RIS����ԤԼ���� = 5
    col_RIS����ԤԼ����ID = 6
End Enum
'����RIS�Ŀ���
Private Enum constRisDepts
    col_Ris����ѡ�� = 0
    col_Ris�������� = 1
    col_Ris���ұ��� = 2
    col_Ris����ID = 3
End Enum
'RISѡ��
Private Const RIS_Checked = "Checked"

'RIS��Ժ����
Private Enum constRisBranchHosp
    col_RIS��Ժ��� = 0
    col_RIS��Ժ���� = 1
    col_ris��Ժ���� = 2
    col_ris��Ժ�û��� = 3
    col_ris��Ժ���� = 4
    col_ris��Ժ���ݿ������ = 5
End Enum

Private Enum constTxt
    txt_��ϲ������� = 0
End Enum
'ҩ�����ϲ��ŵĿ��ұ��
Private Enum mGrdCol
    ѡ�� = 0
    ����
    ����
End Enum

Private Enum constVSGInput
    VSGInput_������Ϣ���������� = 0
    VSGInput_������Ժ���������� = 1
    
    COL_ϵͳ��ʶ = 1
    COL_�������� = 2
    COL_�����ַ = 3
End Enum

Private Enum constCmd
    cmd_����ǩ������ = 0
End Enum
'��¼���༭�Ŀ��ұ�������С��кͱ��ֵ
Private mintLastRow_Drug As Integer          '��
Private mintLastCol_Drug As Integer          '��
Private mstrLastCode_Drug As String          '���

Private mintLastRow_Stuff As Integer          '��
Private mintLastCol_Stuff As Integer          '��
Private mstrLastCode_Stuff As String          '���
Private mrsSvr As ADODB.Recordset   '������������Ŀ¼


Private Sub chkShowSel_Click()
    Dim i As Integer
    
    With vsfRISEnables
        For i = 1 To .Rows - 1
            If .Cell(flexcpChecked, i, col_RIS���ü������) = 2 Then .RowHidden(i) = IIF(chkShowSel.value = 1, True, False)
        Next i
    End With
End Sub

Private Sub cmd_Click(Index As Integer)
    If Index = cmd_����ǩ������ Then
       Call mobjESign.Setup(Me, gcnOracle, glngSys)
    End If
End Sub

Private Sub cmdHelp_Click()
     ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub Form_Activate()
    If Me.Tag = "��ʼ�ɹ�" Then
        Call scbFunc_SelectedChanged(scbFunc.Selected)
        Me.Tag = ""
    End If
End Sub

Private Sub Form_Load()
    Dim strCategory As String
    
    mblnOk = False
    strCategory = "��������,������Ŀ"
    
    'ͼ����,TaskPanelItem��ID(ͬʱҲ�ǲ�������Picture�ؼ������),TaskPanelItem�ı���;......
    marrFunc(0) = "100,0,ϵͳ����;102,1,���˹�����;104,4,������Ϣ����;105,5,������Ժ����;106,6,�����������"
    marrFunc(1) = "103,2,���ݱ������;101,3,����ǩ������;107,7,Ӱ����Ϣϵͳ"
    
    
    '1.��ʼ���������һ�������б�,ȱʡѡ�е�һ��
    Call InitSCBItem(scbFunc, strCategory, picTPL.hwnd)
    Call scbFunc.Icons.AddIcons(imgType.Icons)
      
    '2.��ʼ���������Ķ��������б�,ȱʡѡ�е�һ��
    Call InitTPLItem(sccFunc, tplFunc, scbFunc.Selected.Caption, marrFunc(0))
    Call tplFunc.Icons.AddIcons(imgFunc.Icons)
    
    
    Call InitData
    Call ShowErrParasMsg(Me, mrsPar)
    
    Me.Tag = "��ʼ�ɹ�"
End Sub

Private Sub optBabyWristletPrint_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(optBabyWristletPrint, Index, mrsPar)
End Sub

Private Sub optBabyWristletPrint_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optBabyWristletPrint_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optBabyWristletPrint, Index, mrsPar)
End Sub

Private Sub optDepositMtoZ_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(optDepositMtoZ, Index, mrsPar)
End Sub

Private Sub optDepositMtoZ_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optDepositMtoZ_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optDepositMtoZ, Index, mrsPar)
End Sub

Private Sub optDeptFirst_Click(Index As Integer)
     If Me.Visible Then Call SetParChange(optDeptFirst, Index, mrsPar)
End Sub

Private Sub optDeptFirst_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optDeptFirst_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optDeptFirst, Index, mrsPar)
End Sub

Private Sub optFpagePrint_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(optFpagePrint, Index, mrsPar)
End Sub

Private Sub optFpagePrint_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optFpagePrint_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optFpagePrint, Index, mrsPar)
End Sub

Private Sub OptInDeptTime_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(OptInDeptTime, Index, mrsPar)
End Sub

Private Sub OptInDeptTime_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub OptInDeptTime_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(OptInDeptTime, Index, mrsPar)
End Sub

Private Sub optPatiWristletPrint_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(optPatiWristletPrint, Index, mrsPar)
End Sub

Private Sub optPatiWristletPrint_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optPatiWristletPrint_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optPatiWristletPrint, Index, mrsPar)
End Sub

Private Sub optPrepayPrint_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(optPrepayPrint, Index, mrsPar)
End Sub

Private Sub optPrepayPrint_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optPrepayPrint_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optPrepayPrint, Index, mrsPar)
End Sub

Private Sub optWristletPrint_Click(Index As Integer)
    If Me.Visible Then Call SetParChange(optWristletPrint, Index, mrsPar)
End Sub

Private Sub optWristletPrint_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optWristletPrint_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(optWristletPrint, Index, mrsPar)
End Sub

Private Sub tplFunc_ItemClick(ByVal Item As XtremeSuiteControls.ITaskPanelGroupItem)
    Dim i As Long
    
    For i = 0 To picPar.UBound
        picPar(i).Visible = (i = Item.ID)
    Next
    
    lblLocate(txt_Dept).Visible = (Item.ID = GetFuncID("����ǩ������", marrFunc) Or _
                                   Item.ID = GetFuncID("���ݱ������", marrFunc))
    txtLocate(txt_Dept).Visible = lblLocate(txt_Dept).Visible
    If txtLocate(txt_Dept).Visible Then
        lblPrompt.Left = txtLocate(txt_Dept).Left + txtLocate(txt_Dept).Width + 60
    Else
        lblPrompt.Left = txtLocate(txt_Par).Left + txtLocate(txt_Par).Width + 60
    End If
    lblPrompt.Width = cmdOk.Left - lblPrompt.Left - 120
    mlngPreFind = 1
    
    tplFunc.Tag = Item.ID   '���ڻ�ȡ��ǰѡ�е�TaskPanelItem
End Sub


Private Sub Form_Resize()
    Dim i As Long
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    If picVbar.Left < 1500 Then picVbar.Left = 1500
    If picVbar.Left > Me.ScaleWidth - 3000 Then picVbar.Left = Me.ScaleWidth - 3000
    picVbar.Top = 0
    
    picFunc.Width = picVbar.Left + picVbar.Width
    
    For i = 0 To picPar.UBound
        picPar(i).Top = Me.ScaleTop
        picPar(i).Left = picFunc.Left + picFunc.ScaleWidth
        picPar(i).Width = Me.ScaleWidth - picPar(i).Left
        picPar(i).Height = Me.ScaleHeight - PicBottom.ScaleHeight
    Next
End Sub


Private Sub scbFunc_ExpandButtonDown(CancelMenu As Boolean)
    CancelMenu = True
End Sub

Private Sub picBottom_Resize()
    cmdCancel.Left = PicBottom.ScaleWidth - cmdCancel.Width - 120
    cmdOk.Left = cmdCancel.Left - cmdOk.Width - 120
End Sub


Private Sub picFunc_Resize()
    scbFunc.Top = picFunc.ScaleTop
    scbFunc.Left = picFunc.ScaleLeft + 45
    scbFunc.Width = picFunc.ScaleWidth - picVbar.Width - 45
    scbFunc.Height = picFunc.ScaleHeight
    
    picVbar.Height = picFunc.ScaleHeight
End Sub

Private Sub picTPL_Resize()
    sccFunc.Left = picTPL.ScaleLeft
    sccFunc.Width = picTPL.ScaleWidth
    
    tplFunc.Left = picTPL.ScaleLeft
    tplFunc.Top = sccFunc.Top + sccFunc.Height
    tplFunc.Height = picTPL.ScaleHeight - sccFunc.Height
    tplFunc.Width = picTPL.ScaleWidth
End Sub


Private Sub picVbar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 1 Then
        picVbar.Left = IIF(picVbar.Left + X < 2000, 2000, picVbar.Left + X)
        Call Form_Resize
    End If
End Sub

Private Sub scbFunc_SelectedChanged(ByVal Item As XtremeSuiteControls.IShortcutBarItem)
    If Me.Visible Then
        Call InitTPLItem(sccFunc, tplFunc, Item.Caption, marrFunc(Item.ID - 1)) 'ID�Ǵ�1��ʼ�ģ���ΪͬʱΪͼ����ţ�,�����Ǵ�0��ʼ
        Call tplFunc_ItemClick(tplFunc.Groups(1).Items(1))
    End If
End Sub


Public Sub LocateFuncItem(ByVal lngFunc As Long)
'���ܣ�����IDѡ��һ���Ͷ�������
    Dim i As Long, j As Long, lngId As Long
    Dim arrTmp As Variant
    Dim n As Long
    
    For i = 0 To UBound(marrFunc)
        arrTmp = Split(marrFunc(i), ";")
        For j = 0 To UBound(arrTmp)
            lngId = Split(arrTmp(j), ",")(1)
            If lngFunc = lngId Then
                tplFunc.Tag = lngId
                Set scbFunc.Selected = scbFunc(i)
                
                For n = 1 To tplFunc.Groups(1).Items.Count
                    tplFunc.Groups(1).Items(n).Selected = tplFunc.Groups(1).Items(n).ID = lngId
                Next
            End If
        Next
    Next
End Sub


Private Sub InitData()
'���ܣ���ʼ������ؼ�,��ȡ����������
    '1.��ʼ������
    
    mlngPreFind = 1
    Call InitSystemPara
    
    
    
    '2.��ʼ������ؼ�
    Call InitEnv
    Call Load�����ӿ�
    
    Call Load���ݱ������
    Call LoadҩƷ���Ŀ��ұ��
    
    Call LoadThirdSvr
        
    '3.����ϵͳ����
    Call LoadPar
    
    
End Sub


Private Sub LoadPar()
'���ܣ���ȡ�����ز���������ؼ�
    Dim strValue As String
    Dim i As Long, arrTmp As Variant
    Dim blnFind As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String      'ģ���1:������1:�ؼ��������1,������2:�ؼ��������2,......
    Dim arrObj As Variant  '�������ģ��1,������1,�ؼ�����1,ģ��2,������2,�ؼ�����2,......
    Dim strBillFormat As String, strPrintMode As String 'Ԥ��Ʊ�ݸ�ʽ�ʹ�ӡ��ʽ
    Set rsTmp = GetPar(mrsPar, p������Ϣ���� & "," & P������Ժ���� & "," & p�����������)
        
     '1.����CheckBox�����
    strTmp = "0:10:" & chk_��ȡԤ���� & _
            ",0:11:" & chk_ʱ������￨ & _
            ",0:13:" & chk_���䴲λ�� & _
            ",0:99:" & chk_���ȷ������ȼ� & _
            ",0:191:" & chkֻ����¼���� & _
            ",0:145:" & chk_ÿ��סԺʹ����סԺ�� & _
            ",0:239:" & chk_�¿�һ��ҽ��ǩ��һ�� & _
            ",0:251:" & chk_���˵�ַ�ṹ��¼�� & _
            ",0:252:" & chk_�����ַ�ṹ��¼�� & _
            ",0:253:" & chk_��Ժҽ�������Ƚ��� & _
            ",0:255:" & chk_����ҽѧӰ����Ϣϵͳרҵ��ӿ� & _
            ",0:287:" & chk_ҽ�ƻ�������������¼��

    Call SetParToControl(strTmp, mrsPar, chk)
    
    '������Ϣ���
    strTmp = "1101:ɨ�����֤ǩԼ:" & chk_ɨ�����֤ǩԼ & ",1101:����ͬʱ���뷢��:" & chk_����ͬʱ���뷢�� & ",1101:���Ѽ���:" & chk_���Ѽ���
    Call SetParToControl(strTmp, mrsPar, chk)

    '��Ժ�������
     strTmp = "1131:���Ѽ���:" & chk_��Ժʱ���Ѽ��� & ",1131:������Ϣ:" & chk_���벡�˵�����Ϣ & ",1131:����ģ������:" & chk_��Ժʱ����ģ������ & _
            ",1131:ɨ�����֤ǩԼ:" & chk_��Ժʱɨ�����֤ǩԼ & ",1131:���ü���ʱ��:" & chk_��Ժʱ�Զ�����һ�η��� & ",1131:����ת�������˷�:" & chk_����ת�������˷� & _
        ",1131:�������޿մ����ܵǼ�:" & chk_�������޿մ����ܵǼ�
     Call SetParToControl(strTmp, mrsPar, chk)
     '����������
     strTmp = "1132:Ĭ�����:" & chk_��ԺĬ����� & ",1132:�����������:" & chk_��ס���������Ժ���� & ",1132:����ȼ�Ĭ��Ϊ��:" & chk_ת����סʱ����ȼ�Ĭ��Ϊ�� & _
        ",1132:��Ժ����:" & chk_�´�����ҽ��������������Ժ & ",1132:��סָ��ҽ��С��:" & chk_��סָ��ҽ��С�� & ",1132:ת����ת����:" & chk_ת����ת����
    Call SetParToControl(strTmp, mrsPar, chk)
    
    '2.����ComboBox�����
    strTmp = "0:22:" & cbo_��Ժʱδִ����Ŀ��� & _
            ",0:32:" & cbo_ת��ʱδִ����Ŀ��� & _
            ",0:59:" & cbo_ҽ�������� & _
            ",0:61:" & cbo_���Ʊ���ģʽ & _
            ",0:154:" & cbo_��Ժʱδ��ҩ��Ŀ��� & _
            ",0:155:" & cbo_ת��ʱδ��ҩ��Ŀ��� & _
            ",0:227:" & cmd_ת��ʱδ������ʵ��� & _
            ",0:265:" & cbo_����_����δ��ҩƷ��� & _
            ",0:266:" & cbo_����_����δִ����Ŀ��� & _
            ",0:235:" & cmd_��Ժʱ���ڻ�������

    Call SetParToControl(strTmp, mrsPar, cbo)
    
    strTmp = "0:25:" & cbo_����ǩ����֤����
    Call SetParToControl(strTmp, mrsPar, cbo, 2)
    
    '3.����UpDown�����
    strTmp = "0:147:" & ud_��ͯ����綨���� & _
            ",0:158:" & ud_��¼ʱ��
                
    Call SetParToControl(strTmp, mrsPar, ud)    'mrsPar�洢�Ŀؼ�����txtUD
            
    '4.����TextBox�����
    strTmp = "1131:��ϲ�������:" & txt_��ϲ�������
    Call SetParToControl(strTmp, mrsPar, txt)
    
    '5.����ListBox�����
    strTmp = ""
    'Call SetParToControl(strTmp, mrsPar, lst)
    
    '6.����OptionButton�����
    '������Ժ����
    arrObj = Array(P������Ժ����, "��ѡ����", optDeptFirst, P������Ժ����, "Ԥ����Ʊ�ݴ�ӡ", optPrepayPrint, P������Ժ����, "������ҳ��ӡ", optFpagePrint, P������Ժ����, "���������ӡ", optWristletPrint, P������Ժ����, "����תסԺԤ����ӡ", optDepositMtoZ)
    Call SetParToControl("", mrsPar, arrObj)
    '�����������
    arrObj = Array(p�����������, "ȱʡ���ʱ��", OptInDeptTime, p�����������, "���������ӡ", optPatiWristletPrint, p�����������, "Ӥ�������ӡ", optBabyWristletPrint)
    Call SetParToControl("", mrsPar, arrObj)
    
    '7.����ϵͳ����
    rsTmp.Filter = "ģ��=0"
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!����ֵ
        Select Case rsTmp!������
        
        Case 1    '�������°�ʱ��
            i = InStr(UCase(strValue), "AND")
            strTmp = Mid(strValue, 1, i - 2)
            dtp(dtp_�����ϰ�).value = CDate(strTmp)
            strTmp = Mid(strValue, i + 4)
            dtp(dtp_�����°�).value = CDate(strTmp)
            
            Call SetParRelation(dtp, dtp_�����ϰ�, mrsPar, rsTmp!������)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(dtp, dtp_�����°�, mrsPar)
            
        Case 2    '�������°�ʱ��
            i = InStr(UCase(strValue), "AND")
            strTmp = Mid(strValue, 1, i - 2)
            dtp(dtp_�����ϰ�).value = CDate(strTmp)
            strTmp = Mid(strValue, i + 4)
            dtp(dtp_�����°�).value = CDate(strTmp)
        
            Call SetParRelation(dtp, dtp_�����ϰ�, mrsPar, rsTmp!������)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(dtp, dtp_�����°�, mrsPar)
    
        Case 26    '����ǩ��ʹ�ó���
            strTmp = chk_Sign_���� & "," & chk_Sign_סԺ & "," & chk_Sign_ҽ�� & "," & chk_Sign_���� & "," & _
                    chk_Sign_ҩƷ & "," & chk_Sign_lis & "," & chk_Sign_pacs & "," & chk_sign_Ѫ��
            arrTmp = Split(strTmp, ",")
            For i = 1 To 8
                chk(arrTmp(i - 1)).value = Val(Mid(strValue, i, 1))
                If i = 1 Then
                    Call SetParRelation(chk, arrTmp(i - 1), mrsPar, rsTmp!������)
                Else
                    Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
                    Call SetParRelation(chk, arrTmp(i - 1), mrsPar)
                End If
            Next
            
        Case 44    '�շ���Ŀ��������Ŀ������ƥ�䷽ʽ
            chk(chk_ȫ����ֻ�����).value = IIF(Mid(NVL(strValue, "00"), 1, 1) = "1", 1, 0)
            chk(chk_ȫ��ĸֻ�����).value = IIF(Mid(NVL(strValue, "00"), 2, 1) = "1", 1, 0)
            
            Call SetParRelation(chk, chk_ȫ����ֻ�����, mrsPar, rsTmp!������)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(chk, chk_ȫ��ĸֻ�����, mrsPar)
        Case 55
            cbo(cbo_���������Դ).ListIndex = IIF(Val(strValue) > cbo(cbo_���������Դ).ListCount, 0, Val(strValue) - 1)
            
            Call SetParRelation(cbo, cbo_���������Դ, mrsPar, rsTmp!������)
        Case 65
            cbo(cbo_�����������).ListIndex = Val(Mid(NVL(strValue, "11"), 1, 1)) - 1
            cbo(cbo_סԺ�������).ListIndex = Val(Mid(NVL(strValue, "11"), 2, 1)) - 1
            
            Call SetParRelation(cbo, cbo_�����������, mrsPar, rsTmp!������)
            Call zlDatabase.zlInsertCurrRowData(rsTmp, mrsPar, "")
            Call SetParRelation(cbo, cbo_סԺ�������, mrsPar)
        Case 234
            Call Initת�Ƴ�Ժ�������Ŀ(strValue)
            
            Call SetParRelation(vsUnCheckItem, 0, mrsPar, rsTmp!������)
        End Select
        rsTmp.MoveNext
    Loop
    
    
    '8.����ģ�����
    strBillFormat = "": strPrintMode = ""
    rsTmp.Filter = "ģ��=" & p������Ϣ����
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!����ֵ
        Select Case rsTmp!������
            Case "Ԥ����Ʊ��ʽ" 'Ԥ����Ʊ��ʽ
                strBillFormat = strValue
                Call SetParRelation(vfgBillFormat, 0, mrsPar, rsTmp!������, 1101, , vfgBillFormat.ColIndex("Ʊ�ݸ�ʽ"))
            Case "Ԥ����Ʊ��ӡ��ʽ" 'Ԥ����Ʊ��ӡ��ʽ
                strPrintMode = strValue
                Call SetParRelation(vfgBillFormat, 0, mrsPar, rsTmp!������, 1101, "", vfgBillFormat.ColIndex("Ԥ����ӡ��ʽ"))
            Case "���������" '������Ϣ���� ���������
                If strValue = "" Then strValue = "����|����|ѧ��|����״��|ְҵ|���|��������|����֤��|���֤��|�����ص�|��סַ|��ͥ��ַ�ʱ�|��ͥ�绰|��ϵ������|��ϵ�˹�ϵ|���ڵ�ַ|���ڵ�ַ�ʱ�|����|��ϵ�˵�ַ|��ϵ�˵绰|��ϵ�����֤��|������λ|��λ�绰|��λ�ʱ�|��λ������|��λ�ʺ�|����"
                Call LoadInputItem(VSGInput_������Ϣ����������, strValue)
                Call SetParRelation(vsgInput, VSGInput_������Ϣ����������, mrsPar, rsTmp!������, p������Ϣ����)
        End Select
        rsTmp.MoveNext
    Loop
    
    rsTmp.Filter = "ģ��=" & P������Ժ����
    Do Until rsTmp.EOF
        strValue = "" & rsTmp!����ֵ
        Select Case rsTmp!������
            Case "���������" '������Ϣ���� ���������
                If strValue = "" Then strValue = "����|����|ѧ��|����״��|ְҵ|���|��������|����֤��|���֤��|�����ص�|��סַ|��ͥ��ַ�ʱ�|��ͥ�绰|��ϵ������|��ϵ�˹�ϵ|���ڵ�ַ|���ڵ�ַ�ʱ�|����|��ϵ�˵�ַ|��ϵ�˵绰|��ϵ�����֤��|������λ|��λ�绰|��λ�ʱ�|��λ������|��λ�ʺ�|����"
                Call LoadInputItem(VSGInput_������Ժ����������, strValue)
                Call SetParRelation(vsgInput, VSGInput_������Ժ����������, mrsPar, rsTmp!������, P������Ժ����)
        End Select
        rsTmp.MoveNext
    Loop
    
    
    '����Ԥ��Ʊ�ݸ�ʽ�ʹ�ӡ��ʽ��Ϣ
    Call LaodBillForamt(vfgBillFormat, strBillFormat, strPrintMode)
    
End Sub

Private Sub cmdOK_Click()
    If ValidateData() = False Then Exit Sub
    
    If cbo(cbo_����ǩ����֤����).ListIndex > 0 Then Call Save����ǩ��
    
    Call Save�����ӿ�
    
    Call Save���ݱ������
    Call Save���ұ��
    
    Call SaveThirdSvr
    
    '���桰Ӱ����Ϣϵͳ�����ÿ���
    If chk(52).value = 1 Then Call SaveRisEnable
    
    Call SaveRisBranchHosp
    
    If SavePar(mrsPar, Me) = False Then Exit Sub
    mblnOk = True
    Unload Me
End Sub


Private Function ValidateData() As Boolean
'���ܣ���֤���ݵ���Ч��
    Dim i As Long, strTmp As String
    
    '�Զ��Կ��ұ�����һ���༭��������У��
    If mintLastRow_Drug > 0 And Len(Trim(mstrLastCode_Drug)) > 0 Then
        With BillҩƷ���ұ��
            If .TextMatrix(mintLastRow_Drug, mintLastCol_Drug) <> UCase(mstrLastCode_Drug) Then
                .TextMatrix(mintLastRow_Drug, mintLastCol_Drug) = UCase(mstrLastCode_Drug)
            End If
        End With
    End If
    If mintLastRow_Stuff > 0 And Len(Trim(mstrLastCode_Stuff)) > 0 Then
        With Bill���Ŀ��ұ��
            If .TextMatrix(mintLastRow_Stuff, mintLastCol_Stuff) <> UCase(mstrLastCode_Stuff) Then
                .TextMatrix(mintLastRow_Stuff, mintLastCol_Stuff) = UCase(mstrLastCode_Stuff)
            End If
        End With
    End If
    
      
    If CheckNumberRule_Drug = True Then
        'ͬһ��GRID��Ŀ��ұ�Ų����ظ�
        With BillҩƷ���ұ��
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, 2)) <> "" Then
                    If InStr(1, strTmp & ",", "," & .TextMatrix(i, 2) & ",") > 0 Then
                        MsgBox "ҩƷ���ҵ�" & i & "�б���ظ���", vbQuestion, gstrSysName
                  
                        .Row = i
                        .Col = 2
                        .SetFocus
                        Exit Function
                    End If
                End If
                strTmp = strTmp & "," & .TextMatrix(i, 2)
            Next
        End With
        strTmp = ""
    Else
        With BillҩƷ���ұ��
            For i = 1 To .Rows - 1
                .TextMatrix(i, 2) = ""
            Next
            .Tag = "���޸�"
        End With
    End If
    
    If CheckNumberRule_Stuff = True Then
        'ͬһ��GRID��Ŀ��ұ�Ų����ظ�
        With Bill���Ŀ��ұ��
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, 2)) <> "" Then
                    If InStr(1, strTmp & ",", "," & .TextMatrix(i, 2) & ",") > 0 Then
                        MsgBox "���Ŀ��ҵ�" & i & "�б���ظ���", vbQuestion, gstrSysName
                    
                        .Row = i
                        .Col = 2
                        .SetFocus
                        Exit Function
                    End If
                End If
                strTmp = strTmp & "," & .TextMatrix(i, 2)
            Next
        End With
    Else
        With Bill���Ŀ��ұ��
            For i = 1 To .Rows - 1
                .TextMatrix(i, 2) = ""
            Next
            .Tag = "���޸�"
        End With
    End If
    
    If ValidateRisBranchHosp = False Then Exit Function
    
    ValidateData = True
End Function


Private Sub InitEnv()
'���ܣ���ʼ������ؼ������ػ�������
    Dim strTmp As String
    Dim i As Integer
    Dim rsTmp As New ADODB.Recordset
    Dim blnTmp As Boolean
    Dim arrTemp As Variant
    
    cbo(cbo_סԺ�Ź���).AddItem "0-˳����"
    cbo(cbo_סԺ�Ź���).AddItem "1-����(YYMM)+˳���(0000)"
    cbo(cbo_סԺ�Ź���).AddItem "2-��(YYYY)+˳���(00000)"
    cbo(cbo_סԺ�Ź���).ListIndex = 0
    
    cbo(cbo_���ۺŹ���).AddItem "0-˳����"
    cbo(cbo_���ۺŹ���).AddItem "1-����(YYMM)+˳���(0000)"
    cbo(cbo_���ۺŹ���).AddItem "2-��(YYYY)+˳���(00000)"
    cbo(cbo_���ۺŹ���).ListIndex = 0
    
    cbo(cbo_����Ź���).AddItem "0-˳����"
    cbo(cbo_����Ź���).AddItem "1-������(YYMMDD)+˳���(0000)"
    cbo(cbo_����Ź���).ListIndex = 0

    cbo(cbo_���������Դ).AddItem "1-��ѡ��������Դ"
    cbo(cbo_���������Դ).AddItem "2-����ϱ�׼����"
    cbo(cbo_���������Դ).AddItem "3-��������������"
    cbo(cbo_���������Դ).ListIndex = 0
    
    cbo(cbo_�����������).AddItem "1-������������"
    cbo(cbo_�����������).AddItem "2-�����ݿ���ȡ����"
    cbo(cbo_�����������).AddItem "3-��ҽ�����˴����ݿ�����"
    cbo(cbo_�����������).ListIndex = 0
    
    cbo(cbo_סԺ�������).AddItem "1-������������"
    cbo(cbo_סԺ�������).AddItem "2-�����ݿ���ȡ����"
    cbo(cbo_סԺ�������).AddItem "3-��ҽ�����˴����ݿ�����"
    cbo(cbo_סԺ�������).ListIndex = 0
    
    
    cbo(cbo_���Ʊ���ģʽ).AddItem "˳����"
    cbo(cbo_���Ʊ���ģʽ).AddItem "����+�����+˳����"
        
    cbo(cbo_ҽ��������).AddItem "0-�����м��"
    cbo(cbo_ҽ��������).AddItem "1-��鲢����δ������Ŀ"
    cbo(cbo_ҽ��������).AddItem "2-��鲢��ֹδ������Ŀ"
    cbo(cbo_ҽ��������).ListIndex = 1
    
    
    cbo(cbo_��Ժʱδִ����Ŀ���).AddItem "0-�����"
    cbo(cbo_��Ժʱδִ����Ŀ���).AddItem "1-��鲢��ʾ"
    cbo(cbo_��Ժʱδִ����Ŀ���).AddItem "2-��鲢��ֹ"
    cbo(cbo_��Ժʱδִ����Ŀ���).ListIndex = 0
    
    cbo(cbo_ת��ʱδִ����Ŀ���).AddItem "0-�����"
    cbo(cbo_ת��ʱδִ����Ŀ���).AddItem "1-��鲢��ʾ"
    cbo(cbo_ת��ʱδִ����Ŀ���).AddItem "2-��鲢��ֹ"
    cbo(cbo_ת��ʱδִ����Ŀ���).ListIndex = 0
    
    cbo(cbo_��Ժʱδ��ҩ��Ŀ���).AddItem "0-�����"
    cbo(cbo_��Ժʱδ��ҩ��Ŀ���).AddItem "1-��鲢��ʾ"
    cbo(cbo_��Ժʱδ��ҩ��Ŀ���).AddItem "2-��鲢��ֹ"
    cbo(cbo_��Ժʱδ��ҩ��Ŀ���).ListIndex = 0
    
    cbo(cbo_ת��ʱδ��ҩ��Ŀ���).AddItem "0-�����"
    cbo(cbo_ת��ʱδ��ҩ��Ŀ���).AddItem "1-��鲢��ʾ"
    cbo(cbo_ת��ʱδ��ҩ��Ŀ���).AddItem "2-��鲢��ֹ"
    cbo(cbo_ת��ʱδ��ҩ��Ŀ���).ListIndex = 0
    
    cbo(cmd_ת��ʱδ������ʵ���).AddItem "0-�����"
    cbo(cmd_ת��ʱδ������ʵ���).AddItem "1-��鲢��ʾ"
    cbo(cmd_ת��ʱδ������ʵ���).AddItem "2-��鲢��ֹ"
    cbo(cmd_ת��ʱδ������ʵ���).ListIndex = 0
    
    cbo(cmd_��Ժʱ���ڻ�������).AddItem "0-�����"
    cbo(cmd_��Ժʱ���ڻ�������).AddItem "1-��鲢��ʾ"
    cbo(cmd_��Ժʱ���ڻ�������).AddItem "2-��鲢��ֹ"
    cbo(cmd_��Ժʱ���ڻ�������).ListIndex = 0
    
    cbo(cbo_����_����δ��ҩƷ���).AddItem "0-�����"
    cbo(cbo_����_����δ��ҩƷ���).AddItem "1-��鲢��ʾ"
    cbo(cbo_����_����δ��ҩƷ���).AddItem "2-��鲢��ֹ"
    cbo(cbo_����_����δ��ҩƷ���).ListIndex = 0
    
    cbo(cbo_����_����δִ����Ŀ���).AddItem "0-�����"
    cbo(cbo_����_����δִ����Ŀ���).AddItem "1-��鲢��ʾ"
    cbo(cbo_����_����δִ����Ŀ���).AddItem "2-��鲢��ֹ"
    cbo(cbo_����_����δִ����Ŀ���).ListIndex = 0
    
    
    vsUnCheckItem.ComboList = "..."
    
    '����ǩ����֤����
    If mobjESign Is Nothing Then
        On Error Resume Next
        Set mobjESign = CreateObject("zl9ESign.clsESign")
        Err.Clear: On Error GoTo 0
    End If
    If Not mobjESign Is Nothing Then
        strTmp = mobjESign.GetESignType()
        arrTemp = Split(strTmp, ",")
        For i = LBound(arrTemp) To UBound(arrTemp)
            cbo(cbo_����ǩ����֤����).AddItem arrTemp(i)
        Next
        If cbo(cbo_����ǩ����֤����).ListCount > 0 Then cbo(cbo_����ǩ����֤����).ListIndex = 0
    End If
    
    For i = 0 To sstSign.Tabs - 1
        sstSign.TabVisible(i) = False
        If i = sst_���� Then
            strTmp = " And t.������� IN (1,3)  and T.�������� IN ('�ٴ�','����','����')"
        ElseIf i = sst_סԺҽ�� Then
            strTmp = " And t.������� IN (2,3)  and T.�������� IN ('�ٴ�','����','����')"
        ElseIf i = sst_סԺ��ʿ Then
            strTmp = " And t.������� IN (2,3)  and T.��������='����'"
        ElseIf i = sst_ҽ�� Then
            strTmp = " And t.������� <> 0  and T.�������� IN('���','����','����','����','Ӫ��')"
        ElseIf i = sst_���� Then
            strTmp = " And t.������� IN (2,3)  and T.��������='����'"
        ElseIf i = sst_ҩƷ Then
            strTmp = " and T.�������� in('��ҩ��','��ҩ��','��ҩ��')"
        ElseIf i = sst_lis Then
            strTmp = " And t.������� <> 0  and T.��������='����'"
        ElseIf i = sst_Pacs Then
            strTmp = " And t.������� <> 0  and T.��������='���'"
        ElseIf i = sst_Ѫ�� Then
            strTmp = " And t.������� <> 0  and T.��������='Ѫ��'"
        End If
         '����Ĭ�ϲ���ѡ��
        gstrSQL = "Select Distinct D.ID, d.վ��,D.����, D.����,D.����" & vbNewLine & _
                "From ���ű� D, ��������˵�� T" & vbNewLine & _
                "Where d.Id = t.����id And (d.����ʱ�� Is Null Or d.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) " & strTmp & vbNewLine & _
                "order by վ��,����"
                
        On Error GoTo ErrHandle
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        With vsDept(i)
            .Rows = 1
            .MergeCells = flexMergeFree
            .MergeCol(col_վ��) = True
            .AllowUserResizing = flexResizeBoth
            .SelectionMode = flexSelectionByRow
            .Editable = flexEDKbdMouse
            .ExplorerBar = flexExSortShowAndMove
            .ColSort(col_ѡ��) = flexSortNone
            .Cell(flexcpPicture, 0, col_ѡ��) = imgCheck.ListImages("UnChecked").Picture
            .Cell(flexcpPictureAlignment, 0, col_ѡ��) = flexAlignCenterCenter
            blnTmp = False
            Do While Not rsTmp.EOF
                .AddItem ""
                .RowData(.Rows - 1) = Val(rsTmp!ID & "")
                .TextMatrix(.Rows - 1, col_վ��) = rsTmp!վ�� & ""
                If rsTmp!վ�� & "" <> "" Then
                    blnTmp = True
                Else
                    .TextMatrix(.Rows - 1, col_վ��) = " "
                End If
                .TextMatrix(.Rows - 1, col_����) = rsTmp!���� & ""
                .TextMatrix(.Rows - 1, col_����) = rsTmp!���� & ""
                .TextMatrix(.Rows - 1, col_����) = rsTmp!���� & ""
                
                rsTmp.MoveNext
            Loop
            .ColHidden(col_վ��) = Not blnTmp
        End With
    Next
    
    '����ǩ������
    Call cbo_Click(cbo_����ǩ����֤����)
    Call LoadSign
    
        
    '��ʼ���ؼ�
    With BillҩƷ���ұ��
        .Rows = 2
        .Cols = 3

        .TextMatrix(0, mGrdCol.ѡ��) = "ѡ��"
        .TextMatrix(0, mGrdCol.����) = "����"
        .TextMatrix(0, mGrdCol.����) = "����"

        '-1����ʾ���п���ѡ���ǲ����ͣ�"��"��" "��
        ' 0����ʾ���п���ѡ�񣬵������޸�
        ' 1����ʾ���п������룬�ⲿ��ʾΪ��ťѡ��
        ' 2����ʾ�����������У��ⲿ��ʾΪ��ťѡ�񣬵���������ѡ���
        ' 3����ʾ������ѡ���У��ⲿ��ʾΪ������ѡ��
        '4:  ��ʾ����Ϊ�������ı����û�����
        '5:  ��ʾ���в�����ѡ��

        .ColData(mGrdCol.ѡ��) = 5
        .ColData(mGrdCol.����) = 5
        .ColData(mGrdCol.����) = 4


        .ColWidth(mGrdCol.ѡ��) = 0
        .ColWidth(mGrdCol.����) = 2000
        .ColWidth(mGrdCol.����) = 1600
        
        .ColAlignment(mGrdCol.����) = 1
        
        .PrimaryCol = mGrdCol.����
        .LocateCol = mGrdCol.����
        .AllowAddRow = False
        .Active = True
    End With
    
    With Bill���Ŀ��ұ��
        .Rows = 2
        .Cols = 3

        .TextMatrix(0, mGrdCol.ѡ��) = "ѡ��"
        .TextMatrix(0, mGrdCol.����) = "����"
        .TextMatrix(0, mGrdCol.����) = "����"

        .ColData(mGrdCol.ѡ��) = 5
        .ColData(mGrdCol.����) = 5
        .ColData(mGrdCol.����) = 4


        .ColWidth(mGrdCol.ѡ��) = 0
        .ColWidth(mGrdCol.����) = 2000
        .ColWidth(mGrdCol.����) = 1600
        
        .ColAlignment(mGrdCol.����) = 1
        
        .PrimaryCol = mGrdCol.����
        .LocateCol = mGrdCol.����
        .AllowAddRow = False
        .Active = True
    End With
    
    'Ԥ��Ʊ�ݸ�ʽ
    Call InitBillForamt(vfgBillFormat)
    
    'Ris��������
    LoadRisEnables
    
    'RIS ��Ժ����
    LoadRisBranchHosp
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdCancel_Click()
    mblnOk = False
    Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim blnModi As Boolean, i As Long
    
    If Not mblnOk Then
        For i = 0 To vsDept.Count - 1
            With vsDept(i)
                If .Tag = "���޸�" Then
                    blnModi = True
                    Exit Sub
                End If
            End With
        Next
        If blnModi = False Then
            blnModi = lvw����.Tag = "���޸�" Or lvwNo.Tag = "���޸�" Or cbo(cbo_סԺ�Ź���).Tag = "���޸�" _
                Or cbo(cbo_����Ź���).Tag = "���޸�" Or cbo(cbo_���ۺŹ���).Tag = "���޸�" _
                Or BillҩƷ���ұ��.Tag = "���޸�" Or Bill���Ŀ��ұ��.Tag = "���޸�" Or vsfRISEnables.Tag = "���޸�" _
                Or vsfBranchHosp.Tag = "���޸�" Or txtMainHosp.Tag = "���޸�"
        End If
        
        If Not blnModi Then
            blnModi = ThirdSvrChanged
        End If
        
        mrsPar.Filter = "(�޸�״̬=1 ANd ErrType =Null) OR  (�޸�״̬=1 And ErrType=" & PET_ֵ���� & ")"
        If mrsPar.RecordCount > 0 Or blnModi Then
            If MsgBox("�����޸Ĳ��ֲ����������������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Cancel = 1: Exit Sub
            End If
        End If
    End If
    Set mobjESign = Nothing
    Set mrsPar = Nothing
    Set mrsSvr = Nothing
End Sub


Private Sub chk_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(chk, Index, mrsPar)
End Sub

Private Sub dtp_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(dtp, Index, mrsPar)
End Sub

Private Sub txt_Change(Index As Integer)
    If Me.Visible Then Call SetParChange(txt, Index, mrsPar)
End Sub

Private Sub txt_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txt(Index))
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        zlCommFun.PressKey vbKeyTab
    Else
        If Index = txt_��ϲ������� Then
            If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
        End If
    End If
End Sub

Private Sub txt_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(txt, Index, mrsPar)
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
    If Index = txt_��ϲ������� Then
        If Val(txt(Index).Text) <= 0 Then
            txt(Index).Text = 0
        ElseIf Val(txt(Index).Text) > 999 Then
            txt(Index).Text = 999
        End If
    End If
End Sub

Private Sub txtMainHosp_Change()
    If Me.Visible And txtMainHosp.Tag = "" Then txtMainHosp.Tag = "���޸�"
End Sub

Private Sub txtUD_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(txtUD, Index, mrsPar)
End Sub

Private Sub cbo_GotFocus(Index As Integer)
    If Index = cbo_���ۺŹ��� Then
        Call zlCommFun.ShowTipInfo(cbo(cbo_���ۺŹ���).hwnd, "����˵��" & vbCrLf & "����סԺ���۵Ǽ�ʱ�����ۺŵ����ɹ���", True, True, 8800)
    Else
        Call SetParTip(cbo, Index, mrsPar)
    End If
End Sub


Private Sub Initת�Ƴ�Ժ�������Ŀ(ByVal strIn As String)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    Dim lngRow As Long
    Dim lngCol As Long
    
    If strIn = "" Then Exit Sub
    strIn = Replace(strIn, "|", ",")
    strSQL = "select id,���� from ������ĿĿ¼ where id in (Select Column_Value From Table(Cast(f_Num2list([1]) As zlTools.t_Numlist))) Order by ����"
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, strIn)
    If rsTmp.EOF Then Exit Sub
    
    With vsUnCheckItem
        .Row = 0: .Col = 0
        For i = 1 To rsTmp.RecordCount
            .TextMatrix(.Row, .Col) = rsTmp!���� & ""
            .Cell(flexcpData, .Row, .Col) = rsTmp!ID & ""
            Call EnterNextCell(vsUnCheckItem)
            rsTmp.MoveNext
        Next
    End With
    
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Function Getת�Ƴ�Ժ�������Ŀ() As String
    Dim i As Integer
    Dim j As Integer
    Dim strIds As String
    
    With vsUnCheckItem
        For i = .FixedRows To .Rows - 1
            For j = .FixedCols To .Cols - 1
                If .TextMatrix(i, j) <> "" Then
                    strIds = strIds & "|" & Val(.Cell(flexcpData, i, j))
                End If
            Next
        Next
    End With
    Getת�Ƴ�Ժ�������Ŀ = Mid(strIds, 2)
End Function


Private Sub txtLocate_Change(Index As Integer)
    If Index = txt_Dept Then
        mlngPreFind = 1
    ElseIf Index = txt_Par Then
        txtLocate(Index).Tag = ""
    End If
End Sub

Private Sub txtLocate_GotFocus(Index As Integer)
    txtLocate(Index).SelStart = 0
    txtLocate(Index).SelLength = Len(txtLocate(Index).Text)
End Sub

Private Sub txtLocate_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        Dim strFind As String
        
        If Trim(txtLocate(Index).Text) = "" Then Exit Sub
        strFind = UCase(Trim(txtLocate(Index).Text))
        
        Select Case Index
        Case txt_Par
            Call LocatePar(txtLocate(Index), Me)
        Case txt_Dept
            
            If BillҩƷ���ұ��.Visible Then
                Call LocateDept(strFind, BillҩƷ���ұ��, 1) '��0����������
            Else
                Call LocateDeptSign(strFind)
            End If
        End Select
    End If
End Sub

Private Sub LocateDeptSign(ByVal strFind As String)
'���ܣ��������õ���ǩ���Ŀ���
    Dim i As Long
    
    With vsDept(sstSign.Tab)
        For i = mlngPreFind To .Rows - 1
            If .RowHidden(i) = False Then
                If .TextMatrix(i, col_����) Like IIF(gstrLike <> "", "*", "") & strFind & "*" Or _
                    .TextMatrix(i, col_����) Like IIF(gstrLike <> "", "*", "") & strFind & "*" Or .TextMatrix(i, col_����) = strFind Then
                    .Row = i: .ShowCell i, col_����
                    Exit For
                End If
            End If
        Next
        If i < .Rows Then
            mlngPreFind = i + 1
        Else
            If mlngPreFind = 1 Then
                MsgBox "û���ҵ�ƥ��Ŀ��ң�������������ݡ�", vbInformation, Me.Caption
                txtLocate(txt_Dept).SetFocus
            Else
                MsgBox "ȫ�������ˣ�����û���ˡ�", vbInformation, Me.Caption
                mlngPreFind = 1
            End If
        End If
    End With
End Sub

Private Sub LocateDept(ByVal strFind As String, ByRef objTmp As Object, ByVal lngCol As Long)
'���ܣ����ҿ���
'������lngCol-���в��ҵ���
    Dim i As Long, lngRows As Long, lngStart As Long
    Dim strCode As String, strName As String
    
    With objTmp
        If TypeName(objTmp) = "ListView" Then 'lvw
            lngRows = .ListItems.Count
            For i = mlngPreFind To lngRows
                If .ListItems(i).ListSubItems(lngCol).Text Like IIF(gstrLike <> "", "*", "") & strFind & "*" Then
                    Call .ListItems(i).EnsureVisible
                    .ListItems(i).Selected = True
                    .SetFocus
                    Exit For
                End If
            Next
        ElseIf TypeName(objTmp) = "ListBox" Then 'lst
            With objTmp
                lngRows = .ListCount - 1
                
                lngStart = IIF(mlngPreFind = 1, 0, mlngPreFind)
                For i = lngStart To .ListCount - 1
                    strCode = Split(.List(i), "-")(0)
                    strName = Split(.List(i), "-")(1)
                    If strCode Like strFind & "*" Or strName Like IIF(gstrLike <> "", "*", "") & strFind & "*" Then
                        .ListIndex = i
                        .SetFocus
                        Exit For
                    End If
                Next
            End With
        Else
            lngRows = objTmp.Rows
            For i = mlngPreFind To .Rows - 1
                If InStr(.TextMatrix(i, lngCol), "-") > 0 Then
                    strCode = Split(.TextMatrix(i, lngCol), "-")(0)
                    strName = Split(.TextMatrix(i, lngCol), "-")(1)
                Else
                    strCode = ""
                    strName = .TextMatrix(i, lngCol)
                End If
                
                If strCode Like strFind & "*" Or strName Like IIF(gstrLike <> "", "*", "") & strFind & "*" Then
                    objTmp.SetFocus
                    .Row = i: .Col = lngCol
                    .TopRow = i
                    Exit For
                End If
            Next
        End If
    End With
    If i < lngRows Then
        mlngPreFind = i + 1
    Else
        If mlngPreFind = 1 Then
            MsgBox "û���ҵ�ƥ��ģ�������������ݡ�", vbInformation, Me.Caption
            txtLocate(txt_Dept).SetFocus
        Else
            MsgBox "ȫ�������ˣ�����û���ˡ�", vbInformation, Me.Caption
            mlngPreFind = 1
        End If
    End If
End Sub

Private Sub vfgBillFormat_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim strValue As String, i As Integer
    Dim strCol As String
    
    If Me.Visible Then
        With vfgBillFormat
            If Col = vfgBillFormat.ColIndex("Ʊ�ݸ�ʽ") Or Col = vfgBillFormat.ColIndex("Ԥ����ӡ��ʽ") Then
                For i = 1 To .Rows - 1
                    If Trim(.TextMatrix(i, .ColIndex("ʹ�����"))) <> "" Then
                        If vfgBillFormat.ColIndex("Ʊ�ݸ�ʽ") = Col Then
                            strValue = strValue & "|" & Trim(.Cell(flexcpData, i, .ColIndex("ʹ�����"))) & "," & Val(.TextMatrix(i, .ColIndex("Ʊ�ݸ�ʽ")))
                        Else
                            strValue = strValue & "|" & Trim(.Cell(flexcpData, i, .ColIndex("ʹ�����"))) & "," & Val(Left(.TextMatrix(i, .ColIndex("Ԥ����ӡ��ʽ")), 1))
                        End If
                    End If
                Next
                If strValue <> "" Then strValue = Mid(strValue, 2)
                Call SetParChange(vfgBillFormat, 0, mrsPar, True, strValue, CStr(Col))
            End If
        End With
    End If
End Sub

Private Sub vfgBillFormat_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub vfgBillFormat_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    With vfgBillFormat
        If .MouseCol = .ColIndex("Ʊ�ݸ�ʽ") Then
            Call SetParTip(vfgBillFormat, 0, mrsPar, , , CStr(.ColIndex("Ʊ�ݸ�ʽ")))
        ElseIf .MouseCol = .ColIndex("Ԥ����ӡ��ʽ") Then
            Call SetParTip(vfgBillFormat, 0, mrsPar, , , CStr(.ColIndex("Ԥ����ӡ��ʽ")))
        End If
    End With
End Sub

Private Sub vsDept_CellChanged(Index As Integer, ByVal Row As Long, ByVal Col As Long)
    If Me.Visible And vsDept(Index).Tag = "" Then vsDept(Index).Tag = "���޸�"
End Sub

Private Sub vsfBranchHosp_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call vsfBranchHosp.AutoSize(Row, Col)
End Sub

Private Sub vsfBranchHosp_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim lngLen As Long
    
    If Me.Visible And vsfBranchHosp.Tag = "" Then vsfBranchHosp.Tag = "���޸�"
    
    If Col = col_RIS��Ժ���� Then
        lngLen = 100
    Else
        lngLen = 20
    End If
    
    If Len(vsfBranchHosp.TextMatrix(Row, Col)) > lngLen Then
        MsgBox "��Ӱ����Ϣϵͳ---HISҽԺ����---��Ժ���á��У�" & vbCrLf & vbCrLf & "���ݳ����涨���ȣ�ֻ����ǰ" & lngLen & "λ��", vbInformation, gstrSysName
        vsfBranchHosp.TextMatrix(Row, Col) = Left(vsfBranchHosp.TextMatrix(Row, Col), lngLen)
    End If
End Sub

Private Sub vsfBranchHosp_Click()
    If vsfBranchHosp.Rows = 1 Then
        vsfBranchHosp.Rows = 2
        vsfBranchHosp.TextMatrix(1, col_RIS��Ժ���) = 1
    End If
End Sub

Private Sub vsfBranchHosp_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsfBranchHosp
        If .Rows <= 1 Then Exit Sub
        
        If KeyCode = vbKeyReturn And .Col = col_ris��Ժ���ݿ������ And .Row = .Rows - 1 Then
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, col_RIS��Ժ���) = .Rows - 1
        End If
    End With
End Sub

Private Sub vsfRisDepts_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    If Col = col_Ris����ѡ�� Then
        '��RIS���Ʊ���д�������ƺ�ID��
         Call WriteDeptsIntoVsfRisEnables
    End If
End Sub

Private Sub vsfRisDepts_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> col_Ris����ѡ�� Then Cancel = True
End Sub

Private Sub vsfRisDepts_BeforeSort(ByVal Col As Long, Order As Integer)
    Dim i As Integer
    
    If Col = col_Ris����ѡ�� Then
        With vsfRisDepts
            If .MouseCol = col_Ris����ѡ�� And .MouseRow = 0 Then
                If .ColData(col_Ris����ѡ��) = RIS_Checked Then
                    .Cell(flexcpPicture, 0, col_Ris����ѡ��) = imgCheck.ListImages("UnChecked").Picture
                    .ColData(col_Ris����ѡ��) = ""
                Else
                    .Cell(flexcpPicture, 0, col_Ris����ѡ��) = imgCheck.ListImages("Checked").Picture
                    .ColData(col_Ris����ѡ��) = RIS_Checked
                End If
                
                For i = 1 To .Rows - 1
                    If .ColData(col_Ris����ѡ��) = RIS_Checked Then
                        .Cell(flexcpChecked, i, col_Ris����ѡ��) = 1
                    Else
                        .Cell(flexcpChecked, i, col_Ris����ѡ��) = 2
                    End If
                Next i
                
                Call WriteDeptsIntoVsfRisEnables
                
            End If
        End With
    End If
End Sub

Private Sub vsfRISEnables_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim i As Integer
    
    If Me.Visible And vsfRISEnables.Tag = "" Then vsfRISEnables.Tag = "���޸�"
    '���ѡ����ԤԼ���ҵ�ȫ���������ԤԼ����
    If Col = col_RIS����ԤԼ����ȫ Then
        With vsfRISEnables
            If .Cell(flexcpChecked, Row, col_RIS����ԤԼ����ȫ) = 1 Then
                .TextMatrix(Row, col_RIS����ԤԼ����ID) = ""
                .TextMatrix(Row, col_RIS����ԤԼ����) = ""
                Call .AutoSize(col_RIS���ÿ���, col_RIS����ԤԼ����)
            End If
        End With
    End If
    
    '���ȡ���˳��ϣ���ȡ��RIS���ң�ȡ��ԤԼ����
    If Col = col_RIS���ó��� Then
        With vsfRISEnables
            If .Cell(flexcpChecked, Row, col_RIS���ó���) = 2 Then
                .TextMatrix(Row, col_RIS���ÿ���) = ""
                .TextMatrix(Row, col_RIS���ÿ���ID) = ""
                .Cell(flexcpChecked, Row, col_RIS����ԤԼ����ȫ) = 2
                .TextMatrix(Row, col_RIS����ԤԼ����ID) = ""
                .TextMatrix(Row, col_RIS����ԤԼ����) = ""
                Call .AutoSize(col_RIS���ÿ���, col_RIS����ԤԼ����)
            End If
        End With
    End If
    '���ȡ���˼�������ȡ����������ѡ��
    If Col = col_RIS���ü������ Then
        With vsfRISEnables
            .Cell(flexcpChecked, Row, col_RIS���ó���) = 2
            .TextMatrix(Row, col_RIS���ÿ���) = ""
            .TextMatrix(Row, col_RIS���ÿ���ID) = ""
            .Cell(flexcpChecked, Row, col_RIS����ԤԼ����ȫ) = 2
            .TextMatrix(Row, col_RIS����ԤԼ����ID) = ""
            .TextMatrix(Row, col_RIS����ԤԼ����) = ""
            
            If .TextMatrix(Row + 1, col_RIS���ü������) = .TextMatrix(Row, col_RIS���ü������) Then
                .Cell(flexcpChecked, Row + 1, col_RIS���ó���) = 2
                .TextMatrix(Row + 1, col_RIS���ÿ���) = ""
                .TextMatrix(Row + 1, col_RIS���ÿ���ID) = ""
                .Cell(flexcpChecked, Row + 1, col_RIS����ԤԼ����ȫ) = 2
                .TextMatrix(Row + 1, col_RIS����ԤԼ����ID) = ""
                .TextMatrix(Row + 1, col_RIS����ԤԼ����) = ""
            End If
            
            If .TextMatrix(Row + 2, col_RIS���ü������) = .TextMatrix(Row, col_RIS���ü������) Then
                .Cell(flexcpChecked, Row + 2, col_RIS���ó���) = 2
                .TextMatrix(Row + 2, col_RIS���ÿ���) = ""
                .TextMatrix(Row + 2, col_RIS���ÿ���ID) = ""
                .Cell(flexcpChecked, Row + 2, col_RIS����ԤԼ����ȫ) = 2
                .TextMatrix(Row + 2, col_RIS����ԤԼ����ID) = ""
                .TextMatrix(Row + 2, col_RIS����ԤԼ����) = ""
            End If
            Call .AutoSize(col_RIS���ÿ���, col_RIS����ԤԼ����)
        End With
    End If
End Sub

Private Sub vsfRISEnables_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = col_RIS���ÿ��� Or Col = col_RIS����ԤԼ���� Then Cancel = True
    If Col = col_RIS���ó��� Then
        If vsfRISEnables.Cell(flexcpChecked, Row, col_RIS���ü������) = 2 Then Cancel = True
    End If
    If Col = col_RIS����ԤԼ����ȫ Then
        If vsfRISEnables.Cell(flexcpChecked, Row, col_RIS���ó���) = 2 Then Cancel = True
    End If
End Sub

Private Sub vsfRISEnables_CellChanged(ByVal Row As Long, ByVal Col As Long)
    If Me.Visible And vsfRISEnables.Tag = "" Then vsfRISEnables.Tag = "���޸�"
End Sub

Private Sub vsfRISEnables_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim strDeptIDs As String
    Dim lngSource As Long
    
    With vsfRISEnables
        '�����������ƿ��ҵ�ѡ����ѡ���˼�����ͺͳ��ϣ�����ѡ��RIS���ң���ѡ����RIS���ϣ�����ѡ��ԤԼ����
        lngSource = IIF(.TextMatrix(.MouseRow, col_RIS���ó���) = "����", 1, IIF(.TextMatrix(.MouseRow, col_RIS���ó���) = "סԺ", 2, 4))
        If .MouseRow >= 1 And (.MouseCol = col_RIS���ÿ��� Or .MouseCol = col_RIS����ԤԼ����) And lngSource <> 4 And .Cell(flexcpChecked, .MouseRow, col_RIS���ü������) = 1 And .Cell(flexcpChecked, .MouseRow, col_RIS���ó���) = 1 Then
            strDeptIDs = vsfRISEnables.TextMatrix(vsfRISEnables.MouseRow, IIF(vsfRISEnables.MouseCol = col_RIS���ÿ���, col_RIS���ÿ���ID, col_RIS����ԤԼ����ID))
            Call LoadRisDepts(strDeptIDs, lngSource)
            vsfRisDepts.Visible = True
        Else
            vsfRisDepts.Visible = False
        End If
    End With
End Sub

Private Sub vsgInput_DblClick(Index As Integer)
    Call SetInputItemValue(Index)
End Sub

Private Sub vsgInput_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        With vsgInput(Index)
            Select Case Index
            Case VSGInput_������Ժ����������, VSGInput_������Ϣ����������
                If .Row = .Rows - 1 And .Col = .Cols - 1 Then
                   If cmdOk.Enabled And cmdOk.Visible Then cmdOk.SetFocus
                   Exit Sub
                End If
                
                zlVsMoveGridCell vsgInput(Index), 1, .Cols - 1
            Case Else
            End Select
        End With
        Exit Sub
    End If
    
    If KeyAscii <> vbKeySpace Then Exit Sub
    Call SetInputItemValue(Index)
End Sub

Private Sub vsgInput_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(vsgInput, Index, mrsPar)
End Sub

Private Sub vsUnCheckItem_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    vsUnCheckItem.ComboList = "..."
End Sub

Private Sub vsUnCheckItem_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim vPoint As POINTAPI
    Dim blnCancel As Boolean
    
    On Error GoTo errH
    strSQL = "select A.ID,A.����,A.���� from ������ĿĿ¼ A Where A.��� not in('4','5','6','7') and (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null) Order By ����"
    With vsUnCheckItem
        vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
        Set rsTmp = zlDatabase.ShowSelect(Me, strSQL, 0, "������Ŀ", , , , , , True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, , True)
        If Not rsTmp Is Nothing Then
            If SetItemInput(Row, Col, rsTmp) Then
                Call vsUnCheckItem_AfterRowColChange(-1, -1, Row, Col)
            End If
        Else
            If Not blnCancel Then
                MsgBox "û�п��õ�������Ŀ�����ȵ�������Ŀ���������á�", vbInformation, gstrSysName
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsUnCheckItem_CellChanged(ByVal Row As Long, ByVal Col As Long)
    Dim strValue As String
    
    If Me.Visible Then
        strValue = Getת�Ƴ�Ժ�������Ŀ
        Call SetParChange(vsUnCheckItem, 0, mrsPar, True, strValue)
    End If
End Sub

Private Sub vsUnCheckItem_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTmp As New ADODB.Recordset
    With vsUnCheckItem
        If KeyCode > 127 Then
            '���ֱ�����뺺�ֵ�����
            Call vsUnCheckItem_KeyPress(KeyCode)
            
        ElseIf KeyCode = vbKeyDelete Then
            .TextMatrix(.Row, .Col) = ""
            .Cell(flexcpData, .Row, .Col) = ""
             
        ElseIf KeyCode = vbKeyReturn Then
            Call EnterNextCell(vsUnCheckItem)
        End If
        
    End With
End Sub

Private Sub vsUnCheckItem_KeyPress(KeyAscii As Integer)
    With vsUnCheckItem
        If KeyAscii = 13 Then
            KeyAscii = 0
        Else
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                Call vsUnCheckItem_CellButtonClick(.Row, .Col)
            Else
                If KeyAscii = vbKeyBack Then Exit Sub
                .ComboList = "" 'ʹ��ť״̬��������״̬
            End If
        End If
    End With
End Sub

Private Sub vsUnCheckItem_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Call SetParTip(vsUnCheckItem, 0, mrsPar)
End Sub

Private Sub vsUnCheckItem_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strInput As String
    Dim vPoint As POINTAPI, blnCancel As Boolean
    
    With vsUnCheckItem
        If .EditText = CStr(.TextMatrix(Row, Col)) Then
            Call EnterNextCell(vsUnCheckItem)
            Exit Sub
        End If
        strInput = UCase(.EditText)
        strSQL = "select DISTINCT A.ID,A.����,A.���� from ������ĿĿ¼ A, ������Ŀ���� B where " & _
            " a.Id = b.������Ŀid And B.����=1 And B.����=1 And A.��� not in('4','5','6','7') and (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
            " And (Upper(A.����) Like [1] Or Upper(A.����) Like [2] Or Upper(B.����) Like [2])" & _
            " Order by A.����"
        With vsUnCheckItem
            vPoint = zlControl.GetCoordPos(.hwnd, .CellLeft, .CellTop)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, strSQL, 0, "������Ŀ", False, "", "", False, False, True, _
                vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, strInput & "%", gstrLike & strInput & "%")
            If Not rsTmp Is Nothing Then
                If SetItemInput(Row, Col, rsTmp) Then
                    .EditText = .TextMatrix(Row, Col)
                     Call EnterNextCell(vsUnCheckItem)
                     Exit Sub
                End If
            Else
                If Not blnCancel Then
                    MsgBox "û���ҵ�ƥ��Ŀ��ҡ�", vbInformation, gstrSysName
                End If
                Cancel = True
            End If
        End With
        Call vsUnCheckItem_AfterRowColChange(Row, Col, Row, Col) '������ʾ��ť
    End With
End Sub


Private Function SetItemInput(ByVal lngRow As Long, ByVal lngCol As Long, ByVal rsTmp As ADODB.Recordset) As Boolean
    '�ȼ���±�����Ƿ����
    Dim i As Long, j As Long
    
    With vsUnCheckItem
        For i = .FixedCols To .Cols - 1
            For j = .FixedRows To .Rows - 1
                If .Cell(flexcpData, j, i) = rsTmp!ID & "" Then
                    MsgBox "��������Ŀ�Ѿ������б��У���鿴��", vbInformation, gstrSysName
                    Exit Function
                End If
            Next
        Next
        
        .Cell(flexcpData, lngRow, lngCol) = rsTmp!ID & ""
        .TextMatrix(lngRow, lngCol) = rsTmp!���� & ""
        SetItemInput = True
        
    End With
End Function


Private Sub Save����ǩ��()
    Dim i As Integer, j As Long
    Dim strDept As String
    
    On Error GoTo ErrHandle
    For i = 0 To vsDept.Count - 1
        With vsDept(i)
            If .Tag = "���޸�" Then
                strDept = ""
                For j = 1 To .Rows - 1
                    If .Cell(flexcpChecked, j, col_ѡ��) = 1 Then
                        strDept = strDept & "," & .RowData(j)
                    End If
                Next
                gstrSQL = "Zl_����ǩ�����ò���_Update(" & i & ",'" & Mid(strDept, 2) & "')"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                
                .Tag = ""
            End If
        End With
    Next
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub Save�����ӿ�()
    Dim i As Integer
    
    On Error GoTo ErrHandle
    If lvw����.Tag = "���޸�" Then
        For i = 1 To lvw����.ListItems.Count
            With lvw����.ListItems(i)
                gstrSQL = "Zl_����Ŀ¼_����(" & Mid(.Key, 2) & "," & IIF(.SubItems(4) <> "", 1, 0) & ")"
            End With
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        Next
        lvw����.Tag = ""
    End If
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub lvw����_DblClick()
    If Not lvw����.SelectedItem Is Nothing Then
        If lvw����.SelectedItem.SubItems(4) <> "" Then
            lvw����.SelectedItem.SubItems(4) = ""
        Else
            lvw����.SelectedItem.SubItems(4) = "��"
        End If
        lvw����.Tag = "���޸�"
        
        Call lvw����_ItemClick(lvw����.SelectedItem)
    End If
End Sub

Private Sub lvw����_ItemClick(ByVal Item As MSComctlLib.ListItem)
    cmd��������.Enabled = Item.SubItems(4) <> ""
End Sub

Private Sub lvw����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 32 Then
        Call lvw����_DblClick
    End If
End Sub


Private Sub cmd��������_Click()
    Dim objCommunity As Object
    
    If lvw����.SelectedItem Is Nothing Then Exit Sub
    If lvw����.SelectedItem.SubItems(4) = "" Then
        MsgBox lvw����.SelectedItem.SubItems(1) & "û�����á�", vbInformation, gstrSysName
        Exit Sub
    End If
    
    '�ȱ����������ݣ���Ϊ�ӿڳ�ʼ��Ҫ�ж��Ƿ�����
    Call Save�����ӿ�
    
    '��������
    Err.Clear: On Error Resume Next
    Set objCommunity = CreateObject("zlCommunity.clsCommunity")
    Err.Clear: On Error GoTo 0
    
    '���ù���
    If Not objCommunity Is Nothing Then
        If objCommunity.Initialize(gcnOracle) Then
            Call objCommunity.Setup(Val(Mid(lvw����.SelectedItem.Key, 2)))
        End If
    Else
        MsgBox "���������ӿ�û����ȷ��װ��", vbExclamation, gstrSysName
    End If
    
    Set objCommunity = Nothing
    Exit Sub
errH:
    gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub


Private Function Load�����ӿ�() As Boolean
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String
    Dim ObjItem As ListItem
    
    On Error GoTo errH
    
    strSQL = "Select ���, ����, ˵��, ����, ������ From ����Ŀ¼ Order by ���"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Do While Not rsTmp.EOF
        Set ObjItem = lvw����.ListItems.Add(, "_" & rsTmp!���, rsTmp!���)
        ObjItem.SubItems(1) = rsTmp!����
        ObjItem.SubItems(2) = NVL(rsTmp!˵��)
        ObjItem.SubItems(3) = rsTmp!������
        ObjItem.SubItems(4) = IIF(NVL(rsTmp!����, 0) = 1, "��", "")
        rsTmp.MoveNext
    Loop
    
    If Not lvw����.SelectedItem Is Nothing Then
        Call lvw����_ItemClick(lvw����.SelectedItem)
    End If
    Load�����ӿ� = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function


Private Function Load���ݱ������() As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim lst As ListItem
    
    gstrSQL = "" & _
        "   Select ��Ŀ���,��Ŀ����,��Ź���,decode(��Ź���,2,'2-��ִ�п��ҷ��±��',0,'0-����˳����',1,'1-����˳����','0-����˳����') as ��Ź���˵�� " & _
        "   From ������Ʊ� " & _
        "   where ��Ŀ��� in ( 11,12,13,14,15,16,21,22,23,24,25,26,27,28,29,32,62,68,69,70,71,72,73,74,75,76,77) order by ��Ŀ��� "
    
    Err = 0: On Error GoTo ErrHand:
    Load���ݱ������ = False
    zlDatabase.OpenRecordset rsTmp, gstrSQL, Me.Caption
    
    With rsTmp
        lvwNo.ListItems.Clear
        Do While Not rsTmp.EOF
            Set lst = lvwNo.ListItems.Add(, "K" & NVL(!��Ŀ���, 0), NVL(!��Ŀ����))
            lst.SubItems(1) = NVL(!��Ź���˵��)
            If NVL(!��Ŀ���) >= 1 And NVL(!��Ŀ���) <= 16 Then
                lst.ForeColor = &HC85422
                lvwNo.ListItems("K" & NVL(!��Ŀ���, 0)).ListSubItems(1).ForeColor = &HC85422
            End If
            If NVL(!��Ŀ���) >= 21 And NVL(!��Ŀ���) <= 62 Then
                lst.ForeColor = &H68588
                lvwNo.ListItems("K" & NVL(!��Ŀ���, 0)).ListSubItems(1).ForeColor = &H68588
            End If
            If NVL(!��Ŀ���) >= 68 And NVL(!��Ŀ���) <= 77 Then
                lst.ForeColor = &H856701
                lvwNo.ListItems("K" & NVL(!��Ŀ���, 0)).ListSubItems(1).ForeColor = &H856701
            End If
            lst.Tag = NVL(!��Ź���, 0)
            If lvwNo.SelectedItem Is Nothing Then
                lst.Selected = True
            End If
            .MoveNext
        Loop
    End With
    
    '2-סԺ�ţ�3-����ţ�6-סԺ���غ�
    Set rsTmp = New ADODB.Recordset
    gstrSQL = "Select ��Ŀ���,��Ź��� as ����ֵ From ������Ʊ� Where ��Ŀ��� in (2,3,6)"
    zlDatabase.OpenRecordset rsTmp, gstrSQL, Me.Caption
    
    rsTmp.Filter = "��Ŀ���=2"
    If rsTmp.RecordCount > 0 Then cbo(cbo_סԺ�Ź���).ListIndex = Val("" & rsTmp!����ֵ)
    rsTmp.Filter = "��Ŀ���=3"
    If rsTmp.RecordCount > 0 Then cbo(cbo_����Ź���).ListIndex = Val("" & rsTmp!����ֵ)
    rsTmp.Filter = "��Ŀ���=6"
    If rsTmp.RecordCount > 0 Then cbo(cbo_���ۺŹ���).ListIndex = Val("" & rsTmp!����ֵ)
    
    Load���ݱ������ = True

    Exit Function
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub LoadSign()
'���ܣ����ص���ǩ�����ò���
    Dim rsTmp As New Recordset
    Dim i As Long, lngTmp As Long
    
    gstrSQL = "select ����ID,���� from ����ǩ�����ò���"
    On Error GoTo ErrHandle
     Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
    For i = 0 To vsDept.Count - 1
        With vsDept(i)
            rsTmp.Filter = "����=" & i
            If rsTmp.RecordCount > 0 Then rsTmp.MoveFirst
            Do While Not rsTmp.EOF
                lngTmp = .FindRow(Val(rsTmp!����ID & ""))
                If lngTmp <> -1 Then
                    .Cell(flexcpChecked, lngTmp, col_ѡ��) = 1
                End If
                rsTmp.MoveNext
            Loop
            
        End With
    Next
    For i = 0 To sstSign.Tabs - 1
        If sstSign.TabVisible(i) = True Then sstSign.Tab = i: Exit For
    Next
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub



Private Sub lvwNo_DblClick()
    If lvwNo.SelectedItem Is Nothing Then Exit Sub
    Call Set���ݱ������
End Sub

Private Sub lvwNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = 32 Then
        If lvwNo.SelectedItem Is Nothing Then Exit Sub
        Call Set���ݱ������
    End If
End Sub

Private Sub Set���ݱ������()
'�ı�������
    Dim strNo As String
    
    If lvwNo.SelectedItem Is Nothing Then Exit Sub
    strNo = lvwNo.SelectedItem.SubItems(1) & "-"
    Select Case Split(strNo, "-")(0)
        Case 0
            If Mid(lvwNo.SelectedItem.Key, 2) >= 11 And Mid(lvwNo.SelectedItem.Key, 2) <= 16 Then
                strNo = "1-����˳����"
                lvwNo.SelectedItem.Tag = "1"
            Else
                strNo = "2-��ִ�п��ҷ��±��"
                lvwNo.SelectedItem.Tag = "2"
            End If
        Case 1
            strNo = "0-����˳����"
            lvwNo.SelectedItem.Tag = "0"
        Case 2
            strNo = "0-����˳����"
            lvwNo.SelectedItem.Tag = "0"
    End Select
    lvwNo.SelectedItem.SubItems(1) = strNo
    
    lvwNo.Tag = "���޸�"
End Sub

Sub Save���ݱ������()
    Dim lst As ListItem
    
    On Error GoTo ErrHandle
    If lvwNo.Tag = "���޸�" Then
        For Each lst In lvwNo.ListItems
            gstrSQL = "ZL_������Ʊ�_Rule(" & Mid(lst.Key, 2) & "," & Val(lst.Tag) & ")"
            zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
        Next
        lvwNo.Tag = ""
    End If
    
    '2-סԺ��,3-�����
    If cbo(cbo_סԺ�Ź���).Tag = "���޸�" Then
        gstrSQL = "ZL_������Ʊ�_Rule(2," & cbo(cbo_סԺ�Ź���).ListIndex & ")"
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
        cbo(cbo_סԺ�Ź���).Tag = ""
    End If
    
    If cbo(cbo_����Ź���).Tag = "���޸�" Then
        gstrSQL = "ZL_������Ʊ�_Rule(3," & cbo(cbo_����Ź���).ListIndex & ")"
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
        cbo(cbo_����Ź���).Tag = ""
    End If
    
    If cbo(cbo_���ۺŹ���).Tag = "���޸�" Then
        gstrSQL = "ZL_������Ʊ�_Rule(6," & cbo(cbo_���ۺŹ���).ListIndex & ")"
        zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
        cbo(cbo_���ۺŹ���).Tag = ""
    End If
    
    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub



Private Sub vsDept_BeforeEdit(Index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> col_ѡ�� Then Cancel = True
End Sub

Private Sub vsDept_BeforeSort(Index As Integer, ByVal Col As Long, Order As Integer)
    Dim i As Long
    
    If Col = col_ѡ�� Then
        Order = 0
        With vsDept(Index)
            If .MouseCol = col_ѡ�� And .MouseRow = .FixedRows - 1 Then
                If sstSign.Enabled = False Then Exit Sub
                If .ColData(col_ѡ��) = "Check" Then
                    .Cell(flexcpPicture, 0, col_ѡ��) = imgCheck.ListImages("UnChecked").Picture
                    .ColData(col_ѡ��) = ""
                Else
                    .Cell(flexcpPicture, 0, col_ѡ��) = imgCheck.ListImages("Checked").Picture
                    .ColData(col_ѡ��) = "Check"
                End If
                For i = 1 To .Rows - 1
                    If .ColData(col_ѡ��) = "Check" Then
                        .Cell(flexcpChecked, i, col_ѡ��) = 1
                    Else
                        .Cell(flexcpChecked, i, col_ѡ��) = 2
                    End If
                    
                Next
            End If
        End With
    End If
End Sub

Private Sub vsDept_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeySpace Then
        If vsDept(Index).Row > 0 Then
            vsDept(Index).Cell(flexcpChecked, vsDept(Index).Row, col_ѡ��) = IIF(vsDept(Index).Cell(flexcpChecked, vsDept(Index).Row, col_ѡ��) = 1, 2, 1)
        End If
    End If
End Sub

Private Sub txtUD_Validate(Index As Integer, Cancel As Boolean)
    If Val(txtUD(Index).Text) > ud(Index).Max Or Val(txtUD(Index).Text) < ud(Index).Min Then
        txtUD(Index).Text = ud(Index).value
    End If
End Sub

Private Sub txtUD_Change(Index As Integer)
    If Me.Visible Then Call SetParChange(txtUD, Index, mrsPar)
End Sub

Private Sub txtUD_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txtUD(Index))
End Sub

Private Sub txtUD_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    End If
End Sub

Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub


Private Sub sstSign_Click(PreviousTab As Integer)
    mlngPreFind = 1
End Sub

Private Sub dtp_Change(Index As Integer)
    Dim intNext As Integer
    Dim blnValue As Boolean, strValue As String

    If Index < dtp_�����°� Then
        intNext = Index + 1
        
        dtp(intNext).MinDate = dtp(Index).value
        If dtp(intNext).value < dtp(intNext).MinDate Then
            dtp(intNext).value = dtp(intNext).MinDate
            dtp_Change intNext
        End If
    End If
    
    If Me.Visible Then
        Select Case Index
        Case dtp_�����ϰ�, dtp_�����°�
            blnValue = True
            strValue = Format(dtp(dtp_�����ϰ�).value, "HH:mm") & " AND " & Format(dtp(dtp_�����°�).value, "HH:mm")
            If Index = dtp_�����ϰ� Then
                Call SetParChange(dtp, dtp_�����ϰ�, mrsPar, blnValue, strValue)
            Else
                Call SetParChange(dtp, dtp_�����ϰ�, mrsPar, blnValue, strValue)
            End If
            Exit Sub
        Case dtp_�����ϰ�, dtp_�����°�
            blnValue = True
            strValue = Format(dtp(dtp_�����ϰ�).value, "HH:mm") & " AND " & Format(dtp(dtp_�����°�).value, "HH:mm")
            If Index = dtp_�����ϰ� Then
                Call SetParChange(dtp, dtp_�����°�, mrsPar, blnValue, strValue)
            Else
                Call SetParChange(dtp, dtp_�����ϰ�, mrsPar, blnValue, strValue)
            End If
            Exit Sub
        End Select
        
        Call SetParChange(dtp, Index, mrsPar, blnValue, strValue)
    End If
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo_Click(Index As Integer)
    Dim strTmp As String, i As Long
    Dim arrTmp As Variant
    Dim blnValue As Boolean, strValue As String
    
    Select Case Index
    Case cbo_����ǩ����֤����
        strTmp = chk_Sign_���� & "," & chk_Sign_סԺ & "," & chk_Sign_ҽ�� & "," & chk_Sign_���� & "," & _
                chk_Sign_ҩƷ & "," & chk_Sign_lis & "," & chk_Sign_pacs & "," & chk_sign_Ѫ��
        arrTmp = Split(strTmp, ",")
        
        chk(chk_�¿�һ��ҽ��ǩ��һ��).Enabled = cbo(Index).ListIndex <> 0
        If cbo(Index).ListIndex = 0 Then chk(chk_�¿�һ��ҽ��ǩ��һ��).value = 0
        sstSign.Enabled = cbo(Index).ListIndex <> 0
        
        If cbo(Index).ListIndex = 0 Then
            For i = 1 To 8
                chk(arrTmp(i - 1)).value = 0
                chk(arrTmp(i - 1)).Enabled = False
            Next
            sstSign.TabVisible(sst_����) = True
        Else
            If Not chk(chk_Sign_����).Enabled Then
                chk(chk_Sign_����).value = 1
            End If
            
            For i = 1 To 8
                chk(arrTmp(i - 1)).Enabled = True
            Next
        End If
        
        blnValue = True
        strValue = Val(cbo(Index).List(cbo(Index).ListIndex))
        
        cmd(cmd_����ǩ������).Visible = False
        If cbo(Index).ListIndex <> 0 Then
            If Not mobjESign Is Nothing Then
                cmd(cmd_����ǩ������).Visible = mobjESign.SetEnabled(Val(strValue))
            End If
        End If
    Case cbo_���������Դ
        blnValue = True
        strValue = cbo(cbo_���������Դ).ListIndex + 1
    Case cbo_�����������, cbo_סԺ�������
        blnValue = True
        strValue = (cbo(cbo_�����������).ListIndex + 1) & (cbo(cbo_סԺ�������).ListIndex + 1)
        If Index = cbo_����������� Then
            If Me.Visible Then Call SetParChange(cbo, cbo_סԺ�������, mrsPar, blnValue, strValue)
        Else
            If Me.Visible Then Call SetParChange(cbo, cbo_�����������, mrsPar, blnValue, strValue)
        End If
    Case cbo_סԺ�Ź���, cbo_����Ź���, cbo_���ۺŹ���
        If Me.Visible Then cbo(Index).Tag = "���޸�"
    
    End Select
    
    If Me.Visible Then Call SetParChange(cbo, Index, mrsPar, blnValue, strValue)
End Sub


Private Sub chk_Click(Index As Integer)
    Dim blnValue As Boolean, strValue As String
    
    Select Case Index
    Case chk_Sign_����, chk_Sign_סԺ, chk_Sign_ҽ��, chk_Sign_����, chk_Sign_ҩƷ, chk_Sign_lis, chk_Sign_pacs, chk_sign_Ѫ��
            
        '��ʹ�õ���ǩ��������£�������һ��������Ҫ����ǩ��
        If cbo(cbo_����ǩ����֤����).ListIndex <> 0 Then
            If chk(chk_Sign_����).value = 0 And chk(chk_Sign_סԺ).value = 0 _
                And chk(chk_Sign_ҽ��).value = 0 And chk(chk_Sign_����).value = 0 And chk(chk_Sign_ҩƷ).value = 0 _
                And chk(chk_Sign_lis).value = 0 And chk(chk_Sign_pacs).value = 0 And chk(chk_sign_Ѫ��).value = 0 Then
                    If Index = chk_Sign_���� Then
                        chk(chk_Sign_ҩƷ).value = 1
                    ElseIf Index = chk_Sign_ҩƷ Then
                         chk(chk_Sign_lis).value = 1
                    ElseIf Index = chk_Sign_lis Then
                         chk(chk_Sign_pacs).value = 1
                    ElseIf Index = chk_Sign_pacs Then
                        chk(chk_sign_Ѫ��).value = 1
                    ElseIf Index = chk_sign_Ѫ�� Then
                        chk(chk_Sign_����).value = 1
                    Else
                        chk(((Index - chk_Sign_���� + 1) Mod 4) + chk_Sign_����).value = 1
                    End If
            End If
        End If
        If Index = chk_Sign_���� Then
            sstSign.TabVisible(sst_����) = chk(chk_Sign_����).value = 1
        ElseIf Index = chk_Sign_ҩƷ Then
             sstSign.TabVisible(sst_ҩƷ) = chk(chk_Sign_ҩƷ).value = 1
        ElseIf Index = chk_Sign_lis Then
             sstSign.TabVisible(sst_lis) = chk(chk_Sign_lis).value = 1
        ElseIf Index = chk_Sign_pacs Then
             sstSign.TabVisible(sst_Pacs) = chk(chk_Sign_pacs).value = 1
        ElseIf Index = chk_Sign_���� Then
            sstSign.TabVisible(sst_����) = chk(chk_Sign_����).value = 1
        ElseIf Index = chk_Sign_סԺ Then
            sstSign.TabVisible(sst_סԺ��ʿ) = chk(chk_Sign_סԺ).value = 1
            sstSign.TabVisible(sst_סԺҽ��) = chk(chk_Sign_סԺ).value = 1
        ElseIf Index = chk_Sign_ҽ�� Then
            sstSign.TabVisible(sst_ҽ��) = chk(chk_Sign_ҽ��).value = 1
        ElseIf Index = chk_sign_Ѫ�� Then
            sstSign.TabVisible(sst_Ѫ��) = chk(chk_sign_Ѫ��).value = 1
        End If
        
        blnValue = True
        strValue = chk(chk_Sign_����).value & chk(chk_Sign_סԺ).value & chk(chk_Sign_ҽ��).value & _
                   chk(chk_Sign_����).value & chk(chk_Sign_ҩƷ).value & chk(chk_Sign_lis).value & chk(chk_Sign_pacs).value & chk(chk_sign_Ѫ��).value
                   
    Case chk_ȫ����ֻ�����, chk_ȫ��ĸֻ�����
        
        blnValue = True
        strValue = chk(chk_ȫ����ֻ�����).value & chk(chk_ȫ��ĸֻ�����).value
        If Index = chk_ȫ����ֻ����� Then
            If Me.Visible Then Call SetParChange(chk, chk_ȫ��ĸֻ�����, mrsPar, blnValue, strValue)
        Else
            If Me.Visible Then Call SetParChange(chk, chk_ȫ����ֻ�����, mrsPar, blnValue, strValue)
        End If
    Case chk_���˵�ַ�ṹ��¼��
        chk(chk_�����ַ�ṹ��¼��).Enabled = chk(chk_���˵�ַ�ṹ��¼��).value = 1
        If Not chk(chk_�����ַ�ṹ��¼��).Enabled Then chk(chk_�����ַ�ṹ��¼��).value = 0
    Case chk_����ҽѧӰ����Ϣϵͳרҵ��ӿ�
        sstRIS.Enabled = chk(chk_����ҽѧӰ����Ϣϵͳרҵ��ӿ�).value = 1
        vsfRISEnables.Enabled = sstRIS.Enabled
        chkShowSel.Enabled = vsfRISEnables.Enabled
    End Select
    
    If Me.Visible Then Call SetParChange(chk, Index, mrsPar, blnValue, strValue)
End Sub



Private Sub LoadҩƷ���Ŀ��ұ��()
'���ܣ���ȡ���ݲ���ʾ����
    Dim lng��� As Long, str�ⷿID As String
    Dim rsTmp As New ADODB.Recordset
    Dim strType As String
    Dim strSequence As String
    
    'ҩƷ����
    On Error GoTo ErrHandle
    strType = "('��ҩ��','��ҩ��','��ҩ��','�Ƽ���', '��ҩ��', '��ҩ��', '��ҩ��')"
    strSequence = "(21,22,23,24,25,26,27,28,29,32,62)"
    gstrSQL = "" & _
        "   Select distinct a.ID,a.����,a.����,b.��� " & _
        "   From ���ű� A,���Һ���� b" & _
        "   Where a.id=b.����id and a.ID in (select distinct ����id from ��������˵�� where �������� in " & strType & ")" & _
        "   And b.��Ŀ��� In " & strSequence & " " & _
        "   UNION ALL " & _
        "   Select a.ID,a.����,a.����,'' As ��� " & _
        "   From ���ű� A " & _
        "   Where a.ID in (select distinct ����id from ��������˵�� " & _
        "   where �������� in " & strType & ")" & _
        "   And a.Id Not In(Select Distinct ����id From ���Һ���� Where ����id Is Not null " & _
        "   And ��Ŀ��� In " & strSequence & ") " & _
        "   ORDER BY ���� "
        
    zlDatabase.OpenRecordset rsTmp, gstrSQL, "��ȡ��صĿ���"
    
    With rsTmp
        str�ⷿID = ""
        Do While Not .EOF
            BillҩƷ���ұ��.TextMatrix(BillҩƷ���ұ��.Rows - 1, mGrdCol.����) = NVL(!����)
            BillҩƷ���ұ��.TextMatrix(BillҩƷ���ұ��.Rows - 1, mGrdCol.����) = NVL(!���)
            BillҩƷ���ұ��.RowData(BillҩƷ���ұ��.Rows - 1) = !ID
            BillҩƷ���ұ��.Rows = BillҩƷ���ұ��.Rows + 1
            str�ⷿID = str�ⷿID & "," & rsTmp!ID
            .MoveNext
        Loop
    End With
    
    If str�ⷿID <> "" Then
        str�ⷿID = Mid(str�ⷿID, 2)
        BillҩƷ���ұ��.Rows = BillҩƷ���ұ��.Rows - 1
        BillҩƷ���ұ��.Active = True
    Else
        BillҩƷ���ұ��.Active = False
    End If
    
    rsTmp.Close
    
    '���Ŀ���
    strType = "('�Ƽ���','���Ŀ�','����ⷿ')"
    strSequence = "(68,69,70,71,72,73,74,75,76,77)"
    gstrSQL = "" & _
        "   Select distinct a.ID,a.����,a.����,b.��� " & _
        "   From ���ű� A,���Һ���� b" & _
        "   Where a.id=b.����id and a.ID in (select distinct ����id from ��������˵�� where �������� in " & strType & ")" & _
        "   And b.��Ŀ��� In " & strSequence & " " & _
        "   UNION ALL " & _
        "   Select a.ID,a.����,a.����,'' As ��� " & _
        "   From ���ű� A " & _
        "   Where a.ID in (select distinct ����id from ��������˵�� " & _
        "   where �������� in " & strType & ")" & _
        "   And a.Id Not In(Select Distinct ����id From ���Һ���� Where ����id Is Not null " & _
        "   And ��Ŀ��� In " & strSequence & ") " & _
        "   ORDER BY ���� "

    zlDatabase.OpenRecordset rsTmp, gstrSQL, "��ȡ��صĿ���"
    
    With rsTmp
        str�ⷿID = ""
        Do While Not .EOF
            Bill���Ŀ��ұ��.TextMatrix(Bill���Ŀ��ұ��.Rows - 1, mGrdCol.����) = NVL(!����)
            Bill���Ŀ��ұ��.TextMatrix(Bill���Ŀ��ұ��.Rows - 1, mGrdCol.����) = NVL(!���)
            Bill���Ŀ��ұ��.RowData(Bill���Ŀ��ұ��.Rows - 1) = !ID
            Bill���Ŀ��ұ��.Rows = Bill���Ŀ��ұ��.Rows + 1
            str�ⷿID = str�ⷿID & "," & rsTmp!ID
            .MoveNext
        Loop
    End With
    
    If str�ⷿID <> "" Then
        str�ⷿID = Mid(str�ⷿID, 2)
        Bill���Ŀ��ұ��.Rows = Bill���Ŀ��ұ��.Rows - 1
        Bill���Ŀ��ұ��.Active = True
    Else
        Bill���Ŀ��ұ��.Active = False
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub Save���ұ��()
'���ܣ�������ұ��
    Dim i As Integer
    
    On Error GoTo ErrHandle
    With BillҩƷ���ұ��
        If .Tag = "���޸�" Then
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, mGrdCol.����)) <> "" Then  'And Trim(.TextMatrix(i, mGrdCol.����)) <> "" Then
                    '����ID_IN   IN ���ұ��.����ID%TYPE,
                    '���_IN     IN ���ұ��.���%TYPE
                    gstrSQL = "ZL_���Һ����_UPDATE("
                    gstrSQL = gstrSQL & .RowData(i) & ","
                    gstrSQL = gstrSQL & "'" & Trim(.TextMatrix(i, mGrdCol.����)) & "',1)"
                    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
                End If
            Next
            .Tag = ""
        End If
    End With

    With Bill���Ŀ��ұ��
        If .Tag = "���޸�" Then
            For i = 1 To .Rows - 1
                If Trim(.TextMatrix(i, mGrdCol.����)) <> "" Then 'And Trim(.TextMatrix(i, mGrdCol.����)) <> "" Then
                    '����ID_IN   IN ���ұ��.����ID%TYPE,
                    '���_IN     IN ���ұ��.���%TYPE
                    gstrSQL = "ZL_���Һ����_UPDATE("
                    gstrSQL = gstrSQL & .RowData(i) & ","
                    gstrSQL = gstrSQL & "'" & Trim(.TextMatrix(i, mGrdCol.����)) & "',2)"
                    zlDatabase.ExecuteProcedure gstrSQL, Me.Caption
                End If
            Next
            .Tag = ""
        End If
    End With

    Exit Sub
ErrHandle:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Sub
Function CheckNumberRule_Drug() As Boolean
 '����       ��鵥�ݱ�������Ƿ���"2"��
    Dim i As Integer
    With Me.lvwNo
        For i = 1 To .ListItems.Count
            If Mid(.ListItems(i).Key, 2) >= 21 And Mid(.ListItems(i).Key, 2) <= 62 Then
                If .ListItems(i).SubItems(1) = "2-��ִ�п��ҷ��±��" Then
                    CheckNumberRule_Drug = True
                    Exit For
                End If
            End If
        Next
    End With
    
End Function

Function CheckNumberRule_Stuff() As Boolean
'����       ��鵥�ݱ�������Ƿ���"2"��
    Dim i As Integer
    With Me.lvwNo
        For i = 1 To .ListItems.Count
            If Mid(.ListItems(i).Key, 2) >= 68 And Mid(.ListItems(i).Key, 2) <= 77 Then
                If .ListItems(i).SubItems(1) = "2-��ִ�п��ҷ��±��" Then
                    CheckNumberRule_Stuff = True
                    Exit For
                End If
            End If
        Next
    End With
    
End Function

Private Sub billҩƷ���ұ��_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub billҩƷ���ұ��_EditKeyPress(KeyAscii As Integer)
    If Not CheckNumberRule_Drug Then
        MsgBox "�������ÿ��ұ������Ϊ����ִ�п��ҷ��±�š��������ÿ��ұ��롣", vbOKOnly + vbInformation, gstrSysName
        KeyAscii = 0
        Exit Sub
    End If
End Sub


Private Sub billҩƷ���ұ��_EnterCell(Row As Long, Col As Long)
    With BillҩƷ���ұ��
        Select Case .Col
            Case mGrdCol.����
                .TxtCheck = True
                .MaxLength = 1
                .TextMask = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz123456789"
                mintLastRow_Drug = Row
                mintLastCol_Drug = Col
            End Select
    End With
End Sub

Private Sub billҩƷ���ұ��_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim strSQL As String
     
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    mstrLastCode_Drug = ""
    
    
    With BillҩƷ���ұ��
        .Tag = "���޸�"
        .Text = Replace(UCase(Trim(.Text)), "'", "")
        strKey = UCase(Trim(.Text))
        Select Case .Col
            Case mGrdCol.����
                If strKey <> "" Then
                    .Text = strKey
                End If
                If .Row = .Rows - 1 And .Col = 2 And .TextMatrix(.Row, .Col) <> "" Then
'                    zlCommFun.PressKey vbKeyTab
                    Bill���Ŀ��ұ��.SetFocus
                End If
            Case mGrdCol.����
        End Select
    End With
End Sub

Private Sub billҩƷ���ұ��_KeyPress(KeyAscii As Integer)
    If Not CheckNumberRule_Drug Then
        MsgBox "�������ÿ��ұ������Ϊ����ִ�п��ҷ��±�š��������ÿ��ұ��롣", vbOKOnly + vbInformation, gstrSysName
        KeyAscii = 0
        Exit Sub
    End If
    If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz123456789", Chr(KeyAscii)) > 0 Then
        mstrLastCode_Drug = Chr(KeyAscii)
    End If
End Sub

Private Sub bill���Ŀ��ұ��_BeforeDeleteRow(Row As Long, Cancel As Boolean)
    Cancel = True
End Sub

Private Sub bill���Ŀ��ұ��_EditKeyPress(KeyAscii As Integer)
    If Not CheckNumberRule_Stuff Then
        MsgBox "�������ÿ��ұ������Ϊ����ִ�п��ҷ��±�š��������ÿ��ұ��롣", vbOKOnly + vbInformation, gstrSysName
        KeyAscii = 0
        Exit Sub
    End If
End Sub


Private Sub bill���Ŀ��ұ��_EnterCell(Row As Long, Col As Long)
    With Bill���Ŀ��ұ��
        Select Case .Col
            Case mGrdCol.����
                .TxtCheck = True
                .MaxLength = 1
                .TextMask = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz123456789"
                mintLastRow_Stuff = Row
                mintLastCol_Stuff = Col
            End Select
    End With
End Sub

Private Sub bill���Ŀ��ұ��_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strKey As String
    Dim strSQL As String
     
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    mstrLastCode_Stuff = ""
    
    
    With Bill���Ŀ��ұ��
        .Tag = "���޸�"
        .Text = Replace(UCase(Trim(.Text)), "'", "")
        strKey = UCase(Trim(.Text))
        Select Case .Col
            Case mGrdCol.����
                
                If strKey <> "" Then
                    .Text = strKey
                End If
                If .Row = .Rows - 1 And .Col = 2 And .TextMatrix(.Row, .Col) <> "" Then
                    zlCommFun.PressKey vbKeyTab
                End If
            Case mGrdCol.����
        End Select
    End With
End Sub

Private Sub bill���Ŀ��ұ��_KeyPress(KeyAscii As Integer)
    If Not CheckNumberRule_Stuff Then
        MsgBox "�������ÿ��ұ������Ϊ����ִ�п��ҷ��±�š��������ÿ��ұ��롣", vbOKOnly + vbInformation, gstrSysName
        KeyAscii = 0
        Exit Sub
    End If
    If InStr("ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz123456789", Chr(KeyAscii)) > 0 Then
        mstrLastCode_Stuff = Chr(KeyAscii)
    End If
End Sub

Private Sub InitBillForamt(ByVal vfgBill As VSFlexGrid)
    '��ʼ��Ԥ��Ʊ�ݸ�ʽ��Ϣ
    Dim rsTemp As New ADODB.Recordset, strSQL As String
    
    On Error GoTo ErrHand
    strSQL = "" & _
        "   Select 'ʹ�ñ���ȱʡ��ʽ' as ˵��,0 as ���  From Dual Union ALL " & _
        "   Select B.˵��,B.���  " & _
        "   From zlReports A,zlRptFmts B" & _
        "   Where A.ID=B.����ID And A.���='" & "ZL" & glngSys \ 100 & "_BILL_1103" & "'  " & _
        "   Order by  ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "Ʊ�ݸ�ʽ")
    
    With vfgBill
        .Clear 1
        .ColComboList(.ColIndex("Ʊ�ݸ�ʽ")) = .BuildComboList(rsTemp, "���,*˵��", "���")
        .ColComboList(.ColIndex("Ԥ����ӡ��ʽ")) = "0-����ӡƱ��|1-�Զ���ӡƱ��|2-ѡ���Ƿ��ӡƱ��"
        
        .TextMatrix(1, 0) = "����Ԥ��"
        .Cell(flexcpData, 1, 0) = 1
        .TextMatrix(2, 0) = "סԺԤ��"
        .Cell(flexcpData, 2, 0) = 2
        .ColData(.ColIndex("Ʊ�ݸ�ʽ")) = "0"
        .ColData(.ColIndex("Ԥ����ӡ��ʽ")) = "0"
        .ForeColor = &H80000008:  .ForeColorFixed = &H80000008
        .Editable = flexEDKbdMouse
    End With
        
    Exit Sub
ErrHand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub LaodBillForamt(ByVal vfgBill As VSFlexGrid, ByVal strBillFormat As String, ByVal strPrintMode As String)
    Dim varData As Variant, varType As Variant
    Dim varTemp As Variant, varTemp1 As Variant
    Dim lngRow As Long, i As Long
    
    varData = Split(strBillFormat, "|")
    varType = Split(strPrintMode, "|")
    With vfgBill
        .Tag = ""
        .Clear 1
        .Rows = 3
        For lngRow = 1 To .Cols - 1
            .TextMatrix(lngRow, .ColIndex("Ԥ����ӡ��ʽ")) = "0-����ӡƱ��"
            .TextMatrix(lngRow, .ColIndex("Ʊ�ݸ�ʽ")) = "0"
            For i = 0 To UBound(varData)
                varTemp = Split(varData(i) & "," & ",", ",")
                If Trim(varTemp(0)) = Trim(.Cell(flexcpData, lngRow, 0)) Then
                    .TextMatrix(lngRow, .ColIndex("Ʊ�ݸ�ʽ")) = Val(varTemp(1)): Exit For
                End If
            Next
            For i = 0 To UBound(varType)
                varTemp1 = Split(varType(i) & "," & ",", ",")
                If Trim(varTemp1(0)) = Trim(.Cell(flexcpData, lngRow, 0)) Then
                    .TextMatrix(lngRow, .ColIndex("Ԥ����ӡ��ʽ")) = Decode(Val(varTemp1(1)), 0, "0-����ӡƱ��", 1, "1-�Զ���ӡƱ��", "2-ѡ���Ƿ��ӡƱ��")
                    Exit For
                End If
            Next
        Next
        If Val(.ColData(.ColIndex("Ԥ����ӡ��ʽ"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("Ԥ����ӡ��ʽ"), .Rows - 1, .ColIndex("Ԥ����ӡ��ʽ")) = vbBlue
        End If
        
        If Val(.ColData(.ColIndex("Ʊ�ݸ�ʽ"))) = 1 Then
            .Cell(flexcpForeColor, 0, .ColIndex("Ʊ�ݸ�ʽ"), .Rows - 1, .ColIndex("Ʊ�ݸ�ʽ")) = vbBlue
        End If
    End With
End Sub

Private Sub LoadInputItem(ByVal intIndex As Integer, ByVal strValue As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������������
    '���:intIndex-����ֵ
    '     strValue-ȱʡ����ֵ��������,��ʽ:������Ŀ,��ֹ¼��,����Ƿ�����,������|....
    '����:
    '����:2015-06-11 17:32:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, varTemp As Variant
    Dim intRow As Integer, i As Integer
    
    On Error GoTo ErrHandle
    varData = Split(strValue, "|")
    With vsgInput(intIndex)
        .redraw = flexRDNone
        .Clear 1
        If strValue = "" Then .Rows = 2: Exit Sub
        .Rows = 2: intRow = 1
        For i = 0 To UBound(varData)
            varTemp = Split(varData(i) & ",,,,", ",")
            If varTemp(0) <> "" Then
                .TextMatrix(intRow, .ColIndex("������Ŀ")) = varTemp(0)
                .TextMatrix(intRow, .ColIndex("��ֹ¼��")) = IIF(Val(varTemp(1)) = 1, "��", "")
                .TextMatrix(intRow, .ColIndex("������")) = IIF(Val(varTemp(2)) = 1, "��", "")
                .TextMatrix(intRow, .ColIndex("������")) = IIF(Val(varTemp(3)) = 1, "��", "")
                If .TextMatrix(intRow, .ColIndex("��ֹ¼��")) = "��" Then
                    .Cell(flexcpBackColor, intRow, .ColIndex("������"), intRow, .ColIndex("������")) = &H8000000F
                ElseIf .TextMatrix(intRow, .ColIndex("������")) = "��" _
                    Or .TextMatrix(intRow, .ColIndex("������")) = "��" Then
                    .Cell(flexcpBackColor, intRow, .ColIndex("��ֹ¼��")) = &H8000000F
                End If
                .Rows = .Rows + 1: intRow = intRow + 1
            End If
        Next
        If .Rows > 2 And Trim(.TextMatrix(.Rows - 1, .ColIndex("������Ŀ"))) = "" Then
            .Rows = .Rows - 1
        End If
'        .ColAlignment(.ColIndex("������Ŀ")) = flexAlignCenterCenter
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 0, .Cols - 1
        .redraw = flexRDBuffered
    End With
    Exit Sub
ErrHandle:
    vsgInput(intIndex).redraw = flexRDBuffered
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub SetInputItemValue(ByVal intIndex As Integer)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���õ�ǰ��Ŀ�����ֵ
    '���:intIndex-����ؼ����������ֵ
    '����:
    '����:2015-06-11 17:58:05
    '---------------------------------------------------------------------------------------------------------------------------------------------
       
    On Error GoTo ErrHandle
    With vsgInput(intIndex)
        Select Case .Col
        Case .ColIndex("��ֹ¼��")
            .TextMatrix(.Row, .ColIndex("��ֹ¼��")) = IIF(.TextMatrix(.Row, .ColIndex("��ֹ¼��")) = "", "��", "")
            If .TextMatrix(.Row, .ColIndex("��ֹ¼��")) = "��" Then
                .TextMatrix(.Row, .ColIndex("������")) = ""
                .TextMatrix(.Row, .ColIndex("������")) = ""
                .Cell(flexcpBackColor, .Row, .ColIndex("������"), .Row, .ColIndex("������")) = &H8000000F
            Else
                .Cell(flexcpBackColor, .Row, .ColIndex("������"), .Row, .ColIndex("������")) = &H8000000E
            End If
            .Cell(flexcpBackColor, .Row, .ColIndex("��ֹ¼��")) = &H8000000E
        Case .ColIndex("������")
        
            .TextMatrix(.Row, .ColIndex("������")) = IIF(.TextMatrix(.Row, .ColIndex("������")) = "", "��", "")
            If .TextMatrix(.Row, .ColIndex("������")) = "��" Then
                .TextMatrix(.Row, .ColIndex("��ֹ¼��")) = ""
                .TextMatrix(.Row, .ColIndex("������")) = "��"
                .Cell(flexcpBackColor, .Row, .ColIndex("��ֹ¼��")) = &H8000000F
                .Cell(flexcpBackColor, .Row, .ColIndex("������")) = &H8000000E
            ElseIf .TextMatrix(.Row, .ColIndex("������")) = "��" Then
                .Cell(flexcpBackColor, .Row, .ColIndex("��ֹ¼��")) = &H8000000F
                .Cell(flexcpBackColor, .Row, .ColIndex("������")) = &H8000000E
            Else
                .Cell(flexcpBackColor, .Row, .ColIndex("��ֹ¼��"), .Row, .ColIndex("������")) = &H8000000E
            End If
             .Cell(flexcpBackColor, .Row, .ColIndex("������")) = &H8000000E
        Case .ColIndex("������")
            .TextMatrix(.Row, .ColIndex("������")) = IIF(.TextMatrix(.Row, .ColIndex("������")) = "", "��", "")
             .Cell(flexcpBackColor, .Row, .ColIndex("������")) = &H8000000E
            If .TextMatrix(.Row, .ColIndex("������")) = "��" Then
                .TextMatrix(.Row, .ColIndex("��ֹ¼��")) = ""
                
                .Cell(flexcpBackColor, .Row, .ColIndex("��ֹ¼��")) = &H8000000F
            ElseIf .TextMatrix(.Row, .ColIndex("������")) = "��" Then
                .TextMatrix(.Row, .ColIndex("��ֹ¼��")) = ""
                .Cell(flexcpBackColor, .Row, .ColIndex("��ֹ¼��")) = &H8000000F
            Else
                .Cell(flexcpBackColor, .Row, .ColIndex("��ֹ¼��"), .Row, .ColIndex("������")) = &H8000000E
            End If
        End Select
    End With
    Call SetParChange(vsgInput, intIndex, mrsPar, True, GetInputItemSetValue(intIndex))
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Public Function GetInputItemSetValue(ByVal intIndex As Integer) As String
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�����������ֵ
    '���:intIndex-�ؼ�����
    '����:���������õ�ֵ,��ʽ:������,�Ƿ����,����Ƿ�����,�Ƿ������|....
    '����:
    '����:2015-06-11 18:10:12
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Integer, strTmp As String
    On Error GoTo ErrHandle
        
    With vsgInput(intIndex)
        For i = 1 To .Rows - 1
            strTmp = strTmp & "|" & .TextMatrix(i, .ColIndex("������Ŀ"))
            strTmp = strTmp & "," & IIF(.TextMatrix(i, .ColIndex("��ֹ¼��")) = "��", 1, 0)
            strTmp = strTmp & "," & IIF(.TextMatrix(i, .ColIndex("������")) = "��", 1, 0)
            strTmp = strTmp & "," & IIF(.TextMatrix(i, .ColIndex("������")) = "��", 1, 0)
        Next
    End With
    GetInputItemSetValue = Mid(strTmp, 2)
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub LoadRisEnables()
'-----------------------------------------------------------
'����:����RIS���ÿ����б�
'���:
'����:��..
'-----------------------------------------------------------
   
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    Dim strModality As String
    Dim lngSource As Long
    Dim j As Integer
    Dim strDeptIDs As String
    Dim strDeptNames As String
    
    On Error GoTo Err
    
    '��ѯ���еļ������
    
    strSQL = "select ����, ���� from Ӱ�������"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡӰ�������")
    
    If rsTemp.EOF = True Then Exit Sub
    
    With vsfRISEnables
        .Rows = 1 + rsTemp.RecordCount * 3
        .Cols = 7
        .FixedRows = 1
        .FixedCols = 0
        .RowHeightMin = 400
        .AllowUserResizing = flexResizeColumns
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExNone
        .ExtendLastCol = True

        .TextMatrix(0, col_RIS���ü������) = "�������"
        .TextMatrix(0, col_RIS���ó���) = "����"
        .TextMatrix(0, col_RIS���ÿ���) = "����"
        .TextMatrix(0, col_RIS����ԤԼ����ȫ) = "����ԤԼ����"
        .TextMatrix(0, col_RIS����ԤԼ����) = "����ԤԼ����"

        .ColWidth(col_RIS���ü������) = 850
        .ColWidth(col_RIS���ó���) = 650
        .ColWidth(col_RIS���ÿ���) = 2000
        .ColWidth(col_RIS����ԤԼ����ȫ) = 650
        .ColWidth(col_RIS����ԤԼ����) = 1900
        
        .Cell(flexcpChecked, 1, col_RIS���ü������, .Rows - 1, col_RIS���ó���) = 2
        .Cell(flexcpChecked, 1, col_RIS����ԤԼ����ȫ, .Rows - 1, col_RIS����ԤԼ����ȫ) = 2
        
        For i = 0 To rsTemp.RecordCount - 1
            .TextMatrix(i * 3 + 1, col_RIS���ü������) = rsTemp!����
            .TextMatrix(i * 3 + 2, col_RIS���ü������) = rsTemp!����
            .TextMatrix(i * 3 + 3, col_RIS���ü������) = rsTemp!����
            .RowData(i * 3 + 1) = NVL(rsTemp!����)
            .RowData(i * 3 + 2) = NVL(rsTemp!����)
            .RowData(i * 3 + 3) = NVL(rsTemp!����)
            
            .TextMatrix(i * 3 + 1, col_RIS���ó���) = "����"
            .TextMatrix(i * 3 + 2, col_RIS���ó���) = "סԺ"
            .TextMatrix(i * 3 + 3, col_RIS���ó���) = "���"
            
            .TextMatrix(i * 3 + 1, col_RIS����ԤԼ����ȫ) = "ȫ��"
            .TextMatrix(i * 3 + 2, col_RIS����ԤԼ����ȫ) = "ȫ��"
            .TextMatrix(i * 3 + 3, col_RIS����ԤԼ����ȫ) = "ȫ��"
            rsTemp.MoveNext
        Next i
        
        '���õ�Ԫ��ϲ�
        .MergeCellsFixed = flexMergeFree
        .MergeCol(0) = True
        .MergeRow(0) = True
        
        
        '���ؿ���ID��
        .ColHidden(col_RIS���ÿ���ID) = True
        .ColHidden(col_RIS����ԤԼ����ID) = True
        
        '��ȡRIS�������Ʋ���������ʾ
        strSQL = "select a.�������,a.����,a.����ID,a.�Ƿ�����RIS,a.�Ƿ�����ԤԼ,b.���� from ris���ÿ��� a, ���ű� b where a.����id = b.id(+)"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡRIS���ÿ���")
        If rsTemp.EOF = True Then Exit Sub
        
        'ѭ���б���д����,����һ��ѭ��
        For i = 1 To .Rows - 1 Step 3
            Call loadOneModality(vsfRISEnables, rsTemp, i)
        Next i
        
        '���µ����еĸ߶ȣ�ȷ��������������ʾ
        Call .AutoSize(col_RIS���ÿ���, col_RIS����ԤԼ����)
        
        .Refresh
        
    End With
    Exit Sub
Err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub LoadRisDepts(strDeptIDs As String, lngSource As Long)
'-----------------------------------------------------------
'����:����RIS���ÿ����зֿ��ҿ��Ƶ��б�
'���:  strDeptIDs -- ����ID��
'       lngSource -- ������Դ 1 - ���2 - סԺ�� 4 - ���
'����:��..
'-----------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo Err
    
    vsfRisDepts.Clear
    
    strSQL = "Select Distinct D.ID,D.����, D.����,D.����,t.������� From ���ű� D, ��������˵�� T " & _
            " Where d.Id = t.����id And (d.����ʱ�� Is Null Or d.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD')) "
    If lngSource = 1 Then   '����
        strSQL = strSQL & " And t.������� IN (1,3)  and T.�������� IN ('�ٴ�','����','����','����','���','����','Ӫ��') order by ���� "
    Else    'סԺ
        strSQL = strSQL & " And t.������� IN (2,3)  and T.�������� IN ('�ٴ�','����','����','����','���','����','Ӫ��') order by ���� "
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ����")
    If rsTemp.EOF = True Then Exit Sub
    
    strDeptIDs = "," & strDeptIDs & ","
    
    With vsfRisDepts
        .Rows = rsTemp.RecordCount + 1
        .Cols = 4
        .FixedRows = 1
        .FixedCols = 0
        .RowHeightMin = 400
        .Editable = flexEDKbdMouse
        .AllowUserResizing = flexResizeBoth
        .SelectionMode = flexSelectionByRow
        .ExplorerBar = flexExSort
        .ColSort(col_Ris����ѡ��) = flexSortNone
        .ExtendLastCol = True
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .Cell(flexcpPictureAlignment, 0, col_Ris����ѡ��, .Rows - 1, col_Ris����ѡ��) = flexAlignCenterCenter
        .Cell(flexcpAlignment, 1, col_Ris���ұ���, .Rows - 1, col_Ris���ұ���) = flexAlignLeftCenter
        
        .Cell(flexcpPicture, 0, col_Ris����ѡ��) = imgCheck.ListImages("UnChecked").Picture

        .TextMatrix(0, col_Ris����ѡ��) = ""
        .TextMatrix(0, col_Ris���ұ���) = "����"
        .TextMatrix(0, col_Ris��������) = "����"
        
        .ColWidth(col_Ris����ѡ��) = 400
        .ColWidth(col_Ris���ұ���) = 850
        .ColWidth(col_Ris��������) = 1200
        
        .Cell(flexcpChecked, 1, col_Ris����ѡ��, .Rows - 1, col_Ris����ѡ��) = 2
        
        For i = 1 To rsTemp.RecordCount
            .TextMatrix(i, col_Ris���ұ���) = rsTemp!����
            .TextMatrix(i, col_Ris��������) = rsTemp!����
            .TextMatrix(i, col_Ris����ID) = rsTemp!ID
            If InStr(strDeptIDs, rsTemp!ID) > 0 Then
                .Cell(flexcpChecked, i, col_Ris����ѡ��) = 1
            End If
            rsTemp.MoveNext
        Next i
        
        .ColHidden(col_Ris����ID) = True
        
        .Refresh
    End With
    
    Exit Sub
Err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function loadOneModality(vsfGrid As VSFlexGrid, rsData As ADODB.Recordset, iRow As Integer) As Boolean
'-----------------------------------------------------------
'����:����һ�������������
'���:  vsfGrid -- vsflexGrid�ؼ�
'       rsData -- ����Դ
'       iRow -- Ҫ���ص��к�
'����:True -- �ɹ��� False -- ʧ��
'-----------------------------------------------------------
    Dim lngSource As Long
    Dim strModality As String
    Dim strDeptIDs As String
    Dim strDeptNames As String
    Dim i As Integer
    
    On Error GoTo Err
    
    With vsfGrid
        '�ж�����RIS�ͼ�������Ƿ�ѡ��
        strModality = .RowData(iRow)
        rsData.Filter = " �������='" & strModality & "' and �Ƿ�����RIS=1"
        If rsData.EOF = False Then
            '�ȹ�ѡ������ͣ�������ж����סԺ������ѡ�����
            .Cell(flexcpChecked, iRow, col_RIS���ü������) = 1
            .Cell(flexcpChecked, iRow + 1, col_RIS���ü������) = 1
            .Cell(flexcpChecked, iRow + 2, col_RIS���ü������) = 1
            
            For i = 0 To 2
                lngSource = IIF(.TextMatrix(iRow + i, col_RIS���ó���) = "����", 1, IIF(.TextMatrix(iRow + i, col_RIS���ó���) = "סԺ", 2, 4))
                If GetDeptString(rsData, lngSource, strDeptNames, strDeptIDs) = False Then
                    '����û�б�ѡ�У����ô���
                Else
                    '�ȹ�ѡ���ϣ����ж��Ƿ��п���
                    .Cell(flexcpChecked, iRow + i, col_RIS���ó���) = 1
                    If strDeptIDs = "" Or lngSource = 4 Then
                        'û��ѡ����ң����գ����ô���
                        '�����Ϊһ����һ���ң�ֻ���ֳ��ϣ������ֿ���
                    Else
                        'ѡ���˿��ң�����д����
                        .TextMatrix(iRow + i, col_RIS���ÿ���) = strDeptNames
                        .TextMatrix(iRow + i, col_RIS���ÿ���ID) = strDeptIDs
                    End If
                End If
            Next i
            
            '����ж��Ƿ�ѡ�������סԺ���������֮һ,���û��ѡ����ȡ��������͵Ĺ�ѡ��
            If .Cell(flexcpChecked, iRow, col_RIS���ó���) = 2 And .Cell(flexcpChecked, iRow + 1, col_RIS���ó���) = 2 And .Cell(flexcpChecked, iRow + 2, col_RIS���ó���) = 2 Then
                .Cell(flexcpChecked, iRow, col_RIS���ü������) = 2
                .Cell(flexcpChecked, iRow + 1, col_RIS���ü������) = 2
                .Cell(flexcpChecked, iRow + 2, col_RIS���ü������) = 2
                '�����ٴ���ԤԼ���ݣ�ֱ���˳�
                Exit Function
            End If
            
            '�ж��Ƿ�������ԤԼ
            rsData.Filter = " �������='" & strModality & "' and �Ƿ�����ԤԼ=1"
            If rsData.EOF = False Then
                '����ж����סԺ������ѡ�����
                For i = 0 To 2
                    lngSource = IIF(.TextMatrix(iRow + i, col_RIS���ó���) = "����", 1, IIF(.TextMatrix(iRow + i, col_RIS���ó���) = "סԺ", 2, 4))
                    '�ó���������RIS������дԤԼ
                    If .Cell(flexcpChecked, iRow + i, col_RIS���ó���) = 1 Then
                        If GetDeptString(rsData, lngSource, strDeptNames, strDeptIDs) = True Then
                            If strDeptIDs = "" Or lngSource = 4 Then
                                '������������ԤԼ,���������
                                .Cell(flexcpChecked, iRow + i, col_RIS����ԤԼ����ȫ) = 1
                            Else
                                '������������ԤԼ
                                .TextMatrix(iRow + i, col_RIS����ԤԼ����) = strDeptNames
                                .TextMatrix(iRow + i, col_RIS����ԤԼ����ID) = strDeptIDs
                            End If
                        End If
                    End If
                Next i
            End If
        End If
        
    End With
    
    loadOneModality = True
    Exit Function
Err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function GetDeptString(rsData As ADODB.Recordset, lngSource As Long, ByRef strDeptNames As String, ByRef strDeptIDs As String) As Boolean
'-----------------------------------------------------------
'����:��ȡ����ID�Ͳ������ƴ�
'���:  rsData -- ����Դ
'       lngSource -- ������Դ��1=���2=סԺ��4=���
'       strDeptNames -- ����ֵ���������ƴ�
'       strDeptIDs -- ����ֵ������ID��
'����:True -- �ɹ��� False -- ʧ��
'-----------------------------------------------------------
    Dim i As Integer
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim strFilter As String
    
    On Error GoTo Err
    
    strDeptNames = ""
    strDeptIDs = ""
    
    strFilter = rsData.Filter
    
    rsData.Filter = strFilter & " and ����=" & lngSource
    
    If rsData.EOF = True Then
        rsData.Filter = strFilter
        GetDeptString = False
        Exit Function
    Else
        If rsData.RecordCount = 1 And IsNull(rsData!����ID) Then
            'ʹ��Ĭ�Ϸ���ֵ
        Else
            '��ϲ���ID�����ƴ�
            For i = 1 To rsData.RecordCount
                strDeptIDs = strDeptIDs & "," & NVL(rsData!����ID)
                strDeptNames = strDeptNames & "," & NVL(rsData!����)
                rsData.MoveNext
            Next i
            
            strDeptIDs = Mid(strDeptIDs, 2)
            strDeptNames = Mid(strDeptNames, 2)
        End If
    End If
    
    '�ָ�ԭ����Filter�������Ͳ���Ҫ����һ�����ݼ���
    rsData.Filter = strFilter
    GetDeptString = True
    Exit Function
Err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub SaveRisEnable()
'-----------------------------------------------------------
'����:����RIS�ֿ�����������,���桰Ӱ����Ϣϵͳ�����ÿ���
'���:
'����:
'-----------------------------------------------------------

    Dim i As Integer
    Dim strRISDeptIDs As String
    Dim strSchDeptIDs As String
    Dim strModality As String
    Dim lngSource As Long
    Dim strSQL As String
    
    On Error GoTo Err
    
    With vsfRISEnables
        If .Tag = "���޸�" Then
            strSQL = "b_Zlxwinterface.RIS���ÿ���_Delete()"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
            For i = 1 To .Rows - 1
                'ѡ���˵ļ�����ͣ��ű���
                If .Cell(flexcpChecked, i, col_RIS���ü������) = 1 Then
                    strModality = .RowData(i)
                    'ѡ���˳��ϣ��ű���
                    If .Cell(flexcpChecked, i, col_RIS���ó���) = 1 Then
                        strRISDeptIDs = .TextMatrix(i, col_RIS���ÿ���ID)
                        lngSource = IIF(.TextMatrix(i, col_RIS���ó���) = "����", 1, IIF(.TextMatrix(i, col_RIS���ó���) = "סԺ", 2, 4))
                        
                        strSQL = "b_Zlxwinterface.RIS���ÿ���_Update('" & strModality & "'," & lngSource & ",'" & strRISDeptIDs & "',1)"
                        Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                        
                        'ѡ����ԤԼ���ű���
                        strSchDeptIDs = .TextMatrix(i, col_RIS����ԤԼ����ID)
                        If strSchDeptIDs <> "" Or .Cell(flexcpChecked, i, col_RIS����ԤԼ����ȫ) = 1 Then
                            strSQL = "b_Zlxwinterface.RIS���ÿ���_Update('" & strModality & "'," & lngSource & ",'" & strSchDeptIDs & "',2)"
                            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
                        End If
                    End If
                End If
            Next i
            .Tag = ""
        End If
    End With
    
    Exit Sub
Err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub WriteDeptsIntoVsfRisEnables()
'-----------------------------------------------------------
'����:����������д�ص���RIS���ÿ����б�
'���:
'����:
'-----------------------------------------------------------
    Dim i As Integer
    Dim strDeptNames As String
    Dim strDeptIDs As String
    Dim iSelCount As Integer
    
    On Error GoTo Err
    
    iSelCount = 0
    If vsfRISEnables.ColSel = col_RIS���ÿ��� Or vsfRISEnables.ColSel = col_RIS����ԤԼ���� Then
        With vsfRisDepts
            For i = 1 To .Rows - 1
                If .Cell(flexcpChecked, i, col_Ris����ѡ��) = 1 Then
                    strDeptIDs = strDeptIDs & "," & .TextMatrix(i, col_Ris����ID)
                    strDeptNames = strDeptNames & "," & .TextMatrix(i, col_Ris��������)
                    iSelCount = iSelCount + 1
                End If
            Next i
            
            '�ж��Ƿ�ȫѡ
            If iSelCount = vsfRisDepts.Rows - 1 Then
                'ȫѡ
                If vsfRISEnables.ColSel = col_RIS���ÿ��� Then
                    vsfRISEnables.TextMatrix(vsfRISEnables.RowSel, col_RIS���ÿ���ID) = ""
                    vsfRISEnables.TextMatrix(vsfRISEnables.RowSel, col_RIS���ÿ���) = ""
                Else
                    vsfRISEnables.Cell(flexcpChecked, vsfRISEnables.RowSel, col_RIS����ԤԼ����ȫ) = 1
                    vsfRISEnables.TextMatrix(vsfRISEnables.RowSel, col_RIS����ԤԼ����ID) = ""
                    vsfRISEnables.TextMatrix(vsfRISEnables.RowSel, col_RIS����ԤԼ����) = ""
                End If
            Else
                '����ѡ��
                If strDeptIDs <> "" Then
                    strDeptIDs = Mid(strDeptIDs, 2)
                    strDeptNames = Mid(strDeptNames, 2)
                End If
                
                If vsfRISEnables.ColSel = col_RIS���ÿ��� Then
                    vsfRISEnables.TextMatrix(vsfRISEnables.RowSel, col_RIS���ÿ���ID) = strDeptIDs
                    vsfRISEnables.TextMatrix(vsfRISEnables.RowSel, col_RIS���ÿ���) = strDeptNames
                Else
                    '�����ԤԼ���ң���ȡ��ԤԼȫѡ
                    vsfRISEnables.Cell(flexcpChecked, vsfRISEnables.RowSel, col_RIS����ԤԼ����ȫ) = 2
                    vsfRISEnables.TextMatrix(vsfRISEnables.RowSel, col_RIS����ԤԼ����ID) = strDeptIDs
                    vsfRISEnables.TextMatrix(vsfRISEnables.RowSel, col_RIS����ԤԼ����) = strDeptNames
                End If
            End If
            Call vsfRISEnables.AutoSize(col_RIS���ÿ���, col_RIS����ԤԼ����)
        End With
    End If
    Exit Sub
Err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadRisBranchHosp()
'-----------------------------------------------------------
'����:����RIS�� ��Ժ����
'���:  ��
'����:��
'-----------------------------------------------------------
    Dim strSQL As String
    Dim rsTemp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo Err
    
    vsfBranchHosp.Clear
    
    strSQL = "select a.id ,a.ҽԺ����,a.ҽԺ����,a.�û���,a.����,a.���ݿ������ from ris��Ժ���� a"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, "��ȡRIS��Ժ����")
    
    With vsfBranchHosp
        .Rows = IIF(rsTemp.RecordCount = 0, 1, rsTemp.RecordCount)
        .Cols = 6
        .FixedRows = 1
        .FixedCols = 1
        .RowHeightMin = 400
        .AllowUserResizing = flexResizeColumns
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExNone
        .ExtendLastCol = True
        
        .Cell(flexcpAlignment, 0, 0, 0, col_ris��Ժ���ݿ������) = flexAlignCenterCenter
        
        .TextMatrix(0, col_RIS��Ժ���) = "���"
        .TextMatrix(0, col_RIS��Ժ����) = "ҽԺ����"
        .TextMatrix(0, col_ris��Ժ����) = "ҽԺ����"
        .TextMatrix(0, col_ris��Ժ�û���) = "�û���"
        .TextMatrix(0, col_ris��Ժ����) = "����"
        .TextMatrix(0, col_ris��Ժ���ݿ������) = "���ݿ������"
        
        .ColWidth(col_RIS��Ժ���) = 600
        .ColWidth(col_RIS��Ժ����) = 1600
        .ColWidth(col_ris��Ժ����) = 1600
        .ColWidth(col_ris��Ժ�û���) = 1600
        .ColWidth(col_ris��Ժ����) = 1600
        .ColWidth(col_ris��Ժ���ݿ������) = 1600
        
        i = 1
        While Not rsTemp.EOF
            '��Ժ������д
            If rsTemp!ҽԺ���� = "��Ժ" Then
                txtMainHosp.Text = rsTemp!ҽԺ����
            Else
                .TextMatrix(i, col_RIS��Ժ���) = i
                .TextMatrix(i, col_RIS��Ժ����) = rsTemp!ҽԺ����
                .TextMatrix(i, col_ris��Ժ����) = rsTemp!ҽԺ����
                .TextMatrix(i, col_ris��Ժ�û���) = NVL(rsTemp!�û���)
                .TextMatrix(i, col_ris��Ժ����) = NVL(rsTemp!����)
                .TextMatrix(i, col_ris��Ժ���ݿ������) = rsTemp!���ݿ������
                i = i + 1
            End If
            
            rsTemp.MoveNext
        Wend
        .Refresh
    End With
    
    Exit Sub
Err:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function ValidateRisBranchHosp() As Boolean
'-----------------------------------------------------------
'����:���RIS��Ժ��Ϣ����Ч��
'���:��
'����: True -- ������Ч�����Ա��棻False -- ������Ч�����ܱ���
'-----------------------------------------------------------
    Dim i As Integer
    
    On Error GoTo Err
    
    If txtMainHosp.Tag = "���޸�" Or vsfBranchHosp.Tag = "���޸�" Then
        '�ȼ�����ݵ������ԣ�
        '��1��û�������κ���Ϣ
        '��2�������ɱ�Ժ���룬��������һ����Ժ���ݣ�
        '��3����Ժ�������û������������ͬʱΪ�գ��������ݷǿա�
        
        With vsfBranchHosp
            '�Ȱ�vsfBranchHosp�ж������ɾ����
            Do While .Rows > 1
                If (.TextMatrix(.Rows - 1, col_RIS��Ժ����) = "" And .TextMatrix(.Rows - 1, col_ris��Ժ����) = "" _
                    And .TextMatrix(.Rows - 1, col_ris��Ժ�û���) = "" And .TextMatrix(.Rows - 1, col_ris��Ժ�û���) = "" _
                    And .TextMatrix(.Rows - 1, col_ris��Ժ���ݿ������) = "") Then
                    .Rows = .Rows - 1
                Else
                    Exit Do
                End If
            Loop
        
            If vsfBranchHosp.Rows > 1 Or txtMainHosp.Text <> "" Then
                If txtMainHosp.Text = "" Then
                    MsgBox "��Ӱ����Ϣϵͳ---HISҽԺ����---��Ժ���á��У�ҽԺ���벻��Ϊ�ա�", vbInformation, gstrSysName
                    txtMainHosp.SetFocus
                    Exit Function
                End If
                
                If vsfBranchHosp.Rows <= 1 Then
                    MsgBox "��Ӱ����Ϣϵͳ---HISҽԺ����---��Ժ���á��У�����Ӧ������һ����Ժ��Ϣ��", vbInformation, gstrSysName
                    Exit Function
                End If
                
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, col_RIS��Ժ����) = "" Then
                        MsgBox "��Ӱ����Ϣϵͳ---HISҽԺ����---��Ժ���á��У�" & vbCrLf & vbCrLf & " ҽԺ���Ʋ���Ϊ�գ�����дҽԺ���ơ�", vbInformation, gstrSysName
                        Exit Function
                    End If
                    
                    If .TextMatrix(i, col_ris��Ժ����) = "" Then
                        MsgBox "��Ӱ����Ϣϵͳ---HISҽԺ����---��Ժ���á��У�" & vbCrLf & vbCrLf & " ҽԺ���벻��Ϊ�գ�����дҽԺ���롣", vbInformation, gstrSysName
                        Exit Function
                    End If
                    
                    If .TextMatrix(i, col_ris��Ժ���ݿ������) = "" Then
                        MsgBox "��Ӱ����Ϣϵͳ---HISҽԺ����---��Ժ���á��У�" & vbCrLf & vbCrLf & " ���ݿ����������Ϊ�գ�����д���ݿ��������", vbInformation, gstrSysName
                        Exit Function
                    End If
                    
                    If (.TextMatrix(i, col_ris��Ժ�û���) = "" And .TextMatrix(i, col_ris��Ժ����) = "") Or (.TextMatrix(i, col_ris��Ժ�û���) <> "" And .TextMatrix(i, col_ris��Ժ����) <> "") Then
                        '��ȷ����������ô���
                    Else
                        MsgBox "��Ӱ����Ϣϵͳ---HISҽԺ����---��Ժ���á��У�" & vbCrLf & vbCrLf & " �û������������ͬʱΪ�գ�����ͬʱ��Ϊ�գ��밴�պ�ɫ���������Ĺ�����д�û��������롣", vbInformation, gstrSysName
                        Exit Function
                    End If
                Next i
            End If
        End With
    End If
    
    ValidateRisBranchHosp = True
    Exit Function
Err:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub SaveRisBranchHosp()
'-----------------------------------------------------------
'����:����RIS�ķ�Ժ��Ϣ
'���:��
'����:��
'-----------------------------------------------------------
    Dim blnInTrans As Boolean       '�Ƿ���������֮��
    Dim arrSQL() As Variant
    Dim strSQL As String
    Dim i As Integer

    On Error GoTo Err
    
    If txtMainHosp.Tag = "���޸�" Or vsfBranchHosp.Tag = "���޸�" Then
        arrSQL = Array()
        
        '����գ����������
        strSQL = "b_Zlxwinterface.Ris��Ժ����_Delete()"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSQL
        
        '��ӱ�Ժ����
        strSQL = "b_Zlxwinterface.Ris��Ժ����_Update(1,'��Ժ','" & txtMainHosp.Text & "',null,null,null)"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = strSQL
                
        '��ӷ�Ժ����
        With vsfBranchHosp
            For i = 1 To .Rows - 1
            
                strSQL = "b_Zlxwinterface.Ris��Ժ����_Update(" & Val(.TextMatrix(i, col_RIS��Ժ���)) + 1 _
                     & ",'" & .TextMatrix(i, col_RIS��Ժ����) & "','" & .TextMatrix(i, col_ris��Ժ����) _
                    & "','" & .TextMatrix(i, col_ris��Ժ�û���) & "','" & .TextMatrix(i, col_ris��Ժ����) _
                    & "','" & .TextMatrix(i, col_ris��Ժ���ݿ������) & "')"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSQL
            Next i
        End With
        
        gcnOracle.BeginTrans        '��ʼ�������
        blnInTrans = True
        For i = 0 To UBound(arrSQL)
            Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "����Ris��Ժ����")
        Next i
        gcnOracle.CommitTrans
        blnInTrans = False
    
        '�������֮�����ó�δ�޸�
        txtMainHosp.Tag = ""
        vsfBranchHosp.Tag = ""
    End If
    
    Exit Sub
Err:
    If blnInTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub LoadThirdSvr()
'���ܣ���ʼ�� ������������Ŀ¼
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim i As Long
    
    On Error GoTo errH
    
    strTmp = ",200,1;ϵͳ��ʶ,1000,1;��������,2100,1;�����ַ,3000,1"
    Call Grid.Init(vsThirdSvr, strTmp, , 1)
    vsThirdSvr.Rows = 1
    Set mrsSvr = Nothing
    strSQL = "Select a.ϵͳ��ʶ, a.��������, a.�����ַ From ������������Ŀ¼ A Order By a.ϵͳ��ʶ,a.��������"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ������������Ŀ¼")
    If rsTmp.EOF Then Exit Sub
    Set mrsSvr = zlDatabase.CopyNewRec(rsTmp)
    With vsThirdSvr
        .AllowUserResizing = flexResizeColumns
        .Editable = flexEDKbdMouse
        .ExplorerBar = flexExNone
        .ExtendLastCol = True
        .Rows = rsTmp.RecordCount + 1
        For i = 1 To rsTmp.RecordCount
            .TextMatrix(i, COL_ϵͳ��ʶ) = rsTmp!ϵͳ��ʶ & ""
            .TextMatrix(i, COL_��������) = rsTmp!�������� & ""
            .TextMatrix(i, COL_�����ַ) = rsTmp!�����ַ & ""
            rsTmp.MoveNext
        Next
    End With
      
    Exit Sub
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub vsThirdSvr_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> COL_�����ַ Then
        Cancel = True
    End If
End Sub

Private Function CheckThirdSvr() As Boolean
'���ܣ���� ������������Ŀ¼
'������blnSave true���������ʱ�����
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim i As Long
    Dim strMsg As String
    Dim strTmp1 As String
    
    On Error GoTo errH
    If mrsSvr Is Nothing Then Exit Function
    
    With vsThirdSvr
        mrsSvr.MoveFirst
        For i = 1 To mrsSvr.RecordCount
            .TextMatrix(i, COL_�����ַ) = Trim(.TextMatrix(i, COL_�����ַ))
            If mrsSvr!�����ַ & "" <> .TextMatrix(i, COL_�����ַ) And .TextMatrix(i, COL_�����ַ) <> "" Then
                '��ַ�����˱仯��Ҫ���
                strTmp = ""
                Call Sys.WebAPIByBasic(.TextMatrix(i, COL_�����ַ), "", strTmp1, strTmp)
                If strTmp <> "" Then
                    strMsg = IIF("" = strMsg, "", strMsg & vbCrLf) & .TextMatrix(i, COL_ϵͳ��ʶ) & ":" & .TextMatrix(i, COL_��������) & "  ��֤��" & strTmp
                End If
            End If
            mrsSvr.MoveNext
        Next
    End With
    If strMsg <> "" Then
        MsgBox strMsg, vbInformation, Me.Caption
    Else
        CheckThirdSvr = True
    End If
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub ThirdSvrTest(ByVal lngRow As Long)
'���ܣ����в��Է���Ϸ���
    Dim strTmp As String
    Dim strTmp1 As String
    Dim strMsg As String
    
    With vsThirdSvr
        If lngRow < 1 Then Exit Sub
        .TextMatrix(lngRow, COL_�����ַ) = Trim(.TextMatrix(lngRow, COL_�����ַ))
        If "" = .TextMatrix(lngRow, COL_�����ַ) Then
            MsgBox "�����ַΪ�գ�����д��", vbInformation, Me.Caption
            Exit Sub
        End If
        Call Sys.WebAPIByBasic(.TextMatrix(lngRow, COL_�����ַ), "", strTmp1, strTmp)
        If strTmp <> "" Then
            strMsg = .TextMatrix(lngRow, COL_ϵͳ��ʶ) & ":" & .TextMatrix(lngRow, COL_��������) & "  ��֤��" & strTmp
        End If
    End With
    If strMsg <> "" Then
        MsgBox strMsg, vbInformation, Me.Caption
    Else
        MsgBox "�ɹ���", vbInformation, Me.Caption
    End If
End Sub

Private Function ThirdSvrChanged() As Boolean
'���ܣ��ж����������ַ�Ƿ����仯
    Dim i As Long
    If mrsSvr Is Nothing Then Exit Function
    
    With vsThirdSvr
        mrsSvr.MoveFirst
        For i = 1 To mrsSvr.RecordCount
            .TextMatrix(i, COL_�����ַ) = Trim(.TextMatrix(i, COL_�����ַ))
            If mrsSvr!�����ַ & "" <> .TextMatrix(i, COL_�����ַ) Then
                '��ַ�����˱仯��Ҫ���
                ThirdSvrChanged = True
                Exit Function
            End If
            mrsSvr.MoveNext
        Next
    End With
End Function

Private Sub SaveThirdSvr()
'���ܣ����� ������������Ŀ¼
    Dim blnInTrans As Boolean
    Dim i As Long
    Dim arrSQL As Variant
    Dim strSQL As String
    
    On Error GoTo errH
    
    If Not ThirdSvrChanged Then Exit Sub
    If Not CheckThirdSvr Then Exit Sub
    If mrsSvr Is Nothing Then Exit Sub
    
    arrSQL = Array()
    With vsThirdSvr
        mrsSvr.MoveFirst
        For i = 1 To mrsSvr.RecordCount
            If mrsSvr!�����ַ & "" <> .TextMatrix(i, COL_�����ַ) Then
                '��ַ�����˱仯�ı���
                strSQL = "Zl_������������Ŀ¼_Update('" & .TextMatrix(i, 1) & "','" & .TextMatrix(i, 2) & "','" & .TextMatrix(i, 3) & "')"
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = strSQL
            End If
            mrsSvr.MoveNext
        Next
    End With
    
    gcnOracle.BeginTrans        '��ʼ�������
    blnInTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), "����������������Ŀ¼")
    Next i
    gcnOracle.CommitTrans
    blnInTrans = False
    Exit Sub
errH:
    If blnInTrans = True Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdSvrChk_Click()
'���ܣ���֤ ��������
    Call ThirdSvrTest(vsThirdSvr.Row)
End Sub
