VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{C932BA88-4374-101B-A56C-00AA003668DC}#1.1#0"; "MSMASK32.OCX"
Object = "{D01C2596-4FE0-4EA9-9EE8-D97BE62A1165}#1.1#0"; "ZlPatiAddress.ocx"
Begin VB.Form frmInMedRecEdit_YN 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "��ҳ����"
   ClientHeight    =   7620
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11040
   Icon            =   "frmInMedRecEdit_YN.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   7620
   ScaleWidth      =   11040
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdPrintdown 
      Caption         =   "��"
      Height          =   350
      Left            =   2370
      TabIndex        =   383
      TabStop         =   0   'False
      ToolTipText     =   "ѡ��(*)"
      Top             =   7155
      Width           =   270
   End
   Begin VB.CommandButton cmdPriviewDown 
      Caption         =   "��"
      Height          =   350
      Left            =   1080
      TabIndex        =   382
      TabStop         =   0   'False
      ToolTipText     =   "ѡ��(*)"
      Top             =   7155
      Width           =   270
   End
   Begin VB.CommandButton cmdPriview 
      Caption         =   "Ԥ��"
      Height          =   350
      Left            =   240
      TabIndex        =   381
      Top             =   7155
      Width           =   855
   End
   Begin VB.CommandButton cmdPrint 
      Caption         =   "��ӡ"
      Height          =   350
      Left            =   1485
      TabIndex        =   380
      Top             =   7155
      Width           =   900
   End
   Begin TabDlg.SSTab sstInfo 
      Height          =   6975
      Left            =   120
      TabIndex        =   328
      Top             =   120
      Width           =   10815
      _ExtentX        =   19076
      _ExtentY        =   12303
      _Version        =   393216
      Style           =   1
      Tabs            =   8
      Tab             =   3
      TabsPerRow      =   8
      TabHeight       =   520
      TabCaption(0)   =   "������Ϣ"
      TabPicture(0)   =   "frmInMedRecEdit_YN.frx":058A
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "fraInfo(0)"
      Tab(0).ControlCount=   1
      TabCaption(1)   =   "��ҽ���"
      TabPicture(1)   =   "frmInMedRecEdit_YN.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraInfo(1)"
      Tab(1).Control(1)=   "cmdInfo(35)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).ControlCount=   2
      TabCaption(2)   =   "��ҽ���"
      TabPicture(2)   =   "frmInMedRecEdit_YN.frx":05C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "fraInfo(2)"
      Tab(2).ControlCount=   1
      TabCaption(3)   =   "����������"
      TabPicture(3)   =   "frmInMedRecEdit_YN.frx":05DE
      Tab(3).ControlEnabled=   -1  'True
      Tab(3).Control(0)=   "fraInfo(3)"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "סԺ���"
      TabPicture(4)   =   "frmInMedRecEdit_YN.frx":05FA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "fraInfo(4)"
      Tab(4).ControlCount=   1
      TabCaption(5)   =   "�����뻯��"
      TabPicture(5)   =   "frmInMedRecEdit_YN.frx":0616
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "fraInfo(5)"
      Tab(5).ControlCount=   1
      TabCaption(6)   =   "��ҳ1"
      TabPicture(6)   =   "frmInMedRecEdit_YN.frx":0632
      Tab(6).ControlEnabled=   0   'False
      Tab(6).Control(0)=   "fraInfo(6)"
      Tab(6).ControlCount=   1
      TabCaption(7)   =   "��ҳ2"
      TabPicture(7)   =   "frmInMedRecEdit_YN.frx":064E
      Tab(7).ControlEnabled=   0   'False
      Tab(7).Control(0)=   "fraInfo(7)"
      Tab(7).ControlCount=   1
      Begin VB.Frame fraInfo 
         BorderStyle     =   0  'None
         Height          =   6495
         Index           =   7
         Left            =   -74880
         TabIndex        =   349
         Top             =   360
         Width           =   10455
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   65
            ItemData        =   "frmInMedRecEdit_YN.frx":066A
            Left            =   1920
            List            =   "frmInMedRecEdit_YN.frx":066C
            Style           =   2  'Dropdown List
            TabIndex        =   311
            Top             =   1200
            Width           =   2445
         End
         Begin VB.CheckBox chkInfo 
            Caption         =   "��Ԥ�ڵ��ط���֢ҽѧ��"
            Height          =   195
            Index           =   25
            Left            =   4440
            TabIndex        =   310
            Top             =   840
            Width           =   2850
         End
         Begin VB.CheckBox chkInfo 
            Caption         =   "�����˹������ѳ� "
            Height          =   195
            Index           =   24
            Left            =   1920
            TabIndex        =   309
            Top             =   840
            Width           =   2250
         End
         Begin VB.Frame fra׼ȷ�� 
            Height          =   75
            Index           =   0
            Left            =   2520
            TabIndex        =   375
            Top             =   173
            Width           =   7695
         End
         Begin VB.CommandButton cmdInfo 
            Caption         =   "��"
            Height          =   240
            Index           =   59
            Left            =   4080
            TabIndex        =   308
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
            Top             =   450
            Width           =   270
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   64
            ItemData        =   "frmInMedRecEdit_YN.frx":066E
            Left            =   7890
            List            =   "frmInMedRecEdit_YN.frx":0670
            Style           =   2  'Dropdown List
            TabIndex        =   326
            Top             =   5940
            Width           =   1425
         End
         Begin VB.ComboBox cboinfo 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   63
            ItemData        =   "frmInMedRecEdit_YN.frx":0672
            Left            =   7890
            List            =   "frmInMedRecEdit_YN.frx":0674
            Style           =   2  'Dropdown List
            TabIndex        =   323
            Top             =   4335
            Width           =   1425
         End
         Begin VB.ComboBox cboinfo 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   62
            ItemData        =   "frmInMedRecEdit_YN.frx":0676
            Left            =   7890
            List            =   "frmInMedRecEdit_YN.frx":0678
            Style           =   2  'Dropdown List
            TabIndex        =   324
            Top             =   4710
            Width           =   1425
         End
         Begin VB.ComboBox cboinfo 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   61
            ItemData        =   "frmInMedRecEdit_YN.frx":067A
            Left            =   7890
            List            =   "frmInMedRecEdit_YN.frx":067C
            Style           =   2  'Dropdown List
            TabIndex        =   325
            Top             =   5100
            Width           =   1425
         End
         Begin VB.TextBox txtInfo 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   300
            Index           =   58
            Left            =   7890
            MaxLength       =   5
            TabIndex        =   322
            Top             =   3960
            Width           =   885
         End
         Begin VB.CheckBox chkInfo 
            Caption         =   "סԺ�ڼ�ʹ������Լ��"
            Height          =   195
            Index           =   21
            Left            =   6240
            TabIndex        =   321
            Top             =   3600
            Width           =   2850
         End
         Begin VB.Frame fra׼ȷ�� 
            Height          =   75
            Index           =   7
            Left            =   1440
            TabIndex        =   356
            Top             =   4980
            Width           =   4335
         End
         Begin VB.Frame fraInfection 
            Caption         =   "��Ⱦ����"
            Height          =   1695
            Left            =   6000
            TabIndex        =   319
            Top             =   1680
            Width           =   4335
            Begin VB.ListBox lstInfection 
               Height          =   1320
               ItemData        =   "frmInMedRecEdit_YN.frx":067E
               Left            =   120
               List            =   "frmInMedRecEdit_YN.frx":0685
               Style           =   1  'Checkbox
               TabIndex        =   320
               Top             =   240
               Width           =   4125
            End
         End
         Begin VB.CheckBox chkInfo 
            Caption         =   "1.C&T"
            Height          =   195
            Index           =   12
            Left            =   285
            TabIndex        =   315
            Top             =   5235
            Width           =   675
         End
         Begin VB.CheckBox chkInfo 
            Caption         =   "2.&MRI"
            Height          =   195
            Index           =   13
            Left            =   1140
            TabIndex        =   316
            Top             =   5235
            Width           =   765
         End
         Begin VB.CheckBox chkInfo 
            Caption         =   "3.��ɫ������(&R)"
            Height          =   195
            Index           =   14
            Left            =   2100
            TabIndex        =   317
            Top             =   5235
            Width           =   1665
         End
         Begin VSFlex8Ctl.VSFlexGrid vsTSJC 
            Height          =   930
            Left            =   240
            TabIndex        =   318
            Top             =   5505
            Width           =   5535
            _cx             =   9763
            _cy             =   1640
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
            BackColorSel    =   4210752
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   3
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   2
            FixedRows       =   0
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmInMedRecEdit_YN.frx":0697
            ScrollTrack     =   -1  'True
            ScrollBars      =   0
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
         Begin VSFlex8Ctl.VSFlexGrid vsfMain 
            Height          =   2850
            Left            =   240
            TabIndex        =   314
            Top             =   1920
            Width           =   5565
            _cx             =   9816
            _cy             =   5027
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
            ForeColorSel    =   -2147483642
            BackColorBkg    =   -2147483633
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483632
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483642
            FocusRect       =   2
            HighLight       =   2
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   50
            Cols            =   10
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   100
            ColWidthMax     =   2400
            ExtendLastCol   =   -1  'True
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
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   59
            Left            =   1920
            MaxLength       =   100
            TabIndex        =   307
            Top             =   420
            Width           =   2445
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ط����ʱ��"
            Height          =   180
            Index           =   128
            Left            =   780
            TabIndex        =   376
            Top             =   1260
            Width           =   1080
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��֢�໤������"
            Height          =   180
            Index           =   127
            Left            =   600
            TabIndex        =   374
            Top             =   480
            Width           =   1260
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ס��֢�໤�ң�ICU�����"
            Height          =   180
            Index           =   126
            Left            =   240
            TabIndex        =   373
            Top             =   120
            Width           =   2250
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ժ��ʽ"
            Height          =   180
            Index           =   124
            Left            =   7110
            TabIndex        =   371
            Top             =   6000
            Width           =   720
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�������������"
            Height          =   180
            Index           =   122
            Left            =   6240
            TabIndex        =   370
            Top             =   5640
            Width           =   1260
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Լ������"
            Height          =   180
            Index           =   121
            Left            =   7110
            TabIndex        =   369
            Top             =   4770
            Width           =   720
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Լ����ʽ"
            Height          =   180
            Index           =   120
            Left            =   7110
            TabIndex        =   368
            Top             =   4395
            Width           =   720
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Լ��ԭ��"
            Height          =   180
            Index           =   119
            Left            =   7110
            TabIndex        =   367
            Top             =   5160
            Width           =   720
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Сʱ"
            Height          =   180
            Index           =   116
            Left            =   8895
            TabIndex        =   366
            Top             =   4020
            Width           =   360
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Լ����ʱ��"
            Height          =   180
            Index           =   107
            Left            =   6930
            TabIndex        =   365
            Top             =   4020
            Width           =   900
         End
         Begin VB.Label lbl������Ŀ 
            AutoSize        =   -1  'True
            Caption         =   "����������Ŀ"
            Height          =   180
            Left            =   240
            TabIndex        =   312
            Top             =   1680
            Width           =   1080
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���������"
            Height          =   180
            Index           =   83
            Left            =   285
            TabIndex        =   313
            Top             =   4920
            Width           =   1080
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ml"
            Height          =   180
            Index           =   52
            Left            =   10680
            TabIndex        =   350
            Top             =   3135
            Width           =   180
         End
         Begin VB.Line Line1 
            BorderColor     =   &H8000000B&
            Index           =   8
            X1              =   1440
            X2              =   5800
            Y1              =   1800
            Y2              =   1800
         End
         Begin VB.Line Line1 
            BorderColor     =   &H8000000A&
            Index           =   9
            X1              =   1440
            X2              =   5800
            Y1              =   1785
            Y2              =   1785
         End
      End
      Begin VB.Frame fraInfo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6495
         Index           =   6
         Left            =   -74880
         TabIndex        =   348
         Top             =   390
         Width           =   10455
         Begin VB.CheckBox chkInfo 
            Caption         =   "סԺ�ڼ����Σ��"
            Height          =   195
            Index           =   1
            Left            =   7230
            TabIndex        =   306
            Top             =   3600
            Width           =   2370
         End
         Begin VB.Frame fraAdvEvent 
            Caption         =   "�����¼�"
            Height          =   2955
            Left            =   2760
            TabIndex        =   360
            Top             =   3360
            Width           =   4335
            Begin VB.ListBox lstAdvEvent 
               Height          =   1530
               ItemData        =   "frmInMedRecEdit_YN.frx":0705
               Left            =   120
               List            =   "frmInMedRecEdit_YN.frx":070C
               Style           =   1  'Checkbox
               TabIndex        =   301
               Top             =   240
               Width           =   4125
            End
            Begin VB.ComboBox cboinfo 
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   300
               IMEMode         =   3  'DISABLE
               Index           =   46
               ItemData        =   "frmInMedRecEdit_YN.frx":071E
               Left            =   3315
               List            =   "frmInMedRecEdit_YN.frx":0720
               Style           =   2  'Dropdown List
               TabIndex        =   303
               Top             =   1800
               Width           =   900
            End
            Begin VB.ComboBox cboinfo 
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   300
               IMEMode         =   3  'DISABLE
               Index           =   45
               ItemData        =   "frmInMedRecEdit_YN.frx":0722
               Left            =   1440
               List            =   "frmInMedRecEdit_YN.frx":0724
               Style           =   2  'Dropdown List
               TabIndex        =   302
               Top             =   1800
               Width           =   1335
            End
            Begin VB.ComboBox cboinfo 
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   300
               IMEMode         =   3  'DISABLE
               Index           =   48
               ItemData        =   "frmInMedRecEdit_YN.frx":0726
               Left            =   1440
               List            =   "frmInMedRecEdit_YN.frx":0728
               Style           =   2  'Dropdown List
               TabIndex        =   305
               Top             =   2520
               Width           =   2775
            End
            Begin VB.ComboBox cboinfo 
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   300
               IMEMode         =   3  'DISABLE
               Index           =   47
               ItemData        =   "frmInMedRecEdit_YN.frx":072A
               Left            =   1440
               List            =   "frmInMedRecEdit_YN.frx":072C
               Style           =   2  'Dropdown List
               TabIndex        =   304
               Top             =   2160
               Width           =   2775
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����"
               Height          =   180
               Index           =   91
               Left            =   2835
               TabIndex        =   364
               Top             =   1860
               Width           =   360
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ѹ�������ڼ�"
               Height          =   180
               Index           =   89
               Left            =   300
               TabIndex        =   363
               Top             =   1860
               Width           =   1080
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "������׹��ԭ��"
               Height          =   180
               Index           =   92
               Left            =   120
               TabIndex        =   362
               Top             =   2580
               Width           =   1260
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "������׹���˺�"
               Height          =   180
               Index           =   90
               Left            =   120
               TabIndex        =   361
               Top             =   2220
               Width           =   1260
            End
         End
         Begin VB.Frame fraPath 
            Caption         =   "�ٴ�·����Ϣ"
            Height          =   2955
            Left            =   240
            TabIndex        =   295
            Top             =   3360
            Width           =   2415
            Begin VB.CheckBox chkInfo 
               Caption         =   "����·��"
               Height          =   195
               Index           =   16
               Left            =   120
               TabIndex        =   296
               Top             =   420
               Width           =   1050
            End
            Begin VB.CheckBox chkInfo 
               Caption         =   "���·��"
               Enabled         =   0   'False
               Height          =   195
               Index           =   17
               Left            =   120
               TabIndex        =   297
               TabStop         =   0   'False
               Top             =   720
               Width           =   1050
            End
            Begin VB.CheckBox chkInfo 
               Caption         =   "����"
               Enabled         =   0   'False
               Height          =   195
               Index           =   18
               Left            =   120
               TabIndex        =   299
               TabStop         =   0   'False
               Top             =   1680
               Width           =   690
            End
            Begin VB.TextBox txtInfo 
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   300
               Index           =   61
               Left            =   360
               MaxLength       =   100
               TabIndex        =   298
               TabStop         =   0   'False
               Top             =   1275
               Width           =   1965
            End
            Begin VB.TextBox txtInfo 
               BackColor       =   &H8000000F&
               Enabled         =   0   'False
               Height          =   300
               Index           =   62
               Left            =   360
               MaxLength       =   100
               TabIndex        =   300
               TabStop         =   0   'False
               Top             =   2220
               Width           =   1965
            End
            Begin VB.CommandButton cmdPathLoad 
               Caption         =   "�Զ���ȡ"
               Height          =   350
               Left            =   1320
               TabIndex        =   357
               Top             =   360
               Width           =   975
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�˳�ԭ��"
               Height          =   180
               Index           =   117
               Left            =   375
               TabIndex        =   359
               Top             =   1020
               Width           =   720
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����ԭ��"
               Height          =   180
               Index           =   118
               Left            =   375
               TabIndex        =   358
               Top             =   1980
               Width           =   720
            End
         End
         Begin VB.CommandButton cmdAutoLoad 
            Caption         =   "�Զ���ȡ"
            Height          =   350
            Index           =   0
            Left            =   9240
            TabIndex        =   353
            Top             =   120
            Width           =   1100
         End
         Begin VSFlex8Ctl.VSFlexGrid vsKSS 
            Height          =   2685
            Left            =   240
            TabIndex        =   294
            Top             =   555
            Width           =   10155
            _cx             =   17912
            _cy             =   4736
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
            BackColorSel    =   4210752
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   3
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   4
            Cols            =   8
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmInMedRecEdit_YN.frx":072E
            ScrollTrack     =   -1  'True
            ScrollBars      =   2
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
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ҩ��ʹ���������DDD���������У�"
            Height          =   180
            Index           =   82
            Left            =   360
            TabIndex        =   293
            Top             =   270
            Width           =   3150
         End
      End
      Begin VB.Frame fraInfo 
         BorderStyle     =   0  'None
         Height          =   6345
         Index           =   5
         Left            =   -74880
         TabIndex        =   339
         Top             =   420
         Width           =   10480
         Begin VSFlex8Ctl.VSFlexGrid vs���� 
            Height          =   2715
            Left            =   45
            TabIndex        =   290
            Top             =   345
            Width           =   10440
            _cx             =   18415
            _cy             =   4789
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
            BackColor       =   16777215
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483644
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   16777215
            GridColor       =   -2147483633
            GridColorFixed  =   12632256
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   16777215
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   7
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmInMedRecEdit_YN.frx":0844
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
         Begin VSFlex8Ctl.VSFlexGrid vs���� 
            Height          =   2805
            Left            =   45
            TabIndex        =   292
            Top             =   3480
            Width           =   10440
            _cx             =   18415
            _cy             =   4948
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
            BackColor       =   16777215
            ForeColor       =   -2147483640
            BackColorFixed  =   -2147483644
            ForeColorFixed  =   -2147483630
            BackColorSel    =   -2147483635
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   16777215
            GridColor       =   -2147483633
            GridColorFixed  =   12632256
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   16777215
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   3
            Cols            =   7
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmInMedRecEdit_YN.frx":0971
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
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   3
            Left            =   1680
            TabIndex        =   352
            Top             =   3120
            Width           =   90
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            ForeColor       =   &H00FF0000&
            Height          =   180
            Index           =   2
            Left            =   1680
            TabIndex        =   351
            Top             =   30
            Width           =   90
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "���Ƽ�¼��Ϣ"
            Height          =   180
            Index           =   1
            Left            =   135
            TabIndex        =   289
            Top             =   30
            Width           =   1080
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "���Ƽ�¼��Ϣ"
            Height          =   180
            Index           =   0
            Left            =   105
            TabIndex        =   291
            Top             =   3240
            Width           =   1080
         End
      End
      Begin VB.CommandButton cmdInfo 
         Caption         =   "��"
         Enabled         =   0   'False
         Height          =   240
         Index           =   35
         Left            =   -64800
         TabIndex        =   149
         TabStop         =   0   'False
         ToolTipText     =   "ѡ��(*)"
         Top             =   6060
         Width           =   270
      End
      Begin VB.Frame fraInfo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6495
         Index           =   1
         Left            =   -74880
         TabIndex        =   337
         Top             =   420
         Width           =   10425
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   58
            ItemData        =   "frmInMedRecEdit_YN.frx":0A98
            Left            =   8760
            List            =   "frmInMedRecEdit_YN.frx":0A9A
            Style           =   2  'Dropdown List
            TabIndex        =   131
            Top             =   4425
            Width           =   1470
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   31
            ItemData        =   "frmInMedRecEdit_YN.frx":0A9C
            Left            =   1410
            List            =   "frmInMedRecEdit_YN.frx":0A9E
            Style           =   2  'Dropdown List
            TabIndex        =   133
            Top             =   4770
            Width           =   1470
         End
         Begin VB.CheckBox chkInfo 
            Caption         =   "�Ƿ�ȷ��(&B)"
            Height          =   195
            Index           =   0
            Left            =   5160
            TabIndex        =   117
            Top             =   3780
            Width           =   1290
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   5
            Left            =   1410
            Style           =   2  'Dropdown List
            TabIndex        =   116
            Top             =   3720
            Width           =   1470
         End
         Begin VB.CommandButton cmdInfo 
            Height          =   240
            Index           =   27
            Left            =   9950
            Picture         =   "frmInMedRecEdit_YN.frx":0AA0
            Style           =   1  'Graphical
            TabIndex        =   355
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(F4)"
            Top             =   3750
            Width           =   240
         End
         Begin VB.TextBox txtInfo 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   300
            Index           =   57
            Left            =   1410
            MaxLength       =   50
            TabIndex        =   121
            Top             =   4080
            Width           =   1470
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   34
            ItemData        =   "frmInMedRecEdit_YN.frx":0B96
            Left            =   5010
            List            =   "frmInMedRecEdit_YN.frx":0B98
            Style           =   2  'Dropdown List
            TabIndex        =   135
            Top             =   4770
            Width           =   1470
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   32
            ItemData        =   "frmInMedRecEdit_YN.frx":0B9A
            Left            =   4650
            List            =   "frmInMedRecEdit_YN.frx":0B9C
            Style           =   2  'Dropdown List
            TabIndex        =   146
            Top             =   5610
            Width           =   1470
         End
         Begin VB.CommandButton cmdInfo 
            Caption         =   "��"
            Height          =   240
            Index           =   50
            Left            =   10065
            TabIndex        =   156
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
            Top             =   6150
            Width           =   270
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   50
            Left            =   5640
            MaxLength       =   50
            TabIndex        =   155
            Top             =   6120
            Width           =   4695
         End
         Begin VB.TextBox txtInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   28
            Left            =   3180
            Locked          =   -1  'True
            MaxLength       =   2
            TabIndex        =   153
            Top             =   6120
            Width           =   600
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   26
            Left            =   1425
            MaxLength       =   2
            TabIndex        =   151
            Top             =   6120
            Width           =   600
         End
         Begin VB.CheckBox chkInfo 
            Caption         =   "�·�����(&Q)"
            Height          =   195
            Index           =   5
            Left            =   2040
            TabIndex        =   144
            Top             =   5670
            Width           =   1290
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   4
            Left            =   4650
            MaxLength       =   100
            TabIndex        =   141
            Top             =   5280
            Width           =   3135
         End
         Begin VB.CheckBox chkInfo 
            Caption         =   "��������ʬ��(&P)"
            Enabled         =   0   'False
            Height          =   195
            Index           =   6
            Left            =   240
            TabIndex        =   143
            Top             =   5670
            Width           =   1770
         End
         Begin VB.CheckBox chkInfo 
            Caption         =   "ҽԺ��Ⱦ����ԭѧ���(&O)"
            Height          =   195
            Index           =   9
            Left            =   7920
            TabIndex        =   142
            Top             =   5280
            Width           =   2370
         End
         Begin VB.TextBox txtInfo 
            Enabled         =   0   'False
            Height          =   300
            Index           =   35
            Left            =   8070
            MaxLength       =   150
            TabIndex        =   148
            Top             =   5610
            Width           =   2295
         End
         Begin VB.ComboBox cboinfo 
            Enabled         =   0   'False
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   53
            ItemData        =   "frmInMedRecEdit_YN.frx":0B9E
            Left            =   8760
            List            =   "frmInMedRecEdit_YN.frx":0BA0
            Style           =   2  'Dropdown List
            TabIndex        =   125
            Top             =   4080
            Width           =   1470
         End
         Begin VB.ComboBox cboinfo 
            Enabled         =   0   'False
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   52
            ItemData        =   "frmInMedRecEdit_YN.frx":0BA2
            Left            =   5010
            List            =   "frmInMedRecEdit_YN.frx":0BA4
            Style           =   2  'Dropdown List
            TabIndex        =   123
            Top             =   4080
            Width           =   1470
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   33
            ItemData        =   "frmInMedRecEdit_YN.frx":0BA6
            Left            =   8760
            List            =   "frmInMedRecEdit_YN.frx":0BA8
            Style           =   2  'Dropdown List
            TabIndex        =   137
            Top             =   4770
            Width           =   1470
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   35
            ItemData        =   "frmInMedRecEdit_YN.frx":0BAA
            Left            =   5010
            List            =   "frmInMedRecEdit_YN.frx":0BAC
            Style           =   2  'Dropdown List
            TabIndex        =   129
            Top             =   4425
            Width           =   1470
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   36
            ItemData        =   "frmInMedRecEdit_YN.frx":0BAE
            Left            =   1410
            List            =   "frmInMedRecEdit_YN.frx":0BB0
            Style           =   2  'Dropdown List
            TabIndex        =   127
            Top             =   4425
            Width           =   1470
         End
         Begin VB.Frame fraInput 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   0
            Left            =   5355
            TabIndex        =   338
            Top             =   105
            Width           =   4800
            Begin VB.OptionButton optInput 
               Caption         =   "������ϱ�׼����(&1)"
               ForeColor       =   &H00004000&
               Height          =   180
               Index           =   0
               Left            =   600
               TabIndex        =   112
               TabStop         =   0   'False
               Top             =   0
               Value           =   -1  'True
               Width           =   2010
            End
            Begin VB.OptionButton optInput 
               Caption         =   "���ݼ�����������(&2)"
               ForeColor       =   &H00004000&
               Height          =   180
               Index           =   1
               Left            =   2670
               TabIndex        =   113
               TabStop         =   0   'False
               Top             =   0
               Width           =   2010
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsDiagXY 
            Height          =   3225
            Left            =   45
            TabIndex        =   114
            Top             =   360
            Width           =   10320
            _cx             =   18203
            _cy             =   5689
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
            BackColorSel    =   4210752
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   3
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   9
            Cols            =   14
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmInMedRecEdit_YN.frx":0BB2
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
         Begin MSMask.MaskEdBox txt����ʱ�� 
            Height          =   300
            Left            =   1425
            TabIndex        =   139
            Top             =   5280
            Width           =   1830
            _ExtentX        =   3228
            _ExtentY        =   529
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   19
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Format          =   "yyyy-MM-dd hh:mm:ss"
            Mask            =   "####-##-## ##:##:##"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   27
            Left            =   8420
            Locked          =   -1  'True
            MaxLength       =   16
            TabIndex        =   119
            TabStop         =   0   'False
            Top             =   3720
            Width           =   1785
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000014&
            Index           =   13
            X1              =   0
            X2              =   10275
            Y1              =   6015
            Y2              =   6015
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            Index           =   12
            X1              =   0
            X2              =   10275
            Y1              =   6000
            Y2              =   6000
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������Ժ(&I)"
            Height          =   180
            Index           =   115
            Left            =   7560
            TabIndex        =   130
            Top             =   4485
            Width           =   1170
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ǰ������(&J)"
            Height          =   180
            Index           =   70
            Left            =   240
            TabIndex        =   132
            Top             =   4830
            Width           =   1170
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ҫ���ȷ������(&C)"
            Height          =   180
            Index           =   37
            Left            =   6690
            TabIndex        =   118
            Top             =   3780
            Width           =   1710
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ժ���(&A)"
            Height          =   180
            Index           =   28
            Left            =   420
            TabIndex        =   115
            Top             =   3780
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�����(&D)"
            ForeColor       =   &H00000000&
            Height          =   180
            Index           =   114
            Left            =   600
            TabIndex        =   120
            Top             =   4125
            Width           =   810
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ԭ��(&V)"
            Height          =   180
            Index           =   102
            Left            =   4560
            TabIndex        =   154
            Top             =   6180
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ɹ�����(&U)"
            Height          =   180
            Index           =   36
            Left            =   2145
            TabIndex        =   152
            Top             =   6180
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���ȴ���(&T)"
            Height          =   180
            Index           =   10
            Left            =   405
            TabIndex        =   150
            Top             =   6180
            Width           =   990
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ʱ��(&M)"
            Height          =   180
            Left            =   405
            TabIndex        =   138
            Top             =   5310
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ԭ��(&N)"
            Height          =   180
            Index           =   69
            Left            =   3630
            TabIndex        =   140
            Top             =   5325
            Width           =   990
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ҽԺ��Ⱦ��ԭѧ���(&S)"
            ForeColor       =   &H00808080&
            Height          =   180
            Index           =   61
            Left            =   6150
            TabIndex        =   147
            Top             =   5670
            Width           =   1890
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����������(&F)"
            Height          =   180
            Index           =   104
            Left            =   7380
            TabIndex        =   124
            Top             =   4140
            Width           =   1350
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ֻ��̶�(&E)"
            Height          =   180
            Index           =   103
            Left            =   3960
            TabIndex        =   122
            Top             =   4140
            Width           =   990
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            Index           =   6
            X1              =   45
            X2              =   10320
            Y1              =   5145
            Y2              =   5145
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000014&
            Index           =   7
            X1              =   45
            X2              =   10320
            Y1              =   5160
            Y2              =   5160
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ٴ���ʬ��(&R)"
            Height          =   180
            Index           =   71
            Left            =   3450
            TabIndex        =   145
            Top             =   5670
            Width           =   1170
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ٴ��벡��(&L)"
            Height          =   180
            Index           =   72
            Left            =   7545
            TabIndex        =   136
            Top             =   4830
            Width           =   1170
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�����벡��(&K)"
            Height          =   180
            Index           =   73
            Left            =   3810
            TabIndex        =   134
            Top             =   4830
            Width           =   1170
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ժ���Ժ(&H)"
            Height          =   180
            Index           =   74
            Left            =   3810
            TabIndex        =   128
            Top             =   4485
            Width           =   1170
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�������Ժ(&G)"
            Height          =   180
            Index           =   75
            Left            =   225
            TabIndex        =   126
            Top             =   4485
            Width           =   1170
         End
      End
      Begin VB.Frame fraInfo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6495
         Index           =   3
         Left            =   120
         TabIndex        =   335
         Top             =   420
         Width           =   10395
         Begin VB.CommandButton cmdAutoLoad 
            Caption         =   "�Զ���ȡ"
            Height          =   350
            Index           =   1
            Left            =   9160
            TabIndex        =   384
            Top             =   3215
            Width           =   1100
         End
         Begin VB.CheckBox chkInfo 
            Caption         =   "�����������"
            Height          =   195
            Index           =   23
            Left            =   4440
            TabIndex        =   193
            Top             =   6240
            Width           =   2010
         End
         Begin VB.CheckBox chkInfo 
            Caption         =   "����Χ��������"
            Height          =   195
            Index           =   22
            Left            =   2280
            TabIndex        =   192
            Top             =   6240
            Width           =   2130
         End
         Begin VB.Frame fraInput 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   2
            Left            =   165
            TabIndex        =   336
            Top             =   3300
            Width           =   6360
            Begin VB.CheckBox chkInfo 
               Caption         =   "δ�ҵ�ʱ��������¼��"
               ForeColor       =   &H00004000&
               Height          =   195
               Index           =   19
               Left            =   4200
               TabIndex        =   354
               Top             =   0
               Width           =   2145
            End
            Begin VB.OptionButton optInput 
               Caption         =   "����ICD9-CM3����(&4)"
               ForeColor       =   &H00004000&
               Height          =   180
               Index           =   5
               Left            =   2070
               TabIndex        =   190
               TabStop         =   0   'False
               Top             =   0
               Width           =   2010
            End
            Begin VB.OptionButton optInput 
               Caption         =   "����������Ŀ����(&3)"
               ForeColor       =   &H00004000&
               Height          =   180
               Index           =   4
               Left            =   0
               TabIndex        =   189
               TabStop         =   0   'False
               Top             =   0
               Value           =   -1  'True
               Width           =   2010
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsOPS 
            Height          =   2520
            Left            =   165
            TabIndex        =   191
            Top             =   3615
            Width           =   10095
            _cx             =   17806
            _cy             =   4445
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
            BackColorSel    =   4210752
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   3
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   33
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   250
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmInMedRecEdit_YN.frx":0D7F
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
         Begin VSFlex8Ctl.VSFlexGrid vsAller 
            Height          =   2850
            Left            =   165
            TabIndex        =   188
            Top             =   285
            Width           =   10095
            _cx             =   17806
            _cy             =   5027
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
            BackColorSel    =   4210752
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   3
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   0
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   250
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmInMedRecEdit_YN.frx":11B8
            ScrollTrack     =   -1  'True
            ScrollBars      =   2
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
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������������������"
            Height          =   180
            Index           =   125
            Left            =   240
            TabIndex        =   372
            Top             =   6240
            Width           =   1800
         End
      End
      Begin VB.Frame fraInfo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6495
         Index           =   2
         Left            =   -74880
         TabIndex        =   333
         Top             =   420
         Width           =   10575
         Begin VB.Frame fraSub 
            Caption         =   " ׼ȷ�� "
            Height          =   1635
            Index           =   1
            Left            =   1785
            TabIndex        =   168
            Top             =   4620
            Width           =   2415
            Begin VB.ComboBox cboinfo 
               Height          =   300
               Index           =   2
               ItemData        =   "frmInMedRecEdit_YN.frx":1226
               Left            =   825
               List            =   "frmInMedRecEdit_YN.frx":1228
               Style           =   2  'Dropdown List
               TabIndex        =   170
               Top             =   270
               Width           =   1455
            End
            Begin VB.ComboBox cboinfo 
               Height          =   300
               Index           =   11
               ItemData        =   "frmInMedRecEdit_YN.frx":122A
               Left            =   825
               List            =   "frmInMedRecEdit_YN.frx":122C
               Style           =   2  'Dropdown List
               TabIndex        =   172
               Top             =   720
               Width           =   1455
            End
            Begin VB.ComboBox cboinfo 
               Height          =   300
               Index           =   12
               ItemData        =   "frmInMedRecEdit_YN.frx":122E
               Left            =   825
               List            =   "frmInMedRecEdit_YN.frx":1230
               Style           =   2  'Dropdown List
               TabIndex        =   174
               Top             =   1140
               Width           =   1455
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��֤(&E)"
               Height          =   180
               Index           =   38
               Left            =   165
               TabIndex        =   169
               Top             =   330
               Width           =   630
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�η�(&F)"
               Height          =   180
               Index           =   39
               Left            =   165
               TabIndex        =   171
               Top             =   765
               Width           =   630
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��ҩ(&G)"
               Height          =   180
               Index           =   40
               Left            =   165
               TabIndex        =   173
               Top             =   1200
               Width           =   630
            End
         End
         Begin VB.Frame fraSub 
            Caption         =   " סԺ�ڼ䲡�� "
            Height          =   1635
            Index           =   0
            Left            =   180
            TabIndex        =   164
            Top             =   4620
            Width           =   1500
            Begin VB.CheckBox chkInfo 
               Caption         =   "Σ��(&A)"
               Height          =   195
               Index           =   2
               Left            =   405
               TabIndex        =   165
               Top             =   345
               Width           =   930
            End
            Begin VB.CheckBox chkInfo 
               Caption         =   "��֢(&B)"
               Height          =   195
               Index           =   3
               Left            =   405
               TabIndex        =   166
               Top             =   765
               Width           =   930
            End
            Begin VB.CheckBox chkInfo 
               Caption         =   "����(&D)"
               Height          =   195
               Index           =   4
               Left            =   405
               TabIndex        =   167
               Top             =   1185
               Width           =   930
            End
         End
         Begin VB.Frame fraSub 
            Caption         =   " ���Ʒ��� "
            Height          =   1635
            Index           =   2
            Left            =   4335
            TabIndex        =   175
            Top             =   4620
            Width           =   6090
            Begin VB.ComboBox cboinfo 
               Height          =   300
               Index           =   57
               ItemData        =   "frmInMedRecEdit_YN.frx":1232
               Left            =   4530
               List            =   "frmInMedRecEdit_YN.frx":1234
               Style           =   2  'Dropdown List
               TabIndex        =   187
               Top             =   1140
               Width           =   1410
            End
            Begin VB.ComboBox cboinfo 
               Height          =   300
               Index           =   56
               ItemData        =   "frmInMedRecEdit_YN.frx":1236
               Left            =   4530
               List            =   "frmInMedRecEdit_YN.frx":1238
               Style           =   2  'Dropdown List
               TabIndex        =   185
               Top             =   705
               Width           =   1410
            End
            Begin VB.ComboBox cboinfo 
               Height          =   300
               Index           =   55
               ItemData        =   "frmInMedRecEdit_YN.frx":123A
               Left            =   4530
               List            =   "frmInMedRecEdit_YN.frx":123C
               Style           =   2  'Dropdown List
               TabIndex        =   183
               Top             =   240
               Width           =   1410
            End
            Begin VB.ComboBox cboinfo 
               Height          =   300
               Index           =   13
               ItemData        =   "frmInMedRecEdit_YN.frx":123E
               Left            =   1575
               List            =   "frmInMedRecEdit_YN.frx":1240
               Style           =   2  'Dropdown List
               TabIndex        =   181
               Top             =   1140
               Width           =   1410
            End
            Begin VB.ComboBox cboinfo 
               Height          =   300
               Index           =   14
               ItemData        =   "frmInMedRecEdit_YN.frx":1242
               Left            =   1215
               List            =   "frmInMedRecEdit_YN.frx":1244
               Style           =   2  'Dropdown List
               TabIndex        =   179
               Top             =   705
               Width           =   1410
            End
            Begin VB.ComboBox cboinfo 
               Height          =   300
               Index           =   15
               ItemData        =   "frmInMedRecEdit_YN.frx":1246
               Left            =   1215
               List            =   "frmInMedRecEdit_YN.frx":1248
               Style           =   2  'Dropdown List
               TabIndex        =   177
               Top             =   270
               Width           =   1410
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��֤ʩ��(&P)"
               Height          =   180
               Index           =   113
               Left            =   3480
               TabIndex        =   186
               Top             =   1200
               Width           =   990
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ʹ����ҽ���Ƽ���(&O)"
               Height          =   180
               Index           =   112
               Left            =   2760
               TabIndex        =   184
               Top             =   765
               Width           =   1710
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ʹ����ҽ�����豸(&N)"
               Height          =   180
               Index           =   111
               Left            =   2760
               TabIndex        =   182
               Top             =   300
               Width           =   1710
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "������ҩ�Ƽ�(&K)"
               Height          =   180
               Index           =   41
               Left            =   165
               TabIndex        =   180
               Top             =   1200
               Width           =   1350
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "���ȷ���(&J)"
               Height          =   180
               Index           =   42
               Left            =   165
               TabIndex        =   178
               Top             =   765
               Width           =   990
            End
            Begin VB.Label lblInfo 
               Alignment       =   1  'Right Justify
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "�������(&I)"
               Height          =   180
               Index           =   43
               Left            =   165
               TabIndex        =   176
               Top             =   330
               Width           =   990
            End
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   37
            ItemData        =   "frmInMedRecEdit_YN.frx":124A
            Left            =   4335
            List            =   "frmInMedRecEdit_YN.frx":124C
            Style           =   2  'Dropdown List
            TabIndex        =   163
            Top             =   4035
            Width           =   1395
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   38
            ItemData        =   "frmInMedRecEdit_YN.frx":124E
            Left            =   1470
            List            =   "frmInMedRecEdit_YN.frx":1250
            Style           =   2  'Dropdown List
            TabIndex        =   161
            Top             =   4035
            Width           =   1395
         End
         Begin VB.Frame fraInput 
            Appearance      =   0  'Flat
            BorderStyle     =   0  'None
            ForeColor       =   &H80000008&
            Height          =   180
            Index           =   1
            Left            =   5340
            TabIndex        =   334
            Top             =   105
            Width           =   4800
            Begin VB.OptionButton optInput 
               Caption         =   "���ݼ�����������(&2)"
               ForeColor       =   &H00004000&
               Height          =   180
               Index           =   3
               Left            =   2760
               TabIndex        =   158
               TabStop         =   0   'False
               Top             =   0
               Width           =   2010
            End
            Begin VB.OptionButton optInput 
               Caption         =   "������ϱ�׼����(&1)"
               ForeColor       =   &H00004000&
               Height          =   180
               Index           =   2
               Left            =   720
               TabIndex        =   157
               TabStop         =   0   'False
               Top             =   0
               Value           =   -1  'True
               Width           =   2010
            End
         End
         Begin VSFlex8Ctl.VSFlexGrid vsDiagZY 
            Height          =   3555
            Left            =   165
            TabIndex        =   159
            Top             =   360
            Width           =   10320
            _cx             =   18203
            _cy             =   6271
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
            BackColorSel    =   4210752
            ForeColorSel    =   -2147483634
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   3
            HighLight       =   2
            AllowSelection  =   0   'False
            AllowBigSelection=   0   'False
            AllowUserResizing=   1
            SelectionMode   =   0
            GridLines       =   1
            GridLinesFixed  =   2
            GridLineWidth   =   1
            Rows            =   5
            Cols            =   13
            FixedRows       =   1
            FixedCols       =   1
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmInMedRecEdit_YN.frx":1252
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
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ժ���Ժ(&M)"
            Height          =   180
            Index           =   76
            Left            =   3135
            TabIndex        =   162
            Top             =   4095
            Width           =   1170
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�������Ժ(&L)"
            Height          =   180
            Index           =   77
            Left            =   270
            TabIndex        =   160
            Top             =   4095
            Width           =   1170
         End
      End
      Begin VB.Frame fraInfo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6555
         Index           =   4
         Left            =   -74880
         TabIndex        =   332
         Top             =   360
         Width           =   10455
         Begin VB.CheckBox chkInfo 
            Caption         =   "���Ѳ���(&X)"
            Height          =   195
            Index           =   20
            Left            =   7785
            TabIndex        =   236
            Top             =   1733
            Width           =   1290
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   65
            Left            =   4815
            MaxLength       =   100
            TabIndex        =   227
            Top             =   2400
            Width           =   2880
         End
         Begin VB.CommandButton cmdInfo 
            Caption         =   "��"
            Height          =   240
            Index           =   66
            Left            =   3200
            TabIndex        =   208
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
            Top             =   2430
            Width           =   270
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            Index           =   60
            Left            =   8610
            Style           =   2  'Dropdown List
            TabIndex        =   241
            Top             =   2400
            Width           =   1635
         End
         Begin VB.CommandButton cmdInfo 
            Height          =   240
            Index           =   60
            Left            =   9015
            Picture         =   "frmInMedRecEdit_YN.frx":13EC
            Style           =   1  'Graphical
            TabIndex        =   377
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(F4)"
            Top             =   5850
            Width           =   240
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   59
            ItemData        =   "frmInMedRecEdit_YN.frx":14E2
            Left            =   7860
            List            =   "frmInMedRecEdit_YN.frx":14E4
            Style           =   2  'Dropdown List
            TabIndex        =   280
            Top             =   5454
            Width           =   1425
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   30
            Left            =   4815
            MaxLength       =   5
            TabIndex        =   210
            Top             =   165
            Width           =   1140
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   31
            Left            =   4815
            MaxLength       =   5
            TabIndex        =   213
            Top             =   543
            Width           =   1140
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   32
            Left            =   4815
            MaxLength       =   5
            TabIndex        =   216
            Top             =   921
            Width           =   1140
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   33
            Left            =   4815
            MaxLength       =   5
            TabIndex        =   219
            Top             =   1299
            Width           =   1140
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   34
            Left            =   1320
            MaxLength       =   30
            TabIndex        =   205
            Top             =   2055
            Width           =   1140
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   41
            ItemData        =   "frmInMedRecEdit_YN.frx":14E6
            Left            =   1320
            List            =   "frmInMedRecEdit_YN.frx":14E8
            Style           =   2  'Dropdown List
            TabIndex        =   203
            Top             =   1680
            Width           =   1425
         End
         Begin VB.CheckBox chkInfo 
            Caption         =   "ʾ�̲���"
            Height          =   195
            Index           =   8
            Left            =   7785
            TabIndex        =   234
            Top             =   1358
            Width           =   1100
         End
         Begin VB.CheckBox chkInfo 
            Caption         =   "����(&F)"
            Height          =   195
            Index           =   7
            Left            =   3930
            TabIndex        =   260
            Top             =   4200
            Width           =   930
         End
         Begin VB.ComboBox cboinfo 
            BackColor       =   &H8000000F&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   16
            ItemData        =   "frmInMedRecEdit_YN.frx":14EA
            Left            =   7320
            List            =   "frmInMedRecEdit_YN.frx":14EC
            Locked          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   263
            TabStop         =   0   'False
            Top             =   4140
            Width           =   735
         End
         Begin VB.TextBox txtInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            Index           =   29
            Left            =   6285
            Locked          =   -1  'True
            MaxLength       =   3
            TabIndex        =   262
            TabStop         =   0   'False
            Top             =   4140
            Width           =   1020
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   27
            ItemData        =   "frmInMedRecEdit_YN.frx":14EE
            Left            =   8610
            List            =   "frmInMedRecEdit_YN.frx":14F0
            Style           =   2  'Dropdown List
            TabIndex        =   233
            Top             =   921
            Width           =   1425
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   28
            ItemData        =   "frmInMedRecEdit_YN.frx":14F2
            Left            =   8610
            List            =   "frmInMedRecEdit_YN.frx":14F4
            Style           =   2  'Dropdown List
            TabIndex        =   231
            Top             =   543
            Width           =   1425
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   29
            ItemData        =   "frmInMedRecEdit_YN.frx":14F6
            Left            =   8610
            List            =   "frmInMedRecEdit_YN.frx":14F8
            Style           =   2  'Dropdown List
            TabIndex        =   229
            Top             =   165
            Width           =   1425
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   51
            Left            =   4815
            MaxLength       =   5
            TabIndex        =   222
            Top             =   1680
            Width           =   1140
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   42
            Left            =   4815
            Style           =   2  'Dropdown List
            TabIndex        =   225
            Top             =   2055
            Width           =   1170
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   44
            ItemData        =   "frmInMedRecEdit_YN.frx":14FA
            Left            =   1260
            List            =   "frmInMedRecEdit_YN.frx":14FC
            Style           =   2  'Dropdown List
            TabIndex        =   243
            Top             =   2925
            Width           =   1500
         End
         Begin VB.TextBox txtInfo 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   300
            Index           =   40
            Left            =   3930
            MaxLength       =   100
            TabIndex        =   245
            Top             =   2925
            Width           =   5055
         End
         Begin VB.TextBox txtInfo 
            BackColor       =   &H8000000F&
            Enabled         =   0   'False
            Height          =   300
            Index           =   41
            Left            =   3930
            MaxLength       =   100
            TabIndex        =   249
            Top             =   3330
            Width           =   5055
         End
         Begin VB.OptionButton optInput 
            Caption         =   "��"
            Height          =   255
            Index           =   6
            Left            =   2130
            TabIndex        =   247
            Top             =   3360
            Value           =   -1  'True
            Width           =   495
         End
         Begin VB.OptionButton optInput 
            Caption         =   "�У�Ŀ�ģ�"
            Height          =   255
            Index           =   7
            Left            =   2700
            TabIndex        =   248
            Top             =   3360
            Width           =   1195
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   54
            Left            =   270
            Style           =   2  'Dropdown List
            TabIndex        =   246
            Top             =   3330
            Width           =   1815
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   56
            Left            =   3000
            MaxLength       =   4
            TabIndex        =   251
            Top             =   3735
            Width           =   675
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   55
            Left            =   6960
            MaxLength       =   4
            TabIndex        =   255
            Top             =   3735
            Width           =   675
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   19
            ItemData        =   "frmInMedRecEdit_YN.frx":14FE
            Left            =   1260
            List            =   "frmInMedRecEdit_YN.frx":1500
            TabIndex        =   265
            Top             =   4700
            Width           =   1425
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   30
            ItemData        =   "frmInMedRecEdit_YN.frx":1502
            Left            =   1320
            List            =   "frmInMedRecEdit_YN.frx":1504
            Style           =   2  'Dropdown List
            TabIndex        =   201
            Top             =   1299
            Width           =   1425
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   18
            ItemData        =   "frmInMedRecEdit_YN.frx":1506
            Left            =   1320
            List            =   "frmInMedRecEdit_YN.frx":1508
            Style           =   2  'Dropdown List
            TabIndex        =   199
            Top             =   921
            Width           =   1425
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   43
            ItemData        =   "frmInMedRecEdit_YN.frx":150A
            Left            =   1320
            List            =   "frmInMedRecEdit_YN.frx":150C
            Style           =   2  'Dropdown List
            TabIndex        =   195
            Top             =   165
            Width           =   1425
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   49
            Left            =   1260
            MaxLength       =   4
            TabIndex        =   259
            Top             =   4140
            Width           =   675
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   48
            Left            =   9060
            MaxLength       =   3
            TabIndex        =   257
            Top             =   3735
            Width           =   675
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   47
            Left            =   7920
            MaxLength       =   4
            TabIndex        =   256
            Top             =   3735
            Width           =   675
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   46
            Left            =   5100
            MaxLength       =   3
            TabIndex        =   253
            Top             =   3735
            Width           =   675
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   45
            Left            =   3930
            MaxLength       =   4
            TabIndex        =   252
            Top             =   3735
            Width           =   675
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   50
            ItemData        =   "frmInMedRecEdit_YN.frx":150E
            Left            =   1260
            List            =   "frmInMedRecEdit_YN.frx":1510
            TabIndex        =   288
            Top             =   6210
            Width           =   1425
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   17
            ItemData        =   "frmInMedRecEdit_YN.frx":1512
            Left            =   1320
            List            =   "frmInMedRecEdit_YN.frx":1514
            Style           =   2  'Dropdown List
            TabIndex        =   197
            Top             =   543
            Width           =   1425
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   20
            ItemData        =   "frmInMedRecEdit_YN.frx":1516
            Left            =   4020
            List            =   "frmInMedRecEdit_YN.frx":1518
            TabIndex        =   267
            Top             =   4700
            Width           =   1425
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   21
            ItemData        =   "frmInMedRecEdit_YN.frx":151A
            Left            =   7860
            List            =   "frmInMedRecEdit_YN.frx":151C
            TabIndex        =   268
            Top             =   4700
            Width           =   1425
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   22
            ItemData        =   "frmInMedRecEdit_YN.frx":151E
            Left            =   4020
            List            =   "frmInMedRecEdit_YN.frx":1520
            TabIndex        =   272
            Top             =   5077
            Width           =   1425
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   23
            ItemData        =   "frmInMedRecEdit_YN.frx":1522
            Left            =   7860
            List            =   "frmInMedRecEdit_YN.frx":1524
            TabIndex        =   274
            Top             =   5077
            Width           =   1425
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   24
            ItemData        =   "frmInMedRecEdit_YN.frx":1526
            Left            =   1260
            List            =   "frmInMedRecEdit_YN.frx":1528
            TabIndex        =   270
            Top             =   5077
            Width           =   1425
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   25
            ItemData        =   "frmInMedRecEdit_YN.frx":152A
            Left            =   1260
            List            =   "frmInMedRecEdit_YN.frx":152C
            TabIndex        =   276
            Top             =   5454
            Width           =   1425
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   26
            ItemData        =   "frmInMedRecEdit_YN.frx":152E
            Left            =   4020
            List            =   "frmInMedRecEdit_YN.frx":1530
            TabIndex        =   278
            Top             =   5454
            Width           =   1425
         End
         Begin VB.CheckBox chkInfo 
            Caption         =   "���в���"
            Height          =   195
            Index           =   11
            Left            =   9120
            TabIndex        =   235
            Top             =   1358
            Width           =   1100
         End
         Begin VB.CommandButton cmdSign 
            Caption         =   "ǩ��"
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   5475
            TabIndex        =   340
            Top             =   4693
            Width           =   555
         End
         Begin VB.CommandButton cmdUnSign 
            Caption         =   "ȡ��"
            Enabled         =   0   'False
            Height          =   315
            Index           =   0
            Left            =   6030
            TabIndex        =   341
            Top             =   4693
            Width           =   555
         End
         Begin VB.CommandButton cmdSign 
            Caption         =   "ǩ��"
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   5475
            TabIndex        =   347
            Top             =   5070
            Width           =   555
         End
         Begin VB.CommandButton cmdUnSign 
            Caption         =   "ȡ��"
            Enabled         =   0   'False
            Height          =   315
            Index           =   2
            Left            =   6030
            TabIndex        =   346
            Top             =   5070
            Width           =   555
         End
         Begin VB.CommandButton cmdUnSign 
            Caption         =   "ȡ��"
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   9870
            TabIndex        =   343
            Top             =   4693
            Width           =   555
         End
         Begin VB.CommandButton cmdSign 
            Caption         =   "ǩ��"
            Enabled         =   0   'False
            Height          =   315
            Index           =   1
            Left            =   9315
            TabIndex        =   342
            Top             =   4693
            Width           =   555
         End
         Begin VB.CommandButton cmdUnSign 
            Caption         =   "ȡ��"
            Enabled         =   0   'False
            Height          =   315
            Index           =   3
            Left            =   9870
            TabIndex        =   344
            Top             =   5070
            Width           =   555
         End
         Begin VB.CommandButton cmdSign 
            Caption         =   "ǩ��"
            Enabled         =   0   'False
            Height          =   315
            Index           =   3
            Left            =   9315
            TabIndex        =   345
            Top             =   5070
            Width           =   555
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   39
            ItemData        =   "frmInMedRecEdit_YN.frx":1532
            Left            =   4020
            List            =   "frmInMedRecEdit_YN.frx":1534
            TabIndex        =   284
            Top             =   5831
            Width           =   1425
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   40
            ItemData        =   "frmInMedRecEdit_YN.frx":1536
            Left            =   1260
            List            =   "frmInMedRecEdit_YN.frx":1538
            TabIndex        =   282
            Top             =   5831
            Width           =   1425
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   60
            Left            =   7860
            MaxLength       =   16
            TabIndex        =   285
            Top             =   5820
            Width           =   1425
         End
         Begin MSMask.MaskEdBox txt�������� 
            Height          =   300
            Left            =   8610
            TabIndex        =   238
            Top             =   2040
            Width           =   1035
            _ExtentX        =   1826
            _ExtentY        =   529
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            Format          =   "yyyy-MM-dd"
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txt����ʱ�� 
            Height          =   300
            Left            =   9690
            TabIndex        =   239
            Top             =   2040
            Width           =   585
            _ExtentX        =   1032
            _ExtentY        =   529
            _Version        =   393216
            AutoTab         =   -1  'True
            Enabled         =   0   'False
            MaxLength       =   5
            Format          =   "HH:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   66
            Left            =   1320
            MaxLength       =   100
            TabIndex        =   207
            Top             =   2400
            Width           =   2175
         End
         Begin VB.Label lblInfo 
            Caption         =   "ҽѧ��ʾ"
            Height          =   180
            Index           =   129
            Left            =   535
            TabIndex        =   206
            Top             =   2460
            Width           =   720
         End
         Begin VB.Label lblInfo 
            Caption         =   "����ҽѧ��ʾ"
            Height          =   180
            Index           =   56
            Left            =   3660
            TabIndex        =   226
            Top             =   2460
            Width           =   1080
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����(������)    ҽʦ(&3)"
            Height          =   360
            Index           =   21
            Left            =   6750
            TabIndex        =   379
            Top             =   4680
            Width           =   1140
            WordWrap        =   -1  'True
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ʱ��"
            Height          =   180
            Index           =   21
            Left            =   7800
            TabIndex        =   237
            Top             =   2100
            Width           =   720
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����״��"
            Height          =   180
            Index           =   29
            Left            =   7785
            TabIndex        =   240
            Top             =   2460
            Width           =   720
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "�ʿ�����(&X)"
            Height          =   180
            Index           =   13
            Left            =   6840
            TabIndex        =   286
            Top             =   5880
            Width           =   990
         End
         Begin VB.Label lbl���� 
            AutoSize        =   -1  'True
            Caption         =   "��������(&Y)"
            Height          =   180
            Index           =   8
            Left            =   6840
            TabIndex        =   279
            Top             =   5505
            Width           =   990
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            Index           =   15
            X1              =   120
            X2              =   10440
            Y1              =   2760
            Y2              =   2760
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000014&
            Index           =   14
            X1              =   120
            X2              =   10440
            Y1              =   2775
            Y2              =   2775
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���ϸ��(&L)"
            Height          =   180
            Index           =   47
            Left            =   3750
            TabIndex        =   209
            Top             =   225
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��λ"
            Height          =   180
            Index           =   48
            Left            =   6045
            TabIndex        =   211
            Top             =   225
            Width           =   360
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ѪС��(&M)"
            Height          =   180
            Index           =   49
            Left            =   3750
            TabIndex        =   212
            Top             =   600
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��λ"
            Height          =   180
            Index           =   50
            Left            =   6045
            TabIndex        =   214
            Top             =   600
            Width           =   360
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ѫ��(&N)"
            Height          =   180
            Index           =   51
            Left            =   3930
            TabIndex        =   215
            Top             =   975
            Width           =   810
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ȫѪ(&O)"
            Height          =   180
            Index           =   53
            Left            =   3930
            TabIndex        =   218
            Top             =   1365
            Width           =   810
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ml"
            Height          =   180
            Index           =   54
            Left            =   6045
            TabIndex        =   220
            Top             =   1365
            Width           =   180
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������(&Q)"
            Height          =   180
            Index           =   55
            Left            =   445
            TabIndex        =   204
            Top             =   2115
            Width           =   810
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ѫ��Ӧ(&K)"
            Height          =   180
            Index           =   60
            Left            =   265
            TabIndex        =   202
            Top             =   1740
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������(&G)"
            Height          =   180
            Index           =   44
            Left            =   5265
            TabIndex        =   261
            Top             =   4200
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "H&IV-Ab"
            Height          =   180
            Index           =   65
            Left            =   7965
            TabIndex        =   232
            Top             =   981
            Width           =   540
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "HC&V-Ab"
            Height          =   180
            Index           =   66
            Left            =   7965
            TabIndex        =   230
            Top             =   603
            Width           =   540
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "HB&sAg"
            Height          =   180
            Index           =   67
            Left            =   8055
            TabIndex        =   228
            Top             =   225
            Width           =   450
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ml"
            Height          =   180
            Index           =   105
            Left            =   6045
            TabIndex        =   223
            Top             =   1740
            Width           =   180
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�������(&B)"
            Height          =   180
            Index           =   106
            Left            =   3750
            TabIndex        =   221
            Top             =   1740
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ml"
            Height          =   180
            Index           =   110
            Left            =   6045
            TabIndex        =   217
            Top             =   960
            Width           =   180
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ѫǰ��9����(&E)"
            Height          =   180
            Index           =   81
            Left            =   3120
            TabIndex        =   224
            Top             =   2115
            Width           =   1620
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ժ��ʽ"
            Height          =   180
            Index           =   87
            Left            =   480
            TabIndex        =   242
            Top             =   2985
            Width           =   720
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ת��"
            Height          =   180
            Index           =   88
            Left            =   3525
            TabIndex        =   244
            Top             =   2985
            Width           =   360
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ҽʦ(&1)"
            Height          =   180
            Index           =   57
            Left            =   240
            TabIndex        =   264
            Top             =   4760
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������(&D)"
            Height          =   180
            Index           =   86
            Left            =   270
            TabIndex        =   194
            Top             =   240
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "  ��Ժ��        ��         Сʱ        ����"
            Height          =   180
            Index           =   101
            Left            =   6225
            TabIndex        =   254
            Top             =   3795
            Width           =   3870
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������ʹ��(&C)         Сʱ"
            Height          =   180
            Index           =   100
            Left            =   90
            TabIndex        =   258
            Top             =   4200
            Width           =   2340
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "­�����˻��߻���ʱ��(&P) ��Ժǰ        ��         Сʱ        ����"
            Height          =   180
            Index           =   99
            Left            =   285
            TabIndex        =   250
            Top             =   3795
            Width           =   5850
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���λ�ʿ(&A)"
            Height          =   180
            Index           =   95
            Left            =   240
            TabIndex        =   287
            Top             =   6270
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ѫ��(&J)"
            Height          =   180
            Index           =   45
            Left            =   625
            TabIndex        =   196
            Top             =   603
            Width           =   630
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "&Rh"
            Height          =   180
            Index           =   46
            Left            =   960
            TabIndex        =   198
            Top             =   975
            Width           =   180
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000014&
            Index           =   10
            X1              =   120
            X2              =   10440
            Y1              =   4575
            Y2              =   4575
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            Index           =   11
            X1              =   120
            X2              =   10440
            Y1              =   4560
            Y2              =   4560
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������(&2)"
            Height          =   180
            Index           =   20
            Left            =   3180
            TabIndex        =   266
            Top             =   4755
            Width           =   810
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ҽʦ(&5)"
            Height          =   180
            Index           =   22
            Left            =   3000
            TabIndex        =   271
            Top             =   5137
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "סԺҽʦ(&6)"
            Height          =   180
            Index           =   23
            Left            =   6840
            TabIndex        =   273
            Top             =   5137
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����ҽʦ(&4)"
            Height          =   180
            Index           =   62
            Left            =   240
            TabIndex        =   269
            Top             =   5137
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�о���ҽʦ(&7)"
            Height          =   180
            Index           =   63
            Left            =   60
            TabIndex        =   275
            Top             =   5514
            Width           =   1170
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ʵϰҽʦ(&8)"
            Height          =   180
            Index           =   64
            Left            =   3000
            TabIndex        =   277
            Top             =   5514
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Һ��Ӧ(&S)"
            Height          =   180
            Index           =   68
            Left            =   265
            TabIndex        =   200
            Top             =   1359
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ʿػ�ʿ(&0)"
            Height          =   180
            Index           =   58
            Left            =   3000
            TabIndex        =   283
            Top             =   5891
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ʿ�ҽʦ(&9)"
            Height          =   180
            Index           =   59
            Left            =   240
            TabIndex        =   281
            Top             =   5891
            Width           =   990
         End
      End
      Begin VB.Frame fraInfo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   6495
         Index           =   0
         Left            =   -74880
         TabIndex        =   330
         Top             =   420
         Width           =   10545
         Begin ZlPatiAddress.PatiAddress PatiAddress���� 
            Height          =   360
            Left            =   7755
            TabIndex        =   45
            Top             =   2040
            Width           =   1905
            _ExtentX        =   3360
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Items           =   2
            MaxLength       =   50
         End
         Begin ZlPatiAddress.PatiAddress PatiAddress������ 
            Height          =   360
            Left            =   1290
            TabIndex        =   41
            Top             =   2070
            Width           =   2910
            _ExtentX        =   5133
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Items           =   3
            MaxLength       =   50
         End
         Begin ZlPatiAddress.PatiAddress PatiAddress���ڵ�ַ 
            Height          =   360
            Left            =   1290
            TabIndex        =   61
            Top             =   3240
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   50
         End
         Begin ZlPatiAddress.PatiAddress PatiAddress��סַ 
            Height          =   360
            Left            =   1290
            TabIndex        =   53
            Top             =   2880
            Width           =   5175
            _ExtentX        =   9128
            _ExtentY        =   635
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "����"
               Size            =   9
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            MaxLength       =   50
         End
         Begin VB.CommandButton cmdInfo 
            Caption         =   "��"
            Height          =   240
            Index           =   53
            Left            =   5055
            TabIndex        =   63
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
            Top             =   3270
            Width           =   270
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   54
            Left            =   7215
            MaxLength       =   6
            TabIndex        =   65
            Top             =   3240
            Width           =   1275
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   53
            Left            =   1320
            MaxLength       =   100
            TabIndex        =   62
            Top             =   3240
            Width           =   4035
         End
         Begin VB.CommandButton cmdInfo 
            Caption         =   "��"
            Height          =   240
            Index           =   52
            Left            =   9990
            TabIndex        =   47
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
            Top             =   2070
            Width           =   270
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   52
            Left            =   7785
            MaxLength       =   30
            TabIndex        =   46
            Top             =   2040
            Width           =   2490
         End
         Begin VB.CheckBox chkInfo 
            Caption         =   "��Ժǰ����Ժ����"
            Height          =   195
            Index           =   15
            Left            =   3645
            TabIndex        =   93
            Top             =   5228
            Width           =   2010
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   44
            Left            =   8535
            MaxLength       =   8
            TabIndex        =   32
            Top             =   1350
            Width           =   1710
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   43
            Left            =   4410
            MaxLength       =   8
            TabIndex        =   29
            Top             =   1350
            Width           =   1755
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            Index           =   51
            Left            =   2430
            Style           =   2  'Dropdown List
            TabIndex        =   27
            Top             =   1357
            Width           =   645
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   42
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   26
            Top             =   1357
            Width           =   1080
         End
         Begin VB.CommandButton cmdInfo 
            Caption         =   "��"
            Height          =   240
            Index           =   23
            Left            =   5055
            TabIndex        =   98
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
            Top             =   5550
            Width           =   270
         End
         Begin VB.CommandButton cmdInfo 
            Caption         =   "��"
            Height          =   240
            Index           =   6
            Left            =   6360
            TabIndex        =   43
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
            Top             =   2085
            Width           =   270
         End
         Begin VB.CommandButton cmdInfo 
            Caption         =   "��"
            Height          =   240
            Index           =   15
            Left            =   5055
            TabIndex        =   81
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
            Top             =   4380
            Width           =   270
         End
         Begin VB.CommandButton cmdInfo 
            Caption         =   "��"
            Height          =   240
            Index           =   10
            Left            =   5055
            TabIndex        =   68
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
            Top             =   3645
            Width           =   270
         End
         Begin VB.CommandButton cmdInfo 
            Caption         =   "��"
            Height          =   240
            Index           =   1
            Left            =   5055
            TabIndex        =   55
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
            Top             =   2910
            Width           =   270
         End
         Begin VB.CommandButton cmdInfo 
            Caption         =   "��"
            Height          =   240
            Index           =   24
            Left            =   7305
            TabIndex        =   101
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
            Top             =   5550
            Width           =   270
         End
         Begin VB.CommandButton cmdInfo 
            Caption         =   "��"
            Height          =   240
            Index           =   25
            Left            =   10005
            TabIndex        =   103
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
            Top             =   5550
            Width           =   270
         End
         Begin VB.CommandButton cmdInfo 
            Caption         =   "��"
            Height          =   240
            Index           =   36
            Left            =   9405
            TabIndex        =   87
            TabStop         =   0   'False
            ToolTipText     =   "ѡ��(*)"
            Top             =   4380
            Width           =   270
         End
         Begin VB.TextBox txtInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   0
            Left            =   1320
            Locked          =   -1  'True
            MaxLength       =   18
            TabIndex        =   1
            TabStop         =   0   'False
            Top             =   135
            Width           =   1740
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   36
            Left            =   7215
            MaxLength       =   30
            TabIndex        =   84
            Top             =   4350
            Width           =   2490
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   0
            ItemData        =   "frmInMedRecEdit_YN.frx":153A
            Left            =   7800
            List            =   "frmInMedRecEdit_YN.frx":153C
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   135
            Width           =   2475
         End
         Begin VB.TextBox txtInfo 
            Alignment       =   2  'Center
            BackColor       =   &H8000000F&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   2
            Left            =   3390
            Locked          =   -1  'True
            TabIndex        =   3
            TabStop         =   0   'False
            Top             =   135
            Width           =   375
         End
         Begin VB.TextBox txtInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            Index           =   3
            Left            =   1320
            Locked          =   -1  'True
            MaxLength       =   64
            TabIndex        =   8
            TabStop         =   0   'False
            Top             =   615
            Width           =   1740
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   1
            ItemData        =   "frmInMedRecEdit_YN.frx":153E
            Left            =   4410
            List            =   "frmInMedRecEdit_YN.frx":1540
            Style           =   2  'Dropdown List
            TabIndex        =   10
            Top             =   615
            Width           =   1605
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   5
            Left            =   1320
            MaxLength       =   10
            TabIndex        =   15
            Top             =   990
            Width           =   1080
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   3
            ItemData        =   "frmInMedRecEdit_YN.frx":1542
            Left            =   8535
            List            =   "frmInMedRecEdit_YN.frx":1544
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   990
            Width           =   1740
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   4
            ItemData        =   "frmInMedRecEdit_YN.frx":1546
            Left            =   7785
            List            =   "frmInMedRecEdit_YN.frx":1548
            Style           =   2  'Dropdown List
            TabIndex        =   39
            Top             =   1725
            Width           =   2490
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   7
            Left            =   1320
            Style           =   2  'Dropdown List
            TabIndex        =   35
            Top             =   1725
            Width           =   1740
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   8
            ItemData        =   "frmInMedRecEdit_YN.frx":154A
            Left            =   4410
            List            =   "frmInMedRecEdit_YN.frx":154C
            Style           =   2  'Dropdown List
            TabIndex        =   37
            Top             =   1725
            Width           =   2250
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   6
            Left            =   1320
            MaxLength       =   50
            TabIndex        =   42
            Top             =   2055
            Width           =   5295
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   1
            Left            =   1320
            MaxLength       =   100
            TabIndex        =   54
            Top             =   2880
            Width           =   4035
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   8
            Left            =   7215
            MaxLength       =   20
            TabIndex        =   57
            Top             =   2880
            Width           =   1275
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   9
            Left            =   9330
            MaxLength       =   6
            TabIndex        =   59
            Top             =   2880
            Width           =   945
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   10
            Left            =   1320
            MaxLength       =   100
            TabIndex        =   67
            Top             =   3615
            Width           =   4035
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   11
            Left            =   7215
            MaxLength       =   20
            TabIndex        =   70
            Top             =   3615
            Width           =   1275
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   12
            Left            =   9330
            MaxLength       =   6
            TabIndex        =   72
            Top             =   3615
            Width           =   945
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   13
            Left            =   1320
            MaxLength       =   64
            TabIndex        =   74
            Top             =   3990
            Width           =   1545
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   9
            ItemData        =   "frmInMedRecEdit_YN.frx":154E
            Left            =   3645
            List            =   "frmInMedRecEdit_YN.frx":1550
            Style           =   2  'Dropdown List
            TabIndex        =   76
            Top             =   3990
            Width           =   1700
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   14
            Left            =   7215
            MaxLength       =   20
            TabIndex        =   78
            Top             =   3990
            Width           =   1275
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   15
            Left            =   1320
            MaxLength       =   100
            TabIndex        =   80
            Top             =   4350
            Width           =   4035
         End
         Begin VB.TextBox txtInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   16
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   85
            TabStop         =   0   'False
            Top             =   4815
            Width           =   1545
         End
         Begin VB.TextBox txtInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            Index           =   17
            Left            =   3645
            Locked          =   -1  'True
            TabIndex        =   88
            TabStop         =   0   'False
            Top             =   4815
            Width           =   1695
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   18
            Left            =   6225
            MaxLength       =   100
            TabIndex        =   90
            Top             =   4815
            Width           =   1305
         End
         Begin VB.TextBox txtInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   19
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   105
            TabStop         =   0   'False
            Top             =   5865
            Width           =   1545
         End
         Begin VB.TextBox txtInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            Index           =   20
            Left            =   3645
            Locked          =   -1  'True
            TabIndex        =   107
            TabStop         =   0   'False
            Top             =   5865
            Width           =   1695
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   21
            Left            =   6225
            MaxLength       =   100
            TabIndex        =   109
            Top             =   5865
            Width           =   1275
         End
         Begin VB.TextBox txtInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   22
            Left            =   8610
            Locked          =   -1  'True
            TabIndex        =   111
            TabStop         =   0   'False
            Top             =   5865
            Width           =   1665
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   23
            Left            =   3645
            MaxLength       =   100
            TabIndex        =   97
            Top             =   5520
            Width           =   1695
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   24
            Left            =   5640
            MaxLength       =   100
            TabIndex        =   100
            Top             =   5520
            Width           =   1965
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   25
            Left            =   7830
            MaxLength       =   100
            TabIndex        =   102
            Top             =   5520
            Width           =   2445
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            Index           =   10
            Left            =   2430
            Style           =   2  'Dropdown List
            TabIndex        =   16
            Top             =   990
            Width           =   645
         End
         Begin VB.CheckBox chkInfo 
            Caption         =   "����Ժ"
            Height          =   285
            Index           =   10
            Left            =   4515
            TabIndex        =   4
            Top             =   143
            Width           =   840
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   6
            ItemData        =   "frmInMedRecEdit_YN.frx":1552
            Left            =   1320
            List            =   "frmInMedRecEdit_YN.frx":1554
            TabIndex        =   49
            Top             =   2415
            Width           =   4035
         End
         Begin VB.TextBox txtInfo 
            BackColor       =   &H8000000F&
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   7
            Left            =   1320
            Locked          =   -1  'True
            TabIndex        =   95
            TabStop         =   0   'False
            Top             =   5520
            Width           =   1545
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            Index           =   37
            Left            =   6960
            MaxLength       =   20
            TabIndex        =   51
            Top             =   2400
            Width           =   3315
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   39
            Left            =   6075
            MaxLength       =   5
            TabIndex        =   21
            Top             =   990
            Width           =   555
         End
         Begin VB.TextBox txtInfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   38
            Left            =   4410
            MaxLength       =   5
            TabIndex        =   18
            Top             =   990
            Width           =   555
         End
         Begin VB.ComboBox cboinfo 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Index           =   49
            ItemData        =   "frmInMedRecEdit_YN.frx":1556
            Left            =   1320
            List            =   "frmInMedRecEdit_YN.frx":1558
            Style           =   2  'Dropdown List
            TabIndex        =   92
            Top             =   5175
            Width           =   1545
         End
         Begin MSMask.MaskEdBox txt����ʱ�� 
            Height          =   300
            Left            =   9660
            TabIndex        =   13
            Top             =   615
            Width           =   615
            _ExtentX        =   1085
            _ExtentY        =   529
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   5
            Format          =   "HH:mm"
            Mask            =   "##:##"
            PromptChar      =   "_"
         End
         Begin MSMask.MaskEdBox txt�������� 
            Height          =   300
            Left            =   8535
            TabIndex        =   12
            Top             =   615
            Width           =   1080
            _ExtentX        =   1905
            _ExtentY        =   529
            _Version        =   393216
            AutoTab         =   -1  'True
            MaxLength       =   10
            Format          =   "yyyy-MM-dd"
            Mask            =   "####-##-##"
            PromptChar      =   "_"
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ʱ�(&U)"
            Height          =   180
            Index           =   109
            Left            =   6525
            TabIndex        =   64
            Top             =   3300
            Width           =   630
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���ڵ�ַ(&T)"
            Height          =   180
            Index           =   108
            Left            =   300
            TabIndex        =   60
            Top             =   3300
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����(&M)"
            Height          =   180
            Index           =   93
            Left            =   7125
            TabIndex        =   44
            Top             =   2100
            Width           =   630
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            Height          =   180
            Index           =   2
            Left            =   10320
            TabIndex        =   33
            Top             =   1410
            Width           =   180
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������Ժ����"
            Height          =   180
            Index           =   98
            Left            =   7230
            TabIndex        =   31
            Top             =   1410
            Width           =   1260
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            Height          =   180
            Index           =   1
            Left            =   6240
            TabIndex        =   30
            Top             =   1410
            Width           =   180
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������������"
            Height          =   180
            Index           =   97
            Left            =   3105
            TabIndex        =   28
            Top             =   1410
            Width           =   1260
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "Ӥ�׶�����(&O)"
            Height          =   180
            Index           =   96
            Left            =   120
            TabIndex        =   25
            Top             =   1417
            Width           =   1170
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "סԺ��(&A)"
            Height          =   180
            Index           =   0
            Left            =   480
            TabIndex        =   0
            Top             =   195
            Width           =   810
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��     ��סԺ"
            Height          =   180
            Index           =   3
            Left            =   3180
            TabIndex        =   2
            Top             =   195
            Width           =   1170
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���ѷ�ʽ(&B)"
            Height          =   180
            Index           =   2
            Left            =   6765
            TabIndex        =   5
            Top             =   195
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����(&D)"
            Height          =   180
            Index           =   4
            Left            =   660
            TabIndex        =   7
            Top             =   675
            Width           =   630
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�Ա�(&E)"
            Height          =   180
            Index           =   5
            Left            =   3735
            TabIndex        =   9
            Top             =   675
            Width           =   630
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��������(&F)"
            Height          =   180
            Index           =   6
            Left            =   7485
            TabIndex        =   11
            Top             =   675
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����(&G)"
            Height          =   180
            Index           =   7
            Left            =   660
            TabIndex        =   14
            Top             =   1050
            Width           =   630
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����(&I)"
            Height          =   180
            Index           =   8
            Left            =   7845
            TabIndex        =   23
            Top             =   1050
            Width           =   630
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ְҵ(&J)"
            Height          =   180
            Index           =   9
            Left            =   7125
            TabIndex        =   38
            Top             =   1785
            Width           =   630
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����(&3)"
            Height          =   180
            Index           =   11
            Left            =   6525
            TabIndex        =   83
            Top             =   4410
            Width           =   630
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����(&K)"
            Height          =   180
            Index           =   12
            Left            =   660
            TabIndex        =   34
            Top             =   1785
            Width           =   630
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����(&L)"
            Height          =   180
            Index           =   13
            Left            =   3735
            TabIndex        =   36
            Top             =   1785
            Width           =   630
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�����ص�(&N)"
            Height          =   180
            Index           =   14
            Left            =   300
            TabIndex        =   40
            Top             =   2160
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���֤��(&P)"
            Height          =   180
            Index           =   15
            Left            =   300
            TabIndex        =   48
            Top             =   2475
            Width           =   990
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000014&
            Index           =   0
            X1              =   135
            X2              =   10320
            Y1              =   525
            Y2              =   525
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            Index           =   1
            X1              =   135
            X2              =   10320
            Y1              =   510
            Y2              =   510
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000014&
            Index           =   2
            X1              =   135
            X2              =   10320
            Y1              =   2775
            Y2              =   2775
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            Index           =   3
            X1              =   135
            X2              =   10320
            Y1              =   2760
            Y2              =   2760
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��סַ(&Q)"
            Height          =   180
            Index           =   1
            Left            =   480
            TabIndex        =   52
            Top             =   2940
            Width           =   810
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�绰(&R)"
            Height          =   180
            Index           =   16
            Left            =   6525
            TabIndex        =   56
            Top             =   2940
            Width           =   630
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ʱ�(&S)"
            Height          =   180
            Index           =   17
            Left            =   8670
            TabIndex        =   58
            Top             =   2940
            Width           =   630
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "������λ(&V)"
            Height          =   180
            Index           =   18
            Left            =   300
            TabIndex        =   66
            Top             =   3675
            Width           =   990
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�绰(&W)"
            Height          =   180
            Index           =   19
            Left            =   6525
            TabIndex        =   69
            Top             =   3675
            Width           =   630
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�ʱ�(&X)"
            Height          =   180
            Index           =   123
            Left            =   8670
            TabIndex        =   71
            Top             =   3675
            Width           =   630
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ϵ������(&Y)"
            Height          =   180
            Index           =   79
            Left            =   120
            TabIndex        =   73
            Top             =   4050
            Width           =   1170
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ϵ(&Z)"
            Height          =   180
            Index           =   78
            Left            =   2985
            TabIndex        =   75
            Top             =   4050
            Width           =   630
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "�绰(&1)"
            Height          =   180
            Index           =   80
            Left            =   6525
            TabIndex        =   77
            Top             =   4050
            Width           =   630
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��ϵ�˵�ַ(&2)"
            Height          =   180
            Index           =   24
            Left            =   120
            TabIndex        =   79
            Top             =   4410
            Width           =   1170
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000014&
            Index           =   4
            X1              =   135
            X2              =   10080
            Y1              =   4740
            Y2              =   4740
         End
         Begin VB.Line Line1 
            BorderColor     =   &H80000010&
            Index           =   5
            X1              =   135
            X2              =   10320
            Y1              =   4725
            Y2              =   4725
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ժʱ��"
            Height          =   180
            Index           =   25
            Left            =   570
            TabIndex        =   82
            Top             =   4875
            Width           =   720
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            Height          =   180
            Index           =   26
            Left            =   3255
            TabIndex        =   86
            Top             =   4875
            Width           =   360
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            Height          =   180
            Index           =   27
            Left            =   5835
            TabIndex        =   89
            Top             =   4875
            Width           =   360
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ժʱ��"
            Height          =   180
            Index           =   29
            Left            =   570
            TabIndex        =   104
            Top             =   5925
            Width           =   720
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            Height          =   180
            Index           =   30
            Left            =   3255
            TabIndex        =   106
            Top             =   5925
            Width           =   360
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����"
            Height          =   180
            Index           =   31
            Left            =   5835
            TabIndex        =   108
            Top             =   5925
            Width           =   360
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ת��"
            Height          =   180
            Index           =   33
            Left            =   3255
            TabIndex        =   96
            Top             =   5580
            Width           =   360
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            Height          =   180
            Index           =   34
            Left            =   5415
            TabIndex        =   99
            Top             =   5580
            Width           =   180
         End
         Begin VB.Label lblInfo 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            Height          =   180
            Index           =   35
            Left            =   7635
            TabIndex        =   331
            Top             =   5580
            Width           =   180
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "סԺ����"
            Height          =   180
            Index           =   32
            Left            =   7860
            TabIndex        =   110
            Top             =   5925
            Width           =   720
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���ʱ��"
            Height          =   180
            Index           =   84
            Left            =   570
            TabIndex        =   94
            Top             =   5580
            Width           =   720
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����֤��"
            Height          =   180
            Index           =   85
            Left            =   6195
            TabIndex        =   50
            Top             =   2460
            Width           =   720
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "kg"
            Height          =   180
            Index           =   0
            Left            =   6675
            TabIndex        =   22
            Top             =   1050
            Width           =   180
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����(&W)"
            Height          =   180
            Index           =   24
            Left            =   5400
            TabIndex        =   20
            Top             =   1050
            Width           =   630
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "cm"
            Height          =   180
            Index           =   1
            Left            =   5040
            TabIndex        =   19
            Top             =   1050
            Width           =   180
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "���(&H)"
            Height          =   180
            Index           =   23
            Left            =   3735
            TabIndex        =   17
            Top             =   1050
            Width           =   630
         End
         Begin VB.Label lblInfo 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��Ժ;��"
            Height          =   180
            Index           =   94
            Left            =   555
            TabIndex        =   91
            Top             =   5235
            Width           =   720
         End
      End
   End
   Begin VB.Timer timThis 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   4920
      Top             =   6840
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   9840
      TabIndex        =   329
      Top             =   7155
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   8595
      TabIndex        =   327
      Top             =   7155
      Width           =   1100
   End
   Begin MSComCtl2.MonthView dtpInfo 
      Height          =   2160
      Left            =   2760
      TabIndex        =   378
      TabStop         =   0   'False
      Top             =   1320
      Visible         =   0   'False
      Width           =   2775
      _ExtentX        =   4895
      _ExtentY        =   3810
      _Version        =   393216
      ForeColor       =   -2147483630
      BackColor       =   -2147483633
      BorderStyle     =   1
      Appearance      =   0
      StartOfWeek     =   115802113
      TitleBackColor  =   8421504
      TitleForeColor  =   16777215
      CurrentDate     =   38003
   End
   Begin VB.Image imgButtonDel 
      Height          =   240
      Left            =   3480
      Picture         =   "frmInMedRecEdit_YN.frx":155A
      Top             =   7320
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image imgButtonNew 
      Height          =   240
      Left            =   4320
      Picture         =   "frmInMedRecEdit_YN.frx":7DAC
      Top             =   7320
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Menu menuPriview 
      Caption         =   "Ԥ����ҳ"
      Visible         =   0   'False
      Begin VB.Menu menuPage 
         Caption         =   "����(&1)"
         Index           =   1
      End
      Begin VB.Menu menuPage 
         Caption         =   "����(&2)"
         Index           =   2
      End
      Begin VB.Menu menuPage 
         Caption         =   "��ҳ1(&3)"
         Index           =   3
      End
      Begin VB.Menu menuPage 
         Caption         =   "��ҳ2(&4)"
         Index           =   4
      End
   End
   Begin VB.Menu menuPrint 
      Caption         =   "��ӡ��ҳ"
      Visible         =   0   'False
      Begin VB.Menu menuPagePrint 
         Caption         =   "����(&1)"
         Index           =   1
      End
      Begin VB.Menu menuPagePrint 
         Caption         =   "����(&2)"
         Index           =   2
      End
      Begin VB.Menu menuPagePrint 
         Caption         =   "��ҳ1(&3)"
         Index           =   3
      End
      Begin VB.Menu menuPagePrint 
         Caption         =   "��ҳ2(&4)"
         Index           =   4
      End
      Begin VB.Menu menuPagePrint 
         Caption         =   "����+��ҳ1(&5)"
         Index           =   5
      End
      Begin VB.Menu menuPagePrint 
         Caption         =   "����+��ҳ2(&6)"
         Index           =   6
      End
   End
End
Attribute VB_Name = "frmInMedRecEdit_YN"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public Event Closed(ByVal EditCancel As Boolean, ByVal str����ID As String, ByVal str���ID As String) 'סԺ��ҳ�ر��¼�

Private mcol��ԱSQL As Collection
Private mblnReadOnly As Boolean
Private mstrPrivs As String
Private mlng����ID As Long
Private mlng��ҳID As Long
Private mlng����ID As Long
Private mbln��Ժ As Boolean
Private mint���� As Integer
Private mlngPathState As Long   '·��״̬   -1=δ����,0-�����ϵ���������1-ִ���У�2-����������3-�������
Private mlngDiagnosisType As Long '����·��ʱ���������:1-��ҽ�������;2-��ҽ��Ժ���;11-��ҽ�������;12-��ҽ��Ժ���
Private mstr����ID As String   '���ڱ��漲��ID,��Closed�¼��д��ݸ�������
Private mstr���ID As String   '���ڱ������ID,��Closed�¼��д��ݸ�������
Private mstrPathDiag As String '����һ��סԺ�ڶ���·���Ժ�������б��������1|����ID1|���ID1,�������2|����ID2|���ID2```
Private WithEvents mobjReport As clsReport
Attribute mobjReport.VB_VarHelpID = -1
Private mbln��ʿվ As Boolean

Private mblnIsFirst As Boolean
Private mstrZYDiagInfo As String
Private mstrXYDiagInfo As String
Public mblnDiagChange As Boolean
Private mlng���� As String      '�����Ƿ���  0-����飬2-��ʾ��1-������д
Private mlng�����ж� As Long
Private mlng������� As Long
Private mlngSize As Long '��¼������ҳ�ӱ���Ϣֵ���ֶγ���
Private mstr����������� As String
Private mobjESign As Object           'ǩ����������
Private mblnIsPathOutTime As Boolean   '���·����ʱ���Ƿ�ȳ�Ժ��ϼ�¼ʱ���
Private mlngDateIndex As Long


Private mstr���� As String
Private mblnDiagnose As Boolean

Private mblnOpen As Boolean
Private mbln��ҽ As Boolean

Private mbln�������� As Boolean

Private mstrLike As String
Private mint���� As Integer
Private mblnChange As Boolean
Private mblnOk As Boolean
Private mblnNoClick As Boolean
Private mblnReturn As Boolean
Private mstrDelete As String
Private mlngNum As Long
Private mlngSelNum As Long
Private mlngNumBack As Long
Private mbln��ҳ��� As Boolean
Private Const GRD_LOSTFOCUS_COLORSEL = &H80000010  '�뿪����ʱ,ѡ�����ʾ��ɫ
Private Const GRD_GOTFOCUS_COLORSEL = &H8000000D '16772055 '    '����ؼ�ʱ,ѡ����ʾ��ɫ
Private Declare Function SetFocusHwnd Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long
Private mbln���ýṹ����ַ As Boolean
Private mbln��ʹ����ҽ��Ŀ As Boolean
Private mblnҽ����ʿ������ҳ As Boolean

Private mrsXYDiag  As ADODB.Recordset '��ҽ��ϼ�¼��
Private mrsZYDiag  As ADODB.Recordset '��ҽ��ϼ�¼��

Private Enum Tab�˵�
    TAB_������Ϣ = 0
    TAB_��ҽ��� = 1
    TAB_��ҽ��� = 2
    TAB_���������� = 3
    TAB_סԺ��� = 4
    TAB_�����뻯�� = 5
    TAB_�ض�ҩƷ = 6
    TAB_���� = 7
End Enum

Private Enum COL������
    col������� = 0
    col��ϱ��� = 1
    col������� = 2
    col��ҽ֤�� = 3
    col��ע = 4
    col��Ժ���� = 5
    col��Ժ��� = 6
    col�Ƿ�δ�� = 7
    col�Ƿ����� = 8
    col���� = 9
    colDel = 10
    col���ID = 11
    col����ID = 12
    col���� = 13 '1-��ҽ�������;2-��ҽ��Ժ���;3-��Ժ���(�������);5-Ժ�ڸ�Ⱦ;6-�������;7-�����ж���;10-����֢
    
    colzy���� = 7
    colzyDel = 8
    colzy���ID = 9
    colzy����ID = 10
    colzy֤��ID = 11
    colzy���� = 12
End Enum
Private Enum COL�������
    col�������� = 0
    COL������� = 1
    col�������� = 2
    col�������� = 3
    col�ٴ����� = 4
    col����ҽʦ = 5
    col������ʿ = 6
    col����1 = 7
    col����2 = 8
    col�������� = 9
    colASA�ּ� = 10
    colNNIS�ּ� = 11
    col�������� = 12
    col����ҽʦ = 13
    col�п����� = 14
    colԤ���ÿ���ҩ = 15
    col����ҩ���� = 16
    col��Ԥ�ڵĶ������� = 17
    col������֢ = 18
    col������������ = 19
    col��������֢ = 20
    col�����Ѫ��Ѫ�� = 21
    col�����˿��ѿ� = 22
    col�������Ѫ˨ = 23
    col���������л���� = 24
    col�������˥�� = 25
    col�����˨�� = 26
    col�����Ѫ֢ = 27
    col�����Źؽڹ��� = 28
    col��������ID = 29
    col������ĿID = 30
    col����ID = 31
    col����ʽ = 32
End Enum

Private Enum ������Ϣ
    cbo���ʽ = 0
    cbo�Ա� = 1
    cbo���� = 3
    cboְҵ = 4
    cbo��Ժ���� = 5
    cbo���֤�� = 6
    txt���� = 36
    cbo���� = 7
    cbo���� = 8
    cbo��ϵ�˹�ϵ = 9
    cbo���䵥λ = 10
    txtסԺ�� = 0
    txt��ͥ��ַ = 1
    txtסԺ���� = 2
    txt���� = 3
    'txt�������� = 4
    txt���� = 5
    txt�����ص� = 6
    txt��ͥ�绰 = 8
    txt��ͥ�ʱ� = 9
    txt��λ���� = 10
    txt��λ�绰 = 11
    txt��λ�ʱ� = 12
    txt��ϵ������ = 13
    txt��ϵ�˵绰 = 14
    txt��ϵ�˵�ַ = 15
    txt��Ժʱ�� = 16
    txt��Ժ���� = 17
    txt��Ժ���� = 18
    txt��Ժʱ�� = 19
    txt��Ժ���� = 20
    txt��Ժ���� = 21
    txtסԺ���� = 22
    txt���ʱ�� = 7
    txtת��1 = 23
    txtת��2 = 24
    txtת��3 = 25
    txt����֤�� = 37
    txt��� = 38
    txt���� = 39
    chk����Ժ = 10
    cbo��Ժ��ʽ = 49
    cboӤ�����䵥λ = 51
    txtӤ������ = 42
    txt���������� = 43
    txt��������Ժ���� = 44
    cbo31���7������Ժ = 54
End Enum
Private Enum ��ҽ���
    chk�Ƿ�ȷ�� = 0
    txt���ȴ��� = 26
    txtȷ������ = 27
    txt�ɹ����� = 28
    cbo�������Ժ = 36
    cbo��Ժ���Ժ = 35
    cbo��������Ժ = 58
    cbo�����벡�� = 34
    cbo�ٴ��벡�� = 33
    cbo�ٴ���ʬ�� = 32
    cbo��ǰ������ = 31
    txt����ԭ�� = 50
    cbo�ֻ��̶� = 52
    cbo���������� = 53
    lbl�ֻ��̶� = 103
    lbl���������� = 104
    txt����� = 57
End Enum
Private Enum ��ҽ���
    chkΣ�� = 2
    chk��֢ = 3
    chk���� = 4
    cbo��֤ = 2
    cbo�η� = 11
    cbo��ҩ = 12
    cbo������ҩ = 13
    cbo���ȷ��� = 14
    cbo������� = 15
    cbo��ҽ�������Ժ = 38
    cbo��ҽ��Ժ���Ժ = 37
    cboʹ����ҽ�����豸 = 55
    cboʹ����ҽ���Ƽ��� = 56
    cbo��֤ʩ�� = 57
End Enum
Private Enum ����������
    cboHBsAg = 29
    cboHCVAb = 28
    cboHIVAb = 27
    chk��������¼�� = 19
End Enum
Private Enum סԺ���
    chk�·����� = 5
    chkʬ�� = 6
    chk���� = 7
    chkʾ�̲��� = 8
    chk���в��� = 11
    chk���Ѳ��� = 20
    chk����Ժ���� = 15
    txt����ԭ�� = 4
    txt�������� = 29
    txt���ϸ�� = 30
    txt��ѪС�� = 31
    txt��Ѫ�� = 32
    txt��ȫѪ = 33
    txt������ = 34
    txtҽѧ��ʾ = 66
    txt����ҽѧ��ʾ = 65
    cbo��Һ��Ӧ = 30
    cbo����Ex = 16
    cboѪ�� = 17
    cboRh = 18
    cbo����ҽʦ = 19
    cbo������ = 20
    cbo����ҽʦ = 21
    cbo����ҽʦ = 22
    cboסԺҽʦ = 23
    cbo����ҽʦ = 24
    cbo�о���ҽʦ = 25
    cboʵϰҽʦ = 26
    cbo�ʿػ�ʿ = 39
    cbo�ʿ�ҽʦ = 40
    cbo��Ѫ��Ӧ = 41
    cbo��Ժ��ʽ = 44
    txt��Ժת�� = 40
    lblת��ȥ�� = 88
    txt31��Ŀ�� = 41
    opt31���� = 6
    opt31���� = 7
    cbo���λ�ʿ = 50
    txt��Ժǰ�� = 56
    txt��Ժ���� = 55
    txt��ԺǰСʱ = 45
    txt��Ժǰ���� = 46
    txt��Ժ��Сʱ = 47
    txt��Ժ����� = 48
    txt������Сʱ = 49
    txt���� = 52
    txt���ڵ�ַ = 53
    txt�����ʱ� = 54
    cbo�������� = 59
    txt�ʿ����� = 60
    cbo����״�� = 60
End Enum
Private Enum ��������
    chk��ԭѧ = 9
    txt��ԭѧ = 35
    lbl��ԭѧ = 61
    cbo��Ѫ��� = 42
    cbo�������� = 43
    chkCT = 12
    chkMRI = 13
    chk������ = 14
    picѹ�� = 0
    pic������׹�� = 1
    cboѹ�������ڼ� = 45
    cboѹ������ = 46
    cbo������׹���˺� = 47
    cbo������׹��ԭ�� = 48
    chkסԺ�ڼ�没�ػ�Σ = 1
    chk����·�� = 16
    chk���·�� = 17
    chk���� = 18
    txt�˳�ԭ�� = 61
    txt����ԭ�� = 62
    chk�Ƿ�ʹ������Լ�� = 21
    txtԼ����ʱ�� = 58
    cboԼ����ʽ = 63
    cboԼ������ = 62
    cboԼ��ԭ�� = 61
    cbo��������Ժ��ʽ = 64
    chkΧ�������� = 22
    chk������� = 23
    cbo�ط����ʱ�� = 65
    chk�˹������ѳ� = 24
    chk�ط���֢ҽѧ�� = 25
    txt��֢�໤�� = 59
End Enum
Private Enum ǩ������
    cmd������ = 0
    cmd����ҽʦ = 1
    cmd����ҽʦ = 2
    cmdסԺҽʦ = 3
End Enum
Private Enum ������
    kss���� = 1
    kss��ҩĿ�� = 2
    kssʹ�ý׶� = 3
    kssʹ������ = 4
    KSSһ���п�Ԥ���� = 5
    KSSDDD�� = 6
    KSS������ҩ = 7
End Enum

Private Enum AllerColS
    AC_����ʱ�� = 0
    AC_����ҩ�� = 1
    AC_������Ӧ = 2
End Enum

Private Enum �ɵĵǼ���
    txt������� = 51
End Enum


Private Const ColorUnEditCell = &H8000000B  '����ɫ

Public Function EditMedicalRecord(ByVal lng����ID As Long, ByVal lng��ҳID As Long, ByVal lng����ID As Long, _
                                ByVal lngPathState As Long, ByVal strPrivs As String, frmParent As Object, ByVal blnModal As Boolean, _
                                Optional ByVal str���� As String, Optional blnDiagnose As Boolean, Optional ByVal blnReadOnly As Boolean, _
                                Optional ByRef str����ID As String, Optional ByRef str���ID As String, Optional ByVal bln��ʿվ As Boolean) As Boolean

'������str����=Ҫʾ¼���������ͣ���"3,13"��ʽ
'      blnDiagnose=Ҫ��¼����ϣ���ȱʡ��λ�����
'���أ�blnDiagnose=�Ƿ�¼����ָ�����͵����
    mstrPrivs = strPrivs
    mlng����ID = lng����ID
    mlng��ҳID = lng��ҳID
    mlng����ID = lng����ID
    mlngPathState = lngPathState
    mblnDiagChange = False
    mstrXYDiagInfo = ""
    mstrZYDiagInfo = ""
    
    mstr���� = str����
    mblnDiagnose = blnDiagnose
    mblnReadOnly = blnReadOnly
    mbln��ʿվ = bln��ʿվ
    
    mstr����ID = ""
    mstr���ID = ""

    On Error Resume Next
    If blnModal Then
        Me.Show 1, frmParent
        blnDiagnose = mblnDiagnose
        EditMedicalRecord = mblnOk
        str����ID = mstr����ID
        str���ID = mstr���ID

    Else
        Me.Show , frmParent
    End If
End Function

Public Property Let Opened(ByVal vData As Boolean)
    mblnOpen = vData
End Property

Public Property Get Opened() As Boolean
    Opened = mblnOpen
End Property

Private Sub cboInfo_Change(Index As Integer)
    If cboinfo(Index).Style = 0 Then
        If Visible Then mblnChange = True
    End If
End Sub

Private Sub cboInfo_Click(Index As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String, strTmp As String
    Dim vRect As RECT, blnCancel As Boolean
    Dim intIdx As Integer
    
    On Local Error Resume Next
    
    If Visible Then mblnChange = True
    
    If cboinfo(Index).ItemData(cboinfo(Index).ListIndex) = -1 And Visible Then
        'ѡ����������
        If Index = cbo����ҽʦ Or Index = cbo������ Or Index = cbo����ҽʦ Or Index = cbo����ҽʦ Or Index = cboסԺҽʦ _
            Or Index = cboʵϰҽʦ Or Index = cbo����ҽʦ Or Index = cbo�о���ҽʦ Or Index = cbo�ʿ�ҽʦ Or Index = cbo�ʿػ�ʿ Or Index = cbo���λ�ʿ Then
            
            StrSQL = mcol��ԱSQL("_" & Index)

            vRect = GetControlRect(cboinfo(Index).hwnd)
            Set rsTmp = zlDatabase.ShowSelect(Me, StrSQL, 0, "ҽ����ʿ", , , , , , True, vRect.Left, vRect.Top, cboinfo(Index).Height, blnCancel, , True)
            If Not rsTmp Is Nothing Then
                intIdx = SeekCboIndex(cboinfo(Index), rsTmp!ID)
                If intIdx <> -1 Then
                    cboinfo(Index).ListIndex = intIdx
                Else
                    cboinfo(Index).AddItem rsTmp!����, cboinfo(Index).ListCount - 1
                    cboinfo(Index).ItemData(cboinfo(Index).NewIndex) = rsTmp!ID
                    cboinfo(Index).ListIndex = cboinfo(Index).NewIndex
                End If
            Else
                If Not blnCancel Then
                    MsgBox "û��סԺҽ����ʿ�����ݣ����ȵ�����/��Ա���������á�", vbInformation, gstrSysName
                End If
                '�ָ������е���Ա(������Click)
                intIdx = SeekCboIndex(cboinfo(Index), cboinfo(Index).Tag)
                Call zlControl.CboSetIndex(cboinfo(Index).hwnd, intIdx)
            End If
        End If
    Else
        cboinfo(Index).Tag = cboinfo(Index).Text
    End If
    
    If Index = cbo������ Or Index = cbo����ҽʦ Or Index = cbo����ҽʦ Or Index = cboסԺҽʦ Then
        'ҽʦ����,ˢ��ǩ��״̬
        If Visible Then
            mblnReadOnly = SetSignature(False)
            Call SetFaceEditable(mblnReadOnly)
        End If
    ElseIf Index = cbo����Ex Then
        If cboinfo(Index).Text = "����" Then
            txtInfo(txt��������).Text = ""
            txtInfo(txt��������).Locked = True
            txtInfo(txt��������).TabStop = False
            txtInfo(txt��������).BackColor = vbButtonFace
        Else
            txtInfo(txt��������).Locked = False
            txtInfo(txt��������).TabStop = True
            txtInfo(txt��������).BackColor = vbWindowBackground
            If Visible Then txtInfo(txt��������).SetFocus
        End If
    ElseIf Index = cbo��Ժ��ʽ Then
        If cboinfo(Index).Text = "תԺ" Or cboinfo(Index).Text = "ת����" Then
            txtInfo(txt��Ժת��).Enabled = True
            lblInfo(lblת��ȥ��).Enabled = True
            txtInfo(txt��Ժת��).TabStop = True
            txtInfo(txt��Ժת��).BackColor = vbWindowBackground
        Else
            txtInfo(txt��Ժת��).Enabled = False
            lblInfo(lblת��ȥ��).Enabled = False
            txtInfo(txt��Ժת��).TabStop = False
            txtInfo(txt��Ժת��).BackColor = vbButtonFace
        End If
    End If
End Sub

Private Sub cboInfo_GotFocus(Index As Integer)
    If cboinfo(Index).Style = 0 Then
        Call zlControl.TxtSelAll(cboinfo(Index))
    End If
End Sub

Private Sub cboInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyDelete Then
        If cboinfo(Index).Style = 2 And cboinfo(Index).ListIndex <> -1 Then
            cboinfo(Index).ListIndex = -1
        End If
    End If
End Sub

Private Sub cboInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim lngidx As Long
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Index = cbo��֤ʩ�� Then
            If sstInfo.TabVisible(sstInfo.Tab + IIf(sstInfo.TabVisible(sstInfo.Tab + 1), 1, 2)) Then sstInfo.Tab = sstInfo.Tab + IIf(sstInfo.TabVisible(sstInfo.Tab + 1), 1, 2)
            Call sstInfo_KeyPress(13)
        ElseIf Index = cbo���λ�ʿ Then
            If sstInfo.TabVisible(sstInfo.Tab + IIf(sstInfo.TabVisible(sstInfo.Tab + 1), 1, 2)) Then sstInfo.Tab = sstInfo.Tab + IIf(sstInfo.TabVisible(sstInfo.Tab + 1), 1, 2)
            Call sstInfo_KeyPress(13)
        Else
            Call zlCommFun.PressKey(vbKeyTab)
            If Index = cbo�ط����ʱ�� Then
                If vsfMain.Rows = 1 Then Call zlCommFun.PressKey(vbKeyTab)
            End If
        End If
    ElseIf KeyAscii >= 32 Then
        If Index = cbo���֤�� Then
            '�������볤��
            If zlCommFun.ActualLen(cboinfo(Index).Text) > 18 Then
                KeyAscii = 0: Exit Sub
            End If
            
            '������������
            KeyAscii = Asc(UCase(Chr(KeyAscii)))
            If InStr("1234567890ABCDEFGHIJKLMNOPQRSTUVWXYZ", Chr(KeyAscii)) = 0 Then
                KeyAscii = 0: Exit Sub
            End If
        ElseIf Not cboinfo(Index).Locked And cboinfo(Index).Style = 2 Then
            lngidx = zlControl.CboMatchIndex(cboinfo(Index).hwnd, KeyAscii)
            If lngidx = -1 And cboinfo(Index).ListCount > 0 Then lngidx = 0
            cboinfo(Index).ListIndex = lngidx
        End If
    End If
End Sub

Private Sub cboInfo_LostFocus(Index As Integer)
    Dim strTmp As String, strMsg As String
    
    On Local Error Resume Next
    
    If Index = cbo���䵥λ Then
        If IsNumeric(txtInfo(txt����).Text) And cboinfo(cbo���䵥λ).ListIndex <> -1 Then
            Select Case cboinfo(cbo���䵥λ).Text
                Case "��"
                    If Val(txtInfo(txt����).Text) > 200 Then
                        MsgBox "����ֵ�������������Ƿ���ȷ��", vbInformation, gstrSysName
                        txtInfo(txt����).SetFocus: Exit Sub
                    End If
                Case "��"
                    If Val(txtInfo(txt����).Text) > 2400 Then
                        MsgBox "����ֵ�������������Ƿ���ȷ��", vbInformation, gstrSysName
                        txtInfo(txt����).SetFocus: Exit Sub
                    End If
                Case "��"
                    If Val(txtInfo(txt����).Text) > 73000 Then
                        MsgBox "����ֵ�������������Ƿ���ȷ��", vbInformation, gstrSysName
                        txtInfo(txt����).SetFocus: Exit Sub
                    End If
                Case Else
                    Exit Sub
            End Select
            
            '�������������ڣ�Сʱ���������λ�����з������飩
            If cboinfo(cbo���䵥λ).ListIndex < 3 Then
                If Not IsDate(txt��������.Text) Then
                    txt��������.Text = ReCalcBirth(txtInfo(txt����).Text, cboinfo(cbo���䵥λ).Text)
                Else
                    strTmp = PatiAgeCalc(txt��������.Text, , txtInfo(txt��Ժʱ��).Text)
                    If Right(strTmp, 1) = cboinfo(cbo���䵥λ).Text And IsNumeric(Left(strTmp, Len(strTmp) - 1)) _
                        And strTmp <> txtInfo(txt����).Text & cboinfo(cbo���䵥λ).Text Then
                        
                        strMsg = zlCommFun.ShowMsgBox(gstrSysName, "����ͳ������ڲ�һ�£�" & txt��������.Text & "��������Ӧ����" & strTmp & "��" & _
                            vbCrLf & vbCrLf & "���������������ڵ���ȷ�ԣ���ѡ��������Ӧ�Ĳ�����", "!��������(&R),����(&A),?ȡ��(&C)", Me, vbQuestion)
                        If strMsg = "��������" Then
                            txt��������.Text = ReCalcBirth(txtInfo(txt����).Text, cboinfo(cbo���䵥λ).Text)
                            txt����ʱ��.Text = "__:__"
                        ElseIf strMsg = "����" Then
                        Else
                            txtInfo(txt����).SetFocus: Exit Sub
                        End If
                    End If
                End If
            End If
        End If
    End If
End Sub

Private Sub cboInfo_Validate(Index As Integer, Cancel As Boolean)
'���ܣ��������������,�Զ�ƥ��ִ�п���
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String, intIdx As Long, i As Long
    Dim strInput As String
    Dim vRect As RECT, blnCancel As Boolean
        
    If Index = cbo����ҽʦ Or Index = cbo������ Or Index = cbo����ҽʦ Or Index = cbo����ҽʦ Or Index = cboסԺҽʦ _
        Or Index = cboʵϰҽʦ Or Index = cbo����ҽʦ Or Index = cbo�о���ҽʦ Or Index = cbo�ʿ�ҽʦ Or Index = cbo�ʿػ�ʿ Then
        If cboinfo(Index).ListIndex <> -1 Then Exit Sub '��ѡ��
        If cboinfo(Index).Text = "" Then cboinfo(Index).Tag = "": Exit Sub '������
        
        strInput = UCase(NeedName(cboinfo(Index).Text))
        StrSQL = mcol��ԱSQL("_" & Index)
        StrSQL = Replace(UCase(StrSQL), UCase("Order by"), " And (A.��� Like [1] Or A.���� Like [2] Or A.���� Like [2]) Order by")
        
        On Error GoTo errH
        vRect = GetControlRect(cboinfo(Index).hwnd)
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 0, "ҽ����ʿ", False, "", "", False, False, _
            True, vRect.Left, vRect.Top, cboinfo(Index).Height, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%")
        If Not rsTmp Is Nothing Then
            intIdx = SeekCboIndex(cboinfo(Index), rsTmp!ID)
            If intIdx <> -1 Then
                cboinfo(Index).ListIndex = intIdx
            Else
                cboinfo(Index).AddItem rsTmp!����, cboinfo(Index).ListCount - 1
                cboinfo(Index).ItemData(cboinfo(Index).NewIndex) = rsTmp!ID
                cboinfo(Index).ListIndex = cboinfo(Index).NewIndex
            End If
        Else
            If Not blnCancel Then
                MsgBox "δ�ҵ���Ӧ��ҽ����ʿ��", vbInformation, gstrSysName
            End If
            Cancel = True: Exit Sub
        End If
    ElseIf Index = cbo��ǰ������ Then
        If cboinfo(cbo��ǰ������).ListIndex = 0 And vsOPS.TextMatrix(1, col��������) <> "" Then
            '��������������Ͳ�����ѡδ��
            cboinfo(cbo��ǰ������).ListIndex = 1
        End If
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub chkInfo_Click(Index As Integer)
    If mblnNoClick Then Exit Sub
    If Visible And mblnReadOnly Then
        mblnNoClick = True
        chkInfo(Index).Value = IIf(chkInfo(Index).Value = 1, 0, 1)
        mblnNoClick = False: Exit Sub
    End If
    
    Select Case Index
        Case chk�Ƿ�ȷ��
            If chkInfo(Index).Value = 1 Then
                txtInfo(txtȷ������).Locked = False
                txtInfo(txtȷ������).TabStop = True
                txtInfo(txtȷ������).BackColor = vbWindowBackground
            Else
                txtInfo(txtȷ������).Text = ""
                txtInfo(txtȷ������).Locked = True
                txtInfo(txtȷ������).TabStop = False
                txtInfo(txtȷ������).BackColor = vbButtonFace
            End If
        Case chk����
            If chkInfo(Index).Value = 1 Then
                txtInfo(txt��������).Locked = False
                txtInfo(txt��������).TabStop = True
                txtInfo(txt��������).BackColor = vbWindowBackground
                cboinfo(cbo����Ex).Locked = False
                cboinfo(cbo����Ex).TabStop = True
                cboinfo(cbo����Ex).BackColor = vbWindowBackground
                
                Call cboInfo_Click(cbo����Ex)
            Else
                txtInfo(txt��������).Text = ""
                txtInfo(txt��������).Locked = True
                txtInfo(txt��������).TabStop = False
                txtInfo(txt��������).BackColor = vbButtonFace
                cboinfo(cbo����Ex).Locked = True
                cboinfo(cbo����Ex).TabStop = False
                cboinfo(cbo����Ex).BackColor = vbButtonFace
            End If
        Case chk����·��
            If chkInfo(Index).Value = 0 Then
                chkInfo(chk���·��).Enabled = False
                chkInfo(chk���·��).Value = 0
                txtInfo(txt�˳�ԭ��).Enabled = False
                txtInfo(txt�˳�ԭ��).Text = ""
                txtInfo(txt�˳�ԭ��).BackColor = vbButtonFace
                chkInfo(chk����).Enabled = False
                chkInfo(chk����).Value = 0
                txtInfo(txt����ԭ��).Enabled = False
                txtInfo(txt����ԭ��).Text = ""
                txtInfo(txt����ԭ��).BackColor = vbButtonFace
            Else
                chkInfo(chk���·��).Enabled = True
                chkInfo(chk���·��).TabStop = True
                txtInfo(txt�˳�ԭ��).TabStop = True
                txtInfo(txt�˳�ԭ��).Enabled = True
                txtInfo(txt�˳�ԭ��).BackColor = vbWindowBackground
                chkInfo(chk����).Enabled = True
                chkInfo(chk����).TabStop = True
                chkInfo(Index).TabStop = True
            End If
        Case chk���·��
            If chkInfo(Index).Value = 0 Then
                txtInfo(txt�˳�ԭ��).Enabled = True
                txtInfo(txt�˳�ԭ��).TabStop = True
                txtInfo(txt�˳�ԭ��).BackColor = vbWindowBackground
            Else
                txtInfo(txt�˳�ԭ��).Enabled = False
                txtInfo(txt�˳�ԭ��).Text = ""
                txtInfo(txt�˳�ԭ��).BackColor = vbButtonFace
                chkInfo(Index).TabStop = True
            End If
        Case chk����
            If chkInfo(Index).Value = 1 Then
                txtInfo(txt����ԭ��).Enabled = True
                txtInfo(txt����ԭ��).TabStop = True
                txtInfo(txt����ԭ��).BackColor = vbWindowBackground
                chkInfo(Index).TabStop = True
            Else
                txtInfo(txt����ԭ��).Enabled = False
                txtInfo(txt����ԭ��).Text = ""
                txtInfo(txt����ԭ��).BackColor = vbButtonFace
            End If
        Case chkʬ��
            '������Ϸ������
            If Visible Then Call Set��Ϸ������(cbo�ٴ���ʬ��)
        Case chk��ԭѧ
            If chkInfo(chk��ԭѧ).Value = 0 Then
                txtInfo(txt��ԭѧ).Text = ""
                txtInfo(txt��ԭѧ).Tag = ""
                cmdInfo(txt��ԭѧ).Tag = ""
                txtInfo(txt��ԭѧ).Enabled = False
                cmdInfo(txt��ԭѧ).Enabled = False
                lblInfo(lbl��ԭѧ).ForeColor = &H808080
            ElseIf Not txtInfo(txt��ԭѧ).Enabled Then
                txtInfo(txt��ԭѧ).Enabled = True
                cmdInfo(txt��ԭѧ).Enabled = True
                lblInfo(lbl��ԭѧ).ForeColor = Me.ForeColor
            End If
        Case chk�Ƿ�ʹ������Լ��
            If chkInfo(Index).Value = 0 Then
                txtInfo(txtԼ����ʱ��).Text = ""
                cboinfo(cboԼ����ʽ).ListIndex = -1
                cboinfo(cboԼ������).ListIndex = -1
                cboinfo(cboԼ��ԭ��).ListIndex = -1
                txtInfo(txtԼ����ʱ��).BackColor = vbButtonFace
                cboinfo(cboԼ����ʽ).BackColor = vbButtonFace
                cboinfo(cboԼ������).BackColor = vbButtonFace
                cboinfo(cboԼ��ԭ��).BackColor = vbButtonFace
                txtInfo(txtԼ����ʱ��).Enabled = False
                cboinfo(cboԼ����ʽ).Enabled = False
                cboinfo(cboԼ������).Enabled = False
                cboinfo(cboԼ��ԭ��).Enabled = False
            Else
                txtInfo(txtԼ����ʱ��).BackColor = vbWindowBackground
                cboinfo(cboԼ����ʽ).BackColor = vbWindowBackground
                cboinfo(cboԼ������).BackColor = vbWindowBackground
                cboinfo(cboԼ��ԭ��).BackColor = vbWindowBackground
                txtInfo(txtԼ����ʱ��).Enabled = True
                cboinfo(cboԼ����ʽ).Enabled = True
                cboinfo(cboԼ������).Enabled = True
                cboinfo(cboԼ��ԭ��).Enabled = True
            End If
    End Select
    If Visible Then mblnChange = True
End Sub

Private Sub chkInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        If Index = chk������� Or Index = chkסԺ�ڼ�没�ػ�Σ Then
            If sstInfo.TabVisible(sstInfo.Tab + IIf(sstInfo.TabVisible(sstInfo.Tab + 1), 1, 2)) Then sstInfo.Tab = sstInfo.Tab + IIf(sstInfo.TabVisible(sstInfo.Tab + 1), 1, 2)
            Call sstInfo_KeyPress(13)
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End If
End Sub

Private Function GetStage(ByVal DateUseBegin As Date, ByVal DateUseEnd As Date, ByVal DateSs As Date, ByVal strTime As String) As String
'���ܣ���ÿ�����ʹ�ý׶�
'������DateUseBegin ʹ��ʱ��,DateUseEnd -����ʱ��  DateSs-����ʱ��,strTime ��һ�ε�ʹ�ý׶�
    
    '���û���������򷵻ؿ�
    If DateSs = 0 Then GetStage = " ": Exit Function
    '����Ѿ���Χ�����ڣ�ֱ���˳�
    If strTime = "Χ������" Then GetStage = strTime: Exit Function
    
    If DateUseBegin < DateSs Then
        If DateUseEnd < DateSs Then
            If strTime <> "" Then
                If strTime <> "��ǰ" Then strTime = "Χ������"
            Else
                strTime = "��ǰ"
            End If
        Else
            strTime = "Χ������"
        End If
    ElseIf DateUseBegin > DateSs Then
        If DateUseEnd > DateSs Then
            If strTime <> "" Then
                If strTime <> "����" Then strTime = "Χ������"
            Else
                strTime = "����"
            End If
        Else
            strTime = "Χ������"
        End If
    Else
        If DateUseEnd = DateSs Then
            If strTime <> "" Then
                If strTime <> "����" Then strTime = "Χ������"
            Else
                strTime = "����"
            End If
        Else
            strTime = "Χ������"
        End If
    End If
    GetStage = strTime
End Function

Private Sub cmdAutoLoad_Click(Index As Integer)
    '�Զ���ȡ
    Dim StrSQL As String, rsTmp As Recordset
    Dim i As Long, j As Long
    Dim blnAgain As Boolean, blnIsNull As Boolean
    Dim DateSs As Date          '�ò������������ʱ��
    Dim rsTime As New ADODB.Recordset
    Dim blnStage As Boolean
    Dim strOld���� As String
    Dim lngRow As Long
    Dim blnClear As Boolean
    Dim strPrivs As String
    
    On Error GoTo errH
    Select Case Index
        Case 0
            StrSQL = "Select Min(NVL(to_date(c.�걾��λ,'yyyy-mm-dd hh24:mi:ss'),c.��ʼִ��ʱ��)) as ʹ��ʱ��" & vbNewLine & _
                    " From ������ĿĿ¼ A, ����ҽ����¼ C" & vbNewLine & _
                    " Where  a.Id = c.������Ŀid and a.���='F' And c.����id = [1] And c.��ҳid = [2] And c.ҽ��״̬=8"
        
            Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
            If rsTmp.RecordCount > 0 Then DateSs = CDate(Format(Nvl(rsTmp!ʹ��ʱ��, 0), "yyyy-MM-dd"))
            
            StrSQL = "Select distinct ID, ҽ��id, �ϼ�id, ����, ����, ��λ, ִ��ʱ�䷽��, Ƶ�ʼ��, �����λ, Ƶ�ʴ���, �ϴ�ִ��ʱ��, ��ʼִ��ʱ��, ����ʱ��," & vbNewLine & _
                    "       Sum(Ddd��) Over(Partition By ID) As Ddd��, Count(1) Over(Partition By ���id) As ������ҩ" & vbNewLine & _
                    "From   (Select Distinct ID, ҽ��id, �ϼ�id, ����, ����, ��λ, ִ��ʱ�䷽��, Ƶ�ʼ��, �����λ, Ƶ�ʴ���, �ϴ�ִ��ʱ��, ��ʼִ��ʱ��, ����ʱ��," & vbNewLine & _
                    "                Sum(����) Over(Partition By ID, ҽ��id, ���id) * ����ϵ�� / Decode(Dddֵ, 0, Null, Dddֵ) As Ddd��, ���id" & vbNewLine & _
                    "         From   (Select z.Id, a.Id As ҽ��id, z.����id As �ϼ�id, z.����, z.����, z.���㵥λ As ��λ, a.ִ��ʱ�䷽��, a.Ƶ�ʼ��, a.�����λ, a.Ƶ�ʴ���," & vbNewLine & _
                    "                         a.�ϴ�ִ��ʱ��, a.��ʼִ��ʱ��, Nvl(a.�ϴ�ִ��ʱ��, Nvl(a.ִ����ֹʱ��, a.��ʼִ��ʱ��)) As ����ʱ��, a.���id, f.����, h.����ϵ��," & vbNewLine & _
                    "                         Nvl((Select e.Dddֵ From �����÷����� E Where e.��Ŀid = a.������Ŀid And e.�÷�id = r.������Ŀid), h.Dddֵ) As Dddֵ" & vbNewLine & _
                    "                  From   ����ҽ����¼ A, ����ҽ����¼ R, סԺ���ü�¼ F, ҩƷ��� H, ҩƷ���� B, ������ĿĿ¼ Z" & vbNewLine & _
                    "                  Where  a.������Ŀid = b.ҩ��id And a.������� In ('5', '6') And" & vbNewLine & _
                    "                         (a.ҽ����Ч = 0 And a.�ϴ�ִ��ʱ�� Is Not Null Or a.ҽ����Ч = 1 And a.ҽ��״̬ = 8) And Nvl(b.������, 0) <> 0 And" & vbNewLine & _
                    "                         a.���id = r.Id And a.Id = f.ҽ����� And f.��¼״̬ <> 0 And f.�շ�ϸĿid = h.ҩƷid And b.ҩ��id = z.Id And" & vbNewLine & _
                    "                         a.����id = [1] And a.��ҳid = [2]))" & vbNewLine & _
                    "Order  By Ddd�� Desc"
            Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
            
            rsTime.Fields.Append "�շ�ʱ��", adVarChar, 10
            rsTime.Fields.Append "ҩƷID", adBigInt
            rsTime.CursorLocation = adUseClient
            rsTime.LockType = adLockOptimistic
            rsTime.CursorType = adOpenStatic
            rsTime.Open
                                
            With vsKSS
                
                If rsTmp.RecordCount = 0 Then MsgBox "û���ҵ��ò��˵Ŀ���ҩ��ʹ�ü�¼��", vbInformation, Me.Caption
                Do Until rsTmp.EOF
                    
                    For i = .FixedRows To .Rows - 1
                        '�ж��Ƿ����ظ���
                        If rsTmp!ID = Val(.RowData(i) & "") Then
                            '����ǿղ���ȡ������������ǿգ��򲻸ı��û����ֶ�ѡ��
                            If .Cell(flexcpData, i, kssʹ�ý׶�) = "����" Or Trim(.TextMatrix(i, kssʹ�ý׶�)) = "" Then
                                blnStage = True
                            End If
                            lngRow = i
                            If .TextMatrix(i, KSSDDD��) = "" Then .TextMatrix(i, KSSDDD��) = FormatEx(Val(rsTmp!DDD�� & ""), 2)
                            If Decode(.TextMatrix(i, KSS������ҩ), "����", 1, "����", 2, "����", 3, "����", 4, ">����", 999, 0) < Val(rsTmp!������ҩ & "") Then
                                .TextMatrix(i, KSS������ҩ) = Decode(Val(rsTmp!������ҩ & ""), 1, "����", 2, "����", 3, "����", 4, "����", ">����")
                            End If
                            Exit For
                        End If
                        '�ж�����֮ǰ������û�пյ�
                        If .TextMatrix(i, kss����) & "" = "" Then
                            .TextMatrix(i, kss����) = rsTmp!���� & ""
                            .RowData(i) = Val(rsTmp!ID)
                            .Cell(flexcpData, i, kss����) = rsTmp!���� & ""
                            .TextMatrix(i, KSSDDD��) = FormatEx(Val(rsTmp!DDD�� & ""), 2)
                            .Cell(flexcpData, i, KSSDDD��) = .TextMatrix(i, KSSDDD��)
                            If Decode(.TextMatrix(i, KSS������ҩ), "����", 1, "����", 2, "����", 3, "����", 4, ">����", 999, 0) < Val(rsTmp!������ҩ & "") Then
                                .TextMatrix(i, KSS������ҩ) = Decode(Val(rsTmp!������ҩ & ""), 1, "����", 2, "����", 3, "����", 4, "����", ">����")
                            End If
                            lngRow = i
                            blnStage = True
                            Exit For
                        Else
                            If i = .Rows - 1 Then
                                .AddItem "": .TextMatrix(.Rows - 1, 0) = .Rows - 1
                                .TextMatrix(.Rows - 1, kss����) = rsTmp!���� & ""
                                .RowData(.Rows - 1) = Val(rsTmp!ID)
                                .Cell(flexcpData, .Rows - 1, kss����) = rsTmp!���� & ""
                                .TextMatrix(i, KSSDDD��) = FormatEx(Val(rsTmp!DDD�� & ""), 2)
                                .Cell(flexcpData, i, KSSDDD��) = .TextMatrix(i, KSSDDD��)
                                If Decode(.TextMatrix(i, KSS������ҩ), "����", 1, "����", 2, "����", 3, "����", 4, ">����", 999, 0) < Val(rsTmp!������ҩ & "") Then
                                    .TextMatrix(i, KSS������ҩ) = Decode(Val(rsTmp!������ҩ & ""), 1, "����", 2, "����", 3, "����", 4, "����", ">����")
                                End If
                                lngRow = .Rows - 1
                                blnStage = True
                                Exit For
                            End If
                        End If
                    Next
                    If blnStage Then
                        'ʹ�ý׶�
                        .TextMatrix(lngRow, kssʹ�ý׶�) = _
                            GetStage(CDate(Format(rsTmp!��ʼִ��ʱ�� & "", "yyyy-MM-dd")), CDate(Format(rsTmp!����ʱ�� & "", "yyyy-MM-dd")), DateSs, Trim(.TextMatrix(lngRow, kssʹ�ý׶�)))
                        .Cell(flexcpData, lngRow, kssʹ�ý׶�) = "����"
                        vsKSS.Tag = "": mblnChange = True
                    End If
                    strOld���� = Trim(.TextMatrix(lngRow, kssʹ������))
                        'ʹ������
                    .TextMatrix(lngRow, kssʹ������) = GetUseDay(Val(rsTmp!ҽ��ID), Val(.RowData(lngRow)), Nvl(rsTmp!ִ��ʱ�䷽��) & "", CDate(rsTmp!��ʼִ��ʱ��), CDate(rsTmp!����ʱ��), _
                                Nvl(rsTmp!Ƶ�ʴ���, 0), Nvl(rsTmp!Ƶ�ʼ��, 0), Nvl(rsTmp!�����λ), rsTime) & ""
                    If strOld���� <> Trim(.TextMatrix(lngRow, kssʹ������)) Then
                        .Cell(flexcpData, lngRow, kssʹ������) = "����"
                        vsKSS.Tag = "": mblnChange = True
                    End If
                    
                    rsTmp.MoveNext
                    blnStage = False
                    lngRow = 0
                Loop
            End With
        Case 1
            strPrivs = GetInsidePrivs(p����ӿ�, , 2400)
            If InStr(strPrivs, "�ڲ��ӿ�") > 0 Then
                Set rsTmp = AutoGetOPSInfo(True, mlng����ID, mlng��ҳID)
            Else
                If gbln������ȡ���� Then
                    If MsgBox("������û�С��������ϵͳ-����ӿڹ���ģ����ڲ��ӿ�Ȩ�ޣ�" & vbCrLf & "ϵͳĬ�ϴ�ҽ��ϵͳ�ж�ȡ������Ϣ���Ƿ���� ? " & vbCrLf & "ѡ�������´ν�������ʾ��", vbInformation + vbYesNo + vbDefaultButton1, Me.Caption) = vbYes Then
                        gbln������ȡ���� = False
                    Else
                        Exit Sub
                    End If
                End If
                Set rsTmp = AutoGetOPSInfo(False, mlng����ID, mlng��ҳID)
            End If
            If Not rsTmp.EOF Then
                If MsgBox("�Ƿ����ԭ�е�������Ϣ��", vbInformation + vbYesNo + vbDefaultButton2, Me.Caption) = vbYes Then
                    blnClear = True
                End If
        
                With vsOPS
                    
                    If blnClear Then
                        .Rows = .FixedRows
                        .Rows = .FixedRows + rsTmp.RecordCount + 1
                        lngRow = .FixedRows
                    Else
                        If .Rows > .FixedRows Then
                            If .TextMatrix(.Rows - 1, col��������) = "" Then
                                lngRow = .Rows - 1
                                .Rows = .Rows + rsTmp.RecordCount
                            Else
                                lngRow = .Rows
                                .Rows = .Rows + rsTmp.RecordCount + 1
                            End If
                        End If
                    End If
                    
                    rsTmp.MoveFirst
                    
                    For i = lngRow To lngRow + rsTmp.RecordCount - 1
                        .TextMatrix(i, col��������) = Format(Nvl(rsTmp!��������), "yyyy-MM-dd")
                        .TextMatrix(i, col��������) = Nvl(rsTmp!��������)
                        .TextMatrix(i, col��������) = Nvl(rsTmp!��������)
                        .TextMatrix(i, col����ҽʦ) = Nvl(rsTmp!����ҽʦ)
                        .TextMatrix(i, col������ʿ) = Nvl(rsTmp!������ʿ)
                        .TextMatrix(i, col����1) = Nvl(rsTmp!��һ����)
                        .TextMatrix(i, col����2) = Nvl(rsTmp!�ڶ�����)
                        .TextMatrix(i, col����ʽ) = GetItemField("������ĿĿ¼", Val(Nvl(rsTmp!����ʽ, 0)), "����")
                        .TextMatrix(i, col����ҽʦ) = Nvl(rsTmp!����ҽʦ)
                        If Not IsNull(rsTmp!�п�) And Not IsNull(rsTmp!����) Then
                            .TextMatrix(i, col�п�����) = rsTmp!�п� & "/" & rsTmp!����
                        End If
                        .TextMatrix(i, col��������ID) = Nvl(rsTmp!��������ID)
                        .TextMatrix(i, col������ĿID) = Nvl(rsTmp!������Ŀid)
                        .TextMatrix(i, col����ID) = Nvl(rsTmp!����ʽ)
                        .TextMatrix(i, col��������) = Nvl(rsTmp!��������)
                        .TextMatrix(i, COL�������.COL�������) = Nvl(rsTmp!�������)
                        .TextMatrix(i, colASA�ּ�) = Decode(Nvl(rsTmp!asa�ּ�), "I��", "P1", "II��", "P2", "III��", "P3", "IV��", "P4", "V��", "P5", Nvl(rsTmp!asa�ּ�))
                        .TextMatrix(i, colNNIS�ּ�) = Nvl(rsTmp!NNIS�ּ�)
                        .TextMatrix(i, col��������) = Nvl(rsTmp!��������)
                        .TextMatrix(i, col�ٴ�����) = IIf(Val(rsTmp!�ٴ����� & "") = 1, -1, 0)
                        .TextMatrix(i, col����ҩ����) = rsTmp!������ҩ���� & ""
                        .Cell(flexcpChecked, i, colԤ���ÿ���ҩ) = Val(rsTmp!��ǰ������ҩ & "")
                        .Cell(flexcpChecked, i, col��Ԥ�ڵĶ�������) = Val(rsTmp!��Ԥ�ڵĶ������� & "")
                        .Cell(flexcpChecked, i, col������֢) = Val(rsTmp!������֢ & "")
                        .Cell(flexcpChecked, i, col������������) = Val(rsTmp!������������ & "")
                        .Cell(flexcpChecked, i, col��������֢) = Val(rsTmp!��������֢ & "")
                        .Cell(flexcpChecked, i, col�����Ѫ��Ѫ��) = Val(rsTmp!�����Ѫ��Ѫ�� & "")
                        .Cell(flexcpChecked, i, col�����˿��ѿ�) = Val(rsTmp!�����˿��ѿ� & "")
                        .Cell(flexcpChecked, i, col�������Ѫ˨) = Val(rsTmp!�������Ѫ˨ & "")
                        .Cell(flexcpChecked, i, col���������л����) = Val(rsTmp!���������л���� & "")
                        .Cell(flexcpChecked, i, col�������˥��) = Val(rsTmp!�������˥�� & "")
                        .Cell(flexcpChecked, i, col�����˨��) = Val(rsTmp!�����˨�� & "")
                        .Cell(flexcpChecked, i, col�����Ѫ֢) = Val(rsTmp!�����Ѫ֢ & "")
                        .Cell(flexcpChecked, i, col�����Źؽڹ���) = Val(rsTmp!�����Źؽڹ��� & "")
                        '��¼���ڱ༭�ָ�
                        For j = 0 To .Cols - 1
                            .Cell(flexcpData, i, j) = .TextMatrix(i, j)
                        Next
                        
                        rsTmp.MoveNext
                    Next
                End With
            End If
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function GetUseDay(ByVal AdviceID As Long, ByVal lngҩƷID As Long, ByVal strִ��ʱ�䷽�� As String, ByVal Date��ʼִ��ʱ�� As Date, _
            ByVal Date����ʱ�� As Date, ByVal lngƵ�ʴ��� As Long, ByVal lngƵ�ʼ�� As Long, ByVal str�����λ As String, _
            ByRef rsTime As ADODB.Recordset) As Long
'���ܣ���ȡ�����ص�ʹ������

    Dim strPause As String
    Dim j As Long
    Dim StrDecTime As String, arrDecTime As Variant
    Dim DateStart As String
    Dim strTmp As String
        
    strPause = GetAdvicePause(AdviceID)
    If strִ��ʱ�䷽�� <> "" Then
        StrDecTime = Calc���ڷֽ�ʱ��(Date��ʼִ��ʱ��, Date����ʱ��, strPause, strִ��ʱ�䷽��, lngƵ�ʴ���, lngƵ�ʼ��, str�����λ)
        arrDecTime = Split(StrDecTime, ",")
        For j = 0 To UBound(arrDecTime)
            strTmp = Format(arrDecTime(j), "yyyy-MM-dd")
            rsTime.Filter = "�շ�ʱ��='" & strTmp & "' And " & "ҩƷid=" & lngҩƷID
            If rsTime.EOF Then
                rsTime.AddNew
                rsTime!�շ�ʱ�� = strTmp
                rsTime!ҩƷid = lngҩƷID
                rsTime.Update
            End If
        Next
    Else
        DateStart = CDate(Format(Date��ʼִ��ʱ�� & "", "yyyy-MM-dd"))
        Do While DateStart <= CDate(Format(Date����ʱ�� & "", "yyyy-MM-dd"))
            rsTime.Filter = "�շ�ʱ��='" & strTmp & "' And " & "ҩƷid=" & lngҩƷID
            If rsTime.EOF Then
                rsTime.AddNew
                rsTime!�շ�ʱ�� = Format(CStr(DateStart), "yyyy-MM-dd")
                rsTime!ҩƷid = lngҩƷID
                rsTime.Update
            End If
            DateStart = CDate(DateStart) + 1
        Loop
    End If
    rsTime.Filter = "ҩƷid=" & lngҩƷID
    GetUseDay = rsTime.RecordCount
End Function

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdInfo_Click(Index As Integer)
'˵����ע�������Ҫ��CMD�Ͷ�ӦTXT��Index��ͬ
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String, blnCancel As Boolean
    Dim vPoint As POINTAPI, blnLevel As Boolean
    Dim strResult As String
    
    'ʹ��Lock�ķ�ʽ,������Enabled�ķ�ʽ
    If Not cmdInfo(Index).Enabled Or txtInfo(Index).Locked Then
        If txtInfo(Index).Enabled Then txtInfo(Index).SetFocus: Exit Sub
    End If
    
    Select Case Index
        Case txt�����ص�, txt��ͥ��ַ, txt��ϵ�˵�ַ, txt���ڵ�ַ
            'ѡ���������
            StrSQL = "Select Rownum as ID,����,����,���� From ���� Order by ����"
            vPoint = GetCoordPos(txtInfo(Index).Container.hwnd, txtInfo(Index).Left, txtInfo(Index).Top)
            Set rsTmp = zlDatabase.ShowSelect(Me, StrSQL, 0, , , , , , , True, vPoint.X, vPoint.Y, txtInfo(Index).Height, blnCancel)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "û������""����""���ݣ����ȵ��ֵ�����������á�", vbInformation, gstrSysName
                End If
                txtInfo(Index).SetFocus
            Else
                txtInfo(Index).Text = rsTmp!����
                txtInfo(Index).SetFocus
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Case txt����, txt����
            'ѡ����������
            On Error GoTo errH
            StrSQL = "Select Nvl(����,0) as ���� From ���� Group by Nvl(����,0)"
            Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption)
            If rsTmp.RecordCount > 1 Then blnLevel = True
            
            vPoint = GetCoordPos(txtInfo(Index).Container.hwnd, txtInfo(Index).Left, txtInfo(Index).Top)
            If blnLevel Then
                StrSQL = _
                    " Select ID,�ϼ�id,����,����,����,ĩ��" & _
                    " From (Select ���� As ID,RPad(Substr(����,1,Decode(Nvl(����,0),0,0,1,2,4)),6,'0') As �ϼ�id," & _
                    "       ����,����,����,Decode(Nvl(����,0),2,1,3,1,0) as ĩ��" & _
                    "       From ���� Order By ����)" & _
                    " Start With �ϼ�ID Is Null Connect By Prior ID=�ϼ�id"
                Set rsTmp = zlDatabase.ShowSelect(Me, StrSQL, 2, "����", , , , , , , vPoint.X, vPoint.Y, txtInfo(Index).Height, blnCancel)
            Else
                StrSQL = "Select Rownum as ID,����,����,���� From ���� Order by ����"
                Set rsTmp = zlDatabase.ShowSelect(Me, StrSQL, 0, , , , , , , True, vPoint.X, vPoint.Y, txtInfo(Index).Height, blnCancel)
            End If
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "û������""����""���ݣ����ȵ��ֵ�����������á�", vbInformation, gstrSysName
                End If
                txtInfo(Index).SetFocus
            Else
                txtInfo(Index).Text = rsTmp!����
                txtInfo(Index).SetFocus
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Case txt��λ����
            'ѡ��λ��Ϣ
            StrSQL = "Select ID,�ϼ�ID,ĩ��,����,����,����,��ַ,�绰,��������,�ʺ�,��ϵ��" & _
                " From ��Լ��λ" & _
                " Where (����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or ����ʱ�� is NULL)" & _
                " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID"
            vPoint = GetCoordPos(txtInfo(Index).Container.hwnd, txtInfo(Index).Left, txtInfo(Index).Top)
            Set rsTmp = zlDatabase.ShowSelect(Me, StrSQL, 2, "��Լ��λ", , , , , , True, vPoint.X, vPoint.Y, txtInfo(Index).Height, blnCancel)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "û������""��Լ��λ""���ݣ����ȵ���Լ��λ���������á�", vbInformation, gstrSysName
                End If
                txtInfo(Index).Tag = ""
                txtInfo(Index).SetFocus
            Else
                txtInfo(Index).Text = rsTmp!���� & IIf(Not IsNull(rsTmp!��ַ), "(" & rsTmp!��ַ & ")", "")
                txtInfo(Index).Tag = Val(rsTmp!ID)
                If txtInfo(txt��λ�绰).Text = "" Then
                    txtInfo(txt��λ�绰).Text = Nvl(rsTmp!�绰)
                End If
                txtInfo(Index).SetFocus
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Case txtת��1, txtת��2, txtת��3
            'ѡ��ת�ƿ���
            StrSQL = "Select Distinct A.ID,A.����,A.����,A.����,A.λ��" & _
                " From ���ű� A,��������˵�� B" & _
                " Where A.ID=B.����ID And B.������� IN(2,3) And B.�������� IN('�ٴ�','����')" & _
                " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                " Order by A.����"
            vPoint = GetCoordPos(txtInfo(Index).Container.hwnd, txtInfo(Index).Left, txtInfo(Index).Top)
            Set rsTmp = zlDatabase.ShowSelect(Me, StrSQL, 0, , , , , , , True, vPoint.X, vPoint.Y, txtInfo(Index).Height, blnCancel)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "û������""�ٴ�����""���ݣ����ȵ����Ź��������á�", vbInformation, gstrSysName
                End If
                txtInfo(Index).SetFocus
            Else
                txtInfo(Index).Text = rsTmp!����
                txtInfo(Index).SetFocus
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Case txtȷ������
            If IsDate(txtInfo(txtȷ������).Text) Then
                dtpInfo.Value = CDate(txtInfo(txtȷ������).Text)
            Else
                dtpInfo.Value = zlDatabase.Currentdate
            End If
            mlngDateIndex = Index
            dtpInfo.Left = cmdInfo(Index).Left + cmdInfo(Index).Width - dtpInfo.Width + txtInfo(Index).Container.Left + sstInfo.Left
            dtpInfo.Top = cmdInfo(Index).Top - dtpInfo.Height - 20 + txtInfo(Index).Container.Top + sstInfo.Top
            dtpInfo.ZOrder
            dtpInfo.Visible = True
            dtpInfo.SetFocus
        Case txt�ʿ�����
            If IsDate(txtInfo(txt�ʿ�����).Text) Then
                dtpInfo.Value = CDate(txtInfo(txt�ʿ�����).Text)
            Else
                dtpInfo.Value = zlDatabase.Currentdate
            End If
            mlngDateIndex = Index
            dtpInfo.Left = cmdInfo(Index).Left + cmdInfo(Index).Width - dtpInfo.Width + txtInfo(Index).Container.Left + sstInfo.Left
            dtpInfo.Top = cmdInfo(Index).Top - dtpInfo.Height - 20 + txtInfo(Index).Container.Top + sstInfo.Top
            dtpInfo.ZOrder
            dtpInfo.Visible = True
            dtpInfo.SetFocus
        Case txt��ԭѧ
            'D-ICD-10��������
            Set rsTmp = zlDatabase.ShowILLSelect(Me, "'D'", mlng����ID, cboinfo(cbo�Ա�).Text, False)
            If Not rsTmp Is Nothing Then
                txtInfo(txt��ԭѧ).Text = IIf(Not IsNull(rsTmp!����), "(" & rsTmp!���� & ")", "") & Nvl(rsTmp!����)
                txtInfo(txt��ԭѧ).Tag = txtInfo(txt��ԭѧ).Text
                cmdInfo(txt��ԭѧ).Tag = rsTmp!��ĿID
            End If
            txtInfo(txt��ԭѧ).SetFocus
        Case txt����ԭ��
             'ѡ��λ��Ϣ
            StrSQL = "Select ���� ID,����,���� From ���Ȳ������"
               
            vPoint = GetCoordPos(txtInfo(Index).Container.hwnd, txtInfo(Index).Left, txtInfo(Index).Top)
            Set rsTmp = zlDatabase.ShowSelect(Me, StrSQL, 0, "����ԭ��", , , , , , True, vPoint.X, vPoint.Y, txtInfo(Index).Height, blnCancel)
            If Not rsTmp Is Nothing Then
                txtInfo(Index).Text = rsTmp!����
                txtInfo(Index).SetFocus
            End If
        Case txt��֢�໤��
            StrSQL = " Select Distinct A.ID,A.����,A.����" & _
                    " From ���ű� A,��������˵�� B" & _
                    " Where B.����ID=A.ID And B.��������='ICU'" & _
                    " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                    " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                    " Order by A.����"
            vPoint = GetCoordPos(txtInfo(Index).Container.hwnd, txtInfo(Index).Left, txtInfo(Index).Top)
            Set rsTmp = zlDatabase.ShowSelect(Me, StrSQL, 0, "��֢�໤��", _
                False, "", "", False, False, True, vPoint.X, vPoint.Y, txtInfo(Index).Height, blnCancel)
            
            If rsTmp Is Nothing Then
                If Not blnCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                    MsgBox "û������ICU��֢�໤�ҡ�", vbInformation, Me.Caption
                End If
            Else
                txtInfo(Index).Text = rsTmp!���� & ""
            End If
        Case txtҽѧ��ʾ
            'ѡ��ҽѧ��ʾ
            On Error GoTo errH
            vPoint = GetCoordPos(txtInfo(Index).Container.hwnd, txtInfo(Index).Left, txtInfo(Index).Top)
            StrSQL = "Select Rownum ID,����,����,���� From ҽѧ��ʾ Order by ����"
            Set rsTmp = zlDatabase.ShowSQLMultiSelect(Me, StrSQL, 0, "", True, "", "", True, True, True, vPoint.X, vPoint.Y, txtInfo(Index).Height, blnCancel, True, True)

            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "û������""ҽѧ��ʾ""���ݣ����ȵ��ֵ�����������á�", vbInformation, gstrSysName
                End If
                txtInfo(Index).SetFocus
            Else
                While Not rsTmp.EOF
                    strResult = strResult & "," & rsTmp!����
                    rsTmp.MoveNext
                Wend
                txtInfo(Index).Text = Mid(strResult, 2)
                txtInfo(Index).SetFocus
                Call zlCommFun.PressKey(vbKeyTab)
            End If
    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdOK_Click()
    Dim blnDiagnose As Boolean
    
    If Not CheckPageData(blnDiagnose, False) Then Exit Sub
    If mblnDiagnose And Not blnDiagnose Then
        If MsgBox("Ҫ��������Ϣ��û�����룬Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
    End If
    
    If Not SavePageData(False) Then Exit Sub
    
    mblnDiagnose = blnDiagnose
    mblnOk = True
    Unload Me
End Sub

Private Sub cmdPathLoad_Click()
'���ܣ��Զ���ȡ·����Ϣ
    Dim StrSQL As String, rsTmp As Recordset
    
    
    StrSQL = "Select Decode(c.����, 2, c.����, '') As ����,b.״̬" & vbNewLine & _
            "From ����·������ A, �����ٴ�·�� B, ���쳣��ԭ�� C" & vbNewLine & _
            "Where a.·����¼id(+) = b.Id And b.��ǰ���� = a.����(+) And Nvl(b.��ǰ�׶�id, b.ǰһ�׶�id) = a.�׶�id(+) And b.״̬ <> 0 And a.����ԭ�� = c.����(+) And b.����id = [1] And b.��ҳid = [2]"

    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    If rsTmp.RecordCount > 0 Then
        chkInfo(chk����·��).Value = 1
        If Val(rsTmp!״̬ & "") = 3 Then
            chkInfo(chk���·��).Value = 0
            txtInfo(txt�˳�ԭ��).Text = rsTmp!���� & ""
        ElseIf Val(rsTmp!״̬ & "") = 2 Then
            chkInfo(chk���·��).Value = 1
        End If
    Else
        chkInfo(chk����·��).Value = 0
    End If
    '��ȡ�������
    StrSQL = "Select Count(1) Over(Partition By b.����id, b.��ҳid) As ������, c.���� As ����ԭ��" & vbNewLine & _
            "From ����·������ A, �����ٴ�·�� B, ���쳣��ԭ�� C" & vbNewLine & _
            "Where a.·����¼id = b.Id And c.����(+) = a.����ԭ�� And a.������� = -1 And b.����id = [1] And b.��ҳid = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    If rsTmp.RecordCount > 0 Then
        chkInfo(chk����).Value = 1
        If Val(rsTmp!������ & "") = 1 Then
            txtInfo(txt����ԭ��).Text = rsTmp!����ԭ�� & ""
        End If
    Else
        chkInfo(chk����).Value = 0
    End If
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub cmdPrint_Click()
    Dim blnDiagnose As Boolean
    
    If Not SavePageDataUnit(blnDiagnose, False) Then Exit Sub
    
    mblnDiagnose = blnDiagnose
    
    Call PrintInMedRec(2, mlng����ID, mlng��ҳID, mobjReport, mlng����ID, Me)
End Sub

Private Sub cmdPriview_Click()
    Dim blnDiagnose As Boolean
    
    If Not SavePageDataUnit(blnDiagnose, False) Then Exit Sub
    
    mblnDiagnose = blnDiagnose
    
    Call PrintInMedRec(1, mlng����ID, mlng��ҳID, mobjReport, mlng����ID, Me)
End Sub

Private Sub cmdPrintdown_Click()
    PopupMenu menuPrint, , cmdPrint.Left, cmdPrint.Top + cmdPrint.Height
End Sub

Private Sub cmdPriviewDown_Click()
    PopupMenu menuPriview, , cmdPriview.Left, cmdPriview.Top + cmdPriview.Height
End Sub

Private Sub menuPage_Click(Index As Integer)
    Dim blnDiagnose As Boolean
    
    If Not SavePageDataUnit(blnDiagnose, False) Then Exit Sub
    
    mblnDiagnose = blnDiagnose
    
    Call PrintInMedRec(1, mlng����ID, mlng��ҳID, mobjReport, mlng����ID, Me, Index)
End Sub

Private Sub menuPagePrint_Click(Index As Integer)
    Dim blnDiagnose As Boolean
    
    If Not SavePageDataUnit(blnDiagnose, False) Then Exit Sub
    
    mblnDiagnose = blnDiagnose
    
    Call PrintInMedRec(2, mlng����ID, mlng��ҳID, mobjReport, mlng����ID, Me, Index)
End Sub

Private Sub cmdSign_Click(Index As Integer)
'���ܣ�ǩ��
    Dim StrSQL As String
    Dim rsTmp As Recordset, i As Long
    Dim bln���� As Boolean    '�Ƿ���д��������¼
    
    '�ж��Ƿ���������ǩ��
    If gintCA > 0 And Mid(gstrESign, 2, 1) = "1" Then
        If mobjESign Is Nothing Then
            On Error Resume Next
            Set mobjESign = CreateObject("zl9ESign.clsESign")
            Err.Clear: On Error GoTo 0
            If Not mobjESign Is Nothing Then
                Call mobjESign.Initialize(gcnOracle, glngSys)
            End If
        End If
        If mobjESign Is Nothing Then
                MsgBox "����ǩ������δ����ȷ��װ��ǩ���������ܼ�����", vbInformation, gstrSysName
            Exit Sub
        Else
            If Not mobjESign.CheckCertificate(gstrDBUser) Then Exit Sub
        End If
    End If
    
    '��Ҫȷ������ǩ���������
    If Index = cmdסԺҽʦ Or Index = cmd����ҽʦ Or Index = cmd����ҽʦ Then
        If cboinfo(cbo������).Text = "" Then
            Call ShowMessage(cboinfo(cbo������), "û��ȷ�������Ρ�")
            Exit Sub
        End If
    End If
    If Index = cmdסԺҽʦ Or Index = cmd����ҽʦ Then
        If cboinfo(cbo����ҽʦ).Text = "" Then
            Call ShowMessage(cboinfo(cbo����ҽʦ), "û��ȷ������ҽʦ��")
            Exit Sub
        End If
    End If
    If Index = cmdסԺҽʦ Then
        If cboinfo(cbo����ҽʦ).Text = "" Then
            Call ShowMessage(cboinfo(cbo����ҽʦ), "û��ȷ������ҽʦ��")
            Exit Sub
        End If
    End If
    
    'ǩ��ǰ�Զ�����
    If mblnChange Then
        If Not CheckPageData(False, True) Then Exit Sub
        If Not SavePageData(True) Then Exit Sub
    End If
    
    On Error GoTo errH
    
    '�����������¼������ʾ�Ƿ����
    bln���� = False
    For i = 1 To vsOPS.Rows - 1
        If Trim(vsOPS.TextMatrix(i, col��������)) <> "" Then
            bln���� = True
        End If
    Next
    
    StrSQL = "Select Count(1) As ���� From ����ҽ����¼ Where ����ID=[1] And ��ҳID=[2] And ҽ��״̬=8 And �������='F'"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    If Val(rsTmp!���� & "") > 0 And Not bln���� Then
        vsOPS.Row = 1: vsOPS.Col = col��������
        If ShowMessage(vsOPS, "�ò��˴�������ҽ��������ҳ��û�����������¼���Ƿ������", True) = vbNo Then Exit Sub
    End If
    
    If Index = cmd������ Then
        StrSQL = "Zl_������ҳ�ӱ�_��ҳ����(" & mlng����ID & "," & mlng��ҳID & ",'������ǩ��','" & UserInfo.���� & "')"
    ElseIf Index = cmd����ҽʦ Then
        StrSQL = "Zl_������ҳ�ӱ�_��ҳ����(" & mlng����ID & "," & mlng��ҳID & ",'����ҽʦǩ��','" & UserInfo.���� & "')"
    ElseIf Index = cmd����ҽʦ Then
        StrSQL = "Zl_������ҳ�ӱ�_��ҳ����(" & mlng����ID & "," & mlng��ҳID & ",'����ҽʦǩ��','" & UserInfo.���� & "')"
    ElseIf Index = cmdסԺҽʦ Then
        StrSQL = "Zl_������ҳ�ӱ�_��ҳ����(" & mlng����ID & "," & mlng��ҳID & ",'סԺҽʦǩ��','" & UserInfo.���� & "')"
    End If
    Call zlDatabase.ExecuteProcedure(StrSQL, Me.Caption)
    
    mblnReadOnly = SetSignature()
    Call SetFaceEditable(mblnReadOnly)
    If cmdOK.Enabled Then cmdOK.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmdUnSign_Click(Index As Integer)
'���ܣ�ȡ��ǩ��
    Dim StrSQL As String
    
    If gintCA > 0 And Mid(gstrESign, 2, 1) = "1" Then
        If mobjESign Is Nothing Then
            On Error Resume Next
            Set mobjESign = CreateObject("zl9ESign.clsESign")
            Err.Clear: On Error GoTo 0
            If Not mobjESign Is Nothing Then
                Call mobjESign.Initialize(gcnOracle, glngSys)
            End If
        End If
        If mobjESign Is Nothing Then
                MsgBox "����ǩ������δ����ȷ��װ��ǩ���������ܼ�����", vbInformation, gstrSysName
            Exit Sub
        Else
            If Not mobjESign.CheckCertificate(gstrDBUser) Then Exit Sub
        End If
    End If
    
    '������鲡���Ƿ��Ŀ����ҳ��������״̬
    If Not CheckMecRed(mlng����ID, mlng��ҳID, Me.Caption, "ȡ��ǩ��") Then Exit Sub
        
    On Error GoTo errH
    
    If Index = cmd������ Then
        StrSQL = "Zl_������ҳ�ӱ�_��ҳ����(" & mlng����ID & "," & mlng��ҳID & ",'������ǩ��',Null)"
    ElseIf Index = cmd����ҽʦ Then
        StrSQL = "Zl_������ҳ�ӱ�_��ҳ����(" & mlng����ID & "," & mlng��ҳID & ",'����ҽʦǩ��',Null)"
    ElseIf Index = cmd����ҽʦ Then
        StrSQL = "Zl_������ҳ�ӱ�_��ҳ����(" & mlng����ID & "," & mlng��ҳID & ",'����ҽʦǩ��',Null)"
    ElseIf Index = cmdסԺҽʦ Then
        StrSQL = "Zl_������ҳ�ӱ�_��ҳ����(" & mlng����ID & "," & mlng��ҳID & ",'סԺҽʦǩ��',Null)"
    End If
    Call zlDatabase.ExecuteProcedure(StrSQL, Me.Caption)
    
    mblnReadOnly = SetSignature()
    Call SetFaceEditable(mblnReadOnly)
    If cmdOK.Enabled Then cmdOK.SetFocus
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub dpkInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub dtpInfo_DateClick(ByVal DateClicked As Date)
    Dim strDate As String
    
    If mlngDateIndex = txtȷ������ Then
        If IsDate(txtInfo(txtȷ������).Text) Then
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(txtInfo(txtȷ������).Text, "yyyy-MM-dd HH:mm"), 12, 5)
        Else
            strDate = Format(DateClicked, "yyyy-MM-dd") & " " & Mid(Format(zlDatabase.Currentdate, "yyyy-MM-dd HH:mm"), 12, 5)
        End If
        If Not CheckDateRange(strDate, True) Then
            MsgBox "�������ʱ������ڲ��˵�סԺ�ڼ䡣", vbInformation, Me.Caption
            Exit Sub
        End If
    ElseIf mlngDateIndex = txt�ʿ����� Then
        strDate = Format(DateClicked, "yyyy-MM-dd")
    End If
    txtInfo(mlngDateIndex).Text = strDate
    dtpInfo.Visible = False
    txtInfo(mlngDateIndex).SetFocus
    
    If Visible Then mblnChange = True
End Sub


Private Sub dtpInfo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call dtpInfo_DateClick(dtpInfo.Value)
    End If
End Sub

Private Sub dtpInfo_Validate(Cancel As Boolean)
    dtpInfo.Visible = False
End Sub

Private Sub Form_Activate()
    mblnIsFirst = True
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If dtpInfo.Visible Then
            dtpInfo.Visible = False
        Else
            Call cmdCancel_Click
        End If
    ElseIf KeyCode = vbKeyF1 Then
        '###
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Load()
    Dim i As Integer
    Dim lngW As Long, lngH As Long
    Dim ctlTmp As Control
    Dim StrSQL As String
    Dim rsTmp As Recordset
    
    Me.Opened = True
    mblnIsFirst = False
    mstrPathDiag = ""
    '���Ի����ñ������ǰ�Ŀ�͸߿���̫С
    lngW = Me.Width
    lngH = Me.Height
    Call RestoreWinState(Me, App.ProductName)
    If lngW <> Me.Width Then Me.Width = lngW
    If lngH <> Me.Height Then Me.Height = lngH
    
    On Error Resume Next
    If Val(zlDatabase.GetPara("��ҽ�������", glngSys, pסԺҽ��վ, 0, Array(optInput(0), optInput(1)), InStr(mstrPrivs, "��������") > 0)) = 0 Then
        optInput(0).Value = True
    Else
        optInput(1).Value = True
    End If
    If Val(zlDatabase.GetPara("��ҽ�������", glngSys, pסԺҽ��վ, 0, Array(optInput(2), optInput(3)), InStr(mstrPrivs, "��������") > 0)) = 0 Then
        optInput(2).Value = True
    Else
        optInput(3).Value = True
    End If
    mstr����������� = zlDatabase.GetPara("�����������", glngSys, pסԺҽ��վ, 0, Array(optInput(4), optInput(5), chkInfo(chk��������¼��)), InStr(mstrPrivs, "��������") > 0)
    If Mid(mstr�����������, 1, 1) = "0" Then
        optInput(4).Value = True
    Else
        optInput(5).Value = True
    End If
    chkInfo(chk��������¼��).Value = Val(Mid(mstr�����������, 2, 1))
    
    mlng�����ж� = Val(zlDatabase.GetPara("�����ж����", glngSys, pסԺҽ��վ, 2) & "")
    mlng������� = Val(zlDatabase.GetPara("������ϼ��", glngSys, pסԺҽ��վ, 2) & "")
    mlng���� = Val(zlDatabase.GetPara("������", glngSys, pסԺҽ��վ, 1) & "")
    If InStr(mstrPrivs, "�޸�ҽ�Ƹ��ʽ") > 0 Then
        cboinfo(cbo���ʽ).Enabled = True
    Else
        cboinfo(cbo���ʽ).Enabled = False
    End If
    
    mbln���ýṹ����ַ = Val(zlDatabase.GetPara("���˵�ַ�ṹ��¼��", glngSys, pסԺҽ��վ, 0)) <> 0
    mbln��ʹ����ҽ��Ŀ = Val(zlDatabase.GetPara("��ҽ���Ҳ�ʹ����ҽ������ҳ��Ŀ", glngSys, pסԺҽ��վ, 0)) <> 0
    mblnҽ����ʿ������ҳ = Val(zlDatabase.GetPara("ҽ���ͻ�ʿ�ֱ���д������ҳ", glngSys, pסԺҽ��վ, 0)) = 1
    
    If mbln���ýṹ����ַ Then
        txtInfo(txt�����ص�).Visible = False
        txtInfo(txt����).Visible = False
        txtInfo(txt���ڵ�ַ).Visible = False
        txtInfo(txt��ͥ��ַ).Visible = False
        cmdInfo(txt�����ص�).Visible = False
        cmdInfo(txt����).Visible = False
        cmdInfo(txt���ڵ�ַ).Visible = False
        cmdInfo(txt��ͥ��ַ).Visible = False
    Else
        PatiAddress������.Visible = False
        PatiAddress���ڵ�ַ.Visible = False
        PatiAddress����.Visible = False
        PatiAddress��סַ.Visible = False
    End If
    

    Call optInput_Click(0)
    On Error GoTo 0
    
    '���������Դ
    If gint�����Դ > 1 Then
        optInput(0).Enabled = False
        optInput(1).Enabled = False
        optInput(2).Enabled = False
        optInput(3).Enabled = False
        If gint�����Դ = 2 Then
            optInput(0).Value = True
            optInput(2).Value = True
        ElseIf gint�����Դ = 3 Then
            optInput(1).Value = True
            optInput(3).Value = True
        End If
    End If
    
    mblnOk = False
    mblnChange = False
    mstrLike = IIf(Val(zlDatabase.GetPara("����ƥ��")) = 0, "%", "")
    mint���� = Val(zlDatabase.GetPara("���뷽ʽ")) '����ƥ�䷽ʽ��0-ƴ��,1-���
    
    '��Ƭ����
    If Not Have��������(mlng����ID, "����") Then
        vsOPS.ColHidden(col������ʿ) = True
    End If
    mbln��ҽ = Have��������(mlng����ID, "��ҽ��")
    If Not mbln��ҽ Then
        sstInfo.TabVisible(TAB_��ҽ���) = False
    End If
    For i = 0 To sstInfo.Tabs - 1
        fraInfo(i).BackColor = Me.BackColor
    Next
    If mbln��ʿվ Then
        sstInfo.Tab = TAB_����
    Else
        sstInfo.Tab = TAB_������Ϣ
    End If
    
    '���ƻ���
    mbln�������� = CheckShare(300) '����ϵͳ
    If Not mbln�������� Then
        sstInfo.TabVisible(TAB_�����뻯��) = False
        lblInfo(107).Visible = False
    End If
    StrSQL = "select ��Ϣֵ from ������ҳ�ӱ� where ����id=0 and ��ҳid=0"
    On Error Resume Next
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption)
    mlngSize = rsTmp.Fields("��Ϣֵ").DefinedSize
    
    
    vs����.Tag = "δ�޸�"
    vs����.Tag = "δ�޸�"
    vsfMain.Tag = "δ�޸�"
    
    '��ʼ������
    If Not InitPageData Then Unload Me: Exit Sub
    '��ȡ��ҳ����
    If Not LoadPageData Then Unload Me: Exit Sub
    
    Call SetEditableFrom��Ժ���
    Call Set��ԭѧ
    
    'ȱʡ��λ�����ҳ
    If mblnDiagnose Then
        If mbln��ҽ Then
            sstInfo.Tab = TAB_��ҽ���
        Else
            sstInfo.Tab = TAB_��ҽ���
        End If
    End If

    '����ǩ�����������ֻ�����
    If mblnReadOnly Then
        Call SetSignature
        Call SetFaceEditable(True)
        'ǩ�������ݵ�������
        cboinfo(cbo������).Locked = True: cboinfo(cbo����ҽʦ).Locked = True
        cboinfo(cbo����ҽʦ).Locked = True: cboinfo(cboסԺҽʦ).Locked = True
        cboinfo(cbo������).BackColor = vbButtonFace: cboinfo(cbo����ҽʦ).BackColor = vbButtonFace
        cboinfo(cbo����ҽʦ).BackColor = vbButtonFace: cboinfo(cboסԺҽʦ).BackColor = vbButtonFace
        For i = 0 To cmdSign.UBound
            cmdSign(i).Visible = False: cmdUnSign(i).Visible = False
        Next
    Else
        'ҽ��վ���ж�ǩ������
        If Not mbln��ʿվ Then
            mblnReadOnly = SetSignature
        End If
        Call SetFaceEditable(mblnReadOnly)
    End If
        
    'û�������г�������ʱ����һ������,ֻ�����ѳ�Ժʱ����������
    If txtInfo(txt����).Text = "" And IsDate(txt��������.Text) _
        And Not (mblnReadOnly Or IsDate(txtInfo(txt��Ժʱ��).Text)) Then 'ֻ�����ѳ�Ժʱ���Զ���������
        txt��������.Tag = "": Call txt��������_Validate(False)
    End If
End Sub

Private Sub PatiAddress������_Validate(Cancel As Boolean)
    If PatiAddress������.Tag <> PatiAddress������.Value Then mblnChange = True
End Sub

Private Sub PatiAddress���ڵ�ַ_Validate(Cancel As Boolean)
    If PatiAddress���ڵ�ַ.Tag <> PatiAddress���ڵ�ַ.Value Then mblnChange = True
End Sub

Private Sub PatiAddress����_Validate(Cancel As Boolean)
    If PatiAddress����.Tag <> PatiAddress����.Value Then mblnChange = True
End Sub

Private Sub PatiAddress��סַ_Validate(Cancel As Boolean)
    If PatiAddress��סַ.Tag <> PatiAddress��סַ.Value Then mblnChange = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If mblnChange Then
        If MsgBox("����˳����ղ����޸ĵ����ݽ����ᱻ���档ȷʵҪ�˳���", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
            Cancel = True: Exit Sub
        End If
    End If
    
    If Not mobjESign Is Nothing Then Set mobjESign = Nothing
    Set mcol��ԱSQL = Nothing
    
    Call zlDatabase.SetPara("��ҽ�������", IIf(optInput(0).Value, 0, 1), glngSys, pסԺҽ��վ, InStr(mstrPrivs, "��������") > 0)
    Call zlDatabase.SetPara("��ҽ�������", IIf(optInput(2).Value, 0, 1), glngSys, pסԺҽ��վ, InStr(mstrPrivs, "��������") > 0)
    Call zlDatabase.SetPara("�����������", IIf(optInput(4).Value, "0", "1") & IIf(chkInfo(chk��������¼��).Value = 1, "1", "0"), glngSys, pסԺҽ��վ, InStr(mstrPrivs, "��������") > 0)
    Call SaveWinState(Me, App.ProductName)
        
    Me.Opened = False
    RaiseEvent Closed(Not mblnOk, mstr����ID, mstr���ID)
End Sub

Private Sub fra��ҳ_DragDrop(Index As Integer, Source As Control, X As Single, Y As Single)

End Sub

Private Sub lstAdvEvent_GotFocus()
    lstAdvEvent.ListIndex = 0
End Sub

Private Sub lstAdvEvent_ItemCheck(Item As Integer)
    If lstAdvEvent.List(Item) = "ѹ��" Then
        cboinfo(cboѹ�������ڼ�).Enabled = lstAdvEvent.Selected(Item)
        cboinfo(cboѹ�������ڼ�).TabStop = cboinfo(cboѹ�������ڼ�).Enabled
        cboinfo(cboѹ������).Enabled = lstAdvEvent.Selected(Item)
        cboinfo(cboѹ������).TabStop = cboinfo(cboѹ������).Enabled
        If cboinfo(cboѹ�������ڼ�).Enabled Then
            cboinfo(cboѹ�������ڼ�).BackColor = vbWindowBackground
            cboinfo(cboѹ������).BackColor = vbWindowBackground
        Else
            cboinfo(cboѹ�������ڼ�).BackColor = vbButtonFace
            cboinfo(cboѹ������).BackColor = vbButtonFace
        End If
    ElseIf lstAdvEvent.List(Item) = "ҽԺ�ڵ���/׹��" Then
        cboinfo(cbo������׹���˺�).Enabled = lstAdvEvent.Selected(Item)
        cboinfo(cbo������׹���˺�).TabStop = cboinfo(cbo������׹���˺�).Enabled
        cboinfo(cbo������׹��ԭ��).Enabled = lstAdvEvent.Selected(Item)
        cboinfo(cbo������׹��ԭ��).TabStop = cboinfo(cbo������׹��ԭ��).Enabled
        If cboinfo(cbo������׹���˺�).Enabled Then
            cboinfo(cbo������׹��ԭ��).BackColor = vbWindowBackground
            cboinfo(cbo������׹���˺�).BackColor = vbWindowBackground
        Else
            cboinfo(cbo������׹��ԭ��).BackColor = vbButtonFace
            cboinfo(cbo������׹���˺�).BackColor = vbButtonFace
        End If
    End If
    If mblnIsFirst Then mblnChange = True
End Sub

Private Sub lstAdvEvent_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If lstAdvEvent.ListIndex = lstAdvEvent.ListCount - 1 Then
            If cboinfo(cbo������׹���˺�).Enabled Or cboinfo(cboѹ�������ڼ�).Enabled Then
                If cboinfo(cboѹ�������ڼ�).Enabled Then
                    cboinfo(cboѹ�������ڼ�).SetFocus
                Else
                    cboinfo(cbo������׹���˺�).SetFocus
                End If
            Else
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Else
            lstAdvEvent.ListIndex = lstAdvEvent.ListIndex + 1
        End If
    End If
End Sub

Private Sub lstInfection_GotFocus()
    lstInfection.ListIndex = 0
End Sub

Private Sub lstInfection_ItemCheck(Item As Integer)
    If mblnIsFirst Then mblnChange = True
End Sub

Private Sub lstInfection_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        If lstInfection.ListIndex = lstInfection.ListCount - 1 Then zlCommFun.PressKey vbKeyTab: Exit Sub
        lstInfection.ListIndex = lstInfection.ListIndex + 1
    End If
End Sub

Private Sub optInput_Click(Index As Integer)
    Dim i As Integer

    If Index = opt31���� Then
        txtInfo(txt31��Ŀ��).Enabled = False
        txtInfo(txt31��Ŀ��).BackColor = &H8000000F
    ElseIf Index = opt31���� Then
        txtInfo(txt31��Ŀ��).Enabled = True
        txtInfo(txt31��Ŀ��).BackColor = &H80000005
    End If
End Sub

Private Sub optInput_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub sstInfo_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txtInfo_Change(Index As Integer)
    If Index = txt���� Then
        '��������Ŵ���׼���䵥λ
        If IsNumeric(txtInfo(Index).Text) Or txtInfo(Index).Text = "" Then
            cboinfo(cbo���䵥λ).Visible = True
            If cboinfo(cbo���䵥λ).ListIndex = -1 Then cboinfo(cbo���䵥λ).ListIndex = 0
        Else
            cboinfo(cbo���䵥λ).Visible = False
            cboinfo(cbo���䵥λ).ListIndex = -1
        End If
    ElseIf Index = txtӤ������ Then
        '��������Ŵ���׼���䵥λ
        If IsNumeric(txtInfo(Index).Text) Or txtInfo(Index).Text = "" Then
            cboinfo(cboӤ�����䵥λ).Visible = True
            If cboinfo(cboӤ�����䵥λ).ListIndex = -1 Then cboinfo(cboӤ�����䵥λ).ListIndex = 0
        Else
            cboinfo(cboӤ�����䵥λ).Visible = False
            cboinfo(cboӤ�����䵥λ).ListIndex = -1
        End If
    ElseIf Index = txt���ȴ��� Then
        If Val(txtInfo(Index).Text) > 0 Then
            txtInfo(txt�ɹ�����).Locked = False
            txtInfo(txt�ɹ�����).TabStop = True
            txtInfo(txt�ɹ�����).BackColor = vbWindowBackground
            
            '��Ҫ��ϵĳ�Ժ�����Ϊ����ʱ,ȱʡ���ɹ�����=���ȴ���
            If Visible Then
                If vsDiagXY.TextMatrix(GetRow(3), col��Ժ���) <> "����" Then
                    txtInfo(txt�ɹ�����).Text = txtInfo(txt���ȴ���).Text
                ElseIf Val(txtInfo(txt���ȴ���).Text) > 1 Then
                    txtInfo(txt�ɹ�����).Text = Val(txtInfo(txt���ȴ���).Text) - 1
                End If
            End If
        Else
            txtInfo(txt�ɹ�����).Text = ""
            txtInfo(txt�ɹ�����).Locked = True
            txtInfo(txt�ɹ�����).TabStop = False
            txtInfo(txt�ɹ�����).BackColor = vbButtonFace
        End If
    ElseIf Index = txtת��1 Then
        If txtInfo(Index).Text = "" Then
            txtInfo(txtת��2).Text = ""
            txtInfo(txtת��3).Text = ""
        End If
    ElseIf Index = txtת��2 Then
        If txtInfo(Index).Text = "" Then
            txtInfo(txtת��3).Text = ""
        End If
    End If
    If Visible Then mblnChange = True
End Sub

Private Sub txtInfo_GotFocus(Index As Integer)
    Call zlControl.TxtSelAll(txtInfo(Index))
End Sub

Private Sub txtInfo_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF4 Or (KeyCode = vbKeyDown And Shift = vbAltMask) Then
        If Index = txtȷ������ Then
            Call cmdInfo_Click(Index)
        End If
    ElseIf KeyCode = vbKeyDelete Then
        If Index = txtҽѧ��ʾ Then
            txtInfo(txtҽѧ��ʾ) = ""
        End If
    End If
End Sub

Private Sub txtInfo_KeyPress(Index As Integer, KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String, blnCancel As Boolean
    Dim vPoint As POINTAPI, strMask As String
    
    If KeyAscii = 13 Then
        KeyAscii = 0
        If (Index = txt�����ص� Or Index = txt��ͥ��ַ Or Index = txt��ϵ�˵�ַ Or Index = txt���ڵ�ַ) And txtInfo(Index).Text <> "" Then
            '�����������
            StrSQL = "Select Rownum as ID,����,����,���� From ���� " & _
                " Where (���� Like [1] Or ���� Like [2] Or ���� Like [2])" & _
                " Order by ����"
            vPoint = GetCoordPos(txtInfo(Index).Container.hwnd, txtInfo(Index).Left, txtInfo(Index).Top)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 0, "����", False, "", "", False, _
                False, True, vPoint.X, vPoint.Y, txtInfo(Index).Height, blnCancel, False, False, _
                UCase(txtInfo(Index).Text) & "%", mstrLike & UCase(txtInfo(Index).Text) & "%")
            '������������,��һ��Ҫƥ��
            If Not rsTmp Is Nothing Then
                txtInfo(Index).Text = rsTmp!����
            End If
            txtInfo(Index).SetFocus
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf (Index = txt���� Or Index = txt����) And txtInfo(Index).Text <> "" Then
            '������������
            StrSQL = "Select Rownum as ID,����,����,���� From ���� " & _
                " Where (���� Like [1] Or ���� Like [2] Or ���� Like [2])" & _
                " Order by ����"
            vPoint = GetCoordPos(txtInfo(Index).Container.hwnd, txtInfo(Index).Left, txtInfo(Index).Top)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 0, IIf(Index = txt����, "����", "����"), False, "", "", False, _
                False, True, vPoint.X, vPoint.Y, txtInfo(Index).Height, blnCancel, False, False, _
                UCase(txtInfo(Index).Text) & "%", mstrLike & UCase(txtInfo(Index).Text) & "%")
            '������������,��һ��Ҫƥ��
            If Not rsTmp Is Nothing Then
                txtInfo(Index).Text = rsTmp!����
            End If
            txtInfo(Index).SetFocus
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf Index = txt��λ���� And txtInfo(Index).Text <> "" Then
            '���빤����λ
            StrSQL = "Select ID,����,����,����,��ַ,�绰,��������,�ʺ�,��ϵ�� From ��Լ��λ" & _
                " Where (����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or ����ʱ�� is NULL)" & _
                " And (���� Like [1] Or ���� Like [2] Or ���� Like [2])" & _
                " Order by ����"
            vPoint = GetCoordPos(txtInfo(Index).Container.hwnd, txtInfo(Index).Left, txtInfo(Index).Top)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 0, "������λ", False, "", "", False, _
                False, True, vPoint.X, vPoint.Y, txtInfo(Index).Height, blnCancel, False, False, _
                UCase(txtInfo(Index).Text) & "%", mstrLike & UCase(txtInfo(Index).Text) & "%")
            '������������,��һ��Ҫƥ��
            If Not rsTmp Is Nothing Then
                txtInfo(Index).Text = rsTmp!���� & IIf(Not IsNull(rsTmp!��ַ), "(" & rsTmp!��ַ & ")", "")
                txtInfo(Index).Tag = Val(rsTmp!ID)
                If txtInfo(txt��λ�绰).Text = "" Then
                    txtInfo(txt��λ�绰).Text = Nvl(rsTmp!�绰)
                End If
            Else
                txtInfo(Index).Tag = ""
            End If
            txtInfo(Index).SetFocus
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf (Index = txtת��1 Or Index = txtת��2 Or Index = txtת��3) And txtInfo(Index).Text <> "" Then
            '����ת�ƿ���
            StrSQL = "Select Distinct A.ID,A.����,A.����,A.����,A.λ��" & _
                " From ���ű� A,��������˵�� B" & _
                " Where A.ID=B.����ID And B.������� IN(2,3) And B.�������� IN('�ٴ�','����')" & _
                " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                " And (���� Like [1] Or ���� Like [2] Or ���� Like [2])" & _
                " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                " Order by A.����"
            vPoint = GetCoordPos(txtInfo(Index).Container.hwnd, txtInfo(Index).Left, txtInfo(Index).Top)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 0, "ת�ƿ���", False, "", "", False, _
                False, True, vPoint.X, vPoint.Y, txtInfo(Index).Height, blnCancel, False, False, _
                UCase(txtInfo(Index).Text) & "%", mstrLike & UCase(txtInfo(Index).Text) & "%")
            '������������,��һ��Ҫƥ��
            If Not rsTmp Is Nothing Then
                txtInfo(Index).Text = rsTmp!����
            End If
            txtInfo(Index).SetFocus
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf Index = txtסԺ�� Then
            If txtInfo(Index).Text = "" Then
                txtInfo(Index).Text = zlDatabase.GetNextNo(2)
            End If
            Call zlCommFun.PressKey(vbKeyTab)
        ElseIf Index = txt��Ժ���� Or Index = txt����ԭ�� Then
            If Index = txt����ԭ�� Then
                 'ѡ��λ��Ϣ
                If txtInfo(Index).Text <> "" Then
                    StrSQL = "Select ���� ID,����,���� From ���Ȳ������ where ���� like [1] or ���� like [2] or to_number(����)=[3]"
                       
                    vPoint = GetCoordPos(txtInfo(Index).Container.hwnd, txtInfo(Index).Left, txtInfo(Index).Top)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 0, "����ԭ��", True, 1, "���Ȳ���", False, False, True, vPoint.X, vPoint.Y, txtInfo(Index).Height, blnCancel, False, False, gstrLike & txtInfo(Index).Text & "%", gstrLike & txtInfo(Index).Text & "%", Val(txtInfo(Index).Text))
                    If Not rsTmp Is Nothing Then
                        txtInfo(Index).Text = rsTmp!����
                        txtInfo(Index).SetFocus
                    Else
                        Exit Sub
                    End If
                End If
            End If
            '������һ����Ƭ
            If sstInfo.TabVisible(sstInfo.Tab + IIf(sstInfo.TabVisible(sstInfo.Tab + 1), 1, 2)) Then sstInfo.Tab = sstInfo.Tab + IIf(sstInfo.TabVisible(sstInfo.Tab + 1), 1, 2)
            If Index = txt����ԭ�� And Not mbln��ҽ Then
                vsAller.SetFocus
            Else
                Call sstInfo_KeyPress(13)
            End If
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    ElseIf KeyAscii = vbKeyBack Then
        If Index = txtҽѧ��ʾ Then
            txtInfo(txtҽѧ��ʾ).Text = ""
        End If
    ElseIf Not (KeyAscii >= 0 And KeyAscii < 32) Then
        '�ǿ��ư���
        If Index = txtҽѧ��ʾ Then
            KeyAscii = 0
        End If
        'ѡ���ݼ�
        If KeyAscii = Asc("*") Then
            'ע�������Ҫ��CMD�Ͷ�ӦTXT��Index��ͬ
            On Error Resume Next
            StrSQL = ""
            StrSQL = cmdInfo(Index).Name
            Err.Clear: On Error GoTo 0
            If StrSQL <> "" Then
                KeyAscii = 0
                Call cmdInfo_Click(Index)
                Exit Sub
            End If
        End If
        
        '�������볤��
        If txtInfo(Index).MaxLength <> 0 Then
            If zlCommFun.ActualLen(txtInfo(Index).Text) > txtInfo(Index).MaxLength Then
                KeyAscii = 0: Exit Sub
            End If
        End If
        
        '������������
        Select Case Index
'            Case txt���� '��������¼����
'                strMask = "1234567890"
            'Case txt�������� 'MaskEdit������
                'strMask = "1234567890-"
            Case txtȷ������, txt�ʿ�����
                strMask = "1234567890-: "
            Case txt��ͥ�绰, txt��λ�绰, txt��ϵ�˵绰
                strMask = "1234567890-()"
            Case txtסԺ��, txt�����ʱ�, txt��ͥ�ʱ�, txt��λ�ʱ�, txt���ȴ���, txt�ɹ�����, txt��������, txt��ԺǰСʱ, txt��Ժǰ����, txt��Ժ��Сʱ, txt��Ժ�����, txt������Сʱ
                strMask = "1234567890"
            Case txt���ϸ��, txt��ѪС��, txt��Ѫ��, txt��ȫѪ, txt�������, txtԼ����ʱ��
                strMask = "1234567890."
            Case txt���, txt����, txt����������, txt��������Ժ����
                strMask = "1234567890."
        End Select
        If strMask <> "" Then
            If InStr(strMask, Chr(KeyAscii)) = 0 Then
                KeyAscii = 0: Exit Sub
            End If
        End If
    Else
        If Index = txtҽѧ��ʾ Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub SetCboFromList(ByVal arrList As Variant, ByVal arrCboIdx As Variant, Optional ByVal intDefault As Integer = -1)
'���ܣ���ָ������װ��ָ��ComboBox
'������arrList=List String����
'      arrCboIdx=ComboBox��������,���ComboBoxʱ,װ��������ͬ
'      intDefaut=ȱʡ����
    Dim i As Long, j As Long
    
    For i = 0 To UBound(arrCboIdx)
        cboinfo(arrCboIdx(i)).Clear
        For j = 0 To UBound(arrList)
            cboinfo(arrCboIdx(i)).AddItem arrList(j)
        Next
        cboinfo(arrCboIdx(i)).ListIndex = intDefault 'ȱʡΪδѡ��
    Next
End Sub

Private Sub SetCboFromSQL(ByVal StrSQL As String, ByVal arrCboIdx As Variant, Optional ByVal strSQLExt As String, Optional colsql As Collection)
'���ܣ���ָ������Դ�е�����װ��ָ��������һ������ComboBox
'������strSQL=����"ID,����,����/����,ȱʡ��־/ȱʡ"�ֶΣ�����Order by���������ΪA
'      strSQLExt=���ӵ�SQL����
'      colSQL=Ҫ����SQL�ļ���
    Dim rsTmp As New ADODB.Recordset
    Dim str���� As String, strȱʡ As String
    Dim i As Long, j As Long
    
    For i = 0 To UBound(arrCboIdx)
        '���ԭ������
        cboinfo(arrCboIdx(i)).Clear

        '��¼ԭʼSQL
        If Not colsql Is Nothing Then
            colsql.Add StrSQL, "_" & arrCboIdx(i)
        End If
    Next
    
    If strSQLExt <> "" Then
        StrSQL = Replace(UCase(StrSQL), UCase("Order by"), strSQLExt & " Order by")
    End If
    On Error GoTo errH
    Call zlDatabase.OpenRecordset(rsTmp, StrSQL, Me.Caption)
    
    'װ������
    If Not rsTmp.EOF Then
        For i = 0 To rsTmp.Fields.Count - 1
            If rsTmp.Fields(i).Name = "����" Or rsTmp.Fields(i).Name = "����" Then
                str���� = rsTmp.Fields(i).Name
            ElseIf rsTmp.Fields(i).Name = "ȱʡ��־" Or rsTmp.Fields(i).Name = "ȱʡ" Then
                strȱʡ = rsTmp.Fields(i).Name
            End If
        Next
        For i = 1 To rsTmp.RecordCount
            For j = 0 To UBound(arrCboIdx)
                If IsNull(rsTmp!����) Then
                    cboinfo(arrCboIdx(j)).AddItem rsTmp.Fields(str����).Value
                Else
                    cboinfo(arrCboIdx(j)).AddItem rsTmp!���� & "-" & Chr(13) & rsTmp.Fields(str����).Value
                End If
                cboinfo(arrCboIdx(j)).ItemData(cboinfo(arrCboIdx(j)).NewIndex) = Nvl(rsTmp!ID, 0)
                If strȱʡ <> "" Then
                    If Nvl(rsTmp.Fields(strȱʡ).Value, 0) = 1 Then
                        Call zlControl.CboSetIndex(cboinfo(arrCboIdx(j)).hwnd, cboinfo(arrCboIdx(j)).NewIndex)
                    End If
                End If
            Next
            rsTmp.MoveNext
        Next
    End If
    
    '��ȱʡʱ,Ϊδѡ��
    For i = 0 To UBound(arrCboIdx)
        If cboinfo(arrCboIdx(i)).Style = 0 Then
            cboinfo(arrCboIdx(i)).AddItem "[����...]"
            cboinfo(arrCboIdx(i)).ItemData(cboinfo(arrCboIdx(i)).NewIndex) = -1
        End If
    Next
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function SetCboFromName(ByVal strName As String, objCbo As Object) As Boolean
'���ܣ���ָ����������Ա���뵽��������
    Static rsTmp As ADODB.Recordset
    Dim StrSQL As String, intIdx As Integer
    
    On Error GoTo errH
    
    If rsTmp Is Nothing Then
        StrSQL = "Select A.ID,A.���,A.����,Null As ����" & _
            " From ��Ա�� A,��Ա����˵�� B" & _
            " Where A.ID=B.��ԱID And B.��Ա���� IN('ҽ��','��ʿ')" & _
            " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " Order by A.����"
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, StrSQL, "SetCboFromName")
    End If
    
    rsTmp.Filter = "����='" & strName & "'"
    If Not rsTmp.EOF Then
        intIdx = objCbo.ListCount
        If objCbo.ListCount > 0 Then
            If objCbo.ItemData(objCbo.ListCount - 1) = -1 Then
                intIdx = objCbo.ListCount - 1
            End If
        End If
        
        If IsNull(rsTmp!����) Then
            objCbo.AddItem rsTmp!����, intIdx
        Else
            objCbo.AddItem rsTmp!���� & "-" & Chr(13) & rsTmp!����, intIdx
        End If
        objCbo.ItemData(objCbo.NewIndex) = Val(rsTmp!ID)
        
        objCbo.ListIndex = objCbo.NewIndex
    End If
    
    SetCboFromName = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function InitPageData() As Boolean
'���ܣ���ʼ����ҳ�༭ʱ����Ҫ��һЩ����
    Dim StrSQL As String, strSQLExt As String
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    
    '���ò���������ĸ߶ȵĿ��
    Call zlControl.CboSetWidth(cboinfo(cboְҵ).hwnd, cboinfo(cboְҵ).Width + 500)
    Call zlControl.CboSetWidth(cboinfo(cbo����).hwnd, cboinfo(cbo����).Width * 2)
    Call zlControl.CboSetHeight(cboinfo(cbo����), cboinfo(cbo����).Height * 16)
    Call zlControl.CboSetHeight(cboinfo(cbo����), cboinfo(cbo����).Height * 16)
    Call zlControl.CboSetHeight(cboinfo(cboְҵ), cboinfo(cboְҵ).Height * 16)
    Call zlControl.CboSetHeight(cboinfo(cbo����ҽʦ), cboinfo(cbo����ҽʦ).Height * 16)
    Call zlControl.CboSetHeight(cboinfo(cbo������), cboinfo(cbo������).Height * 16)
    Call zlControl.CboSetHeight(cboinfo(cbo����ҽʦ), cboinfo(cbo����ҽʦ).Height * 16)
    Call zlControl.CboSetHeight(cboinfo(cbo����ҽʦ), cboinfo(cbo����ҽʦ).Height * 16)
    Call zlControl.CboSetHeight(cboinfo(cboסԺҽʦ), cboinfo(cboסԺҽʦ).Height * 16)
    Call zlControl.CboSetHeight(cboinfo(cbo����ҽʦ), cboinfo(cbo����ҽʦ).Height * 16)
    Call zlControl.CboSetHeight(cboinfo(cbo�о���ҽʦ), cboinfo(cbo�о���ҽʦ).Height * 16)
    Call zlControl.CboSetHeight(cboinfo(cboʵϰҽʦ), cboinfo(cboʵϰҽʦ).Height * 16)
    Call zlControl.CboSetHeight(cboinfo(cbo�ʿ�ҽʦ), cboinfo(cbo�ʿ�ҽʦ).Height * 16)
    Call zlControl.CboSetHeight(cboinfo(cbo�ʿػ�ʿ), cboinfo(cbo�ʿػ�ʿ).Height * 16)
    
    '���̶ֹ����ݵ�������
    Call SetCboFromList(Array("δ��", "��ʧ����", "δ��"), Array(cbo���֤��), 0)
    Call SetCboFromList(Array("��", "��", "��", "Сʱ", "����"), Array(cbo���䵥λ), 0) '�����Ŀʱ��ע��cboInfo(cbo���䵥λ).listIndex<3���ж�
    Call SetCboFromList(Array("��", "��", "��", "��", "����"), Array(cbo����Ex), 0)
    Call SetCboFromList(Array("0-δ��", "1-��", "2-��", "3-����"), Array(cboRh))
    Call SetCboFromList(Array("1.1-��", "1.2-����", "2-����", "3-��"), Array(cbo�������, cbo���ȷ���))
    Call SetCboFromList(Array("0-δ֪", "1-��", "2-��"), Array(cbo������ҩ))
    Call SetCboFromList(Array(" ", "1-��", "2-��"), Array(cboʹ����ҽ�����豸))
    Call SetCboFromList(Array(" ", "1-��", "2-��"), Array(cboʹ����ҽ���Ƽ���))
    Call SetCboFromList(Array(" ", "1-��", "2-��"), Array(cbo��֤ʩ��))
    Call SetCboFromList(Array("0-δ��", "1-׼ȷ", "2-����׼ȷ", "3-�ش�ȱ��", "4-����"), Array(cbo��֤, cbo�η�, cbo��ҩ))
    Call SetCboFromList(Array("0-δ��", "1-����", "2-����", "3-������"), Array(cboHBsAg))
    Call SetCboFromList(Array("0-δ��", "1-����", "2-����"), Array(cboHCVAb, cboHIVAb))
    Call SetCboFromList(Array("1-��", "2-��", "3-δ��"), Array(cbo��Һ��Ӧ))
    Call SetCboFromList(Array("0-��", "1-��", "2-δ��"), Array(cbo��Ѫ��Ӧ))
    Call SetCboFromList(Array("1-��", "2-��", "3-����"), Array(cbo��Ѫ���))
    Call SetCboFromList(Array("0-δ��", "1-����", "2-������", "3-���϶�"), Array(cbo�������Ժ, cbo��������Ժ, cbo��Ժ���Ժ, cbo�����벡��, cbo�ٴ��벡��, cbo�ٴ���ʬ��, cbo��ǰ������, cbo��ҽ�������Ժ, cbo��ҽ��Ժ���Ժ))
    Call SetCboFromList(Array(" ", "0-��Ժǰ", "1-סԺ�ڼ�"), Array(cboѹ�������ڼ�))
    Call SetCboFromList(Array(" ", "1��", "2��", "3��", "4��"), Array(cboѹ������))
    Call SetCboFromList(Array(" ", "һ��", "����", "����", "δ����˺�"), Array(cbo������׹���˺�))
    Call SetCboFromList(Array(" ", "����ԭ��", "���ơ�ҩ�����ԭ��", "��������", "����ԭ��"), Array(cbo������׹��ԭ��))
    Call SetCboFromList(Array("��", "��", "Сʱ", "����"), Array(cboӤ�����䵥λ))
    Call SetCboFromList(Array("31������סԺ�ƻ�", "7������סԺ�ƻ�"), Array(cbo31���7������Ժ))
    Call SetCboFromList(Array("", "1-��", "2-��", "3-��"), Array(cbo��������))
    Call SetCboFromList(Array("", "һ��", "����", "����", "����"), Array(cboԼ����ʽ))
    Call SetCboFromList(Array("", "��ʽ��", "Ӳʽ��", "����", "������", "Լ����", "����"), Array(cboԼ������))
    Call SetCboFromList(Array("", "��֪�ϰ�", "���ܵ���", "��Ϊ����", "������Ҫ", "�궯", "ҽ������", "����"), Array(cboԼ��ԭ��))
    Call SetCboFromList(Array("", "ҽ����Ժ", "ת����", "תԺ", "��ҽ����Ժ", "����"), Array(cbo��������Ժ��ʽ))
    Call SetCboFromList(Array("���ط�", "24h��", "24-48h", "��48h"), Array(cbo�ط����ʱ��))
    Call SetCboFromList(Array("", "0-δ����", "1-����1̥", "2-����2̥������", "4-����"), Array(cbo����״��), 0)
    cboinfo(cbo31���7������Ժ).ListIndex = 0
    cboinfo(cboӤ�����䵥λ).ListIndex = 0
    cboinfo(cboӤ�����䵥λ).ListIndex = 0
    
    '����һЩ�ֵ���������������
    Call SetCboFromSQL("Select 0 as ID,���� as ����,����,ȱʡ��־ From ҽ�Ƹ��ʽ Order by ����", Array(cbo���ʽ))
    Call SetCboFromSQL("Select 0 as ID,���� as ����,����,ȱʡ��־ From �Ա� Order by ����", Array(cbo�Ա�))
    Call SetCboFromSQL("Select 0 as ID,���� as ����,����,ȱʡ��־ From ����״�� Order by ����", Array(cbo����))
    Call SetCboFromSQL("Select 0 as ID,���� as ����,����,ȱʡ��־ From ְҵ Order by ����", Array(cboְҵ))
    Call SetCboFromSQL("Select 0 as ID,���� as ����,����,ȱʡ��־ From ���� Order by ����", Array(cbo����))
    Call SetCboFromSQL("Select 0 as ID,���� as ����,����,ȱʡ��־ From ���� Order by ����", Array(cbo����))
    Call SetCboFromSQL("Select 0 as ID,���� as ����,����,ȱʡ��־ From Ѫ�� Order by ����", Array(cboѪ��))
    Call SetCboFromSQL("Select 0 as ID,���� as ����,����,ȱʡ��־ From ����ϵ Order by ����", Array(cbo��ϵ�˹�ϵ))
    Call SetCboFromSQL("Select 0 as ID,���� as ����,����,0 as ȱʡ��־ From ���� Order by ����", Array(cbo��Ժ����))
    Call SetCboFromSQL("Select 0 as ID,���� as ����,����,0 as ȱʡ��־ From �ٴ��������� Order by ����", Array(cbo��������))
    Call SetCboFromSQL("Select 0 as ID,���� as ����,����,ȱʡ��־ From ��Ժ��ʽ Order by ����", Array(cbo��Ժ��ʽ))
    Call SetCboFromSQL("Select 0 as ID,���� as ����,����,ȱʡ��־ From �ֻ��̶� Order by ����", Array(cbo�ֻ��̶�))
    Call SetCboFromSQL("Select 0 as ID,���� as ����,����,ȱʡ��־ From ���������� Order by ����", Array(cbo����������))
    cboinfo(cbo��������).AddItem " "
    
    'ҽ������ʿ����----------------------------------------------------------------
    Set mcol��ԱSQL = New Collection
    strSQLExt = " And Exists(Select 1 From ������Ա Where ��ԱID=A.ID And ����ID IN(Select B.����ID From �ϻ���Ա�� A,������Ա B Where A.�û���=User And A.��ԱID=B.��ԱID))"
    
    '����ҽ��
    StrSQL = "Select Distinct A.ID,A.���,A.����,Null as ����" & _
        " From ��Ա�� A,��Ա����˵�� B,������Ա C,��������˵�� D" & _
        " Where A.ID=B.��ԱID And B.��Ա����='ҽ��' And A.ID=C.��ԱID And C.����ID=D.����ID And D.������� IN(1,2,3)" & _
        " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        " Order by A.����"
    Call SetCboFromSQL(StrSQL, Array(cbo����ҽʦ), , mcol��ԱSQL)
    
    'ҽ��
    StrSQL = "Select A.ID,A.���,A.����,Null as ����" & _
        " From ��Ա�� A,��Ա����˵�� B" & _
        " Where A.ID=B.��ԱID And B.��Ա����='ҽ��'" & _
        " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        " Order by A.����"
    Call SetCboFromSQL(StrSQL, Array(cbo����ҽʦ, cbo�о���ҽʦ, cboʵϰҽʦ, cbo�ʿ�ҽʦ), strSQLExt, mcol��ԱSQL)
    
    'סԺҽʦ
    StrSQL = "Select A.ID,A.���,A.����,Null as ����" & _
        " From ��Ա�� A,��Ա����˵�� B" & _
        " Where A.ID=B.��ԱID And B.��Ա����='ҽ��' And A.רҵ����ְ�� IN('����ҽʦ','������ҽʦ','����ҽʦ','ҽʦ','ҽʿ')" & _
        " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        " Order by A.����"
    Call SetCboFromSQL(StrSQL, Array(cboסԺҽʦ), strSQLExt, mcol��ԱSQL)
    
    '����ҽʦ
    StrSQL = "Select A.ID,A.���,A.����,Null as ����" & _
        " From ��Ա�� A,��Ա����˵�� B" & _
        " Where A.ID=B.��ԱID And B.��Ա����='ҽ��' And A.רҵ����ְ�� IN('����ҽʦ','������ҽʦ','����ҽʦ')" & _
        " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        " Order by A.����"
    Call SetCboFromSQL(StrSQL, Array(cbo����ҽʦ), strSQLExt, mcol��ԱSQL)
    
    '����ҽʦ
    StrSQL = "Select A.ID,A.���,A.����,Null as ����" & _
        " From ��Ա�� A,��Ա����˵�� B" & _
        " Where A.ID=B.��ԱID And B.��Ա����='ҽ��' And A.רҵ����ְ�� IN('����ҽʦ','������ҽʦ')" & _
        " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        " Order by A.����"
    Call SetCboFromSQL(StrSQL, Array(cbo����ҽʦ), strSQLExt, mcol��ԱSQL)
    
    '������
    StrSQL = "Select A.ID,A.���,A.����,Null as ����" & _
        " From ��Ա�� A,��Ա����˵�� B" & _
        " Where A.ID=B.��ԱID And B.��Ա����='ҽ��' And A.����ְ�� IN('��������','���Ҹ�����')" & _
        " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        " Order by A.����"
    Call SetCboFromSQL(StrSQL, Array(cbo������), strSQLExt, mcol��ԱSQL)
    
    '�ʿػ�ʿ
    StrSQL = "Select A.ID,A.���,A.����,Null as ����" & _
        " From ��Ա�� A,��Ա����˵�� B" & _
        " Where A.ID=B.��ԱID And B.��Ա����='��ʿ'" & _
        " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        " Order by A.����"
    Call SetCboFromSQL(StrSQL, Array(cbo�ʿػ�ʿ), strSQLExt, mcol��ԱSQL)
    
    '���λ�ʿ
    StrSQL = "Select A.ID,A.���,A.����,Null as ����" & _
        " From ��Ա�� A,��Ա����˵�� B" & _
        " Where A.ID=B.��ԱID And B.��Ա����='��ʿ'" & _
        " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        " Order by A.����"
    Call SetCboFromSQL(StrSQL, Array(cbo���λ�ʿ), strSQLExt, mcol��ԱSQL)
    
    '��Ժ��ʽ
    cboinfo(cbo��Ժ��ʽ).Clear
    StrSQL = "select ���� AS ID,����,'' ����,ȱʡ��־ from ��Ժ��ʽ order by ����"
    Call SetCboFromSQL(StrSQL, Array(cbo��Ժ��ʽ))
    
    '-------------------
    Call SetKSSSerial
    Call vsKSS_AfterRowColChange(-1, -1, vsKSS.Row, vsKSS.Col)
    Call vsTSJC_AfterRowColChange(-1, -1, vsTSJC.Row, vsTSJC.Col)
    
    If mbln�������� Then Call Init���������Grid
    Call FillVsf

    Screen.MousePointer = 0
    InitPageData = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function Load��ҳ����(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '-------------------------------------------------------------------------------------------------------------------------
    '����:���ظ�ҳ����
    '����:lng����id-����id
    '     lng��ҳid -��ҳid
    '����:���سɹ�,����true,���򷵻�False
    '-------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    Dim StrSQL As String
    
    Err = 0: On Error GoTo Errhand:
     
    
    '����֢���
    StrSQL = "" & _
        " Select �໤������,�˹������ѳ�,�ط���֢ҽѧ��," & _
        "      �ط����ʱ�� " & _
        " From ������֢�໤��� " & _
        " where ����id=[1] and ��ҳid=[2] " & _
        " order by ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, lng����ID, lng��ҳID)
    If rsTemp.RecordCount > 0 Then
        txtInfo(txt��֢�໤��).Text = rsTemp!�໤������ & ""
        chkInfo(chk�˹������ѳ�).Value = Val(rsTemp!�˹������ѳ� & "")
        chkInfo(chk�ط���֢ҽѧ��).Value = Val(rsTemp!�ط���֢ҽѧ�� & "")
        Call GetCboIndex(cboinfo(cbo�ط����ʱ��), Nvl(rsTemp!�ط����ʱ��))
    End If
    
    Load��ҳ���� = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Function Get���ƽ��() As String
    Dim rsTmp As New ADODB.Recordset
    Dim StrSQL As String
        
    On Error GoTo errH
    StrSQL = "Select ����,����,���� From ���ƽ�� Order by ����"
    Call zlDatabase.OpenRecordset(rsTmp, StrSQL, Me.Caption)
    
    StrSQL = ""
    Do While Not rsTmp.EOF
        StrSQL = StrSQL & "|" & rsTmp!���� & "-" & rsTmp!����
        rsTmp.MoveNext
    Loop
    If StrSQL = "" Then
        Get���ƽ�� = "1-����|2-��ת|3-δ��|4-����|5-����"
    Else
        Get���ƽ�� = Mid(StrSQL, 2)
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Init���������Grid()
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʼ�������뻯������ؼ���Ĭ������
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2008-10-21 15:16:08
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Set rsTemp = Get���������(True)
        
    With vs����
        .Rows = 2
        .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
        .Clear 1
        .Editable = flexEDKbdMouse

        .ColComboList(.ColIndex("��ѧ���Ʊ���")) = .BuildComboList(rsTemp, "������Ϣ", "ID")
        If rsTemp.RecordCount = 1 Then
            .ColData(.ColIndex("��ѧ���Ʊ���")) = Nvl(rsTemp!ID) & ";" & Nvl(rsTemp!������Ϣ)
        Else
            rsTemp.Filter = "ȱʡ��־=1"
            If rsTemp.EOF = False Then
                .ColData(.ColIndex("��ѧ���Ʊ���")) = Nvl(rsTemp!ID) & ";" & Nvl(rsTemp!������Ϣ)
            Else
                .ColData(.ColIndex("��ѧ���Ʊ���")) = ";"
                lblEdit(2).Caption = "û�п��õĻ������Ʊ��룬�뵽����ϵͳ�����á�"
            End If
        End If
        Call vs����_LostFocus
        zl_vsGrid_Para_Restore glngModul, vs����, Me.Caption, "����"
    End With
    Set rsTemp = Get���������(False)
    With vs����
        .Rows = 2
        .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
        .Clear 1
        .Editable = flexEDKbdMouse
        .ColComboList(.ColIndex("�������Ʊ���")) = .BuildComboList(rsTemp, "������Ϣ", "ID")
        If rsTemp.RecordCount = 1 Then
            .ColData(.ColIndex("�������Ʊ���")) = Nvl(rsTemp!ID) & ";" & Nvl(rsTemp!������Ϣ)
        Else
            rsTemp.Filter = "ȱʡ��־=1"
            If rsTemp.EOF = False Then
                .ColData(.ColIndex("�������Ʊ���")) = Nvl(rsTemp!ID) & ";" & Nvl(rsTemp!������Ϣ)
            Else
                .ColData(.ColIndex("�������Ʊ���")) = ";"
                lblEdit(2).Caption = "û�п��õķ������Ʊ��룬�뵽����ϵͳ�����á�"
            End If
        End If
        Call vs����_LostFocus
        zl_vsGrid_Para_Restore glngModul, vs����, Me.Caption, "����"
    End With
End Sub

Private Function Get���������(ByVal bln���� As Boolean, Optional ByVal arrControl As Variant, Optional blnSetup As Boolean = False) As ADODB.Recordset
    '-----------------------------------------------------------------------------------------------------------
    '����:��ʼ����������ƵĲ���
    '���:
    '����:
    '����:
    '����:���˺�
    '����:2008-10-21 10:37:13
    '-----------------------------------------------------------------------------------------------------------
    Dim strTemp As String, StrSQL As String
    Dim arrData As Variant, strDefaultCode As String, strCodeIN As String
    Dim rsTemp As New ADODB.Recordset, i As Long
    
    
    '��ȡ���ƺͻ���
    '   zlDatabase.SetPara IIf(bln����, "������Ŀ", "������Ŀ"), strSaveData, glngSys, mlngModule, False
    strTemp = zlDatabase.GetPara(IIf(Not bln����, "������Ŀ", "������Ŀ"), glngSys * 3, 200, , arrControl, blnSetup)
    If strTemp <> "" Then
        arrData = Split(strTemp, ";")
        For i = 0 To UBound(arrData)
            If InStr(1, arrData(i), ",") > 0 Then
                If Val(Split(arrData(i), ",")(1)) = 1 Then
                    strDefaultCode = Split(arrData(i), ",")(0)
                End If
                strCodeIN = strCodeIN & "," & Split(arrData(i), ",")(0)
            Else
                strCodeIN = strCodeIN & "," & arrData(i)
            End If
        Next
    End If
    If strCodeIN <> "" Then
        strCodeIN = Mid(strCodeIN, 2)
    Else
        strCodeIN = ";-"
    End If
    StrSQL = "" & _
    "   Select /*+ Rule*/ A.id,A.����,A.����||'-'||A.���� as ������Ϣ,decode(A.����,[2],1,0) as ȱʡ��־ " & _
    "   From ��������Ŀ¼ A, " & _
    "       Table(Cast(f_Str2list([1]) As zlTools.t_Strlist)) B " & _
    "   Where A.���� = B.Column_Value"
    On Error GoTo errH
    Set Get��������� = zlDatabase.OpenSQLRecord(StrSQL, "��ȡ�����������Ϣ", strCodeIN, strDefaultCode)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function zl_vsGrid_Para_Restore(ByVal lngModule As Long, ByVal vsGrid As VSFlexGrid, ByVal strCaption, ByVal strKEY As String, _
    Optional blnSaveToDataBase As Boolean = False, Optional blnǿ�ƻָ����� As Boolean = False) As Boolean
    '------------------------------------------------------------------------------
    '����:�����ݿ��лָ�����Ŀ�ȵ���Ϣ
    '����:vsGrid-��Ӧ������ؼ�
    '     strCaption-������
    '     strKey-����
    '     blnSaveToDataBase-�Ƿ��������ݿ��б������(����������ݿ��б���,��ǿ�Ʊ���Ϊtrue,��������Ƿ�ʹ�ø��Ի������ȷ��)
    '     blnǿ�ƻָ�����-�����Ƿ񽫱���ע���Ĳ���ֵ,����ǿ�ƻָ�
    '����:�ָ��ɹ�,����True,���򷵻�False
    '����:���˺�
    '����:2008/03/03
    '------------------------------------------------------------------------------
    Dim strParaValue As String, intCols As Integer, arrReg As Variant, ArrTemp As Variant, intCol As Integer, intRow As Integer
    Dim intTemp As Integer, strColName As String
    
    If blnSaveToDataBase = False Then
        'ֻ���ڱ���ע����вŻᴦ����Ի�����
        zl_vsGrid_Para_Restore = True
        If blnǿ�ƻָ����� = False Then
            If Val(zlDatabase.GetPara("ʹ�ø��Ի����")) = 0 Then Exit Function
        End If
        Call GetRegInFor(g˽��ģ��, strCaption, strKEY, strParaValue)
    Else
        strParaValue = zlDatabase.GetPara(strKEY, glngSys, lngModule)
    End If
    
    zl_vsGrid_Para_Restore = False
    If strParaValue = "" Then Exit Function
    'strParaValue:�����ʽ:������,�п�,������|������,�п�,������|...
    Err = 0: On Error GoTo Errhand:
    arrReg = Split(strParaValue, "|")
    If vsGrid.Cols <> UBound(arrReg) + 1 Then Exit Function
    intCols = UBound(arrReg) + 1
    With vsGrid
        For intCol = 0 To intCols - 1
            ArrTemp = Split(arrReg(intCol) & ",,", ",")
            strColName = ArrTemp(0)
            intTemp = .ColIndex(strColName)
            If intTemp <> -1 Then
                .ColWidth(intTemp) = Val(ArrTemp(1))
                If Val(ArrTemp(2)) = 1 Then
                    .ColHidden(intTemp) = True
                Else
                    .ColHidden(intTemp) = False
                End If
                If .ColWidth(intTemp) = 0 Then .ColHidden(intTemp) = True
                .ColPosition(.ColIndex(strColName)) = intCol
            End If
        Next
    End With
    zl_vsGrid_Para_Restore = True
    Exit Function
Errhand:
End Function

Private Sub GetRegInFor(ByVal RegType As gRegType, ByVal strSection As String, _
                ByVal strKEY As String, ByRef strKeyValue As String)
    '--------------------------------------------------------------------------------------------------------------
    '����:  ��ָ����ע����Ϣ��ȡ����
    '�����:  RegType-ע������
    '       strSection-ע���Ŀ¼
    '       StrKey-����
    '������:
    '       strKeyValue-���صļ�ֵ
    '����:
    '--------------------------------------------------------------------------------------------------------------
    Dim strValue As String
    Err = 0
    On Error GoTo Errhand:
    Select Case RegType
        Case gע����Ϣ
            SaveSetting "ZLSOFT", "ע����Ϣ\" & strSection, strKEY, strKeyValue
            strKeyValue = GetSetting("ZLSOFT", "ע����Ϣ\" & strSection, strKEY, "")
        Case g����ȫ��
            strKeyValue = GetSetting("ZLSOFT", "����ȫ��\" & strSection, strKEY, "")
        Case g����ģ��
            strKeyValue = GetSetting("ZLSOFT", "����ģ��" & "\" & App.ProductName & "\" & strSection, strKEY, "")
        Case g˽��ȫ��
            strKeyValue = GetSetting("ZLSOFT", "˽��ȫ��\" & gstrDBUser & "\" & strSection, strKEY, "")
        Case g˽��ģ��
            strKeyValue = GetSetting("ZLSOFT", "˽��ģ��\" & gstrDBUser & "\" & App.ProductName & "\" & strSection, strKEY, "")
    End Select
Errhand:
End Sub

Private Function GetIDTmp(ByVal strName As String) As Long
'���ܣ��������ڽ�������ҳ�ӱ�Ŀ����� �Ƶ����±� ���˿����ؼ�¼�У���ǰû�м�¼ҩƷid�����ڸ������ƽ�id�����
    Dim rsTmp As Recordset, StrSQL As String
    
    On Error GoTo errH
    StrSQL = "Select Distinct a.Id" & vbNewLine & _
                "From ������ĿĿ¼ A, ������Ŀ���� B, ҩƷ���� C" & vbNewLine & _
                "Where a.Id = b.������Ŀid And a.Id = c.ҩ��id And Nvl(c.������, 0) <> 0 And A.����=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, strName)
    If rsTmp.RecordCount > 0 Then
        GetIDTmp = Val(rsTmp!ID)
    Else
        GetIDTmp = 0
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function LoadPageData() As Boolean
'���ܣ���ȡ���˵���ҳ��Ϣ
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String, i As Long, j As Long
    Dim lngRow As Long, varTmp As Variant
    Dim str���ƽ�� As String, blnDo As Boolean
    Dim lngCol As Long
    Dim bln�ֻ��̶� As Boolean
    Dim strTmp As String
    
    On Error GoTo errH
    
    Screen.MousePointer = 11
    
    If mlngPathState <> -1 Then
        'ֻ������ҳ���������ϣ���ǰû��ģ�ȱʡ���������ڡ���ҽ��Ժ��ϡ�
        StrSQL = "Select Nvl(�������,2) as �������,NVL(����ID,0) As ����ID,NVL(���ID,0) as ���ID,״̬ From �����ٴ�·�� Where ����ID=[1] And ��ҳID=[2] And (�����Դ = 3 or �����Դ is null) Order By ����ʱ��"
        Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
        If rsTmp.RecordCount > 0 Then
            mlngDiagnosisType = rsTmp!�������
            '����ж���·������ȡ��һ����״̬
            If rsTmp.RecordCount >= 2 Then mlngPathState = Val(rsTmp!״̬ & "")
            rsTmp.MoveNext
            Do While Not rsTmp.EOF
                mstrPathDiag = mstrPathDiag & "," & rsTmp!������� & "|" & rsTmp!����id & "|" & rsTmp!���id
                rsTmp.MoveNext
            Loop
            mstrPathDiag = Mid(mstrPathDiag, 2)
        Else
            mlngDiagnosisType = 0
        End If
        '���·����ʱ���Ƿ�ȳ�Ժ��ϼ�¼ʱ���()ȡ��һ��·��
        If mlngPathState = 2 Then
            StrSQL = "Select Sign(Nvl(a.����ʱ��, Null)-Nvl(b.��¼����, Sysdate)) As �ж�" & vbNewLine & _
                    "From �����ٴ�·�� A, (Select ����id, ��ҳid, ��¼���� From ������ϼ�¼ Where ��¼��Դ = 3 And ��ϴ��� = 1 And ������� = [3]) B" & vbNewLine & _
                    " Where a.����id = b.����id(+) And a.��ҳid = b.��ҳid(+) And a.����ID=[1] And A.��ҳID=[2]" & _
                    " and a.����ʱ��=(Select Min(����ʱ��) From �����ٴ�·�� Where ����ID=[1] and ��ҳID=[2])"
            Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID, IIf(mlngDiagnosisType > 10, 13, 3))
            If rsTmp.RecordCount > 0 Then
                mblnIsPathOutTime = Val(rsTmp!�ж� & "") = 1
            Else
                mblnIsPathOutTime = False
            End If
        End If
    End If
    
    '������Ϣ����
    '---------------------------------------------------------------
    StrSQL = "Select ����,�Ա�,��������,�����ص�,���֤��,����֤��,����,����,סԺ��,���� From ������Ϣ Where ����ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID)
        
    txtInfo(txtסԺ����).Text = mlng��ҳID
    txtInfo(txt����).Text = Nvl(rsTmp!����)
    Call GetCboIndex(cboinfo(cbo�Ա�), Nvl(rsTmp!�Ա�))
    
    If Not IsNull(rsTmp!��������) Then
        txt��������.Text = Format(rsTmp!��������, "yyyy-MM-dd")
        If Format(rsTmp!��������, "HH:mm") <> "00:00" Then
            txt����ʱ��.Text = Format(rsTmp!��������, "HH:mm")
        End If
    End If
    txt��������.Tag = txt��������.Text '���ڼ�¼����仯
    If mbln���ýṹ����ַ Then
        '������
        Call SetStrucAddress(PatiAddress������, GetStrucAddress(mlng����ID, mlng��ҳID, 1), Nvl(rsTmp!�����ص�))
        PatiAddress������.Tag = PatiAddress������.Value
        '����
        Call SetStrucAddress(PatiAddress����, GetStrucAddress(mlng����ID, mlng��ҳID, 2), Nvl(rsTmp!����))
        PatiAddress����.Tag = PatiAddress����.Value
    Else
        txtInfo(txt�����ص�).Text = Nvl(rsTmp!�����ص�)
        txtInfo(txt����).Text = Nvl(rsTmp!����)
    End If
    cboinfo(cbo���֤��).Text = Nvl(rsTmp!���֤��)
    txtInfo(txt����֤��).Text = Nvl(rsTmp!����֤��)
    Call GetCboIndex(cboinfo(cbo����), Nvl(rsTmp!����))
    
    txtInfo(txt����).Text = Nvl(rsTmp!����)
    
    '������ҳ����
    '---------------------------------------------------------------
    StrSQL = "Select A.*,B.���� as ��Ժ����,C.���� as ��Ժ����" & _
        " From ������ҳ A,���ű� B,���ű� C" & _
        " Where A.��Ժ����ID=B.ID And A.��Ժ����ID=C.ID" & _
        " And A.����ID=[1] And A.��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    
    mint���� = Nvl(rsTmp!����, 0)
    
    '���۲�����סԺ��
    If Nvl(rsTmp!��������, 0) <> 0 Then
        lblInfo(0).Visible = False
        txtInfo(txtסԺ��).Visible = False
        txtInfo(txtסԺ��).Enabled = False '��־Ϊ�����
    End If
    
    Call GetCboIndex(cboinfo(cbo���ʽ), Nvl(rsTmp!ҽ�Ƹ��ʽ))
    
    Call LoadOldData("" & rsTmp!����)
    
    Call GetCboIndex(cboinfo(cbo����), Nvl(rsTmp!����״��))
    Call GetCboIndex(cboinfo(cboְҵ), Nvl(rsTmp!ְҵ))
    
    Call GetCboIndex(cboinfo(cbo����), Nvl(rsTmp!����))
    If Not IsNull(rsTmp!����) Then
        txtInfo(txt����).Text = Nvl(rsTmp!����)
    End If
    txtInfo(txtסԺ��).Text = Nvl(rsTmp!סԺ��)
    If mbln���ýṹ����ַ Then
        '��סַ
        Call SetStrucAddress(PatiAddress��סַ, GetStrucAddress(mlng����ID, mlng��ҳID, 3), Nvl(rsTmp!��ͥ��ַ))
        PatiAddress��סַ.Tag = PatiAddress��סַ.Value
        '���ڵ�ַ
        Call SetStrucAddress(PatiAddress���ڵ�ַ, GetStrucAddress(mlng����ID, mlng��ҳID, 4), Nvl(rsTmp!���ڵ�ַ))
        PatiAddress���ڵ�ַ.Tag = PatiAddress���ڵ�ַ.Value
    Else
        txtInfo(txt��ͥ��ַ).Text = Nvl(rsTmp!��ͥ��ַ)
        txtInfo(txt���ڵ�ַ).Text = Nvl(rsTmp!���ڵ�ַ)
    End If
    txtInfo(txt��ͥ�绰).Text = Nvl(rsTmp!��ͥ�绰)
    txtInfo(txt��ͥ�ʱ�).Text = Nvl(rsTmp!��ͥ��ַ�ʱ�)
    txtInfo(txt��λ����).Text = Nvl(rsTmp!��λ��ַ)
    txtInfo(txt��λ�绰).Text = Nvl(rsTmp!��λ�绰)
    txtInfo(txt��λ�ʱ�).Text = Nvl(rsTmp!��λ�ʱ�)
    txtInfo(txt�����ʱ�).Text = Nvl(rsTmp!���ڵ�ַ�ʱ�)
    txtInfo(txt��ϵ������).Text = Nvl(rsTmp!��ϵ������)
    Call GetCboIndex(cboinfo(cbo��ϵ�˹�ϵ), Nvl(rsTmp!��ϵ�˹�ϵ))
    txtInfo(txt��ϵ�˵绰).Text = Nvl(rsTmp!��ϵ�˵绰)
    txtInfo(txt��ϵ�˵�ַ).Text = Nvl(rsTmp!��ϵ�˵�ַ)
    
    chkInfo(chk����Ժ).Value = Nvl(rsTmp!����Ժ, 0)
    txtInfo(txt��Ժʱ��).Text = Format(rsTmp!��Ժ����, "yyyy-MM-dd HH:mm")
    
    txtInfo(txt��Ժ����).Text = rsTmp!��Ժ����
    Call GetCboIndex(cboinfo(cbo��Ժ����), Nvl(rsTmp!��Ժ����))
    
    txtInfo(txt��Ժʱ��).Text = Format(Nvl(rsTmp!��Ժ����), "yyyy-MM-dd HH:mm")
    
    txtInfo(txt��Ժ����).Text = rsTmp!��Ժ����
    
    Call GetCboIndex(cboinfo(cbo��Ժ��ʽ), Nvl(rsTmp!��Ժ��ʽ))
    
    If Not IsNull(rsTmp!��Ժ����) Then
        txtInfo(txtסԺ����).Text = DateDiff("d", rsTmp!��Ժ����, rsTmp!��Ժ����)
    Else
        txtInfo(txtסԺ����).Text = DateDiff("d", rsTmp!��Ժ����, zlDatabase.Currentdate)
    End If
    If Val(txtInfo(txtסԺ����).Text) = 0 Then txtInfo(txtסԺ����).Text = "1"
    
    chkInfo(chk�Ƿ�ȷ��).Value = Nvl(rsTmp!�Ƿ�ȷ��, 0)
    If chkInfo(chk�Ƿ�ȷ��).Value = 1 Then
        txtInfo(txtȷ������).Text = Format(Nvl(rsTmp!ȷ������), "yyyy-MM-dd HH:mm")
    End If
    txtInfo(txt���ȴ���).Text = Nvl(rsTmp!���ȴ���)
    If Val(txtInfo(txt���ȴ���).Text) <> 0 Then
        txtInfo(txt�ɹ�����).Text = Nvl(rsTmp!�ɹ�����)
    End If
    chkInfo(chk�·�����).Value = Nvl(rsTmp!�·�����, 0)
    Call GetCboIndex(cboinfo(cbo�������), Nvl(rsTmp!��ҽ�������))
    chkInfo(chkʬ��).Value = Nvl(rsTmp!ʬ���־, 0)
    
    chkInfo(chk����).Value = IIf(Nvl(rsTmp!�����־, 0) = 0, 0, 1)
    If chkInfo(chk����).Value = 1 Then
        cboinfo(cbo����Ex).Text = Decode(Nvl(rsTmp!�����־, 0), 1, "��", 2, "��", 3, "��", 4, "��", 9, "����")
        txtInfo(txt��������).Text = IIf(Nvl(rsTmp!�����־, 0) = 9, "", Nvl(rsTmp!��������, 0))
    End If
    
    Call GetCboIndex(cboinfo(cbo����ҽʦ), Nvl(rsTmp!����ҽʦ))
    If Not IsNull(rsTmp!����ҽʦ) And cboinfo(cbo����ҽʦ).ListIndex = -1 Then Call SetCboFromName(rsTmp!����ҽʦ, cboinfo(cbo����ҽʦ))
    
    Call GetCboIndex(cboinfo(cboסԺҽʦ), Nvl(rsTmp!סԺҽʦ))
    If Not IsNull(rsTmp!סԺҽʦ) And cboinfo(cboסԺҽʦ).ListIndex = -1 Then Call SetCboFromName(rsTmp!סԺҽʦ, cboinfo(cboסԺҽʦ))
    
    Call GetCboIndex(cboinfo(cbo���λ�ʿ), Nvl(rsTmp!���λ�ʿ))
    If Not IsNull(rsTmp!���λ�ʿ) And cboinfo(cbo���λ�ʿ).ListIndex = -1 Then Call SetCboFromName(rsTmp!���λ�ʿ, cboinfo(cbo���λ�ʿ))
    
    '����������  δ֪ ��Ϊ ����
    Call GetCboIndex(cboinfo(cboѪ��), IIf(Nvl(rsTmp!Ѫ��) = "δ֪", "����", Nvl(rsTmp!Ѫ��)))
    '�������
    txtInfo(txt���).Text = IIf(rsTmp!��� & "" = "0", "", rsTmp!��� & "")
    txtInfo(txt����).Text = IIf(rsTmp!���� & "" = "0", "", rsTmp!���� & "")
    
    '��Ժ��ʽ
    Call GetCboIndex(cboinfo(cbo��Ժ��ʽ), Nvl(rsTmp!��Ժ��ʽ))
    
    '���ʱ��
    If Nvl(rsTmp!״̬, 0) = 1 Then
        txtInfo(txt���ʱ��).Text = "��δ���"
    Else
        StrSQL = "Select ��ʼʱ�� From ���˱䶯��¼" & _
            " Where ����ID=[1] And ��ҳID=[2] And ��ʼԭ�� IN(2,1) And ��ʼʱ�� is Not Null Order by ��ʼԭ�� Desc"
        Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
        If Not rsTmp.EOF Then
            txtInfo(txt���ʱ��).Text = Format(rsTmp!��ʼʱ��, "yyyy-MM-dd HH:mm")
        End If
    End If
    
    '�����ӱ���
    '---------------------------------------------------------------
    StrSQL = "Select a.����ID,a.��ҳID,a.��Ϣ��,a.��Ϣֵ,b.���� From ������ҳ�ӱ� a " & _
            ",������Ŀ b" & " where a.��Ϣ��=b.����(+) And a.����ID=[1] And a.��ҳID=[2] Order by a.��Ϣ��"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    For i = 1 To rsTmp.RecordCount
        Select Case UCase(Nvl(rsTmp!��Ϣ��))
            Case "��������"
                Call GetCboIndex(cboinfo(cbo��������), Nvl(rsTmp!��Ϣֵ))
                If cboinfo(cbo��������).ListIndex = -1 And Not IsNull(rsTmp!��Ϣֵ) Then    '����ϵͳ��ǰ���ܶ����в��淶��ֵ
                    cboinfo(cbo��������).AddItem rsTmp!��Ϣֵ
                    cboinfo(cbo��������).ListIndex = cboinfo(cbo��������).NewIndex
                End If
            Case "��Ժ����"
                txtInfo(txt��Ժ����).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��Ժ����"
                txtInfo(txt��Ժ����).Text = Nvl(rsTmp!��Ϣֵ)
            Case "ת�Ƽ�¼"
                varTmp = Split(Nvl(rsTmp!��Ϣֵ), ",")
                If UBound(varTmp) >= 0 Then txtInfo(txtת��1).Text = varTmp(0)
                If UBound(varTmp) >= 1 Then txtInfo(txtת��2).Text = varTmp(1)
                If UBound(varTmp) >= 2 Then txtInfo(txtת��3).Text = varTmp(2)
            Case UCase("HBsAg")
                Call GetCboIndex(cboinfo(cboHBsAg), Nvl(rsTmp!��Ϣֵ))
            Case UCase("HCV-Ab")
                Call GetCboIndex(cboinfo(cboHCVAb), Nvl(rsTmp!��Ϣֵ))
            Case UCase("HIV-Ab")
                Call GetCboIndex(cboinfo(cboHIVAb), Nvl(rsTmp!��Ϣֵ))
            Case "��ҽΣ��"
                chkInfo(chkΣ��).Value = Val(Nvl(rsTmp!��Ϣֵ, 0))
            Case "��ҽ��֢"
                chkInfo(chk��֢).Value = Val(Nvl(rsTmp!��Ϣֵ, 0))
            Case "��ҽ����"
                chkInfo(chk����).Value = Val(Nvl(rsTmp!��Ϣֵ, 0))
            Case "��ҽ���ȷ���"
                Call GetCboIndex(cboinfo(cbo���ȷ���), Nvl(rsTmp!��Ϣֵ))
            Case "������ҩ�Ƽ�"
                Call GetCboIndex(cboinfo(cbo������ҩ), Nvl(rsTmp!��Ϣֵ))
            Case "��������ԭ��"
                txtInfo(txt����ԭ��).Text = Nvl(rsTmp!��Ϣֵ)
            Case "����ʱ��"
                If IsNull(rsTmp!��Ϣֵ) Then
                    txt����ʱ��.Text = "____-__-__ __:__:__"
                ElseIf Not IsDate(rsTmp!��Ϣֵ) Then
                    txt����ʱ��.Text = "____-__-__ __:__:__"
                Else
                    txt����ʱ��.Text = rsTmp!��Ϣֵ
                End If
            Case "��Ժǰ����Ժ����"
                chkInfo(chk����Ժ����).Value = Val(Nvl(rsTmp!��Ϣֵ, 0))
            Case "ʾ�̲���"
                chkInfo(chkʾ�̲���).Value = Val(Nvl(rsTmp!��Ϣֵ, 0))
            Case "���в���"
                chkInfo(chk���в���).Value = Val(Nvl(rsTmp!��Ϣֵ, 0))
            Case "���Ѳ���"
                chkInfo(chk���Ѳ���).Value = Val(Nvl(rsTmp!��Ϣֵ))
            Case UCase("Rh")
                '���������ݣ�δ�� ��Ϊ δ��
                Call GetCboIndex(cboinfo(cboRh), IIf(Nvl(rsTmp!��Ϣֵ) = "δ��", "δ��", Nvl(rsTmp!��Ϣֵ)))
            Case "��Ѫ��Ӧ"
                cboinfo(cbo��Ѫ��Ӧ).ListIndex = Val(Nvl(rsTmp!��Ϣֵ, 0))
            Case "���ϸ��"
                txtInfo(txt���ϸ��).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��ѪС��"
                txtInfo(txt��ѪС��).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��Ѫ��"
                txtInfo(txt��Ѫ��).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��ȫѪ"
                txtInfo(txt��ȫѪ).Text = Nvl(rsTmp!��Ϣֵ)
            Case "������"
                txtInfo(txt������).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��Һ��Ӧ"
                Call GetCboIndex(cboinfo(cbo��Һ��Ӧ), Nvl(rsTmp!��Ϣֵ))
            Case "ҽѧ��ʾ"
                txtInfo(txtҽѧ��ʾ).Text = Nvl(rsTmp!��Ϣֵ)
            Case "����ҽѧ��ʾ"
                txtInfo(txt����ҽѧ��ʾ).Text = Nvl(rsTmp!��Ϣֵ)
            Case "������"
                Call GetCboIndex(cboinfo(cbo������), Nvl(rsTmp!��Ϣֵ))
                If Not IsNull(rsTmp!��Ϣֵ) And cboinfo(cbo������).ListIndex = -1 Then Call SetCboFromName(rsTmp!��Ϣֵ, cboinfo(cbo������))
            Case "����ҽʦ"
                Call GetCboIndex(cboinfo(cbo����ҽʦ), Nvl(rsTmp!��Ϣֵ))
                If Not IsNull(rsTmp!��Ϣֵ) And cboinfo(cbo����ҽʦ).ListIndex = -1 Then Call SetCboFromName(rsTmp!��Ϣֵ, cboinfo(cbo����ҽʦ))
            Case "����ҽʦ"
                Call GetCboIndex(cboinfo(cbo����ҽʦ), Nvl(rsTmp!��Ϣֵ))
                If Not IsNull(rsTmp!��Ϣֵ) And cboinfo(cbo����ҽʦ).ListIndex = -1 Then Call SetCboFromName(rsTmp!��Ϣֵ, cboinfo(cbo����ҽʦ))
            Case "����ҽʦ"
                Call GetCboIndex(cboinfo(cbo����ҽʦ), Nvl(rsTmp!��Ϣֵ))
                If Not IsNull(rsTmp!��Ϣֵ) And cboinfo(cbo����ҽʦ).ListIndex = -1 Then Call SetCboFromName(rsTmp!��Ϣֵ, cboinfo(cbo����ҽʦ))
            Case "�о���ʵϰҽʦ"
                Call GetCboIndex(cboinfo(cbo�о���ҽʦ), Nvl(rsTmp!��Ϣֵ))
                If Not IsNull(rsTmp!��Ϣֵ) And cboinfo(cbo�о���ҽʦ).ListIndex = -1 Then Call SetCboFromName(rsTmp!��Ϣֵ, cboinfo(cbo�о���ҽʦ))
            Case "ʵϰҽʦ"
                Call GetCboIndex(cboinfo(cboʵϰҽʦ), Nvl(rsTmp!��Ϣֵ))
                If Not IsNull(rsTmp!��Ϣֵ) And cboinfo(cboʵϰҽʦ).ListIndex = -1 Then Call SetCboFromName(rsTmp!��Ϣֵ, cboinfo(cboʵϰҽʦ))
            Case "�ʿ�ҽʦ"
                Call GetCboIndex(cboinfo(cbo�ʿ�ҽʦ), Nvl(rsTmp!��Ϣֵ))
                If Not IsNull(rsTmp!��Ϣֵ) And cboinfo(cbo�ʿ�ҽʦ).ListIndex = -1 Then Call SetCboFromName(rsTmp!��Ϣֵ, cboinfo(cbo�ʿ�ҽʦ))
            Case "�ʿػ�ʿ"
                Call GetCboIndex(cboinfo(cbo�ʿػ�ʿ), Nvl(rsTmp!��Ϣֵ))
                If Not IsNull(rsTmp!��Ϣֵ) And cboinfo(cbo�ʿػ�ʿ).ListIndex = -1 Then Call SetCboFromName(rsTmp!��Ϣֵ, cboinfo(cbo�ʿػ�ʿ))
            Case "��ԭѧ���"
                chkInfo(chk��ԭѧ).Value = Val(Nvl(rsTmp!��Ϣֵ, 0))
            Case "��Ѫ���"
                Call GetCboIndex(cboinfo(cbo��Ѫ���), Nvl(rsTmp!��Ϣֵ))
            Case "CT"
                chkInfo(chkCT).Value = Val(Nvl(rsTmp!��Ϣֵ, 0))
            Case "MRI"
                chkInfo(chkMRI).Value = Val(Nvl(rsTmp!��Ϣֵ, 0))
            Case "��ɫ������"
                chkInfo(chk������).Value = Val(Nvl(rsTmp!��Ϣֵ, 0))
            Case "������4"
                vsTSJC.TextMatrix(vsTSJC.FixedRows + 0, 1) = Nvl(rsTmp!��Ϣֵ)
                vsTSJC.Cell(flexcpData, vsTSJC.FixedRows + 0, 1) = vsTSJC.TextMatrix(vsTSJC.FixedRows + 0, 1)
            Case "������5"
                vsTSJC.TextMatrix(vsTSJC.FixedRows + 1, 1) = Nvl(rsTmp!��Ϣֵ)
                vsTSJC.Cell(flexcpData, vsTSJC.FixedRows + 1, 1) = vsTSJC.TextMatrix(vsTSJC.FixedRows + 1, 1)
            Case "������6"
                vsTSJC.TextMatrix(vsTSJC.FixedRows + 2, 1) = Nvl(rsTmp!��Ϣֵ)
                vsTSJC.Cell(flexcpData, vsTSJC.FixedRows + 2, 1) = vsTSJC.TextMatrix(vsTSJC.FixedRows + 2, 1)
            Case "��Ժת��"
                txtInfo(txt��Ժת��).Text = Nvl(rsTmp!��Ϣֵ)
            Case "ѹ�������ڼ�"
                cboinfo(cboѹ�������ڼ�).Text = Nvl(rsTmp!��Ϣֵ, " ")
            Case "ѹ������"
                cboinfo(cboѹ������).Text = Nvl(rsTmp!��Ϣֵ, " ")
            Case "������׹���˺�"
                cboinfo(cbo������׹���˺�).Text = Nvl(rsTmp!��Ϣֵ, " ")
            Case "������׹��ԭ��"
                cboinfo(cbo������׹��ԭ��).Text = Nvl(rsTmp!��Ϣֵ, " ")
            Case "31������סԺ"
                If Nvl(rsTmp!��Ϣֵ) <> "" Then
                    optInput(opt31����).Value = True
                    txtInfo(txt31��Ŀ��).Text = Nvl(rsTmp!��Ϣֵ)
                    txtInfo(txt31��Ŀ��).Enabled = True
                End If
            Case "����Ժ�ƻ�����"
                cboinfo(cbo31���7������Ժ).ListIndex = Val(Nvl(rsTmp!��Ϣֵ))
            Case "������������"
                Call LoadOldData("" & rsTmp!��Ϣֵ, txtӤ������)
            Case "��������������"
                txtInfo(txt����������).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��������Ժ����"
                txtInfo(txt��������Ժ����).Text = Nvl(rsTmp!��Ϣֵ)
            Case "������ʹ��ʱ��"
                txtInfo(txt������Сʱ).Text = Nvl(rsTmp!��Ϣֵ)
            Case "����ʱ��"
                '�����ʽ:��Ժǰ(�죬Сʱ,����)|��Ժ��(�죬Сʱ,����)
                txtInfo(txt��Ժǰ��).Text = Split(Split(Nvl(rsTmp!��Ϣֵ), "|")(0) & ",", ",")(0)
                txtInfo(txt��ԺǰСʱ).Text = Split(Split(Nvl(rsTmp!��Ϣֵ), "|")(0) & ",", ",")(1)
                txtInfo(txt��Ժǰ����).Text = Split(Split(Nvl(rsTmp!��Ϣֵ), "|")(0) & ",", ",")(2)
                txtInfo(txt��Ժ����).Text = Split(Split(Nvl(rsTmp!��Ϣֵ), "|")(1) & ",", ",")(0)
                txtInfo(txt��Ժ��Сʱ).Text = Split(Split(Nvl(rsTmp!��Ϣֵ) & "|", "|")(1) & ",", ",")(1)
                txtInfo(txt��Ժ�����).Text = Split(Split(Nvl(rsTmp!��Ϣֵ) & "|", "|")(1) & ",", ",")(2)
            Case "���Ȳ���"
                txtInfo(txt����ԭ��).Text = Nvl(rsTmp!��Ϣֵ)
            Case "�������"
                txtInfo(txt�������).Text = Nvl(rsTmp!��Ϣֵ)
            Case "����"
                txtInfo(txt����).Text = Nvl(rsTmp!��Ϣֵ)
            Case "����������"
                If Nvl(rsTmp!��Ϣֵ) <> "" Then
                    Call GetCboIndex(cboinfo(cbo����������), Nvl(rsTmp!��Ϣֵ))
                End If
            Case "�ֻ��̶�"
                If Nvl(rsTmp!��Ϣֵ) <> "" Then
                    Call GetCboIndex(cboinfo(cbo�ֻ��̶�), Nvl(rsTmp!��Ϣֵ))
                End If
            Case "��ҽ�豸"
                Call GetCboIndex(cboinfo(cboʹ����ҽ�����豸), Nvl(rsTmp!��Ϣֵ))
            Case "��ҽ����"
                Call GetCboIndex(cboinfo(cboʹ����ҽ���Ƽ���), Nvl(rsTmp!��Ϣֵ))
            Case "��֤ʩ��"
                Call GetCboIndex(cboinfo(cbo��֤ʩ��), Nvl(rsTmp!��Ϣֵ))
            Case "�����"
                txtInfo(txt�����).Text = Nvl(rsTmp!��Ϣֵ)
            Case "��������"
                Call GetCboIndex(cboinfo(cbo��������), Nvl(rsTmp!��Ϣֵ))
            Case "��ҳ��������"
                txtInfo(txt�ʿ�����).Text = Nvl(rsTmp!��Ϣֵ)
            Case "�没�ز�Σ"
                chkInfo(chkסԺ�ڼ�没�ػ�Σ).Value = Val(Nvl(rsTmp!��Ϣֵ))
            Case "�ٴ�·��"
                chkInfo(chk����·��).Value = Val(Nvl(rsTmp!��Ϣֵ))
            Case "�˳�ԭ��"
                If Nvl(rsTmp!��Ϣֵ) = "1" Then
                    chkInfo(chk���·��).Value = 1
                Else
                    chkInfo(chk���·��).Value = 0
                    txtInfo(txt�˳�ԭ��).Text = Nvl(rsTmp!��Ϣֵ)
                End If
            Case "����ԭ��"
                If Nvl(rsTmp!��Ϣֵ) = "0" Then
                    chkInfo(chk����).Value = 0
                Else
                    chkInfo(chk����).Value = 1
                    txtInfo(txt����ԭ��).Text = Trim(Nvl(rsTmp!��Ϣֵ))
                End If
            Case "����Լ��"
                chkInfo(chk�Ƿ�ʹ������Լ��).Value = Val(Nvl(rsTmp!��Ϣֵ))
            Case "Լ����ʱ��"
                txtInfo(txtԼ����ʱ��).Text = Nvl(rsTmp!��Ϣֵ)
            Case "Լ����ʽ"
                Call GetCboIndex(cboinfo(cboԼ����ʽ), Nvl(rsTmp!��Ϣֵ))
            Case "Լ������"
                Call GetCboIndex(cboinfo(cboԼ������), Nvl(rsTmp!��Ϣֵ))
            Case "Լ��ԭ��"
                Call GetCboIndex(cboinfo(cboԼ��ԭ��), Nvl(rsTmp!��Ϣֵ))
            Case "��������Ժ��ʽ"
                Call GetCboIndex(cboinfo(cbo��������Ժ��ʽ), Nvl(rsTmp!��Ϣֵ))
            Case "Χ��������"
                chkInfo(chkΧ��������).Value = Val(Nvl(rsTmp!��Ϣֵ))
            Case "�������"
                chkInfo(chk�������).Value = Val(Nvl(rsTmp!��Ϣֵ))
            Case "����״��"
                Call GetCboIndex(cboinfo(cbo����״��), Nvl(rsTmp!��Ϣֵ))
            Case "����ʱ��"
                If Nvl(rsTmp!��Ϣֵ) <> "" Then
                    txt��������.Text = Format(rsTmp!��Ϣֵ, "yyyy-MM-dd")
                    If Format(rsTmp!��Ϣֵ, "HH:mm") <> "00:00" Then
                        txt����ʱ��.Text = Format(rsTmp!��Ϣֵ, "HH:mm")
                    End If
                End If
            Case Else
                '�������������
                If Left(Nvl(rsTmp!��Ϣ��), 3) = "������" And Not IsNull(rsTmp!��Ϣֵ) Then
                    With vsKSS
                        For j = .FixedRows To .Rows - 1
                            If .TextMatrix(j, 1) = "" Then
                                '���������ݣ�����ҳ�ӱ����ȶ�����
                                .RowData(j) = GetIDTmp(rsTmp!��Ϣֵ)
                                If .RowData(j) <> 0 Then
                                    .TextMatrix(j, 1) = rsTmp!��Ϣֵ
                                    .Cell(flexcpData, j, 1) = .TextMatrix(j, 1)
                                End If
                                Exit For
                            End If
                        Next
                        If j > .Rows - 1 Then
                            .AddItem ""
                             '���������ݣ�����ҳ�ӱ����ȶ�����
                            .RowData(.Rows - 1) = GetIDTmp(rsTmp!��Ϣֵ)
                            If .RowData(.Rows - 1) <> 0 Then
                                .TextMatrix(.Rows - 1, 1) = rsTmp!��Ϣֵ
                                .Cell(flexcpData, .Rows - 1, 1) = .TextMatrix(.Rows - 1, 1)
                            End If
                        End If
                    End With
                Else
                    '������Ŀ
                    If Not IsNull(rsTmp("����")) Then
                        With vsfMain
                            For lngCol = 0 To vsfMain.Cols - 1 Step 3
                                lngRow = vsfMain.FindRow(rsTmp("��Ϣ��"), , lngCol)
                                If lngRow >= 0 Then
                                    If vsfMain.TextMatrix(lngRow, lngCol) = rsTmp("��Ϣ��") Then
                                        If vsfMain.TextMatrix(lngRow, lngCol + 2) = "�Ƿ�" Then
                                            vsfMain.Cell(flexcpChecked, lngRow, lngCol + 1) = IIf(rsTmp("��Ϣֵ") = 0, 2, 1)
                                            Exit For
                                        Else
                                            vsfMain.TextMatrix(lngRow, lngCol + 1) = rsTmp("��Ϣֵ") & ""
                                            Exit For
                                        End If
                                    End If
                                End If
                            Next lngCol
                        End With
                    End If
                End If
                Call SetKSSSerial
        End Select
        rsTmp.MoveNext
    Next
    
    '�Զ���ȡת�ƿ��Ҽ��������(�����)
    '---------------------------------------------------------------
    If txtInfo(txtת��1).Text = "" And txtInfo(txtת��2).Text = "" And txtInfo(txtת��3).Text = "" Then
        StrSQL = _
            " Select B.����" & _
            " From ���˱䶯��¼ A,���ű� B" & _
            " Where A.����ID=[1] And A.��ҳID=[2]" & _
            " And A.����ID=B.ID And A.��ʼԭ��=3 And A.��ʼʱ�� is Not NULL" & _
            " Order by A.��ʼʱ��"
        Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
        For i = 1 To rsTmp.RecordCount
            If i = 1 Then
                txtInfo(txtת��1).Text = rsTmp!����
            ElseIf i = 2 Then
                txtInfo(txtת��2).Text = rsTmp!����
            ElseIf i = 3 Then
                txtInfo(txtת��3).Text = rsTmp!����
                Exit For
            End If
            rsTmp.MoveNext
        Next
    End If
    
    If txtInfo(txt��Ժ����).Text = "" Then
        StrSQL = "Select B.�����" & _
            " From ������ҳ A,��λ״����¼ B" & _
            " Where A.����ID=[1] And A.��ҳID=[2]" & _
            " And A.��Ժ����ID=B.����ID And A.��Ժ����=B.����"
        Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
        If Not rsTmp.EOF Then txtInfo(txt��Ժ����).Text = Nvl(rsTmp!�����)
    End If
    If txtInfo(txt��Ժ����).Text = "" Then
        StrSQL = "Select B.�����" & _
            " From ������ҳ A,��λ״����¼ B" & _
            " Where A.����ID=[1] And A.��ҳID=[2]" & _
            " And A.��ǰ����ID=B.����ID And A.��Ժ����=B.����"
        Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
        If Not rsTmp.EOF Then txtInfo(txt��Ժ����).Text = Nvl(rsTmp!�����)
    End If
    
    '������Ϣ:����סԺ��,������
    '---------------------------------------------------------------
    StrSQL = "Select ��¼��Դ,NVL(����ʱ��,��¼ʱ��) as ����ʱ��,ҩ��ID,ҩ����,������Ӧ From ���˹�����¼ A" & _
        " Where ���=1 And ����ID=[1] And ��ҳID=[2]" & _
        " And Not Exists(Select ҩ��ID From ���˹�����¼" & _
            " Where (Nvl(ҩ��ID,0)=Nvl(A.ҩ��ID,0) Or Nvl(ҩ����,'Null')=Nvl(A.ҩ����,'Null'))" & _
            " And Nvl(���,0)=0 And ��¼ʱ��>A.��¼ʱ�� And ����ID=[1] And ��ҳID=[2])" & _
        " Order by NVL(����ʱ��,��¼ʱ��),ҩ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    If Not rsTmp.EOF Then
        rsTmp.Filter = "��¼��Դ=3" '��ҳ������д��
        If rsTmp.EOF Then rsTmp.Filter = "��¼��Դ<>3" '������Դ����Ϊȱʡ��ʾ
        With vsAller
            .Rows = rsTmp.RecordCount + 2 '�̶���+����
            For i = 1 To rsTmp.RecordCount
                '������Դ�Ŀ������ظ�
                lngRow = -1
                If Not IsNull(rsTmp!ҩ��ID) Then
                    lngRow = .FindRow(CLng(rsTmp!ҩ��ID))
                ElseIf Not IsNull(rsTmp!ҩ����) Then
                    lngRow = .FindRow(CStr(rsTmp!ҩ����), , AC_����ҩ��)
                End If
                If lngRow = -1 Then
                    .TextMatrix(i, AC_����ʱ��) = Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm")
                    .Cell(flexcpData, i, AC_����ʱ��) = Format(rsTmp!����ʱ��, "yyyy-MM-dd HH:mm")  '���ڱ���
                    .TextMatrix(i, AC_����ҩ��) = Nvl(rsTmp!ҩ����)
                    .Cell(flexcpData, i, AC_����ҩ��) = .TextMatrix(i, AC_����ҩ��) '��������ָ�
                    .TextMatrix(i, AC_������Ӧ) = Nvl(rsTmp!������Ӧ)
                    .Cell(flexcpData, i, AC_������Ӧ) = .TextMatrix(i, AC_������Ӧ)   '��������ָ�

                End If
                rsTmp.MoveNext
            Next
        End With
    End If
    vsAller.Row = 1: vsAller.Col = AC_����ҩ��
    vsAller.Tag = "δ�޸�"
    
    '��ҽ���
    '---------------------------------------------------------------
    str���ƽ�� = Get���ƽ��
    vsDiagXY.ColData(col��Ժ���) = str���ƽ��
    
    '�ж���ҳ�Ƿ�������
    StrSQL = "Select 1 From ������ϼ�¼ Where ����ID=[1] And ��ҳID=[2] And ��¼��Դ=3  And RowNum<2"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    mbln��ҳ��� = rsTmp.RecordCount > 0
    If mbln��ҳ��� Then
        strTmp = " and a.��¼��Դ=3 "
    Else
        strTmp = " And a.��¼��Դ IN(1,2,3,4) "
    End If
    'ȱʡ����ʼ��
    With vsDiagXY
        '1-��ҽ�������;2-��ҽ��Ժ���;3-��Ժ���(�������);5-Ժ�ڸ�Ⱦ;6-�������;7-�����ж���;10-����֢
        .TextMatrix(1, col����) = 1
        .TextMatrix(2, col����) = 2
        .TextMatrix(3, col����) = 3
        .TextMatrix(4, col����) = 3
        .TextMatrix(5, col����) = 5
        .TextMatrix(6, col����) = 10
        .TextMatrix(7, col����) = 6
        .TextMatrix(8, col����) = 7
    End With
    
    '��ȡ������Դ�����
    StrSQL = "Select a.��ע,a.ID,a.����ID,a.��ҳID,a.ҽ��ID,a.��¼��Դ,a.��ϴ���,a.�������,a.����ID,a.�������,a.����ID,a.��Ժ����," & _
        " a.���ID,a.֤��ID,a.�������,a.��Ժ���,a.�Ƿ�δ��,a.�Ƿ�����,a.��¼����,a.��¼��,a.ȡ��ʱ��,a.ȡ����,a.����ID, b.���� As ��������, c.���� As ��ϱ��� " & _
        " From ������ϼ�¼ A, ��������Ŀ¼ B, �������Ŀ¼ C" & _
        " Where a.����id = b.Id(+) And a.���id = c.Id(+)  And a.������� IN(1,2,3,5,6,7,10,21)" & _
        strTmp & _
        " And a.ȡ��ʱ�� is Null And a.����ID=[1] And a.��ҳID=[2]" & _
        " Order by a.�������,a.��¼��Դ Desc,a.��ϴ���,a.ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    Set mrsXYDiag = zlDatabase.CopyNewRec(rsTmp)
    If Not rsTmp.EOF Then
        With vsDiagXY
            StrSQL = "1,2,3,5,6,7,10,21"
            For i = 0 To UBound(Split(StrSQL, ","))
                rsTmp.Filter = "��¼��Դ=3 And �������=" & Split(StrSQL, ",")(i)
                If Val(Split(StrSQL, ",")(i)) <> 21 Then
                    If rsTmp.EOF Then
                        rsTmp.Filter = "��¼��Դ=2 And �������=" & Split(StrSQL, ",")(i)
                    End If
                    If rsTmp.EOF Then
                        rsTmp.Filter = "��¼��Դ=1 And �������=" & Split(StrSQL, ",")(i)
                    End If
                End If
                If rsTmp.EOF Then
                    rsTmp.Filter = "��¼��Դ=4 And �������=" & Split(StrSQL, ",")(i)
                End If
                
                If Val(Split(StrSQL, ",")(i)) = 21 Then
                    '21-��ԭѧ���
                    If Not rsTmp.EOF Then
                        txtInfo(txt��ԭѧ).Text = Nvl(rsTmp!�������)
                        txtInfo(txt��ԭѧ).Tag = txtInfo(txt��ԭѧ).Text
                        cmdInfo(txt��ԭѧ).Tag = Nvl(rsTmp!����id, 0)
                    End If
                Else
                    Do While Not rsTmp.EOF
                        If Val("" & rsTmp!��¼��Դ) = 3 And Val("" & rsTmp!�������) = 2 And Val("" & rsTmp!��ϴ���) = 1 Then
                            mstrXYDiagInfo = "" & rsTmp!�������
                        End If
                        'ȷ����ǰ��ʾ��
                        lngRow = .FindRow(CStr(Split(StrSQL, ",")(i)), , col����)
                        For j = lngRow To .Rows - 1
                            If Val(.TextMatrix(j, col����)) = Val(Split(StrSQL, ",")(i)) Then
                                lngRow = j
                                If .TextMatrix(j, col�������) = "" Then Exit For
                            Else
                                Exit For
                            End If
                        Next
                        If .TextMatrix(lngRow, col�������) <> "" Then
                            lngRow = lngRow + 1: .AddItem "", lngRow
                            .TextMatrix(lngRow, col����) = Split(StrSQL, ",")(i)
                        End If
                        
                        If IsNull(rsTmp!�������) Then
                            .TextMatrix(lngRow, col��ϱ���) = ""
                            .TextMatrix(lngRow, col�������) = ""
                        Else
                            If Mid(rsTmp!�������, 1, 1) <> "(" Or (Val(rsTmp!���id & "") = 0 And Val(rsTmp!����id & "") = 0) Then '��ҽ���������������ˣ���֢��������ֻ�жϵ�һ���ַ�
                                '���ڼ����������Ͽ��Զ�Ӧ�������������Ϊ�յ�ʱ�����жϼ������룬��ȡ��������
                                If Val(rsTmp!����id & "") <> 0 Then
                                    .TextMatrix(lngRow, col��ϱ���) = Nvl(rsTmp!��������)
                                ElseIf Val(rsTmp!���id & "") <> 0 Then
                                    .TextMatrix(lngRow, col��ϱ���) = Nvl(rsTmp!��ϱ���)
                                Else
                                    .TextMatrix(lngRow, col��ϱ���) = ""
                                End If
                                .TextMatrix(lngRow, col�������) = rsTmp!�������
                            Else
                                .TextMatrix(lngRow, col��ϱ���) = Mid(rsTmp!�������, 2, InStr(rsTmp!�������, ")") - 2)
                                .TextMatrix(lngRow, col�������) = Mid(rsTmp!�������, InStr(rsTmp!�������, ")") + 1)
                            End If
                        End If
                        If Not IsNull(rsTmp!����id) Or Not IsNull(rsTmp!���id) Then
                            .Cell(flexcpData, lngRow, col�������) = Get�������(Val("" & rsTmp!���id), Val("" & rsTmp!����id))    '��ȡԭʼ�����Ա��޸�ʱ�ж�
                        Else
                            .Cell(flexcpData, lngRow, col�������) = .TextMatrix(lngRow, col�������)
                        End If
                        
                        '�ֻ��̶Ⱥ�����������
                        If Val("" & rsTmp!�������) = 3 And Val("" & rsTmp!��ϴ���) = 1 Then
                            If Trim(Nvl(rsTmp!��������)) = "" Then
                                bln�ֻ��̶� = False
                            Else
                                bln�ֻ��̶� = ((InStr("C", UCase(Left(Nvl(rsTmp!��������), 1)))) > 0) Or ((InStr("D0", UCase(Left(Nvl(rsTmp!��������), 2)))) > 0) Or ((InStr("D32.,D33.,", UCase(Left(Nvl(rsTmp!��������), 4)))) > 0)
                            End If
                        End If
                        
                        cboinfo(cbo�ֻ��̶�).Enabled = bln�ֻ��̶�
                        lblInfo(lbl�ֻ��̶�).Enabled = bln�ֻ��̶�
                        lblInfo(lbl����������).Enabled = bln�ֻ��̶�
                        cboinfo(cbo����������).Enabled = bln�ֻ��̶�
                        .TextMatrix(lngRow, col��ע) = Nvl(rsTmp!��ע)
                       .Cell(flexcpData, lngRow, col�Ƿ�����) = Val(rsTmp!ID & "")
                        .TextMatrix(lngRow, col��Ժ���) = Nvl(rsTmp!��Ժ���)
                        .TextMatrix(lngRow, col��Ժ����) = Nvl(rsTmp!��Ժ����)
                        .TextMatrix(lngRow, col�Ƿ�δ��) = IIf(Nvl(rsTmp!�Ƿ�δ��, 0) = 1, "��", "")
                        .TextMatrix(lngRow, col�Ƿ�����) = IIf(Nvl(rsTmp!�Ƿ�����, 0) = 1, "��", "")
                        .TextMatrix(lngRow, col���ID) = Nvl(rsTmp!���id, 0)
                        .TextMatrix(lngRow, col����ID) = Nvl(rsTmp!����id, 0)
                        rsTmp.MoveNext
                    Loop
                End If
            Next
        End With
    End If
    
    vsDiagXY.Cell(flexcpForeColor, 1, col�Ƿ�����, vsDiagXY.Rows - 1, col�Ƿ�����) = vbRed
    vsDiagXY.Cell(flexcpBackColor, GetRow(3), vsDiagXY.FixedRows, GetRow(3), vsDiagXY.Cols - 1) = &HC0FFC0
    vsDiagXY.Cell(flexcpBackColor, 1, col��ϱ���, vsDiagXY.Rows - 1, col��ϱ���) = ColorUnEditCell      '����ɫ
    vsDiagXY.Row = 1: vsDiagXY.Col = col�������
    Call vsDiagXY_AfterRowColChange(-1, -1, vsDiagXY.Row, vsDiagXY.Col)
    vsDiagXY.Tag = "δ�޸�"
    If vsDiagXY.TextMatrix(GetRow(6), col�������) <> "" Then
        txtInfo(txt�����).Enabled = True
        txtInfo(txt�����).BackColor = vbWindowBackground
    End If
        
    '��ҽ���
    '---------------------------------------------------------------
    vsDiagZY.ColData(col��Ժ���) = str���ƽ��
    
    'ȱʡ����ʼ��
    With vsDiagZY
        '11-��ҽ�������;12-��ҽ��Ժ���;13-��ҽ��Ժ���(��Ҫ��ϡ��������)
        .TextMatrix(1, colzy����) = 11
        .TextMatrix(2, colzy����) = 12
        .TextMatrix(3, colzy����) = 13
        .TextMatrix(4, colzy����) = 13
    End With
    
    If mbln��ҳ��� Then
        strTmp = " and a.��¼��Դ=3 "
    Else
        strTmp = " And a.��¼��Դ IN(1,2,3,4) "
    End If
    
    '��ȡ������Դ�����
    StrSQL = "Select a.��ע, a.Id, a.����id, a.��ҳid, a.ҽ��id, a.��¼��Դ, a.��ϴ���, a.�������, a.����id, a.�������,a.��Ժ����," & _
        " a.����id, a.���id, a.֤��id, a.�������,a.��Ժ���, a.�Ƿ�δ��, a.�Ƿ�����, a.��¼����, a.��¼��, a.ȡ��ʱ��," & _
        " a.ȡ����, a.����id, b.���� As ��������, c.���� As ��ϱ���,d.���� as ֤����� From ������ϼ�¼ A, ��������Ŀ¼ B, �������Ŀ¼ C,��������Ŀ¼ D" & _
        " Where a.����id = b.Id(+) And a.���id = c.Id(+) And a.֤��ID=d.ID(+) And a.������� IN(11,12,13)" & _
        strTmp & _
        " And ȡ��ʱ�� Is Null And ����ID=[1] And ��ҳID=[2]" & _
        " Order by a.�������,a.��¼��Դ Desc,a.��ϴ���,a.�������,a.ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    strTmp = ""
    Set mrsZYDiag = zlDatabase.CopyNewRec(rsTmp)
    If Not rsTmp.EOF Then
        With vsDiagZY
            StrSQL = "11,12,13"
            For i = 0 To UBound(Split(StrSQL, ","))
                rsTmp.Filter = "��¼��Դ=3 And �������=" & Split(StrSQL, ",")(i)
                If rsTmp.EOF Then
                    rsTmp.Filter = "��¼��Դ=2 And �������=" & Split(StrSQL, ",")(i)
                End If
                If rsTmp.EOF Then
                    rsTmp.Filter = "��¼��Դ=1 And �������=" & Split(StrSQL, ",")(i)
                End If
                If rsTmp.EOF Then
                    rsTmp.Filter = "��¼��Դ=4 And �������=" & Split(StrSQL, ",")(i)
                End If
                
                Do While Not rsTmp.EOF
                    If Val("" & rsTmp!��¼��Դ) = 3 And Val("" & rsTmp!�������) = 12 And Val("" & rsTmp!��ϴ���) = 1 Then
                        mstrZYDiagInfo = "" & rsTmp!�������
                    End If
                    'ȷ����ǰ��ʾ��
                    lngRow = .FindRow(CStr(Split(StrSQL, ",")(i)), , colzy����)
                    For j = lngRow To .Rows - 1
                        If Val(.TextMatrix(j, colzy����)) = Val(Split(StrSQL, ",")(i)) Then
                            lngRow = j
                            If .TextMatrix(j, col�������) = "" Then Exit For
                        Else
                            Exit For
                        End If
                    Next
                    If .TextMatrix(lngRow, col�������) <> "" Then
                        lngRow = lngRow + 1: .AddItem "", lngRow
                        .TextMatrix(lngRow, colzy����) = Split(StrSQL, ",")(i)
                    End If
                    
                    If IsNull(rsTmp!�������) Then
                        .TextMatrix(lngRow, col��ϱ���) = ""
                        .TextMatrix(lngRow, col�������) = ""
                    Else
                        If Mid(rsTmp!�������, 1, 1) <> "(" Or (Val(rsTmp!���id & "") = 0 And Val(rsTmp!����id & "") = 0) Then     '��ҽ���������������ˣ���֢��������ֻ�жϵ�һ���ַ�
                            '���ڼ����������Ͽ��Զ�Ӧ�������������Ϊ�յ�ʱ�����жϼ������룬��ȡ��������
                            If Val(rsTmp!����id & "") <> 0 Then
                                .TextMatrix(lngRow, col��ϱ���) = Nvl(rsTmp!��������)
                            ElseIf Val(rsTmp!���id & "") <> 0 Then
                                .TextMatrix(lngRow, col��ϱ���) = Nvl(rsTmp!��ϱ���)
                            Else
                                .TextMatrix(lngRow, col��ϱ���) = ""
                            End If
                            .TextMatrix(lngRow, col�������) = rsTmp!�������
                        Else
                            .TextMatrix(lngRow, col��ϱ���) = Mid(rsTmp!�������, 2, InStr(rsTmp!�������, ")") - 2)
                            .TextMatrix(lngRow, col�������) = Mid(rsTmp!�������, InStr(rsTmp!�������, ")") + 1)
                        End If
                    End If
                    If Not IsNull(rsTmp!����id) Or Not IsNull(rsTmp!���id) Then
                        .Cell(flexcpData, lngRow, col�������) = Get�������(Val("" & rsTmp!���id), Val("" & rsTmp!����id))    '��ȡԭʼ�����Ա��޸�ʱ�ж�
                    Else
                        .Cell(flexcpData, lngRow, col�������) = .TextMatrix(lngRow, col�������)
                    End If
                        
                    .TextMatrix(lngRow, col��ע) = Nvl(rsTmp!��ע)
                    .Cell(flexcpData, lngRow, col�Ƿ�����) = Val(rsTmp!ID & "")
                    .TextMatrix(lngRow, col��Ժ���) = Nvl(rsTmp!��Ժ���)
                    .TextMatrix(lngRow, col��Ժ����) = Nvl(rsTmp!��Ժ����)
                    .TextMatrix(lngRow, colzy���ID) = Nvl(rsTmp!���id, 0)
                    .TextMatrix(lngRow, colzy����ID) = Nvl(rsTmp!����id, 0)
                    .TextMatrix(lngRow, colzy֤��ID) = Nvl(rsTmp!֤��id, 0)
                    'ȡ֤������
                    If InStr(.TextMatrix(lngRow, col�������), "(") > 0 And InStr(.TextMatrix(lngRow, col�������), ")") > 0 Then
                        strTmp = Mid(.TextMatrix(lngRow, col�������), InStrRev(.TextMatrix(lngRow, col�������), "(") + 1)
                        strTmp = Mid(strTmp, 1, Len(strTmp) - 1)
                        '��ȡ֤��
                        .TextMatrix(lngRow, col��ҽ֤��) = strTmp
                        'ȥ�����������֤��
                        .TextMatrix(lngRow, col�������) = Mid(.TextMatrix(lngRow, col�������), 1, InStrRev(.TextMatrix(lngRow, col�������), "(") - 1)
                    Else
                       .TextMatrix(lngRow, col��ҽ֤��) = ""
                    End If
                    
                    
                    rsTmp.MoveNext
                Loop
            Next
        End With
    End If
    vsDiagZY.Cell(flexcpBackColor, GetRow(13), vsDiagZY.FixedRows, GetRow(13), vsDiagZY.Cols - 1) = &HC0FFC0
    vsDiagZY.Cell(flexcpBackColor, 1, col��ϱ���, vsDiagZY.Rows - 1, col��ϱ���) = ColorUnEditCell      '����ɫ
    vsDiagZY.Row = 1: vsDiagZY.Col = col�������
    Call vsDiagZY_AfterRowColChange(-1, -1, vsDiagXY.Row, vsDiagXY.Col)
    vsDiagZY.Tag = "δ�޸�"
    
    If Not mbln��ҳ��� Then
        vsDiagZY.Tag = ""
        vsDiagXY.Tag = ""
    End If
    
    '�������
    '---------------------------------------------------------------
    StrSQL = "Select ����,���� From �����п�����"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption)
    If rsTmp.RecordCount > 0 Then
        strTmp = " "
        Do While Not rsTmp.EOF
            strTmp = strTmp & "|" & rsTmp!���� & "-" & rsTmp!����
            rsTmp.MoveNext
        Loop
        vsOPS.ColComboList(col�п�����) = strTmp
    Else
        vsOPS.ColComboList(col�п�����) = " |0-0 / |1-��/��|2-��/��|3-��/��|4-��/����|5-��/��|6-��/��|7-��/��|8-��/����|9-��/��|10-��/��|11-��/��|12-��/����|13-IV/��|14-IV/��|15-IV/��|16-IV/����"
    End If
    'col��������
    StrSQL = "Select ����,���� From ������������"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption)
    If rsTmp.RecordCount > 0 Then
        strTmp = " "
        Do While Not rsTmp.EOF
            strTmp = strTmp & "|" & rsTmp!����
            rsTmp.MoveNext
        Loop
        vsOPS.ColComboList(col��������) = strTmp
    Else
        vsOPS.ColComboList(col��������) = " |����|ȫ��|��Ӳ|����|����|�۴�|����"
    End If
    '�������
    vsOPS.ColComboList(COL�������.COL�������) = " |����|����|����"
    'ASA�ּ�
    vsOPS.ColComboList(COL�������.colASA�ּ�) = " |P1|P2|P3|P4|P5|P6"
    'colNNIS�ּ�
    vsOPS.ColComboList(colNNIS�ּ�) = " |NNIS0��|NNIS1��|NNIS2��|NNIS3��"
    '�����ּ�
    vsOPS.ColComboList(col��������) = " |һ������|��������|��������|�ļ�����"
    vsOPS.ColDataType(col�ٴ�����) = flexDTBoolean
    
    '�׶�ȡ��ҳ�����������
    StrSQL = "Select a.������ʿ,a.�������,a.�п�,a.����,a.��¼����,a.��¼��,a.ȡ��ʱ��,a.ȡ����,NVl(B.����,C.����) as ��������,a.ID,a.����ID,a.��ҳID,a.��¼��Դ,a.��������,a.������ʼʱ��,a.��������ʱ��,a.��������,a.��������ID,a.������ĿID,a.��������,a.����ҽʦ,a.��һ����," & _
    "a.�ڶ�����,a.������ʿ,a.����ʼʱ��,a.�������ʱ��,a.����ʽ,a.��������,a.��������,a.��Һ����,a.����ҽʦ,a.������ʼʱ��,a.��������ʱ��,a.�������,a.ASA�ּ�,a.�ٴ�����,a.NNIS�ּ�,decode(a.��������,1,'һ������',2,'��������',3,'��������',4,'�ļ�����',' ') as ��������" & _
    ",a.��ǰ������ҩ,a.������ҩ����,a.��Ԥ�ڵĶ�������,a.������֢,a.������������,a.��������֢,a.�����Ѫ��Ѫ��,a.�����˿��ѿ�,a.�������Ѫ˨,a.���������л����,a.�������˥��,a.�����˨��,a.�����Ѫ֢,a.�����Źؽڹ���" & _
            " From ���������¼  A,��������Ŀ¼ B,������ĿĿ¼ C Where c.ID(+)=a.������ĿID And A.��������ID=B.ID(+) and ����ID=[1] And ��ҳID=[2] And ��¼��Դ=3 Order by ID"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    If rsTmp.EOF Then 'û��ʱ��ȡ������Դ�����
        '��������������ʱ��дȡ��
        StrSQL = "Select Max(��¼����) From ���������¼" & _
            " Where ����ID=" & mlng����ID & " And ��ҳID=" & mlng��ҳID & _
            " And ��¼��Դ=1 And ȡ��ʱ�� is NULL"
         StrSQL = "Select a.������ʿ,a.�������,a.�п�,a.����,a.��¼����,a.��¼��,a.ȡ��ʱ��,a.ȡ����,NVl(B.����,C.����) as ��������,a.ID,a.����ID,a.��ҳID,a.��¼��Դ,a.��������,a.������ʼʱ��,a.��������ʱ��,a.��������,a.��������ID,a.������ĿID,a.��������,a.����ҽʦ,a.��һ����," & _
         "a.�ڶ�����,a.������ʿ,a.����ʼʱ��,a.�������ʱ��,a.����ʽ,a.��������,a.��������,a.��Һ����,a.����ҽʦ,a.������ʼʱ��,a.��������ʱ��,a.�������,a.ASA�ּ�,a.�ٴ�����,a.NNIS�ּ�,decode(a.��������,1,'һ������',2,'��������',3,'��������',4,'�ļ�����',' ') as ��������" & _
         ",a.��ǰ������ҩ,a.������ҩ����,a.��Ԥ�ڵĶ�������,a.������֢,a.������������,a.��������֢,a.�����Ѫ��Ѫ��,a.�����˿��ѿ�,a.�������Ѫ˨,a.���������л����,a.�������˥��,a.�����˨��,a.�����Ѫ֢,a.�����Źؽڹ���" & _
            " From ���������¼  A,��������Ŀ¼ B,������ĿĿ¼ C Where c.ID(+)=a.������ĿID And " & _
            " A.��������ID=B.ID(+) and ����ID=[1] And ��ҳID=[2]" & _
            " And ��¼��Դ=1 And ȡ��ʱ�� is NULL And ��¼����=(" & StrSQL & ")" & _
            " Order by ID"
        Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
        If rsTmp.EOF Then '����
            StrSQL = "Select a.������ʿ,a.�������,a.�п�,a.����,a.��¼����,a.��¼��,a.ȡ��ʱ��,a.ȡ����,NVl(B.����,C.����) as ��������,a.ID,a.����ID,a.��ҳID,a.��¼��Դ,a.��������,a.������ʼʱ��,a.��������ʱ��,a.��������,a.��������ID,a.������ĿID,a.��������,a.����ҽʦ,a.��һ����," & _
            "a.�ڶ�����,a.������ʿ,a.����ʼʱ��,a.�������ʱ��,a.����ʽ,a.��������,a.��������,a.��Һ����,a.����ҽʦ,a.������ʼʱ��,a.��������ʱ��,a.�������,a.ASA�ּ�,a.�ٴ�����,a.NNIS�ּ�,decode(a.��������,1,'һ������',2,'��������',3,'��������',4,'�ļ�����',' ') as ��������" & _
            ",a.��ǰ������ҩ,a.������ҩ����,a.��Ԥ�ڵĶ�������,a.������֢,a.������������,a.��������֢,a.�����Ѫ��Ѫ��,a.�����˿��ѿ�,a.�������Ѫ˨,a.���������л����,a.�������˥��,a.�����˨��,a.�����Ѫ֢,a.�����Źؽڹ���" & _
                " From ���������¼  A,��������Ŀ¼ B,������ĿĿ¼ C Where c.ID(+)=a.������ĿID And  A.��������ID=B.ID(+) and ����ID=[1] And ��ҳID=[2] And ��¼��Դ=4 Order by ID"
            Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
        End If
    End If
    If Not rsTmp.EOF Then
        With vsOPS
            .Rows = .FixedRows + rsTmp.RecordCount + 1
            For i = 1 To rsTmp.RecordCount
                .TextMatrix(i, col��������) = Format(Nvl(rsTmp!��������), "yyyy-MM-dd")
                .TextMatrix(i, col��������) = Nvl(rsTmp!��������)
                .TextMatrix(i, col��������) = Nvl(rsTmp!��������)
                .TextMatrix(i, col����ҽʦ) = Nvl(rsTmp!����ҽʦ)
                .TextMatrix(i, col������ʿ) = Nvl(rsTmp!������ʿ)
                .TextMatrix(i, col����1) = Nvl(rsTmp!��һ����)
                .TextMatrix(i, col����2) = Nvl(rsTmp!�ڶ�����)
                .TextMatrix(i, col����ʽ) = GetItemField("������ĿĿ¼", Val(Nvl(rsTmp!����ʽ, 0)), "����")
                .TextMatrix(i, col����ҽʦ) = Nvl(rsTmp!����ҽʦ)
                If Not IsNull(rsTmp!�п�) And Not IsNull(rsTmp!����) Then
                    .TextMatrix(i, col�п�����) = rsTmp!�п� & "/" & rsTmp!����
                End If
                .TextMatrix(i, col��������ID) = Nvl(rsTmp!��������ID)
                .TextMatrix(i, col������ĿID) = Nvl(rsTmp!������Ŀid)
                .TextMatrix(i, col����ID) = Nvl(rsTmp!����ʽ)
                .TextMatrix(i, col��������) = Nvl(rsTmp!��������)
                .TextMatrix(i, COL�������.COL�������) = Nvl(rsTmp!�������)
                .TextMatrix(i, colASA�ּ�) = Decode(Nvl(rsTmp!asa�ּ�), "I��", "P1", "II��", "P2", "III��", "P3", "IV��", "P4", "V��", "P5", Nvl(rsTmp!asa�ּ�))
                .TextMatrix(i, colNNIS�ּ�) = Nvl(rsTmp!NNIS�ּ�)
                .TextMatrix(i, col��������) = Nvl(rsTmp!��������)
                .TextMatrix(i, col�ٴ�����) = IIf(Val(rsTmp!�ٴ����� & "") = 1, -1, 0)
                .TextMatrix(i, col����ҩ����) = rsTmp!������ҩ���� & ""
                .Cell(flexcpChecked, i, colԤ���ÿ���ҩ) = Val(rsTmp!��ǰ������ҩ & "")
                .Cell(flexcpChecked, i, col��Ԥ�ڵĶ�������) = Val(rsTmp!��Ԥ�ڵĶ������� & "")
                .Cell(flexcpChecked, i, col������֢) = Val(rsTmp!������֢ & "")
                .Cell(flexcpChecked, i, col������������) = Val(rsTmp!������������ & "")
                .Cell(flexcpChecked, i, col��������֢) = Val(rsTmp!��������֢ & "")
                .Cell(flexcpChecked, i, col�����Ѫ��Ѫ��) = Val(rsTmp!�����Ѫ��Ѫ�� & "")
                .Cell(flexcpChecked, i, col�����˿��ѿ�) = Val(rsTmp!�����˿��ѿ� & "")
                .Cell(flexcpChecked, i, col�������Ѫ˨) = Val(rsTmp!�������Ѫ˨ & "")
                .Cell(flexcpChecked, i, col���������л����) = Val(rsTmp!���������л���� & "")
                .Cell(flexcpChecked, i, col�������˥��) = Val(rsTmp!�������˥�� & "")
                .Cell(flexcpChecked, i, col�����˨��) = Val(rsTmp!�����˨�� & "")
                .Cell(flexcpChecked, i, col�����Ѫ֢) = Val(rsTmp!�����Ѫ֢ & "")
                .Cell(flexcpChecked, i, col�����Źؽڹ���) = Val(rsTmp!�����Źؽڹ��� & "")
                '��¼���ڱ༭�ָ�
                For j = 0 To .Cols - 1
                    .Cell(flexcpData, i, j) = .TextMatrix(i, j)
                Next
                
                rsTmp.MoveNext
            Next
        End With
    End If
    vsOPS.Tag = "δ�޸�"
    vsKSS.Tag = "δ�޸�"
    
    '��Ϸ������
    '---------------------------------------------------------------
    '������Ϸ������ȱʡֵ
    Call Set��Ϸ������(cbo�������Ժ)
    Call Set��Ϸ������(cbo��Ժ���Ժ)
    Call Set��Ϸ������(cbo��������Ժ)
    Call Set��Ϸ������(cbo�����벡��)
    Call Set��Ϸ������(cbo�ٴ��벡��)
    Call Set��Ϸ������(cbo�ٴ���ʬ��)
    Call Set��Ϸ������(cbo��ǰ������)
    Call Set��Ϸ������(cbo��ҽ�������Ժ)
    Call Set��Ϸ������(cbo��ҽ��Ժ���Ժ)
    
    StrSQL = "Select ��������,������� From ��Ϸ������ Where ����ID=[1] And ��ҳID=[2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    Do While Not rsTmp.EOF
        Select Case rsTmp!��������
        Case 1 '�������Ժ
            If Nvl(rsTmp!�������, 0) >= 0 Then cboinfo(cbo�������Ժ).ListIndex = rsTmp!�������
        Case 2 '��Ժ���Ժ
            If Nvl(rsTmp!�������, 0) >= 0 Then cboinfo(cbo��Ժ���Ժ).ListIndex = rsTmp!�������
        Case 3 '�����벡��
            If Nvl(rsTmp!�������, 0) >= 0 Then cboinfo(cbo�����벡��).ListIndex = rsTmp!�������
        Case 4 '�ٴ��벡��
            If Nvl(rsTmp!�������, 0) >= 0 Then cboinfo(cbo�ٴ��벡��).ListIndex = rsTmp!�������
        Case 5 '�ٴ���ʬ��
            If Nvl(rsTmp!�������, 0) >= 0 Then cboinfo(cbo�ٴ���ʬ��).ListIndex = rsTmp!�������
        Case 6 '��ǰ������
            If Nvl(rsTmp!�������, 0) >= 0 Then cboinfo(cbo��ǰ������).ListIndex = rsTmp!�������
        Case 7 '��������Ժ
            If Nvl(rsTmp!�������, 0) >= 0 Then cboinfo(cbo��������Ժ).ListIndex = rsTmp!�������
        Case 11 '��ҽ�������Ժ
            If Nvl(rsTmp!�������, 0) >= 0 Then cboinfo(cbo��ҽ�������Ժ).ListIndex = rsTmp!�������
        Case 12 '��ҽ��Ժ���Ժ
            If Nvl(rsTmp!�������, 0) >= 0 Then cboinfo(cbo��ҽ��Ժ���Ժ).ListIndex = rsTmp!�������
        Case 13 '��ҽ��֤
            If Nvl(rsTmp!�������, 0) >= 0 Then cboinfo(cbo��֤).ListIndex = rsTmp!�������
        Case 14 '��ҽ�η�
            If Nvl(rsTmp!�������, 0) >= 0 Then cboinfo(cbo�η�).ListIndex = rsTmp!�������
        Case 15 '��ҽ��ҩ
            If Nvl(rsTmp!�������, 0) >= 0 Then cboinfo(cbo��ҩ).ListIndex = rsTmp!�������
        End Select
        rsTmp.MoveNext
    Loop
    
    '������Ϣ
    '---------------------------------------------------------------
    '�����¼�
    lstAdvEvent.Clear
    StrSQL = "Select ����,���� From �����¼� order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        If Nvl(rsTmp!����) = "����������" Or Nvl(rsTmp!����) = "���������������" Then
            If Have��������(mlng����ID, "����") Then
                lstAdvEvent.AddItem Nvl(rsTmp!����)
                lstAdvEvent.ItemData(lstAdvEvent.NewIndex) = Val(rsTmp!����)
            End If
        Else
            lstAdvEvent.AddItem Nvl(rsTmp!����)
            lstAdvEvent.ItemData(lstAdvEvent.NewIndex) = Val(rsTmp!����)
        End If
        rsTmp.MoveNext
    Next
    If lstAdvEvent.ListCount > 0 Then lstAdvEvent.ListIndex = 0
    StrSQL = "Select ��Ϣֵ From ������ҳ�ӱ� Where ����id=[1] And ��ҳID=[2] And ��Ϣ��='�����¼�'"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    If rsTmp.RecordCount > 0 Then
        StrSQL = "Select /*+ Rule*/  * From  Table(f_Str2list([1]))"
        Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, rsTmp!��Ϣֵ & "")
        For i = 1 To rsTmp.RecordCount
            For j = 0 To lstAdvEvent.ListCount - 1
                If lstAdvEvent.ItemData(j) = Val(rsTmp!COLUMN_VALUE & "") Then
                    lstAdvEvent.Selected(j) = True
                    If lstAdvEvent.List(j) = "ѹ��" Or lstAdvEvent.List(j) = "ҽԺ�ڵ���/׹��" Then Call lstAdvEvent_ItemCheck(CInt(j))
                End If
            Next
        rsTmp.MoveNext
        Next
    End If
    If lstAdvEvent.ListCount > 0 Then lstAdvEvent.ListIndex = 0
    '��Ⱦ����
    lstInfection.Clear
    StrSQL = "Select ����,���� From ��Ⱦ���� order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption)
    For i = 1 To rsTmp.RecordCount
        lstInfection.AddItem Nvl(rsTmp!����)
        lstInfection.ItemData(lstInfection.NewIndex) = Val(rsTmp!����)
        rsTmp.MoveNext
    Next
    If lstInfection.ListCount > 0 Then lstInfection.ListIndex = 0
    StrSQL = "Select ��Ϣֵ From ������ҳ�ӱ� Where ����id=[1] And ��ҳID=[2] And ��Ϣ��='��Ⱦ����'"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    If rsTmp.RecordCount > 0 Then
        StrSQL = "Select /*+ Rule*/  * From  Table(f_Str2list([1]))"
        Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, rsTmp!��Ϣֵ & "")
        For i = 1 To rsTmp.RecordCount
            For j = 0 To lstInfection.ListCount - 1
                If lstInfection.ItemData(j) = Val(rsTmp!COLUMN_VALUE & "") Then
                    lstInfection.Selected(j) = True
                End If
            Next
        rsTmp.MoveNext
        Next
    End If
    If lstInfection.ListCount > 0 Then lstInfection.ListIndex = 0
    '--------------------------------------------------------------
    '����ҩ��
    StrSQL = "Select a.ҩ��id, a.��ҩĿ��, a.ʹ�ý׶�, a.ʹ������,a.ҩƷ���� ����,һ���п�Ԥ����,DDD��,������ҩ " & vbNewLine & _
            " From ���˿����ؼ�¼ A" & vbNewLine & _
            " Where a.����id = [1] And a.��ҳid = [2] Order By DDD�� Desc"
    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    
    Do While Not rsTmp.EOF
        With vsKSS
            For j = .FixedRows To .Rows - 1
                If .TextMatrix(j, 1) = "" Then
                    .RowData(j) = Val(rsTmp!ҩ��id & "")
                    If .RowData(j) <> 0 Then
                        .TextMatrix(j, 1) = Nvl(rsTmp!����)
                        .Cell(flexcpData, j, 1) = .TextMatrix(j, 1)
                        .TextMatrix(j, kss��ҩĿ��) = Nvl(rsTmp!��ҩĿ��)
                        .TextMatrix(j, kssʹ�ý׶�) = Nvl(rsTmp!ʹ�ý׶�)
                        .TextMatrix(j, kssʹ������) = IIf(Val(rsTmp!ʹ������ & "") = 0, "", Val(rsTmp!ʹ������ & "") & "")
                        .Cell(flexcpChecked, j, KSSһ���п�Ԥ����) = Val(rsTmp!һ���п�Ԥ���� & "")
                        .TextMatrix(j, KSSDDD��) = IIf(Val(rsTmp!DDD�� & "") > 0 And Val(rsTmp!DDD�� & "") < 1, "0", "") & Val(rsTmp!DDD�� & "")
                        .TextMatrix(j, KSS������ҩ) = rsTmp!������ҩ & ""
                    End If
                    Exit For
                ElseIf .RowData(j) = Val(rsTmp!ҩ��id & "") Then
                '�ų��ظ�ֵ��������ظ��ģ��򽫺�����е���Ϣ���ϡ�
                    If .RowData(j) <> 0 Then
                        .TextMatrix(j, 1) = Nvl(rsTmp!����)
                        .Cell(flexcpData, j, 1) = .TextMatrix(j, 1)
                        .TextMatrix(j, kss��ҩĿ��) = Nvl(rsTmp!��ҩĿ��)
                        .TextMatrix(j, kssʹ�ý׶�) = Nvl(rsTmp!ʹ�ý׶�)
                        .TextMatrix(j, kssʹ������) = IIf(Val(rsTmp!ʹ������ & "") = 0, "", Val(rsTmp!ʹ������ & "") & "")
                        .Cell(flexcpChecked, j, KSSһ���п�Ԥ����) = Val(rsTmp!һ���п�Ԥ���� & "")
                        .TextMatrix(j, KSSDDD��) = IIf(Val(rsTmp!DDD�� & "") > 0 And Val(rsTmp!DDD�� & "") < 1, "0", "") & Val(rsTmp!DDD�� & "")
                        .TextMatrix(j, KSS������ҩ) = rsTmp!������ҩ & ""
                    End If
                    Exit For
                End If
            Next
            '���û������û�п����ˣ�������һ��
            If j > .Rows - 1 Then
                .AddItem ""
                .RowData(.Rows - 1) = Val(rsTmp!ҩ��id & "")
                If .RowData(.Rows - 1) <> 0 Then
                    .TextMatrix(.Rows - 1, 1) = rsTmp!����
                    .Cell(flexcpData, .Rows - 1, 1) = .TextMatrix(.Rows - 1, 1)
                    .TextMatrix(.Rows - 1, kss��ҩĿ��) = Nvl(rsTmp!��ҩĿ��)
                    .TextMatrix(.Rows - 1, kssʹ�ý׶�) = Nvl(rsTmp!ʹ�ý׶�)
                    .TextMatrix(.Rows - 1, kssʹ������) = IIf(Val(rsTmp!ʹ������ & "") = 0, "", Val(rsTmp!ʹ������ & "") & "")
                    .Cell(flexcpChecked, .Rows - 1, KSSһ���п�Ԥ����) = Val(rsTmp!һ���п�Ԥ���� & "")
                    .TextMatrix(.Rows - 1, KSSDDD��) = IIf(Val(rsTmp!DDD�� & "") > 0 And Val(rsTmp!DDD�� & "") < 1, "0", "") & Val(rsTmp!DDD�� & "")
                    .TextMatrix(.Rows - 1, KSS������ҩ) = rsTmp!������ҩ & ""
                End If
            End If
        End With
        rsTmp.MoveNext
    Loop
    Call SetKSSSerial
    
    If mbln�������� Then
        '���ƻ���
        Call Load���������(mlng����ID, mlng��ҳID)
        
    End If
    Call Load��ҳ����(mlng����ID, mlng��ҳID)
        
    Screen.MousePointer = 0
    LoadPageData = True
    Exit Function
errH:
    Screen.MousePointer = 0
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SetKSSSerial()
    Dim i As Long
    
    With vsKSS
        For i = .FixedRows To .Rows - 1
            .TextMatrix(i, 0) = i
        Next
    End With
End Sub

Private Sub txtInfo_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txtInfo(Index).Locked Then
        glngTXTProc = GetWindowLong(txtInfo(Index).hwnd, GWL_WNDPROC)
        Call SetWindowLong(txtInfo(Index).hwnd, GWL_WNDPROC, AddressOf WndMessage)
    End If
End Sub

Private Sub txtInfo_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And txtInfo(Index).Locked Then
        Call SetWindowLong(txtInfo(Index).hwnd, GWL_WNDPROC, glngTXTProc)
    End If
End Sub

Private Sub txtInfo_Validate(Index As Integer, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String, blnCancel As Boolean
    Dim str�Ա� As String, int������� As Integer
    Dim strInput As String, vPoint As POINTAPI

    Select Case Index
        Case txt����
            'û�������г�������ʱ����һ������
            If txtInfo(txt����).Text = "" And IsDate(txt��������.Text) Then
                txt��������.Tag = "": Call txt��������_Validate(False)
            End If
        Case txtȷ������
            txtInfo(Index).Text = GetFullDate(txtInfo(Index).Text)
            If Not IsDate(txtInfo(Index).Text) Then
                txtInfo(Index).Text = ""
            ElseIf Not CheckDateRange(txtInfo(Index).Text, True) Then
                txtInfo(Index).Text = ""
            End If
        Case txt���ȴ���, txt�ɹ�����, txt��������, txt���ϸ��, txt��ѪС��, txt��Ѫ��, txt��ȫѪ, txt�������
            If txtInfo(Index).Text <> "" Then
                If Not IsNumeric(txtInfo(Index).Text) Then
                    txtInfo(Index).Text = ""
                ElseIf Val(txtInfo(Index).Text) <= 0 Then
                    txtInfo(Index).Text = ""
                End If
            End If
            
            If Index = txt���ȴ��� Or Index = txt�ɹ����� Or Index = txt�������� Then
                If IsNumeric(txtInfo(Index).Text) Then
                    txtInfo(Index).Text = Int(Val(txtInfo(Index).Text))
                End If
            End If
        Case txt��ԭѧ
            If txtInfo(txt��ԭѧ).Text = "" Then
                txtInfo(txt��ԭѧ).Tag = ""
                cmdInfo(txt��ԭѧ).Tag = ""
            ElseIf txtInfo(txt��ԭѧ).Text = txtInfo(txt��ԭѧ).Tag Then
                'Nothing
            Else
                int������� = Val(Mid(gstr�������, 2, 1))
                If int������� = 0 Then int������� = 1
                
                strInput = UCase(txtInfo(txt��ԭѧ).Text)
                
                If cboinfo(cbo�Ա�).Text Like "*��*" Then
                    str�Ա� = "��"
                ElseIf cboinfo(cbo�Ա�).Text Like "*Ů*" Then
                    str�Ա� = "Ů"
                End If
                If zlCommFun.IsCharChinese(strInput) Then
                    StrSQL = "���� Like [2]" '���뺺��ʱֻƥ������
                Else
                    StrSQL = "���� Like [1] Or ���� Like [2] Or " & IIf(mint���� = 0, "����", "�����") & " Like [2]"
                End If
                StrSQL = _
                    " Select ID,ID as ��ĿID,����,����,����," & IIf(mint���� = 0, "����", "����� as ����") & ",˵��" & _
                    " From ��������Ŀ¼ Where Instr([3],���)>0 And (" & StrSQL & ")" & _
                    IIf(str�Ա� <> "", " And (�Ա�����=[4] Or �Ա����� is NULL)", "") & _
                    " And (����ʱ�� is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " Order by ����"

                If int������� = 1 And zlCommFun.IsCharChinese(strInput) Then
                    '�����ж��룺Y-�����ж����ⲿԭ�򣻲����������M-������̬ѧ���룻������ϣ�D-ICD-10��������
                    On Error GoTo errH
                    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, strInput & "%", mstrLike & strInput & "%", "'D'", str�Ա�)
                    If rsTmp.EOF Then
                        Set rsTmp = Nothing
                    ElseIf rsTmp.RecordCount > 1 Then
                        Set rsTmp = Nothing '����¼��ʱ�ж��ƥ�䲻����ѡ��
                    End If
                Else
                    vPoint = GetCoordPos(txtInfo(txt��ԭѧ).hwnd, 0, 0)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 0, "��ԭѧ���", False, "", "", False, False, True, vPoint.X, vPoint.Y, txtInfo(txt��ԭѧ).Height, blnCancel, False, True, _
                        strInput & "%", mstrLike & strInput & "%", "'D'", str�Ա�)
                    If blnCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                        Cancel = True
                    Else
                        '���������뷽ʽ
                        If rsTmp Is Nothing And (int������� = 2 Or int������� = 3 And mint���� <> 0) Then
                            MsgBox "û���ҵ�������ƥ������ݡ�", vbInformation, gstrSysName
                            Cancel = True
                        End If
                    End If
                End If
                
                If Not Cancel Then
                    If rsTmp Is Nothing Then
                        cmdInfo(txt��ԭѧ).Tag = ""
                    Else
                        txtInfo(txt��ԭѧ).Text = IIf(Not IsNull(rsTmp!����), "(" & rsTmp!���� & ")", "") & Nvl(rsTmp!����)
                        txtInfo(txt��ԭѧ).Tag = txtInfo(txt��ԭѧ).Text
                        cmdInfo(txt��ԭѧ).Tag = rsTmp!��ĿID
                    End If
                End If
            End If
        Case txt��֢�໤��
            strInput = UCase(txtInfo(Index).Text)
            If strInput = "" Then Exit Sub
            StrSQL = " Select Distinct A.ID,A.����,A.����" & _
                    " From ���ű� A,��������˵�� B" & _
                    " Where B.����ID=A.ID And B.��������='ICU'" & _
                    " And (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                    " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                    " And (A.���� Like [1] Or A.���� Like [2] Or A.���� Like [2])" & _
                    " Order by A.����"
            vPoint = GetCoordPos(txtInfo(Index).hwnd, 0, 0)
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 0, "��֢�໤��", _
                False, "", "", False, False, True, vPoint.X, vPoint.Y, txtInfo(Index).Height, blnCancel, False, True, _
                strInput & "%", mstrLike & strInput & "%")
            
            If rsTmp Is Nothing Then
                If Not blnCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                    MsgBox "û���ҵ�ָ����ICU��֢�໤�ҡ�", vbInformation, Me.Caption
                End If
                Cancel = True
                Exit Sub
            Else
                txtInfo(Index).Text = rsTmp!���� & ""
            End If

    End Select
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txt��������_Change()
    If Visible Then mblnChange = True
End Sub

Private Sub txt��������_GotFocus()
    Call zlControl.TxtSelAll(txt��������)
End Sub

Private Sub txt��������_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txt��������_Validate(Cancel As Boolean)
    Dim str���� As String
    
    If IsDate(txt��������.Text) Then
        If txt��������.Tag = txt��������.Text Then Exit Sub
        txt��������.Tag = txt��������.Text '���ڼ�¼����仯
        'Сʱ���������λ�򲻽��з���
        If cboinfo(cbo���䵥λ).ListIndex < 3 Then
            str���� = PatiAgeCalc(txt��������.Text, , txtInfo(txt��Ժʱ��).Text)
            Call LoadOldData(str����)
        End If
    ElseIf txt��������.Text = "____-__-__" Then
        txt����ʱ��.Text = "__:__"
    Else
        txt��������.Text = "____-__-__"
        txt����ʱ��.Text = "__:__"
        Cancel = True
    End If
End Sub

Private Sub txt����ʱ��_Change()
    If Visible Then mblnChange = True
End Sub

Private Sub txt����ʱ��_GotFocus()
    Call zlControl.TxtSelAll(txt����ʱ��)
End Sub

Private Sub txt����ʱ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf Not IsDate(txt��������.Text) Then
        KeyAscii = 0: txt����ʱ��.Text = "__:__"
    End If
End Sub

Private Sub txt����ʱ��_Validate(Cancel As Boolean)
    If txt����ʱ��.Text <> "__:__" And Not IsDate(txt����ʱ��.Text) Then
        txt����ʱ��.Text = "__:__": Cancel = True
    End If
End Sub

Private Sub txt��������_Change()
    If Visible Then mblnChange = True
    
    If IsDate(txt��������.Text) Then
        txt����ʱ��.Enabled = True
    Else
        txt����ʱ��.Enabled = False
    End If
End Sub

Private Sub txt��������_GotFocus()
    Call zlControl.TxtSelAll(txt��������)
End Sub

Private Sub txt��������_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txt��������_Validate(Cancel As Boolean)
    If txt��������.Text <> "____-__-__" And Not IsDate(txt��������.Text) Then
        txt��������.Text = "____-__-__": Cancel = True
    End If
End Sub

Private Sub txt����ʱ��_GotFocus()
    Call zlControl.TxtSelAll(txt����ʱ��)
End Sub

Private Sub txt����ʱ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        KeyAscii = 0
        Call zlCommFun.PressKey(vbKeyTab)
    End If
End Sub

Private Sub txt����ʱ��_Validate(Cancel As Boolean)
    If txt����ʱ��.Text <> "__:__" And Not IsDate(txt����ʱ��.Text) Then
        txt����ʱ��.Text = "__:__": Cancel = True
    End If
End Sub

Private Sub txt����ʱ��_Change()
    If Visible Then mblnChange = True
End Sub

Private Sub txt����ʱ��_GotFocus()
    Call zlControl.TxtSelAll(txt����ʱ��)
End Sub

Private Sub txt����ʱ��_Validate(Cancel As Boolean)
    If Not IsDate(txt����ʱ��.Text) And txt����ʱ��.Text <> "____-__-__ __:__:__" Then
        Cancel = True
    End If
End Sub

Private Sub timThis_Timer()
    Dim lngSelNum As Long
    
    If vsAller.Col = AC_����ʱ�� Then
        lngSelNum = vsAller.EditSelStart
        If lngSelNum <> mlngSelNum And lngSelNum <> 16 And lngSelNum <> 0 Then
            Call Vs_EditSelChange(lngSelNum)
            mlngSelNum = lngSelNum
        End If
    End If
End Sub

Private Sub Vs_EditSelChange(ByVal lngSelNum As Long)
'���û��л�����ʱ�򴥷�
    With vsAller
        If lngSelNum <= 4 Then
            .EditSelStart = 0
            .EditSelLength = 4
            mlngNum = 0
            mlngNumBack = 4
        ElseIf lngSelNum <= 7 Then
            .EditSelStart = 5
            .EditSelLength = 2
            mlngNum = 5
            mlngNumBack = 7
        ElseIf lngSelNum <= 10 Then
            .EditSelStart = 8
            .EditSelLength = 2
            mlngNum = 8
            mlngNumBack = 10
        ElseIf lngSelNum <= 13 Then
            .EditSelStart = 11
            .EditSelLength = 2
            mlngNum = 11
            mlngNumBack = 13
        ElseIf lngSelNum < 16 Then
            .EditSelStart = 14
            .EditSelLength = 2
            mlngNum = 14
            mlngNumBack = 16
        End If
    End With
End Sub

Private Sub vsAller_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call vsAller_AfterRowColChange(-1, -1, Row, Col)
End Sub

Private Sub vsAller_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsAller
        If NewCol = AC_����ҩ�� Then
            .ComboList = "..."
            .FocusRect = flexFocusSolid
        Else
            .FocusRect = IIf(Trim(vsAller.TextMatrix(NewRow, AC_����ҩ��)) = "", flexFocusLight, flexFocusSolid)
            .ComboList = ""
        End If
    End With
End Sub

Private Sub vsAller_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = AC_����ʱ�� And Trim(vsAller.Cell(flexcpData, Row, AC_����ҩ��)) = "" Then Cancel = True
End Sub

Private Sub vsAller_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String, blnCancel As Boolean
    Dim int�Ա� As Integer
    
    With vsAller
        If cboinfo(cbo�Ա�).Text Like "*��*" Then
            int�Ա� = 1
        ElseIf cboinfo(cbo�Ա�).Text Like "*Ů*" Then
            int�Ա� = 2
        End If
        
        StrSQL = _
            " Select -1 as ID,-NULL as �ϼ�ID,0 as ĩ��,NULL as ����,'����ҩ' as ����,NULL as ��λ,NULL as ����,NULL as �������,NULL as Ƥ�� From Dual Union ALL" & _
            " Select -2 as ID,-NULL as �ϼ�ID,0 as ĩ��,NULL as ����,'�г�ҩ' as ����,NULL as ��λ,NULL as ����,NULL as �������,NULL as Ƥ�� From Dual Union ALL" & _
            " Select -3 as ID,-NULL as �ϼ�ID,0 as ĩ��,NULL as ����,'�в�ҩ' as ����,NULL as ��λ,NULL as ����,NULL as �������,NULL as Ƥ�� From Dual Union ALL" & _
            " Select ID,Nvl(�ϼ�ID,-����) as �ϼ�ID,0 as ĩ��,NULL as ����,����," & _
            " NULL as ��λ,NULL as ����,NULL as �������,NULL as Ƥ��" & _
            " From ���Ʒ���Ŀ¼ Where ���� IN (1,2,3) And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID" & _
            " Union All" & _
            " Select Distinct A.ID,A.����ID as �ϼ�ID,1 as ĩ��,A.����,A.����," & _
            " A.���㵥λ as ��λ,B.ҩƷ���� as ����,B.�������,Decode(B.�Ƿ�Ƥ��,1,'��','') as Ƥ��" & _
            " From ������ĿĿ¼ A,ҩƷ���� B" & _
            " Where A.��� IN('5','6','7') And A.ID=B.ҩ��ID" & _
            IIf(int�Ա� <> 0, " And Nvl(A.�����Ա�,0) IN(0,[1])", "") & _
            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)"
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 2, "����ҩ��", False, "", "", False, False, False, 0, 0, 0, blnCancel, False, False, int�Ա�)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "û��ҩƷ���ݿ���ѡ��", vbInformation, gstrSysName
            End If
        Else
            Call SetAllerInput(Row, rsTmp)
            Call AllerEnterNextCell
        End If
    End With
End Sub

Private Sub vsAller_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    If mbln��ʿվ Or mblnReadOnly Then Exit Sub
    
    With vsAller
        If KeyCode = vbKeyF4 Then
            If .Col = 1 Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, AC_����ҩ��) <> "" Then
                If MsgBox("ȷʵҪ������й���ҩ����", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    .RemoveItem .Row
                    vsAller.Tag = ""
                    mblnChange = True
                End If
            End If
        ElseIf KeyCode > 127 Then
            '���ֱ�����뺺�ֵ�����
            Call vsAller_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsAller_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    If KeyCode = vbKeyLeft Then
        If mlngNum <= 4 Then Exit Sub
        If mlngNum <= 7 Then Vs_EditSelChange (4): Exit Sub
        If mlngNum <= 10 Then Vs_EditSelChange (7): Exit Sub
        If mlngNum <= 13 Then Vs_EditSelChange (10): Exit Sub
        If mlngNum <= 16 Then Vs_EditSelChange (13): Exit Sub
    End If
End Sub

Private Sub vsAller_KeyPress(KeyAscii As Integer)
    If mbln��ʿվ Or mblnReadOnly Then Exit Sub

    With vsAller
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call AllerEnterNextCell
        ElseIf .Col = AC_����ҩ�� Then
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                Call vsAller_CellButtonClick(.Row, .Col)
            Else
                .ComboList = "" 'ʹ��ť״̬��������״̬
            End If
        End If
    End With
End Sub

Private Sub vsAller_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim blnIsNextchr As Boolean
    Dim strChr As String
    
    If KeyAscii = 13 Then
        mblnReturn = True
    Else
        mblnReturn = False
    End If
    With vsAller
        If Col = AC_����ʱ�� Then
            If KeyAscii = 13 Then .Col = .Col + 1: .ShowCell Row, Col: Exit Sub
            If KeyAscii = vbKeyBack Then
                If mlngNumBack <= 16 Then
                    If mlngNumBack = 0 Then KeyAscii = 0: Exit Sub
                    blnIsNextchr = InStr("1234567890", Mid(.TextMatrix(.Row, .Col), mlngNumBack, 1)) = 0
                    strChr = Mid(.TextMatrix(.Row, .Col), mlngNumBack - IIf(blnIsNextchr, 1, 0), 1)
                    mlngNumBack = mlngNumBack - IIf(blnIsNextchr, 2, 1)
                    .EditText = Mid(.EditText, 1, mlngNumBack) & strChr & Mid(.EditText, mlngNumBack + 2)
                    mlngNum = mlngNumBack
                    KeyAscii = 0
                    If mlngNum <= 4 Then
                        .EditSelStart = 0
                        .EditSelLength = 4
                    ElseIf mlngNum <= 8 Then
                        .EditSelStart = 5
                        .EditSelLength = 2
                    ElseIf mlngNum <= 11 Then
                        .EditSelStart = 8
                        .EditSelLength = 2
                    ElseIf mlngNum <= 14 Then
                        .EditSelStart = 11
                        .EditSelLength = 2
                    ElseIf mlngNum <= 16 Then
                        .EditSelStart = 14
                        .EditSelLength = 2
                    End If
                End If
            Else
                If Not IsNumeric(Chr(KeyAscii)) Then KeyAscii = 0
                If Len(.EditText) <= 16 And mlngNum <> 16 Then
                    blnIsNextchr = InStr("1234567890", Mid(.TextMatrix(.Row, .Col), mlngNum + 2, 1)) = 0
                    strChr = Chr(KeyAscii)
                    .EditText = Mid(.EditText, 1, mlngNum) & strChr & Mid(.EditText, mlngNum + 2)
                    mlngNum = mlngNum + IIf(blnIsNextchr, 2, 1)
                    mlngNumBack = mlngNum
                End If
                KeyAscii = 0
                If mlngNum <= 4 Then
                    .EditSelStart = 0
                    .EditSelLength = 4
                ElseIf mlngNum <= 7 Then
                    .EditSelStart = 5
                    .EditSelLength = 2
                ElseIf mlngNum <= 10 Then
                    .EditSelStart = 8
                    .EditSelLength = 2
                ElseIf mlngNum <= 13 Then
                    .EditSelStart = 11
                    .EditSelLength = 2
                ElseIf mlngNum <= 16 Then
                    .EditSelStart = 14
                    .EditSelLength = 2
                End If
            End If
        End If
    End With
End Sub

Private Sub vsAller_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    If Col = AC_����ҩ�� Then
        vsAller.EditSelStart = 0
        vsAller.EditSelLength = zlCommFun.ActualLen(vsAller.EditText)
    ElseIf Col = AC_����ʱ�� Then
        vsAller.EditSelStart = 0
        vsAller.EditSelLength = 4
        mlngNum = 0
        timThis.Enabled = True
    End If
End Sub

Private Sub vsAller_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = AC_������Ӧ And Trim(vsAller.TextMatrix(Row, AC_����ҩ��)) = "" Then Cancel = True
End Sub

Private Sub vsAller_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String, blnCancel As Boolean
    Dim strInput As String, vPoint As POINTAPI
    Dim int�Ա�  As Integer
    Dim curDate As Date
    
    With vsAller
        If Col = AC_����ҩ�� Then
            If .EditText = "" Then
                .EditText = .Cell(flexcpData, Row, Col)
                If mblnReturn Then Call AllerEnterNextCell
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call AllerEnterNextCell
            Else
                If LenB(StrConv(.EditText, vbFromUnicode)) > 60 Then
                    MsgBox "ҩ�����Ʋ��ܳ���30�����ֵĳ��ȡ�", vbInformation, Me.Caption
                    Cancel = True
                    Exit Sub
                End If
                strInput = UCase(.EditText)
                If cboinfo(cbo�Ա�).Text Like "*��*" Then
                    int�Ա� = 1
                ElseIf cboinfo(cbo�Ա�).Text Like "*Ů*" Then
                    int�Ա� = 2
                End If
                StrSQL = _
                    " Select Distinct A.ID,A.����,A.����,A.���㵥λ as ��λ," & _
                    " B.ҩƷ���� as ����,B.�������,Decode(B.�Ƿ�Ƥ��,1,'��','') as Ƥ��" & _
                    " From ������ĿĿ¼ A,ҩƷ���� B,������Ŀ���� C" & _
                    " Where A.��� IN('5','6','7') And A.ID=B.ҩ��ID And A.ID=C.������ĿID" & _
                    " And (A.���� Like [1] Or A.���� Like [2] Or C.���� Like [2] Or C.���� Like [2])" & _
                    IIf(int�Ա� <> 0, " And Nvl(A.�����Ա�,0) IN(0,[3])", "") & _
                    Decode(mint����, 0, " And C.����=[4]", 1, " And C.����=[4]", "") & _
                    " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                    " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                    " Order by A.����"
                
                vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 0, "����ҩ��", False, "", "", False, _
                    False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    strInput & "%", mstrLike & strInput & "%", int�Ա�, mint���� + 1)
                If blnCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                    Cancel = True
                Else
                    Call SetAllerInput(Row, rsTmp): .EditText = .Text
                    If mblnReturn Then Call AllerEnterNextCell
                End If
            End If
            mblnReturn = False
        ElseIf Col = AC_����ʱ�� Then
            If Not IsDate(.EditText) And .EditText <> "" Then
                MsgBox "����������ڸ�ʽ����ȷ����ʽ�磺2010-10-10 18:30��"
                Cancel = True
                .EditText = vsAller.TextMatrix(Row, Col)
            Else
                If .EditText <> "" Then
                    curDate = zlDatabase.Currentdate
                    If CDate(.EditText) > curDate Then
                        MsgBox "����������ڲ��ܴ��ڵ�ǰʱ�䡣��ǰʱ�䣺" & curDate & "��"
                        Cancel = True
                        .EditText = .TextMatrix(Row, Col)
                    End If
                End If
                timThis.Enabled = False
                If .Cell(flexcpData, Row, Col) <> .EditText Then
                    .Cell(flexcpData, Row, Col) = .EditText
                    mblnChange = True
                End If
                .Tag = ""
            End If
        Else
            If LenB(StrConv(.EditText, vbFromUnicode)) > 100 Then
                MsgBox "ҩ�����Ʋ��ܳ���50�����ֵĳ��ȡ�", vbInformation, Me.Caption
                Cancel = True
                Exit Sub
            End If
        End If
    End With
End Sub

Private Sub vsDiagXY_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsDiagXY
        If Col = col��Ժ��� Then
            '��Ҫ����ǻس��뿪:����ComboIndex,ȡ���༭ʱ����
            .TextMatrix(Row, Col) = NeedName(.TextMatrix(Row, Col))
            If Not XYCellEditable(Row, col�Ƿ�δ��) Then
                .TextMatrix(Row, col�Ƿ�δ��) = ""
            End If
            Call SetEditableFrom��Ժ���
            mblnChange = True
            .Tag = ""
        End If
        
        If Col = col������� Then
            ' .EditText = "" �ų���Ԫ�������ݲ����س���״��
            If .EditText = "" And .Cell(flexcpData, Row, Col) <> "" Then
                '�ڵ���vsDiagXY_KeyDown(vbKeyDelete, 0)���ǿ���ɾ����ǰ�У������ָ�ԭʼ����
                .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                Call vsDiagXY_KeyDown(vbKeyDelete, 0)
            End If
        End If
        
        Call vsDiagXY_AfterRowColChange(-1, -1, .Row, .Col)
        '�ж��Ƿ������޸�
        If vsDiagXY.Tag = "δ�޸�" Then
            vsDiagXY.Tag = ""
        End If
    End With
End Sub

Private Sub SetEditableFrom��Ժ���()
    With vsDiagXY
'        '��Ҫ��ϵĳ�Ժ���Ϊ�����򲻿��������ȴ���
'        If .TextMatrix(GetRow(3), col��Ժ���) = "����" Then
'            txtInfo(txt���ȴ���).Text = ""
'            txtInfo(txt���ȴ���).Locked = True
'            txtInfo(txt���ȴ���).TabStop = False
'            txtInfo(txt���ȴ���).BackColor = vbButtonFace
'        Else
'            txtInfo(txt���ȴ���).Locked = False
'            txtInfo(txt���ȴ���).TabStop = True
'            txtInfo(txt���ȴ���).BackColor = vbWindowBackground
'        End If
'        Call txtInfo_Change(txt���ȴ���)
        
        '��Ҫ��ϵĳ�Ժ���Ϊ����ʱ�ſ���ʬ��
        If .TextMatrix(GetRow(3), col��Ժ���) = "����" Then
            txt����ʱ��.Enabled = True: txt����ʱ��.TabStop = True: txt����ʱ��.BackColor = vbWindowBackground
            txtInfo(txt����ԭ��).Enabled = True
            txtInfo(txt����ԭ��).TabStop = True
            txtInfo(txt����ԭ��).BackColor = vbWindowBackground
            chkInfo(chkʬ��).Enabled = True
            chkInfo(chkʬ��).TabStop = True
        Else
            txt����ʱ��.Text = "____-__-__ __:__:__"
            txt����ʱ��.Enabled = False: txt����ʱ��.TabStop = False: txt����ʱ��.BackColor = vbButtonFace
            txtInfo(txt����ԭ��).Text = ""
            txtInfo(txt����ԭ��).Enabled = False
            txtInfo(txt����ԭ��).TabStop = False
            txtInfo(txt����ԭ��).BackColor = vbButtonFace
            chkInfo(chkʬ��).Value = 0
            chkInfo(chkʬ��).Enabled = False
            chkInfo(chkʬ��).TabStop = False
        End If
        
        '��Ҫ��ϵĳ�Ժ�����Ϊ����ʱ�ſ�������
        If .TextMatrix(GetRow(3), col��Ժ���) <> "����" Then
            chkInfo(chk����).Enabled = True
            chkInfo(chk����).TabStop = True
            cboinfo(cbo��Ժ��ʽ).Enabled = True
        Else
            '��������������Ժ�������Ϊ����
            cboinfo(cbo��Ժ��ʽ).Text = "����"
            cboinfo(cbo��Ժ��ʽ).Enabled = False
            
            chkInfo(chk����).Value = 0
            chkInfo(chk����).Enabled = False
            chkInfo(chk����).TabStop = False
        End If
        Call chkInfo_Click(chk����)
    End With
End Sub

Private Sub vsDiagXY_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim i As Long
    
    With vsDiagXY
        '���ͼƬ
        For i = .FixedRows To .Rows - 1
            If Not .Cell(flexcpPicture, i, col����) Is Nothing Then
                Set .Cell(flexcpPicture, i, col����) = Nothing
            End If
            If Not .Cell(flexcpPicture, i, colDel) Is Nothing Then
               Set .Cell(flexcpPicture, i, colDel) = Nothing
            End If
        Next
        
        If Not XYCellEditable(NewRow, NewCol) Then
            .ComboList = ""
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
            Set .CellButtonPicture = Nothing
            
            If NewCol = col������� Then
                .ComboList = "..."
            ElseIf NewCol = col��Ժ��� Then
                .ComboList = .ColData(NewCol)
            ElseIf NewCol = col��Ժ���� Then
                If .TextMatrix(NewRow, 0) = "��Ժ���" Or .TextMatrix(NewRow, 0) = "�������" Or .TextMatrix(NewRow, 0) = "" Then
                    .ComboList = "��|�ٴ�δȷ��|�������|��"
                Else
                    .ComboList = ""
                    .FocusRect = flexFocusLight
                End If
            ElseIf NewCol = col���� Then
                .ComboList = "..."
                .FocusRect = flexFocusNone
                Set .CellButtonPicture = imgButtonNew.Picture
            ElseIf NewCol = colDel Then
                .ComboList = "..."
                .FocusRect = flexFocusNone
                Set .CellButtonPicture = imgButtonDel.Picture
            Else
                .ComboList = ""
            End If
        End If
        If NewRow >= .FixedRows Then
            '��ʾͼƬ
            If NewCol <> col���� And .TextMatrix(NewRow, col�������) <> "" And .TextMatrix(NewRow, 0) <> "��Ժ���" Then
                Set .Cell(flexcpPicture, NewRow, col����) = imgButtonNew.Picture
            End If
            '��ʾͼƬ
            If NewCol <> colDel Then
                Set .Cell(flexcpPicture, NewRow, colDel) = imgButtonDel.Picture
            End If
        End If
    End With
End Sub

Private Sub vsDiagXY_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = col���� Then Cancel = True
End Sub

Private Sub vsDiagXY_Click()
    With vsDiagXY
        If (.MouseCol = col���� Or .MouseCol = colDel) And .MouseRow >= .FixedRows Then
            
            If .MouseCol = col���� Then
                If .TextMatrix(.MouseRow, col�������) = "" Or .TextMatrix(.MouseRow, 0) = "��Ժ���" Then Exit Sub
            End If
            
            .Select .MouseRow, .MouseCol
            Call vsDiagXY_CellButtonClick(.MouseRow, .MouseCol)
        End If
    End With
End Sub

Private Sub vsDiagXY_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    With vsDiagXY
        If Col = col��Ժ��� Then
            '��λ��ƥ����
            For i = 0 To .ComboCount - 1
                If NeedName(.ComboItem(i)) = .TextMatrix(Row, Col) Then
                    .ComboIndex = i: Exit For
                End If
            Next
        End If
    End With
End Sub

Private Sub vsDiagXY_DblClick()
    Call vsDiagXY_KeyPress(32)
    '����Ϊ���޸�
    If vsDiagXY.Col = col�Ƿ�δ�� Or vsDiagXY.Col = col�Ƿ����� Then
        If vsDiagXY.Tag = "δ�޸�" Then vsDiagXY.Tag = "": mblnChange = True
    End If
End Sub

Private Sub vsDiagXY_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long, j As Long
    
    If mbln��ʿվ Or mblnReadOnly Then Exit Sub
    
    With vsDiagXY
        If KeyCode = vbKeyF4 Then
            If .Col = col������� Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, col�������) <> "" Then
                If mlngPathState = 1 Or mlngPathState = 2 Then
                    If .TextMatrix(.Row, col�������) = "��Ժ���" And mlngDiagnosisType = 2 Or .TextMatrix(.Row, col�������) = "�������" And mlngDiagnosisType = 1 Then
                        If .TextMatrix(.Row, col�������) <> .TextMatrix(.Row - 1, col�������) Then
                            '��Ҫ��ϲ������
                            Exit Sub
                        End If
                    End If
                End If
                '�ϲ�·��
                If Not CheckMergePath(mlng����ID, mlng��ҳID, Val(.TextMatrix(.Row, col����)), Val(.TextMatrix(.Row, col����ID))) Then Exit Sub

                '����·������
                If mstrPathDiag <> "" And mlngPathState > 0 Then
                    If InStr("," & mstrPathDiag & ",", "," & .TextMatrix(.Row, col����) & "|" & Val(.TextMatrix(.Row, col����ID)) & "|" & Val(.TextMatrix(.Row, col���ID)) & ",") > 0 Then
                        '������ϲ������
                        Exit Sub
                    End If
                End If
                If mlngPathState = 2 And mblnIsPathOutTime Then
                    If .TextMatrix(.Row, col�������) = "��Ժ���" And mlngDiagnosisType <= 2 Then
                        '������ɵĳ�Ժ��ϲ������
                        Exit Sub
                    End If
                End If
                If MsgBox("ȷʵҪ������������Ϣ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    i = Val(.TextMatrix(.Row, col����))
                    .Cell(flexcpText, .Row, .FixedCols, .Row, .Cols - 1) = ""
                    .Cell(flexcpData, .Row, .FixedCols, .Row, .Cols - 1) = Empty
                    .TextMatrix(.Row, col����) = i
                    
                    '�����ͬ�������������
                    If .TextMatrix(.Row, col�������) = "" Then
                        .RemoveItem .Row
                    Else
                        If .Row + 1 <= .Rows - 1 Then
                            If .TextMatrix(.Row + 1, col�������) = "" Then
                                '��һ��Ϊ�ޱ����������ʱ�����ݲ����ƣ�����ǰ��Ϊ�б���ʱֻ�����
                                For i = .Row + 1 To .Rows - 1
                                    If Val(.TextMatrix(i, col����)) = Val(.TextMatrix(.Row, col����)) Then
                                        For j = .FixedCols To .Cols - 1
                                            .TextMatrix(i - 1, j) = .TextMatrix(i, j)
                                            .Cell(flexcpData, i - 1, j) = .Cell(flexcpData, i, j)
                                        Next
                                        .Cell(flexcpText, i, .FixedCols, i, .Cols - 1) = ""
                                        .Cell(flexcpData, i, .FixedCols, i, .Cols - 1) = Empty
                                        .TextMatrix(i, col����) = Val(.TextMatrix(.Row, col����))
                                        
                                        If i = .Rows - 1 Then
                                            If .TextMatrix(i, col�������) = "" Then .RemoveItem i
                                            Exit For
                                        ElseIf Val(.TextMatrix(i + 1, col����)) <> Val(.TextMatrix(i, col����)) Then
                                            If .TextMatrix(i, col�������) = "" Then .RemoveItem i
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    End If
                    
'                    .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = .BackColor
'                    .Cell(flexcpBackColor, GetRow(3), .FixedRows, GetRow(3), .Cols - 1) = &HC0FFC0
                    
                    '������Ϸ������
                    Call Set��ҽ������(.Row)
                    Call Set��ԭѧ
                    
                    mblnChange = True
                    .Tag = ""
                End If
            ElseIf .TextMatrix(.Row, col�������) = "" Then
                .RemoveItem .Row
            End If
        ElseIf KeyCode > 127 Then
            '���ֱ�����뺺�ֵ�����
            Call vsDiagXY_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsDiagXY_KeyPress(KeyAscii As Integer)
    If mbln��ʿվ Or mblnReadOnly Then Exit Sub
    
    With vsDiagXY
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call XYEnterNextCell
        ElseIf KeyAscii = 32 And (.Col = col�Ƿ�δ�� Or .Col = col�Ƿ�����) Then
            If XYCellEditable(.Row, .Col) Then
                KeyAscii = 0
                If .Col = col�Ƿ����� Then
                    .TextMatrix(.Row, .Col) = IIf(.TextMatrix(.Row, .Col) = "", "��", "")
                ElseIf .Col = col�Ƿ�δ�� Then
                    .TextMatrix(.Row, .Col) = IIf(.TextMatrix(.Row, .Col) = "", "��", "")
                End If
            End If
        Else
            If .Col = col������� Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsDiagXY_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" 'ʹ��ť״̬��������״̬
                End If
            End If
        End If
    End With
End Sub

Private Sub vsDiagXY_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsDiagXY.EditSelStart = 0
    vsDiagXY.EditSelLength = zlCommFun.ActualLen(vsDiagXY.EditText)
End Sub

Private Sub vsDiagXY_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not XYCellEditable(Row, Col) Then
        Cancel = True
    ElseIf Col = col�Ƿ�δ�� Or Col = col�Ƿ����� Then
        Cancel = True '��ֱ�ӱ༭
    End If
End Sub

Private Sub vsDiagXY_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String
    Dim str�Ա� As String, lngRow As Long
    
    With vsDiagXY
        If Col = col������� Then
            If optInput(0).Value Then
                '���������:��ҽ���ݣ�һ����Ͽ������ڶ������
                Set rsTmp = zlDatabase.ShowILLSelect(Me, "1", mlng����ID, , True, False)
            Else
                '7-�����ж���Y-�����ж����ⲿԭ��6-������ϣ�M-������̬ѧ���룻������ϣ�D-ICD-10��������
                Set rsTmp = zlDatabase.ShowILLSelect(Me, Decode(Val(.TextMatrix(Row, col����)), 7, "'Y'", 6, "'M,D'", "'D'"), mlng����ID, cboinfo(cbo�Ա�).Text, True)
            End If
            If Not rsTmp Is Nothing Then
                Call XYSetDiagInput(Row, rsTmp)
                Call XYEnterNextCell
            End If
        ElseIf Col = col���� Then
            lngRow = Row + 1: .AddItem "", lngRow
            .TextMatrix(lngRow, col����) = .TextMatrix(Row, col����)
            .Cell(flexcpBackColor, lngRow, col��ϱ���) = ColorUnEditCell      '����ɫ
            
'            .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = .BackColor
'            .Cell(flexcpBackColor, GetRow(3), .FixedRows, GetRow(3), .Cols - 1) = &HC0FFC0
            
            .Row = lngRow: .Col = col�������
            .ShowCell .Row, .Col
        ElseIf Col = colDel Then
            Call vsDiagXY_KeyDown(vbKeyDelete, 0)
        End If
    End With
End Sub

Private Sub vsDiagXY_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        mblnReturn = True
        
        With vsDiagXY
            If Col = col��Ժ��� Then
                KeyAscii = 0
                If .ComboIndex <> -1 Then
                    '��ʱ.TextMatrix��δ����,����ȡComboItem
                    .TextMatrix(Row, Col) = NeedName(.ComboItem(.ComboIndex))
                    If Not XYCellEditable(Row, col�Ƿ�δ��) Then
                        .TextMatrix(Row, col�Ƿ�δ��) = ""
                    End If
                    .Tag = ""
                    mblnChange = True
                    Call XYEnterNextCell
                End If
            End If
        End With
    Else
        mblnReturn = False
    End If
End Sub

Private Sub XYSetDiagInput(ByVal lngRow As Long, rsInput As ADODB.Recordset)
'���ܣ�������ҽ�����Ŀ������
    Dim rsTmp As New ADODB.Recordset
    Dim StrSQL As String, i As Long, j As Long
    Dim bln�ֻ��̶� As Boolean
    
    With vsDiagXY
        If Not rsInput Is Nothing Then
            For i = 1 To rsInput.RecordCount
                If i > 1 Then
                    '�����ж�ѡ�����ʱ�Ĵ���
                    If lngRow = .Rows - 1 Then
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, col����) = .TextMatrix(lngRow, col����)
                    End If
                    'ȷ����ǰ��ʾ��
                    If Val(.TextMatrix(lngRow + 1, col����)) = Val(.TextMatrix(lngRow, col����)) Then
                        For j = lngRow + 1 To .Rows - 1
                            If Val(.TextMatrix(j, col����)) = Val(.TextMatrix(lngRow, col����)) Then
                                lngRow = j
                                If .TextMatrix(j, col�������) = "" Then Exit For
                            Else
                                Exit For
                            End If
                        Next
                        If .TextMatrix(lngRow, col�������) <> "" Then
                            lngRow = lngRow + 1: .AddItem "", lngRow
                            .TextMatrix(lngRow, col����) = .TextMatrix(lngRow - 1, col����)
                        End If
                    Else
                        lngRow = lngRow + 1: .AddItem "", lngRow
                        .TextMatrix(lngRow, col����) = .TextMatrix(lngRow - 1, col����)
                    End If
                End If
                
                If .TextMatrix(lngRow, col�������) = "��Ժ���" Then
                    If Nvl(rsInput!����) = "" Then
                        bln�ֻ��̶� = False
                    Else
                        bln�ֻ��̶� = ((InStr("C", UCase(Left(rsInput!����, 1)))) > 0) Or ((InStr("D0", UCase(Left(rsInput!����, 2)))) > 0) Or ((InStr("D32.,D33.,", UCase(Left(rsInput!����, 4)))) > 0)
                    End If
                    cboinfo(cbo�ֻ��̶�).Enabled = bln�ֻ��̶�
                    lblInfo(lbl�ֻ��̶�).Enabled = bln�ֻ��̶�
                    lblInfo(lbl����������).Enabled = bln�ֻ��̶�
                    cboinfo(cbo����������).Enabled = bln�ֻ��̶�
                End If
                .TextMatrix(lngRow, col��ϱ���) = "" & rsInput!����
                .TextMatrix(lngRow, col�������) = "" & rsInput!����
                .Cell(flexcpData, lngRow, col�������) = .TextMatrix(lngRow, col�������)
                
                '�������ȷ������,����ݼ���ȷ�����
                If optInput(0).Value Then
                    .TextMatrix(lngRow, col���ID) = rsInput!��ĿID
                    .TextMatrix(lngRow, col����ID) = ""
                    StrSQL = "Select ����ID as ID From ������϶��� Where ���ID=[1]"
                Else
                    .TextMatrix(lngRow, col����ID) = rsInput!��ĿID
                    .TextMatrix(lngRow, col���ID) = ""
                    StrSQL = "Select ���ID as ID From ������϶��� Where ����ID=[1]"
                End If
                On Error GoTo errH
                Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, Val(rsInput!��ĿID))
                If Not rsTmp.EOF Then
                    If optInput(0).Value Then
                        .TextMatrix(lngRow, col����ID) = Nvl(rsTmp!ID)
                    Else
                        .TextMatrix(lngRow, col���ID) = Nvl(rsTmp!ID)
                    End If
                End If
                
                rsInput.MoveNext
            Next
        Else
            .TextMatrix(lngRow, col��ϱ���) = ""
            .TextMatrix(lngRow, col�������) = .EditText
            .Cell(flexcpData, lngRow, col�������) = .TextMatrix(lngRow, col�������)
            .TextMatrix(lngRow, col���ID) = ""
            .TextMatrix(lngRow, col����ID) = ""
        End If
        
        .Cell(flexcpForeColor, 1, col�Ƿ�����, .Rows - 1, col�Ƿ�����) = vbRed
        
        '������Ϸ������
        Call Set��ҽ������(lngRow)
        Call Set��ԭѧ
        
        .Tag = ""
        mblnChange = True
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Set��ҽ������(ByVal lngRow As Long)
    With vsDiagXY
        If lngRow > .Rows - 1 Then Exit Sub
        
        '��Ϸ������
        If .TextMatrix(lngRow, 0) = "�������" Or .TextMatrix(lngRow, 0) = "��Ժ���" Then
            Call Set��Ϸ������(cbo�������Ժ)
            Call Set��Ϸ������(cbo��������Ժ)
        End If
        If .TextMatrix(lngRow, 0) = "��Ժ���" Or .TextMatrix(lngRow, 0) = "��Ժ���" Then
            Call Set��Ϸ������(cbo��Ժ���Ժ)
            Call Set��Ϸ������(cbo��������Ժ)
        End If
        If .TextMatrix(lngRow, 0) = "�������" Then
            Call Set��Ϸ������(cbo�����벡��)
            Call Set��Ϸ������(cbo�ٴ��벡��)
            If vsDiagXY.TextMatrix(lngRow, col�������) <> "" Then
                txtInfo(txt�����).Enabled = True
                txtInfo(txt�����).BackColor = vbWindowBackground
            Else
                txtInfo(txt�����).Enabled = False
                txtInfo(txt�����).BackColor = &H8000000F
            End If
        End If
    End With
End Sub

Private Sub Set��ҽ������(ByVal lngRow As Long)
    With vsDiagZY
        If lngRow > .Rows - 1 Then Exit Sub
        If .TextMatrix(lngRow, 0) = "�������" Or .TextMatrix(lngRow, 0) = "��Ҫ���" Then
            Call Set��Ϸ������(cbo��ҽ�������Ժ)
        End If
        If .TextMatrix(lngRow, 0) = "��Ժ���" Or .TextMatrix(lngRow, 0) = "��Ҫ���" Then
            Call Set��Ϸ������(cbo��ҽ��Ժ���Ժ)
        End If
    End With
End Sub

Private Sub Set��ԭѧ()
    With vsDiagXY
        'Ժ�ڸ�Ⱦ�벡ԭѧ���
        If Trim(.TextMatrix(GetRow(5), col�������)) = "" Then
            chkInfo(chk��ԭѧ).Value = 0
            chkInfo(chk��ԭѧ).Enabled = False
            chkInfo(chk��ԭѧ).TabStop = False
            Call chkInfo_Click(chk��ԭѧ)
        ElseIf Not chkInfo(chk��ԭѧ).Enabled Then
            chkInfo(chk��ԭѧ).Enabled = True
            chkInfo(chk��ԭѧ).TabStop = True
        End If
    End With
End Sub

Private Function GetRow(ByVal lng������� As Long) As Long
'���ܣ�����ָ��������͵ĵ�һ�����
    If InStr(",11,12,13,", "," & lng������� & ",") > 0 Then
        GetRow = vsDiagZY.FindRow(CStr(lng�������), , colzy����)
    Else
        GetRow = vsDiagXY.FindRow(CStr(lng�������), , col����)
    End If
End Function

Private Sub Set��Ϸ������(ByVal intIdx As Integer)
'���ܣ�����Ϸ����������ȱʡֵ�����Լ�����Ƿ��������
'������intIdx=Ҫ���õķ�������ؼ�
    Dim i As Long
    
    With vsDiagXY
        '�������Ժ��������Ϻͳ�Ժ�����ͬʱ"����"������һ��������ʱ"���϶�"����ͬʱ"������"
        If intIdx = cbo�������Ժ Then
            If Trim(.TextMatrix(GetRow(1), col�������)) = "" And Trim(.TextMatrix(GetRow(3), col�������)) = "" Then
                Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 1) '���Ը�ʱȱʡΪ����
            Else
                If Trim(.TextMatrix(GetRow(1), col�������)) = "" Or Trim(.TextMatrix(GetRow(3), col�������)) = "" Then
                    Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 3)
                ElseIf .TextMatrix(GetRow(1), col�������) <> .TextMatrix(GetRow(3), col�������) Then
                    Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 2)
                ElseIf .TextMatrix(GetRow(1), col�������) = .TextMatrix(GetRow(3), col�������) Then
                    Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 1)
                End If
            End If
        End If
        
        '��Ժ���Ժ����Ժ��Ϻͳ�Ժ�����ͬʱ"����"������һ��������ʱ"���϶�"����ͬʱ"������"
        If intIdx = cbo��Ժ���Ժ Then
            If Trim(.TextMatrix(GetRow(2), col�������)) = "" And Trim(.TextMatrix(GetRow(3), col�������)) = "" Then
                Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 1) '���Ը�ʱȱʡΪ����
            Else
                If Trim(.TextMatrix(GetRow(2), col�������)) = "" Or Trim(.TextMatrix(GetRow(3), col�������)) = "" Then
                    Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 3)
                ElseIf .TextMatrix(GetRow(2), col�������) <> .TextMatrix(GetRow(3), col�������) Then
                    Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 2)
                ElseIf .TextMatrix(GetRow(2), col�������) = .TextMatrix(GetRow(3), col�������) Then
                    Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 1)
                End If
            End If
        End If
        
        '��������Ժ��������Ϻ���Ժ�����ͬʱ"����"������һ��������ʱ"���϶�"����ͬʱ"������"
        If intIdx = cbo��������Ժ Then
            If Trim(.TextMatrix(GetRow(1), col�������)) = "" And Trim(.TextMatrix(GetRow(2), col�������)) = "" Then
                Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 1) '���Ը�ʱȱʡΪ����
            Else
                If Trim(.TextMatrix(GetRow(1), col�������)) = "" Or Trim(.TextMatrix(GetRow(2), col�������)) = "" Then
                    Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 3)
                ElseIf .TextMatrix(GetRow(1), col�������) <> .TextMatrix(GetRow(2), col�������) Then
                    Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 2)
                ElseIf .TextMatrix(GetRow(1), col�������) = .TextMatrix(GetRow(2), col�������) Then
                    Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 1)
                End If
            End If
        End If
        
        '�����벡���ٴ��벡��¼�벡����Ϻ����¼�룬ȱʡΪ���ϡ�
        If intIdx = cbo�����벡�� Or intIdx = cbo�ٴ��벡�� Then
            cboinfo(intIdx).Enabled = .TextMatrix(GetRow(6), col�������) <> ""
            If Not cboinfo(intIdx).Enabled Then
                Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 0) '�����Ը�ʱȱʡΪδ��
                cboinfo(intIdx).BackColor = vbButtonFace
            Else
                Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 1)
                cboinfo(intIdx).BackColor = vbWindowBackground
            End If
        End If
    End With
    
    '�ٴ���ʬ�죺��ѡʬ������¼�룬ȱʡΪ���ϡ�
    If intIdx = cbo�ٴ���ʬ�� Then
        cboinfo(intIdx).Enabled = chkInfo(chkʬ��).Value = 1
        If Not cboinfo(intIdx).Enabled Then
            Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 0) '�����Ը�ʱȱʡΪδ��
            cboinfo(intIdx).BackColor = vbButtonFace
        Else
            Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 1)
            cboinfo(intIdx).BackColor = vbWindowBackground
        End If
    End If
    
    '��ǰ����������������������¼�룬ȱʡΪ���ϡ�
    If intIdx = cbo��ǰ������ Then
        With vsOPS
            For i = .FixedRows To .Rows - 1
                If Trim(.TextMatrix(i, col��������)) <> "" Then Exit For
            Next
            If Not i <= .Rows - 1 Then
                Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 0) '�����Ը�ʱȱʡΪδ��
            Else
                Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 1)
            End If
        End With
    End If
    
    With vsDiagZY
        '��ҽ�������Ժ��������Ϻͳ�Ժ�����ͬʱ"����"������һ��������ʱ"���϶�"����ͬʱ"������"
        If intIdx = cbo��ҽ�������Ժ Then
            If Trim(.TextMatrix(GetRow(11), col�������)) = "" And Trim(.TextMatrix(GetRow(13), col�������)) = "" Then
                Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 1) '���Ը�ʱȱʡΪ����
            Else
                If Trim(.TextMatrix(GetRow(11), col�������)) = "" Or Trim(.TextMatrix(GetRow(13), col�������)) = "" Then
                    Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 3)
                ElseIf .TextMatrix(GetRow(11), col�������) <> .TextMatrix(GetRow(13), col�������) Then
                    Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 2)
                ElseIf .TextMatrix(GetRow(11), col�������) = .TextMatrix(GetRow(13), col�������) Then
                    Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 1)
                End If
            End If
        End If
        
        '��ҽ��Ժ���Ժ����Ժ��Ϻͳ�Ժ�����ͬʱ"����"������һ��������ʱ"���϶�"����ͬʱ"������"
        If intIdx = cbo��ҽ��Ժ���Ժ Then
            If Trim(.TextMatrix(GetRow(12), col�������)) = "" And Trim(.TextMatrix(GetRow(13), col�������)) = "" Then
                Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 1) '���Ը�ʱȱʡΪ����
            Else
                If Trim(.TextMatrix(GetRow(12), col�������)) = "" Or Trim(.TextMatrix(GetRow(13), col�������)) = "" Then
                    Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 3)
                ElseIf .TextMatrix(GetRow(12), col�������) <> .TextMatrix(GetRow(13), col�������) Then
                    Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 2)
                ElseIf .TextMatrix(GetRow(12), col�������) = .TextMatrix(GetRow(13), col�������) Then
                    Call zlControl.CboSetIndex(cboinfo(intIdx).hwnd, 1)
                End If
            End If
        End If
    End With
End Sub

Private Sub KSSEnterNextCell()
    With vsKSS
        If .Row = .Rows - 1 And .Col = .Cols - 1 And .TextMatrix(.Row, .FixedCols) = "" Then
            Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        Else
            If .Row + 1 > .Rows - 1 And .Col = .Cols - 1 Then
                If .Rows - .FixedRows >= 10 Then
                    Call zlCommFun.PressKey(vbKeyTab)
                    Exit Sub
                End If
                .AddItem "": Call SetKSSSerial
            End If
            If .Col = .Cols - 1 Then
                .Row = .Row + 1
                .Col = .FixedCols
                .ShowCell .Row, .Col
            Else
                .Col = .Col + 1
                .ShowCell .Row, .Col
            End If
        End If
    End With
End Sub

Private Sub KSSSetDiagInput(ByVal lngRow As Long, rsInput As ADODB.Recordset)
'���ܣ�����������Ŀ������
    With vsKSS
        If Not rsInput Is Nothing Then
            '�ж��Ƿ����޸�
            If .RowData(lngRow) & "" <> "" Then
                If InStr(mstrDelete, .RowData(lngRow) & "") <= 0 Then
                    mstrDelete = mstrDelete & IIf(mstrDelete <> "", ",", "") & .RowData(lngRow)
                End If
            End If
            .TextMatrix(lngRow, 1) = Nvl(rsInput!����)
            .RowData(lngRow) = Val(rsInput!ID)
        Else
            .TextMatrix(lngRow, 1) = .EditText
        End If
        .Cell(flexcpData, lngRow, 1) = .TextMatrix(lngRow, 1)
        mblnChange = True
        .Tag = ""
    End With
End Sub

Private Sub TSJCEnterNextCell()
    With vsTSJC
        If .Row = .Rows - 1 Then
            Call zlCommFun.PressKey(vbKeyTab)
        Else
            If .Row + 1 > .Rows - 1 Then
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                .Row = .Row + 1
            End If
        End If
    End With
End Sub

Private Sub TSJCSetDiagInput(ByVal lngRow As Long, rsInput As ADODB.Recordset)
'���ܣ�������������Ŀ������
    With vsTSJC
        If Not rsInput Is Nothing Then
            .TextMatrix(lngRow, 1) = Nvl(rsInput!����)
        Else
            .TextMatrix(lngRow, 1) = .EditText
        End If
        .Cell(flexcpData, lngRow, 1) = .TextMatrix(lngRow, 1)
        mblnChange = True
    End With
End Sub

Private Sub XYEnterNextCell()
    Dim i As Long, j As Long
    
    With vsDiagXY
        '����һ��Ԫ��ʼѭ������
        For i = .Row To .Rows - 1
            For j = IIf(i = .Row, .Col + 1, col�������) To col����
                If XYCellEditable(i, j) And .ColWidth(j) <> 0 Then Exit For
            Next
            If j <= col���� Then Exit For
        Next
        If i <= .Rows - 1 Then
            .Row = i: .Col = j
            .ShowCell .Row, .Col
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End With
End Sub

Private Function XYCellEditable(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
    With vsDiagXY
        '�����в��ɱ༭
        If .ColHidden(lngCol) Then Exit Function
        
        If lngCol = col������� And (mlngPathState = 1 Or mlngPathState = 2) Then
            If .TextMatrix(lngRow, col�������) = "��Ժ���" And mlngDiagnosisType = 2 Or .TextMatrix(lngRow, col�������) = "�������" And mlngDiagnosisType = 1 Then
                If .TextMatrix(lngRow, col�������) <> "" And .TextMatrix(lngRow, col�������) <> .TextMatrix(lngRow - 1, col�������) Then
                    '��Ҫ��ϲ������
                    Exit Function
                End If
            End If
            '�ϲ�·��
            If Not CheckMergePath(mlng����ID, mlng��ҳID, Val(.TextMatrix(lngRow, col����)), Val(.TextMatrix(lngRow, col����ID))) Then Exit Function
        End If
        If lngCol = col������� Then
            '����·������
            If mstrPathDiag <> "" And mlngPathState > 0 Then
                If InStr("," & mstrPathDiag & ",", "," & .TextMatrix(.Row, col����) & "|" & Val(.TextMatrix(.Row, col����ID)) & "|" & Val(.TextMatrix(.Row, col���ID)) & ",") > 0 Then
                    '������ϲ������
                    Exit Function
                End If
            End If
        End If
        If lngCol = col������� And mlngPathState = 2 And mblnIsPathOutTime Then
            If .TextMatrix(.Row, col�������) = "��Ժ���" And mlngDiagnosisType <= 2 Then
                '������ɵĳ�Ժ��ϲ������
                Exit Function
            End If
        End If
        '�������������
        If .TextMatrix(lngRow, col�������) = "" Then
            If lngCol = col��Ժ��� Or lngCol = col��ע Or lngCol = col�Ƿ�δ�� Or lngCol = col�Ƿ����� Or lngCol = col���� Then
                Exit Function
            End If
        End If
        If lngCol = col��ϱ��� Then Exit Function
        
        If lngCol = col���� Then
            If Val(.TextMatrix(lngRow, col����)) = 3 Then
                If .TextMatrix(lngRow, col�������) = "��Ժ���" Then Exit Function
            End If
        End If
        
        '��Ժ��Ϻ�Ժ�ڸ�Ⱦ���������Ժ���(��Ϊ����Ժ�ڸ�Ⱦ�ڳ�Ժʱ�Ѿ���ת��������)
        If Val(.TextMatrix(lngRow, col����)) = 3 Or Val(.TextMatrix(lngRow, col����)) = 5 Or Val(.TextMatrix(lngRow, col����)) = 10 Then
            '��Ժ��ϱ�����������(��δ����ʱ)
            If .TextMatrix(lngRow, col�������) = "" And Val(.TextMatrix(lngRow, col����)) = 3 Then
                If Val(.TextMatrix(lngRow - 1, col����)) = 3 And .TextMatrix(lngRow - 1, col�������) = "" Then
                    Exit Function
                End If
            End If

            '��Ժ���Ϊ"����"ʱ�ſ��������Ƿ�δ��
            If .TextMatrix(lngRow, col��Ժ���) <> "����" And lngCol = col�Ƿ�δ�� Then
                Exit Function
            End If
        ElseIf lngCol = col��Ժ��� Or lngCol = col�Ƿ�δ�� Then
            Exit Function
        End If
        
        '��Ժ����ֻ���ڳ�Ժ��Ϻ������������д
        If lngCol = col��Ժ���� Then
            If .TextMatrix(lngRow, col����) <> "3" Then
                Exit Function
            End If
        End If
    End With
    XYCellEditable = True
End Function

Private Function ZYCellEditable(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
    With vsDiagZY
        '�����в��ɱ༭
        If .ColHidden(lngCol) Then Exit Function
        
        If lngCol = col������� And (mlngPathState = 1 Or mlngPathState = 2) Then
            If .TextMatrix(lngRow, col�������) = "��Ժ���" And mlngDiagnosisType = 12 Or .TextMatrix(lngRow, col�������) = "�������" And mlngDiagnosisType = 11 Then
                If .TextMatrix(lngRow, col�������) <> "" And .TextMatrix(lngRow, col�������) <> .TextMatrix(lngRow - 1, col�������) Then
                    '��Ҫ��ϲ������
                    Exit Function
                End If
            End If
            '�ϲ�·��
            If Not CheckMergePath(mlng����ID, mlng��ҳID, Val(.TextMatrix(lngRow, colzy����)), Val(.TextMatrix(lngRow, colzy����ID))) Then Exit Function
        End If
        If lngCol = col������� Then
            '����·������
            If mstrPathDiag <> "" And mlngPathState > 0 Then
                If InStr("," & mstrPathDiag & ",", "," & .TextMatrix(.Row, colzy����) & "|" & Val(.TextMatrix(.Row, col����ID)) & "|" & Val(.TextMatrix(.Row, col���ID)) & ",") > 0 Then
                    '������ϲ������
                    Exit Function
                End If
            End If
        End If
        If lngCol = col������� And mlngPathState = 2 And mblnIsPathOutTime Then
            If .TextMatrix(.Row, col�������) = "��Ҫ���" And mlngDiagnosisType > 10 Then
                '������ɵĳ�Ժ��ϲ������
                Exit Function
            End If
        End If
        '�������������
        If .TextMatrix(lngRow, col�������) = "" Then
            If lngCol = col��Ժ��� Or lngCol = col��ע Or lngCol = colzy���� Then Exit Function
        End If
        If lngCol = col��ϱ��� Then Exit Function
        
        If lngCol = colzy���� Then
            If Val(.TextMatrix(lngRow, colzy����)) = 13 Then
                If .TextMatrix(lngRow, col�������) = "��Ҫ���" Then Exit Function
            End If
        End If
        
        If Val(.TextMatrix(lngRow, colzy����)) = 13 Then
            '��Ժ��ϱ�����������(��δ����ʱ)
            If .TextMatrix(lngRow, col�������) = "" Then
                If Val(.TextMatrix(lngRow - 1, colzy����)) = 13 And .TextMatrix(lngRow - 1, col�������) = "" Then
                    Exit Function
                End If
            End If
        ElseIf lngCol = col��Ժ��� Then
            '�ǳ�Ժ���ʱ����������
            If Val(.TextMatrix(lngRow, colzy����)) <> 13 Then Exit Function
        End If
        '��Ժ����ֻ������Ҫ��Ϻ������������д
        If lngCol = col��Ժ���� Then
            If .TextMatrix(lngRow, colzy����) <> "13" Then
                Exit Function
            End If
        End If
        '���������������֤��
        If lngCol = col��ҽ֤�� Then
            If .TextMatrix(lngRow, col�������) = "" Then Exit Function
        End If
    End With
    ZYCellEditable = True
End Function

Private Sub ZYEnterNextCell()
    Dim i As Long, j As Long
    
    With vsDiagZY
        '����һ��Ԫ��ʼѭ������
        For i = .Row To .Rows - 1
            For j = IIf(i = .Row, .Col + 1, col�������) To colzy����
                If ZYCellEditable(i, j) And .ColWidth(j) <> 0 Then Exit For
            Next
            If j <= colzy���� Then Exit For
        Next
        If i <= .Rows - 1 Then
            .Row = i: .Col = j
            .ShowCell .Row, .Col
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End With
End Sub

Private Sub ZYSetDiagInput(ByVal lngRow As Long, rsInput As ADODB.Recordset)
'���ܣ�������ҽ�����Ŀ������
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String, blnCancel As Boolean
    Dim vPoint As POINTAPI
    Dim i As Long, j As Long
    Dim strTmp As String
    
    With vsDiagZY
        If Not rsInput Is Nothing Then
            For i = 1 To rsInput.RecordCount
                If i > 1 Then
                    '�������ѡ�����ʱ�Ĵ���
                    If lngRow = .Rows - 1 Then
                        .Rows = .Rows + 1
                        .TextMatrix(.Rows - 1, colzy����) = .TextMatrix(lngRow, colzy����)
                    End If
                    'ȷ����ǰ��ʾ��
                    If Val(.TextMatrix(lngRow + 1, colzy����)) = Val(.TextMatrix(lngRow, colzy����)) Then
                        For j = lngRow + 1 To .Rows - 1
                            If Val(.TextMatrix(j, colzy����)) = Val(.TextMatrix(lngRow, colzy����)) Then
                                lngRow = j
                                If .TextMatrix(j, col�������) = "" Then Exit For
                            Else
                                Exit For
                            End If
                        Next
                        If .TextMatrix(lngRow, col�������) <> "" Then
                            lngRow = lngRow + 1: .AddItem "", lngRow
                            .TextMatrix(lngRow, colzy����) = .TextMatrix(lngRow - 1, colzy����)
                        End If
                    Else
                        lngRow = lngRow + 1: .AddItem "", lngRow
                        .TextMatrix(lngRow, colzy����) = .TextMatrix(lngRow - 1, colzy����)
                    End If
                End If
                
                If InStr(.TextMatrix(lngRow, col�������), "(") > 0 And InStr(.TextMatrix(lngRow, col�������), ")") > 0 Then
                    strTmp = Mid(.TextMatrix(lngRow, col�������), InStrRev(.TextMatrix(lngRow, col�������), "("))
                End If
                                        
                .TextMatrix(lngRow, col��ϱ���) = "" & rsInput!����
                .TextMatrix(lngRow, col�������) = "" & rsInput!���� & strTmp
                .Cell(flexcpData, lngRow, col�������) = .TextMatrix(lngRow, col�������)
                                
                
                '�������ȷ������,����ݼ���ȷ�����
                If optInput(2).Value Then
                    .TextMatrix(lngRow, colzy���ID) = rsInput!��ĿID
                    .TextMatrix(lngRow, colzy����ID) = ""
                    StrSQL = "Select ����ID as ID From ������϶��� Where ���ID=[1]"
                Else
                    .TextMatrix(lngRow, colzy����ID) = rsInput!��ĿID
                    .TextMatrix(lngRow, colzy���ID) = ""
                    StrSQL = "Select ���ID as ID From ������϶��� Where ����ID=[1]"
                End If
                Set rsTmp = New ADODB.Recordset
                On Error GoTo errH
                Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, Val(rsInput!��ĿID))
                If Not rsTmp.EOF Then
                    If optInput(2).Value Then
                        .TextMatrix(lngRow, colzy����ID) = Nvl(rsTmp!ID)
                    Else
                        .TextMatrix(lngRow, colzy���ID) = Nvl(rsTmp!ID)
                    End If
                End If
                
                '��ҽ���ݼ�����ϲο�ȡ֤��
                Call Set��ҽ֤��(lngRow, Val(.TextMatrix(lngRow, colzy���ID)))
                
                rsInput.MoveNext
            Next
        Else
            .TextMatrix(lngRow, col��ϱ���) = ""
            .TextMatrix(lngRow, col�������) = .EditText
            .Cell(flexcpData, lngRow, col�������) = .TextMatrix(lngRow, col�������)
            .TextMatrix(lngRow, colzy���ID) = ""
            .TextMatrix(lngRow, colzy����ID) = ""
            .TextMatrix(lngRow, colzy֤��ID) = ""
        End If
        
        '������Ϸ������
        Call Set��ҽ������(lngRow)
        
        .Tag = ""
        mblnChange = True
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function Set��ҽ֤��(ByVal lngRow As Long, ByVal lng���ID As Long, Optional ByVal rsInput As Recordset) As Boolean
'���ܣ���ҽ���ݼ�����ϲο�ȡ֤��
'������rsInput-�����Ϊ�գ������ָ������ҩ֤���¼��
'���أ��Ƿ��ж�Ӧ��ϵ
    Dim rsTmp As Recordset
    Dim StrSQL As String
    Dim blnCancel As Boolean
    Dim vPoint As POINTAPI
    Dim strTmp As String
    
    With vsDiagZY
        'ȥ�����е�֤��
        If InStr(.TextMatrix(lngRow, col�������), "(") > 0 And InStr(.TextMatrix(lngRow, col�������), ")") > 0 Then
            strTmp = Mid(.TextMatrix(lngRow, col�������), 1, InStrRev(.TextMatrix(lngRow, col�������), "(") - 1)
        Else
            strTmp = .TextMatrix(lngRow, col�������)
        End If
        If rsInput Is Nothing Then
            If lng���ID <> 0 Then
                StrSQL = "Select Distinct a.֤����� as ID,a.֤��ID,a.֤������,b.���� as ֤�����" & _
                    " From ������ϲο� A,��������Ŀ¼ B" & _
                    " Where a.֤��ID=b.ID(+) And a.���ID=[1] And a.֤������ is Not NULL" & _
                    " Order by a.֤�����"
                vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = Nothing
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 0, "��ҽ֤��", False, "", "", False, False, True, _
                    vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, lng���ID)
                If Not rsTmp Is Nothing Then
                    .TextMatrix(lngRow, colzy֤��ID) = Nvl(rsTmp!֤��id)
                    If Not IsNull(rsTmp!֤������) Then
                        .TextMatrix(lngRow, col�������) = strTmp
                        .Cell(flexcpData, lngRow, col�������) = .TextMatrix(lngRow, col�������)
                        .TextMatrix(lngRow, col��ҽ֤��) = Nvl(rsTmp!֤������)
                        .Cell(flexcpData, lngRow, col��ҽ֤��) = .TextMatrix(lngRow, col��ҽ֤��)
                        If .EditText <> "" Then .EditText = .TextMatrix(lngRow, col��ҽ֤��)
                        mblnChange = True
                        .Tag = ""
                    End If
                    Set��ҽ֤�� = True
                Else
                    If blnCancel Then
                        Set��ҽ֤�� = True
                        If .EditText <> "" Then .EditText = .Cell(flexcpData, lngRow, col��ҽ֤��)
                    Else
                        Set��ҽ֤�� = False
                    End If
                End If
            Else
                Set��ҽ֤�� = False
            End If
        Else
            .TextMatrix(lngRow, colzy֤��ID) = Nvl(rsInput!��ĿID)
            .TextMatrix(lngRow, col�������) = strTmp
            .Cell(flexcpData, lngRow, col�������) = .TextMatrix(lngRow, col�������)
            .TextMatrix(lngRow, col��ҽ֤��) = Nvl(rsInput!����)
            .Cell(flexcpData, lngRow, col��ҽ֤��) = .TextMatrix(lngRow, col��ҽ֤��)
            If .EditText <> "" Then .EditText = .TextMatrix(lngRow, col��ҽ֤��)
            .Tag = ""
            mblnChange = True
        End If
    End With
End Function

Private Function GetSQL(ByVal intType As Integer, ByVal strInput As String, ByRef str�Ա� As String, Optional ByVal strOtherInfo As String) As String
'���ܣ���ò�ѯ��ҽ��ϵ�SQL
'������intType:��ȡ��SQL����,0-��ҽ��ϣ�1-��ҽ��ϣ�2-��������
'    strInput-��ѯ������str�Ա�--���˵��Ա�
'   strOtherInfo:��ҽ���-������������
'���أ�strsql--��ѯ��ϵ�SQL
    Dim StrSQL As String
    
    If cboinfo(cbo�Ա�).Text Like "*��*" Then
        str�Ա� = "��"
    ElseIf cboinfo(cbo�Ա�).Text Like "*Ů*" Then
        str�Ա� = "Ů"
    End If
    
    Select Case intType
        Case 0 '��ҽ���
            If optInput(0).Value Then
                '���������:��ҽ���ݣ�һ����Ͽ������ڶ������
                If zlCommFun.IsCharChinese(strInput) Then
                    StrSQL = "B.���� Like [2]" '���뺺��ʱֻƥ������
                Else
                    StrSQL = "A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2]"
                End If
                StrSQL = _
                    " Select Distinct A.ID,A.ID as ��ĿID,A.����,A.����,A.˵��,A.����" & _
                    " From �������Ŀ¼ A,������ϱ��� B" & _
                    " Where A.ID=B.���ID And A.���=1" & _
                    " And B.����=[5] And (" & StrSQL & ")" & _
                    " Order by A.����"
            Else
                If zlCommFun.IsCharChinese(strInput) Then
                    StrSQL = "���� Like [2]" '���뺺��ʱֻƥ������
                Else
                    StrSQL = "���� Like [1] Or ���� Like [2] Or " & IIf(mint���� = 0, "����", "�����") & " Like [2]"
                End If
                StrSQL = _
                    " Select ID,ID as ��ĿID,����,����,����," & IIf(mint���� = 0, "����", "����� as ����") & ",˵��" & _
                    " From ��������Ŀ¼ Where Instr([3],���)>0 And (" & StrSQL & ")" & _
                    IIf(str�Ա� <> "", " And (�Ա�����=[4] Or �Ա����� is NULL)", "") & _
                    " And (����ʱ�� is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " Order by ����"
            End If
        
        Case 1 '��ҽ���
            If optInput(2).Value And strOtherInfo <> "Z" Then
                '���������:��ҽ���ݣ�һ����Ͽ������ڶ������
                If zlCommFun.IsCharChinese(strInput) Then
                    StrSQL = "B.���� Like [2]" '���뺺��ʱֻƥ������
                Else
                    StrSQL = "A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2]"
                End If
                StrSQL = _
                    " Select Distinct A.ID,A.ID as ��ĿID,A.����,A.����,A.˵��,A.����" & _
                    " From �������Ŀ¼ A,������ϱ��� B" & _
                    " Where A.ID=B.���ID And A.���=2" & _
                    " And B.����=[4] And (" & StrSQL & ")" & _
                    " Order by A.����"
            Else
                'B-��ҽ��������
                If zlCommFun.IsCharChinese(strInput) Then
                    StrSQL = "���� Like [2]" '���뺺��ʱֻƥ������
                Else
                    StrSQL = "���� Like [1] Or ���� Like [2] Or " & IIf(mint���� = 0, "����", "�����") & " Like [2]"
                End If
                StrSQL = _
                    " Select ID,ID as ��ĿID,����,����,����," & IIf(mint���� = 0, "����", "����� as ����") & ",˵��" & _
                    " From ��������Ŀ¼" & _
                    " Where ���='" & IIf(strOtherInfo = "", "B", strOtherInfo) & "' And (" & StrSQL & ")" & _
                    IIf(str�Ա� <> "", " And (�Ա�����=[3] Or �Ա����� is NULL)", "") & _
                    " And (����ʱ�� is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " Order by ����"
            End If
        Case 2 '��������
            If optInput(4).Value Then
                '��������Ŀ����
                StrSQL = "Select distinct A.ID,A.����,A.����,A.�������� as ��ģ" & _
                    " From ������ĿĿ¼ A,������Ŀ���� B" & _
                    " Where A.���='F' And A.������� IN(2,3) And A.ID=B.������ĿID" & _
                    IIf(str�Ա� <> "", " And Nvl(A.�����Ա�,0) IN(0,[4])", "") & _
                    " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                    " And (A.���� Like [1] Or A.���� Like [2] Or B.���� Like [2] Or B.���� Like [2])" & _
                    " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                    " Order by A.����"
            Else
                '��ICD9-CM3����
                StrSQL = _
                    " Select distinct ID,����,����,����,����,˵��" & _
                    " From ��������Ŀ¼ Where ���='S'" & _
                    IIf(str�Ա� <> "", " And (�Ա�����=[3] Or �Ա����� is NULL)", "") & _
                    " And (����ʱ�� is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " And (���� Like [1] Or ���� Like [2] Or ���� Like [2])" & _
                    " Order by ����"
            End If
    End Select
    GetSQL = StrSQL
End Function

Private Sub vsDiagXY_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String, blnCancel As Boolean
    Dim str�Ա� As String, int������� As Integer
    Dim strInput As String, vPoint As POINTAPI
    
    With vsDiagXY
        If Col = col������� Then
            '.Cell(flexcpData, Row, Col) <> ""�ų����лس�
            If .EditText = "" And .Cell(flexcpData, Row, Col) <> "" Then
                .EditText = ""
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call XYEnterNextCell
            ElseIf .TextMatrix(Row, col��ϱ���) <> "" And .Cell(flexcpData, Row, Col) <> "" And .EditText Like "*" & .Cell(flexcpData, Row, Col) & "*" Then
                '�жϼ���ǰ׺��������Ƿ������������ϱ���
                strInput = UCase(.EditText)
                StrSQL = GetSQL(0, strInput, str�Ա�)
                On Error GoTo errH
                Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, strInput, strInput, _
                        Decode(Val(.TextMatrix(Row, col����)), 7, "'Y'", 6, "'M,D'", "'D'"), str�Ա�, mint���� + 1)
                If rsTmp.RecordCount <> 1 Then
                    '�����ڱ�׼������ǰ�����븽����Ϣ
                    .TextMatrix(Row, col�������) = .EditText
                Else
                    Call XYSetDiagInput(Row, rsTmp)
                    .EditText = .Text
                End If
                '������.Cell(flexcpData, Row, Col)���Ա��޸�����ʱ�ٴ�ʹ��like�ж�
                .Tag = ""
                mblnChange = True
            Else
                If Val(.TextMatrix(Row, col����)) = 1 Then
                    int������� = Val(Mid(gstr�������, 1, 1))
                Else
                    int������� = Val(Mid(gstr�������, 2, 1))
                End If
                If int������� = 0 Then int������� = 1
                
                strInput = UCase(.EditText)
                StrSQL = GetSQL(0, strInput, str�Ա�)
                If int������� = 1 And zlCommFun.IsCharChinese(strInput) Then
                    '�����ж��룺Y-�����ж����ⲿԭ�򣻲����������M-������̬ѧ���룻������ϣ�D-ICD-10��������
                    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, strInput & "%", mstrLike & strInput & "%", _
                        Decode(Val(.TextMatrix(Row, col����)), 7, "'Y'", 6, "'M,D'", "'D'"), str�Ա�, mint���� + 1)
                    If rsTmp.EOF Then
                        Set rsTmp = Nothing
                    ElseIf rsTmp.RecordCount > 1 Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, strInput, strInput, _
                        Decode(Val(.TextMatrix(Row, col����)), 7, "'Y'", 6, "'M,D'", "'D'"), str�Ա�, mint���� + 1)
                        If rsTmp.RecordCount <> 1 Then Set rsTmp = Nothing '����¼��ʱ�ж��ƥ�䲻����ѡ��
                    End If
                    Call XYSetDiagInput(Row, rsTmp)
                    .EditText = .Text
                    If mblnReturn And rsTmp Is Nothing Then Call XYEnterNextCell '��������¼��ʱ���ݲ�������һ�У���Ϊ���ܻ�Ҫ����������
                Else
                    vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 0, IIf(optInput(0).Value, "�������", "��������"), _
                        False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                        strInput & "%", mstrLike & strInput & "%", Decode(Val(.TextMatrix(Row, col����)), 7, "'Y'", 6, "'M,D'", "'D'"), str�Ա�, mint���� + 1)
                    If blnCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                        Cancel = True
                    Else
                        '���������뷽ʽ
                        If rsTmp Is Nothing And ((int������� = 2 Or int������� = 3 And mint���� <> 0)) Then
                            MsgBox "û���ҵ�������ƥ������ݡ�", vbInformation, gstrSysName
                            Cancel = True
                        Else
                            Call XYSetDiagInput(Row, rsTmp): .EditText = .Text
                            'If mblnReturn Then Call XYEnterNextCell    '�ݲ�������һ�У���Ϊ���ܻ�Ҫ����������
                        End If
                    End If
                End If
            End If
            mblnReturn = False
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsDiagZY_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    With vsDiagZY
        If Col = col��Ժ��� Then
            .TextMatrix(Row, Col) = NeedName(.TextMatrix(Row, Col))
            .Tag = ""
            mblnChange = True
        End If
        If Col = col������� Then
            ' .EditText = "" �ų���Ԫ�������ݲ����س���״��
            If .EditText = "" And .Cell(flexcpData, Row, Col) <> "" Then
                '�ڵ���vsDiagZY_KeyDown(vbKeyDelete, 0)���ǿ���ɾ����ǰ�У������ָ�ԭʼ����
                .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
                Call vsDiagZY_KeyDown(vbKeyDelete, 0)
            End If
        End If
        Call vsDiagZY_AfterRowColChange(-1, -1, .Row, .Col)
    End With
    If vsDiagZY.Tag = "δ�޸�" Then vsDiagZY.Tag = "": mblnChange = True
End Sub

Private Sub vsDiagZY_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim i As Long
    
    With vsDiagZY
        '���ͼƬ
        For i = .FixedRows To .Rows - 1
            If Not .Cell(flexcpPicture, i, colzy����) Is Nothing Then
                Set .Cell(flexcpPicture, i, colzy����) = Nothing
            End If
            If Not .Cell(flexcpPicture, i, colzyDel) Is Nothing Then
               Set .Cell(flexcpPicture, i, colzyDel) = Nothing
            End If
        Next
        
        If Not ZYCellEditable(NewRow, NewCol) Then
            .ComboList = ""
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
            Set .CellButtonPicture = Nothing
            
            If NewCol = col������� Then
                .ComboList = "..."
            ElseIf NewCol = col��ҽ֤�� Then
                If .TextMatrix(NewRow, col�������) = "" Then
                    .ComboList = ""
                    .FocusRect = flexFocusLight
                Else
                    .ComboList = "..."
                End If
            ElseIf NewCol = col��Ժ��� Then
                .ComboList = .ColData(NewCol)
            ElseIf NewCol = col��Ժ���� Then
                If .TextMatrix(NewRow, colzy����) = "13" Then
                    .ComboList = "��|�ٴ�δȷ��|�������|��"
                Else
                    .ComboList = ""
                    .FocusRect = flexFocusLight
                End If
            ElseIf NewCol = colzy���� Then
                .ComboList = "..."
                .FocusRect = flexFocusNone
                Set .CellButtonPicture = imgButtonNew.Picture
            ElseIf NewCol = colzyDel Then
                .ComboList = "..."
                .FocusRect = flexFocusNone
                Set .CellButtonPicture = imgButtonDel.Picture
            Else
                .ComboList = ""
            End If
        End If
        If NewRow >= .FixedRows Then
            '��ʾͼƬ
            If NewCol <> colzy���� And .TextMatrix(NewRow, col�������) <> "" And .TextMatrix(NewRow, 0) <> "��Ҫ���" Then
                Set .Cell(flexcpPicture, NewRow, colzy����) = imgButtonNew.Picture
            End If
            '��ʾͼƬ
            If NewCol <> colzyDel Then
                Set .Cell(flexcpPicture, NewRow, colzyDel) = imgButtonDel.Picture
            End If
        End If
    End With
End Sub

Private Sub vsDiagZY_BeforeUserResize(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = colzy���� Then Cancel = True
End Sub

Private Sub vsDiagZY_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String
    Dim str�Ա� As String, lngRow As Long
    Dim blnCancle As Boolean
    
    With vsDiagZY
        If Col = col������� Then
            If optInput(2).Value Then
                '���������:��ҽ���ݣ�һ����Ͽ������ڶ������
                Set rsTmp = zlDatabase.ShowILLSelect(Me, "2", mlng����ID, , True, False)
            Else
                'B-��ҽ��������
                Set rsTmp = zlDatabase.ShowILLSelect(Me, "B", mlng����ID, cboinfo(cbo�Ա�).Text, True)
            End If
            If Not rsTmp Is Nothing Then
                Call ZYSetDiagInput(Row, rsTmp)
                Call ZYEnterNextCell
            End If
        ElseIf Col = col��ҽ֤�� Then
            If optInput(2).Value Then
                '���������:�Ȳ��Ƿ��ж�Ӧ
                If Not Set��ҽ֤��(Row, Val(.TextMatrix(Row, colzy���ID))) Then
                    Set rsTmp = zlDatabase.ShowILLSelect(Me, "Z", mlng����ID, cboinfo(cbo�Ա�).Text, True)
                Else
                    Exit Sub
                End If
            Else
                'Z-��ҽ��������
                Set rsTmp = zlDatabase.ShowILLSelect(Me, "Z", mlng����ID, cboinfo(cbo�Ա�).Text, True)
            End If
            If Not rsTmp Is Nothing Then
                Call Set��ҽ֤��(Row, 0, rsTmp)
                Call ZYEnterNextCell
            End If
        ElseIf Col = colzy���� Then
            lngRow = Row + 1: .AddItem "", lngRow
            .TextMatrix(lngRow, colzy����) = .TextMatrix(Row, colzy����)
            .Cell(flexcpBackColor, lngRow, col��ϱ���) = ColorUnEditCell      '����ɫ
            
'            .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = .BackColor
'            .Cell(flexcpBackColor, GetRow(13), .FixedRows, GetRow(13), .Cols - 1) = &HC0FFC0
            
            .Row = lngRow: .Col = col�������
            .ShowCell .Row, .Col
        ElseIf Col = colzyDel Then
            Call vsDiagZY_KeyDown(vbKeyDelete, 0)
        End If
    End With
End Sub

Private Sub vsDiagZY_Click()
    With vsDiagZY
        If (.MouseCol = colzy���� Or .MouseCol = colzyDel) And .MouseRow >= .FixedRows Then
            If .MouseCol = colzy���� Then
                If .TextMatrix(.MouseRow, col�������) = "" Or .TextMatrix(.MouseRow, 0) = "��Ҫ���" Then Exit Sub
            End If
        
            .Select .MouseRow, .MouseCol
            Call vsDiagZY_CellButtonClick(.MouseRow, .MouseCol)
        End If
    End With
End Sub

Private Sub vsDiagZY_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    With vsDiagZY
        If Col = col��Ժ��� Then
            '��λ��ƥ����
            For i = 0 To .ComboCount - 1
                If NeedName(.ComboItem(i)) = .TextMatrix(Row, Col) Then
                    .ComboIndex = i: Exit For
                End If
            Next
        End If
    End With
End Sub

Private Sub vsDiagZY_DblClick()
    Call vsDiagZY_KeyPress(32)
End Sub

Private Sub vsDiagZY_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long, j As Long
    
    If mbln��ʿվ Or mblnReadOnly Then Exit Sub
    
    With vsDiagZY
        If KeyCode = vbKeyF4 Then
            If .Col = col������� Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, col�������) <> "" Then
                If mlngPathState = 1 Or mlngPathState = 2 Then
                    If .TextMatrix(.Row, col�������) = "��Ժ���" And mlngDiagnosisType = 12 Or .TextMatrix(.Row, col�������) = "�������" And mlngDiagnosisType = 11 Then
                        If .TextMatrix(.Row, col�������) <> .TextMatrix(.Row - 1, col�������) Then
                            '��Ҫ��ϲ������
                            Exit Sub
                        End If
                    End If
                End If
                '�ϲ�·��
                If Not CheckMergePath(mlng����ID, mlng��ҳID, Val(.TextMatrix(.Row, colzy����)), Val(.TextMatrix(.Row, colzy����ID))) Then Exit Sub
                '����·������
                If mstrPathDiag <> "" And mlngPathState > 0 Then
                    If InStr("," & mstrPathDiag & ",", "," & .TextMatrix(.Row, colzy����) & "|" & Val(.TextMatrix(.Row, col����ID)) & "|" & Val(.TextMatrix(.Row, col���ID)) & ",") > 0 Then
                        '������ϲ������
                        Exit Sub
                    End If
                End If
                If mlngPathState = 2 And mblnIsPathOutTime Then
                    If .TextMatrix(.Row, col�������) = "��Ҫ���" And mlngDiagnosisType > 10 Then
                        '������ɵĳ�Ժ��ϲ������
                        Exit Sub
                    End If
                End If
                If MsgBox("ȷʵҪ������������Ϣ��", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    i = Val(.TextMatrix(.Row, colzy����))
                    .Cell(flexcpText, .Row, .FixedRows, .Row, .Cols - 1) = ""
                    .Cell(flexcpData, .Row, .FixedRows, .Row, .Cols - 1) = Empty
                    .TextMatrix(.Row, colzy����) = i
                    
                    '�����ͬ�������������
                    If .TextMatrix(.Row, col�������) = "" Then
                        .RemoveItem .Row
                    Else
                        If .Row + 1 <= .Rows - 1 Then
                            If .TextMatrix(.Row + 1, col�������) = "" Then
                                '��һ��Ϊ�ޱ����������ʱ�����ݲ����ƣ�����ǰ��Ϊ�б���ʱֻ�����
                                For i = .Row + 1 To .Rows - 1
                                    If Val(.TextMatrix(i, colzy����)) = Val(.TextMatrix(.Row, colzy����)) Then
                                        For j = .FixedCols To .Cols - 1
                                            .TextMatrix(i - 1, j) = .TextMatrix(i, j)
                                            .Cell(flexcpData, i - 1, j) = .Cell(flexcpData, i, j)
                                        Next
                                        .Cell(flexcpText, i, .FixedCols, i, .Cols - 1) = ""
                                        .Cell(flexcpData, i, .FixedCols, i, .Cols - 1) = Empty
                                        .TextMatrix(i, colzy����) = Val(.TextMatrix(.Row, colzy����))
                                        
                                        If i = .Rows - 1 Then
                                            If .TextMatrix(i, col�������) = "" Then .RemoveItem i
                                            Exit For
                                        ElseIf Val(.TextMatrix(i + 1, colzy����)) <> Val(.TextMatrix(i, colzy����)) Then
                                            If .TextMatrix(i, col�������) = "" Then .RemoveItem i
                                            Exit For
                                        End If
                                    End If
                                Next
                            End If
                        End If
                    End If
                    
'                    .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = .BackColor
'                    .Cell(flexcpBackColor, GetRow(13), .FixedRows, GetRow(13), .Cols - 1) = &HC0FFC0
                    
                    '������Ϸ������
                    Call Set��ҽ������(.Row)

                    mblnChange = True
                    .Tag = ""
                End If
            ElseIf .TextMatrix(.Row, col�������) = "" Then
                .RemoveItem .Row
            End If
        ElseIf KeyCode > 127 Then
            '���ֱ�����뺺�ֵ�����
            Call vsDiagZY_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsDiagZY_KeyPress(KeyAscii As Integer)
    If mbln��ʿվ Or mblnReadOnly Then Exit Sub
    
    With vsDiagZY
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call ZYEnterNextCell
        Else
            If .Col = col������� Or .Col = col��ҽ֤�� Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsDiagZY_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" 'ʹ��ť״̬��������״̬
                End If
            End If
        End If
    End With
End Sub

Private Sub vsDiagZY_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        mblnReturn = True
        
        With vsDiagZY
            If Col = col��Ժ��� Then
                KeyAscii = 0
                
                '��ʱ.TextMatrix��δ����,����ȡComboItem
                .TextMatrix(Row, Col) = NeedName(.ComboItem(.ComboIndex))
                mblnChange = True
                .Tag = ""
                Call ZYEnterNextCell
            End If
        End With
    Else
        mblnReturn = False
    End If
End Sub

Private Sub vsDiagZY_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsDiagZY.EditSelStart = 0
    vsDiagZY.EditSelLength = zlCommFun.ActualLen(vsDiagZY.EditText)
End Sub

Private Sub vsDiagZY_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not ZYCellEditable(Row, Col) Then
        Cancel = True
    End If
End Sub

Private Sub vsDiagZY_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String, blnCancel As Boolean
    Dim strInput As String, vPoint As POINTAPI
    Dim str�Ա� As String, int������� As Integer
    
    With vsDiagZY
        If Col = col������� Or Col = col��ҽ֤�� Then
            '.Cell(flexcpData, Row, Col) <> ""�ų����лس�
            If .EditText = "" And .Cell(flexcpData, Row, Col) <> "" Then
                .EditText = ""
                '��ҽ֢���������������
                If Col = col��ҽ֤�� Then
                    .Cell(flexcpData, Row, Col) = ""
                End If
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call ZYEnterNextCell
            ElseIf Col = col������� And .TextMatrix(Row, col��ϱ���) <> "" And .Cell(flexcpData, Row, Col) <> "" And .EditText Like "*" & .Cell(flexcpData, Row, Col) & "*" Then
                strInput = UCase(.EditText)
                StrSQL = GetSQL(1, strInput, str�Ա�)
                On Error GoTo errH
                Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, strInput, strInput, str�Ա�, mint���� + 1)
                If rsTmp.RecordCount = 1 Then
                    Call ZYSetDiagInput(Row, rsTmp):
                    .EditText = .Text
                Else
                    '�����ڱ�׼������ǰ�����븽����Ϣ
                    .TextMatrix(Row, col�������) = .EditText
                End If
                '������.Cell(flexcpData, Row, Col)���Ա��޸�����ʱ�ٴ�ʹ��like�ж�
                .Tag = ""
                mblnChange = True
            Else
                If Val(.TextMatrix(Row, colzy����)) = 11 Then
                    int������� = Val(Mid(gstr�������, 1, 1))
                Else
                    int������� = Val(Mid(gstr�������, 2, 1))
                End If
                If int������� = 0 Then int������� = 1
                
                strInput = UCase(.EditText)
                StrSQL = GetSQL(1, strInput, str�Ա�, IIf(Col = col�������, "B", "Z"))
                If Col = col������� Then
                    If int������� = 1 And zlCommFun.IsCharChinese(strInput) Then
                        Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, strInput & "%", mstrLike & strInput & "%", str�Ա�, mint���� + 1)
                        If rsTmp.EOF Then
                            Set rsTmp = Nothing
                        ElseIf rsTmp.RecordCount > 1 Then
                            Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, strInput, strInput, str�Ա�, mint���� + 1)
                            If rsTmp.RecordCount <> 1 Then Set rsTmp = Nothing '����¼��ʱ�ж��ƥ�䲻����ѡ��
                        End If
                        Call ZYSetDiagInput(Row, rsTmp): .EditText = .Text
                        If mblnReturn And rsTmp Is Nothing Then Call ZYEnterNextCell '��������¼��ʱ���ݲ�������һ�У���Ϊ���ܻ�Ҫ����������
                    Else
                        vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                        Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 0, IIf(optInput(2).Value, "�������", "��������"), False, "", "", False, False, True, _
                            vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%", str�Ա�, mint���� + 1)
                        If blnCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                            Cancel = True
                        Else
                            '���������뷽ʽ
                            If rsTmp Is Nothing And ((int������� = 2 Or int������� = 3 And mint���� <> 0)) Then
                                MsgBox "û���ҵ�������ƥ������ݡ�", vbInformation, gstrSysName
                                Cancel = True
                            Else
                                Call ZYSetDiagInput(Row, rsTmp): .EditText = .Text
                                'If mblnReturn Then Call ZYEnterNextCell '�ݲ�������һ�У���Ϊ���ܻ�Ҫ����������
                            End If
                        End If
                    End If
                ElseIf Col = col��ҽ֤�� Then
                    If optInput(2).Value Then
                        '���������:�Ȳ��Ƿ��ж�Ӧ
                        If Set��ҽ֤��(Row, Val(.TextMatrix(Row, colzy���ID))) Then
                            mblnReturn = False
                            Exit Sub
                        End If
                    End If
                    vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 0, "��ҽ֤��", False, "", "", False, False, True, _
                        vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%", str�Ա�, mint���� + 1)
                    If blnCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                        Cancel = True
                    Else
                        '���������뷽ʽ
                        If rsTmp Is Nothing Then
                            MsgBox "û���ҵ�������ƥ������ݡ�", vbInformation, gstrSysName
                            Cancel = True
                        Else
                            Call Set��ҽ֤��(Row, 0, rsTmp)
                        End If
                    End If
                End If
            End If
            mblnReturn = False
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub VsGriedFocuesMove(ByVal vsBill As Object, ByVal lngRow As Long, ByVal lngCol As Long, ByVal KeyCode As Integer, _
        Optional lngFiexCol As Long = 0, Optional lngFiexCol1 As Long = -1)
    '------------------------------------------------------------------------------------------------------------
    '����:��һ�������ƶ���Ԫ��
    '����:vsBill-���ؼ�
    '       lngRow-��ǰ��
    '       lngCol-��ǰ��
    '       KeyCode-����
    '       lngFiexCol-�ж��Ƿ��Ƶ�������еĹ̶���
    '       lngFiexCol1-�ж��Ƿ��Ƶ�������еĹ̶���(��ͬʱҪ����lngFiexCol��)
    '����:���˺�
    '����:2007/05/18
    '------------------------------------------------------------------------------------------------------------
    If KeyCode <> vbKeyReturn Then Exit Sub
    Dim strCurrValue As String
    If lngCol = lngFiexCol Then
        strCurrValue = vsBill.EditText
    Else
        strCurrValue = ""
    End If
    
    With vsBill
        
        Select Case lngCol
        Case 0
            If Trim(.TextMatrix(lngRow, lngFiexCol)) = "" And strCurrValue = "" Then
                zlCommFun.PressKey vbKeyTab
                Exit Sub
            End If
            .Col = lngCol + 1
            GoTo ShowCell:
        Case Else
            If lngCol >= .Cols - 1 Then
                If lngRow < .Rows - 1 Then
                    .Row = lngRow + 1
                    .Col = 0
                    GoTo ShowCell:
                    Exit Sub
                End If
                If Trim(.TextMatrix(lngRow, lngFiexCol)) <> "" Then
                    If lngFiexCol1 > 0 Then
                        If Trim(.TextMatrix(lngRow, lngFiexCol1)) <> "" Then
                            .Rows = .Rows + 1
                            .Row = .Rows - 1
                            .Col = 0
                        End If
                    Else
                        .Rows = .Rows + 1
                        .Row = .Rows - 1
                        .Col = 0
                    End If
                End If
                GoTo ShowCell:
                Exit Sub
            End If
            .Col = lngCol + 1
         End Select
ShowCell:
        .ShowCell .Row, .Col
    End With
End Sub


Private Function CheckInPutIsDate(ByVal vsObj As Object, lngRow As Long, lngCol As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------
    '����:���������������Ƿ�Ϸ�
    '����:lngRow -��,lngCol -��
    '����:���ںϷ�,����true,���򷵻�False
    '����:���˺�
    '����:2007/05/21
    '---------------------------------------------------------------------------------------------------------
    Dim strKEY As String
    Dim str����ʱ�� As String, str�˳�ʱ�� As String
    Dim str��Ժʱ�� As String
    str��Ժʱ�� = txtInfo(txt��Ժʱ��).Text
    
        
    strKEY = Trim(vsObj.EditText)
    strKEY = Replace(strKEY, Chr(vbKeyReturn), "")
    strKEY = Replace(strKEY, Chr(10), "")
    If strKEY <> "" Then
        
        If Not IsDate(strKEY) Then
            MsgBox vsObj.TextMatrix(0, lngCol) & "����Ϊ������,���������룡", vbInformation + vbDefaultButton1, Me.Caption
             vsObj.EditSelStart = 0
             vsObj.EditSelLength = 1000
            Exit Function
        End If
        Select Case lngCol
        Case 1
            str����ʱ�� = strKEY
            str�˳�ʱ�� = Trim(vsObj.TextMatrix(lngRow, 2))
            If str�˳�ʱ�� <> "" And str����ʱ�� > str�˳�ʱ�� Then
                MsgBox "ע:" & vbCrLf & "  ����ʱ��������˳�ʱ��,���飡", vbInformation + vbDefaultButton1, Me.Caption
                Exit Function
            End If
        Case Else
            str����ʱ�� = Trim(vsObj.TextMatrix(lngRow, 1))
            str�˳�ʱ�� = strKEY
 
            If str����ʱ�� <> "" And CDate(str����ʱ��) >= CDate(str�˳�ʱ��) Then
                MsgBox "ע:" & vbCrLf & "  �˳�ʱ��С���˽���ʱ��,���飡", vbInformation + vbDefaultButton1, Me.Caption
                Exit Function
            End If
        End Select
    End If
    CheckInPutIsDate = True
End Function

Private Sub vsfMain_EnterCell()
    Select Case vsfMain.Col
        Case 0, 3, 6
            vsfMain.Editable = flexEDNone
        Case 1, 4
            If vsfMain.BackColor = vbButtonFace Then Exit Sub
            If InStr(vsfMain.TextMatrix(vsfMain.Row, vsfMain.Col + 1), ",") > 0 Then
                vsfMain.ColComboList(vsfMain.Col) = Replace(vsfMain.TextMatrix(vsfMain.Row, vsfMain.Col + 1), ",", "|")
            Else
                vsfMain.ColComboList(vsfMain.Col) = ""
            End If
            If vsfMain.TextMatrix(vsfMain.Row, vsfMain.Col - 1) <> "" Then
                vsfMain.Editable = flexEDKbdMouse
            Else
                vsfMain.Editable = flexEDNone
            End If
    End Select
End Sub
Private Sub vsfMain_KeyPress(KeyAscii As Integer)
    If vsfMain.Rows <= 1 Then Exit Sub
    If mbln��ʿվ Or mblnReadOnly Then Exit Sub
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Select Case vsfMain.Col
            Case 0, 3, 6
                vsfMain.Col = vsfMain.Col + 1
            Case 1, 4
                If vsfMain.Col = 4 And vsfMain.Row <> vsfMain.Rows - 1 Then
                    vsfMain.Col = 0
                    vsfMain.Row = vsfMain.Row + 1
                ElseIf vsfMain.Col = 4 And vsfMain.Row = vsfMain.Rows - 1 Then
                    Call zlCommFun.PressKey(vbKeyTab)
                Else
                    vsfMain.Col = vsfMain.Col + 3
                End If
        End Select
        vsfMain.ShowCell vsfMain.Row, vsfMain.Col
    End If
End Sub

Private Sub vsfMain_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim sngNum1, sngNum2 As Single
    If InStr(vsfMain.TextMatrix(Row, Col + 1), "...") > 0 Then
        sngNum1 = Mid(vsfMain.TextMatrix(Row, Col + 1), 1, InStr(vsfMain.TextMatrix(Row, Col + 1), "...") - 1)
        sngNum2 = Mid(vsfMain.TextMatrix(Row, Col + 1), InStr(vsfMain.TextMatrix(Row, Col + 1), "...") + 3)
        If Not IsNumeric(vsfMain.EditText) Then
            Cancel = True
        ElseIf CSng(vsfMain.EditText) < sngNum1 Or CSng(vsfMain.EditText) > sngNum2 Then
            MsgBox "����Ӧ����" & vsfMain.TextMatrix(Row, Col + 1) & "�ķ�Χ����!", vbInformation, gstrSysName
            Cancel = True
        End If
    ElseIf InStr(vsfMain.TextMatrix(Row, Col + 1), "-") > 0 Then
        If InStr(vsfMain.TextMatrix(Row, Col + 1), "-") = 1 Then
            sngNum1 = Mid(vsfMain.TextMatrix(Row, Col + 1), 2, InStr(2, vsfMain.TextMatrix(Row, Col + 1), "-") - 1)
            sngNum2 = Mid(vsfMain.TextMatrix(Row, Col + 1), InStr(2, vsfMain.TextMatrix(Row, Col + 1), "-") + 1)
        Else
            sngNum1 = Mid(vsfMain.TextMatrix(Row, Col + 1), 1, InStr(1, vsfMain.TextMatrix(Row, Col + 1), "-") - 1)
            sngNum2 = Mid(vsfMain.TextMatrix(Row, Col + 1), InStr(1, vsfMain.TextMatrix(Row, Col + 1), "-") + 1)
        End If
        If Not IsNumeric(vsfMain.EditText) Then
            Cancel = True
        ElseIf CSng(vsfMain.EditText) < sngNum1 Or CSng(vsfMain.EditText) > sngNum2 Then
            MsgBox "����Ӧ����" & vsfMain.TextMatrix(Row, Col + 1) & "�ķ�Χ����!", vbInformation, gstrSysName
            Cancel = True
        End If
    ElseIf vsfMain.TextMatrix(Row, Col + 1) = "" Then
        If zlCommFun.ActualLen(vsfMain.EditText) > mlngSize Then
            MsgBox "���볤�Ȳ��ܴ���" & "[" & mlngSize & "]", vbInformation, gstrSysName
            Cancel = True
        End If
    End If
    If Cancel = False Then mblnChange = True: vsfMain.Tag = ""
    
End Sub

Private Sub vsKSS_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call vsKSS_AfterRowColChange(-1, -1, vsKSS.Row, vsKSS.Col)
End Sub

Private Sub vsKSS_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    vsKSS.ColComboList(kss����) = "..."
    vsKSS.ColComboList(kssʹ�ý׶�) = " |��ǰ|����|����|Χ������"
    vsKSS.ColComboList(KSS������ҩ) = "����|����|����|����|>����"
    vsKSS.ColComboList(kss��ҩĿ��) = " |Ԥ��|����|Ԥ��������"
End Sub

Private Sub vsKSS_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String, blnCancel As Boolean
    Dim strSQLItem As String
    
    With vsKSS
        If Col = kss���� Then
            strSQLItem = _
                " From ������ĿĿ¼ A,ҩƷ���� B" & _
                " Where A.ID=B.ҩ��ID And A.���='5' And A.������� IN(2,3) And Nvl(b.������, 0) <> 0" & _
                " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
                " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null) "
            StrSQL = "Select 0 as ĩ��,Max(Level) as ��ID,ID,�ϼ�ID,����,����,NULL as ��λ" & _
                " From ���Ʒ���Ŀ¼ Where ����=1 And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " Start With ID In (Select A.����ID" & strSQLItem & ") Connect by Prior �ϼ�ID=ID" & _
                " Group by ID,�ϼ�ID,����,����"
            StrSQL = StrSQL & " Union ALL" & _
                " Select 1 as ĩ��,1 as ��ID,A.ID,����ID as �ϼ�ID,A.����,A.����,A.���㵥λ as ��λ" & _
                strSQLItem & " Order By ĩ��,��ID Desc,����"
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 2, "����ҩ��", False, "", "", False, True, False, 0, 0, 0, blnCancel, False, False)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "û�п���ҩ�����ݿ���ѡ��", vbInformation, gstrSysName
                End If
            Else
                Call KSSSetDiagInput(Row, rsTmp)
                Call KSSEnterNextCell
            End If
        End If
    End With
End Sub

Private Sub vsKSS_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long, j As Long
    
    If mbln��ʿվ Or mblnReadOnly Then Exit Sub
    
    If KeyCode = vbKeyF4 Then
        Call zlCommFun.PressKey(vbKeySpace)
    ElseIf KeyCode = vbKeyDelete Then
        If MsgBox("ȷʵҪɾ������������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            With vsKSS
                '�ж��Ƿ����޸�
                If .RowData(.Row) & "" <> "" Then
                    If InStr(mstrDelete, .RowData(.Row) & "") <= 0 Then
                        mstrDelete = mstrDelete & IIf(mstrDelete <> "", ",", "") & .RowData(.Row)
                    End If
                End If
                .RemoveItem .Row
                If .Rows < 4 Then .Rows = 4
                Call SetKSSSerial
                vsKSS.Tag = ""
            End With
            mblnChange = True
        End If
    ElseIf KeyCode > 127 Then
        '���ֱ�����뺺�ֵ�����
        Call vsKSS_KeyPress(KeyCode)
    End If
End Sub

Private Sub vsKSS_KeyPress(KeyAscii As Integer)
    If mbln��ʿվ Or mblnReadOnly Then Exit Sub
    
    With vsKSS
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call KSSEnterNextCell
        Else
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                Call vsKSS_CellButtonClick(.Row, .Col)
            Else
                .ColComboList(kss����) = "" 'ʹ��ť״̬��������״̬
            End If
        End If
    End With
End Sub

Private Sub vsKSS_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        mblnReturn = True
    Else
        mblnReturn = False
    End If
    If Col = kssʹ������ And Len(vsKSS.EditText) > 18 And KeyAscii <> vbKeyBack And vsKSS.EditSelLength = 0 Then KeyAscii = 0
    If Col = kss��ҩĿ�� And LenB(StrConv(vsKSS.EditText, vbFromUnicode)) >= 200 And KeyAscii <> vbKeyBack And vsKSS.EditSelLength = 0 Then KeyAscii = 0
End Sub

Private Sub vsKSS_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsKSS.EditSelStart = 0
    vsKSS.EditSelLength = zlCommFun.ActualLen(vsKSS.EditText)
End Sub

Private Sub vsKSS_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String, blnCancel As Boolean
    Dim strInput As String, vPoint As POINTAPI
    
    With vsKSS
        If Col = kss���� Then
            If .EditText = "" Then
                .EditText = .Cell(flexcpData, Row, Col)
                If mblnReturn Then Call KSSEnterNextCell
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call KSSEnterNextCell
            Else
                strInput = UCase(.EditText)
                If zlCommFun.IsCharChinese(strInput) Then
                    StrSQL = "B.���� Like [2]" '���뺺��ʱֻƥ������
                Else
                    StrSQL = "A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2]"
                End If
                StrSQL = _
                    " Select Distinct A.ID,A.����,A.����,A.���㵥λ as ��λ" & _
                    " From ������ĿĿ¼ A,������Ŀ���� B,ҩƷ���� C" & _
                    " Where A.ID=B.������ĿID And A.ID=C.ҩ��ID And Nvl(c.������, 0) <> 0" & _
                    " And (A.����ʱ�� Is Null Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                    " And A.���='5' And A.������� IN(2,3) And B.����=[3] And (" & StrSQL & ")" & _
                    " Order by A.����"
                If zlCommFun.IsCharChinese(strInput) Then
                    On Error GoTo errH
                    Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, strInput & "%", mstrLike & strInput & "%", mint���� + 1)
                    '�ж��Ƿ�������
                    If rsTmp.RecordCount = 0 Then
                        MsgBox "û���ҵ�ָ���Ŀ���ҩ�", vbInformation, gstrSysName
                        Cancel = True: .EditText = "": Exit Sub
                    End If
                    If rsTmp.EOF Then
                        Set rsTmp = Nothing
                    ElseIf rsTmp.RecordCount > 1 Then
                        Set rsTmp = Nothing '����¼��ʱ�ж��ƥ�䲻����ѡ��
                    End If
                    Call KSSSetDiagInput(Row, rsTmp)
                    .EditText = .Text
                    If mblnReturn Then Call KSSEnterNextCell
                Else
                    vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                    Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 0, "����ҩ��", _
                        False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                        strInput & "%", mstrLike & strInput & "%", mint���� + 1)
                    If blnCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                        Cancel = True
                    Else
                        '�ж��Ƿ�������
                        If rsTmp Is Nothing Then
                            MsgBox "û���ҵ�ָ���Ŀ���ҩ�", vbInformation, gstrSysName
                            Cancel = True: .EditText = "": Exit Sub
                        End If
                        Call KSSSetDiagInput(Row, rsTmp)
                        .EditText = .Text
                        If mblnReturn Then Call KSSEnterNextCell
                    End If
                End If
            End If
            mblnReturn = False
        ElseIf Col = kssʹ������ Or Col = KSSDDD�� Then
            If (Not IsNumeric(.EditText) Or InStr(.EditText, "-") > 0 Or InStr(.EditText, "+") > 0) And .EditText <> "" Then
                MsgBox "��������Ч�����֡�", vbInformation, Me.Caption
                Cancel = True
            Else
                If Len(.EditText) > 12 Then
                    MsgBox "������12λ���µ����֡�", vbInformation, Me.Caption
                    Cancel = True
                    Exit Sub
                End If
                If .TextMatrix(Row, Col) <> .EditText Then .Tag = "": mblnChange = True
            End If
        Else
            '����û��޸��ˣ�����ȡ��ʱ��Ӱ����һ��
            If .Cell(flexcpData, Row, Col) = "����" Then .Cell(flexcpData, Row, Col) = ""
            If .TextMatrix(Row, Col) <> .EditText Then .Tag = "": mblnChange = True
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsOPS_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim strInput As String
    
    With vsOPS
        If Col = col�������� Then
            strInput = GetFullDate(.TextMatrix(Row, Col), False)
            If Not IsDate(strInput) Then
                .TextMatrix(Row, Col) = .Cell(flexcpData, Row, Col)
            Else
                .TextMatrix(Row, Col) = strInput
                .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                mblnChange = True
                .Tag = ""
            End If
        ElseIf Col = col�п����� Then
            .TextMatrix(Row, Col) = NeedName(.TextMatrix(Row, Col))
            mblnChange = True
            .Tag = ""
        ElseIf Col = col�������� Or Col = col�������� Then
            mblnChange = True
            .Tag = ""
        End If
    End With
    Call vsOPS_AfterRowColChange(-1, -1, vsOPS.Row, vsOPS.Col)
End Sub

Private Sub vsOPS_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    With vsOPS
        If Not OPSCellEditable(NewRow, NewCol) Then
            .ComboList = ""
            .FocusRect = flexFocusLight
        Else
            .FocusRect = flexFocusSolid
            If NewCol = col�������� Or NewCol = col����ҽʦ _
                Or NewCol = col������ʿ Or NewCol = col����1 Or NewCol = col����2 _
                Or NewCol = col����ʽ Or NewCol = col����ҽʦ Or (NewCol = col�������� And chkInfo(chk��������¼��).Value) Then
                .ComboList = "..."
            ElseIf NewCol = col�п����� Then
                .ComboList = .ColData(NewCol)
            Else
                .ComboList = ""
            End If
        End If
    End With
End Sub

Private Sub vsOPS_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String, blnCancel As Boolean
    Dim str�Ա� As String, int�Ա� As Integer
    Dim vPoint As POINTAPI
    
    With vsOPS
        If Col = col�������� Or Col = col�������� Then
            If optInput(4).Value Then
                '��������Ŀ����
                If cboinfo(cbo�Ա�).Text Like "*��*" Then
                    int�Ա� = 1
                ElseIf cboinfo(cbo�Ա�).Text Like "*Ů*" Then
                    int�Ա� = 2
                End If
                            
                StrSQL = "Select 0 as ĩ��,ID,�ϼ�ID,����,����,NULL as ��ģ" & _
                    " From ���Ʒ���Ŀ¼ Where ����=5 And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                    " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID" & _
                    " Union ALL " & _
                    " Select 1 as ĩ��,ID,����ID as �ϼ�ID,����,����,�������� as ��ģ" & _
                    " From ������ĿĿ¼" & _
                    " Where ���='F' And ������� IN(2,3) And (վ��='" & gstrNodeNo & "' Or վ�� is Null)" & _
                    IIf(int�Ա� <> 0, " And Nvl(�����Ա�,0) IN(0,[2])", "") & _
                    " And (����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or ����ʱ�� is NULL)"
            Else
                '��ICD9-CM3����
                If cboinfo(cbo�Ա�).Text Like "*��*" Then
                    str�Ա� = "��"
                ElseIf cboinfo(cbo�Ա�).Text Like "*Ů*" Then
                    str�Ա� = "Ů"
                End If
                StrSQL = _
                    " Select 0 as ĩ��,ID,�ϼ�ID," & _
                    " ���||LPAD(���,3,'0') as ����," & _
                    " NULL as ����,����,����,NULL as ˵��" & _
                    " From ����������� Where ���='S'" & _
                    " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID" & _
                    " Union ALL " & _
                    " Select 1 as ĩ��,ID,����ID as �ϼ�ID,����,����,����,����,˵��" & _
                    " From ��������Ŀ¼ Where ���='S'" & _
                    IIf(str�Ա� <> "", " And (�Ա�����=[1] Or �Ա����� is NULL)", "") & _
                    " And (����ʱ�� is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))"
            End If
            Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 2, IIf(optInput(4).Value, "������Ŀ", "��������"), _
                False, "", "", False, True, False, 0, 0, 0, blnCancel, False, False, str�Ա�, int�Ա�)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "û��" & IIf(optInput(4).Value, "������Ŀ", "��������") & "����ѡ��", vbInformation, gstrSysName
                End If
            Else
                Call OPSSetInput(Row, Col, rsTmp)
                Call OPSEnterNextCell
            End If
        ElseIf Col = col����ʽ Then
            StrSQL = "Select 0 as ĩ��,ID,�ϼ�ID,����,����,NULL as ��������" & _
                " From ���Ʒ���Ŀ¼ Where ����=5 And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " Start With �ϼ�ID is NULL Connect by Prior ID=�ϼ�ID" & _
                " Union ALL " & _
                " Select 1 as ĩ��,ID,����ID as �ϼ�ID,����,����,�������� as ��������" & _
                " From ������ĿĿ¼ Where ���='G'" & _
                " And (����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or ����ʱ�� is NULL)" & _
                " And (վ��='" & gstrNodeNo & "' Or վ�� is Null)"
            Set rsTmp = zlDatabase.ShowSelect(Me, StrSQL, 2, "������Ŀ", , , , , True, , , , , blnCancel)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "û��������Ŀ����ѡ��", vbInformation, gstrSysName
                End If
            Else
                Call OPSSetInput(Row, Col, rsTmp)
                Call OPSEnterNextCell
            End If
        ElseIf Col = col����ҽʦ Or Col = col����1 Or Col = col����2 Or Col = col����ҽʦ Then
            StrSQL = "Select A.ID,A.���,A.����,A.����" & _
                " From ��Ա�� A,��Ա����˵�� B" & _
                " Where A.ID=B.��ԱID And B.��Ա����='ҽ��'" & _
                " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
                " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                " Order by A.���"
            vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
            Set rsTmp = zlDatabase.ShowSelect(Me, StrSQL, 0, "ҽ��", , , , , , True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, , True)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "û��ҽ������ѡ��", vbInformation, gstrSysName
                End If
            Else
                Call OPSSetInput(Row, Col, rsTmp)
                Call OPSEnterNextCell
            End If
        ElseIf Col = col������ʿ Then
            StrSQL = "Select A.ID,A.���,A.����,A.����" & _
                " From ��Ա�� A,��Ա����˵�� B" & _
                " Where A.ID=B.��ԱID And B.��Ա����='��ʿ'" & _
                " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
                " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                " Order by A.���"
            vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
            Set rsTmp = zlDatabase.ShowSelect(Me, StrSQL, 0, "��ʿ", , , , , , True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, , True)
            If rsTmp Is Nothing Then
                If Not blnCancel Then
                    MsgBox "û�л�ʿ����ѡ��", vbInformation, gstrSysName
                End If
            Else
                Call OPSSetInput(Row, Col, rsTmp)
                Call OPSEnterNextCell
            End If
        End If
    End With
End Sub

Private Sub SetAllerInput(ByVal lngRow As Long, rsInput As ADODB.Recordset)
'���ܣ��������ҩ�������
    Dim StrSQL As String, curDate As Date
    
    With vsAller
        If Not rsInput Is Nothing Then
            .RowData(lngRow) = CLng(rsInput!ID)
            .TextMatrix(lngRow, AC_����ҩ��) = Nvl(rsInput!����)
        Else
            .RowData(lngRow) = 0
            .TextMatrix(lngRow, AC_����ҩ��) = .EditText
        End If
        .Cell(flexcpData, lngRow, AC_����ҩ��) = .TextMatrix(lngRow, AC_����ҩ��)
        
        If .Cell(flexcpData, lngRow, AC_����ʱ��) = "" Then
            curDate = zlDatabase.Currentdate
            .TextMatrix(lngRow, AC_����ʱ��) = Format(curDate, "yyyy-MM-dd HH:mm")
            .Cell(flexcpData, lngRow, AC_����ʱ��) = Format(curDate, "yyyy-MM-dd HH:mm")
        End If
        
        'ʼ�ձ���һ����
        If lngRow = .Rows - 1 Then
            .AddItem "", lngRow + 1
        End If
        .Tag = ""
        mblnChange = True
    End With
End Sub

Private Sub OPSSetInput(ByVal lngRow As Long, ByVal lngCol As Long, rsInput As ADODB.Recordset)
'���ܣ�������������������������ñ������
    Dim rsTmp As New ADODB.Recordset
    Dim StrSQL As String, i As Long
    
    With vsOPS
        If lngCol = col�������� Or lngCol = col�������� Then
            If Not rsInput Is Nothing Then
                .TextMatrix(lngRow, col��������) = rsInput!����
                 .Cell(flexcpData, lngRow, col��������) = .TextMatrix(lngRow, col��������)
                .TextMatrix(lngRow, col��������) = rsInput!����
                If optInput(4).Value Then
                    .TextMatrix(lngRow, col������ĿID) = rsInput!ID
                    .TextMatrix(lngRow, col��������ID) = ""
                    StrSQL = "Select ����ID as ID From ������϶��� Where ����ID=[1]"
                Else
                    .TextMatrix(lngRow, col��������ID) = rsInput!ID
                    .TextMatrix(lngRow, col������ĿID) = ""
                    StrSQL = "Select ����ID as ID From ������϶��� Where ����ID=[1]"
                End If
                On Error GoTo errH
                Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, Val(rsInput!ID))
                If Not rsTmp.EOF Then
                    If optInput(4).Value Then
                        .TextMatrix(lngRow, col��������ID) = Val(rsTmp!ID)
                    Else
                        .TextMatrix(lngRow, col������ĿID) = Val(rsTmp!ID)
                    End If
                End If
            Else
                .TextMatrix(lngRow, lngCol) = .EditText
                .TextMatrix(lngRow, col��������ID) = ""
                .TextMatrix(lngRow, col������ĿID) = ""
            End If
            .Cell(flexcpData, lngRow, lngCol) = .TextMatrix(lngRow, lngCol)
            
            '����������ͬʱ��������������Ĭ������һ����ͬ
            If Not rsInput Is Nothing And lngRow > .FixedRows And lngRow = .Rows - 1 Then
                If .TextMatrix(lngRow, col��������) = .TextMatrix(lngRow - 1, col��������) Then
                    .TextMatrix(lngRow, col����ҽʦ) = .TextMatrix(lngRow - 1, col����ҽʦ)
                    .TextMatrix(lngRow, col������ʿ) = .TextMatrix(lngRow - 1, col������ʿ)
                    .TextMatrix(lngRow, col����1) = .TextMatrix(lngRow - 1, col����1)
                    .TextMatrix(lngRow, col����2) = .TextMatrix(lngRow - 1, col����2)
                    .TextMatrix(lngRow, col����ʽ) = .TextMatrix(lngRow - 1, col����ʽ)
                    .TextMatrix(lngRow, col����ҽʦ) = .TextMatrix(lngRow - 1, col����ҽʦ)
                    .TextMatrix(lngRow, col�п�����) = .TextMatrix(lngRow - 1, col�п�����)
                    .TextMatrix(lngRow, col����ID) = .TextMatrix(lngRow - 1, col����ID)
                    .TextMatrix(lngRow, col��������) = .TextMatrix(lngRow - 1, col��������)
                    
                    For i = col����ҽʦ To .Cols - 1
                        .Cell(flexcpData, lngRow, i) = .TextMatrix(lngRow, i)
                    Next
                End If
            End If
            
            '�����ʼ�ձ���һ����
            If lngRow = .Rows - 1 Then .AddItem ""
        ElseIf lngCol = col����ʽ Then
            .TextMatrix(lngRow, lngCol) = rsInput!����
            .Cell(flexcpData, lngRow, lngCol) = .TextMatrix(lngRow, lngCol)
            .TextMatrix(lngRow, col����ID) = rsInput!ID
            .TextMatrix(lngRow, col��������) = Nvl(rsInput!��������)
        ElseIf lngCol = col����ҽʦ Or lngCol = col������ʿ Or lngCol = col����1 Or lngCol = col����2 Or lngCol = col����ҽʦ Then
            .TextMatrix(lngRow, lngCol) = rsInput!����
            .Cell(flexcpData, lngRow, lngCol) = .TextMatrix(lngRow, lngCol)
        End If
        
        '������Ϸ������
        Call Set��Ϸ������(cbo��ǰ������)
        
        .Tag = ""
        mblnChange = True
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsOPS_ComboDropDown(ByVal Row As Long, ByVal Col As Long)
    Dim i As Long
    With vsOPS
        If Col = col�п����� Then
            For i = 0 To .ComboCount - 1
                If NeedName(.ComboItem(i)) = .TextMatrix(Row, Col) Then
                    .ComboIndex = i: Exit For
                End If
            Next
        End If
    End With
End Sub

Private Sub vsOPS_DblClick()
    Call vsOPS_KeyPress(32)
End Sub

Private Sub vsOPS_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long
    
    If mbln��ʿվ Or mblnReadOnly Then Exit Sub
    
    With vsOPS
        If KeyCode = vbKeyF4 Then
            If .ComboList = "..." Then
                Call zlCommFun.PressKey(vbKeySpace)
            End If
        ElseIf KeyCode = vbKeyDelete Then
            If .TextMatrix(.Row, col��������) <> "" Then
                If MsgBox("ȷʵҪɾ������������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
                    .RemoveItem .Row
                    
                    '������Ϸ������
                    Call Set��Ϸ������(cbo��ǰ������)

                    mblnChange = True
                    .Tag = ""
                End If
            End If
        ElseIf KeyCode > 127 Then
            '���ֱ�����뺺�ֵ�����
            Call vsOPS_KeyPress(KeyCode)
        End If
    End With
End Sub

Private Sub vsOPS_KeyPress(KeyAscii As Integer)
    If mbln��ʿվ Or mblnReadOnly Then Exit Sub
    
    With vsOPS
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call OPSEnterNextCell
        Else
            If .ComboList = "..." Then
                If KeyAscii = Asc("*") Then
                    KeyAscii = 0
                    Call vsOPS_CellButtonClick(.Row, .Col)
                Else
                    .ComboList = "" 'ʹ��ť״̬��������״̬
                End If
            End If
        End If
    End With
End Sub

Private Sub vsOPS_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    Dim strInput As String
    
    With vsOPS
        If KeyAscii = 13 Then
            mblnReturn = True
            
            If Col = col�������� Then
                KeyAscii = 0
                strInput = GetFullDate(.EditText, False)
                If IsDate(strInput) Then
                    .TextMatrix(Row, Col) = strInput
                    .Cell(flexcpData, Row, Col) = .TextMatrix(Row, Col)
                    mblnChange = True
                    .Tag = ""
                    Call OPSEnterNextCell
                End If
            ElseIf Col = col�п����� Then
                KeyAscii = 0
                If .ComboIndex <> -1 Then
                    .TextMatrix(Row, Col) = NeedName(.ComboItem(.ComboIndex))
                    mblnChange = True
                    .Tag = ""
                    Call OPSEnterNextCell
                End If
            End If
        Else
            mblnReturn = False
            
            If Col = col�������� Then
                If InStr("0123456789-" & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                End If
            ElseIf Col = col����ҩ���� Then
                If InStr("0123456789" & Chr(8) & Chr(27), Chr(KeyAscii)) = 0 Then
                    KeyAscii = 0
                End If
            End If
        End If
    End With
End Sub

Private Sub vsOPS_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsOPS.EditSelStart = 0
    vsOPS.EditSelLength = zlCommFun.ActualLen(vsOPS.EditText)
End Sub

Private Function OPSCellEditable(ByVal lngRow As Long, ByVal lngCol As Long) As Boolean
    With vsOPS
        If .ColHidden(lngCol) Then Exit Function
        
        '������������������,��������
        If Not IsDate(.TextMatrix(lngRow, col��������)) Then
            If lngCol > col�������� Then Exit Function
        End If
        If .TextMatrix(lngRow, col��������) = "" Then
            If lngCol > col�������� Then Exit Function
        End If
        
        '��������������ҽʦ
        If .TextMatrix(lngRow, col����ҽʦ) = "" Then
            If lngCol = col����1 Or lngCol = col����2 Then Exit Function
        End If
        
        '�����������1����
        If .TextMatrix(lngRow, col����1) = "" Then
            If lngCol = col����2 Then Exit Function
        End If
        
        '��������������ʽ
        If Trim(.TextMatrix(lngRow, col��������)) = "" Then
            If lngCol = col����ҽʦ Then Exit Function
        End If
        
        '�������Ʋ�������
        If lngCol = col�������� And chkInfo(chk��������¼��).Value = 0 Then Exit Function
    End With
    OPSCellEditable = True
End Function

Private Sub OPSEnterNextCell()
    Dim i As Long, j As Long
    
    With vsOPS
        '����һ��Ԫ��ʼѭ������
        For i = .Row To .Rows - 1
            For j = IIf(i = .Row, .Col + 1, col��������) To col�����Źؽڹ���
                If OPSCellEditable(i, j) Then Exit For
            Next
            If j <= col�����Źؽڹ��� Then Exit For
        Next
        If i <= .Rows - 1 Then
            Call .Select(i, j)
            .ShowCell .Row, .Col
        Else
            Call zlCommFun.PressKey(vbKeyTab)
        End If
    End With
End Sub

Private Sub AllerEnterNextCell()
    Dim i As Long, j As Long
    
    With vsAller
        If .Col = AC_������Ӧ Then
            If .Row + 1 <= .Rows - 1 Then
                .Row = .Row + 1
                .Col = AC_����ҩ��
                .ShowCell .Row, .Col
            Else
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Else
            .Col = .Col + 1
            .ShowCell .Row, .Col
        End If
    End With
End Sub

Private Function ShowMessage(objTmp As Object, ByVal strMsg As String, Optional ByVal blnAsk As Boolean) As VbMsgBoxResult
'���ܣ���ʾ��ʾ��Ϣ����λ��������Ŀ��
    Dim lngColor As Long

    If UCase(objTmp.Container.Name) <> UCase("fraInfo") Then
        If UCase(objTmp.Container.Container.Name) = UCase("fraInfo") Then sstInfo.Tab = objTmp.Container.Container.Index
    Else
        sstInfo.Tab = objTmp.Container.Index
    End If
    If UCase(TypeName(objTmp)) <> UCase("VSFlexGrid") Then
        lngColor = objTmp.BackColor: objTmp.BackColor = &HC0C0FF
    Else
        lngColor = objTmp.CellBackColor: objTmp.CellBackColor = &HC0C0FF
        Call objTmp.ShowCell(objTmp.Row, objTmp.Col)
    End If
    If Not blnAsk Then
        MsgBox strMsg, vbInformation, gstrSysName
    Else
        ShowMessage = MsgBox(strMsg, vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName)
    End If
    If UCase(TypeName(objTmp)) <> UCase("VSFlexGrid") Then
        objTmp.BackColor = lngColor
    Else
        objTmp.CellBackColor = lngColor
    End If
    If objTmp.Enabled And objTmp.Visible Then objTmp.SetFocus
    Me.Refresh
End Function

Private Function CheckPageData(ByRef blnDiagnose As Boolean, ByVal blnBeforSign As Boolean) As Boolean
'���ܣ������ҳ�������ݺϷ���
'���أ�blnDiagnose=�Ƿ���д�����
'������blnBeforSign-�Ƿ�ǩ��ʱ����ǰ����
    Dim objTmp As Object, curDate As Date
    Dim arrInfo() As Variant, arrName As Variant
    Dim str���֤ As String, str�������� As String, lng�Ա� As Long
    Dim lng�������� As Long, str���� As String
    Dim str����IDs As String, str���IDs As String
    Dim i As Long, j As Long
    Dim StrSQL As String, rsTmp As Recordset
    
    
    blnDiagnose = False
    
    '��Ŀ����ĳ��ȼ��
    '-----------------------------------------------------------------------------------------
    For Each objTmp In txtInfo
        If objTmp.Enabled And Not objTmp.Locked And objTmp.MaxLength <> 0 Then
            If zlCommFun.ActualLen(objTmp.Text) > objTmp.MaxLength Then
                Call ShowMessage(objTmp, "�������ݹ��������顣(����Ŀ������� " & objTmp.MaxLength & " ���ַ��� " & objTmp.MaxLength \ 2 & " ������)")
                Exit Function
            End If
        End If
    Next
    If Not mbln��ʿվ Then
        curDate = zlDatabase.Currentdate
        
        '����Ҫ��������ݼ��
        '-----------------------------------------------------------------------------------------
        arrInfo = Array(txtסԺ��, txt����, txt����, txt����)
        arrName = Array("סԺ��", "����", "����", "����")
        For i = 0 To UBound(arrInfo)
            If txtInfo(arrInfo(i)).Enabled And Not txtInfo(arrInfo(i)).Locked And txtInfo(arrInfo(i)).Text = "" Then
                If arrName(i) <> "����" Or mlng���� = 1 Then
                    Call ShowMessage(txtInfo(arrInfo(i)), "�������벡�˵�" & arrName(i) & "��")
                    Exit Function
                ElseIf mlng���� = 2 Then
                    If ShowMessage(txtInfo(arrInfo(i)), "û�����벡�˵�" & arrName(i) & ",�Ƿ������", True) = vbNo Then
                        Exit Function
                    End If
                End If
            End If
        Next
        
        Select Case cboinfo(cbo���䵥λ).Text
            Case "��"
                If Val(txtInfo(txt����).Text) > 200 Then
                    Call ShowMessage(txtInfo(txt����), "����ֵ�������������200�꣬���������Ƿ���ȷ��")
                    Exit Function
                End If
            Case "��"
                If Val(txtInfo(txt����).Text) > 2400 Then
                    Call ShowMessage(txtInfo(txt����), "����ֵ�������������2400�£����������Ƿ���ȷ��")
                    Exit Function
                End If
            Case "��"
                If Val(txtInfo(txt����).Text) > 73000 Then
                    Call ShowMessage(txtInfo(txt����), "����ֵ�����������73000�죬���������Ƿ���ȷ��")
                    Exit Function
                End If
            Case "Сʱ" '���ܴ���30�켴720Сʱ
                If Val(txtInfo(txt����).Text) > 720 Then
                    Call ShowMessage(txtInfo(txt����), "����ֵ�������������720Сʱ����ʹ�ú��ʵ����䵥λ��")
                    Exit Function
                End If
            Case "����" '���ܴ���24Сʱ��1440����
                If Val(txtInfo(txt����).Text) > 1440 Then
                    Call ShowMessage(txtInfo(txt����), "����ֵ�������������1440���ӣ���ʹ�ú��ʵ����䵥λ��")
                    Exit Function
                End If
        End Select
        
        Select Case cboinfo(cboӤ�����䵥λ).Text
            Case "��" 'һ��
                If Val(txtInfo(txtӤ������).Text) > 12 Then
                    Call ShowMessage(txtInfo(txtӤ������), "Ӥ������ֵ�������������12�£����������Ƿ���ȷ��")
                    Exit Function
                End If
            Case "��" '365��
                If Val(txtInfo(txtӤ������).Text) > 365 Then
                    Call ShowMessage(txtInfo(txtӤ������), "Ӥ������ֵ�������������365�죬��ʹ�ú��ʵ����䵥λ��")
                    Exit Function
                End If
            Case "Сʱ" '���ܴ���30�켴720Сʱ
                If Val(txtInfo(txtӤ������).Text) > 720 Then
                    Call ShowMessage(txtInfo(txtӤ������), "Ӥ������ֵ�������������720Сʱ����ʹ�ú��ʵ����䵥λ��")
                    Exit Function
                End If
            Case "����" '���ܴ���24Сʱ��1440����
                If Val(txtInfo(txtӤ������).Text) > 1440 Then
                    Call ShowMessage(txtInfo(txtӤ������), "Ӥ������ֵ�������������1440���ӣ���ʹ�ú��ʵ����䵥λ��")
                    Exit Function
                End If
        End Select
        If Not IsDate(txt��������.Text) Then
            Call ShowMessage(txt��������, "�������벡�˵ĳ������ڡ�")
            Exit Function
        ElseIf txt����ʱ��.Text <> "__:__" And Not IsDate(txt����ʱ��.Text) Then
            Call ShowMessage(txt����ʱ��, "��������ȷ�Ĳ��˳���ʱ�䡣")
            Exit Function
        End If
        
        arrInfo = Array(cbo���ʽ, cbo�Ա�, cbo����, cboְҵ, cbo��Ժ����)
        arrName = Array("���ʽ", "�Ա�", "����", "ְҵ", "��Ժ����")
        For i = 0 To UBound(arrInfo)
            If cboinfo(arrInfo(i)).Enabled And Not cboinfo(arrInfo(i)).Locked And cboinfo(arrInfo(i)).ListIndex = -1 Then
                Call ShowMessage(cboinfo(arrInfo(i)), "�������벡�˵�" & arrName(i) & "��")
                Exit Function
            End If
        Next
        
        If txtInfo(txtת��3).Text <> "" And txtInfo(txtת��2).Text = "" Then
            Call ShowMessage(txtInfo(txtת��2), "����������ת�ƿ��ҡ�")
            Exit Function
        End If
        If txtInfo(txtת��2).Text <> "" And txtInfo(txtת��1).Text = "" Then
            Call ShowMessage(txtInfo(txtת��1), "����������ת�ƿ��ҡ�")
            Exit Function
        End If
        If txtInfo(txtת��1).Text = txtInfo(txtת��2).Text And txtInfo(txtת��1).Text <> "" Then
            Call ShowMessage(txtInfo(txtת��2), "ת�Ƶ��������Ҳ�Ӧ����ͬ��")
            Exit Function
        End If
        If txtInfo(txtת��2).Text = txtInfo(txtת��3).Text And txtInfo(txtת��2).Text <> "" Then
            Call ShowMessage(txtInfo(txtת��3), "ת�Ƶ��������Ҳ�Ӧ����ͬ��")
            Exit Function
        End If
        
        If cboinfo(cbo������).ListIndex = -1 And cboinfo(cbo����ҽʦ).ListIndex = -1 _
            And cboinfo(cbo����ҽʦ).ListIndex = -1 And cboinfo(cboסԺҽʦ).ListIndex = -1 Then
            Call ShowMessage(cboinfo(cbo������), "���ڿ����Ρ�����ҽʦ������ҽʦ��סԺҽʦ֮������ѡ��һλ��")
            Exit Function
        End If
            
    
        '���䳤��Ҫ���ϵ�λ
        If zlCommFun.ActualLen(txtInfo(txt����).Text & cboinfo(cbo���䵥λ).Text) > txtInfo(txt����).MaxLength Then
            Call ShowMessage(txtInfo(txt����), "�������ݹ��������顣(����Ŀ������� " & txtInfo(txt����).MaxLength & " ���ַ��� " & txtInfo(txt����).MaxLength \ 2 & " ������)")
        End If
        
        If txt����ʱ��.Text <> "____-__-__ __:__:__" Then
            If Not IsDate(txt����ʱ��.Text) Then
                Call ShowMessage(txt����ʱ��, "����ʱ�䲻����Ч�����ڸ�ʽ��")
                Exit Function
            End If
            If Format(txt����ʱ��.Text, "yyyy-MM-dd HH:mm:ss") <= Format(txtInfo(txt��Ժʱ��).Text, "yyyy-MM-dd HH:mm:ss") Then
                Call ShowMessage(txt����ʱ��, "����ʱ��Ӧ����Ժʱ����")
                Exit Function
            End If
        End If
        
        '�������ݵ���Ч�Լ��
        '-----------------------------------------------------------------------------------------
        '�������ڱ������ڵ�ǰʱ��
        If Format(txt��������.Text, "yyyy-MM-dd") > Format(curDate, "yyyy-MM-dd") Then
            Call ShowMessage(txt��������, "�������ڲ�Ӧ�ñȵ�ǰ���ڻ���")
            Exit Function
        End If
        
        '�������ڱ���������Ժʱ��
        If Trim(txtInfo(txt��Ժʱ��).Text) <> "" Then
            If Format(txt��������.Text, "yyyy-MM-dd") > Format(txtInfo(txt��Ժʱ��).Text, "yyyy-MM-dd") Then
                Call ShowMessage(txt��������, "�������ڲ�Ӧ�ñ���Ժʱ�仹��")
                Exit Function
            End If
        End If
        
        '������������ڵ�ƥ����
        If IsNumeric(txtInfo(txt����).Text) And cboinfo(cbo���䵥λ).ListIndex <> -1 And cboinfo(cbo���䵥λ).ListIndex < 3 Then
            str���� = PatiAgeCalc(txt��������.Text, , txtInfo(txt��Ժʱ��).Text)
            If Right(str����, 1) = cboinfo(cbo���䵥λ).Text And IsNumeric(Left(str����, Len(str����) - 1)) _
                And str���� <> txtInfo(txt����).Text & cboinfo(cbo���䵥λ).Text Then
                If ShowMessage(txt��������, "����ͳ������ڲ�һ�£�" & txt��������.Text & "��������Ӧ����" & str���� & "��" & _
                    vbCrLf & vbCrLf & "���������������ڵ���ȷ�ԣ�Ҫ������", True) = vbNo Then
                    Exit Function
                End If
            End If
        End If
        
        '15������ӦΪδ��
        If DateDiff("yyyy", CDate(txt��������.Text), curDate) < 15 Then
            If InStr(cboinfo(cbo����).Text, "δ��") = 0 Then
                Call ShowMessage(cboinfo(cbo����), "�ò�������С��15�꣬����״��Ӧ��дΪδ�顣")
                Exit Function
            End If
        End If
                
        '���֤������
        '�����֤�Ž�����֤
        str���֤ = cboinfo(cbo���֤��).Text
        If str���֤ <> "" Then
            If zlCommFun.ActualLen(str���֤) = Len(str���֤) Then
                If Len(str���֤) <> 15 And Len(str���֤) <> 18 Then
                    Call ShowMessage(cboinfo(cbo���֤��), "���֤����ĳ��Ȳ���ȷ��ӦΪ15λ��18λ��")
                    Exit Function
                End If
                
                If Len(str���֤) = 15 Then
                    str�������� = Mid(str���֤, 7, 6)
                    str�������� = Format(GetFullDate(str��������), "yyyy-MM-dd")
                    lng�Ա� = Val(Right(str���֤, 1))
                Else
                    str�������� = Mid(str���֤, 7, 8)
                    str�������� = Format(GetFullDate(str��������), "yyyy-MM-dd")
                    lng�Ա� = Val(Mid(str���֤, 17, 1))
                End If
                If Not IsDate(str��������) Then
                    If ShowMessage(cboinfo(cbo���֤��), "���֤�����еĳ���������Ϣ����ȷ���Ƿ������", True) = vbNo Then Exit Function
                Else
                    If Format(str��������, "yyyy-MM-dd") <> Format(txt��������.Text, "yyyy-MM-dd") Then
                        If ShowMessage(cboinfo(cbo���֤��), "���֤�����еĳ���������Ϣ�벡�˵ĳ������ڲ������Ƿ������", True) = vbNo Then Exit Function
                    End If
                End If
                If (lng�Ա� Mod 2 = 1 And InStr(cboinfo(cbo�Ա�).Text, "Ů") > 0) Or (lng�Ա� Mod 2 = 0 And InStr(cboinfo(cbo�Ա�).Text, "��") > 0) Then
                    If ShowMessage(cboinfo(cbo���֤��), "���֤�����е��Ա���Ϣ�벡�˵��Ա𲻷����Ƿ������", True) = vbNo Then Exit Function
                End If
            Else
                If zlCommFun.ActualLen(str���֤) > 18 Then
                    Call ShowMessage(cboinfo(cbo���֤��), "���ܳ���9�����ֵĳ��ȣ����顣")
                    Exit Function
                End If
            End If
        End If
        
        'ȷ��ʱ���������Ժʱ��ͳ�Ժʱ��֮��
        If IsDate(txtInfo(txtȷ������).Text) Then
            If Not Between(Format(txtInfo(txtȷ������).Text, "yyyy-MM-dd"), Format(txtInfo(txt��Ժʱ��).Tag, "yyyy-MM-dd"), _
                Format(IIf(txtInfo(txt��Ժʱ��).Text = "", zlDatabase.Currentdate, txtInfo(txt��Ժʱ��).Text), "yyyy-MM-dd")) Then
                Call ShowMessage(txtInfo(txtȷ������), "ȷ��ʱ���������Ժʱ��ͳ�Ժʱ��֮�䡣")
                Exit Function
            End If
        ElseIf chkInfo(chk�Ƿ�ȷ��).Value = 1 Then
            Call ShowMessage(txtInfo(txtȷ������), "ȷ��ʱ���������")
            Exit Function
        End If
        
        '��Ժ����ΪΣʱ��Ҫ��������
        If InStr(cboinfo(cbo��Ժ����).Text, "Σ") > 0 And Val(txtInfo(txt���ȴ���).Text) = 0 Then
            If ShowMessage(txtInfo(txt���ȴ���), "�ò�����Ժ����ΪΣ����û�н������ȣ��Ƿ������", True) = vbNo Then Exit Function
        End If
        
        '�ɹ��������ܳ������ȴ���
        If Val(txtInfo(txt�ɹ�����).Text) > Val(txtInfo(txt���ȴ���).Text) Then
            Call ShowMessage(txtInfo(txt�ɹ�����), "�ɹ��������ܳ������ȴ�����")
            Exit Function
        End If
        '�ɹ�����С�����ȴ���ʱ��Ժ���ӦΪ���� 2010-03-23 27224 ����ʱ�򣬳ɹ��������Ե������ȴ�������Ϊ��   ����û�����Ⱦ����ˡ�
        If InStr(vsDiagXY.TextMatrix(GetRow(3), col��Ժ���), "����") > 0 Then
            If Val(txtInfo(txt�ɹ�����).Text) > Val(txtInfo(txt���ȴ���).Text) And txtInfo(txt���ȴ���).Text <> "" Then
                Call ShowMessage(txtInfo(txt�ɹ�����), "���˳�Ժ���Ϊ�������ɹ��������ܴ������ȴ�����")
                Exit Function
            End If
        Else
            If Val(txtInfo(txt�ɹ�����).Text) <> Val(txtInfo(txt���ȴ���).Text) And txtInfo(txt���ȴ���).Text <> "" Then
                If InStr(vsDiagXY.TextMatrix(GetRow(3), col��Ժ���), "����") > 0 Then
                    If ShowMessage(txtInfo(txt�ɹ�����), "���˳�Ժ�����Ϊ�������ɹ�����Ӧ�������ȴ������Ƿ������", True) = vbNo Then Exit Function
                Else
                    Call ShowMessage(txtInfo(txt�ɹ�����), "���˳�Ժ�����Ϊ�������ɹ�����Ӧ�������ȴ�����")
                    Exit Function
                End If
            End If
        End If
        '�ɹ������������ȴ�����һ��
        If Val(txtInfo(txt���ȴ���).Text) - Val(txtInfo(txt�ɹ�����).Text) > 1 And txtInfo(txt���ȴ���).Text <> "" Then
            Call ShowMessage(txtInfo(txt�ɹ�����), "�ɹ������������ȴ�����һ�Ρ�")
            Exit Function
        End If
        
        '������
        If chkInfo(chk����).Value = 1 Then
            If Val(txtInfo(txt��������).Text) <= 0 And cboinfo(cbo����Ex).Text <> "����" Then
                Call ShowMessage(txtInfo(txt��������), "��������ȷ���������ޡ�")
                Exit Function
            End If
        End If
        
        '31������סԺ�ƻ���Ŀ��
        If optInput(opt31����).Value Then
            If Trim(txtInfo(txt31��Ŀ��).Text) = "" Then
                Call ShowMessage(txtInfo(txt31��Ŀ��), "����д" & cboinfo(cbo31���7������Ժ).Text & "��Ŀ�ġ�")
                Exit Function
            End If
        End If
        '��д����ǰ�����󣬱�����д�������
        If vsOPS.TextMatrix(1, col��������) = "" And cboinfo(cbo��ǰ������).ListIndex > 0 Then
            Call ShowMessage(cboinfo(cbo��ǰ������), "û����д�������,��ǰ������ֻ��ѡ��""δ��""��")
            Exit Function
        End If
        
        '����ʱ����
        If txt��������.Text <> "____-__-__" Then
            If Not IsDate(txt��������.Text) Then
                Call ShowMessage(txt��������, "��������ȷ�ķ������ڡ�")
                Exit Function
            Else
                If txt����ʱ��.Text <> "__:__" Then
                    If Not IsDate(txt����ʱ��.Text) Then
                        Call ShowMessage(txt����ʱ��, "��������ȷ�ķ���ʱ�䡣")
                        Exit Function
                    End If
                End If
                
                If txt��������.Text & IIf(txt����ʱ��.Text = "__:__", "", " " & txt����ʱ��.Text) _
                    >= Format(curDate, txt��������.Format & IIf(txt����ʱ��.Text = "__:__", "", " " & txt����ʱ��.Format)) Then
                    Call ShowMessage(txt��������, "����ʱ��Ӧ�����ڵ�ǰʱ�䡣")
                    Exit Function
                End If
            End If
        End If
        
        '���ļ��
        '-----------------------------------------------------------------------------------------
        With vsOPS
            For i = .FixedRows To .Rows - 1
                If Trim(.TextMatrix(i, col��������)) <> "" Then
                    lng�������� = lng�������� + 1
                End If
            Next
        End With
        
        With vsDiagXY
            For i = .FixedRows To .Rows - 1
                If .TextMatrix(i, col�������) <> "" And .TextMatrix(i - 1, col�������) = "" _
                    And Val(.TextMatrix(i, col����)) = Val(.TextMatrix(i - 1, col����)) Then
                    .Row = i - 1: .Col = col�������
                    Call ShowMessage(vsDiagXY, "���������������Ϣ��")
                    Exit Function
                End If
                
                If Trim(.TextMatrix(i, col�������)) <> "" Then
                    If zlCommFun.ActualLen(.TextMatrix(i, col�������)) > 200 Then
                        .Row = i: .Col = col�������
                        Call ShowMessage(vsDiagXY, IIf(.TextMatrix(i, col�������) = "", "��Ժ���", .TextMatrix(i, col�������)) & "����̫����ֻ����200���ַ���100�����֡�")
                        Exit Function
                    End If
                    If zlCommFun.ActualLen(.TextMatrix(i, col��ע)) > 50 Then
                        .Row = i: .Col = col��ע
                        Call ShowMessage(vsDiagXY, """" & .TextMatrix(i, col�������) & """�ı�ע����̫����ֻ����50���ַ���25�����֡�")
                        Exit Function
                    End If
                    If Val(.TextMatrix(i, col����)) = 5 Then    'Ժ�ڸ�Ⱦ
                        If .TextMatrix(i, col��Ժ���) = "" Then
                            .Row = i: .Col = col��Ժ���
                            If ShowMessage(vsDiagXY, "Ժ�ڸ�Ⱦ�ĳ�Ժ���û����д���Ƿ������", True) = vbNo Then Exit Function
                        End If
                    End If
                    If Val(.TextMatrix(i, col����)) = 3 Then
                        If .TextMatrix(i, col��Ժ���) = "" Then
                            .Row = i: .Col = col��Ժ���
                            Call ShowMessage(vsDiagXY, "����д��Ժ��ϵĳ�Ժ�����")
                            Exit Function
                        ElseIf Val(.TextMatrix(i - 1, col����)) <> 3 And InStr(.TextMatrix(i, col��Ժ���), "����") > 0 And lng�������� > 0 Then
                            .Row = i: .Col = col��Ժ���
                            If ShowMessage(vsDiagXY, "�ò��˽���������������Ժ���ѡ��Ϊ�������Ƿ������", True) = vbNo Then Exit Function
    '                    ElseIf Val(.TextMatrix(i - 1, col����)) = 3 And InStr(.TextMatrix(GetRow(3), col��Ժ���), "����") > 0 And InStr(.TextMatrix(i, col��Ժ���), "����") = 0 Then
    '                        .Row = i: .Col = col��Ժ���
    '                        Call ShowMessage(vsDiagXY, "��Ҫ��ϵĳ�Ժ���Ϊ��������������ϵĳ�Ժ���ȴ����""" & .TextMatrix(i, col��Ժ���) & """��")
    '                        Exit Function
                        ElseIf Val(.TextMatrix(i - 1, col����)) = 3 And InStr(.TextMatrix(GetRow(3), col��Ժ���), "����") = 0 And InStr(.TextMatrix(i, col��Ժ���), "����") > 0 Then
                            .Row = i: .Col = col��Ժ���
                            Call ShowMessage(vsDiagXY, "��Ҫ��ϵĳ�Ժ�����Ϊ��������������ϵĳ�Ժ���ȴΪ������")
                            Exit Function
                        ElseIf Val(.TextMatrix(i - 1, col����)) <> 3 And InStr(.TextMatrix(i, col��Ժ���), "����") > 0 And Val(txtInfo(txtסԺ����).Text) < 3 Then
                            .Row = i: .Col = col��Ժ���
                            If ShowMessage(vsDiagXY, "�ò���סԺ��ԺΪ " & Val(txtInfo(txtסԺ����).Text) & " �죬��Ժ���ȴΪ�������Ƿ������", True) = vbNo Then Exit Function
                        ElseIf .TextMatrix(i, col�������) = "��Ժ���" Then
                            If mlng�����ж� <> 0 Then
                                '��Ҫ�����Ҫ�����˵��ⲿԭ��
                                If InStr("ST", Left(.TextMatrix(i, col��ϱ���), 1)) > 0 And Left(.TextMatrix(i, col��ϱ���), 1) <> "" Then
                                    '��Ҫ�����ж��ⲿԭ��
                                    If .TextMatrix(GetRow(7), col�������) = "" Then
                                        If Not sstInfo.TabVisible(TAB_��ҽ���) Then
                                            .Row = GetRow(7): .Col = col�������
                                            If mlng�����ж� = 1 Then
                                                Call ShowMessage(vsDiagXY, "����д�����ж���ԭ��")
                                                Exit Function
                                            Else
                                                If ShowMessage(vsDiagXY, "û����д�����ж���ԭ��,�Ƿ������", True) = vbNo Then Exit Function
                                            End If
                                        End If
                                    End If
                                Else
                                    If .TextMatrix(GetRow(7), col�������) <> "" Then
                                        .Row = GetRow(7): .Col = col�������
                                        If mlng�����ж� = 1 Then
                                            Call ShowMessage(vsDiagXY, "������д�����ж���ԭ��")
                                            Exit Function
                                        Else
                                            If ShowMessage(vsDiagXY, "��Ժ����������ж���ԭ�򲻷�,�Ƿ������", True) = vbNo Then Exit Function
                                        End If
                                    End If
                                End If
                            End If
                            If mlng������� <> 0 Then
                                '��Ҫ�����Ҫ��д������ϵ��ⲿԭ��
                                If InStr("CD", Left(.TextMatrix(i, col��ϱ���), 1)) > 0 And Left(.TextMatrix(i, col��ϱ���), 1) <> "" Then
                                    '��Ҫ������ϵ��ⲿԭ��
                                    If .TextMatrix(GetRow(6), col�������) = "" Then
                                        If Not sstInfo.TabVisible(TAB_��ҽ���) Then
                                            .Row = GetRow(6): .Col = col�������
                                            If mlng������� = 1 Then
                                                Call ShowMessage(vsDiagXY, "����д������ϡ�")
                                                Exit Function
                                            Else
                                                If ShowMessage(vsDiagXY, "û����д�������,�Ƿ������", True) = vbNo Then Exit Function
                                            End If
                                        End If
                                    End If
                                Else
                                    If .TextMatrix(GetRow(6), col�������) <> "" Then
                                        .Row = GetRow(6): .Col = col�������
                                        If mlng������� = 1 Then
                                            Call ShowMessage(vsDiagXY, "������д������ϡ�")
                                            Exit Function
                                        Else
                                            If ShowMessage(vsDiagXY, "��Ժ����벡����ϲ���,�Ƿ������", True) = vbNo Then Exit Function
                                        End If
                                    End If
                                End If
                            End If
                        End If
                        
                        For j = GetRow(3) To .Rows - 1
                            If Val(.TextMatrix(j, col����)) = 3 Then
                                If j <> i And .TextMatrix(j, col�������) <> "" Then
                                    If .TextMatrix(j, col�������) = .TextMatrix(i, col�������) Then
                                        .Row = i: .Col = col�������
                                        Call ShowMessage(vsDiagXY, "���ִ���������ͬ�ĳ�Ժ�����Ϣ��")
                                        Exit Function
                                    ElseIf Val(.TextMatrix(i, col����ID)) <> 0 Then
                                        If Val(.TextMatrix(j, col����ID)) = Val(.TextMatrix(i, col����ID)) Then
                                            .Row = i: .Col = col�������
                                            Call ShowMessage(vsDiagXY, "���ִ���������ͬ�ĳ�Ժ�����Ϣ��")
                                            Exit Function
                                        End If
                                    ElseIf Val(.TextMatrix(i, col���ID)) <> 0 Then
                                        If Val(.TextMatrix(j, col���ID)) = Val(.TextMatrix(i, col���ID)) Then
                                            .Row = i: .Col = col�������
                                            Call ShowMessage(vsDiagXY, "���ִ���������ͬ�ĳ�Ժ�����Ϣ��")
                                            Exit Function
                                        End If
                                    End If
                                End If
                            End If
                        Next
                    End If
                    
                    If Val(.TextMatrix(i, col����ID)) <> 0 Then str����IDs = str����IDs & "," & Val(.TextMatrix(i, col����ID))
                    If Val(.TextMatrix(i, col���ID)) <> 0 Then str���IDs = str���IDs & "," & Val(.TextMatrix(i, col���ID))
                    
                    '�Ƿ�������Ҫ����������
                    If InStr("," & mstr���� & ",", "," & Val(.TextMatrix(i, col����)) & ",") > 0 Then
                        blnDiagnose = True
                    End If
                End If
            Next
        End With
            
        If mbln��ҽ Then
            With vsDiagZY
                For i = .FixedRows To .Rows - 1
                    If .TextMatrix(i, col�������) <> "" And .TextMatrix(i - 1, col�������) = "" _
                        And Val(.TextMatrix(i, colzy����)) = Val(.TextMatrix(i - 1, colzy����)) Then
                        .Row = i - 1: .Col = col�������
                        Call ShowMessage(vsDiagZY, "���������������Ϣ��")
                        Exit Function
                    End If
                
                    If Trim(.TextMatrix(i, col�������)) <> "" Then
                        If zlCommFun.ActualLen(.TextMatrix(i, col�������)) > 200 Then
                            .Row = i: .Col = col�������
                            Call ShowMessage(vsDiagZY, IIf(.TextMatrix(i, col�������) = "", "��Ժ���", .TextMatrix(i, col�������)) & "����̫����ֻ����200���ַ���100�����֡�")
                            Exit Function
                        End If
                        If zlCommFun.ActualLen(.TextMatrix(i, col��ע)) > 50 Then
                            .Row = i: .Col = col��ע
                            Call ShowMessage(vsDiagZY, """" & .TextMatrix(i, col�������) & """�ı�ע����̫����ֻ����50���ַ���25�����֡�")
                            Exit Function
                        End If
                        If Val(.TextMatrix(i, colzy����)) = 13 Then
                            If .TextMatrix(i, col��Ժ���) = "" Then
                                .Row = i: .Col = col��Ժ���
                                Call ShowMessage(vsDiagZY, "����д��Ժ��ϵĳ�Ժ�����")
                                Exit Function
    '                        ElseIf Val(.TextMatrix(i - 1, colzy����)) = 13 And InStr(.TextMatrix(GetRow(13), col��Ժ���), "����") > 0 And InStr(.TextMatrix(i, col��Ժ���), "����") = 0 Then
    '                            .Row = i: .Col = col��Ժ���
    '                            Call ShowMessage(vsDiagZY, "��Ҫ��ϵĳ�Ժ���Ϊ��������������ϵĳ�Ժ���ȴ����""" & .TextMatrix(i, col��Ժ���) & """��")
    '                            Exit Function
                            ElseIf Val(.TextMatrix(i - 1, colzy����)) = 13 And InStr(.TextMatrix(GetRow(13), col��Ժ���), "����") = 0 And InStr(.TextMatrix(i, col��Ժ���), "����") > 0 Then
                                .Row = i: .Col = col��Ժ���
                                Call ShowMessage(vsDiagZY, "��Ҫ��ϵĳ�Ժ�����Ϊ��������������ϵĳ�Ժ���ȴΪ������")
                                Exit Function
                            End If
                            
                            For j = GetRow(13) To .Rows - 1
                                If j <> i And .TextMatrix(j, col�������) <> "" Then
                                    If .TextMatrix(j, col�������) = .TextMatrix(i, col�������) Then
                                        .Row = i: .Col = col�������
                                        Call ShowMessage(vsDiagZY, "���ִ���������ͬ�ĳ�Ժ�����Ϣ��")
                                        Exit Function
                                    ElseIf Val(.TextMatrix(i, colzy����ID)) <> 0 Then
                                        If Val(.TextMatrix(j, colzy����ID)) = Val(.TextMatrix(i, colzy����ID)) Then
                                            .Row = i: .Col = col�������
                                            Call ShowMessage(vsDiagZY, "���ִ���������ͬ�ĳ�Ժ�����Ϣ��")
                                            Exit Function
                                        End If
                                    ElseIf Val(.TextMatrix(i, colzy���ID)) <> 0 Then
                                        '����ҽ��ϴ�֤��,�����޶�Ӧ֤��ID,���ID����ͬ
    '                                    If Val(.TextMatrix(j, colzy���ID)) = Val(.TextMatrix(i, colzy���ID)) Then
    '                                        .Row = i: .Col = col�������
    '                                        Call ShowMessage(vsDiagZY, "���ִ���������ͬ�ĳ�Ժ�����Ϣ��")
    '                                        Exit Function
    '                                    End If
                                    End If
                                End If
                            Next
                        End If
                        
                        If Val(.TextMatrix(i, colzy����ID)) <> 0 Then str����IDs = str����IDs & "," & Val(.TextMatrix(i, colzy����ID))
                        If Val(.TextMatrix(i, colzy���ID)) <> 0 Then str���IDs = str���IDs & "," & Val(.TextMatrix(i, colzy���ID))
                        
                        '�Ƿ�������Ҫ����������
                        If InStr("," & mstr���� & ",", "," & Val(.TextMatrix(i, colzy����)) & ",") > 0 Then
                            blnDiagnose = True
                        End If
                    End If
                Next
            End With
        End If
        
        With vsOPS
            For i = .FixedRows To .Rows - 1
                If Trim(.TextMatrix(i, col��������)) <> "" Then
                    If Not IsDate(.TextMatrix(i, col��������)) Then
                        .Row = i: .Col = col��������
                        Call ShowMessage(vsOPS, "�����������벻��ȷ��")
                        Exit Function
                    ElseIf txtInfo(txt��Ժʱ��).Text <> "" And Format(.TextMatrix(i, col��������), "yyyy-MM-dd") > Format(txtInfo(txt��Ժʱ��).Text, "yyyy-MM-dd") Or _
                        Format(.TextMatrix(i, col��������), "yyyy-MM-dd") < Format(txtInfo(txt��Ժʱ��).Text, "yyyy-MM-dd") Then
                        .Row = i: .Col = col��������    '��������û�о�ȷ��ʱ��
                        Call ShowMessage(vsOPS, "�������ڲ������Ժ���ڷ�Χ�ڡ�")
                        Exit Function
                    End If
                    If zlCommFun.ActualLen(.TextMatrix(i, col��������)) > 100 Then
                        .Row = i: .Col = col��������
                        Call ShowMessage(vsOPS, "������������̫����ֻ����100���ַ���50�����֡�")
                        Exit Function
                    End If
                    If .ColHidden(col������ʿ) Then
                        If .TextMatrix(i, col����ҽʦ) = "" Then
                            .Row = i: .Col = col����ҽʦ
                            Call ShowMessage(vsOPS, "����������ҽʦ��")
                            Exit Function
                        End If
                    Else
                        If .TextMatrix(i, col����ҽʦ) = "" And .TextMatrix(i, col������ʿ) = "" Then
                            .Row = i: .Col = col����ҽʦ
                            Call ShowMessage(vsOPS, "����������ҽʦ��������ʿ��")
                            Exit Function
                        End If
                    End If
                End If
            Next
        End With
            
        With vsKSS
            For i = .FixedRows To .Rows - 1
                If i > .FixedRows Then
                    If Trim(.TextMatrix(i, 1)) <> "" And Trim(.TextMatrix(i - 1, 1)) = "" Then
                        .Row = i - 1: .Col = 1
                        Call ShowMessage(vsKSS, "���������뿹��ҩ�����ݡ�")
                        Exit Function
                    End If
                End If
                If Trim(.TextMatrix(i, 1)) <> "" Then
                    For j = .FixedRows To .Rows - 1
                        If j <> i And Trim(.TextMatrix(j, 1)) = Trim(.TextMatrix(i, 1)) Then
                            .Row = j: .Col = 1
                            Call ShowMessage(vsKSS, "���ִ���������ͬ�Ŀ���ҩ����Ϣ��")
                            Exit Function
                        End If
                    Next
                End If
            Next
        End With
        
        With vsTSJC
            For i = .FixedRows To .Rows - 1
                If i > .FixedRows Then
                    If Trim(.TextMatrix(i, 1)) <> "" And Trim(.TextMatrix(i - 1, 1)) = "" Then
                        .Row = i - 1: .Col = 1
                        Call ShowMessage(vsTSJC, "�������������������ݡ�")
                        Exit Function
                    End If
                End If
                If Trim(.TextMatrix(i, 1)) <> "" Then
                    For j = .FixedRows To .Rows - 1
                        If j <> i And Trim(.TextMatrix(j, 1)) = Trim(.TextMatrix(i, 1)) Then
                            .Row = j: .Col = 1
                            Call ShowMessage(vsTSJC, "���ִ���������ͬ����������Ϣ��")
                            Exit Function
                        End If
                    Next
                End If
            Next
        End With
            
        '����ҩ��
        With vsKSS
        
            For i = .FixedRows To .Rows - 1
                If (Len(.TextMatrix(i, kssʹ������)) > 18 Or Val(.TextMatrix(i, kssʹ������)) = 0) And Len(.TextMatrix(i, kssʹ������)) > 0 Then
                    .Row = i: .Col = kssʹ������
                    Call ShowMessage(vsKSS, "����дʮ��λ�����ڵ�����������")
                    Exit Function
                End If
                If LenB(StrConv(.TextMatrix(i, kss��ҩĿ��), vbFromUnicode)) > 200 And LenB(StrConv(.TextMatrix(i, kss��ҩĿ��), vbFromUnicode)) > 0 Then
                    .Row = i: .Col = kss��ҩĿ��
                    Call ShowMessage(vsKSS, "����д100���������ڵ���ҩĿ�ġ�")
                    Exit Function
                End If
            Next
        
        End With
        
         
        mstr����ID = Mid(str����IDs, 2)
        mstr���ID = Mid(str���IDs, 2)
    End If
    '������鲡���Ƿ��Ŀ����ҳ��������״̬
    If Not CheckMecRed(mlng����ID, mlng��ҳID, Me.Caption, "�޸���ҳ") Then Exit Function
         
    CheckPageData = True
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub LoadOldData(strOld As String, Optional ByVal lngIndex As Long)
'����:�����ݿ��б�������䰴���Ƶĸ�ʽ���ص�����
'����:lngIndex-�����������������ΪӤ���ĺͲ��˱����
    Dim strTmp As Long
    Dim lng��λ As Long

    If lngIndex = 0 Then lngIndex = txt����: lng��λ = cbo���䵥λ
    If lngIndex = txtӤ������ Then lng��λ = cboӤ�����䵥λ
    If InStr(strOld, "��") > 0 And lngIndex = txt���� Then
        If InStr(strOld, "��") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "��") - 1)
            txtInfo(lngIndex).Text = strTmp
            If cboinfo(lng��λ).ListCount > 0 Then cboinfo(lng��λ).ListIndex = 0
        Else
            txtInfo(lngIndex).Text = strOld
            cboinfo(lng��λ).ListIndex = -1
        End If
    ElseIf InStr(strOld, "��") > 0 Then
        If InStr(strOld, "��") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "��") - 1)
            txtInfo(lngIndex).Text = strTmp
            If cboinfo(lng��λ).ListCount > 1 Then
                cboinfo(lng��λ).ListIndex = IIf(lngIndex = txt����, 1, 0)
            End If
        Else
            txtInfo(lngIndex).Text = strOld
            cboinfo(lng��λ).ListIndex = -1
        End If
    ElseIf InStr(strOld, "��") > 0 Then
        If InStr(strOld, "��") = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "��") - 1)
            txtInfo(lngIndex).Text = strTmp
            If cboinfo(lng��λ).ListCount > 1 Then cboinfo(lng��λ).ListIndex = IIf(lngIndex = txt����, 2, 1)
        Else
            txtInfo(lngIndex).Text = strOld
            cboinfo(lng��λ).ListIndex = -1
        End If
    ElseIf InStr(strOld, "Сʱ") > 0 Then
        If InStr(strOld, "Сʱ") + 1 = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "Сʱ") - 1)
            txtInfo(lngIndex).Text = strTmp
            If cboinfo(lng��λ).ListCount > 1 Then cboinfo(lng��λ).ListIndex = IIf(lngIndex = txt����, 3, 2)
        Else
            txtInfo(lngIndex).Text = strOld
            cboinfo(lng��λ).ListIndex = -1
        End If
    ElseIf InStr(strOld, "����") > 0 Then
        If InStr(strOld, "����") + 1 = Len(strOld) Then
            strTmp = Mid(strOld, 1, InStr(strOld, "����") - 1)
            txtInfo(lngIndex).Text = strTmp
            If cboinfo(lng��λ).ListCount > 1 Then cboinfo(lng��λ).ListIndex = IIf(lngIndex = txt����, 4, 3)
        Else
            txtInfo(lngIndex).Text = strOld
            cboinfo(lng��λ).ListIndex = -1
        End If
    ElseIf IsNumeric(strOld) Then
        txtInfo(lngIndex).Text = strOld
        If cboinfo(lng��λ).ListCount > 0 Then cboinfo(lng��λ).ListIndex = 0
    Else
        txtInfo(lngIndex).Text = strOld
        cboinfo(lng��λ).ListIndex = -1
    End If
End Sub

Private Function SavePageData(ByVal blnBeforSign As Boolean) As Boolean
'���ܣ����没����ҳ����
'������blnBeforSign-�Ƿ�ǩ��ʱ����ǰ����
    Dim arrSQL() As Variant, i As Long
    Dim strȷ������ As String, str�����־ As String
    Dim arrFieldҽ��() As Variant, arrValueҽ��() As Variant
    Dim arrField��ʿ() As Variant, arrValue��ʿ() As Variant
    Dim strת�ƿ��� As String, curDate As Date
    Dim str�п� As String, str���� As String
    Dim lng�����ּ� As Long
    Dim intIdx As Integer, str���� As String, str�������� As String
    Dim lng��λID As Long, str������� As String
    Dim str��Ⱦ���� As String
    Dim str��Ժȥ�� As String
    Dim str�����¼� As String
    Dim str����״�� As String
    Dim str���� As String
    Dim ArrDel As Variant
    Dim blnIsYCcheck As Boolean
    Dim blnIsDDcheck As Boolean
    Dim StrSQL As String
    Dim strTemp As String
    Dim lngCol As Long
    Dim lngRow As Long
    Dim blnTrans As Boolean, blnDiagChange As Boolean
    Dim strFilter As String, strTmp As String
    
    arrSQL = Array()
    arrFieldҽ�� = Array()
    arrField��ʿ = Array()
    
    '������ҳ�ӱ�
    
    If txtInfo(txtת��1).Text <> "" Then
        strת�ƿ��� = txtInfo(txtת��1).Text
        If txtInfo(txtת��2).Text <> "" Then
            strת�ƿ��� = strת�ƿ��� & "," & txtInfo(txtת��2).Text
            If txtInfo(txtת��3).Text <> "" Then
                strת�ƿ��� = strת�ƿ��� & "," & txtInfo(txtת��3).Text
            End If
        End If
    End If
    str�������� = Trim(cboinfo(cbo��������).Text)
    If InStr(str��������, "-") > 0 Then  '����ǹ淶�����ݣ���ֻ�����
        str�������� = Mid(str��������, 1, InStr(str��������, "-") - 1)
    End If
    
    '��Ⱦ����
    For i = 0 To lstInfection.ListCount - 1
        If lstInfection.Selected(i) = True Then
            str��Ⱦ���� = str��Ⱦ���� & IIf(i <> 0, ",", "") & lstInfection.ItemData(i)
        End If
    Next
    '�����¼�
    For i = 0 To lstAdvEvent.ListCount - 1
        If lstAdvEvent.Selected(i) = True Then
            str�����¼� = str�����¼� & IIf(i <> 0, ",", "") & lstAdvEvent.ItemData(i)
            If lstAdvEvent.List(i) = "ѹ��" Then blnIsYCcheck = True
            If lstAdvEvent.List(i) = "ҽԺ�ڵ���/׹��" Then blnIsDDcheck = True
        End If
    Next
    
    '��Ժȥ��
    str��Ժȥ�� = cboinfo(cbo��Ժ��ʽ).Text
    
    '����ʱ��
    str���� = ""
    If IsDate(txt��������.Text) Then
        If IsDate(txt����ʱ��.Text) Then
            str���� = txt��������.Text & " " & txt����ʱ��.Text
        Else
            str���� = txt��������.Text
        End If
    End If
    '����״��
    If cboinfo(cbo����״��).ListIndex > 0 Then
        str����״�� = Mid(cboinfo(cbo����״��), 1, InStr(cboinfo(cbo����״��), "-") - 1)
    End If
    '��ʿվ����ʱ���滤ʿ��Ϣ���֣�����ҽ��վ������ʱ��ʿ��Ϣ��
    If mbln��ʿվ And mblnҽ����ʿ������ҳ Or Not mbln��ʿվ And Not mblnҽ����ʿ������ҳ Then
        arrField��ʿ = Array("�����¼�", "ѹ�������ڼ�", "ѹ������", "������׹���˺�", "������׹��ԭ��", _
                    "����Լ��", "Լ����ʱ��", "Լ����ʽ", "Լ������", "Լ��ԭ��")
        arrValue��ʿ = Array(str�����¼�, IIf(blnIsYCcheck, cboinfo(cboѹ�������ڼ�).Text, ""), IIf(blnIsYCcheck, cboinfo(cboѹ������).Text, ""), _
                    IIf(blnIsDDcheck, cboinfo(cbo������׹���˺�).Text, ""), IIf(blnIsDDcheck, cboinfo(cbo������׹��ԭ��).Text, ""), _
                    chkInfo(chk�Ƿ�ʹ������Լ��).Value, txtInfo(txtԼ����ʱ��).Text, NeedName(cboinfo(cboԼ����ʽ).Text), _
                    NeedName(cboinfo(cboԼ������).Text), NeedName(cboinfo(cboԼ��ԭ��).Text))
    End If
    'ҽ��վ����ҽ��������Ϣ����
    If Not mbln��ʿվ Then
        arrFieldҽ�� = Array("��Ժ����", "��Ժ����", "ת�Ƽ�¼", "HBsAg", "HCV-Ab", "HIV-Ab", _
                    "��ҽΣ��", "��ҽ��֢", "��ҽ����", "��ҽ���ȷ���", "������ҩ�Ƽ�", "��������ԭ��", "����ʱ��", _
                    "��Ժǰ����Ժ����", "ʾ�̲���", "���в���", "���Ѳ���", "Rh", "��Ѫ��Ӧ", "���ϸ��", "��ѪС��", "��Ѫ��", "��ȫѪ", "������", _
                    "��Һ��Ӧ", "������", "����ҽʦ", "����ҽʦ", "����ҽʦ", "�о���ʵϰҽʦ", "ʵϰҽʦ", _
                    "�ʿ�ҽʦ", "�ʿػ�ʿ", "��ԭѧ���", "��Ѫ���", "CT", "MRI", "��ɫ������", "������4", "������5", "������6", _
                    "��������", "��Ⱦ����", "��Ժ��ʽ", "��Ժת��", _
                    "����Ժ�ƻ�����", "31������סԺ", "������������", "��������������", "��������Ժ����", "������ʹ��ʱ��", "����ʱ��", "���Ȳ���", _
                    "�������", "�ֻ��̶�", "����������", "��ҽ�豸", "��ҽ����", "��֤ʩ��", "�����", "����", "��������", "��ҳ��������", _
                    "�没�ز�Σ", "�ٴ�·��", "�˳�ԭ��", "����ԭ��", _
                    "��������Ժ��ʽ", "Χ��������", "�������", "����״��", "����ʱ��", "ҽѧ��ʾ", "����ҽѧ��ʾ")
                
    
    
        arrValueҽ�� = Array(txtInfo(txt��Ժ����).Text, txtInfo(txt��Ժ����).Text, strת�ƿ���, _
                    NeedName(cboinfo(cboHBsAg).Text), NeedName(cboinfo(cboHCVAb).Text), NeedName(cboinfo(cboHIVAb).Text), _
                    chkInfo(chkΣ��).Value, chkInfo(chk��֢).Value, chkInfo(chk����).Value, NeedName(cboinfo(cbo���ȷ���).Text), _
                    NeedName(cboinfo(cbo������ҩ).Text), txtInfo(txt����ԭ��).Text, IIf(txt����ʱ��.Text = "____-__-__ __:__:__", "", txt����ʱ��.Text), chkInfo(chk����Ժ����).Value, chkInfo(chkʾ�̲���).Value, _
                    chkInfo(chk���в���).Value, chkInfo(chk���Ѳ���).Value, NeedName(cboinfo(cboRh).Text), IIf(cboinfo(cbo��Ѫ��Ӧ).ListIndex = -1, "", cboinfo(cbo��Ѫ��Ӧ).ListIndex), _
                    txtInfo(txt���ϸ��).Text, txtInfo(txt��ѪС��).Text, txtInfo(txt��Ѫ��).Text, _
                    txtInfo(txt��ȫѪ).Text, txtInfo(txt������).Text, NeedName(cboinfo(cbo��Һ��Ӧ).Text), _
                    NeedName(cboinfo(cbo������).Text), NeedName(cboinfo(cbo����ҽʦ).Text), _
                    NeedName(cboinfo(cbo����ҽʦ).Text), NeedName(cboinfo(cbo����ҽʦ).Text), _
                    NeedName(cboinfo(cbo�о���ҽʦ).Text), NeedName(cboinfo(cboʵϰҽʦ).Text), _
                    NeedName(cboinfo(cbo�ʿ�ҽʦ).Text), NeedName(cboinfo(cbo�ʿػ�ʿ).Text), _
                    chkInfo(chk��ԭѧ).Value, NeedName(cboinfo(cbo��Ѫ���).Text), _
                    chkInfo(chkCT).Value, chkInfo(chkMRI).Value, chkInfo(chk������).Value, _
                    vsTSJC.TextMatrix(vsTSJC.FixedRows + 0, 1), vsTSJC.TextMatrix(vsTSJC.FixedRows + 1, 1), _
                    vsTSJC.TextMatrix(vsTSJC.FixedRows + 2, 1), str��������, str��Ⱦ����, str��Ժȥ��, IIf(cboinfo(cbo��Ժ��ʽ).Text = "תԺ" Or cboinfo(cbo��Ժ��ʽ).Text = "ת����", txtInfo(txt��Ժת��).Text, ""), _
                    cboinfo(cbo31���7������Ժ).ListIndex, IIf(txtInfo(txt31��Ŀ��).Enabled, Trim(txtInfo(txt31��Ŀ��).Text), ""), _
                    IIf(Trim(txtInfo(txtӤ������).Text) <> "", txtInfo(txtӤ������).Text & IIf(cboinfo(cboӤ�����䵥λ).Visible, cboinfo(cboӤ�����䵥λ).Text, ""), ""), txtInfo(txt����������).Text, txtInfo(txt��������Ժ����).Text, _
                    txtInfo(txt������Сʱ).Text, txtInfo(txt��Ժǰ��).Text & "," & txtInfo(txt��ԺǰСʱ).Text & "," & txtInfo(txt��Ժǰ����).Text & "|" & txtInfo(txt��Ժ����).Text & "," & txtInfo(txt��Ժ��Сʱ).Text & "," & txtInfo(txt��Ժ�����).Text, _
                    txtInfo(txt����ԭ��).Text, txtInfo(txt�������).Text, IIf(cboinfo(cbo�ֻ��̶�).Enabled, NeedName(cboinfo(cbo�ֻ��̶�).Text), ""), IIf(cboinfo(cbo����������).Enabled, NeedName(cboinfo(cbo����������).Text), ""), _
                    NeedName(cboinfo(cboʹ����ҽ�����豸).Text), NeedName(cboinfo(cboʹ����ҽ���Ƽ���).Text), NeedName(cboinfo(cbo��֤ʩ��).Text), IIf(txtInfo(txt�����).Enabled, txtInfo(txt�����).Text, ""), "", NeedName(cboinfo(cbo��������).Text), txtInfo(txt�ʿ�����).Text _
                    , chkInfo(chkסԺ�ڼ�没�ػ�Σ).Value, chkInfo(chk����·��).Value, IIf(chkInfo(chk���·��).Value = 1, "1", txtInfo(txt�˳�ԭ��).Text), IIf(chkInfo(chk����).Value = 0, "0", IIf(txtInfo(txt����ԭ��).Text = "", " ", txtInfo(txt����ԭ��).Text)) _
                    , NeedName(cboinfo(cbo��������Ժ��ʽ).Text), chkInfo(chkΧ��������).Value, chkInfo(chk�������).Value _
                    , str����״��, str����, txtInfo(txtҽѧ��ʾ).Text, txtInfo(txt����ҽѧ��ʾ).Text)
    End If

    '��ʿ����
    For i = 0 To UBound(arrField��ʿ)
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_������ҳ�ӱ�_��ҳ����(" & _
            mlng����ID & "," & mlng��ҳID & ",'" & arrField��ʿ(i) & "','" & arrValue��ʿ(i) & "')"
    Next
    'ҽ������
    For i = 0 To UBound(arrFieldҽ��)
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_������ҳ�ӱ�_��ҳ����(" & _
            mlng����ID & "," & mlng��ҳID & ",'" & arrFieldҽ��(i) & "','" & arrValueҽ��(i) & "')"
    Next
    
    If Not mbln��ʿվ Then
        curDate = zlDatabase.Currentdate
        
        If IsDate(txt����ʱ��.Text) Then
            str���� = "To_Date('" & Format(txt��������.Text & " " & txt����ʱ��.Text, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')"
        Else
            str���� = "To_Date('" & Format(txt��������.Text, "yyyy-MM-dd") & "','YYYY-MM-DD')"
        End If
        If Trim(txtInfo(txt��λ����).Text) <> "" Then
            lng��λID = Val(txtInfo(txt��λ����).Tag)
        End If
        
        '������Ϣ
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_������Ϣ_��ҳ����(" & _
            mlng����ID & "," & IIf(txtInfo(txtסԺ��).Text = "", "NULL", "'" & txtInfo(txtסԺ��).Text & "'") & "," & _
            "'" & txtInfo(txt����).Text & "','" & NeedName(cboinfo(cbo�Ա�).Text) & "','" & txtInfo(txt����).Text & cboinfo(cbo���䵥λ).Text & "'," & _
            str���� & ",'" & IIf(mbln���ýṹ����ַ, PatiAddress������.Value, txtInfo(txt�����ص�).Text) & "','" & cboinfo(cbo���֤��).Text & "'," & _
            "'" & NeedName(cboinfo(cbo����).Text) & "','" & NeedName(cboinfo(cbo����).Text) & "','" & txtInfo(txt����).Text & "'," & _
            "'" & NeedName(cboinfo(cbo����).Text) & "','" & NeedName(cboinfo(cboְҵ).Text) & "'," & _
            "'" & NeedName(cboinfo(cbo���ʽ).Text) & "','" & IIf(mbln���ýṹ����ַ, PatiAddress��סַ.Value, txtInfo(txt��ͥ��ַ).Text) & "'," & _
            "'" & txtInfo(txt��ͥ�绰).Text & "','" & txtInfo(txt��ͥ�ʱ�).Text & "'," & _
            "'" & txtInfo(txt��λ����).Text & "','" & txtInfo(txt��λ�绰).Text & "'," & _
            "'" & txtInfo(txt��λ�ʱ�).Text & "','" & txtInfo(txt��ϵ������).Text & "'," & _
            "'" & NeedName(cboinfo(cbo��ϵ�˹�ϵ).Text) & "','" & txtInfo(txt��ϵ�˵绰).Text & "'," & _
            "'" & txtInfo(txt��ϵ�˵�ַ).Text & "',null,null,null,null,null,null,'" & Trim(txtInfo(txt����֤��).Text) & "'," & _
            ZVal(lng��λID) & ",'" & IIf(mbln���ýṹ����ַ, PatiAddress���ڵ�ַ.Value, txtInfo(txt���ڵ�ַ).Text) & "','" & txtInfo(txt�����ʱ�).Text & "','" & IIf(mbln���ýṹ����ַ, PatiAddress����.Value, txtInfo(txt����).Text) & "')"
    
        '�ṹ����ַ
        If mbln���ýṹ����ַ Then
            '������
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            If PatiAddress������.valueʡ <> "" Or PatiAddress������.value�� <> "" Or PatiAddress������.value���� <> "" Then
                arrSQL(UBound(arrSQL)) = "zl_���˵�ַ��Ϣ_update(1," & mlng����ID & "," & mlng��ҳID & ",1,'" & PatiAddress������.valueʡ & "','" & _
                    PatiAddress������.value�� & "','" & PatiAddress������.value���� & "','" & PatiAddress������.value��ϸ��ַ & "')"
            Else
                arrSQL(UBound(arrSQL)) = "zl_���˵�ַ��Ϣ_update(2," & mlng����ID & "," & mlng��ҳID & ",1)"
            End If
            '����
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            If PatiAddress����.valueʡ <> "" Or PatiAddress����.value�� <> "" Or PatiAddress����.value���� <> "" Then
                arrSQL(UBound(arrSQL)) = "zl_���˵�ַ��Ϣ_update(1," & mlng����ID & "," & mlng��ҳID & ",2,'" & PatiAddress����.valueʡ & "','" & _
                    PatiAddress����.value�� & "','" & PatiAddress����.value���� & "','" & PatiAddress����.value��ϸ��ַ & "')"
            Else
                arrSQL(UBound(arrSQL)) = "zl_���˵�ַ��Ϣ_update(2," & mlng����ID & "," & mlng��ҳID & ",2)"
            End If
        End If
    
        '������ҳ
        If IsDate(txtInfo(txtȷ������).Text) Then
            strȷ������ = "To_Date('" & Format(txtInfo(txtȷ������).Text, "yyyy-MM-dd HH:mm") & "','YYYY-MM-DD HH24:MI')"
        Else
            strȷ������ = "NULL"
        End If
        If chkInfo(chk����).Value = 1 Then
            str�����־ = Decode(NeedName(cboinfo(cbo����Ex).Text), "��", 1, "��", 2, "��", 3, "��", 4, "����", 9)
        Else
            str�����־ = "NULL"
        End If
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "ZL_������ҳ_��ҳ����(" & _
            mlng����ID & "," & mlng��ҳID & ",'" & NeedName(cboinfo(cbo����).Text) & "'," & _
            "'" & txtInfo(txt����).Text & cboinfo(cbo���䵥λ).Text & "','" & NeedName(cboinfo(cboְҵ).Text) & "'," & _
            "'" & NeedName(cboinfo(cbo����).Text) & "','" & txtInfo(txt����).Text & "'," & _
            "'" & NeedName(cboinfo(cbo���ʽ).Text) & "','" & IIf(mbln���ýṹ����ַ, PatiAddress��סַ.Value, txtInfo(txt��ͥ��ַ).Text) & "'," & _
            "'" & txtInfo(txt��ͥ�绰).Text & "','" & txtInfo(txt��ͥ�ʱ�).Text & "'," & _
            "'" & txtInfo(txt��λ����).Text & "','" & txtInfo(txt��λ�绰).Text & "'," & _
            "'" & txtInfo(txt��λ�ʱ�).Text & "','" & txtInfo(txt��ϵ������).Text & "'," & _
            "'" & NeedName(cboinfo(cbo��ϵ�˹�ϵ).Text) & "','" & txtInfo(txt��ϵ�˵绰).Text & "'," & _
            "'" & txtInfo(txt��ϵ�˵�ַ).Text & "','" & NeedName(cboinfo(cbo��Ժ����).Text) & "'," & _
            "'" & chkInfo(chk�Ƿ�ȷ��).Value & "'," & strȷ������ & "," & _
            IIf(Val(txtInfo(txt���ȴ���).Text) <> 0, Val(txtInfo(txt���ȴ���).Text), "NULL") & "," & _
            IIf(Val(txtInfo(txt�ɹ�����).Text) <> 0, Val(txtInfo(txt�ɹ�����).Text), "NULL") & "," & _
            IIf(chkInfo(chkʬ��).Enabled, chkInfo(chkʬ��).Value, "NULL") & "," & _
            str�����־ & "," & IIf(Val(txtInfo(txt��������).Text) <> 0, Val(txtInfo(txt��������).Text), "NULL") & "," & _
            "'" & NeedName(cboinfo(cboѪ��).Text) & "','" & NeedName(cboinfo(cbo����ҽʦ).Text) & "'," & _
            "'" & NeedName(cboinfo(cboסԺҽʦ).Text) & "','" & NeedName(cboinfo(cbo����ҽʦ).Text) & "','" & NeedName(cboinfo(cbo����ҽʦ).Text) & _
            "','" & UserInfo.��� & "','" & UserInfo.���� & "'," & chkInfo(chk�·�����).Value & "," & _
            "'" & NeedName(cboinfo(cbo�������).Text) & "'," & chkInfo(chk����Ժ).Value & "," & Val(txtInfo(txt���).Text) & "," & Val(txtInfo(txt����).Text) & ",'" & _
            NeedName(cboinfo(cbo��Ժ��ʽ).Text) & "','" & NeedName(cboinfo(cbo��Ժ��ʽ).Text) & "','" & NeedName(cboinfo(cbo���λ�ʿ).Text) & "','" & IIf(mbln���ýṹ����ַ, PatiAddress���ڵ�ַ.Value, txtInfo(txt���ڵ�ַ).Text) & "','" & txtInfo(txt�����ʱ�).Text & "')"
    
        '�ṹ����ַ
        If mbln���ýṹ����ַ Then
            '��סַ
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            If PatiAddress��סַ.valueʡ <> "" Or PatiAddress��סַ.value�� <> "" Or PatiAddress��סַ.value���� <> "" Then
                arrSQL(UBound(arrSQL)) = "zl_���˵�ַ��Ϣ_update(1," & mlng����ID & "," & mlng��ҳID & ",3,'" & PatiAddress��סַ.valueʡ & "','" & _
                    PatiAddress��סַ.value�� & "','" & PatiAddress��סַ.value���� & "','" & PatiAddress��סַ.value��ϸ��ַ & "')"
            Else
                arrSQL(UBound(arrSQL)) = "zl_���˵�ַ��Ϣ_update(2," & mlng����ID & "," & mlng��ҳID & ",3)"
            End If
            '���ڵ�ַ
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            If PatiAddress���ڵ�ַ.valueʡ <> "" Or PatiAddress���ڵ�ַ.value�� <> "" Or PatiAddress���ڵ�ַ.value���� <> "" Then
                arrSQL(UBound(arrSQL)) = "zl_���˵�ַ��Ϣ_update(1," & mlng����ID & "," & mlng��ҳID & ",4,'" & PatiAddress���ڵ�ַ.valueʡ & "','" & _
                    PatiAddress���ڵ�ַ.value�� & "','" & PatiAddress���ڵ�ַ.value���� & "','" & PatiAddress���ڵ�ַ.value��ϸ��ַ & "')"
            Else
                arrSQL(UBound(arrSQL)) = "zl_���˵�ַ��Ϣ_update(2," & mlng����ID & "," & mlng��ҳID & ",4)"
            End If
        End If
        
        '��Ϸ������
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_��Ϸ������_Insert(" & mlng����ID & "," & mlng��ҳID & ",1," & _
            IIf(cboinfo(cbo�������Ժ).ListIndex = -1, "NULL", cboinfo(cbo�������Ժ).ListIndex) & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_��Ϸ������_Insert(" & mlng����ID & "," & mlng��ҳID & ",2," & _
            IIf(cboinfo(cbo��Ժ���Ժ).ListIndex = -1, "NULL", cboinfo(cbo��Ժ���Ժ).ListIndex) & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_��Ϸ������_Insert(" & mlng����ID & "," & mlng��ҳID & ",3," & _
            IIf(cboinfo(cbo�����벡��).ListIndex = -1, "NULL", cboinfo(cbo�����벡��).ListIndex) & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_��Ϸ������_Insert(" & mlng����ID & "," & mlng��ҳID & ",4," & _
            IIf(cboinfo(cbo�ٴ��벡��).ListIndex = -1, "NULL", cboinfo(cbo�ٴ��벡��).ListIndex) & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_��Ϸ������_Insert(" & mlng����ID & "," & mlng��ҳID & ",5," & _
            IIf(cboinfo(cbo�ٴ���ʬ��).ListIndex = -1, "NULL", cboinfo(cbo�ٴ���ʬ��).ListIndex) & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_��Ϸ������_Insert(" & mlng����ID & "," & mlng��ҳID & ",6," & _
            IIf(cboinfo(cbo��ǰ������).ListIndex = -1, "NULL", cboinfo(cbo��ǰ������).ListIndex) & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_��Ϸ������_Insert(" & mlng����ID & "," & mlng��ҳID & ",7," & _
            IIf(cboinfo(cbo��������Ժ).ListIndex = -1, "NULL", cboinfo(cbo��������Ժ).ListIndex) & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_��Ϸ������_Insert(" & mlng����ID & "," & mlng��ҳID & ",11," & _
            IIf(cboinfo(cbo��ҽ�������Ժ).ListIndex = -1, "NULL", cboinfo(cbo��ҽ�������Ժ).ListIndex) & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_��Ϸ������_Insert(" & mlng����ID & "," & mlng��ҳID & ",12," & _
            IIf(cboinfo(cbo��ҽ��Ժ���Ժ).ListIndex = -1, "NULL", cboinfo(cbo��ҽ��Ժ���Ժ).ListIndex) & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_��Ϸ������_Insert(" & mlng����ID & "," & mlng��ҳID & ",13," & _
            IIf(cboinfo(cbo��֤).ListIndex = -1, "NULL", cboinfo(cbo��֤).ListIndex) & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_��Ϸ������_Insert(" & mlng����ID & "," & mlng��ҳID & ",14," & _
            IIf(cboinfo(cbo�η�).ListIndex = -1, "NULL", cboinfo(cbo�η�).ListIndex) & ")"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_��Ϸ������_Insert(" & mlng����ID & "," & mlng��ҳID & ",15," & _
            IIf(cboinfo(cbo��ҩ).ListIndex = -1, "NULL", cboinfo(cbo��ҩ).ListIndex) & ")"
        '������Ϣ�ӱ�
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "zl_������Ϣ�ӱ�_Update(" & mlng����ID & ",'Ѫ��','" & IIf(InStr(";A��;B��;O��;AB��;����;", ";" & NeedName(cboinfo(cboѪ��).Text) & ";") > 0, NeedName(cboinfo(cboѪ��).Text), "") & "')"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "zl_������Ϣ�ӱ�_Update(" & mlng����ID & ",'RH','" & NeedName(cboinfo(cboRh).Text) & "')"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "zl_������Ϣ�ӱ�_Update(" & mlng����ID & ",'ҽѧ��ʾ','" & txtInfo(txtҽѧ��ʾ) & "')"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "zl_������Ϣ�ӱ�_Update(" & mlng����ID & ",'����ҽѧ��ʾ','" & txtInfo(txt����ҽѧ��ʾ) & "')"
        
        '����ҩ��
        If vsAller.Tag = "" Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "zl_���˹�����¼_Delete(" & mlng����ID & "," & mlng��ҳID & ",3)"
            With vsAller
                For i = .FixedRows To .Rows - 1
                    If Trim(.TextMatrix(i, AC_����ҩ��)) <> "" Then
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = _
                            "zl_���˹�����¼_Insert(" & mlng����ID & "," & mlng��ҳID & "," & _
                            "3," & ZVal(.RowData(i)) & ",'" & .TextMatrix(i, AC_����ҩ��) & "',1," & _
                            "To_Date('" & .Cell(flexcpData, i, AC_����ʱ��) & "','YYYY-MM-DD HH24:MI:SS')," & _
                            "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),'" & .TextMatrix(i, AC_������Ӧ) & "')"
                    End If
                Next
            End With
        End If
        
        '��ҽ���
        If vsDiagXY.Tag = "" Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_������ϼ�¼_DELETE(" & mlng����ID & "," & mlng��ҳID & ",3,NULL,'1,2,3,5,6,7,10')"
            With vsDiagXY
                intIdx = 0
                For i = .FixedRows To .Rows - 1
                    If Trim(.TextMatrix(i, col�������)) <> "" Then
                        If Trim(.TextMatrix(i, col��ϱ���)) = "" Then
                            str������� = .TextMatrix(i, col�������)
                        Else
                            str������� = "(" & .TextMatrix(i, col��ϱ���) & ")" & .TextMatrix(i, col�������)
                        End If
                        blnDiagChange = True
                        If Val(.Cell(flexcpData, i, col�Ƿ�����) & "") > 0 Then
                            strFilter = "�������=" & Val(.TextMatrix(i, col����)) & " And ��¼��Դ=3 And ����id=" & ZVal(.TextMatrix(i, col����ID)) & " And ���id=" & ZVal(.TextMatrix(i, col���ID))

                            strTmp = IIf(str������� = "", "Null", "'" & str������� & "'")
                            strFilter = strFilter & " And �������= " & strTmp

                            strTmp = .TextMatrix(i, col��Ժ����)
                            strTmp = IIf(strTmp = "", "Null", "'" & strTmp & "'")
                            strFilter = strFilter & " And  ��Ժ����= " & strTmp

                            strTmp = NeedName(.TextMatrix(i, col��Ժ���))
                            strTmp = IIf(strTmp = "", "Null", "'" & strTmp & "'")
                            strFilter = strFilter & " And  ��Ժ���= " & strTmp
                            
                            strTmp = .TextMatrix(i, col��ע)
                            strTmp = IIf(strTmp = "", "Null", "'" & strTmp & "'")
                            strFilter = strFilter & " And  ��ע= " & strTmp
                            
                            strFilter = strFilter & " And �Ƿ�δ��=" & IIf(.TextMatrix(i, col�Ƿ�δ��) = "", 0, 1) & " And �Ƿ�����=" & IIf(.TextMatrix(i, col�Ƿ�����) = "", 0, 1)
                            mrsXYDiag.Filter = strFilter
                            blnDiagChange = mrsXYDiag.EOF
                        End If
                        
                        
                        If Val(.TextMatrix(i, col����)) <> Val(.TextMatrix(i - 1, col����)) Then intIdx = 0
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1): intIdx = intIdx + 1
                        If mblnChange Then
                            arrSQL(UBound(arrSQL)) = "ZL_������ϼ�¼_INSERT(" & mlng����ID & "," & mlng��ҳID & ",3,NULL," & _
                                Val(.TextMatrix(i, col����)) & "," & ZVal(.TextMatrix(i, col����ID)) & "," & ZVal(.TextMatrix(i, col���ID)) & "," & _
                                "NULL,'" & str������� & "','" & NeedName(.TextMatrix(i, col��Ժ���)) & "'," & _
                                IIf(.TextMatrix(i, col�Ƿ�δ��) = "", 0, 1) & "," & IIf(.TextMatrix(i, col�Ƿ�����) = "", 0, 1) & "," & _
                                "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                "Null," & intIdx & ",'" & .TextMatrix(i, col��ע) & "','" & .TextMatrix(i, col��Ժ����) & "',Null,'" & UserInfo.���� & "')"
                        Else
                            arrSQL(UBound(arrSQL)) = "ZL_������ϼ�¼_INSERT(" & mlng����ID & "," & mlng��ҳID & ",3,NULL," & _
                                Val(.TextMatrix(i, col����)) & "," & ZVal(.TextMatrix(i, col����ID)) & "," & ZVal(.TextMatrix(i, col���ID)) & "," & _
                                "NULL,'" & str������� & "','" & NeedName(.TextMatrix(i, col��Ժ���)) & "'," & _
                                IIf(.TextMatrix(i, col�Ƿ�δ��) = "", 0, 1) & "," & IIf(.TextMatrix(i, col�Ƿ�����) = "", 0, 1) & "," & _
                                "To_Date('" & Format(CDate(mrsXYDiag!��¼����), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                "Null," & intIdx & ",'" & .TextMatrix(i, col��ע) & "','" & .TextMatrix(i, col��Ժ����) & "',Null,'" & mrsXYDiag!��¼�� & "')"
                        End If
                        If Val(.TextMatrix(i, col����)) = 2 And intIdx = 1 Then mblnDiagChange = mstrXYDiagInfo <> str�������
                    End If
                Next
            End With
        End If
        
        '��ԭѧ���
        If txtInfo(txt��ԭѧ).Enabled Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_������ϼ�¼_DELETE(" & mlng����ID & "," & mlng��ҳID & ",3,NULL,'21')"
            If txtInfo(txt��ԭѧ).Text <> "" Then
                blnDiagChange = True
                If Not mrsXYDiag Is Nothing Then
                    strFilter = "�������=21 And ��¼��Դ=3 And ����id=" & ZVal(cmdInfo(txt��ԭѧ).Tag)
                    strTmp = IIf(txtInfo(txt��ԭѧ).Text = "", "Null", "'" & txtInfo(txt��ԭѧ).Text & "'")
                    strFilter = strFilter & " And �������= " & strTmp
                    
                    mrsXYDiag.Filter = strFilter
                    blnDiagChange = mrsXYDiag.EOF
                End If
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                If blnDiagChange Then
                    arrSQL(UBound(arrSQL)) = "ZL_������ϼ�¼_INSERT(" & mlng����ID & "," & mlng��ҳID & ",3,NULL,21," & _
                        ZVal(cmdInfo(txt��ԭѧ).Tag) & ",NULL,NULL,'" & txtInfo(txt��ԭѧ).Text & "',NULL,NULL,NULL," & _
                        "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),Null,1,Null,Null,Null,'" & UserInfo.���� & "')"
                Else
                    arrSQL(UBound(arrSQL)) = "ZL_������ϼ�¼_INSERT(" & mlng����ID & "," & mlng��ҳID & ",3,NULL,21," & _
                        ZVal(cmdInfo(txt��ԭѧ).Tag) & ",NULL,NULL,'" & txtInfo(txt��ԭѧ).Text & "',NULL,NULL,NULL," & _
                        "To_Date('" & Format(CDate(mrsXYDiag!��¼����), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),Null,1,Null,Null,Null, '" & mrsXYDiag!��¼�� & "')"
                End If
            End If
        End If
        
        '��ҽ���
        If mbln��ҽ And vsDiagZY.Tag = "" Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_������ϼ�¼_DELETE(" & mlng����ID & "," & mlng��ҳID & ",3,NULL,'11,12,13')"
            With vsDiagZY
                intIdx = 0
                For i = .FixedRows To .Rows - 1
                    If Trim(.TextMatrix(i, col�������)) <> "" Then
                        If Trim(.TextMatrix(i, col��ϱ���)) = "" Then
                            str������� = .TextMatrix(i, col�������) & IIf(.TextMatrix(i, col��ҽ֤��) <> "", "(" & .TextMatrix(i, col��ҽ֤��) & ")", "")
                        Else
                            str������� = "(" & .TextMatrix(i, col��ϱ���) & ")" & .TextMatrix(i, col�������) & IIf(.TextMatrix(i, col��ҽ֤��) <> "", "(" & .TextMatrix(i, col��ҽ֤��) & ")", "")
                        End If
                        blnDiagChange = True
                        If Val(.Cell(flexcpData, i, col�Ƿ�����) & "") > 0 Then
                            strFilter = "�������=" & Val(.TextMatrix(i, colzy����)) & " And ��¼��Դ=3 And ����id=" & ZVal(.TextMatrix(i, colzy����ID)) & _
                                        " And ���id=" & ZVal(.TextMatrix(i, colzy���ID)) & " And ֤��ID=" & ZVal(.TextMatrix(i, colzy֤��ID))

                            strTmp = IIf(str������� = "", "Null", "'" & str������� & "'")
                            strFilter = strFilter & " And �������= " & strTmp

                            strTmp = .TextMatrix(i, col��Ժ����)
                            strTmp = IIf(strTmp = "", "Null", "'" & strTmp & "'")
                            strFilter = strFilter & " And  ��Ժ����= " & strTmp

                            strTmp = NeedName(.TextMatrix(i, col��Ժ���))
                            strTmp = IIf(strTmp = "", "Null", "'" & strTmp & "'")
                            strFilter = strFilter & " And  ��Ժ���= " & strTmp
                                                        
                            strTmp = .TextMatrix(i, col��ע)
                            strTmp = IIf(strTmp = "", "Null", "'" & strTmp & "'")
                            strFilter = strFilter & " And  ��ע= " & strTmp
                            
                            mrsZYDiag.Filter = strFilter
                            blnDiagChange = mrsZYDiag.EOF
                        End If
                        
                        If Val(.TextMatrix(i, colzy����)) <> Val(.TextMatrix(i - 1, colzy����)) Then intIdx = 0
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1): intIdx = intIdx + 1
                        If blnDiagChange Then
                            arrSQL(UBound(arrSQL)) = "ZL_������ϼ�¼_INSERT(" & mlng����ID & "," & mlng��ҳID & ",3,NULL," & _
                                Val(.TextMatrix(i, colzy����)) & "," & ZVal(.TextMatrix(i, colzy����ID)) & "," & ZVal(.TextMatrix(i, colzy���ID)) & "," & _
                                ZVal(.TextMatrix(i, colzy֤��ID)) & ",'" & str������� & "','" & NeedName(.TextMatrix(i, col��Ժ���)) & "'," & _
                                "NULL,NULL,To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                "Null," & intIdx & ",'" & .TextMatrix(i, col��ע) & "','" & .TextMatrix(i, col��Ժ����) & "',Null,'" & UserInfo.���� & "')"
                        Else
                            arrSQL(UBound(arrSQL)) = "ZL_������ϼ�¼_INSERT(" & mlng����ID & "," & mlng��ҳID & ",3,NULL," & _
                                Val(.TextMatrix(i, colzy����)) & "," & ZVal(.TextMatrix(i, colzy����ID)) & "," & ZVal(.TextMatrix(i, colzy���ID)) & "," & _
                                ZVal(.TextMatrix(i, colzy֤��ID)) & ",'" & str������� & "','" & NeedName(.TextMatrix(i, col��Ժ���)) & "'," & _
                                "NULL,NULL,To_Date('" & Format(CDate(mrsZYDiag!��¼����), "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')," & _
                                "Null," & intIdx & ",'" & .TextMatrix(i, col��ע) & "','" & .TextMatrix(i, col��Ժ����) & "',Null,'" & mrsZYDiag!��¼�� & "')"
                        
                        End If
                        If Val(.TextMatrix(i, colzy����)) = 12 And intIdx = 1 Then mblnDiagChange = mstrZYDiagInfo <> str�������
                    End If
                Next
            End With
        End If
        
        '�������
        If vsOPS.Tag = "" Then
            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
            arrSQL(UBound(arrSQL)) = "ZL_���������¼_DELETE(" & mlng����ID & "," & mlng��ҳID & ",3)"
            
            With vsOPS
                For i = .FixedRows To .Rows - 1
                    If Trim(.TextMatrix(i, col��������)) <> "" Then
                        If Trim(.TextMatrix(i, col�п�����)) = "" Then
                            str�п� = "NULL": str���� = "NULL"
                        Else
                            str�п� = "'" & Split(.TextMatrix(i, col�п�����), "/")(0) & "'"
                            str���� = "'" & Split(.TextMatrix(i, col�п�����), "/")(1) & "'"
                        End If
                        If .TextMatrix(i, col��������) = "һ������" Then
                            lng�����ּ� = 1
                        ElseIf .TextMatrix(i, col��������) = "��������" Then
                            lng�����ּ� = 2
                        ElseIf .TextMatrix(i, col��������) = "��������" Then
                            lng�����ּ� = 3
                        ElseIf .TextMatrix(i, col��������) = "�ļ�����" Then
                            lng�����ּ� = 4
                        Else
                            lng�����ּ� = 0
                        End If
                        
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "ZL_���������¼_Insert(" & _
                            zlDatabase.GetNextId("���������¼") & "," & mlng����ID & "," & mlng��ҳID & ",3," & _
                            "To_Date('" & Format(.TextMatrix(i, col��������), "yyyy-MM-dd") & "','YYYY-MM-DD')," & _
                            "NULL,NULL,NULL," & ZVal(.TextMatrix(i, col��������ID)) & "," & ZVal(.TextMatrix(i, col������ĿID)) & "," & _
                            "'" & .TextMatrix(i, col��������) & "','" & .TextMatrix(i, col����ҽʦ) & "','" & .TextMatrix(i, col������ʿ) & "'," & _
                            "'" & .TextMatrix(i, col����1) & "','" & .TextMatrix(i, col����2) & "',NULL,NULL,NULL," & _
                            ZVal(.TextMatrix(i, col����ID)) & ",'" & .TextMatrix(i, col��������) & "',NULL,NULL," & _
                            "'" & .TextMatrix(i, col����ҽʦ) & "',NULL,NULL," & str�п� & "," & str���� & "," & _
                            "To_Date('" & Format(curDate, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS'),'" & _
                            .TextMatrix(i, COL�������.COL�������) & "','" & .TextMatrix(i, colASA�ּ�) & "'," & Abs(Val(.TextMatrix(i, col�ٴ�����))) & ",'" & .TextMatrix(i, colNNIS�ּ�) & "'," & _
                            lng�����ּ� & ",NULL,NULL,NULL,NULL,NULL,NULL,NULL," & .Cell(flexcpChecked, i, colԤ���ÿ���ҩ) & "," & ZVal(Val(.TextMatrix(i, col����ҩ����))) & "," & .Cell(flexcpChecked, i, col��Ԥ�ڵĶ�������) & _
                            "," & .Cell(flexcpChecked, i, col������֢) & "," & .Cell(flexcpChecked, i, col������������) & "," & .Cell(flexcpChecked, i, col��������֢) & "," & .Cell(flexcpChecked, i, col�����Ѫ��Ѫ��) & _
                            "," & .Cell(flexcpChecked, i, col�����˿��ѿ�) & "," & .Cell(flexcpChecked, i, col�������Ѫ˨) & "," & .Cell(flexcpChecked, i, col���������л����) & "," & .Cell(flexcpChecked, i, col�������˥��) & _
                            "," & .Cell(flexcpChecked, i, col�����˨��) & "," & .Cell(flexcpChecked, i, col�����Ѫ֢) & "," & .Cell(flexcpChecked, i, col�����Źؽڹ���) & ")"
                    End If
                Next
            End With
        End If
            
        'ʹ�ÿ����صļ�¼
        If vsKSS.Tag = "" Then
            With vsKSS
                '��ɾ���û��������ļ�¼
                ArrDel = Split(mstrDelete, ",")
                mstrDelete = ""
                For i = 0 To UBound(ArrDel)
                    ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                    arrSQL(UBound(arrSQL)) = "Zl_���˿����ؼ�¼_Update(" & _
                            "2," & mlng����ID & "," & mlng��ҳID & "," & ArrDel(i) & ",'" & .TextMatrix(i, kss����) & "',NULL,NULL,NULL,'" & UserInfo.���� & "',To_Date('" & curDate & "','YYYY-MM-DD HH24:MI:SS'))"
                Next
                '���롢����������޸Ľ����ϵ�����
                For i = 1 To .Rows - 1
                    If Val(.RowData(i) & "") <> 0 Then
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "Zl_���˿����ؼ�¼_Update(" & _
                                "1," & mlng����ID & "," & mlng��ҳID & "," & Val(.RowData(i) & "") & ",'" & .TextMatrix(i, kss����) & "','" & .TextMatrix(i, kss��ҩĿ��) & _
                                "','" & .TextMatrix(i, kssʹ�ý׶�) & "'," & Val(.TextMatrix(i, kssʹ������)) & ",'" & UserInfo.���� & "',To_Date('" & curDate & "','YYYY-MM-DD HH24:MI:SS')" & _
                                "," & IIf(.Cell(flexcpChecked, i, KSSһ���п�Ԥ����) = "", "Null", .Cell(flexcpChecked, i, KSSһ���п�Ԥ����)) & "," & ZVal(.TextMatrix(i, KSSDDD��)) & ",'" & .TextMatrix(i, KSS������ҩ) & "')"
                    End If
                Next
                '��������-������ҳ�ӱ������һ��ɾȥ
                For i = 1 To 10
                    If .FixedRows + i - 1 <= .Rows - 1 Then
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        'ɾ��
                        arrSQL(UBound(arrSQL)) = "ZL_������ҳ�ӱ�_��ҳ����(" & mlng����ID & "," & mlng��ҳID & ",'������" & i & "',NULL)"
                    End If
                Next
            End With
        End If
        
        
        If mbln�������� Then
            '���ƻ���
            '��ɾ����Ϣ
            'Zl_�������Ƽ�¼_Delete
            If vs����.Tag = "" Then
                StrSQL = "Zl_�������Ƽ�¼_Delete("
                '  ����id_In In �������Ƽ�¼.����id%Type,
                StrSQL = StrSQL & "" & mlng����ID & ","
                '  ��ҳid_In In �������Ƽ�¼.��ҳid%Type
                StrSQL = StrSQL & "" & mlng��ҳID & ")"
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = StrSQL
                
                With vs����
                    For i = 1 To .Rows - 1
                        If Trim(.TextMatrix(i, .ColIndex("��ʼ����"))) <> "" And _
                              Val(.Cell(flexcpData, i, .ColIndex("��ѧ���Ʊ���"))) <> 0 Then
                            'Zl_�������Ƽ�¼_Insert
                            StrSQL = "Zl_�������Ƽ�¼_Insert("
                            '  ����id_In   In �������Ƽ�¼.����id%Type,
                            StrSQL = StrSQL & "" & mlng����ID & ","
                            '  ��ҳid_In   In �������Ƽ�¼.��ҳid%Type,
                            StrSQL = StrSQL & "" & mlng��ҳID & ","
                            '  ���_In     In �������Ƽ�¼.���%Type,
                            StrSQL = StrSQL & "" & i & ","
                            '  ����id_In   In �������Ƽ�¼.����id%Type,
                            StrSQL = StrSQL & "" & Val(.Cell(flexcpData, i, .ColIndex("��ѧ���Ʊ���"))) & ","
                            '  ��ʼ����_In In �������Ƽ�¼.��ʼ����%Type,
                            StrSQL = StrSQL & "to_date('" & .TextMatrix(i, .ColIndex("��ʼ����")) & "','yyyy-mm-dd'),"
                            '  ��������_In In �������Ƽ�¼.��������%Type,
                            StrSQL = StrSQL & "to_date('" & .TextMatrix(i, .ColIndex("��������")) & "','yyyy-mm-dd'),"
                            '  �Ƴ���_In   In �������Ƽ�¼.�Ƴ���%Type,
                            StrSQL = StrSQL & "" & Val(.TextMatrix(i, .ColIndex("�Ƴ���"))) & ","
                            '  ����_In     In �������Ƽ�¼.����%Type,
                            StrSQL = StrSQL & "" & Val(.TextMatrix(i, .ColIndex("����"))) & ","
                            '  ���Ʒ���_In In �������Ƽ�¼.���Ʒ���%Type,
                            StrSQL = StrSQL & "'" & Trim(.TextMatrix(i, .ColIndex("���Ʒ���"))) & "',"
                            '  ����Ч��_In In �������Ƽ�¼.����Ч��%Type
                            StrSQL = StrSQL & "'" & Trim(.TextMatrix(i, .ColIndex("����Ч��"))) & "')"
                            
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = StrSQL
                        End If
                    Next
                End With
            End If
            
            If vs����.Tag = "" Then
                '��ɾ����Ϣ
                'Zl_�������Ƽ�¼_Delete
                StrSQL = "Zl_�������Ƽ�¼_Delete("
                '  ����id_In In �������Ƽ�¼.����id%Type,
                StrSQL = StrSQL & "" & mlng����ID & ","
                '  ��ҳid_In In �������Ƽ�¼.��ҳid%Type
                StrSQL = StrSQL & "" & mlng��ҳID & ")"
                
                ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                arrSQL(UBound(arrSQL)) = StrSQL
                With vs����
                    For i = 1 To .Rows - 1
                        If Trim(.TextMatrix(i, .ColIndex("��ʼ����"))) <> "" And _
                              Val(.Cell(flexcpData, i, .ColIndex("�������Ʊ���"))) <> 0 Then
                            'Zl_�������Ƽ�¼_Insert
                            StrSQL = "Zl_�������Ƽ�¼_Insert("
                            '  ����id_In   In �������Ƽ�¼.����id%Type,
                            StrSQL = StrSQL & "" & mlng����ID & ","
                            '  ��ҳid_In   In �������Ƽ�¼.��ҳid%Type,
                            StrSQL = StrSQL & "" & mlng��ҳID & ","
                            '  ���_In     In �������Ƽ�¼.���%Type,
                            StrSQL = StrSQL & "" & i & ","
                            '  ����id_In   In �������Ƽ�¼.����id%Type,
                            StrSQL = StrSQL & "" & Val(.Cell(flexcpData, i, .ColIndex("�������Ʊ���"))) & ","
                            '  ��ʼ����_In In �������Ƽ�¼.��ʼ����%Type,
                            StrSQL = StrSQL & "to_date('" & .TextMatrix(i, .ColIndex("��ʼ����")) & "','yyyy-mm-dd'),"
                            '  ��������_In In �������Ƽ�¼.��������%Type,
                            StrSQL = StrSQL & "to_date('" & .TextMatrix(i, .ColIndex("��������")) & "','yyyy-mm-dd'),"
                            '  ��Ұ��λ_In In �������Ƽ�¼.��Ұ��λ%Type,
                            StrSQL = StrSQL & "'" & Trim(.TextMatrix(i, .ColIndex("��Ұ��λ"))) & "',"
                            '  �������_In In �������Ƽ�¼.�������%Type,
                            StrSQL = StrSQL & "" & Val(.TextMatrix(i, .ColIndex("�������"))) & ","
                            '  �ۼ���_In   In �������Ƽ�¼.�ۼ���%Type,
                            StrSQL = StrSQL & "" & Val(.TextMatrix(i, .ColIndex("�ۼ���"))) & ","
                            '  ����Ч��_In In �������Ƽ�¼.����Ч��%Type
                            StrSQL = StrSQL & "'" & Trim(.TextMatrix(i, .ColIndex("����Ч��"))) & "')"
                            
                            ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                            arrSQL(UBound(arrSQL)) = StrSQL
                        End If
                    Next
                End With
            End If
            
            
        End If
        '����������Ϣ
        If vsfMain.Tag = "" Then
            For lngRow = 1 To vsfMain.Rows - 1
                For lngCol = 0 To vsfMain.Cols - 1 Step 3
                    If vsfMain.TextMatrix(lngRow, lngCol + 2) = "�Ƿ�" Then
                        strTemp = IIf(vsfMain.Cell(flexcpChecked, lngRow, lngCol + 1) = 2, 0, 1)
                    Else
                        strTemp = vsfMain.TextMatrix(lngRow, lngCol + 1)
                    End If
                    If vsfMain.TextMatrix(lngRow, lngCol) <> "" And strTemp <> "" Then
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "ZL_������ҳ�ӱ�_��ҳ����(" & mlng����ID & "," & mlng��ҳID & ",'" & vsfMain.TextMatrix(lngRow, lngCol) & "','" & strTemp & "')"
                    ElseIf vsfMain.TextMatrix(lngRow, lngCol) <> "" And strTemp = "" Then
                        '���˺�:11557:2007/09/14:����
                        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
                        arrSQL(UBound(arrSQL)) = "ZL_������ҳ�ӱ�_��ҳ����(" & mlng����ID & "," & mlng��ҳID & ",'" & vsfMain.TextMatrix(lngRow, lngCol) & "',NULL)"
                    End If
                Next lngCol
            Next lngRow
        End If
        '��֢�໤��¼
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = "Zl_������֢�໤���_Delete(" & mlng����ID & "," & mlng��ҳID & ")"
        
        StrSQL = "Zl_������֢�໤���_Insert("
        StrSQL = StrSQL & "" & mlng����ID & ","
        StrSQL = StrSQL & "" & mlng��ҳID & ","
        StrSQL = StrSQL & "1,"
        StrSQL = StrSQL & "'" & txtInfo(txt��֢�໤��).Text & "',NULL,NULL,null,null,"
        StrSQL = StrSQL & chkInfo(chk�˹������ѳ�).Value & ","
        StrSQL = StrSQL & chkInfo(chk�ط���֢ҽѧ��).Value & ","
        StrSQL = StrSQL & "'" & cboinfo(cbo�ط����ʱ��).Text & "')"
        ReDim Preserve arrSQL(UBound(arrSQL) + 1)
        arrSQL(UBound(arrSQL)) = StrSQL
    End If
    Screen.MousePointer = 11
    On Error GoTo errH
    gcnOracle.BeginTrans: blnTrans = True
    For i = 0 To UBound(arrSQL)
        Call zlDatabase.ExecuteProcedure(CStr(arrSQL(i)), Me.Caption)
    Next
    '����ҽ��������Ϣ�޸Ľӿ�
    If mint���� <> 0 Then
        If Not gclsInsure.ModiPatiSwap(mlng����ID, mlng��ҳID, mint����, "2") Then
            gcnOracle.RollbackTrans: Screen.MousePointer = 0: Exit Function
        End If
    End If
    gcnOracle.CommitTrans: blnTrans = False
    
    
    On Error GoTo 0
    Screen.MousePointer = 0
    mblnChange = False
    SavePageData = True
    
    Exit Function
errH:
    Screen.MousePointer = 0
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsOPS_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Not OPSCellEditable(Row, Col) Then
        Cancel = True
    End If
End Sub

Private Sub SetFaceEditable(ByVal blnReadOnly As Boolean)
'���ܣ����ݵ�ǰ�Ƿ�ֻ�������ý���Ŀɱ༭����
    Dim objControl As Object, blnTmp As Boolean
    Dim bln��ҳ As Boolean, strTypeName As String
    
    bln��ҳ = InStr(mstrPrivs, "��ҳ������Ϣ") = 0
    For Each objControl In Me.Controls
        blnTmp = blnReadOnly
        strTypeName = TypeName(objControl)
        If InStr("TextBox;MaskEdBox;ComboBox;CheckBox;VSFlexGrid;ListBox;OptionButton;CommandButton;DTPicker;PatiAddress", TypeName(objControl)) > 0 Then
            'TabStop=False��ʾ��ǰȷʵ���ɱ༭��
            If TypeName(objControl.Container) = "Frame" And (objControl.TabStop = True Or TypeName(objControl) = "OptionButton" And objControl.TabStop = False) Then
                '��ҳ�ɱ༭
                If Not blnTmp Then
                    If mblnҽ����ʿ������ҳ And mbln��ʿվ Then '��ʿվ������ҳʱֻ����д�����¼�
                        If TypeName(objControl) <> "PatiAddress" Then
                            If objControl.Container.hwnd <> fraAdvEvent.hwnd And InStr("," & chkInfo(chk�Ƿ�ʹ������Լ��).hwnd & "," & txtInfo(txtԼ����ʱ��).hwnd & _
                                    "," & cboinfo(cboԼ����ʽ).hwnd & "," & cboinfo(cboԼ������).hwnd & "," & cboinfo(cboԼ��ԭ��).hwnd & ",", "," & objControl.hwnd & ",") = 0 Then
                                blnTmp = True
                                objControl.TabStop = False
                            End If
                        Else
                            blnTmp = True
                            objControl.TabStop = False
                        End If
                    Else
                        '�ж��û��Ƿ�����ҳ������ϢȨ��
                        If bln��ҳ Then
                            If objControl.Container.hwnd = fraInfo(0).hwnd Then
                                '��Ժʱ��֮��Ĳ��ܿ���
                                If objControl.TabIndex < 85 Then
                                    Select Case strTypeName
                                        Case "TextBox", "ComboBox"
                                            If objControl.Text <> "" Then
                                                blnTmp = True
                                            End If
                                        Case "MaskEdBox"
                                            If objControl.Name = "txt��������" And IsDate(objControl.Text) Then
                                                blnTmp = True
                                            ElseIf objControl.Name = "txt����ʱ��" And objControl.Text <> "__:__" Then
                                                blnTmp = True
                                            End If
                                        Case "CheckBox"
                                            If objControl.Value = 1 Then
                                                blnTmp = True
                                            End If
                                        Case "CommandButton"
                                            If txtInfo(objControl.Index).Text <> "" Then
                                                blnTmp = True
                                            End If
                                        Case "PatiAddress"
                                            If objControl.value���� & objControl.valueʡ & objControl.value�� & objControl.value��ϸ��ַ <> "" Then
                                                blnTmp = True
                                            End If
                                    End Select
                                End If
                            End If
                        End If
                        '��ҽ���Ҳ�ʹ����ҽ������ҳ��Ŀ
                        If mbln��ʹ����ҽ��Ŀ And mbln��ҽ Then
                            If InStr("," & cboinfo(cbo��������).hwnd & "," & cboinfo(cbo��Һ��Ӧ).hwnd & "," & cboinfo(cbo��Ѫ��Ӧ).hwnd & _
                            "," & cboinfo(cbo��Ѫ���).hwnd & "," & cboinfo(cboHBsAg).hwnd & "," & cboinfo(cboHCVAb).hwnd & "," & cboinfo(cboHIVAb).hwnd & _
                             "," & cboinfo(cbo����Ex).hwnd & "," & cboinfo(cbo�о���ҽʦ).hwnd & "," & txtInfo(txt���ϸ��).hwnd & "," & txtInfo(txt������).hwnd & _
                              "," & txtInfo(txt��ȫѪ).hwnd & "," & txtInfo(txt��Ѫ��).hwnd & "," & txtInfo(txt�������).hwnd & "," & txtInfo(txt��ѪС��).hwnd & _
                              "," & txtInfo(txt��������).hwnd & "," & txtInfo(txt������Сʱ).hwnd & "," & chkInfo(chk���в���).hwnd & "," & chkInfo(chkʾ�̲���).hwnd & _
                              "," & chkInfo(chk����).hwnd & ",", "," & objControl.hwnd & ",") > 0 Then
                                blnTmp = True
                                objControl.TabStop = False
                            End If
                        End If
                        If mblnҽ����ʿ������ҳ And Not mbln��ʿվ And TypeName(objControl) <> "PatiAddress" Then 'ҽ��վ������д�����¼�
                            If objControl.Container.hwnd = fraAdvEvent.hwnd Or InStr("," & chkInfo(chk�Ƿ�ʹ������Լ��).hwnd & "," & txtInfo(txtԼ����ʱ��).hwnd & _
                                     "," & cboinfo(cboԼ����ʽ).hwnd & "," & cboinfo(cboԼ������).hwnd & "," & cboinfo(cboԼ��ԭ��).hwnd & ",", "," & objControl.hwnd & ",") > 0 Then
                                blnTmp = True
                                objControl.TabStop = False
                            End If
                        End If
                    End If
                End If
                
                If TypeName(objControl) = "TextBox" And objControl.Enabled Then
                    objControl.BackColor = IIf(blnTmp, vbButtonFace, vbWindowBackground)
                    objControl.Locked = blnTmp
                ElseIf TypeName(objControl) = "MaskEdBox" Then
                    'û��Locked����,��Enabledʵ��
                    objControl.Enabled = Not blnTmp
                    objControl.BackColor = IIf(blnTmp, vbButtonFace, vbWindowBackground)
                ElseIf TypeName(objControl) = "ComboBox" And objControl.Enabled Then
                    If Not ((objControl Is cboinfo(cbo������) Or objControl Is cboinfo(cbo����ҽʦ) _
                        Or objControl Is cboinfo(cbo����ҽʦ) Or objControl Is cboinfo(cboסԺҽʦ)) And Not mbln��ʿվ) Then
                        objControl.BackColor = IIf(blnTmp, vbButtonFace, vbWindowBackground)
                        objControl.Locked = blnTmp
                    End If
                ElseIf TypeName(objControl) = "DTPicker" Then
                    objControl.Enabled = Not blnTmp
                ElseIf TypeName(objControl) = "CheckBox" Then
                    'û��Locked����,��Enabledʵ��
                    objControl.Enabled = Not blnTmp
                ElseIf TypeName(objControl) = "VSFlexGrid" Then
                    'ͬʱע��Ҫ�ڼ�������¼��н���һЩ����
                    objControl.Editable = IIf(blnTmp, flexEDNone, flexEDKbdMouse)
                    objControl.BackColor = IIf(blnTmp, vbButtonFace, vbWindowBackground)
                    objControl.BackColorBkg = IIf(blnTmp, vbButtonFace, vbWindowBackground)
                ElseIf TypeName(objControl) = "ListBox" Then
                    objControl.Enabled = IIf(blnTmp, False, True)
                    objControl.BackColor = IIf(blnTmp, vbButtonFace, vbWindowBackground)
                ElseIf TypeName(objControl) = "CommandButton" Then
                    If objControl.Name = "cmdAutoLoad" Or objControl.Name = "cmdPathLoad" Then
                        objControl.Enabled = IIf(blnTmp, False, True)
                    End If
                ElseIf TypeName(objControl) = "PatiAddress" Then
                    objControl.ControlLock = blnTmp
                End If
            End If
            '"OptionButton"������Enabled�ж�
            If TypeName(objControl) = "OptionButton" And TypeName(objControl.Container) = "Frame" And objControl.Enabled = True Then
                objControl.Enabled = IIf(blnTmp, False, True)
                objControl.BackColor = IIf(blnTmp, vbButtonFace, &H8000000F)
            End If
        End If
    Next
End Sub

Public Function BinToDec(ByVal strBin As String) As Long
'���ܣ��������ƴ�ת��Ϊʮ��������
    Dim i As Byte, X As Long
    
    For i = 1 To Len(strBin)
        X = X * 2 + Val(Mid(strBin, i, 1))
    Next i
    
    BinToDec = X
End Function

Private Function SetSignature(Optional ByVal blnReload As Boolean = True) As Boolean
'���ܣ����ݵ�ǰ���˵�ҽʦ��ǩ�������ȷ��ǩ�����������ݵĿɱ༭��
'���أ������Ƿ���ǩ��ֻ�����ܱ༭
    Static rsTmp As ADODB.Recordset
    Dim intCurr As Integer, intHave As Integer
    Dim StrSQL As String, blnReadOnly As Boolean
    Dim i As Integer
    
    '��ʼ��ǩ����ؽ���
    blnReadOnly = False
    For i = 0 To cmdSign.UBound
        cmdSign(i).Enabled = False: cmdUnSign(i).Enabled = False
    Next
    cboinfo(cbo������).ForeColor = Me.ForeColor: lblInfo(cbo������).ForeColor = Me.ForeColor
    cboinfo(cbo����ҽʦ).ForeColor = Me.ForeColor: lblInfo(cbo����ҽʦ).ForeColor = Me.ForeColor
    cboinfo(cbo����ҽʦ).ForeColor = Me.ForeColor: lblInfo(cbo����ҽʦ).ForeColor = Me.ForeColor
    cboinfo(cboסԺҽʦ).ForeColor = Me.ForeColor: lblInfo(cboסԺҽʦ).ForeColor = Me.ForeColor
    cboinfo(cbo������).Locked = False: cboinfo(cbo������).BackColor = vbWindowBackground
    cboinfo(cbo����ҽʦ).Locked = False: cboinfo(cbo����ҽʦ).BackColor = vbWindowBackground
    cboinfo(cbo����ҽʦ).Locked = False: cboinfo(cbo����ҽʦ).BackColor = vbWindowBackground
    cboinfo(cboסԺҽʦ).Locked = False: cboinfo(cboסԺҽʦ).BackColor = vbWindowBackground
    
    '��ȡ��ǰ��Ա���ǩ������
    If NeedName(cboinfo(cboסԺҽʦ).Text) = UserInfo.���� Then
        '�иü���ǩ��Ȩ�޳�ʼ
        intCurr = 1: cmdSign(cmdסԺҽʦ).Enabled = True: cmdUnSign(cmdסԺҽʦ).Enabled = False
    End If
    If NeedName(cboinfo(cbo����ҽʦ).Text) = UserInfo.���� Then
        intCurr = 2: cmdSign(cmd����ҽʦ).Enabled = True: cmdUnSign(cmd����ҽʦ).Enabled = False
    End If
    If NeedName(cboinfo(cbo����ҽʦ).Text) = UserInfo.���� Then
        intCurr = 3: cmdSign(cmd����ҽʦ).Enabled = True: cmdUnSign(cmd����ҽʦ).Enabled = False
    End If
    If NeedName(cboinfo(cbo������).Text) = UserInfo.���� Then
        intCurr = 4: cmdSign(cmd������).Enabled = True: cmdUnSign(cmd������).Enabled = False
    End If
    
    '��ȡ��ҳ�Ѿ�ǩ����߼���
    If rsTmp Is Nothing Or blnReload Then
        On Error GoTo errH
        Set rsTmp = Nothing
        StrSQL = "Select ��Ϣ��,��Ϣֵ From ������ҳ�ӱ� Where ����ID=[1] And ��ҳID=[2] And ��Ϣֵ is Not Null"
        Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, mlng����ID, mlng��ҳID)
    End If
    rsTmp.Filter = "��Ϣ��='סԺҽʦǩ��'"
    If Not rsTmp.EOF Then
        intHave = 1
        
        '��ǩ������ɫ�ֱ�ʾ
        cboinfo(cboסԺҽʦ).ForeColor = vbBlue: lblInfo(cboסԺҽʦ).ForeColor = vbBlue
        
        'ǩ����ť�ɲ���״̬
        If rsTmp!��Ϣֵ = UserInfo.���� Then
            cmdSign(cmdסԺҽʦ).Enabled = False: cmdUnSign(cmdסԺҽʦ).Enabled = True
        Else '�����ѵ�ǩ������ȡ��
            cmdSign(cmdסԺҽʦ).Enabled = False: cmdUnSign(cmdסԺҽʦ).Enabled = False
        End If
    End If
    rsTmp.Filter = "��Ϣ��='����ҽʦǩ��'"
    If Not rsTmp.EOF Then
        intHave = 2
        
        '��ǩ������ɫ�ֱ�ʾ
        cboinfo(cbo����ҽʦ).ForeColor = vbBlue: lblInfo(cbo����ҽʦ).ForeColor = vbBlue
        
        'ǩ����ť�ɲ���״̬
        If rsTmp!��Ϣֵ = UserInfo.���� Then
            cmdSign(cmd����ҽʦ).Enabled = False: cmdUnSign(cmd����ҽʦ).Enabled = True
        Else '�����ѵ�ǩ������ȡ��
            cmdSign(cmd����ҽʦ).Enabled = False: cmdUnSign(cmd����ҽʦ).Enabled = False
        End If
        
        '�ͼ���ǩ�����ܱ��
        cmdSign(cmdסԺҽʦ).Enabled = False: cmdUnSign(cmdסԺҽʦ).Enabled = False
    End If
    rsTmp.Filter = "��Ϣ��='����ҽʦǩ��'"
    If Not rsTmp.EOF Then
        intHave = 3
        
        '��ǩ������ɫ�ֱ�ʾ
        cboinfo(cbo����ҽʦ).ForeColor = vbBlue: lblInfo(cbo����ҽʦ).ForeColor = vbBlue
        
        'ǩ����ť�ɲ���״̬
        If rsTmp!��Ϣֵ = UserInfo.���� Then
            cmdSign(cmd����ҽʦ).Enabled = False: cmdUnSign(cmd����ҽʦ).Enabled = True
        Else '�����ѵ�ǩ������ȡ��
            cmdSign(cmd����ҽʦ).Enabled = False: cmdUnSign(cmd����ҽʦ).Enabled = False
        End If
        
        '�ͼ���ǩ�����ܱ��
        cmdSign(cmdסԺҽʦ).Enabled = False: cmdUnSign(cmdסԺҽʦ).Enabled = False
        cmdSign(cmd����ҽʦ).Enabled = False: cmdUnSign(cmd����ҽʦ).Enabled = False
    End If
    rsTmp.Filter = "��Ϣ��='������ǩ��'"
    If Not rsTmp.EOF Then
        intHave = 4
        
        '��ǩ������ɫ�ֱ�ʾ
        cboinfo(cbo������).ForeColor = vbBlue
        lblInfo(cbo������).ForeColor = vbBlue
        
        'ǩ����ť�ɲ���״̬
        If rsTmp!��Ϣֵ = UserInfo.���� Then
            cmdSign(cmd������).Enabled = False: cmdUnSign(cmd������).Enabled = True
        Else '�����ѵ�ǩ������ȡ��
            cmdSign(cmd������).Enabled = False: cmdUnSign(cmd������).Enabled = False
        End If
        
        '�ͼ���ǩ�����ܱ��
        cmdSign(cmdסԺҽʦ).Enabled = False: cmdUnSign(cmdסԺҽʦ).Enabled = False
        cmdSign(cmd����ҽʦ).Enabled = False: cmdUnSign(cmd����ҽʦ).Enabled = False
        cmdSign(cmd����ҽʦ).Enabled = False: cmdUnSign(cmd����ҽʦ).Enabled = False
    End If
    If intHave > 0 Then
        '�漰ǩ������������ٸ���,��ȻȨ�޻���
        cboinfo(cbo������).Locked = True: cboinfo(cbo������).BackColor = vbButtonFace
        cboinfo(cbo����ҽʦ).Locked = True: cboinfo(cbo����ҽʦ).BackColor = vbButtonFace
        cboinfo(cbo����ҽʦ).Locked = True: cboinfo(cbo����ҽʦ).BackColor = vbButtonFace
        cboinfo(cboסԺҽʦ).Locked = True: cboinfo(cboסԺҽʦ).BackColor = vbButtonFace
    End If
    
    '�����ǰ��Աǩ�����𲻸�����ǩ�������򲻿ɱ༭
    If intCurr <= intHave And intHave > 0 Then
        blnReadOnly = True
    End If
    
    SetSignature = blnReadOnly
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub vsOPS_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String, blnCancel As Boolean
    Dim str�Ա� As String, int�Ա� As Integer
    Dim strInput As String, vPoint As POINTAPI
    
    On Error GoTo errH
    With vsOPS
        If Col = col�������� Or Col = col�������� Then
            If .EditText = "" Then
                .EditText = .Cell(flexcpData, Row, Col)
                If mblnReturn Then Call OPSEnterNextCell
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call OPSEnterNextCell
            ElseIf Col = col�������� And .TextMatrix(Row, col��������) <> "" And .Cell(flexcpData, Row, Col) <> "" And .EditText Like "*" & .Cell(flexcpData, Row, Col) & "*" Then
                '�жϼ���ǰ׺��������Ƿ������������ϱ���
                strInput = UCase(.EditText)
                StrSQL = GetSQL(2, strInput, str�Ա�)
                Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, strInput & "%", mstrLike & strInput & "%", str�Ա�, int�Ա�)
                If rsTmp.RecordCount <> 1 Then
                    '�����ڱ�׼������ǰ�����븽����Ϣ
                    .TextMatrix(Row, col��������) = .EditText
                Else
                    Call OPSSetInput(Row, Col, rsTmp)
                    .EditText = .Text
                End If
'                '������.Cell(flexcpData, Row, Col)���Ա��޸�����ʱ�ٴ�ʹ��like�ж�
                .Tag = ""
                mblnChange = True
            Else
                strInput = UCase(.EditText)
                StrSQL = GetSQL(2, strInput, str�Ա�)
                If str�Ա� = "��" Then
                    int�Ա� = 1
                ElseIf str�Ա� = "Ů" Then
                    int�Ա� = 2
                End If
                vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 0, IIf(optInput(4).Value, "������Ŀ", "��������"), False, "", "", False, True, True, _
                    vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%", str�Ա�, int�Ա�)
                If rsTmp Is Nothing Then
                    If Not blnCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                        If chkInfo(chk��������¼��).Value = 0 Or Col = col�������� Then
                            MsgBox "û���ҵ������ҵ�������Ŀ��", vbInformation, Me.Caption
                            Cancel = True
                        Else
                            .TextMatrix(Row, col��������) = ""
                            .Cell(flexcpData, Row, col��������) = ""
                            .TextMatrix(Row, col������ĿID) = ""
                            .TextMatrix(Row, col��������ID) = ""
                            .Tag = ""
                            '�����ʼ�ձ���һ����
                            If Row = .Rows - 1 Then .AddItem ""
                        End If
                    Else
                        Cancel = True
                    End If
                Else
                    Call OPSSetInput(Row, Col, rsTmp): .EditText = .Text
                    If mblnReturn Then Call OPSEnterNextCell
                End If
            End If
            mblnReturn = False
        ElseIf Col = col����ʽ Then
            If .EditText = "" Then
                .EditText = .Cell(flexcpData, Row, Col)
                If mblnReturn Then Call OPSEnterNextCell
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call OPSEnterNextCell
            Else
                strInput = UCase(.EditText)
                StrSQL = _
                    " Select A.ID,A.����,A.����,A.�������� as ��������" & _
                    " From ������ĿĿ¼ A,������Ŀ���� B" & _
                    " Where A.���='G' And A.ID=B.������ĿID" & _
                    " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
                    " And (A.���� Like [1] Or A.���� Like [2] Or B.���� Like [2] Or B.���� Like [2])" & _
                    " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                    " Order by A.����"
                
                vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 0, "������Ŀ", False, "", "", False, True, True, _
                    vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%")
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "û���ҵ�ƥ���������Ŀ��", vbInformation, gstrSysName
                    End If
                    Cancel = True
                Else
                    Call OPSSetInput(Row, Col, rsTmp): .EditText = .Text
                    If mblnReturn Then Call OPSEnterNextCell
                End If
            End If
            mblnReturn = False
        ElseIf Col = col����ҽʦ Or Col = col����1 Or Col = col����2 Or Col = col����ҽʦ Then
            If (Col = col����1 Or Col = col����2) And .EditText = "" Then
                .TextMatrix(Row, Col) = "": .Cell(flexcpData, Row, Col) = ""
                If Col = col����1 Then
                    .TextMatrix(Row, col����2) = "": .Cell(flexcpData, Row, col����2) = ""
                End If
                If mblnReturn Then Call OPSEnterNextCell
            ElseIf .EditText = "" Then
                .EditText = .Cell(flexcpData, Row, Col)
                If mblnReturn Then Call OPSEnterNextCell
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call OPSEnterNextCell
            Else
                strInput = UCase(.EditText)
                StrSQL = "Select A.ID,A.���,A.����,A.����" & _
                    " From ��Ա�� A,��Ա����˵�� B" & _
                    " Where A.ID=B.��ԱID And B.��Ա����='ҽ��'" & _
                    " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
                    " And (A.��� Like [1] Or A.���� Like [2] Or A.���� Like [2])" & _
                    " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                    " Order by A.���"
                vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 0, "ҽ��", False, "", "", False, False, True, _
                    vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%")
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        If (Col = col����ҽʦ Or Col = col����1 Or Col = col����2) And zlCommFun.IsCharChinese(.EditText) Then
                            If MsgBox("û���ҵ�ƥ��ı�Ժҽ�����Ƿ�¼��δ�ڱ�Ժ������ҽ����", vbQuestion + vbYesNo + vbDefaultButton1, Me.Caption) = vbYes Then
                                .Tag = ""
                                mblnChange = True
                                If mblnReturn Then Call OPSEnterNextCell
                                Exit Sub
                            End If
                        Else
                            MsgBox "û���ҵ�ƥ���ҽ����", vbInformation, gstrSysName
                        End If
                    End If
                    Cancel = True
                Else
                    Call OPSSetInput(Row, Col, rsTmp): .EditText = .Text
                    If mblnReturn Then Call OPSEnterNextCell
                End If
            End If
            mblnReturn = False
        ElseIf Col = col������ʿ Then
            If .EditText = "" Then
                .EditText = .Cell(flexcpData, Row, Col)
                If mblnReturn Then Call OPSEnterNextCell
            ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
                If mblnReturn Then Call OPSEnterNextCell
            Else
                strInput = UCase(.EditText)
                StrSQL = "Select A.ID,A.���,A.����,A.����" & _
                    " From ��Ա�� A,��Ա����˵�� B" & _
                    " Where A.ID=B.��ԱID And B.��Ա����='��ʿ'" & _
                    " And (A.����ʱ�� = To_Date('3000-01-01', 'YYYY-MM-DD') Or A.����ʱ�� Is Null)" & _
                    " And (A.��� Like [1] Or A.���� Like [2] Or A.���� Like [2])" & _
                    " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                    " Order by A.���"
                vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 0, "��ʿ", False, "", "", False, False, True, _
                    vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, strInput & "%", mstrLike & strInput & "%")
                If rsTmp Is Nothing Then
                    If Not blnCancel Then
                        MsgBox "û���ҵ�ƥ��Ļ�ʿ��", vbInformation, gstrSysName
                    End If
                    Cancel = True
                Else
                    Call OPSSetInput(Row, Col, rsTmp): .EditText = .Text
                    If mblnReturn Then Call OPSEnterNextCell
                End If
            End If
            mblnReturn = False
        ElseIf Col = COL�������.COL������� Or Col = colASA�ּ� Or Col = colNNIS�ּ� Then
            If .TextMatrix(Row, Col) <> .EditText Then
                If .Tag = "δ�޸�" Then .Tag = "": mblnChange = True
            End If
        ElseIf Col = col�ٴ����� Or Col = colԤ���ÿ���ҩ Or Col = col��Ԥ�ڵĶ������� Or Col = col������֢ Or Col = col������������ Or Col = col��������֢ _
                Or Col = col�����Ѫ��Ѫ�� Or Col = col�����˿��ѿ� Or Col = col�������Ѫ˨ Or Col = col���������л���� Or Col = col�������˥�� _
                Or Col = col�����˨�� Or Col = col�����Ѫ֢ Or Col = col�����Źؽڹ��� Then
            If .TextMatrix(Row, Col) <> IIf(.EditText = 2, 0, -1) Then
                If .Tag = "δ�޸�" Then .Tag = "": mblnChange = True
            End If
        ElseIf Col = col����ҩ���� Then
            If Len(Trim(.EditText)) > 5 Then
                MsgBox "������ҩ�������ܳ���5λ����", vbInformation, Me.Caption
                Cancel = True
                Exit Sub
            Else
                If .EditText <> .TextMatrix(Row, Col) And .Tag = "δ�޸�" Then .Tag = "": mblnChange = True
            End If
        End If
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsTSJC_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Call vsTSJC_AfterRowColChange(-1, -1, vsTSJC.Row, vsTSJC.Col)
End Sub

Private Sub vsTSJC_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    vsTSJC.ComboList = "..."
End Sub

Private Sub vsTSJC_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String, blnCancel As Boolean
    Dim strSQLItem As String
    
    With vsTSJC
        strSQLItem = _
            " From ������ĿĿ¼ A" & _
            " Where A.���='D' And A.������� IN(2,3) And A.����Ӧ��=1" & _
            " And (A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� IS NULL)" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)"
        StrSQL = "Select 0 as ĩ��,Max(Level) as ��ID,ID,�ϼ�ID,����,����,NULL as ��λ" & _
            " From ���Ʒ���Ŀ¼ Where ����=5 And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " Start With ID In (Select A.����ID" & strSQLItem & ") Connect by Prior �ϼ�ID=ID" & _
            " Group by ID,�ϼ�ID,����,����"
        StrSQL = StrSQL & " Union ALL" & _
            " Select 1 as ĩ��,1 as ��ID,A.ID,����ID as �ϼ�ID,A.����,A.����,A.���㵥λ as ��λ" & _
            strSQLItem & " Order By ĩ��,��ID Desc,����"
        Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 2, "������", False, "", "", False, True, False, 0, 0, 0, blnCancel, False, False)
        If rsTmp Is Nothing Then
            If Not blnCancel Then
                MsgBox "û�м����Ŀ���ݿ���ѡ��", vbInformation, gstrSysName
            End If
        Else
            Call TSJCSetDiagInput(Row, rsTmp)
            Call TSJCEnterNextCell
        End If
    End With
End Sub

Private Sub vsTSJC_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim i As Long, j As Long
    
    If mbln��ʿվ Or mblnReadOnly Then Exit Sub
    
    If KeyCode = vbKeyF4 Then
        Call zlCommFun.PressKey(vbKeySpace)
    ElseIf KeyCode = vbKeyDelete Then
        If MsgBox("ȷʵҪɾ������������", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbYes Then
            With vsTSJC
                .TextMatrix(.Row, 1) = ""
            End With
            mblnChange = True
        End If
    ElseIf KeyCode > 127 Then
        '���ֱ�����뺺�ֵ�����
        Call vsTSJC_KeyPress(KeyCode)
    End If
End Sub

Private Sub vsTSJC_KeyPress(KeyAscii As Integer)
    If mbln��ʿվ Or mblnReadOnly Then Exit Sub
    
    With vsTSJC
        If KeyAscii = 13 Then
            KeyAscii = 0
            Call TSJCEnterNextCell
        Else
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                Call vsTSJC_CellButtonClick(.Row, .Col)
            Else
                .ComboList = "" 'ʹ��ť״̬��������״̬
            End If
        End If
    End With
End Sub

Private Sub vsTSJC_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = 13 Then
        mblnReturn = True
    Else
        mblnReturn = False
    End If
End Sub

Private Sub vsTSJC_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    vsTSJC.EditSelStart = 0
    vsTSJC.EditSelLength = zlCommFun.ActualLen(vsTSJC.EditText)
End Sub

Private Sub vsTSJC_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim rsTmp As ADODB.Recordset
    Dim StrSQL As String, blnCancel As Boolean
    Dim strInput As String, vPoint As POINTAPI
    
    With vsTSJC
        If .EditText = "" Then
            .EditText = .Cell(flexcpData, Row, Col)
            If mblnReturn Then Call TSJCEnterNextCell
        ElseIf .EditText = .Cell(flexcpData, Row, Col) Then
            If mblnReturn Then Call TSJCEnterNextCell
        Else
            strInput = UCase(.EditText)
            If LenB(StrConv(strInput, vbFromUnicode)) > 100 Then
                MsgBox "����������ݲ��ܳ���50�����֡�"
                Cancel = True
                Exit Sub
            End If
            If zlCommFun.IsCharChinese(strInput) Then
                StrSQL = "B.���� Like [2]" '���뺺��ʱֻƥ������
            Else
                StrSQL = "A.���� Like [1] Or B.���� Like [2] Or B.���� Like [2]"
            End If
            StrSQL = _
                " Select Distinct A.ID,A.����,A.����,A.���㵥λ as ��λ" & _
                " From ������ĿĿ¼ A,������Ŀ���� B" & _
                " Where A.ID=B.������ĿID And A.���='D' And A.������� IN(2,3)" & _
                " And (A.����ʱ�� Is Null Or A.����ʱ��=To_Date('3000-01-01','YYYY-MM-DD'))" & _
                " And A.����Ӧ��=1 And B.����=[3] And (" & StrSQL & ")" & _
                " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
                " Order by A.����"
            If zlCommFun.IsCharChinese(strInput) Then
                On Error GoTo errH
                Set rsTmp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, strInput & "%", mstrLike & strInput & "%", mint���� + 1)
                If rsTmp.EOF Then
                    Set rsTmp = Nothing
                ElseIf rsTmp.RecordCount > 1 Then
                    Set rsTmp = Nothing '����¼��ʱ�ж��ƥ�䲻����ѡ��
                End If
                Call TSJCSetDiagInput(Row, rsTmp)
                .EditText = .Text
                If mblnReturn Then Call TSJCEnterNextCell
            Else
                vPoint = GetCoordPos(.hwnd, .CellLeft + 15, .CellTop)
                Set rsTmp = zlDatabase.ShowSQLSelect(Me, StrSQL, 0, "������", _
                    False, "", "", False, False, True, vPoint.X, vPoint.Y, .CellHeight, blnCancel, False, True, _
                    strInput & "%", mstrLike & strInput & "%", mint���� + 1)
                If blnCancel Then '��ƥ������ʱ,���������봦��,ȡ����ͬ
                    Cancel = True
                Else
                    Call TSJCSetDiagInput(Row, rsTmp)
                    .EditText = .Text
                    If mblnReturn Then Call TSJCEnterNextCell
                End If
            End If
        End If
        mblnReturn = False
    End With
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function CheckDateRange(ByVal strDate As String, Optional ByVal blnCheckData As Boolean) As Boolean
    '���¼�������Ƿ������Ժ���ڷ�Χ�ڣ�
    'blnCheckData true:ֻ������ڷ�Χ�������ʱ�䷶Χ��false:������ʱ�䷶Χ
    ' ��Ժ����Ϊ�գ�����false,��Ժ����Ϊ������Ϊ3000-01-01
    
    Dim DateStart As Date, dateEnd As Date
    
    On Error GoTo errH
    CheckDateRange = False
    If Not IsDate(strDate) Then Exit Function
    
    If Trim("" & txtInfo(txt��Ժʱ��).Text) = "" Then
        DateStart = CDate(0)
    Else
        DateStart = CDate(Trim("" & txtInfo(txt��Ժʱ��).Text))
    End If
    If Trim("" & txtInfo(txt��Ժʱ��).Text) = "" Then
        dateEnd = CDate(0)
    Else
        dateEnd = CDate(Trim("" & txtInfo(txt��Ժʱ��).Text))
    End If
    
    If DateStart = CDate(0) Then Exit Function
    If dateEnd = CDate(0) Then dateEnd = zlDatabase.Currentdate
    
    If blnCheckData Then
        If Between(Format(strDate, "yyyy-MM-dd"), Format(DateStart, "yyyy-MM-dd"), Format(dateEnd, "yyyy-MM-dd")) Then
            CheckDateRange = True
        End If
    Else
        If CDate(strDate) >= DateStart And CDate(strDate) <= dateEnd Then
            CheckDateRange = True
        End If
    End If
    Exit Function
errH:
    CheckDateRange = False
End Function

Private Sub vs����_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '--------------------------------------------------------------------------------
    '������صĸ�ʽ
    '���˺�:2007/09/17
    '--------------------------------------------------------------------------------
    With vs����
        Select Case Col
        Case .ColIndex("�������Ʊ���")
'            .ColComboList(Col) = "..."
            If .ComboIndex < 0 Then Exit Sub
            .Cell(flexcpData, Row, Col) = .ComboData(.ComboIndex)
        Case .ColIndex("��Ұ��λ")
           ' .ColComboList(Col) = "..."
        End Select
    End With
End Sub
 

Private Sub vs����_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
  Call zl_VsGridRowChange(vs����, OldRow, NewRow, OldCol, NewCol)
End Sub

Private Sub vs����_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    
    With vs����
        Select Case Col
        Case .ColIndex("�������Ʊ���"), .ColIndex("��Ұ��λ")
        Case .ColIndex("��ʼ����"), .ColIndex("��������")
        Case .ColIndex("�������"), .ColIndex("�ۼ���")
        Case .ColIndex("����Ч��")
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub vs����_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    '--------------------------------------------------------------------------
    '����:��ťѡ��
    '����:
    '--------------------------------------------------------------------------
    With vs����
        Select Case Col
        Case .ColIndex("�������Ʊ���")
'            If Select���������("", False) = False Then
'                Exit Sub
'            End If
        Case Else
        End Select
    End With
End Sub

Private Sub vs����_GotFocus()
    Call zl_VsGridGotFocus(vs����)
End Sub

Private Sub vs����_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lngCol As Long, lngRow As Long, strKEY As String
   If mbln��ʿվ Or mblnReadOnly Then Exit Sub
    With vs����
        If (.Col = .ColIndex("��Ұ��λ")) And KeyCode <> vbKeyReturn Then
          '  .ColComboList(.Col) = ""
        End If
        If KeyCode = vbKeyDelete Then
            If MsgBox("���Ƿ����Ҫɾ�����еķ�����Ϣ��?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
            If .Row = .Rows - 1 And .Row = 1 Then
                For lngCol = 0 To .Cols - 1
                    .TextMatrix(.Row, lngCol) = ""
                    .Cell(flexcpData, .Row, lngCol) = ""
                    .RowData(.Row) = ""
                Next
            Else
                .RemoveItem .Row
            End If
            zlCtlSetFocus vs����, True
        End If
    End With
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With vs����
        If Val(.Cell(flexcpData, .Row, .ColIndex("�������Ʊ���"))) = 0 Or Trim(.TextMatrix(.Row, .ColIndex("��ʼ����"))) = "" Then
            Err = 0: On Error Resume Next
            If sstInfo.TabVisible(sstInfo.Tab + IIf(sstInfo.TabVisible(sstInfo.Tab + 1), 1, 2)) Then sstInfo.Tab = sstInfo.Tab + IIf(sstInfo.TabVisible(sstInfo.Tab + 1), 1, 2)
            Call vsKSS.SetFocus
            Exit Sub
        End If
        Select Case .Col
        Case .Cols - 1
            If Not .Row >= .Rows - 1 Then
                .Col = .ColIndex("�������Ʊ���")
                .Row = .Row + 1
            Else
                Call vs����_KeyDownEdit(.Row, .Col, KeyCode, Shift)
            End If
            .SetFocus
        Case Else
            zlCommFun.PressKey vbKeyRight
        End Select
    End With
End Sub

Private Sub vs����_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim intCol As Integer, lngRow As Long, strKEY As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With vs����
        Select Case Col
        Case .ColIndex("�������Ʊ���")
'            strKey = Trim(.EditText)
'            strKey = Replace(strKey, Chr(vbKeyReturn), "")
'            strKey = Replace(strKey, Chr(10), "")
'            If strKey = "" Then Exit Sub
'            If Select���������(strKey, True) = False Then
'                .TextMatrix(Row, Col) = .EditText: .Cell(flexcpData, Row, Col) = ""
'                For intCol = 0 To .Cols - 1
'                    If intCol <> Col Then
'                        .TextMatrix(Row, intCol) = ""
'                        .Cell(flexcpData, Row, intCol) = ""
'                    End If
'                Next
'                Exit Sub
'            End If
'            .EditText = .TextMatrix(Row, Col)
        Case Else
        End Select
        Call zlVsMoveGridCell(vs����, .ColIndex("��ʼ����"), .Cols - 1, True, lngRow)
        If lngRow > 0 Then
            '��ʾ��������һ��,��Ҫ������ص�ȱʡֵ
            strKEY = .ColData(.ColIndex("�������Ʊ���"))
            If InStr(1, strKEY, ";") > 0 Then
                .TextMatrix(lngRow, .ColIndex("�������Ʊ���")) = Mid(strKEY, InStr(1, strKEY, ";") + 1)
                .Cell(flexcpData, lngRow, .ColIndex("�������Ʊ���")) = Mid(strKEY, 1, InStr(1, strKEY, ";") - 1)
            End If
        End If
    End With
     
End Sub

Private Sub vs����_KeyPress(KeyAscii As Integer)
    If mbln��ʿվ Or mblnReadOnly Then Exit Sub
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub vs����_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0: Exit Sub
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Exit Sub
    End If
    With vs����
        Select Case Col
        Case .ColIndex("�������Ʊ���"), .ColIndex("��Ұ��λ")
            Call VsFlxGridCheckKeyPress(vs����, Row, Col, KeyAscii, m�ı�ʽ)
        Case .ColIndex("��ʼ����"), .ColIndex("��������")
            Call VsFlxGridCheckKeyPress(vs����, Row, Col, KeyAscii, m�ı�ʽ)
        Case .ColIndex("�������"), .ColIndex("�ۼ���")
            Call VsFlxGridCheckKeyPress(vs����, Row, Col, KeyAscii, m���ʽ)
        Case .ColIndex("����Ч��")
        Case Else
        End Select
    End With
End Sub

Private Sub vs����_LostFocus()
    Call zl_VsGridLOSTFOCUS(vs����)
End Sub

Private Sub vs����_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKEY As String
    Dim intCol As Integer
    Dim strTemp As String
    
    With vs����
        strKEY = Trim(.EditText): strKEY = Replace(strKEY, Chr(vbKeyReturn), ""): strKEY = Replace(strKEY, Chr(10), "")
        Select Case Col
        Case .ColIndex("�������Ʊ���")
        Case .ColIndex("��ʼ����")
            If strKEY = "" Then Exit Sub
            strKEY = CheckIsDate(strKEY, "��ʼ����")
            If strKEY = "" Then Cancel = True: Exit Sub
            If Check������Ч��(strKEY, "��ʼ����") = False Then Cancel = True: Exit Sub
            If Trim(.TextMatrix(Row, .ColIndex("��������"))) <> "" Then
                If strKEY > Trim(.TextMatrix(Row, .ColIndex("��������"))) Then
                    MsgBox "��ʼ���ڲ��ܴ��ڽ�������,����!", vbInformation, Me.Caption
                    Cancel = True
                    Exit Sub
                End If
            End If
            
            .EditText = strKEY
        Case .ColIndex("��������")
            If strKEY = "" Then Exit Sub
            strKEY = CheckIsDate(strKEY, "��������")
            If strKEY = "" Then Cancel = True: Exit Sub
            If Check������Ч��(strKEY, "��������") = False Then Cancel = True: Exit Sub
            If Trim(.TextMatrix(Row, .ColIndex("��ʼ����"))) <> "" Then
                If strKEY < Trim(.TextMatrix(Row, .ColIndex("��ʼ����"))) Then
                    MsgBox "�������ڲ���С�ڿ�ʼ����,����!", vbInformation, Me.Caption
                    Cancel = True
                    Exit Sub
                End If
            End If
            .EditText = strKEY
        Case .ColIndex("��Ұ��λ")
            If strKEY = "" Then Exit Sub
            If zlCommFun.StrIsValid(strKEY, 50, 0, "��Ұ��λ") = False Then
                Cancel = True: Exit Sub
            End If
        Case .ColIndex("�������"), .ColIndex("�ۼ���")
            If strKEY = "" Then Exit Sub
            If DblIsValid(strKEY, 10, True, False, 0, .ColKey(Col)) = False Then Cancel = True: Exit Sub
            If strKEY = "" Then Cancel = True: Exit Sub
            .EditText = strKEY
        Case .ColIndex("����Ч��")
        End Select
        mblnChange = True
        vs����.Tag = ""
    End With
End Sub
Private Function Load���������(ByVal lng����ID As Long, ByVal lng��ҳID As Long) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:���ط����뻯����Ϣ
    '���:
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-10-21 15:55:27
    '����:13999
    '-----------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    Dim StrSQL As String
    
    Err = 0: On Error GoTo Errhand:
    StrSQL = " " & _
    "   Select A.����id, A.��ҳid, A.���, A.����id, A.��ʼ����, A.��������, A.�Ƴ���, A.����, A.���Ʒ���, A.����Ч��, " & _
    "          B.���� || '-' || B.���� As ������Ϣ " & _
    "   From �������Ƽ�¼ A, ��������Ŀ¼ B " & _
    "   Where A.����id = B.Id And a.����id=[1] And a.��ҳid=[2] " & _
    "   Order By ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, lng����ID, lng��ҳID)
    With vs����
        .Rows = 2
        .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
        .Clear 1
        If rsTemp.RecordCount <> 0 Then .Rows = rsTemp.RecordCount + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("��ѧ���Ʊ���")) = Nvl(rsTemp!������Ϣ)
            .Cell(flexcpData, lngRow, .ColIndex("��ѧ���Ʊ���")) = Nvl(rsTemp!����id)
            .TextMatrix(lngRow, .ColIndex("��ʼ����")) = Format(rsTemp!��ʼ����, "yyyy-MM-DD")
            .TextMatrix(lngRow, .ColIndex("��������")) = Format(rsTemp!��������, "yyyy-MM-DD")
            .TextMatrix(lngRow, .ColIndex("�Ƴ���")) = Format(Val(Nvl(rsTemp!�Ƴ���)), "###;-###;;")
            .TextMatrix(lngRow, .ColIndex("����")) = Format(Val(Nvl(rsTemp!����)), "###;-###;;")
            .TextMatrix(lngRow, .ColIndex("���Ʒ���")) = Trim(Nvl(rsTemp!���Ʒ���))
            .TextMatrix(lngRow, .ColIndex("����Ч��")) = Trim(Nvl(rsTemp!����Ч��))
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    StrSQL = " " & _
    "   Select A.����id, A.��ҳid, A.���, A.����id, A.��ʼ����, A.��������,A.��Ұ��λ, A.�������, A.�ۼ���, A.����Ч��, " & _
    "          B.���� || '-' || B.���� As ������Ϣ " & _
    "   From �������Ƽ�¼ A, ��������Ŀ¼ B " & _
    "   Where A.����id = B.Id And a.����id=[1] And a.��ҳid=[2] " & _
    "   Order By ���"
    Set rsTemp = zlDatabase.OpenSQLRecord(StrSQL, Me.Caption, lng����ID, lng��ҳID)
    With vs����
        .Rows = 2
        .Cell(flexcpData, 1, 0, 1, .Cols - 1) = ""
        .Clear 1
        If rsTemp.RecordCount <> 0 Then .Rows = rsTemp.RecordCount + 1
        lngRow = 1
        Do While Not rsTemp.EOF
            .TextMatrix(lngRow, .ColIndex("�������Ʊ���")) = Nvl(rsTemp!������Ϣ)
            .Cell(flexcpData, lngRow, .ColIndex("�������Ʊ���")) = Nvl(rsTemp!����id)
            .TextMatrix(lngRow, .ColIndex("��ʼ����")) = Format(rsTemp!��ʼ����, "yyyy-MM-DD")
            .TextMatrix(lngRow, .ColIndex("��������")) = Format(rsTemp!��������, "yyyy-MM-DD")
            .TextMatrix(lngRow, .ColIndex("�������")) = Format(Val(Nvl(rsTemp!�������)), "###;-###;;")
            .TextMatrix(lngRow, .ColIndex("�ۼ���")) = Format(Val(Nvl(rsTemp!�ۼ���)), "###;-###;;")
            .TextMatrix(lngRow, .ColIndex("��Ұ��λ")) = Trim(Nvl(rsTemp!��Ұ��λ))
            .TextMatrix(lngRow, .ColIndex("����Ч��")) = Trim(Nvl(rsTemp!����Ч��))
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    Load��������� = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub vs����_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '--------------------------------------------------------------------------------
    '������صĸ�ʽ
    '���˺�:2007/09/17
    '--------------------------------------------------------------------------------
    With vs����
        Select Case Col
        Case .ColIndex("��ѧ���Ʊ���")
            '.ColComboList(Col) = "..."
             If .ComboIndex < 0 Then Exit Sub
            .Cell(flexcpData, Row, Col) = .ComboData(.ComboIndex)
        End Select
    End With
End Sub
 

Private Sub vs����_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
  Call zl_VsGridRowChange(vs����, OldRow, NewRow, OldCol, NewCol)
End Sub

Private Sub vs����_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vs����
        Select Case Col
        Case .ColIndex("��ѧ���Ʊ���"), .ColIndex("���Ʒ���")
        Case .ColIndex("��ʼ����"), .ColIndex("��������")
        Case .ColIndex("�Ƴ���"), .ColIndex("����")
        Case .ColIndex("����Ч��")
        Case Else
            Cancel = True
        End Select
    End With
End Sub

Private Sub vs����_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    '--------------------------------------------------------------------------
    '����:��ťѡ��
    '����:
    '--------------------------------------------------------------------------
    With vs����
        Select Case Col
'        Case .ColIndex("��ѧ���Ʊ���")
'            If Select���������("", True) = False Then
'                Exit Sub
'            End If
        Case Else
        End Select
    End With
End Sub

Private Sub vs����_GotFocus()
    Call zl_VsGridGotFocus(vs����)
End Sub

Private Sub vs����_KeyDown(KeyCode As Integer, Shift As Integer)
   Dim lngCol As Long, lngRow As Long, strKEY As String
   If mbln��ʿվ Or mblnReadOnly Then Exit Sub
    With vs����
        If (.Col = .ColIndex("��ѧ���Ʊ���")) And KeyCode <> vbKeyReturn Then
           ' .ColComboList(.Col) = ""
        End If
        If KeyCode = vbKeyDelete Then
            If MsgBox("���Ƿ����Ҫɾ�����еĻ�����Ϣ��?", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Sub
            If .Row = .Rows - 1 And .Row = 1 Then
                For lngCol = 0 To .Cols - 1
                    .TextMatrix(.Row, lngCol) = ""
                    .Cell(flexcpData, .Row, lngCol) = ""
                    .RowData(.Row) = ""
                Next
            Else
                .RemoveItem .Row
            End If
            zlCtlSetFocus vs����, True
        End If
    End With
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With vs����
        If Val(.Cell(flexcpData, .Row, .ColIndex("��ѧ���Ʊ���"))) = 0 Or Trim(.TextMatrix(.Row, .ColIndex("��ʼ����"))) = "" Then
            zlCtlSetFocus vs����, True
            Exit Sub
        End If
        Select Case .Col
        Case .Cols - 1
            If Not .Row >= .Rows - 1 Then
                .Col = 0
                .Row = .Row + 1
            Else
                Call vs����_KeyDownEdit(.Row, .Col, KeyCode, Shift)
            End If
            .SetFocus
        Case Else
            zlCommFun.PressKey vbKeyRight
        End Select
    End With
End Sub

Private Sub vs����_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim intCol As Integer, lngRow As Long
    Dim strKEY As String
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With vs����
        Select Case Col
        Case .ColIndex("��ѧ���Ʊ���")
'            strKey = Trim(.EditText)
'            strKey = Replace(strKey, Chr(vbKeyReturn), "")
'            strKey = Replace(strKey, Chr(10), "")
'            If strKey = "" Then Exit Sub
'            If Select���������(strKey, True) = False Then
'                .TextMatrix(Row, Col) = .EditText: .Cell(flexcpData, Row, Col) = ""
'                For intCol = 0 To .Cols - 1
'                    If intCol <> Col Then
'                        .TextMatrix(Row, intCol) = ""
'                        .Cell(flexcpData, Row, intCol) = ""
'                    End If
'                Next
'                Exit Sub
'            End If
'            .EditText = .TextMatrix(Row, Col)
        Case Else
        End Select
        Call zlVsMoveGridCell(vs����, .ColIndex("��ʼ����"), .Cols - 1, True, lngRow)
        If lngRow > 0 Then
            '��ʾ��������һ��,��Ҫ������ص�ȱʡֵ
            strKEY = .ColData(.ColIndex("��ѧ���Ʊ���"))
            If InStr(1, strKEY, ";") > 0 Then
                .TextMatrix(lngRow, .ColIndex("��ѧ���Ʊ���")) = Mid(strKEY, InStr(1, strKEY, ";") + 1)
                .Cell(flexcpData, lngRow, .ColIndex("��ѧ���Ʊ���")) = Mid(strKEY, 1, InStr(1, strKEY, ";") - 1)
                .TextMatrix(lngRow, .ColIndex("�Ƴ���")) = 1
            End If
        End If
    End With
     
End Sub

Private Sub vs����_KeyPress(KeyAscii As Integer)
    If mbln��ʿվ Or mblnReadOnly Then Exit Sub
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
    End If
End Sub

Private Sub vs����_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0: Exit Sub
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        Exit Sub
    End If
    With vs����
        Select Case Col
        Case .ColIndex("��ѧ���Ʊ���"), .ColIndex("���Ʒ���")
            Call VsFlxGridCheckKeyPress(vs����, Row, Col, KeyAscii, m�ı�ʽ)
        Case .ColIndex("��ʼ����"), .ColIndex("��������")
            Call VsFlxGridCheckKeyPress(vs����, Row, Col, KeyAscii, m�ı�ʽ)
        Case .ColIndex("�Ƴ���"), .ColIndex("����")
            Call VsFlxGridCheckKeyPress(vs����, Row, Col, KeyAscii, m���ʽ)
        Case .ColIndex("����Ч��")
        Case Else
        End Select
    End With
End Sub

Private Sub vs����_LostFocus()
    Call zl_VsGridLOSTFOCUS(vs����)
End Sub

Private Sub vs����_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKEY As String
    Dim intCol As Integer
    Dim strTemp As String
    
    With vs����
        strKEY = Trim(.EditText): strKEY = Replace(strKEY, Chr(vbKeyReturn), ""): strKEY = Replace(strKEY, Chr(10), "")
        Select Case Col
        Case .ColIndex("��ѧ���Ʊ���")
        Case .ColIndex("��ʼ����")
            If strKEY = "" Then Exit Sub
            strKEY = CheckIsDate(strKEY, "��ʼ����")
            If strKEY = "" Then Cancel = True: Exit Sub
            If Check������Ч��(strKEY, "��ʼ����") = False Then Cancel = True: Exit Sub
            If Trim(.TextMatrix(Row, .ColIndex("��������"))) <> "" Then
                If strKEY > Trim(.TextMatrix(Row, .ColIndex("��������"))) Then
                    MsgBox "��ʼ���ڲ��ܴ��ڽ�������,����!", vbInformation, Me.Caption
                    Cancel = True
                    Exit Sub
                End If
            End If
            .EditText = strKEY
        Case .ColIndex("��������")
            If strKEY = "" Then Exit Sub
            strKEY = CheckIsDate(strKEY, "��������")
            If strKEY = "" Then Cancel = True: Exit Sub
            If Check������Ч��(strKEY, "��������") = False Then Cancel = True: Exit Sub
            If Trim(.TextMatrix(Row, .ColIndex("��ʼ����"))) <> "" Then
                If strKEY < Trim(.TextMatrix(Row, .ColIndex("��ʼ����"))) Then
                    MsgBox "�������ڲ���С�ڿ�ʼ����,����!", vbInformation, Me.Caption
                    Cancel = True
                    Exit Sub
                End If
            End If
            .EditText = strKEY
        Case .ColIndex("���Ʒ���")
            If strKEY = "" Then Exit Sub
            If zlCommFun.StrIsValid(strKEY, 50, 0, "���Ʒ���") = False Then
                
                Cancel = True: Exit Sub
            End If
        Case .ColIndex("�Ƴ���")
            If strKEY = "" Then Exit Sub
            If DblIsValid(strKEY, 3, True, False, 0, .ColKey(Col)) = False Then Cancel = True: Exit Sub
            If strKEY = "" Then Cancel = True: Exit Sub
            .EditText = strKEY
        Case .ColIndex("����")
            If strKEY = "" Then Exit Sub
            If DblIsValid(strKEY, 10, True, False, 0, .ColKey(Col)) = False Then Cancel = True: Exit Sub
            If strKEY = "" Then Cancel = True: Exit Sub
            .EditText = strKEY
        End Select
        mblnChange = True
        vs����.Tag = ""
    End With
End Sub

Private Sub zl_VsGridLOSTFOCUS(ByVal vsGrid As VSFlexGrid, Optional CustomColor As OLE_COLOR = -1)
    '------------------------------------------------------------------------------------------------------------------------
   '���ܣ��뿪����ؼ�ʱѡ�����ɫ
    '��Σ�CustomColor-�Ƿ����Զ�����ɫ������(BackColor)�ķ�ʽ������)
    '���ƣ����˺�
    '���ڣ�2010-03-23 11:03:05
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    With vsGrid
        If CustomColor <> -1 Then
             If .Row >= .FixedRows Then
                .Cell(flexcpBackColor, .Row, .FixedCols, .Row, .Cols - 1) = CustomColor
            End If
        Else
            .SelectionMode = flexSelectionByRow
            .FocusRect = IIf(vsGrid.Editable = flexEDNone, flexFocusHeavy, flexFocusSolid)
            .HighLight = flexHighlightAlways
            .BackColorSel = GRD_LOSTFOCUS_COLORSEL
        End If
    End With
End Sub

Private Sub zl_VsGridRowChange(ByVal vsGrid As VSFlexGrid, ByVal lngOldRow As Long, ByVal lngNewRow As Long, _
    ByVal lngoldCol As Long, ByVal lngNewCol As Long, Optional CustomColor As OLE_COLOR = -1)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����иı�ʱ,������ص���ɫ
    '��Σ�CustomColor-�Զ�����ɫ
    '���Σ�
    '���أ�
    '���ƣ����˺�
    '���ڣ�2010-03-23 11:22:38
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    '�иı�ʱ
    Err = 0: On Error Resume Next
    If lngOldRow = lngNewRow Then
        vsGrid.Cell(flexcpBackColor, lngNewRow, vsGrid.FixedCols, lngNewRow, vsGrid.Cols - 1) = IIf(CustomColor <> -1, CustomColor, 16772055)
        Exit Sub
    End If
    With vsGrid
        .Cell(flexcpBackColor, lngOldRow, vsGrid.FixedCols, lngOldRow, .Cols - 1) = .BackColor
        .Cell(flexcpBackColor, lngNewRow, vsGrid.FixedCols, lngNewRow, .Cols - 1) = IIf(CustomColor <> -1, CustomColor, 16772055)
    End With
End Sub

Private Sub zl_VsGridGotFocus(ByVal vsGrid As VSFlexGrid, Optional CustomColor As OLE_COLOR = -1)
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ���������ؼ�ʱѡ�����ɫ
    '��Σ�CustomColor-�Զ���ɫ
    '���ƣ����˺�
    '���ڣ�2010-03-23 10:52:23
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Err = 0: On Error Resume Next
    '����ؼ�
    With vsGrid
         If CustomColor <> -1 Then
             .FocusRect = flexFocusSolid
             .HighLight = flexHighlightNever
             If .Row >= .FixedRows Then
                If .Rows - 1 > .FixedRows Then  '���ѡ����ɫ
                    .Cell(flexcpBackColor, .FixedRows, .FixedCols, .Rows - 1, .Cols - 1) = .BackColor
                End If
                 .Cell(flexcpBackColor, .Row, .FixedCols, .Row, .Cols - 1) = CustomColor
             End If
         Else
            .FocusRect = flexFocusSolid 'IIf(vsGrid.Editable = flexEDNone, flexFocusNone, flexFocusSolid)
            .HighLight = flexHighlightNever
            .BackColorSel = GRD_GOTFOCUS_COLORSEL
        End If
    End With
    Call zl_VsGridRowChange(vsGrid, vsGrid.Row, vsGrid.Row, 0, 0)
End Sub

Private Sub zlCtlSetFocus(ByVal objCtl As Object, Optional blnDoEvnts As Boolean = False)
    '����:�������ƶ��ؼ���:2008-07-08 16:48:35
    Err = 0: On Error Resume Next
    If blnDoEvnts Then DoEvents
    If IsCtrlSetFocus(objCtl) = True Then: objCtl.SetFocus
End Sub

Private Sub zlVsMoveGridCell(ByVal vsGrid As VSFlexGrid, _
    Optional lng���� As Long = -1, Optional lngβ�� As Long = -1, _
    Optional blnEdit As Boolean = False, Optional ByRef lngRow As Long = -1)
    '-----------------------------------------------------------------------------------------------------------
    '����:�ƶ���Ԫ�����
    '���:blnEdit-��ǰ�����ڱ༭״̬,����������
    '     lng����-����,���<0,������Ϊ0��,����Ϊָ������
    '     lngβ��-β��,���<0,������Ϊ.cols-1,����Ϊָ������
    '����:lngRow-������ڲ�����,�򷵻ر�������к�,���򷵻�-1
    '����:
    '����:���˺�
    '����:2008-11-06 14:24:12
    '-----------------------------------------------------------------------------------------------------------
    Dim lngCol As Long, lngLastCol As Long, arrSplit As Variant
    Dim i As Long
    
    Err = 0: On Error GoTo Errhand:
    
    'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
    If lng���� <> -1 Then
        lngCol = lng����
    Else
        lngCol = vsGrid.ColIndex(Split(vsGrid.Tag & "|", "|")(1))
    End If
    If lngCol = -1 Then lngCol = 0
    lngLastCol = IIf(lngβ�� < 0, vsGrid.Cols - 1, lngβ��)
    lngRow = -1
    With vsGrid
        If lngLastCol = .Col Then
            .Col = lngCol
            If .Row < .Rows - 1 Then
                .Row = .Row + 1
            Else
                If blnEdit = True Then
                    If Trim(.TextMatrix(.Row, lngCol)) <> "" Then
                        Call zlVsInsertIntoRow(vsGrid, .Row)
                        .Row = .Rows - 1
                        lngRow = .Row
                    End If
                End If
            End If
        Else
            .Col = .Col + 1
            For i = .Col To .Cols - 1
                'ColData(i):����������(1-�̶�,-1-����ѡ,0-��ѡ)||������(0-��������,1-��ֹ����,2-��������,�����س���������)
                arrSplit = Split(.ColData(i) & "||", "||")
                If .ColHidden(i) Or Val(arrSplit(1)) >= 1 Then
                    If .Col >= .Cols - 1 Then
                        If .Row < .Rows - 1 Then
                             .Row = .Row + 1
                             .Col = lngCol
                        Else
                            If blnEdit = True Then
                                If Trim(.TextMatrix(.Row, lngCol)) <> "" Then
                                    Call zlVsInsertIntoRow(vsGrid, .Row)
                                    .Row = .Rows - 1
                                    lngRow = .Row
                                End If
                            End If
                            .Col = lngCol
                        End If
                    Else
                        .Col = .Col + 1
                    End If
                Else
                    Exit For
                End If
            Next
        End If
        If .RowIsVisible(.Row) = False Then
            .TopRow = .Row
        End If
        If .ColIsVisible(.Col) = False Then
            .LeftCol = .Col
        Else
            If .CellLeft + .CellWidth > vsGrid.Width Then .LeftCol = .Col
        End If
        .SetFocus
    End With
    Exit Sub
Errhand:
End Sub

Private Sub VsFlxGridCheckKeyPress(ByVal objCtl As Object, Row As Long, Col As Long, KeyAscii As Integer, ByVal TextType As mTextType)
    '------------------------------------------------------------------------------------------------------------------
    '����:ֻ���������ֺͻس����˸�
    '����:
    '   objctl:Vsgrid8.0�ؼ�
    '   Keyascii:
    '           Keyascii:8 (�˸�)
    '   Row-��ǰ��
    '   Col-��ǰ��
    '   TextType:(0-�ı�ʽ;1-����ʽ;2-���ʽ)
    '����:һ��KeyAscii
    '------------------------------------------------------------------------------------------------------------------
    Err = 0
    On Error GoTo Errhand:
    
    If TextType = m�ı�ʽ Then
        If KeyAscii = Asc("'") Then
            KeyAscii = 0
        End If
        Exit Sub
    End If

    If KeyAscii < Asc("0") Or KeyAscii > Asc("9") Then
        Select Case KeyAscii
        Case vbKeyReturn       '�س�
        Case 8                 '�˸�
        Case Asc(".")
            If TextType = m���ʽ Or TextType = m�����ʽ Then
                If InStr(objCtl.EditText, ".") <> 0 Then     'ֻ�ܴ���һ��С����
                    KeyAscii = 0
                End If
            Else
                KeyAscii = 0
            End If
        Case Asc("-")          '����
            Dim iRow As Long
            Dim icol As Long
            If Trim(objCtl.EditText) = "" Then Exit Sub
            If TextType <> m�����ʽ Then KeyAscii = 0: Exit Sub
            If objCtl.EditSelStart <> 0 Then KeyAscii = 0: Exit Sub      '��겻���һλ,�������븺��
            If InStr(1, objCtl.EditText, "-") <> 0 Then   'ֻ�ܴ���һ������
                KeyAscii = 0
            End If
        Case Else
            KeyAscii = 0
        End Select
    End If
    Exit Sub
Errhand:
    KeyAscii = 0
End Sub

Private Function CheckIsDate(ByVal strKEY As String, ByVal strTittle As String) As String
    '------------------------------------------------------------------------------
    '����:����Ƿ�Ϸ���������,����Ϊ:(20070101��2007-01-01)����(01-01��0101)����(01<01-31>)
    '����:strKey-��Ҫ���Ĺؽ���
    '����:�Ϸ�������,���ر�׼��ʽ(yyyy-mm-dd),���򷵻�""
    '����:���˺�
    '����:2008/01/24
    '------------------------------------------------------------------------------
    If Len(strKEY) = 4 And InStr(1, strKEY, "-") = 0 Then
        '0101,��Ҫ��ǰ�����
        strKEY = Year(Now) & strKEY
    ElseIf Len(Replace(strKEY, "-", "")) = 4 And InStr(1, strKEY, "-") > 0 Then
        '01-01��ʽ,��Ҫ����
        strKEY = Year(Now) & Replace(strKEY, "-", "")
    ElseIf Len(strKEY) <= 2 And IsNumeric(strKEY) Then
        'ָ����
        strKEY = Format(Now, "YYYYMM") & IIf(Len(strKEY) = 2, strKEY, "0" & strKEY)
    End If
    If Len(strKEY) = 8 And InStr(1, strKEY, "-") = 0 Then
        strKEY = TranNumToDate(strKEY)
        If strKEY = "" Then
            MsgBox strTittle & "����Ϊ������,���飡", vbInformation, Me.Caption
            Exit Function
        End If
    End If
    If Not IsDate(strKEY) Then
        MsgBox strTittle & "����Ϊ��������(2000-10-10) ��20001010��,���飡", vbInformation, Me.Caption
        Exit Function
    End If
    CheckIsDate = strKEY
End Function

Private Function Check������Ч��(ByVal strDate As String, ByVal strTittle As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:������ڵ���Ч��
    '���:strDate-��ǰ����
    '     strTittle-����:��:�����ڵڼ���
    '����:
    '����:��Ч��strDate="",����true,���򷵻�False
    '����:���˺�
    '����:2008-10-21 17:03:30
    '-----------------------------------------------------------------------------------------------------------
    Dim strTemp As String, strCurDate As String
    Dim str��Ժʱ�� As String, str��Ժʱ�� As String
    
    If strDate = "" Then Check������Ч�� = True: Exit Function
    '��������Ƿ�Ϸ�
    If IsDate(strDate) = False Or IsNumeric(strDate) Then
        MsgBox strTittle & "����һ����Ч�����ڷ�Χ,����!", vbInformation, Me.Caption
        Exit Function
    End If
    str��Ժʱ�� = Format(txtInfo(txt��Ժʱ��).Text, "yyyy-mm-dd")
    If txtInfo(txt��Ժʱ��).Text <> "" Then str��Ժʱ�� = Format(txtInfo(txt��Ժʱ��).Text, "yyyy-mm-dd")
    strCurDate = Format(zlDatabase.Currentdate, "yyyy-mm-dd")
    If strDate > strCurDate Then
        MsgBox strTittle & "�ȵ�ǰ���ڻ�Ҫ��,����!", vbInformation, Me.Caption
        Exit Function
    End If
    
    If strDate < str��Ժʱ�� Then
        MsgBox strTittle & "����Ժ���ڻ�ҪС,����!", vbInformation, Me.Caption
        Exit Function
    End If
    If str��Ժʱ�� <> "" Then
        If str��Ժʱ�� < strDate Then
            MsgBox strTittle & "�ȳ�Ժ���ڻ�Ҫ��,����!", vbInformation, Me.Caption
            Exit Function
        End If
    End If
    Check������Ч�� = True
End Function

Private Function DblIsValid(ByVal strInput As String, ByVal intMax As Integer, Optional blnNegative As Boolean = True, Optional blnZero As Boolean = True, _
        Optional ByVal hwnd As Long = 0, Optional str��Ŀ As String) As Boolean
    '-----------------------------------------------------------------------------------------------------------
    '����:����ַ����Ƿ�Ϸ��Ľ��
    '���:strInput        ������ַ���
    '     intMax          ������λ��
    '     blnNegative     �Ƿ���и������
    '     blnZero         �Ƿ������ļ��
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008-10-20 15:16:08
    '-----------------------------------------------------------------------------------------------------------
   
    Dim dblValue As Double
    If blnZero = True Then
        If strInput = "" Then
            MsgBox str��Ŀ & "δ���룬����!", vbInformation, gstrSysName
            If hwnd <> 0 Then SetFocusHwnd hwnd
            Exit Function
        End If
    End If
    If strInput = "" Then DblIsValid = True: Exit Function
    If IsNumeric(strInput) = False Then
        MsgBox str��Ŀ & "������Ч�����ָ�ʽ��", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '���ý���
        Exit Function
    End If
    
    dblValue = Val(strInput)
    If dblValue >= 10 ^ intMax - 1 Then
        MsgBox str��Ŀ & "��ֵ���󣬲��ܳ���" & 10 ^ intMax - 1 & "��", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '���ý���
        Exit Function
    End If
    If blnNegative = True And dblValue < 0 Then
        MsgBox str��Ŀ & "�������븺����", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '���ý���
        Exit Function
    End If
    
    If Abs(dblValue) >= 10 ^ intMax And dblValue < 0 Then
        MsgBox str��Ŀ & "��ֵ��С������С��-" & 10 ^ intMax - 1 & "λ��", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '���ý���
        Exit Function
    End If
    
    
    If blnZero = True And dblValue = 0 Then
        MsgBox str��Ŀ & "���������㡣", vbInformation, gstrSysName
        If hwnd <> 0 Then SetFocusHwnd hwnd              '���ý���
        Exit Function
    End If
    DblIsValid = True
End Function

Private Function IsCtrlSetFocus(ByVal objCtl As Object) As Boolean
    '------------------------------------------------------------------------------
    '����:�жϿؼ��Ƿ��
    '����:����ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008/01/24
    '------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Err = 0: On Error GoTo Errhand:
    
    IsCtrlSetFocus = objCtl.Enabled And objCtl.Visible
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

Private Function zlVsInsertIntoRow(ByVal vsGrid As VSFlexGrid, ByVal lngRow As Long, Optional blnBefor As Boolean = False, _
    Optional blnMoveNewRow As Boolean = True) As Boolean
    '------------------------------------------------------------------------------
    '����:������
    '����:vsGrid-�����е�������
    '     lngRow-��ǰ��
    '     blnBefor-��lngrow֮���֮��.true:֮��,false-֮��
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2008/01/24
    '------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Err = 0: On Error GoTo Errhand:
    With vsGrid
        If blnBefor Then
            .AddItem "", lngRow
        Else
            .AddItem "", lngRow + 1
        End If
        If blnMoveNewRow = True Then
            If blnBefor Then '
                .Row = lngRow
            Else
                .Row = lngRow + 1
            End If
        End If
    End With
    zlVsInsertIntoRow = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
    Call SaveErrLog
End Function

'ת����ֵΪ����
Private Function TranNumToDate(ByVal strNum As String, Optional ByVal blnDec As Boolean = False) As String
    Dim strYear As String
    Dim strMonth As String
    Dim strDay As String
    Dim strDate As String
    
    TranNumToDate = ""
    strYear = Mid(strNum, 1, 4)
    strMonth = Mid(strNum, 5, 2)
    strDay = Mid(strNum, 7, 2)
        
    If strYear < 1000 Or strYear > 5000 Then Exit Function
    If strMonth = "" Then strMonth = "01"
    If strDay = "" Then strDay = "01"
    
    If strMonth > 12 Or strMonth < 1 Then Exit Function
    strDate = strYear & "-" & strMonth & "-" & strDay
        
    If Not IsDate(strDate) Then Exit Function
    
    strDate = Format(strDate, "yyyy-mm-dd")
    If blnDec Then strDate = DateAdd("d", -1, Format(strDate, "yyyy-mm-dd"))
    TranNumToDate = strDate
End Function

Private Sub FillVsf()
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    Dim lngCol As Long
    Dim StrSQL As String
    
    On Error GoTo errH
    StrSQL = "select ����,���� from ������Ŀ order by ����"
    vsfMain.Clear
    
    Call zlDatabase.OpenRecordset(rsTemp, StrSQL, Me.Caption)
    If rsTemp.RecordCount = 0 Then vsfMain.Rows = 1: vsfMain.Cols = 1: Exit Sub
    If (rsTemp.RecordCount Mod 2) <> 0 Then
        vsfMain.Rows = rsTemp.RecordCount \ 2 + 2
    Else
        vsfMain.Rows = rsTemp.RecordCount \ 2 + 1
    End If
    With vsfMain
        .Cols = 6
        For lngRow = 0 To 3 Step 3
            .TextMatrix(0, lngRow) = "��Ŀ"
            .TextMatrix(0, lngRow + 1) = "����"
            .TextMatrix(0, lngRow + 2) = "˵��"
            .ColWidth(0 + lngRow) = 1500
            .ColWidth(1 + lngRow) = 1200
            .ColHidden(2 + lngRow) = True
        Next lngRow
        .Cell(flexcpAlignment, 0, 0, 0, vsfMain.Cols - 1) = 4
        .Cell(flexcpBackColor, 1, 0, .Rows - 1, 0) = &HFCE7D8
        .Cell(flexcpBackColor, 1, 3, .Rows - 1, 3) = &HFCE7D8
        .ColAlignment(0) = flexAlignCenterCenter
        .ColAlignment(3) = flexAlignCenterCenter
    End With
    lngRow = 1
    lngCol = 0
    While Not rsTemp.EOF
        If lngCol < 4 Then
            With vsfMain
                .TextMatrix(lngRow, lngCol + 0) = rsTemp!����
                .TextMatrix(lngRow, lngCol + 2) = rsTemp!���� & ""
                If InStr(rsTemp!����, "�Ƿ�") > 0 Then
                    vsfMain.TextMatrix(lngRow, lngCol + 1) = "��"
                    vsfMain.Cell(flexcpChecked, lngRow, lngCol + 1) = 2
                End If
            End With
            lngCol = lngCol + 3
            rsTemp.MoveNext
        Else
            lngCol = 0
            lngRow = lngRow + 1
        End If
    Wend
    vsfMain.Editable = flexEDKbdMouse
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function SavePageDataUnit(ByRef blnDiagnose As Boolean, ByVal blnBeforSign As Boolean) As Boolean

'���ܣ���鱣����һ�����ҳ���淽��
'������blnBeforSign-�Ƿ�ǩ��ʱ����ǰ����
'���أ�blnDiagnose=�Ƿ���д�����
'���أ�SavePageDataUnit=����ɹ�

    If Not CheckPageData(blnDiagnose, blnBeforSign) Then Exit Function
    
    If mblnDiagnose And Not blnDiagnose Then
        If MsgBox("Ҫ��������Ϣ��û�����룬Ҫ������", vbQuestion + vbYesNo + vbDefaultButton2) = vbNo Then Exit Function
    End If
    
    If Not SavePageData(blnBeforSign) Then Exit Function
    '���ý��������
    Call SetFaceEditable(mblnReadOnly)
    
    SavePageDataUnit = True
    
    
End Function
