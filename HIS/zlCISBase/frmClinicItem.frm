VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Begin VB.Form frmClinicItem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������Ŀ�༭"
   ClientHeight    =   8115
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   10275
   Icon            =   "frmClinicItem.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8115
   ScaleWidth      =   10275
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Visible         =   0   'False
   Begin VB.CheckBox chkGoOn 
      Caption         =   "��������������Ŀ"
      Height          =   180
      Left            =   5880
      TabIndex        =   106
      Top             =   7785
      Visible         =   0   'False
      Width           =   1935
   End
   Begin MSComctlLib.TreeView tvwClass 
      Height          =   3570
      Left            =   2040
      TabIndex        =   55
      TabStop         =   0   'False
      Tag             =   "1000"
      Top             =   7440
      Visible         =   0   'False
      Width           =   3540
      _ExtentX        =   6244
      _ExtentY        =   6297
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   353
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      FullRowSelect   =   -1  'True
      ImageList       =   "imgList"
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin TabDlg.SSTab stbInfo 
      Height          =   7020
      Left            =   120
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   555
      Width           =   10005
      _ExtentX        =   17648
      _ExtentY        =   12383
      _Version        =   393216
      Style           =   1
      Tabs            =   5
      TabsPerRow      =   5
      TabHeight       =   520
      TabCaption(0)   =   "��Ŀ����(&B)"
      TabPicture(0)   =   "frmClinicItem.frx":058A
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lblComment"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "imgComment"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Label2"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbl��Ŀ����"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbl��Ŀ����"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lblִ��Ƶ��"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lbl�����Ա�"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lbl���㷽ʽ"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lbl���Ƽ���"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lbl��������"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lbl��������"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lbl��������"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lbl���㵥λ"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "Label3"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "Label4"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "lblӢ��"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "lbl����˵��"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "lblִ�з���"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "lbl�������"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "lbl�������"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "Label5"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "Label6"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "lblML"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "lbllel"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "lbl��Һ����"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "lblZLPL"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "lbl�Թܱ���"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "cboZLPL"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "cbo�������"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "fra�걾��λ"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "fra¼����"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "fra��鲿λ"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "fra��׼����"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "chk�������"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txt�ο�"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "cbo��������"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txt��Ŀ����"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).Control(37)=   "txt��Ŀ����"
      Tab(0).Control(37).Enabled=   0   'False
      Tab(0).Control(38)=   "cbo�����Ա�"
      Tab(0).Control(38).Enabled=   0   'False
      Tab(0).Control(39)=   "cbo���㷽ʽ"
      Tab(0).Control(39).Enabled=   0   'False
      Tab(0).Control(40)=   "txt����ƴ��"
      Tab(0).Control(40).Enabled=   0   'False
      Tab(0).Control(41)=   "txt�������"
      Tab(0).Control(41).Enabled=   0   'False
      Tab(0).Control(42)=   "txt��������"
      Tab(0).Control(42).Enabled=   0   'False
      Tab(0).Control(43)=   "chk����Ӧ��"
      Tab(0).Control(43).Enabled=   0   'False
      Tab(0).Control(44)=   "chkִ�а���"
      Tab(0).Control(44).Enabled=   0   'False
      Tab(0).Control(45)=   "txt����ƴ��"
      Tab(0).Control(45).Enabled=   0   'False
      Tab(0).Control(46)=   "txt�������"
      Tab(0).Control(46).Enabled=   0   'False
      Tab(0).Control(47)=   "txt���㵥λ"
      Tab(0).Control(47).Enabled=   0   'False
      Tab(0).Control(48)=   "txt¼������"
      Tab(0).Control(48).Enabled=   0   'False
      Tab(0).Control(49)=   "cbo¼��������Χ"
      Tab(0).Control(49).Enabled=   0   'False
      Tab(0).Control(50)=   "txtӢ��"
      Tab(0).Control(50).Enabled=   0   'False
      Tab(0).Control(51)=   "cbo����˵��"
      Tab(0).Control(51).Enabled=   0   'False
      Tab(0).Control(52)=   "cboִ��Ƶ��"
      Tab(0).Control(52).Enabled=   0   'False
      Tab(0).Control(53)=   "picFound"
      Tab(0).Control(53).Enabled=   0   'False
      Tab(0).Control(54)=   "chk����"
      Tab(0).Control(54).Enabled=   0   'False
      Tab(0).Control(55)=   "Frame3"
      Tab(0).Control(55).Enabled=   0   'False
      Tab(0).Control(56)=   "txtML"
      Tab(0).Control(56).Enabled=   0   'False
      Tab(0).Control(57)=   "cbo��Һ����"
      Tab(0).Control(57).Enabled=   0   'False
      Tab(0).Control(58)=   "cbo�������"
      Tab(0).Control(58).Enabled=   0   'False
      Tab(0).Control(59)=   "vsfBloodLis"
      Tab(0).Control(59).Enabled=   0   'False
      Tab(0).Control(60)=   "cmd�ο�"
      Tab(0).Control(60).Enabled=   0   'False
      Tab(0).Control(61)=   "cmdDel�ο�"
      Tab(0).Control(61).Enabled=   0   'False
      Tab(0).Control(62)=   "picTestTubeCode"
      Tab(0).Control(62).Enabled=   0   'False
      Tab(0).Control(63)=   "cboִ�з���"
      Tab(0).Control(63).Enabled=   0   'False
      Tab(0).Control(64)=   "cboBloodType"
      Tab(0).Control(64).Enabled=   0   'False
      Tab(0).Control(65)=   "chkNoTMSY"
      Tab(0).Control(65).Enabled=   0   'False
      Tab(0).Control(66)=   "chkYYPS"
      Tab(0).Control(66).Enabled=   0   'False
      Tab(0).ControlCount=   67
      TabCaption(1)   =   "ִ�п���(&E)"
      TabPicture(1)   =   "frmClinicItem.frx":05A6
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fraDeptFind"
      Tab(1).Control(1)=   "picDept"
      Tab(1).Control(2)=   "fraִ�в���"
      Tab(1).ControlCount=   3
      TabCaption(2)   =   "��鲿λ(&L)"
      TabPicture(2)   =   "frmClinicItem.frx":05C2
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "optList(0)"
      Tab(2).Control(1)=   "optList(1)"
      Tab(2).Control(2)=   "vfgList"
      Tab(2).ControlCount=   3
      TabCaption(3)   =   "Ƥ�Խ��(&P)"
      TabPicture(3)   =   "frmClinicItem.frx":05DE
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "fraƤ�Խ��"
      Tab(3).ControlCount=   1
      TabCaption(4)   =   "Ƶ������(&R)"
      TabPicture(4)   =   "frmClinicItem.frx":05FA
      Tab(4).ControlEnabled=   0   'False
      Tab(4).Control(0)=   "lblFreq"
      Tab(4).Control(1)=   "vsfFreq"
      Tab(4).ControlCount=   2
      Begin VB.CheckBox chkYYPS 
         Caption         =   "ԴҺƤ��"
         Height          =   270
         Left            =   8775
         TabIndex        =   155
         Top             =   3075
         Width           =   1100
      End
      Begin VB.CheckBox chkNoTMSY 
         Caption         =   "����������ʹ��"
         Height          =   180
         Left            =   7080
         TabIndex        =   154
         Top             =   3075
         Visible         =   0   'False
         Width           =   1665
      End
      Begin VB.ComboBox cboBloodType 
         Height          =   300
         Left            =   5115
         Style           =   2  'Dropdown List
         TabIndex        =   147
         Top             =   2295
         Visible         =   0   'False
         Width           =   1080
      End
      Begin VB.ComboBox cboִ�з��� 
         Height          =   300
         Left            =   5130
         Style           =   2  'Dropdown List
         TabIndex        =   88
         Top             =   2290
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.PictureBox picTestTubeCode 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   5115
         ScaleHeight     =   300
         ScaleWidth      =   1785
         TabIndex        =   150
         Top             =   2280
         Width           =   1785
         Begin VB.PictureBox picTubeColor 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   270
            Left            =   1515
            ScaleHeight     =   240
            ScaleWidth      =   240
            TabIndex        =   152
            Top             =   15
            Visible         =   0   'False
            Width           =   270
         End
         Begin VB.ComboBox cboTestTubeCode 
            Height          =   300
            Left            =   0
            Style           =   2  'Dropdown List
            TabIndex        =   151
            Top             =   15
            Width           =   1500
         End
      End
      Begin VB.CommandButton cmdDel�ο� 
         Height          =   285
         Left            =   6650
         Picture         =   "frmClinicItem.frx":0616
         Style           =   1  'Graphical
         TabIndex        =   149
         Top             =   3360
         Width           =   285
      End
      Begin VB.CommandButton cmd�ο� 
         Caption         =   "��"
         Height          =   285
         Left            =   6360
         TabIndex        =   148
         TabStop         =   0   'False
         Top             =   3360
         Width           =   285
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfBloodLis 
         Height          =   1000
         Left            =   7080
         TabIndex        =   144
         Top             =   3090
         Visible         =   0   'False
         Width           =   2760
         _cx             =   4860
         _cy             =   1764
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
         BackColorSel    =   16769992
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483633
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483632
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   2
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   255
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmClinicItem.frx":09D9
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
         VirtualData     =   0   'False
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
      Begin VB.ComboBox cbo������� 
         Height          =   300
         Left            =   8160
         Style           =   2  'Dropdown List
         TabIndex        =   143
         Top             =   3030
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.ComboBox cbo��Һ���� 
         Height          =   300
         Left            =   8160
         Style           =   2  'Dropdown List
         TabIndex        =   142
         Top             =   2295
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.TextBox txtML 
         Height          =   300
         Left            =   7080
         TabIndex        =   138
         Top             =   2660
         Width           =   675
      End
      Begin VB.Frame Frame3 
         Caption         =   "���÷�Χ"
         Height          =   2800
         Left            =   120
         TabIndex        =   119
         Top             =   4100
         Width           =   9735
         Begin VB.Frame Frame4 
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   1080
            TabIndex        =   128
            Top             =   2400
            Width           =   8385
            Begin VB.OptionButton OptAppUse 
               Caption         =   "Ӧ���ڱ���"
               Height          =   225
               Index           =   0
               Left            =   75
               TabIndex        =   132
               Top             =   60
               Value           =   -1  'True
               Width           =   1725
            End
            Begin VB.OptionButton OptAppUse 
               Caption         =   "Ӧ����ͬ��"
               Height          =   225
               Index           =   1
               Left            =   1920
               TabIndex        =   131
               Top             =   60
               Width           =   1605
            End
            Begin VB.OptionButton OptAppUse 
               Caption         =   "Ӧ���ڷ���������"
               Height          =   225
               Index           =   2
               Left            =   3720
               TabIndex        =   130
               Top             =   60
               Width           =   2235
            End
            Begin VB.OptionButton OptAppUse 
               Caption         =   "Ӧ���ڵ�ǰ���"
               Height          =   225
               Index           =   3
               Left            =   6120
               TabIndex        =   129
               Top             =   60
               Width           =   2055
            End
         End
         Begin VB.ComboBox cmbStationNo 
            Height          =   300
            Left            =   5340
            Style           =   2  'Dropdown List
            TabIndex        =   123
            Top             =   240
            Visible         =   0   'False
            Width           =   1500
         End
         Begin VB.CheckBox chk������� 
            Caption         =   "���(&P)"
            Height          =   225
            Index           =   2
            Left            =   3060
            TabIndex        =   122
            Top             =   285
            Width           =   930
         End
         Begin VB.CheckBox chk������� 
            Caption         =   "סԺ(&I)"
            Height          =   225
            Index           =   1
            Left            =   2100
            TabIndex        =   121
            Top             =   285
            Value           =   1  'Checked
            Width           =   930
         End
         Begin VB.CheckBox chk������� 
            Caption         =   "����(&W)"
            Height          =   225
            Index           =   0
            Left            =   1140
            TabIndex        =   120
            Top             =   285
            Value           =   1  'Checked
            Width           =   930
         End
         Begin VSFlex8Ctl.VSFlexGrid vsUseDept 
            Height          =   1800
            Left            =   1080
            TabIndex        =   124
            Top             =   600
            Width           =   8565
            _cx             =   15108
            _cy             =   3175
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
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483638
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
            Rows            =   7
            Cols            =   10
            FixedRows       =   0
            FixedCols       =   0
            RowHeightMin    =   245
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmClinicItem.frx":0A31
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
            Caption         =   "ʹ�ÿ���"
            Height          =   255
            Left            =   240
            TabIndex        =   133
            Top             =   2460
            Width           =   855
         End
         Begin VB.Label lblStationNo 
            AutoSize        =   -1  'True
            Caption         =   "Ժ�����(&Z)"
            Height          =   180
            Left            =   4320
            TabIndex        =   127
            Top             =   300
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.Label Label7 
            Caption         =   "����Χ"
            Height          =   255
            Left            =   240
            TabIndex        =   126
            Top             =   285
            Width           =   855
         End
         Begin VB.Label Label8 
            Caption         =   "ʹ�ÿ���"
            Height          =   255
            Left            =   240
            TabIndex        =   125
            Top             =   660
            Width           =   855
         End
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "�����Ի�����ִ��(&J)"
         Height          =   240
         Left            =   7200
         TabIndex        =   118
         Top             =   3825
         Visible         =   0   'False
         Width           =   2280
      End
      Begin VB.Frame fraDeptFind 
         BorderStyle     =   0  'None
         Height          =   350
         Left            =   -72970
         TabIndex        =   103
         Top             =   1750
         Width           =   5475
         Begin VB.TextBox txtLocate 
            Height          =   320
            Left            =   4380
            TabIndex        =   111
            ToolTipText     =   "������һ��F3��س�����λ�����F4"
            Top             =   57
            Width           =   1000
         End
         Begin VB.OptionButton optDeptKind 
            Caption         =   "ִ�п���(&E)"
            Height          =   375
            Index           =   0
            Left            =   720
            TabIndex        =   110
            Top             =   30
            Value           =   -1  'True
            Width           =   1300
         End
         Begin VB.OptionButton optDeptKind 
            Caption         =   "���˿���(&B)"
            Height          =   375
            Index           =   1
            Left            =   2160
            TabIndex        =   109
            Top             =   30
            Width           =   1300
         End
         Begin VB.Label lblX 
            Caption         =   "/"
            Height          =   195
            Left            =   2040
            TabIndex        =   137
            Top             =   120
            Width           =   375
         End
         Begin VB.Label lblAn 
            Caption         =   "��"
            Height          =   195
            Left            =   480
            TabIndex        =   136
            Top             =   120
            Width           =   375
         End
         Begin VB.Label lblLocate 
            Caption         =   "����(&F)"
            Height          =   195
            Left            =   3540
            TabIndex        =   112
            Top             =   120
            Width           =   735
         End
      End
      Begin VB.PictureBox picFound 
         BorderStyle     =   0  'None
         Height          =   210
         Left            =   6360
         ScaleHeight     =   210
         ScaleWidth      =   3570
         TabIndex        =   101
         Top             =   0
         Width           =   3570
         Begin VB.Label lblFound 
            AutoSize        =   -1  'True
            Caption         =   "����Ŀ��2002-12-20����"
            Height          =   180
            Left            =   1560
            TabIndex        =   102
            Top             =   0
            Width           =   1980
         End
      End
      Begin VB.PictureBox picDept 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   3135
         Left            =   -69840
         ScaleHeight     =   3105
         ScaleWidth      =   4890
         TabIndex        =   96
         Top             =   3120
         Visible         =   0   'False
         Width           =   4920
         Begin VB.CommandButton cmdFindCancle 
            Caption         =   "ȡ��"
            Height          =   270
            Left            =   4200
            TabIndex        =   117
            Top             =   360
            Width           =   615
         End
         Begin VB.CommandButton cmdFindOk 
            Caption         =   "ȷ��"
            Height          =   270
            Left            =   3480
            TabIndex        =   116
            Top             =   360
            Width           =   615
         End
         Begin VB.CommandButton cmdFind 
            Caption         =   "����"
            Height          =   270
            Left            =   1740
            TabIndex        =   108
            Top             =   360
            Width           =   615
         End
         Begin VB.TextBox txtFind 
            Height          =   270
            Left            =   50
            TabIndex        =   107
            Top             =   360
            Width           =   1575
         End
         Begin VB.CheckBox ChkSelect 
            Appearance      =   0  'Flat
            BackColor       =   &H00E0E0E0&
            Caption         =   "ȫѡ"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   2115
            TabIndex        =   98
            TabStop         =   0   'False
            Top             =   88
            Width           =   675
         End
         Begin VB.ComboBox cboProperty 
            Height          =   300
            Left            =   795
            Style           =   2  'Dropdown List
            TabIndex        =   97
            Top             =   45
            Width           =   1215
         End
         Begin MSComctlLib.ListView lvwItems 
            Height          =   2280
            Left            =   75
            TabIndex        =   99
            Top             =   795
            Width           =   4755
            _ExtentX        =   8387
            _ExtentY        =   4022
            View            =   3
            LabelEdit       =   1
            Sorted          =   -1  'True
            LabelWrap       =   -1  'True
            HideSelection   =   0   'False
            FullRowSelect   =   -1  'True
            GridLines       =   -1  'True
            _Version        =   393217
            Icons           =   "imgList"
            SmallIcons      =   "imgList"
            ForeColor       =   -2147483640
            BackColor       =   -2147483643
            BorderStyle     =   1
            Appearance      =   0
            NumItems        =   0
         End
         Begin VB.Label lbl�������� 
            AutoSize        =   -1  'True
            BackColor       =   &H00E0E0E0&
            Caption         =   "��������"
            Height          =   180
            Left            =   50
            TabIndex        =   100
            Top             =   110
            Width           =   720
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vfgList 
         Height          =   6135
         Left            =   -74910
         TabIndex        =   72
         Top             =   720
         Width           =   9795
         _cx             =   17277
         _cy             =   10821
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
         BackColorFixed  =   15790320
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   8
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.ComboBox cboִ��Ƶ�� 
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   90
         Top             =   2290
         Width           =   2115
      End
      Begin VB.ComboBox cbo����˵�� 
         Height          =   300
         Left            =   8160
         TabIndex        =   87
         Text            =   "cbo����˵��"
         Top             =   3780
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.Frame fraƤ�Խ�� 
         Caption         =   "Ƥ�Խ��"
         Height          =   6375
         Left            =   -74760
         TabIndex        =   77
         Top             =   480
         Width           =   9495
         Begin VB.CommandButton cmdTestDel 
            Caption         =   "ɾ��(&D)"
            Height          =   350
            Left            =   6480
            TabIndex        =   84
            Top             =   2280
            Width           =   1100
         End
         Begin VB.CommandButton cmdTestAdd 
            Caption         =   "����(&A)"
            Height          =   350
            Left            =   5280
            TabIndex        =   83
            Top             =   2280
            Width           =   1100
         End
         Begin VB.CheckBox chkƤ�Թ��� 
            Caption         =   "����"
            Height          =   180
            Left            =   6000
            TabIndex        =   82
            Top             =   1560
            Width           =   735
         End
         Begin VB.TextBox txtƤ������ 
            Height          =   300
            Left            =   6000
            MaxLength       =   13
            TabIndex        =   81
            Top             =   1020
            Width           =   1515
         End
         Begin VB.TextBox txtƤ�Ա�ע 
            Height          =   300
            Left            =   6000
            MaxLength       =   8
            TabIndex        =   79
            Top             =   540
            Width           =   1515
         End
         Begin VSFlex8Ctl.VSFlexGrid vsfTest 
            Height          =   5775
            Left            =   240
            TabIndex        =   85
            Top             =   360
            Width           =   4755
            _cx             =   8387
            _cy             =   10186
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
            BackColorFixed  =   15790320
            ForeColorFixed  =   -2147483630
            BackColorSel    =   16772055
            ForeColorSel    =   -2147483640
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483636
            GridColorFixed  =   -2147483636
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483643
            FocusRect       =   0
            HighLight       =   1
            AllowSelection  =   0   'False
            AllowBigSelection=   -1  'True
            AllowUserResizing=   1
            SelectionMode   =   1
            GridLines       =   1
            GridLinesFixed  =   1
            GridLineWidth   =   1
            Rows            =   2
            Cols            =   3
            FixedRows       =   1
            FixedCols       =   0
            RowHeightMin    =   300
            RowHeightMax    =   0
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   0   'False
            FormatString    =   ""
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
            AllowUserFreezing=   1
            BackColorFrozen =   0
            ForeColorFrozen =   0
            WallPaperAlignment=   8
            AccessibleName  =   ""
            AccessibleDescription=   ""
            AccessibleValue =   ""
            AccessibleRole  =   24
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "����(&W)"
            Height          =   180
            Left            =   5280
            TabIndex        =   80
            Top             =   1080
            Width           =   630
         End
         Begin VB.Label lbl��ע 
            AutoSize        =   -1  'True
            Caption         =   "��ע(&B)"
            Height          =   180
            Left            =   5280
            TabIndex        =   78
            Top             =   600
            Width           =   630
         End
      End
      Begin VB.TextBox txtӢ�� 
         Height          =   300
         Left            =   7845
         MaxLength       =   12
         TabIndex        =   75
         Top             =   3405
         Width           =   1980
      End
      Begin VB.OptionButton optList 
         Caption         =   "����ʾ��ѡ��λ"
         Height          =   375
         Index           =   1
         Left            =   -66960
         TabIndex        =   74
         Top             =   360
         Width           =   1935
      End
      Begin VB.OptionButton optList 
         Caption         =   "��ʾ���в�λ"
         Height          =   375
         Index           =   0
         Left            =   -68760
         TabIndex        =   73
         Top             =   360
         Value           =   -1  'True
         Width           =   1935
      End
      Begin VB.ComboBox cbo¼��������Χ 
         Height          =   300
         Left            =   4800
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   3765
         Width           =   1650
      End
      Begin VB.TextBox txt¼������ 
         Height          =   300
         Left            =   1200
         MaxLength       =   13
         TabIndex        =   31
         Top             =   3765
         Width           =   2115
      End
      Begin VB.TextBox txt���㵥λ 
         Height          =   300
         Left            =   5115
         TabIndex        =   25
         Top             =   2660
         Width           =   1785
      End
      Begin VB.TextBox txt������� 
         Height          =   300
         Left            =   4080
         MaxLength       =   12
         TabIndex        =   19
         Top             =   1940
         Width           =   2115
      End
      Begin VB.TextBox txt����ƴ�� 
         Height          =   300
         Left            =   1200
         MaxLength       =   12
         TabIndex        =   18
         Top             =   1940
         Width           =   2115
      End
      Begin VB.CheckBox chkִ�а��� 
         Caption         =   "��Ҫִ�а���(&W)"
         Height          =   210
         Left            =   3465
         TabIndex        =   28
         Top             =   3075
         Width           =   1740
      End
      Begin VB.CheckBox chk����Ӧ�� 
         Caption         =   "������Ӧ��(&Y)"
         Height          =   210
         Left            =   5250
         TabIndex        =   21
         Top             =   3075
         Value           =   1  'Checked
         Width           =   1680
      End
      Begin VB.TextBox txt�������� 
         Height          =   300
         Left            =   1200
         MaxLength       =   40
         TabIndex        =   16
         Top             =   1560
         Width           =   5715
      End
      Begin VB.TextBox txt������� 
         Height          =   300
         Left            =   4080
         MaxLength       =   12
         TabIndex        =   14
         Top             =   1215
         Width           =   2115
      End
      Begin VB.TextBox txt����ƴ�� 
         Height          =   300
         Left            =   1200
         MaxLength       =   12
         TabIndex        =   13
         Top             =   1215
         Width           =   2115
      End
      Begin VB.ComboBox cbo���㷽ʽ 
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   2660
         Width           =   2115
      End
      Begin VB.ComboBox cbo�����Ա� 
         Height          =   300
         Left            =   1200
         Style           =   2  'Dropdown List
         TabIndex        =   27
         Top             =   3015
         Width           =   2115
      End
      Begin VB.TextBox txt��Ŀ���� 
         Height          =   300
         Left            =   1200
         TabIndex        =   11
         Top             =   820
         Width           =   5715
      End
      Begin VB.TextBox txt��Ŀ���� 
         Height          =   300
         Left            =   1200
         MaxLength       =   13
         TabIndex        =   7
         Top             =   465
         Width           =   2115
      End
      Begin VB.ComboBox cbo�������� 
         Height          =   300
         ItemData        =   "frmClinicItem.frx":0BDD
         Left            =   5085
         List            =   "frmClinicItem.frx":0BDF
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   450
         Width           =   1815
      End
      Begin VB.TextBox txt�ο� 
         Height          =   300
         Left            =   1200
         TabIndex        =   30
         Top             =   3377
         Width           =   5460
      End
      Begin VB.CheckBox chk������� 
         Caption         =   "��ϼ�����Ŀ"
         Height          =   210
         Left            =   7860
         TabIndex        =   60
         Top             =   3825
         Visible         =   0   'False
         Width           =   1575
      End
      Begin VB.Frame fraִ�в��� 
         BorderStyle     =   0  'None
         Caption         =   "ִ�п���"
         Height          =   6405
         Left            =   -74760
         TabIndex        =   41
         Top             =   480
         Width           =   9795
         Begin VSFlex8Ctl.VSFlexGrid msf����ִ�� 
            Height          =   4455
            Left            =   180
            TabIndex        =   115
            Top             =   1680
            Width           =   9405
            _cx             =   16589
            _cy             =   7858
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
            GridColor       =   -2147483633
            GridColorFixed  =   -2147483638
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
            RowHeightMin    =   245
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
         Begin VB.OptionButton optִ�в��� 
            Caption         =   "���������ڿ���(&6)"
            Height          =   195
            Index           =   6
            Left            =   4500
            TabIndex        =   66
            Top             =   555
            Width           =   1860
         End
         Begin VB.Frame Frame2 
            BorderStyle     =   0  'None
            Height          =   330
            Left            =   120
            TabIndex        =   61
            Top             =   6135
            Width           =   8985
            Begin VB.OptionButton OptApp 
               Caption         =   "Ӧ���ڵ�ǰ���"
               Height          =   225
               Index           =   3
               Left            =   6240
               TabIndex        =   65
               Top             =   60
               Width           =   2415
            End
            Begin VB.OptionButton OptApp 
               Caption         =   "Ӧ���ڷ���������"
               Height          =   225
               Index           =   2
               Left            =   3840
               TabIndex        =   64
               Top             =   60
               Width           =   2235
            End
            Begin VB.OptionButton OptApp 
               Caption         =   "Ӧ����ͬ��"
               Height          =   225
               Index           =   1
               Left            =   1920
               TabIndex        =   63
               Top             =   60
               Width           =   1605
            End
            Begin VB.OptionButton OptApp 
               Caption         =   "Ӧ���ڱ���"
               Height          =   225
               Index           =   0
               Left            =   75
               TabIndex        =   62
               Top             =   60
               Value           =   -1  'True
               Width           =   1605
            End
         End
         Begin VB.Frame Frame1 
            Height          =   120
            Left            =   120
            TabIndex        =   59
            Top             =   780
            Width           =   9550
         End
         Begin VB.TextBox txtסԺִ�� 
            Enabled         =   0   'False
            Height          =   300
            Left            =   7725
            MaxLength       =   30
            TabIndex        =   50
            Top             =   940
            Width           =   1860
         End
         Begin VB.TextBox txt����ִ�� 
            Enabled         =   0   'False
            Height          =   300
            Left            =   3750
            MaxLength       =   30
            TabIndex        =   49
            Top             =   940
            Width           =   1890
         End
         Begin VB.OptionButton optִ�в��� 
            Caption         =   "ҽԺ��ִ��(&5)"
            Height          =   180
            Index           =   5
            Left            =   6660
            TabIndex        =   47
            Top             =   280
            Width           =   1485
         End
         Begin VB.OptionButton optִ�в��� 
            Caption         =   "ָ������ִ��(&4)"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   4
            Left            =   2385
            TabIndex        =   46
            Top             =   555
            Width           =   1665
         End
         Begin VB.OptionButton optִ�в��� 
            Caption         =   "����Ա���ڿ���(&3)"
            Height          =   180
            Index           =   3
            Left            =   90
            TabIndex        =   45
            Top             =   555
            Width           =   2025
         End
         Begin VB.OptionButton optִ�в��� 
            Caption         =   "�ɲ��˲���ִ��(&2)"
            Height          =   180
            Index           =   2
            Left            =   4500
            TabIndex        =   44
            Top             =   280
            Width           =   1845
         End
         Begin VB.OptionButton optִ�в��� 
            Caption         =   "�ɲ��˿���ִ��(&1)"
            Height          =   180
            Index           =   1
            Left            =   2385
            TabIndex        =   43
            Top             =   280
            Width           =   1845
         End
         Begin VB.OptionButton optִ�в��� 
            Caption         =   "������ִ�еĶ���(&0)"
            Height          =   180
            Index           =   0
            Left            =   90
            TabIndex        =   42
            Top             =   280
            Width           =   2025
         End
         Begin VB.Label lblסԺִ�� 
            AutoSize        =   -1  'True
            Caption         =   "סԺ����ִ�п���"
            Height          =   180
            Left            =   6210
            TabIndex        =   58
            Top             =   1005
            Width           =   1440
         End
         Begin VB.Label lblһ����� 
            AutoSize        =   -1  'True
            Caption         =   "1����ָ�����˿����⣺"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   90
            TabIndex        =   57
            Top             =   1000
            Width           =   1890
         End
         Begin VB.Label lbl����ִ�� 
            AutoSize        =   -1  'True
            Caption         =   "2��ָ�����˿��ң�"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   90
            TabIndex        =   51
            Top             =   1380
            Width           =   1530
         End
         Begin VB.Label lbl����ִ�� 
            AutoSize        =   -1  'True
            Caption         =   "���ﲡ��ִ�п���"
            Height          =   180
            Left            =   2250
            TabIndex        =   48
            Top             =   1005
            Width           =   1440
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfFreq 
         Height          =   6135
         Left            =   -74880
         TabIndex        =   91
         Top             =   720
         Width           =   9735
         _cx             =   17171
         _cy             =   10821
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
         BackColorFixed  =   15790320
         ForeColorFixed  =   -2147483630
         BackColorSel    =   16772055
         ForeColorSel    =   -2147483640
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483636
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483643
         FocusRect       =   0
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   1
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   3
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   ""
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
         AllowUserFreezing=   1
         BackColorFrozen =   0
         ForeColorFrozen =   0
         WallPaperAlignment=   8
         AccessibleName  =   ""
         AccessibleDescription=   ""
         AccessibleValue =   ""
         AccessibleRole  =   24
      End
      Begin VB.Frame fra��׼���� 
         Caption         =   "������׼����"
         Height          =   1400
         Left            =   7395
         TabIndex        =   33
         Top             =   375
         Visible         =   0   'False
         Width           =   2160
         Begin VB.CommandButton cmdѡ�� 
            Caption         =   "��"
            Height          =   285
            Left            =   1720
            TabIndex        =   135
            TabStop         =   0   'False
            Top             =   240
            Width           =   285
         End
         Begin VB.TextBox txt��׼���� 
            Height          =   300
            Left            =   135
            TabIndex        =   34
            Top             =   255
            Width           =   1875
         End
         Begin VB.Label lbl��׼���� 
            Height          =   600
            Left            =   180
            TabIndex        =   35
            Top             =   615
            Width           =   1800
            WordWrap        =   -1  'True
         End
      End
      Begin VB.Frame fra��鲿λ 
         Height          =   1380
         Left            =   7395
         TabIndex        =   36
         Top             =   375
         Visible         =   0   'False
         Width           =   2160
         Begin VB.TextBox txt��鲿λ 
            Height          =   300
            Left            =   390
            MaxLength       =   40
            TabIndex        =   38
            Top             =   435
            Width           =   1665
         End
         Begin VB.OptionButton opt��鲿λ 
            Caption         =   "��ѡ�ಿλ���(&X)"
            Height          =   285
            Index           =   1
            Left            =   120
            TabIndex        =   39
            Top             =   795
            Width           =   1980
         End
         Begin VB.OptionButton opt��鲿λ 
            Caption         =   "�̶�����λ���(&G)"
            Height          =   285
            Index           =   0
            Left            =   120
            TabIndex        =   37
            Top             =   165
            Value           =   -1  'True
            Width           =   1980
         End
      End
      Begin VB.Frame fra¼���� 
         Caption         =   "��¼�����"
         Height          =   1395
         Left            =   7380
         TabIndex        =   93
         Top             =   375
         Visible         =   0   'False
         Width           =   2160
         Begin VB.CheckBox chk���� 
            Caption         =   "����(&E)"
            Height          =   255
            Left            =   120
            TabIndex        =   94
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame fra�걾��λ 
         Caption         =   "Ĭ�ϱ걾��λ"
         Height          =   1350
         Left            =   7380
         TabIndex        =   69
         Top             =   375
         Visible         =   0   'False
         Width           =   2160
         Begin VB.CommandButton cmd�걾 
            Caption         =   "��"
            Height          =   285
            Left            =   1740
            TabIndex        =   71
            TabStop         =   0   'False
            Top             =   263
            Width           =   285
         End
         Begin VB.TextBox txt�걾��λ 
            Height          =   300
            Left            =   135
            TabIndex        =   70
            Top             =   255
            Width           =   1575
         End
      End
      Begin VB.ComboBox cbo������� 
         Height          =   300
         Left            =   8160
         Style           =   2  'Dropdown List
         TabIndex        =   105
         Top             =   450
         Width           =   1575
      End
      Begin VB.ComboBox cboZLPL 
         Height          =   300
         Left            =   5115
         Style           =   2  'Dropdown List
         TabIndex        =   145
         Top             =   2280
         Visible         =   0   'False
         Width           =   1785
      End
      Begin VB.Label lbl�Թܱ��� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�Թܱ���(&C)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3960
         TabIndex        =   153
         Top             =   2340
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lblZLPL 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "����Ƶ��(&T)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   3960
         TabIndex        =   146
         Top             =   2340
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lbl��Һ���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��Һ����"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   7200
         TabIndex        =   141
         Top             =   2355
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lbllel 
         Caption         =   "="
         Height          =   180
         Left            =   6960
         TabIndex        =   140
         Top             =   2720
         Width           =   135
      End
      Begin VB.Label lblML 
         Caption         =   "����"
         Height          =   180
         Left            =   7800
         TabIndex        =   139
         Top             =   2720
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "(���)"
         Height          =   255
         Left            =   6195
         TabIndex        =   114
         Top             =   1260
         Width           =   615
      End
      Begin VB.Label Label5 
         Caption         =   "(���)"
         Height          =   255
         Left            =   6195
         TabIndex        =   113
         Top             =   1963
         Width           =   615
      End
      Begin VB.Label lbl������� 
         Caption         =   "�ű�����"
         Height          =   180
         Left            =   7200
         TabIndex        =   104
         Top             =   510
         Width           =   855
      End
      Begin VB.Label lbl������� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�������(&G)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   7080
         TabIndex        =   95
         Top             =   3090
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lblFreq 
         AutoSize        =   -1  'True
         Caption         =   "��ѡ���������Ŀ��ִ��Ƶ�ʣ��ڵ�һ�д򹴡�"
         Height          =   180
         Left            =   -74880
         TabIndex        =   92
         Top             =   480
         Width           =   3780
      End
      Begin VB.Label lblִ�з��� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ִ�з���(&Z)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4065
         TabIndex        =   89
         Top             =   2350
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lbl����˵�� 
         AutoSize        =   -1  'True
         Caption         =   "����˵��"
         Height          =   180
         Left            =   7200
         TabIndex        =   86
         Top             =   3840
         Visible         =   0   'False
         Width           =   720
      End
      Begin VB.Label lblӢ�� 
         AutoSize        =   -1  'True
         Caption         =   "Ӣ����д"
         Height          =   255
         Left            =   7080
         TabIndex        =   76
         Top             =   3450
         Width           =   720
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "¼������Ӧ����"
         Height          =   180
         Left            =   3480
         TabIndex        =   68
         Top             =   3825
         Width           =   1260
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "¼������(&X)"
         Height          =   180
         Left            =   180
         TabIndex        =   67
         Top             =   3825
         Width           =   990
      End
      Begin VB.Label lbl���㵥λ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���㵥λ(&U)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4080
         TabIndex        =   24
         Top             =   2720
         Width           =   990
      End
      Begin VB.Label lbl�������� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������(&N)                        (ƴ��)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   165
         TabIndex        =   17
         Top             =   2000
         Width           =   3690
      End
      Begin VB.Label lbl�������� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������(&A)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   165
         TabIndex        =   15
         Top             =   1620
         Width           =   990
      End
      Begin VB.Label lbl�������� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��������(&E)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   4080
         TabIndex        =   8
         Top             =   510
         Width           =   990
      End
      Begin VB.Label lbl���Ƽ��� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���Ƽ���(&S)                        (ƴ��)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   165
         TabIndex        =   12
         Top             =   1275
         Width           =   3690
      End
      Begin VB.Label lbl���㷽ʽ 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "���㷽ʽ(&M)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   165
         TabIndex        =   22
         Top             =   2720
         Width           =   990
      End
      Begin VB.Label lbl�����Ա� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "�����Ա�(&R)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   165
         TabIndex        =   26
         Top             =   3075
         Width           =   990
      End
      Begin VB.Label lblִ��Ƶ�� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "ִ��Ƶ��(&Q)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   165
         TabIndex        =   20
         Top             =   2350
         Width           =   990
      End
      Begin VB.Label lbl��Ŀ���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��Ŀ����(&N)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   165
         TabIndex        =   10
         Top             =   880
         Width           =   990
      End
      Begin VB.Label lbl��Ŀ���� 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "��Ŀ����(&D)"
         ForeColor       =   &H80000008&
         Height          =   180
         Left            =   165
         TabIndex        =   6
         Top             =   525
         Width           =   990
      End
      Begin VB.Label Label2 
         Caption         =   "�ο���Ŀ(&F)"
         Height          =   255
         Left            =   180
         TabIndex        =   29
         Top             =   3400
         Width           =   1095
      End
      Begin VB.Image imgComment 
         Height          =   240
         Left            =   7305
         Picture         =   "frmClinicItem.frx":0BE1
         Top             =   1845
         Width           =   240
      End
      Begin VB.Label lblComment 
         Caption         =   $"frmClinicItem.frx":1A23
         Height          =   1545
         Left            =   7305
         TabIndex        =   40
         Top             =   1905
         Width           =   2430
         WordWrap        =   -1  'True
      End
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   9025
      TabIndex        =   53
      Top             =   7680
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   105
      Picture         =   "frmClinicItem.frx":1AC7
      TabIndex        =   54
      TabStop         =   0   'False
      Top             =   7695
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   7920
      TabIndex        =   52
      Top             =   7680
      Width           =   1100
   End
   Begin VB.TextBox txt���� 
      Height          =   300
      Left            =   1185
      Locked          =   -1  'True
      MaxLength       =   40
      TabIndex        =   1
      Top             =   75
      Width           =   5010
   End
   Begin VB.CommandButton cmd���� 
      Caption         =   "&P"
      Height          =   285
      Left            =   6225
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   75
      Width           =   285
   End
   Begin VB.ComboBox cbo��� 
      Height          =   300
      ItemData        =   "frmClinicItem.frx":1C11
      Left            =   8310
      List            =   "frmClinicItem.frx":1C13
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   60
      Width           =   1710
   End
   Begin VB.Frame fraLine 
      Height          =   120
      Left            =   -240
      TabIndex        =   56
      Top             =   360
      Width           =   10290
   End
   Begin MSComctlLib.ImageList imgList 
      Left            =   8160
      Top             =   7200
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicItem.frx":1C15
            Key             =   "close"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicItem.frx":21AF
            Key             =   "expend"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicItem.frx":2749
            Key             =   "Dept"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicItem.frx":2CE3
            Key             =   "unCheck"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClinicItem.frx":327D
            Key             =   "Check"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ListView lvwItem 
      Height          =   2280
      Left            =   1320
      TabIndex        =   134
      Top             =   7680
      Visible         =   0   'False
      Width           =   4755
      _ExtentX        =   8387
      _ExtentY        =   4022
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "imgList"
      SmallIcons      =   "imgList"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   0
   End
   Begin VB.Label lbl���� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "��Ŀ����(&T)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   120
      TabIndex        =   0
      Top             =   150
      Width           =   990
   End
   Begin VB.Label lbl��� 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "�������(&K)"
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   7230
      TabIndex        =   3
      Top             =   135
      Width           =   990
   End
End
Attribute VB_Name = "frmClinicItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'---------------------------------------------------
'˵����
'   1���ϼ�����ͨ��������ShowMe�������������塢Ȩ�ޡ��༭��Ŀ�ķ���ID��ID,�༭״̬����Ϣ���ݽ��뱾����
'   2���༭״̬����Me.tag��ţ��ֱ�Ϊ"����"��"�޸�"��"����"�����ϼ�����ͨ��ShowMe����
'---------------------------------------------------
Private lngClassId As Long       '���༭�ķ���ID���ϼ�����ͨ��ShowMe���ݽ���
Private lngItemID As Long        '���༭����ĿID���޸ġ�����ʱ���ϼ�����ͨ��ShowMe���ݽ���,����ʱΪ0��
Private lngVItemID As Long       '������Ŀ��ص�ָ��ID������������������У�
Private mlngOldId As Long

Private strInputed As String
Dim rsTemp As New ADODB.Recordset
Dim objNode As Node
Dim objItem As ListItem
Dim strTemp As String, aryTemp() As String
Dim intCount As Integer
Dim mstrMatch As String, strRefer As String '�ο�����
Dim mbln��ϲ�λ��Ŀ As Boolean             '�Ƿ��������Ŀ�Ĳ�λ��Ŀ
Dim mbln�����Ŀ As Boolean                 '�Ƿ��������Ŀ��������Ŀ��
Private mlng���볤�� As Long
Private mLast�������� As String
Private mFromLoad As Boolean                '�Ƿ��һ�ε���
Private mbln�������� As Boolean             '�Ƿ���������
Private mblnIniTest As Boolean
Private mstr��ѡִ�п��� As String
Private mstr��ѡʹ�ÿ��� As String
Private mblnRefresh As Boolean
Private mrs���ʷ��� As ADODB.Recordset
Private mstrPageCaption     '������¼�ϴε�ҳ�еı���
Private mblnOK As Boolean
Private mlngFind As Long
Private mstrFindStyle As String 'ƥ�䷽ʽ
Private mstrOldBlood As String  '��¼�޸�ǰ��Ѫ�����¼�б��е�ֵ
Private mblnPACSInterface As Boolean        '����Ӱ����Ϣϵͳ�ӿ�
Private mstrӦ�÷�Χ As String
Private Enum ִ�п���COL
    col���˿���ID = 0
    col���˿��� = 1
    colִ�п���ID = 2
    colִ�п��� = 3
End Enum

Private Sub Ini���ʷ���()
    'ȡ�������ʷ��࣬����Ѿ���ȡ�����˳�
    On Error GoTo ErrHandle
    If Not mrs���ʷ��� Is Nothing Then
        mrs���ʷ���.Filter = ""
        If Not mrs���ʷ���.EOF Then
            Exit Sub
        End If
    End If
    
    gstrSql = "Select ����,������ From �������ʷ���"
    Set mrs���ʷ��� = zlDatabase.OpenSQLRecord(gstrSql, "ȡ�������ʷ���")
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Init����Ƶ��(Optional ByVal strNo As String)
    Dim rsTemp As Recordset
    Dim strTemp As String
    Dim intIndex As Integer
    Dim i As Integer
    
    'ȡ����Ƶ����Ŀ�м����λΪСʱ���߷ֵ���Ŀ
    On Error GoTo ErrHandle
    
    gstrSql = "select ����,���� from ����Ƶ����Ŀ where �����λ='Сʱ' or �����λ='����'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "ȡ����Ƶ����Ŀ")
    With Me.cboZLPL
        .Clear
        .AddItem ""
        strTemp = "|"
        
        Do While Not rsTemp.EOF
            .AddItem rsTemp!����
            strTemp = strTemp & rsTemp!���� & "-" & rsTemp!���� & "|"
            i = i + 1
            If strNo = rsTemp!���� Then
                intIndex = i
            End If
            rsTemp.MoveNext
        Loop
        
        .ListIndex = intIndex
        Me.lblZLPL.Tag = strTemp
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
    
End Sub
Private Sub load���ʷ���(ByVal intType As Integer)
    'intType:0-ִ�п��ң��������ʣ������ڲ��ˣ���1-���˿��ң��ٴ����ʣ�;ʹ�ÿ���(���ݷ���Χ��վ�㣺�ٴ�����顢���顢���������ơ�������ʵĿ���)
    
    mblnRefresh = True
    
    With cboProperty
        .Clear
        
        If mrs���ʷ��� Is Nothing Then Exit Sub
        
        If intType = 0 Then
            mrs���ʷ���.Filter = "������=1 Or ������=2 Or ������=3"
        ElseIf intType = 1 Then
            mrs���ʷ���.Filter = "����='�ٴ�'"
        ElseIf intType = 2 Then
            mrs���ʷ���.Filter = "����='�ٴ�' Or ����='���' Or ����='����' Or ����='����' Or ����='����'" & IIf(chk�������(2).Value = 1, " Or ����='���'", "")
        End If
        
        If mrs���ʷ���.RecordCount = 0 Then Exit Sub
        
        If intType = 0 Or intType = 2 Then
            .AddItem "��������"
            
            Do While Not mrs���ʷ���.EOF
                .AddItem mrs���ʷ���!����
                
                mrs���ʷ���.MoveNext
            Loop
        ElseIf intType = 1 Then
            .AddItem "�ٴ�"
        End If
        
        If .ListCount > 0 Then .ListIndex = 0
    End With
    
    DoEvents
    
    mblnRefresh = False
End Sub
Private Sub Load����(ByVal intType As Integer, ByVal str�������� As String)
    'intType:0-ִ�п��ң��������ʣ������ڲ��ˣ���1-���˿��ң��ٴ����ʣ�;2-ʹ�ÿ���
    Dim rsData As ADODB.Recordset
    Dim objItem As ListItem
    Dim strTmp As String
    Dim strվ�� As String
    
    On Error GoTo ErrHandle
    If intType = 1 Then
        gstrSql = "select distinct ID,����,����" & _
                " from ���ű� D,��������˵�� T" & _
                " where D.ID=T.����ID and ��������=[1] " & _
                "       and (D.����ʱ�� is null or D.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))" & _
                " order by ����"
    ElseIf intType = 0 Then
        gstrSql = "select distinct ID,����,����" & _
                " from ���ű� D,��������˵�� T" & _
                " where D.ID=T.����ID and T.������� in (1,2,3) " & _
                " and (D.����ʱ�� is null or D.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))"
                
        If str�������� <> "��������" Then
            gstrSql = gstrSql & " and ��������=[1] "
        End If
                
        gstrSql = gstrSql & " order by ����"
    ElseIf intType = 2 Then
        If chk�������(1).Value = 1 Then strTmp = " T.�������=2"
        If chk�������(2).Value = 1 Or chk�������(0).Value = 1 Then strTmp = strTmp & IIf(strTmp = "", "", " Or") & " T.�������=1"
        If strTmp <> "" Then strTmp = " And (" & strTmp & " Or T.�������=3)"
        gstrSql = "select distinct ID,����,����" & _
                " from ���ű� D,��������˵�� T" & _
                " where D.ID=T.����ID " & _
                " and (D.����ʱ�� is null or D.����ʱ��=to_date('3000-01-01','YYYY-MM-DD')) And t.�������<>0 " & strTmp
        If cmbStationNo.Text <> "" Then
            strվ�� = Mid(cmbStationNo.Text, 1, InStr(1, cmbStationNo.Text, "-") - 1)
            gstrSql = gstrSql & " And (D.վ��=[2] Or D.վ�� is Null)"
        End If
                
        If str�������� <> "��������" Then
            gstrSql = gstrSql & " and ��������=[1] "
        Else
            gstrSql = gstrSql & " and �������� In('�ٴ�','���','����','����','����'" & IIf(chk�������(2).Value = 1, ",'���'", "") & ") "
        End If
                
        gstrSql = gstrSql & " order by ����"
    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, str��������, strվ��)
    
    Me.lvwItems.ListItems.Clear
    
    Me.lvwItems.Checkboxes = True
   
    Do Until rsData.EOF
        Set objItem = Me.lvwItems.ListItems.Add(, "_" & rsData!ID, rsData!����)
        objItem.Icon = "Dept": objItem.SmallIcon = "Dept"
        objItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = rsData!����
        objItem.Checked = False
        If Me.lvwItems.Tag = "����" Then
            If InStr(Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, col���˿���ID) & ",", rsData!ID & ",") > 0 Then
                objItem.Checked = True
            End If
        End If
        
        If Me.lvwItems.Tag = "ִ��" Then
            If InStr(mstr��ѡִ�п���, rsData!ID & "," & rsData!����) > 0 Then
                objItem.Checked = True
            End If
        End If
        
        If Me.lvwItems.Tag = "ʹ��" Then
            If InStr(mstr��ѡʹ�ÿ���, rsData!ID & "," & rsData!����) > 0 Then
                objItem.Checked = True
            End If
        End If
        
        rsData.MoveNext
    Loop
    rsData.Close
    
    'û��ʱ�˳�
    If Me.lvwItems.ListItems.Count = 0 Then Exit Sub
    
    Me.lvwItems.ListItems(1).Selected = True
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Load�Թܱ���(Optional ByVal strCode As String = "")
    Dim rsTemp As Recordset
    Dim strTmp As String
    Dim i As Integer, intIndex As Integer
    'ȡ�������ʷ��࣬����Ѿ���ȡ�����˳�
    On Error GoTo ErrHandle
    With Me.cboTestTubeCode
        If strCode = "" Then
            If .ListCount > 0 Then strTmp = .Text
        Else
            strTmp = strCode
        End If
        If .ListCount <= 1 Then
            gstrSql = "Select ���� || '-' || ���� ����,��ɫ  From ��Ѫ������"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "��Ѫ������")
            .Clear
            .AddItem "<δ����>": .ItemData(0) = 0
            Do While Not rsTemp.EOF
                i = i + 1
                .AddItem rsTemp!����
                .ItemData(.NewIndex) = Val(rsTemp!��ɫ)
                If Split(rsTemp!����, "-")(0) = strTmp Or rsTemp!���� = strTmp Then
                    intIndex = i
                End If
                rsTemp.MoveNext
            Loop
            .ListIndex = intIndex
        Else
            For i = 1 To .ListCount - 1
                If Split(.List(i), "-")(0) = strTmp Or .List(i) = strTmp Then
                    intIndex = i
                    Exit For
                End If
            Next
            .ListIndex = intIndex
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cboProperty_Click()
    If picDept.Tag = "2" Then
        Load���� 2, cboProperty.Text
    Else
        If Me.msf����ִ��.Col = col���˿��� Then
            Load���� 1, cboProperty.Text
        Else
            Load���� 0, cboProperty.Text
        End If
    End If
    
    ChkSelect.Value = 0
End Sub

Private Sub cboTestTubeCode_Click()
    If cboTestTubeCode.ListIndex > 0 And cboTestTubeCode.ListIndex < cboTestTubeCode.ListCount - 1 Then
        picTubeColor.Visible = True
        picTubeColor.BackColor = Val(cboTestTubeCode.ItemData(cboTestTubeCode.ListIndex))
    Else
        picTubeColor.Visible = False
        picTubeColor.BackColor = picTestTubeCode.BackColor
    End If
End Sub

Private Sub cboTestTubeCode_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo¼��������Χ_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cboִ�з���_Click()
    Me.lbl��Һ����.Visible = False
    Me.cbo��Һ����.Visible = False
    
    If Left(Me.cboִ�з���.Text, 1) = "1" Then      '��Һ
        Me.lbl��Һ����.Visible = True
        Me.cbo��Һ����.Visible = True
    End If
End Sub


Private Sub ChkSelect_Click()
    Dim i As Integer
    Dim str���� As String
    
    On Error GoTo ErrHandle
    If mblnRefresh = True Then Exit Sub
    
    If ChkSelect.Value = 2 Then Exit Sub
    Call SetSelect(lvwItems, ChkSelect.Value)
    
    If cboProperty.Text = "��������" Then
        If lvwItems.Tag = "ִ��" Then
            mstr��ѡִ�п��� = ""
        ElseIf lvwItems.Tag = "ʹ��" Then
            mstr��ѡʹ�ÿ��� = ""
        End If
    End If
    
    If ChkSelect.Value = 1 Then
        '��ǰ����ȫѡ
        For i = 1 To lvwItems.ListItems.Count
            str���� = Mid(lvwItems.ListItems(i).Key, 2) & "," & lvwItems.ListItems(i).Text
            
            If InStr(mstr��ѡִ�п���, str����) = 0 Or cboProperty.Text = "��������" Then
                If lvwItems.Tag = "ִ��" Then
                    mstr��ѡִ�п��� = IIf(mstr��ѡִ�п��� = "", "", mstr��ѡִ�п��� & ";") & str����
                ElseIf lvwItems.Tag = "ʹ��" Then
                    mstr��ѡʹ�ÿ��� = IIf(mstr��ѡʹ�ÿ��� = "", "", mstr��ѡʹ�ÿ��� & ";") & str����
                End If
            End If
        Next
    ElseIf cboProperty.Text <> "��������" Then
        '��ǰ����ȫ��

        For i = 1 To lvwItems.ListItems.Count
            str���� = Mid(lvwItems.ListItems(i).Key, 2) & "," & "[" & lvwItems.ListItems(i).SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) & "]" & lvwItems.ListItems(i).Text
            If lvwItems.Tag = "ִ��" Then
                If InStr(mstr��ѡִ�п���, str����) > 0 Then
                    mstr��ѡִ�п��� = Replace(mstr��ѡִ�п���, str����, "")
                End If
            ElseIf lvwItems.Tag = "ʹ��" Then
                If InStr(mstr��ѡʹ�ÿ���, str����) > 0 Then
                    mstr��ѡʹ�ÿ��� = Replace(mstr��ѡʹ�ÿ���, str����, "")
                End If
            End If
        Next
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub SetSelect(ByVal lvwObj As Object, Optional ByVal BlnSelect As Boolean = True)
    Dim intSelect As Integer
    With lvwObj
        For intSelect = 1 To .ListItems.Count
            .ListItems(intSelect).Checked = BlnSelect
        Next
    End With
End Sub

Private Sub cmbStationNo_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
End Sub

Private Sub cmdDel�ο�_Click()
    Me.txt�ο�.Text = ""
    Me.txt�ο�.Tag = ""
End Sub

Private Sub cmdFind_Click()
    Dim strFind As String
    Dim i As Long
    Dim blnIsFind As Boolean
    
    strFind = UCase(Trim(txtFind.Text))
    If strFind = "" Then Exit Sub
    For i = mlngFind To lvwItems.ListItems.Count
        If zlCommFun.SpellCode(Mid(lvwItems.ListItems(i).Text, InStr(lvwItems.ListItems(i).Text, "-") + 1)) Like UCase(IIf(gstrMatch <> "", "*", "") & strFind & "*") Or _
                UCase(lvwItems.ListItems(i).Text) Like UCase(IIf(gstrMatch <> "", "*", "") & strFind & "*") Then
            lvwItems.ListItems(i).Selected = True
            lvwItems.ListItems(i).EnsureVisible
            blnIsFind = True
            mlngFind = i + 1
            Exit For
        End If
    Next
    If blnIsFind = False Then
        If mlngFind = 1 Then
            MsgBox "û���ҵ������ҵĿ��ҡ�", vbInformation, Me.Caption
        Else
            MsgBox "�Ѿ������һ�������ˡ�", vbInformation, Me.Caption
            mlngFind = 1
        End If
    End If
End Sub

Private Sub cmdFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        picDept.Visible = False
        txtFind.Text = ""
    End If
End Sub

Private Sub cmdFindCancle_Click()
    Call lvwItems_KeyPress(vbKeyEscape)
End Sub

Private Sub cmdFindOk_Click()
    Call lvwItems_DblClick
End Sub

Private Sub cmdѡ��_Click()
    Dim rsTemp As Recordset

    On Error GoTo ErrHand
    gstrSql = "select A.ID,A.����,A.�������� ��������,A.����,A.����" & _
            " from ��������Ŀ¼ A" & _
            " where A.���='S' and (A.����ʱ�� is null or A.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)

    With rsTemp
        If .RecordCount = 0 Then
            MsgBox "δ�ҵ�ָ��������׼����", vbExclamation, gstrSysName
            Me.txt��׼����.SetFocus
            Exit Sub
        End If
        If .RecordCount = 1 Then
            Me.txt��׼����.Tag = !ID
            Me.txt��׼����.Text = IIf(IsNull(!����), "", !����)
            Me.lbl��׼����.Caption = IIf(IsNull(!��������), "", "��" & NVL(!��������) & "��") & IIf(IsNull(!����), "", !����)
            Me.stbInfo.Tab = 1: Me.chk�������(0).SetFocus
            Exit Sub
        End If

        Me.lvwItem.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItem.ListItems.Add(, "_" & !ID, !����, "expend", "expend")
            objItem.SubItems(Me.lvwItem.ColumnHeaders("����").Index - 1) = !����
            objItem.SubItems(Me.lvwItem.ColumnHeaders("���").Index - 1) = NVL(!��������)
            .MoveNext
        Loop
        With Me.lvwItem
            .ListItems(1).Selected = True
            .Tag = "����"
            .Left = Me.stbInfo.Left + Me.fra��׼����.Left + Me.fra��׼����.Width - .Width
            .Top = Me.stbInfo.Top + Me.fra��׼����.Top + Me.txt��׼����.Top + Me.txt��׼����.Height
            .ZOrder 0: .Visible = True
            .SetFocus
        End With
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub lvwItems_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim str���� As String
    
    If Me.lvwItems.Tag = "ִ��" Then
        str���� = Mid(Item.Key, 2) & "," & Item.Text
        
        If Item.Checked = True Then
            If InStr(mstr��ѡִ�п���, str����) = 0 Then
                mstr��ѡִ�п��� = IIf(mstr��ѡִ�п��� = "", "", mstr��ѡִ�п��� & ";") & str����
            End If
        Else
            If InStr(mstr��ѡִ�п���, str����) > 0 Then
                mstr��ѡִ�п��� = Replace(mstr��ѡִ�п���, str����, "")
            End If
        End If
    ElseIf Me.lvwItems.Tag = "ʹ��" Then
        str���� = Mid(Item.Key, 2) & "," & Item.Text
        
        If Item.Checked = True Then
            If InStr(mstr��ѡʹ�ÿ���, str����) = 0 Then
                mstr��ѡʹ�ÿ��� = IIf(mstr��ѡʹ�ÿ��� = "", "", mstr��ѡʹ�ÿ��� & ";") & str����
            End If
        Else
            If InStr(mstr��ѡʹ�ÿ���, str����) > 0 Then
                mstr��ѡʹ�ÿ��� = Replace(mstr��ѡʹ�ÿ���, str����, "")
            End If
        End If
    End If
End Sub

Private Sub lvwItems_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mlngFind = Item.Index + 1
End Sub

Private Sub lvwItems_KeyDown(KeyCode As Integer, Shift As Integer)
    If Me.lvwItems.Tag = "����" Or Me.lvwItems.Tag = "ִ��" Or Me.lvwItems.Tag = "ʹ��" Then
        If KeyCode = vbKeyA And Shift = vbCtrlMask Then 'ȫѡ Ctrl+A
            If Me.lvwItems.Tag = "ִ��" Or Me.lvwItems.Tag = "ʹ��" Then
                If Me.ChkSelect.Value = 0 Then
                    Me.ChkSelect.Value = 1
                    Call SetSelect(lvwItems, True)
                End If
            Else
                Call SetSelect(lvwItems, True)
            End If
        End If
        
        If KeyCode = vbKeyR And Shift = vbCtrlMask Then     'ȫ�� Ctrl+R
            If Me.lvwItems.Tag = "ִ��" Or Me.lvwItems.Tag = "ʹ��" Then
                If Me.ChkSelect.Value = 1 Then
                    Me.ChkSelect.Value = 0
                    Call SetSelect(lvwItems, False)
                End If
            Else
                Call SetSelect(lvwItems, False)
            End If
        End If
    End If
End Sub
Private Sub cboProperty_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyEscape
         picDept.Visible = False
         txtFind.Text = ""
    End Select
End Sub

Private Sub cboProperty_LostFocus()
    Call picDept_LostFocus
End Sub

Private Sub msf����ִ��_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If msf����ִ��.Editable = flexEDNone Then
        msf����ִ��.FocusRect = flexFocusLight
        msf����ִ��.ComboList = ""
    Else
        msf����ִ��.FocusRect = flexFocusSolid
        msf����ִ��.ComboList = "..."
    End If
End Sub

Private Sub msf����ִ��_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    msf����ִ��.AutoSize msf����ִ��.FixedCols, msf����ִ��.Cols - 1
End Sub

Private Sub msf����ִ��_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If msf����ִ��.TextMatrix(NewRow, OldCol) <> msf����ִ��.Cell(flexcpData, NewRow, OldCol) Then
        msf����ִ��.TextMatrix(NewRow, OldCol) = msf����ִ��.Cell(flexcpData, NewRow, OldCol)
    End If
End Sub

Private Sub msf����ִ��_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As New ADODB.Recordset
    Dim objItem As ListItem
    Dim i As Integer
    
    mstr��ѡִ�п��� = ""
    If Me.msf����ִ��.Col = colִ�п��� Then
        With Me.msf����ִ��
            For i = 1 To .Rows - 1
                If .TextMatrix(i, colִ�п���) <> "" Then
                    mstr��ѡִ�п��� = IIf(mstr��ѡִ�п��� = "", "", mstr��ѡִ�п��� & ";") & .TextMatrix(i, colִ�п���ID) & "," & .TextMatrix(i, colִ�п���)
                End If
            Next
        End With
    End If
    
    With Me.picDept
        If Me.msf����ִ��.Col = col���˿��� Then
            .Tag = ""
            Me.lvwItems.Tag = "����"
            .Left = Me.fraִ�в���.Left + Me.msf����ִ��.Left + Me.msf����ִ��.ColWidth(colִ�п���ID)
            .Width = IIf(Me.msf����ִ��.ColWidth(col���˿���) < 3000, 3000, Me.msf����ִ��.ColWidth(col���˿���))
        Else
            .Tag = "1"
            Me.lvwItems.Tag = "ִ��"
            .Left = Me.fraִ�в���.Left + Me.msf����ִ��.Left + Me.msf����ִ��.ColWidth(colִ�п���ID) + Me.msf����ִ��.ColWidth(col���˿���) + Me.msf����ִ��.ColWidth(col���˿���ID)
            .Width = IIf(Me.msf����ִ��.ColWidth(col���˿���ID) < 5000, 5000, Me.msf����ִ��.ColWidth(col���˿���ID))
            If .Left > Me.Width - .Width - stbInfo.Left - Me.fraִ�в���.Left - Me.msf����ִ��.Left Then .Left = Me.Width - .Width - stbInfo.Left - Me.fraִ�в���.Left - Me.msf����ִ��.Left
        End If
        
        .Top = 50
        .Height = Me.fraִ�в���.Top + Me.msf����ִ��.Top + (IIf(Me.msf����ִ��.Row > 14, 14, Me.msf����ִ��.Row) - Me.msf����ִ��.FixedRows + 1) * Me.msf����ִ��.RowHeight(col���˿���)
        
        lbl��������.Visible = (Me.msf����ִ��.Col = colִ�п���)
        cboProperty.Visible = lbl��������.Visible
        ChkSelect.Visible = lbl��������.Visible
        
        If Me.lvwItems.Tag = "ִ��" Then
            lbl��������.Left = 50
            ChkSelect.Left = .Width - ChkSelect.Width - 50
            cboProperty.Width = ChkSelect.Left - cboProperty.Left - 50
        End If
        
        cmdFind.Visible = True
        txtFind.Visible = True
        cmdFindOk.Visible = True
        cmdFindCancle.Visible = True
        .ZOrder 0
        .Visible = True
    End With

    With Me.lvwItems
        If .Tag = "ִ��" Then
            .Left = lbl��������.Left
            .Top = cboProperty.Top + cboProperty.Height + 50 + txtFind.Height + 50
            .Width = Me.picDept.Width - .Left - 50
            .Height = Me.picDept.Height - .Top - 10
            txtFind.Top = cboProperty.Top + cboProperty.Height + 50
            cmdFind.Top = cboProperty.Top + cboProperty.Height + 50
            cmdFindOk.Left = .Width + .Left - cmdFind.Width - 80 - cmdFindCancle.Width
            cmdFindCancle.Left = .Width + .Left - cmdFind.Width - 50
            cmdFindOk.Top = cmdFind.Top
            cmdFindCancle.Top = cmdFind.Top
        Else
            .Left = 0
            .Top = txtFind.Height + 100
            .Width = Me.picDept.Width
            .Height = Me.picDept.Height - txtFind.Height - 50 - 50
            txtFind.Top = 50
            cmdFind.Top = 50
            cmdFindOk.Left = .Width + .Left - cmdFind.Width - 80 - cmdFindCancle.Width
            cmdFindCancle.Left = .Width + .Left - cmdFind.Width - 50
            cmdFindOk.Top = cmdFind.Top
            cmdFindCancle.Top = cmdFind.Top
        End If
        
        .SetFocus
        .Refresh
    End With
    
    If Me.msf����ִ��.Col = col���˿��� Then
        load���ʷ��� 1
    Else
        load���ʷ��� 0
    End If
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub msf����ִ��_EnterCell()
    strInputed = Me.msf����ִ��.TextMatrix(msf����ִ��.Row, msf����ִ��.Col)
End Sub

Private Sub msf����ִ��_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTmp As New ADODB.Recordset
    If KeyCode > 127 Then
        '���ֱ�����뺺�ֵ�����
        Call msf����ִ��_KeyPress(KeyCode)
    ElseIf KeyCode = vbKeyDelete Then
        If msf����ִ��.TextMatrix(msf����ִ��.Row, msf����ִ��.Col) <> "" Then
            If MsgBox("��ȷ��Ҫɾ����һ�����ݣ�", vbYesNo + vbQuestion + vbDefaultButton1, Me.Caption) = vbYes Then
                If msf����ִ��.Rows <= 2 Then
                    msf����ִ��.Cell(flexcpText, 1, col���˿���ID, 1, colִ�п���) = ""
                    msf����ִ��.Cell(flexcpData, 1, col���˿���ID, 1, colִ�п���) = ""
                Else
                    msf����ִ��.RemoveItem msf����ִ��.Row
                End If
            End If
        End If
    End If
    If KeyCode <> vbKeyReturn Then Exit Sub
    With Me.msf����ִ��
        If .Editable = flexEDNone Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        If .Col = colִ�п��� And .TextMatrix(.Row, colִ�п���) = "" Then
            If .Row = 1 Then Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        If .Col = col���˿��� And .TextMatrix(.Row, col���˿���) = "" Then
            .TextMatrix(.Row, col���˿���) = "�����в��ţ�"
            .TextMatrix(.Row, col���˿���ID) = "�����в��ţ�"
            Exit Sub
        End If
    End With
End Sub

Private Sub msf����ִ��_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim rsTmp As New ADODB.Recordset
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    With Me.msf����ִ��
        If .Editable = flexEDNone Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        If .Col = colִ�п��� And Trim(.EditText) = "" Then
            If .Row = 1 Then .SetFocus: Call zlCommFun.PressKey(vbKeyTab)
            Exit Sub
        End If
        strTemp = UCase(Trim(.EditText))
    End With
    If strTemp = strInputed Then Exit Sub
    If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
    
    If strTemp = "" Then Exit Sub
    
    err = 0: On Error GoTo ErrHand

    If Me.msf����ִ��.Col = col���˿��� Then
        gstrSql = "select distinct ID,����,����" & _
                " from ���ű� D,��������˵�� T" & _
                " where D.ID=T.����ID and ��������='�ٴ�'" & _
                "       and (D.����ʱ�� is null or D.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))" & _
                "       and (D.���� like [1] or D.���� like [1] or D.���� like [1])" & _
                " order by ����"
    Else
        gstrSql = "select distinct ID,����,����" & _
                " from ���ű� D,��������˵�� T" & _
                " where D.ID=T.����ID and T.������� in (1,2,3)" & _
                "       and (D.����ʱ�� is null or D.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))" & _
                "       and (D.���� like [1] or D.���� like [1] or D.���� like [1])" & _
                " order by ����"
    End If

    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, gstrMatch & strTemp & "%")

    With rsTmp
        If .BOF Or .EOF Then
            MsgBox "δ�ҵ�ָ�����ţ����������룡", vbExclamation, gstrSysName
            Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, msf����ִ��.Col) = msf����ִ��.Cell(flexcpData, msf����ִ��.Row, msf����ִ��.Col)
            msf����ִ��.EditText = msf����ִ��.Cell(flexcpData, msf����ִ��.Row, msf����ִ��.Col)
            Exit Sub
        End If
        If .RecordCount = 1 Then
            Me.msf����ִ��.Text = !����
            If Me.msf����ִ��.Col = colִ�п��� Then
                Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, colִ�п���ID) = !ID
                Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, colִ�п���) = Me.msf����ִ��.Text
                msf����ִ��.EditText = Me.msf����ִ��.Text
                msf����ִ��.Cell(flexcpData, msf����ִ��.Row, msf����ִ��.Col) = Me.msf����ִ��.Text
            Else
                Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, col���˿���ID) = !ID
                Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, col���˿���) = Me.msf����ִ��.Text
                msf����ִ��.EditText = Me.msf����ִ��.Text
                msf����ִ��.Cell(flexcpData, msf����ִ��.Row, msf����ִ��.Col) = Me.msf����ִ��.Text
            End If
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Me.lvwItems.Checkboxes = (Me.msf����ִ��.Col = col���˿���)
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !����)
            objItem.Icon = "Dept": objItem.SmallIcon = "Dept"
            objItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = !����

            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    
    With Me.picDept
        If Me.msf����ִ��.Col = col���˿��� Then
            .Tag = ""
            Me.lvwItems.Tag = "����"
            .Left = Me.fraִ�в���.Left + Me.msf����ִ��.Left
            .Width = IIf(Me.msf����ִ��.ColWidth(col���˿���) < 3000, 3000, Me.msf����ִ��.ColWidth(col���˿���))
        Else
            .Tag = "1"
            Me.lvwItems.Tag = "ִ��"
            .Left = Me.fraִ�в���.Left + Me.msf����ִ��.Left + Me.msf����ִ��.ColWidth(colִ�п���ID) + Me.msf����ִ��.ColWidth(col���˿���) + Me.msf����ִ��.ColWidth(col���˿���ID)
            .Width = IIf(Me.msf����ִ��.ColWidth(col���˿���ID) < 3000, 3000, Me.msf����ִ��.ColWidth(col���˿���ID))
            If .Left > Me.Width - .Width - stbInfo.Left - Me.fraִ�в���.Left - Me.msf����ִ��.Left Then .Left = Me.Width - .Width - stbInfo.Left - Me.fraִ�в���.Left - Me.msf����ִ��.Left
        End If
        
        .Top = 50
        .Height = Me.fraִ�в���.Top + Me.msf����ִ��.Top + (IIf(Me.msf����ִ��.Row > 14, 14, Me.msf����ִ��.Row) - Me.msf����ִ��.FixedRows + 1) * Me.msf����ִ��.RowHeight(col���˿���)
        
        lbl��������.Visible = False
        cboProperty.Visible = lbl��������.Visible
        ChkSelect.Visible = lbl��������.Visible
        
        If Me.msf����ִ��.Col = colִ�п��� Then
            lbl��������.Left = 50
            ChkSelect.Left = .Width - ChkSelect.Width - 50
            cboProperty.Width = ChkSelect.Left - cboProperty.Left - 50
        End If
        
        txtFind.Visible = False
        cmdFind.Visible = False
        cmdFindOk.Visible = False
        cmdFindCancle.Visible = False
        .ZOrder 0
        .Visible = True
    End With
    
    With Me.lvwItems
        .Left = 0
        .Top = 0
        .Width = Me.picDept.Width
        .Height = Me.picDept.Height
        
        .SetFocus
        .Refresh
    End With
    KeyCode = 0
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub msf����ִ��_KeyPress(KeyAscii As Integer)
    If msf����ִ��.Editable = flexEDNone Then Exit Sub

    With msf����ִ��
        If KeyAscii = 13 Then
            KeyAscii = 0
            If .Col = colִ�п��� Then
                If .Row = .Rows - 1 Then
                    If (.TextMatrix(.Row, colִ�п���) <> "" Or .TextMatrix(.Row, col���˿���) <> "") Then
                        .Rows = .Rows + 1
                    Else
                        zlCommFun.PressKey (vbKeyTab)
                        Exit Sub
                    End If
                End If
                .Row = .Row + 1
                .Col = col���˿���
            ElseIf .Col = col���˿��� Then
                .Col = colִ�п���
            End If
        Else
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                Call msf����ִ��_CellButtonClick(.Row, .Col)
            Else
                If KeyAscii = vbKeyBack Then Exit Sub
                .ComboList = "" 'ʹ��ť״̬��������״̬
            End If
        End If
    End With
End Sub

Private Sub msf����ִ��_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    If msf����ִ��.Editable = flexEDNone Then
        msf����ִ��.FocusRect = flexFocusLight
        msf����ִ��.ComboList = ""
    Else
        msf����ִ��.FocusRect = flexFocusSolid
        msf����ִ��.ComboList = "..."
    End If
End Sub

Private Sub msf����ִ��_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Call msf����ִ��_KeyDownEdit(Row, Col, vbKeyReturn, 0)
End Sub

Private Sub OptApp_Click(Index As Integer)
    Dim i As Integer
    
    For i = 1 To OptApp.UBound
        If i = Index Then
            OptApp(i).FontBold = True
        Else
            OptApp(i).FontBold = False
        End If
    Next
End Sub

Private Sub OptAppUse_Click(Index As Integer)
    Dim i As Integer
    
    For i = 1 To OptAppUse.UBound
        If i = Index Then
            OptAppUse(i).FontBold = True
        Else
            OptAppUse(i).FontBold = False
        End If
    Next
End Sub

Private Sub optDeptKind_Click(Index As Integer)
    lblLocate.Tag = ""
End Sub

Private Sub picDept_LostFocus()
    Dim strActive As String
    
    strActive = UCase(Me.ActiveControl.Name)
    
    If InStr(1, "CMDOKDEPT,CMDCANCELDEPT,LVWITEMS,CBOPROPERTY,PICDEPT,CHKSELECT,TXTFIND,CMDFIND,CMDFINDOK", strActive) <> 0 Then
        Exit Sub
    End If

    picDept.Visible = False
    If Me.lvwItems.Tag = "ʹ��" Then
        vsUseDept.TextMatrix(vsUseDept.Row, vsUseDept.Col) = vsUseDept.Cell(flexcpData, vsUseDept.Row, vsUseDept.Col)
        vsUseDept.AutoSize vsUseDept.FixedCols, vsUseDept.Cols - 1
        Call vsUseDept.SetFocus
    Else
        msf����ִ��.TextMatrix(msf����ִ��.Row, msf����ִ��.Col) = msf����ִ��.Cell(flexcpData, msf����ִ��.Row, msf����ִ��.Col)
        msf����ִ��.AutoSize msf����ִ��.FixedCols, msf����ִ��.Cols - 1
        Call msf����ִ��.SetFocus
    End If
    txtFind.Text = ""
    mlngFind = 1
End Sub
Private Sub ChkSelect_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyEscape
         picDept.Visible = False
         txtFind.Text = ""
    End Select
End Sub
Private Sub GetDefineSize()
    '���ܣ��õ����ݿ�ı��ֶεĳ���
    On Error GoTo ErrHandle
    Dim rsTmp As New ADODB.Recordset
    gstrSql = "Select A.����, A.���㵥λ, B.����, B.���� �������� From ������ĿĿ¼ A, ������Ŀ���� B " & _
            " Where A.ID = B.������Ŀid And A.ID = 0 And B.���� = 1"
    Call zlDatabase.OpenRecordset(rsTmp, gstrSql, Me.Caption)

    mlng���볤�� = rsTmp.Fields("����").DefinedSize

    txt��Ŀ����.MaxLength = rsTmp.Fields("����").DefinedSize
    txt���㵥λ.MaxLength = rsTmp.Fields("���㵥λ").DefinedSize
    txt����ƴ��.MaxLength = mlng���볤��
    txt�������.MaxLength = mlng���볤��
    txt����ƴ��.MaxLength = mlng���볤��
    txt�������.MaxLength = mlng���볤��
    txt��������.MaxLength = rsTmp.Fields("��������").DefinedSize
    txt��Ŀ����.MaxLength = rsTmp.Fields("��������").DefinedSize
    txtӢ��.MaxLength = 40
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
Private Sub InitVsfFreq()
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    
    On Error GoTo ErrHandle
    With vsfFreq
        '��ʼ�����
        .Clear
        .FixedCols = 0
        .FixedRows = 1
        .Rows = 1
        .Cols = 7
        
        .RowHeightMin = 300
        
        .TextMatrix(0, 0) = "ѡ��"
        .TextMatrix(0, 1) = "����"
        .TextMatrix(0, 2) = "����"
        .TextMatrix(0, 3) = "Ӣ������"
        .TextMatrix(0, 4) = "Ƶ�ʴ���"
        .TextMatrix(0, 5) = "Ƶ�ʼ��"
        .TextMatrix(0, 6) = "�����λ"
        
        .ColWidth(0) = 500
        .ColWidth(1) = 0
        .ColWidth(2) = 2000
        .ColWidth(3) = 1500
        .ColWidth(4) = 1000
        .ColWidth(5) = 1000
        .ColWidth(6) = 1000
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .ColAlignment(0) = flexAlignCenterCenter
        
        .Editable = flexEDNone
        
        '��ȡ����,�����
        gstrSql = "Select A.��Ŀid, B.����, B.����, B.Ӣ������, B.Ƶ�ʴ���, B.Ƶ�ʼ��, B.�����λ " & _
            " From (Select ��Ŀid, Ƶ�� From �����÷����� Where ��Ŀid = [1]) A, ����Ƶ����Ŀ B " & _
            " Where A.Ƶ��(+) = B.���� And B.���÷�Χ = 1 " & _
            " Order By A.��Ŀid, B.����"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption & " ������ĿƵ��", lngItemID)
        
        Do While Not rsTmp.EOF
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = IIf(IsNull(rsTmp!��Ŀid), "", "��")
            .TextMatrix(.Rows - 1, 1) = rsTmp!����
            .TextMatrix(.Rows - 1, 2) = rsTmp!����
            .TextMatrix(.Rows - 1, 3) = IIf(IsNull(rsTmp!Ӣ������), "", rsTmp!Ӣ������)
            .TextMatrix(.Rows - 1, 4) = rsTmp!Ƶ�ʴ���
            .TextMatrix(.Rows - 1, 5) = rsTmp!Ƶ�ʼ��
            .TextMatrix(.Rows - 1, 6) = IIf(IsNull(rsTmp!�����λ), "", rsTmp!�����λ)
            rsTmp.MoveNext
        Loop
    End With
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub InivsfTest()
    Dim rsTemp As ADODB.Recordset
    Dim strTest As String
    Dim strTemp As String
    Dim strArr
    Dim n As Integer
    Const strDefault As String = "����(+);����(-)"
    
    On Error GoTo ErrHandle
    With vsfTest
        .Clear
        .Cols = 3
        .Rows = 1
        
        .TextMatrix(0, 0) = "��ע"
        .TextMatrix(0, 1) = "����"
        .TextMatrix(0, 2) = "����"
        
        .ColWidth(0) = 1000
        .ColWidth(1) = 1500
        .ColWidth(2) = 800
        
        .FixedAlignment(0) = flexAlignCenterCenter
        .FixedAlignment(1) = flexAlignCenterCenter
        .FixedAlignment(2) = flexAlignCenterCenter
        
        .ColAlignment(0) = flexAlignLeftCenter
        .ColAlignment(1) = flexAlignLeftCenter
        .ColAlignment(2) = flexAlignCenterCenter
        
        strTest = strDefault
        
        If lngItemID > 0 Then
            gstrSql = "Select �걾��λ From ������ĿĿ¼ Where ID = [1] "
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "ȡƤ�Խ��", lngItemID)

            If rsTemp.RecordCount > 0 Then
                strTemp = IIf(IsNull(rsTemp!�걾��λ), "", rsTemp!�걾��λ)
                
                If strTemp <> "" And InStrB(strTemp, ";") > 0 Then
                    strTest = strTemp
                End If
            End If
        End If
        
        '���Խ��
        strArr = Split(Split(strTest, ";")(1), ",")
        For n = 0 To UBound(strArr)
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = MidB(strArr(n), InStrB(strArr(n), "("))
            .TextMatrix(.Rows - 1, 1) = MidB(strArr(n), 1, InStrB(strArr(n), "(") - 1)
        Next
        
        '���Խ��
        strArr = Split(Split(strTest, ";")(0), ",")
        For n = 0 To UBound(strArr)
            .Rows = .Rows + 1
            .TextMatrix(.Rows - 1, 0) = MidB(strArr(n), InStrB(strArr(n), "("))
            .TextMatrix(.Rows - 1, 1) = MidB(strArr(n), 1, InStrB(strArr(n), "(") - 1)
            .TextMatrix(.Rows - 1, 2) = "��"
        Next
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Public Function ShowMe(ByVal frmParent As Object, ByVal byt״̬ As Byte, ByVal lng����id As Long, Optional ByVal lng��Ŀid As Long) As Boolean
    '---------------------------------------------------
    '���ܣ��ϼ�������ñ�����ģ����ݲ���������ʾ����
    '---------------------------------------------------
    Me.Tag = Switch(byt״̬ = 0, "����", byt״̬ = 1, "�޸�", byt״̬ = 2, "����", byt״̬ = 3, "��������")
    lngClassId = lng����id: lngItemID = lng��Ŀid: lngVItemID = 0
    mlngOldId = lng��Ŀid
    
    '��д��Ҫѡ�������
    aryTemp = Split("0-��ѡƵ��;1-һ����;2-������", ";")
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cboִ��Ƶ��.AddItem aryTemp(intCount)
    Next
    Me.cboִ��Ƶ��.ListIndex = 0

    aryTemp = Split("0-����ȷ;1-����;2-��ʱ;3-�ƴ�", ";")
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo���㷽ʽ.AddItem aryTemp(intCount)
    Next
    Me.cbo���㷽ʽ.ListIndex = 0

    aryTemp = Split("0-���Ա�����;1-����;2-Ů��", ";")
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo�����Ա�.AddItem aryTemp(intCount)
    Next
    Me.cbo�����Ա�.ListIndex = 0
    
    aryTemp = Split("0-��������;1-ȡ������", ";")
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cbo�������.AddItem aryTemp(intCount)
    Next
    Me.cbo�������.ListIndex = 0

    err = 0: On Error GoTo ErrHand
    
    gstrSql = "select ID,�ϼ�ID,����,����,����" & _
            " From ���Ʒ���Ŀ¼" & _
            " Where ���� = 5" & _
            " start with �ϼ�ID is null" & _
            " connect by prior ID=�ϼ�ID"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "ShowMe")
'        Call SQLTest
    With rsTemp
        If .BOF Or .EOF Then MsgBox "�����Ƚ������Ʒ�����Ŀ֮��������Ŀ", vbExclamation, gstrSysName: Unload Me: Exit Function
        Me.tvwClass.Nodes.Clear
        Do While Not .EOF
            If IsNull(!�ϼ�ID) Then
                Set objNode = Me.tvwClass.Nodes.Add(, , "_" & !ID, "[" & !���� & "]" & !����, "close")
            Else
                Set objNode = Me.tvwClass.Nodes.Add("_" & !�ϼ�ID, tvwChild, "_" & !ID, "[" & !���� & "]" & !����, "close")
            End If
            objNode.Sorted = True
            objNode.Tag = IIf(IsNull(!����), "", !����)
            objNode.ExpandedImage = "expend"
            .MoveNext
        Loop
        Me.tvwClass.Nodes("_" & lng����id).Selected = True
        Me.txt����.Text = Me.tvwClass.SelectedItem.Text
        Me.txt����.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
    End With
    gstrSql = "select ����||'-'||���� from ������Ŀ��� where ����>'9' order by ����"
'        If .State = adStateOpen Then .Close
'        Call SQLTest(App.ProductName, Me.Caption, gstrSql)
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "ShowMe")
'        Call SQLTest
    With rsTemp
        If .BOF Or .EOF Then MsgBox "������Ŀ������ݶ�ʧ(��ϵͳ����Ա����)", vbExclamation, gstrSysName: Unload Me: Exit Function
        Me.cbo���.Clear
        Do While Not .EOF
            Me.cbo���.AddItem .Fields(0).Value
            .MoveNext
        Loop
        If Me.cbo���.ListCount > 0 Then Me.cbo���.ListIndex = 0
    End With
    'Me.cbo��Ŀ����.ListIndex = 0: Me.cbo�������.ListIndex = 0
    
    'ȡ��ҩ;���ķ���˵��
    With Me.cbo����˵��
        .Clear
        gstrSql = "Select Distinct �걾��λ From ������ĿĿ¼ Where ��� = 'E' And �������� = '2' And �걾��λ Is Not Null"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "ȡ��ҩ;���ķ���˵��")
        
        If rsTemp.RecordCount > 0 Then
            Do While Not rsTemp.EOF
                .AddItem rsTemp.Fields(0).Value
                rsTemp.MoveNext
            Loop
        End If
    End With
    
    '��ʾ����
    Me.Show 1, frmParent
    ShowMe = mblnOK
    Exit Function

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cbo��������_Click()
    Dim i As Long
    
    stbInfo.TabVisible(3) = False
    Me.lbl����˵��.Visible = False
    Me.cbo����˵��.Visible = False
    Me.fra¼����.Visible = False
    cboִ�з���.Visible = False
    cboBloodType.Visible = False
    lblִ�з���.Visible = False
    Me.lbl��Һ����.Visible = False
    Me.cbo��Һ����.Visible = False
    Me.picTestTubeCode.Visible = False
    Me.lbl�Թܱ���.Visible = False
    Me.chkNoTMSY.Visible = False
    Me.chkYYPS.Visible = False
    
    imgComment.Top = lbl��������.Top + 50
    lblComment.Top = lbl��������.Top + 50
                
    If Left(Me.cbo���.Text, 1) = "E" Then      '����
        Select Case Val(Left(Me.cbo��������.Text, 1))
        Case 0, 5     '0-��ͨ;5-��������
            Me.chk����Ӧ��.Enabled = True
            Me.cboִ��Ƶ��.Enabled = True
            Me.cbo���㷽ʽ.Enabled = True
            Me.optִ�в���(5).Enabled = True
        Case 1      '1-��������
            Me.chk����Ӧ��.Enabled = True
            Me.chkNoTMSY.Visible = True
            Me.chkYYPS.Visible = True
            Me.cboִ��Ƶ��.ListIndex = 1: Me.cboִ��Ƶ��.Enabled = False
            Me.cbo���㷽ʽ.ListIndex = 3: Me.cbo���㷽ʽ.Enabled = False: Me.txt���㵥λ.Text = "��"
            If Me.optִ�в���(5).Value = True Then Me.optִ�в���(5).Value = False: Me.optִ�в���(2).Value = True
            Me.optִ�в���(5).Enabled = False
            stbInfo.TabVisible(3) = True
            Call InivsfTest
        Case 2, 3, 4, 6, 9  '2-��ҩ����(��ҩ);3-��ҩ�巨;4-��ҩ��(��)��;6-�걾�ɼ�;9-��Ѫ�ɼ�
            Me.chk����Ӧ��.Value = 0: Me.chk����Ӧ��.Enabled = False
            Me.cboִ��Ƶ��.ListIndex = 0: Me.cboִ��Ƶ��.Enabled = False
            Me.cbo���㷽ʽ.ListIndex = 3: Me.cbo���㷽ʽ.Enabled = False: Me.txt���㵥λ.Text = "��"
            If Me.optִ�в���(5).Value = True Then Me.optִ�в���(5).Value = False: Me.optִ�в���(2).Value = True
            Me.optִ�в���(5).Enabled = False
            Me.lbl����˵��.Visible = True
            Me.cbo����˵��.Visible = True
            
            If Val(Left(Me.cbo��������.Text, 1)) = 2 Then
                cboִ�з���.Visible = True
                lblִ�з���.Visible = True
                imgComment.Top = lbl���Ƽ���.Top + 50
                lblComment.Top = lbl���Ƽ���.Top + 50
            ElseIf Val(Left(Me.cbo��������.Text, 1)) = 9 Then
                Call Load�Թܱ���
                Me.lbl�Թܱ���.Visible = True
                Me.picTestTubeCode.Visible = True
                Me.cboTestTubeCode.Visible = True
            End If
            
            If cboִ�з���.Visible = True And Val(Left(Me.cboִ�з���.Text, 1)) = 1 Then
                Me.lbl��Һ����.Visible = True
                Me.cbo��Һ����.Visible = True
            End If
        Case 7 '7-��Ѫ����
            Me.cboִ��Ƶ��.ListIndex = 1: Me.cboִ��Ƶ��.Enabled = False
            Me.cbo���㷽ʽ.ListIndex = 3: Me.cbo���㷽ʽ.Enabled = False: Me.txt���㵥λ.Text = "��"
            Me.chk����Ӧ��.Value = 0: Me.chk����Ӧ��.Enabled = False
        Case 8 '8-��Ѫ;��
            Me.cboִ��Ƶ��.ListIndex = 0: Me.cboִ��Ƶ��.Enabled = False
            Me.cbo���㷽ʽ.ListIndex = 3: Me.cbo���㷽ʽ.Enabled = False: Me.txt���㵥λ.Text = "��"
            Me.chk����Ӧ��.Enabled = True: Me.chk����Ӧ��.Value = IIf(Val(Me.chk����Ӧ��.Tag) = 1, 1, 0)
            lblִ�з���.Visible = True
            cboBloodType.Visible = True
        End Select
    End If
    
    If Left(Me.cbo���.Text, 1) = "H" Then
        Select Case Val(Left(Me.cbo��������.Text, 1))
        Case 0
            Me.cboִ��Ƶ��.Enabled = True
        Case 1
            Me.cboִ��Ƶ��.ListIndex = 1
            Me.cboִ��Ƶ��.Enabled = False
        End Select
    End If

'    If mLast�������� <> cbo��������.Text Then
        If Left(Me.cbo���.Text, 1) = "D" Then
            If cbo���.Text = "D-���" And cbo��������.Text = "18-����" Then
                lbl�������.Visible = True
                cbo�������.Visible = True
                stbInfo.TabCaption(2) = "����걾"
                optList(0).Caption = "��ʾ���б걾"
                optList(1).Caption = "����ʾ��ѡ��걾"
            Else
                lbl�������.Visible = False
                cbo�������.Visible = False
                
                If cbo��������.Text <> "18-����" Then
                    stbInfo.TabCaption(2) = "��鲿λ(&L)"
                    optList(0).Caption = "��ʾ���в�λ"
                    optList(1).Caption = "����ʾ��ѡ��λ"
                End If
            End If
            
            '�����Ŀ ��ʾ��λ����
            Call initVfgList
            
        End If
'    End If
    
    If Left(Me.cbo���.Text, 1) = "Z" Then      '����
        Me.cbo��������.Width = Me.fra��׼����.Left + Me.fra��׼����.Width - Me.cbo��������.Left
        Me.txt��Ŀ����.Width = Me.fra��׼����.Left + Me.fra��׼����.Width - Me.txt��Ŀ����.Left
    
        Select Case Val(Mid(Me.cbo��������.Text, 1, InStr(1, Me.cbo��������.Text, "-") - 1))
        Case 0      '0-��ͨ
            Me.cboִ��Ƶ��.Enabled = True: Me.cbo���㷽ʽ.Enabled = True
            Me.chk�������(0).Enabled = True: Me.chk�������(1).Enabled = True: Me.chk�������(2).Enabled = True
            Me.vsUseDept.Editable = flexEDKbdMouse
            For i = 0 To OptAppUse.Count - 1
                If i = 0 Then
                    OptAppUse(i).Enabled = True
                Else
                    '���ݲ�����ȷ���Ƿ����
                    OptAppUse(i).Enabled = (Val(Mid(mstrӦ�÷�Χ, i, 1)) = 1)
                End If
            Next
            For intCount = Me.optִ�в���.LBound To Me.optִ�в���.UBound
                Me.optִ�в���(intCount).Enabled = True
            Next
            For intCount = Me.OptApp.LBound To Me.OptApp.UBound
                If intCount = 0 Then
                    Me.OptApp(intCount).Enabled = True
                Else
                    Me.OptApp(intCount).Enabled = (Val(Mid(mstrӦ�÷�Χ, intCount, 1)) = 1)
                End If
            Next
        Case 1, 2     '1-����,2-סԺ
            Me.cboִ��Ƶ��.ListIndex = 1: Me.cboִ��Ƶ��.Enabled = False
            Me.cbo���㷽ʽ.ListIndex = 3: Me.cbo���㷽ʽ.Enabled = False: Me.txt���㵥λ.Text = ""
            Me.chk�������(0).Value = 1: Me.chk�������(0).Enabled = False
            Me.chk�������(1).Value = 0: Me.chk�������(1).Enabled = False
            Me.chk�������(2).Value = 0: Me.chk�������(2).Enabled = False
            Me.optִ�в���(1).Value = True
            vsUseDept.Editable = flexEDNone
            For i = 0 To OptAppUse.Count - 1
                OptAppUse(i).Enabled = False
            Next
            For intCount = Me.optִ�в���.LBound To Me.optִ�в���.UBound
                Me.optִ�в���(intCount).Enabled = False
            Next
            For intCount = Me.OptApp.LBound To Me.OptApp.UBound
                Me.OptApp(intCount).Enabled = False
            Next
        Case 3, 5    '3-ת��,5-��Ժ
            Me.cboִ��Ƶ��.ListIndex = 1: Me.cboִ��Ƶ��.Enabled = False
            Me.cbo���㷽ʽ.ListIndex = 3: Me.cbo���㷽ʽ.Enabled = False: Me.txt���㵥λ.Text = ""
            Me.chk�������(0).Value = 0: Me.chk�������(0).Enabled = False
            Me.chk�������(1).Value = 1: Me.chk�������(1).Enabled = False
            Me.chk�������(2).Value = 0: Me.chk�������(2).Enabled = False
            Me.optִ�в���(1).Value = True
            Me.vsUseDept.Editable = flexEDNone
            For i = 0 To OptAppUse.Count - 1
                OptAppUse(i).Enabled = False
            Next
            For intCount = Me.optִ�в���.LBound To Me.optִ�в���.UBound
                Me.optִ�в���(intCount).Enabled = False
            Next
            For intCount = Me.OptApp.LBound To Me.OptApp.UBound
                Me.OptApp(intCount).Enabled = False
            Next
        Case 4, 14     '4-����; 14-��ǰ
            Me.cboִ��Ƶ��.ListIndex = 2: Me.cboִ��Ƶ��.Enabled = False
            Me.cbo���㷽ʽ.ListIndex = 0: Me.cbo���㷽ʽ.Enabled = False: Me.txt���㵥λ.Text = ""
            Me.chk�������(0).Value = 0: Me.chk�������(0).Enabled = False
            Me.chk�������(1).Value = 1: Me.chk�������(1).Enabled = False
            Me.chk�������(2).Value = 0: Me.chk�������(2).Enabled = False
            Me.vsUseDept.Editable = flexEDNone
            For i = 0 To OptAppUse.Count - 1
                OptAppUse(i).Enabled = False
            Next
            Me.optִ�в���(1).Value = True
            For intCount = Me.optִ�в���.LBound To Me.optִ�в���.UBound
                Me.optִ�в���(intCount).Enabled = False
            Next
            For intCount = Me.OptApp.LBound To Me.OptApp.UBound
                Me.OptApp(intCount).Enabled = False
            Next
        Case 6   '6-תԺ
            Me.cboִ��Ƶ��.ListIndex = 1: Me.cboִ��Ƶ��.Enabled = False
            Me.cbo���㷽ʽ.ListIndex = 3: Me.cbo���㷽ʽ.Enabled = False: Me.txt���㵥λ.Text = ""
            Me.chk�������(0).Value = 1: Me.chk�������(0).Enabled = False
            Me.chk�������(1).Value = 1: Me.chk�������(1).Enabled = False
            Me.chk�������(2).Value = 0: Me.chk�������(2).Enabled = False
            Me.vsUseDept.Editable = flexEDNone
            For i = 0 To OptAppUse.Count - 1
                OptAppUse(i).Enabled = False
            Next
            Me.optִ�в���(1).Value = True
            For intCount = Me.optִ�в���.LBound To Me.optִ�в���.UBound
                Me.optִ�в���(intCount).Enabled = False
            Next
            For intCount = Me.OptApp.LBound To Me.OptApp.UBound
                Me.OptApp(intCount).Enabled = False
            Next
        Case 7   '7-����
            Me.cboִ��Ƶ��.ListIndex = 1: Me.cboִ��Ƶ��.Enabled = False
            Me.cbo���㷽ʽ.ListIndex = 3: Me.cbo���㷽ʽ.Enabled = False: Me.txt���㵥λ.Text = ""
            Me.chk�������(0).Value = 0: Me.chk�������(0).Enabled = False
            Me.chk�������(1).Value = 1: Me.chk�������(1).Enabled = False
            Me.chk�������(2).Value = 0: Me.chk�������(2).Enabled = False
            Me.vsUseDept.Editable = flexEDNone
            For i = 0 To OptAppUse.Count - 1
                OptAppUse(i).Enabled = False
            Next
            Me.optִ�в���(1).Value = True
            For intCount = Me.optִ�в���.LBound To Me.optִ�в���.UBound
                Me.optִ�в���(intCount).Enabled = False
            Next
            For intCount = Me.OptApp.LBound To Me.OptApp.UBound
                Me.OptApp(intCount).Enabled = False
            Next
        Case 8, 11
            Me.cboִ��Ƶ��.ListIndex = 1: Me.cboִ��Ƶ��.Enabled = False
            Me.cbo���㷽ʽ.ListIndex = 3: Me.cbo���㷽ʽ.Enabled = False: Me.txt���㵥λ.Text = ""
            Me.chk�������(0).Value = 0: Me.chk�������(0).Enabled = False
            Me.chk�������(1).Value = 1: Me.chk�������(1).Enabled = False
            Me.chk�������(2).Value = 0: Me.chk�������(2).Enabled = False
            Me.vsUseDept.Editable = flexEDNone
            For i = 0 To OptAppUse.Count - 1
                OptAppUse(i).Enabled = False
            Next
            Me.optִ�в���(1).Value = True
            For intCount = Me.optִ�в���.LBound To Me.optִ�в���.UBound
                Me.optִ�в���(intCount).Enabled = False
            Next
            For intCount = Me.OptApp.LBound To Me.OptApp.UBound
                Me.OptApp(intCount).Enabled = False
            Next
        Case 9, 10
            Me.cboִ��Ƶ��.ListIndex = 2: Me.cboִ��Ƶ��.Enabled = False
            Me.cbo���㷽ʽ.ListIndex = 0: Me.cbo���㷽ʽ.Enabled = False: Me.txt���㵥λ.Text = ""
            Me.chk�������(0).Value = 0: Me.chk�������(0).Enabled = False
            Me.chk�������(1).Value = 1: Me.chk�������(1).Enabled = False
            Me.vsUseDept.Editable = flexEDNone
            For i = 0 To OptAppUse.Count - 1
                OptAppUse(i).Enabled = False
            Next
            Me.chk�������(2).Value = 0: Me.chk�������(2).Enabled = False
            Me.optִ�в���(1).Value = True
            For intCount = Me.optִ�в���.LBound To Me.optִ�в���.UBound
                Me.optִ�в���(intCount).Enabled = False
            Next
            For intCount = Me.OptApp.LBound To Me.OptApp.UBound
                Me.OptApp(intCount).Enabled = False
            Next
        Case 12
            Me.cbo��������.Width = Me.txt��������.Left + Me.txt��������.Width - Me.cbo��������.Left
            Me.txt��Ŀ����.Width = Me.txt��������.Left + Me.txt��������.Width - Me.txt��Ŀ����.Left

            Me.cboִ��Ƶ��.ListIndex = 2: Me.cboִ��Ƶ��.Enabled = False
            Me.cbo���㷽ʽ.ListIndex = 0: Me.cbo���㷽ʽ.Enabled = False: Me.txt���㵥λ.Text = ""
            Me.chk�������(0).Value = 0: Me.chk�������(0).Enabled = False
            Me.chk�������(1).Value = 1: Me.chk�������(1).Enabled = False
            Me.chk�������(2).Value = 0: Me.chk�������(2).Enabled = False
            Me.vsUseDept.Editable = flexEDNone
            For i = 0 To OptAppUse.Count - 1
                OptAppUse(i).Enabled = False
            Next
            Me.optִ�в���(2).Value = True
            Me.fra¼����.Visible = True
            For intCount = Me.optִ�в���.LBound To Me.optִ�в���.UBound
                Me.optִ�в���(intCount).Enabled = False
            Next
            Me.OptApp(0).Value = True
            For intCount = Me.OptApp.LBound To Me.OptApp.UBound
                Me.OptApp(intCount).Enabled = False
            Next
        End Select
    End If
End Sub

Private Sub cbo��������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo���㷽ʽ_Click()
    If cboִ��Ƶ��.ListIndex = 0 And cbo���㷽ʽ.ListIndex > 0 Then
        lbl�������.Visible = True
        cbo�������.Visible = True
    Else
        lbl�������.Visible = False
        cbo�������.Visible = False
    End If
End Sub

Private Sub cbo���㷽ʽ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo���_Click()
    Dim i As Long
    
    Me.fra��鲿λ.Visible = False: Me.fra��׼����.Visible = False
    Me.fra�걾��λ.Visible = False: lblӢ��.Visible = False: Me.txtӢ��.Visible = False
    Me.chk����Ӧ��.Value = 1: Me.chk����Ӧ��.Enabled = True
    Me.chk�������(0).Enabled = True: Me.chk�������(1).Enabled = True: Me.chk�������(2).Enabled = True
    Me.lbl����˵��.Visible = False
    Me.cbo����˵��.Visible = False
    Me.cboִ�з���.Visible = False
    Me.lblִ�з���.Visible = False
    Me.lbl��Һ����.Visible = False
    Me.cbo��Һ����.Visible = False
    Me.cboBloodType.Visible = False
    vsfBloodLis.Visible = False
    
    Me.vsUseDept.Editable = flexEDKbdMouse
    For i = 0 To OptAppUse.Count - 1
        If i = 0 Then
            OptAppUse(i).Enabled = True
        Else
            '���ݲ�����ȷ���Ƿ����
            OptAppUse(i).Enabled = (Val(Mid(mstrӦ�÷�Χ, i, 1)) = 1)
        End If
    Next
    
    Me.imgComment.Top = Me.txt��������.Top + 250
    Me.lblComment.Top = Me.imgComment.Top + 50
    
    On Error GoTo ErrHand
    For intCount = Me.optִ�в���.LBound To Me.optִ�в���.UBound
        Me.optִ�в���(intCount).Enabled = True
    Next
    For intCount = Me.OptApp.LBound To Me.OptApp.UBound
        If intCount = 0 Then
            Me.OptApp(intCount).Enabled = True
        Else
            Me.OptApp(intCount).Enabled = (Val(Mid(mstrӦ�÷�Χ, intCount, 1)) = 1)
        End If
    Next
    Me.cbo��������.Width = Me.fra��׼����.Left + Me.fra��׼����.Width - Me.cbo��������.Left
    Me.txt��Ŀ����.Width = Me.fra��׼����.Left + Me.fra��׼����.Width - Me.txt��Ŀ����.Left
    Me.chk�������.Visible = False
    
    Me.stbInfo.TabVisible(2) = False '��鲿λ
    chk����.Visible = False
    
    If cbo���.Text = "D-���" And cbo��������.Text = "18-����" Then
        lbl�������.Visible = True
        cbo�������.Visible = True
        stbInfo.TabCaption(2) = "����걾"
    Else
        lbl�������.Visible = False
        cbo�������.Visible = False
    End If
    
    Me.cboִ��Ƶ��.Clear
    Select Case Left(Me.cbo���.Text, 1)
    Case "C", "D"
        aryTemp = Split("0-��ѡƵ��;1-һ����", ";")
    Case "H"
        aryTemp = Split("0-��ѡƵ��;2-������", ";")
    Case Else
        aryTemp = Split("0-��ѡƵ��;1-һ����;2-������", ";")
    End Select
    For intCount = LBound(aryTemp) To UBound(aryTemp)
        Me.cboִ��Ƶ��.AddItem aryTemp(intCount)
    Next

    Select Case Left(Me.cbo���.Text, 1)
    Case "C"        '����
        Me.cbo��������.Width = Me.txt��������.Left + Me.txt��������.Width - Me.cbo��������.Left
        Me.txt��Ŀ����.Width = Me.txt��������.Left + Me.txt��������.Width - Me.txt��Ŀ����.Left
        Me.fra�걾��λ.Visible = True

        Me.chk�������.Visible = True
        Me.txtӢ��.Visible = True: lblӢ��.Visible = True
        
        Me.lbl��������.Caption = "��������(&T)": Me.lbl��������.Visible = True
        Me.cbo��������.Clear: Me.cbo��������.Visible = True
        err = 0: On Error GoTo ErrHand
        
        gstrSql = "select ����||'-'||���� from ���Ƽ������� order by ����"
'            If .State = adStateOpen Then .Close
'            Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "cbo���_Click")
'            Call SQLTest
        With rsTemp
            Do While Not .EOF
                Me.cbo��������.AddItem .Fields(0).Value
                .MoveNext
            Loop
            If Me.cbo��������.ListCount > 0 Then Me.cbo��������.ListIndex = 0
        End With
        Me.cboִ��Ƶ��.ListIndex = 1: Me.cboִ��Ƶ��.Enabled = True
        Me.cbo���㷽ʽ.ListIndex = 3: Me.cbo���㷽ʽ.Enabled = False: Me.txt���㵥λ.Text = "��"
        Me.imgComment.Top = Me.fra�걾��λ.Top + Me.fra�걾��λ.Height + 50
        Me.lblComment.Top = Me.imgComment.Top + 50
        Me.lblComment.Caption = Space(4) & "��������Ŀֻ����һ������Ŀ��ֻ�����������סԺ��ʱҽ����ʹ�ã�Ϊ��Чִ�У���Ҫ�ڼ�����Ŀ������ָ����걾�Ͳο�ȡֵ�����������Ŀ��Ӧ�������Ӧ�Ļ���ָ���"
        Me.chk�������(0).Value = 1: Me.chk�������(1).Value = 1
        Me.optִ�в���(0).Value = False: Me.optִ�в���(0).Enabled = False
        Me.optִ�в���(1).Value = True
        Me.lbllel.Visible = False
        Me.txtML.Visible = False: Me.txtML.Text = ""
        Me.lblML.Visible = False
    Case "D"        '���
        Me.cbo��������.Width = Me.txt��������.Left + Me.txt��������.Width - Me.cbo��������.Left
        Me.txt��Ŀ����.Width = Me.txt��������.Left + Me.txt��������.Width - Me.txt��Ŀ����.Left
        Me.lbl��������.Caption = "�������(&T)": Me.lbl��������.Visible = True
        Me.cbo��������.Clear: Me.cbo��������.Visible = True
        err = 0: On Error GoTo ErrHand
        gstrSql = "select ����||'-'||���� from ���Ƽ������ order by ����"
'            If .State = adStateOpen Then .Close
'            Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "cbo���_Click")
'            Call SQLTest
        With rsTemp
            Do While Not .EOF
                Me.cbo��������.AddItem .Fields(0).Value
                .MoveNext
            Loop
            If Me.cbo��������.ListCount > 0 Then Me.cbo��������.ListIndex = 0
        End With
        Me.cboִ��Ƶ��.ListIndex = 1: Me.cboִ��Ƶ��.Enabled = True
        Me.cbo���㷽ʽ.ListIndex = 3: Me.cbo���㷽ʽ.Enabled = False: Me.txt���㵥λ.Text = "��"
        Me.lblComment.Caption = Space(4) & "�������Ŀֻ����һ������Ŀ���̶���λ������ָ����λ����ѡ��λ��Ŀ��Ҫͨ����λ���ɳ���ָ�����ѡ����λ��Ŀ����ʹ�á�"
        'Me.fra��鲿λ.Visible = True
        Me.stbInfo.TabVisible(2) = True '��鲿λ
        stbInfo.TabCaption(2) = "��鲿λ"
        Me.optִ�в���(0).Value = False: Me.optִ�в���(0).Enabled = False
        Me.optִ�в���(1).Value = True
        chk����.Visible = True
        Me.lbllel.Visible = False
        Me.txtML.Visible = False: Me.txtML.Text = ""
        Me.lblML.Visible = False
    Case "E"        '����
        Me.lbl��������.Caption = "��������(&T)": Me.lbl��������.Visible = True
        Me.cbo��������.Clear: Me.cbo��������.Visible = True
        aryTemp = Split("0-��ͨ;1-��������;2-��ҩ����(��ҩ);3-��ҩ�巨;4-��ҩ��(��)��;5-��������;6-�걾�ɼ�;7-��Ѫ����;8-��Ѫ;��;9-��Ѫ�ɼ�", ";")
        For intCount = LBound(aryTemp) To UBound(aryTemp)
            Me.cbo��������.AddItem aryTemp(intCount)
        Next
        Me.cbo��������.ListIndex = 0
        Me.cboִ��Ƶ��.ListIndex = 1: Me.cboִ��Ƶ��.Enabled = True
        Me.cbo���㷽ʽ.ListIndex = 3: Me.cbo���㷽ʽ.Enabled = True: Me.txt���㵥λ.Text = "��"
        Me.lblComment.Caption = Space(4) & "��������Ŀ������ͨ���ƺ͹������顢��ҩ;������ҩ�巨�ȣ���������Ŀ��׼ȷ���������ʣ��Ա�ҽ������ִ��ʱ���á�"
        If Me.cbo��������.ListIndex = 1 Then
            stbInfo.TabVisible(3) = True
            Call InivsfTest
        Else
            stbInfo.TabVisible(3) = False
        End If

        If Me.cbo��������.ListIndex = 2 Then
            Me.lbl����˵��.Visible = True
            Me.cbo����˵��.Visible = True
            
            Me.cboִ�з���.Visible = True
            Me.lblִ�з���.Visible = True
        End If
        
        Me.lbllel.Visible = False
        Me.txtML.Visible = False: Me.txtML.Text = ""
        Me.lblML.Visible = False
    Case "F"        '����
        Me.cbo��������.Width = Me.txt��������.Left + Me.txt��������.Width - Me.cbo��������.Left
        Me.txt��Ŀ����.Width = Me.txt��������.Left + Me.txt��������.Width - Me.txt��Ŀ����.Left

        Me.lbl��������.Caption = "������ģ(&T)": Me.lbl��������.Visible = True
        Me.cbo��������.Clear: Me.cbo��������.Visible = True
        err = 0: On Error GoTo ErrHand
        gstrSql = "select ����||'-'||���� from ����������ģ order by ����"
'            If .State = adStateOpen Then .Close
'            Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "cbo���_Click")
'            Call SQLTest
        With rsTemp
            Do While Not .EOF
                Me.cbo��������.AddItem .Fields(0).Value
                .MoveNext
            Loop
            If Me.cbo��������.ListCount > 0 Then Me.cbo��������.ListIndex = 0
        End With
        Me.cbo��������.ListIndex = 0
        Me.cboִ��Ƶ��.ListIndex = 1: Me.cboִ��Ƶ��.Enabled = False
        Me.cbo���㷽ʽ.ListIndex = 3: Me.cbo���㷽ʽ.Enabled = False: Me.txt���㵥λ.Text = "��"
        Me.lblComment.Caption = Space(4) & "��������Ŀֻ����һ������Ŀ������Ϊ��ͬ��ģ��������Ŀǰϵͳֻ�����������סԺ��ʱҽ����ʹ�á�"
        Me.fra��׼����.Visible = True
        Me.optִ�в���(0).Value = False: Me.optִ�в���(0).Enabled = False
        Me.optִ�в���(1).Value = True
        
        Me.lbllel.Visible = False
        Me.txtML.Visible = False: Me.txtML.Text = ""
        Me.lblML.Visible = False
    Case "G"        '����
        Me.lbl��������.Caption = "��������(&T)": Me.lbl��������.Visible = True
        Me.cbo��������.Clear: Me.cbo��������.Visible = True
        err = 0: On Error GoTo ErrHand
        gstrSql = "select ����||'-'||���� from ������������ order by ����"
'            If .State = adStateOpen Then .Close
'            Call SQLTest(App.ProductName, Me.Caption, gstrSql)
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "cbo���_Click")
'            Call SQLTest
        With rsTemp
            Do While Not .EOF
                Me.cbo��������.AddItem .Fields(0).Value
                .MoveNext
            Loop
            If Me.cbo��������.ListCount > 0 Then Me.cbo��������.ListIndex = 0
        End With
        Me.cbo��������.ListIndex = 0
        Me.cboִ��Ƶ��.ListIndex = 1: Me.cboִ��Ƶ��.Enabled = False
        Me.cbo���㷽ʽ.ListIndex = 3: Me.cbo���㷽ʽ.Enabled = False: Me.txt���㵥λ.Text = "��"
        Me.chk����Ӧ��.Value = 0: Me.chk����Ӧ��.Enabled = False
        Me.lblComment.Caption = Space(4) & "��������Ŀֻ����һ������Ŀ��ֻ��������ҽ���и�����Ҫָ������������������Ŀ����ʹ�á�"
        Me.optִ�в���(0).Value = False: Me.optִ�в���(0).Enabled = False
        Me.optִ�в���(1).Value = True
        
        Me.lbllel.Visible = False
        Me.txtML.Visible = False: Me.txtML.Text = ""
        Me.lblML.Visible = False
    Case "H"        '����
        Me.lbl��������.Caption = "��Ŀ����(&T)": Me.lbl��������.Visible = True
        Me.cbo��������.Clear: Me.cbo��������.Visible = True
        aryTemp = Split("0-������;1-����ȼ�", ";")
        For intCount = LBound(aryTemp) To UBound(aryTemp)
            Me.cbo��������.AddItem aryTemp(intCount)
        Next
        Me.cbo��������.ListIndex = 0
        Me.cboִ��Ƶ��.ListIndex = 1
        Me.cbo���㷽ʽ.ListIndex = 0: Me.cbo���㷽ʽ.Enabled = False: Me.txt���㵥λ.Text = " "
        Me.chk�������(0).Value = 0: Me.chk�������(0).Enabled = False
        Me.chk�������(1).Value = 1: Me.chk�������(1).Enabled = False
        Me.chk�������(2).Value = 0: Me.chk�������(2).Enabled = False
        Me.vsUseDept.Editable = flexEDNone
        For i = 0 To OptAppUse.Count - 1
            OptAppUse(i).Enabled = False
        Next
        Me.lblComment.Caption = Space(4) & "��������Ŀ����������ͻ���ȼ���Ϊ�����Ե���Ŀ��ֻ��סԺ����ҽ����ʹ�á�"
        
        Me.lbllel.Visible = False
        Me.txtML.Visible = False: Me.txtML.Text = ""
        Me.lblML.Visible = False
    Case "I"        '��ʳ
        Me.lbl��������.Visible = False
        Me.cbo��������.Clear: Me.cbo��������.Visible = False
        Me.cboִ��Ƶ��.ListIndex = 2: Me.cboִ��Ƶ��.Enabled = False
        Me.cbo���㷽ʽ.ListIndex = 0: Me.cbo���㷽ʽ.Enabled = True: Me.txt���㵥λ.Text = " "
        Me.chk�������(0).Value = 0: Me.chk�������(1).Value = 1
        Me.lblComment.Caption = Space(4) & "��ʳ����Ŀ��ҽ�������������ҽ�Ƶ���ʳҪ��Ϊ�����Ե���Ŀ��ͨ��ֻ��סԺ����ҽ����ʹ�á�"
        
        Me.lbllel.Visible = False
        Me.txtML.Visible = False: Me.txtML.Text = ""
        Me.lblML.Visible = False
    Case "K"        '��Ѫ
        Me.optִ�в���(0).Enabled = False
        Me.lbl��������.Visible = False
        Me.cbo��������.Clear: Me.cbo��������.Visible = False
        Me.cboִ��Ƶ��.ListIndex = 1: Me.cboִ��Ƶ��.Enabled = False
        Me.cbo���㷽ʽ.ListIndex = 1: Me.cbo���㷽ʽ.Enabled = True: Me.txt���㵥λ.Text = " "
        Me.chk�������(0).Value = 0: Me.chk�������(1).Value = 1
        Me.lblComment.Caption = Space(4) & "��Ѫͨ����Ϊ���˸������ƴ�ʩһ����Ӧ�ã�����ʵ����Ѫ��ִ�С�"
        Me.lbllel.Visible = True
        Me.txtML.Visible = True: Me.txtML.Text = ""
        Me.lblML.Visible = True
        vsfBloodLis.Visible = True
    Case "L"        '����
        Me.lbl��������.Visible = False
        Me.cbo��������.Clear: Me.cbo��������.Visible = False
        Me.cboִ��Ƶ��.ListIndex = 2: Me.cboִ��Ƶ��.Enabled = True
        Me.cbo���㷽ʽ.ListIndex = 0: Me.cbo���㷽ʽ.Enabled = True: Me.txt���㵥λ.Text = " "
        Me.chk�������(0).Value = 0: Me.chk�������(1).Value = 1
        Me.lblComment.Caption = Space(4) & "����ͨ����Ϊ���˳��õĸ������ƴ�ʩ����Ӧ�ã�����ʵ��������ʱ�����ִ�С�"
        
        Me.lbllel.Visible = False
        Me.txtML.Visible = False: Me.txtML.Text = ""
        Me.lblML.Visible = False
    Case "M"        '����
        Me.lbl��������.Visible = False
        Me.cbo��������.Clear: Me.cbo��������.Visible = False
        Me.cboִ��Ƶ��.ListIndex = 0: Me.cboִ��Ƶ��.Enabled = True
        Me.cbo���㷽ʽ.ListIndex = 1: Me.cbo���㷽ʽ.Enabled = True: Me.txt���㵥λ.Text = " "
        Me.chk�������(0).Value = 0: Me.chk�������(1).Value = 1
        Me.lblComment.Caption = Space(4) & "�ڲ������ƹ����У����Ը���ʵ�����Ӧ��ĳЩ���ϣ�ִ��Ƶ�ʺͼ��㷽ʽ������������ʵ����Ŀ���á�"
        stbInfo.TabCaption(2) = "Ƶ������"
        
        Me.lbllel.Visible = False
        Me.txtML.Visible = False: Me.txtML.Text = ""
        Me.lblML.Visible = False
    Case "Z"        '����
        Me.lbl��������.Caption = "�����־(&T)": Me.lbl��������.Visible = True
        Me.cbo��������.Clear: Me.cbo��������.Visible = True
        aryTemp = Split("0-��ͨ;1-����;2-סԺ;3-ת��;4-����;5-��Ժ;6-תԺ;7-����;8-����;9-����;10-��Σ;11-����;12-��¼�����;14-��ǰ", ";")
        For intCount = LBound(aryTemp) To UBound(aryTemp)
            Me.cbo��������.AddItem aryTemp(intCount)
        Next
        Me.cbo��������.ListIndex = 0
        Me.cboִ��Ƶ��.ListIndex = 0: Me.cboִ��Ƶ��.Enabled = True
        Me.cbo���㷽ʽ.ListIndex = 0: Me.cbo���㷽ʽ.Enabled = True: Me.txt���㵥λ.Text = " "
        Me.chk�������(0).Value = 0: Me.chk�������(1).Value = 1
        Me.lblComment.Caption = Space(4) & "����ȷ���Ե�������Ŀ�����������������ݾ�����Ŀ�ص���ȷ���������ԣ�ֱ��Ӱ��ҽ�����´����Чִ�С�"
        
        Me.lbllel.Visible = False
        Me.txtML.Visible = False: Me.txtML.Text = ""
        Me.lblML.Visible = False
    End Select
    
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cbo���_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cbo�����Ա�_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub cboִ��Ƶ��_Click()
    If Left(Me.cbo���.Text, 1) = "H" Then
        If cboִ��Ƶ��.ListIndex = 0 Then
            cbo���㷽ʽ.ListIndex = 2
        ElseIf cboִ��Ƶ��.ListIndex = 1 Then
            cbo���㷽ʽ.ListIndex = 0
        End If
    End If
    
    '��ѡƵ�ʼ��Ǽ�����Ŀʱ����������Ƶ��
    If cboִ��Ƶ��.ListIndex = 0 And Left(Me.cbo���.Text, 1) <> "C" Then
        stbInfo.TabVisible(4) = True
        Call InitVsfFreq
    Else
        stbInfo.TabVisible(4) = False
    End If
    
    If cboִ��Ƶ��.ListIndex = 0 And cbo���㷽ʽ.ListIndex > 0 Then
        lbl�������.Visible = True
        cbo�������.Visible = True
    Else
        lbl�������.Visible = False
        cbo�������.Visible = False
    End If
    
    If InStr(1, cboִ��Ƶ��.Text, "������") > 0 Then
        Me.lblZLPL.Visible = True
        Me.cboZLPL.Visible = True
    Else
        Me.lblZLPL.Visible = False
        Me.cboZLPL.Visible = False
    End If
End Sub

Private Sub cboִ��Ƶ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub
Private Sub chk����Ӧ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk�������_Click(Index As Integer)
    Dim i As Long, j As Long
    Dim blnIsUse As Boolean
    
    If Me.chk�������(0).Enabled Then
        If Me.chk�������(0).Value = 1 Then
            Me.txt����ִ��.Enabled = True
            txt����ִ��.BackColor = vbWindowBackground
        Else
            Me.txt����ִ��.Enabled = False
            txt����ִ��.BackColor = vbButtonFace
        End If
        If Me.chk�������(1).Value = 1 Then
            Me.txtסԺִ��.Enabled = True
            txtסԺִ��.BackColor = vbWindowBackground
            Me.chk����.Enabled = True
        Else
            Me.txtסԺִ��.Enabled = False
            Me.chk����.Enabled = False
            txtסԺִ��.BackColor = vbButtonFace
            Me.chk����.Value = 0
        End If
    Else
        
    End If
    If Me.chk�������(0).Value = 0 And Me.chk�������(1).Value = 0 And Me.chk�������(2).Value = 0 Then
        If chk�������(0).Enabled = True Then
            For i = 0 To vsUseDept.Rows - 1
                For j = 0 To vsUseDept.Cols - 1
                    If vsUseDept.ColHidden(j) = False Then
                        If vsUseDept.TextMatrix(i, j) <> "" Then
                            blnIsUse = True
                            Exit For
                        End If
                    End If
                Next
            Next
            If blnIsUse Then
                chk�������(Index).Value = 1
            Else
                vsUseDept.Editable = flexEDNone
                For i = 0 To OptAppUse.Count - 1
                    OptAppUse(i).Enabled = False
                Next
            End If
        Else
            vsUseDept.Editable = flexEDNone
            For i = 0 To OptAppUse.Count - 1
                OptAppUse(i).Enabled = False
            Next
        End If
    Else
        If Me.chk�������(0).Enabled Then
            vsUseDept.Editable = flexEDKbdMouse
            For i = 0 To OptAppUse.Count - 1
                If i = 0 Then
                    OptAppUse(i).Enabled = True
                Else
                    '���ݲ�����ȷ���Ƿ����
                    OptAppUse(i).Enabled = (Val(Mid(mstrӦ�÷�Χ, i, 1)) = 1)
                End If
            Next
        End If
    End If
    '�����Ŀ�����סԺ����
    If Index = 0 Or Index = 1 Then
        If chk�������(Index).Value = 1 And chk�������(2).Value = 1 Then
            chk�������(2).Value = 0
        End If
    ElseIf chk�������(Index).Value = 1 Then
        If chk�������(0).Value = 1 Then
            chk�������(0).Value = 0
        End If
        If chk�������(1).Value = 1 Then
            chk�������(1).Value = 0
        End If
    End If
End Sub

Private Sub chk�������_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub chk�������_Click()
    If Me.chk�������.Value = 1 Then
        Me.chk����Ӧ��.Enabled = False: Me.chk����Ӧ��.Value = 1
    Else
        Me.chk����Ӧ��.Enabled = True
    End If
'    Me.stbInfo.TabVisible(2) = Not Me.chk�������.Value = 1
End Sub

Private Sub chkִ�а���_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Left(Me.cbo���.Text, 1) = "D" Then
        Call zlCommFun.PressKey(vbKeyTab)
    ElseIf Me.fra��׼����.Visible Then
        Call zlCommFun.PressKey(vbKeyTab)
    Else
        Me.stbInfo.Tab = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
    Exit Sub
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Function CheckUseDept(ByVal strDept As String) As String
'���ʹ�ÿ��ҵ�վ��ͷ�������Ƿ�ͽ����ϵ��Ǻ�
'���أ�����в��Ǻϵģ�������ʾ��Ϣ
    Dim rsTmp As Recordset
    Dim strSql As String
    Dim strMsg As String
    Dim strվ�� As String
    
    On Error GoTo errH
    If strDept = "" Then Exit Function
    If cmbStationNo.Text <> "" Then
        strվ�� = Mid(cmbStationNo.Text, 1, InStr(1, cmbStationNo.Text, "-") - 1)
        
        strSql = "Select a.ID,a.����,a.վ�� From ���ű� A Where ID In(" & strDept & ") And (a.վ��<>[1] And a.վ�� Is Not Null)"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strվ��)
        If rsTmp.RecordCount > 0 Then
            Do While Not rsTmp.EOF
                strMsg = strMsg & "," & rsTmp!����
                rsTmp.MoveNext
            Loop
            strMsg = Mid(strMsg, 2)
            CheckUseDept = strMsg & " ���� " & strվ�� & " վ��Ŀ��ң����顣"
        End If
    End If
    strSql = ""
    If chk�������(0).Value = 0 And chk�������(2).Value = 0 Then
        'û�й�ѡ�������Ƿ���ֻ����������Ŀ��ҡ�
        strSql = "Select ID,���� From (Select a.ID,a.����,Decode(Max(�������), 3, 3, 2, Decode(Min(�������), 1, 3, 2, 2), 1, 1) As ������� " & vbNewLine & _
                " From ���ű� A, ��������˵�� B" & vbNewLine & _
                " Where a.Id = b.����id And b.������� <> 0" & vbNewLine & _
                " And ID In(" & strDept & ") " & _
                " Group By a.Id,a.����, վ�� ) Where �������=1"

    ElseIf chk�������(1).Value = 0 Then
        'û�й�ѡסԺ������Ƿ���ֻ������סԺ�Ŀ��ҡ�
        strSql = "Select ID,���� From (Select a.ID,a.����,Decode(Max(�������), 3, 3, 2, Decode(Min(�������), 1, 3, 2, 2), 1, 1) As ������� " & vbNewLine & _
                " From ���ű� A, ��������˵�� B" & vbNewLine & _
                " Where a.Id = b.����id And b.������� <> 0" & vbNewLine & _
                " And ID In(" & strDept & ")" & _
                " Group By a.Id,a.����, վ��) Where �������=2"
    End If
    If strSql <> "" Then
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
        If rsTmp.RecordCount > 0 Then
            Do While Not rsTmp.EOF
                strMsg = strMsg & "," & rsTmp!����
                rsTmp.MoveNext
            Loop
            strMsg = Mid(strMsg, 2)
            If chk�������(0).Value = 0 And chk�������(2).Value = 0 Then
                CheckUseDept = strMsg & " ��ֻ����������Ŀ��ң����顣"
            Else
                CheckUseDept = strMsg & " ��ֻ������סԺ�Ŀ��ң����顣"
            End If
            
        End If
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cmdOK_Click()
    Dim strFormula As String
    Dim strErrorMsg As String, iErrorPos As Integer
    Dim strMid As Variant
    Dim i As Integer, lngVItemID0 As Long
    Dim mAppType As Integer                 'Ӧ������ =0Ӧ���ڱ���;=1Ӧ����ͬ��;=2Ӧ���ڱ���������;=3Ӧ�����������
    Dim strTmp As String, blnBegin As Boolean
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset
    Dim strFreq As String
    Dim strվ�� As String
    Dim j As Long
    Dim strDeptId As String
    Dim strMsg As String
    Dim str���� As String
    Dim intѭ�� As Integer
    Dim lngNum As Long
    Dim lngSum As Long
    Dim intFirst As Integer
    Dim strLast As String
    Dim intRow As Integer
    Dim str��Ѫ������� As String
    Dim str�Թܱ��� As String
    Dim blnRisTrans As Boolean
    Dim varTemp As Variant
    Dim strItem As String
    
    'Ƥ�Խ��
    Dim strTest As String
    Dim str���� As String
    Dim str���� As String
    
    lngNum = 1
    intѭ�� = 1
    '���¼�����ƣ���ȥ�������ַ�
    strTmp = MoveSpecialChar(txt��Ŀ����.Text)
    If txt��Ŀ����.Text <> strTmp Then
        txt��Ŀ����.Text = strTmp
        Me.txt����ƴ��.Text = zlStr.GetCodeByORCL(strTmp, False, mlng���볤��)
        Me.txt�������.Text = zlStr.GetCodeByORCL(strTmp, True, mlng���볤��)
    End If
    
    '���ʹ�ÿ���
    With Me.vsUseDept
        For i = 0 To .Rows - 1
            For j = 0 To 4
                If .TextMatrix(i, j) <> "" Then
                    strDeptId = strDeptId & "," & .TextMatrix(i, j + 5)
                End If
            Next
        Next
        strDeptId = Mid(strDeptId, 2)
        strMsg = CheckUseDept(strDeptId)
        If strMsg <> "" Then
            MsgBox strMsg, vbInformation, gstrSysName
            Me.stbInfo.Tab = 0: Me.vsUseDept.SetFocus: Exit Sub
        End If
    End With
    strTmp = MoveSpecialChar(txt��������.Text)
    If txt��������.Text <> strTmp Then
        txt��������.Text = strTmp
        Me.txt����ƴ��.Text = zlStr.GetCodeByORCL(strTmp, False, mlng���볤��)
        Me.txt�������.Text = zlStr.GetCodeByORCL(strTmp, True, mlng���볤��)
    End If

    'һ�����Լ��
    If Trim(Me.txt��Ŀ����.Text) = "" Then
        MsgBox "��������Ŀ���룡", vbInformation, gstrSysName
        Me.stbInfo.Tab = 0: Me.txt��Ŀ����.SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt��Ŀ����.Text), vbFromUnicode)) > Me.txt��Ŀ����.MaxLength Then
        MsgBox "��Ŀ����ĳ��ȳ��������" & Me.txt��Ŀ����.MaxLength & " ���ַ�����", vbInformation, gstrSysName
        Me.stbInfo.Tab = 0: Me.txt��Ŀ����.SetFocus: Exit Sub
    End If
    If Trim(Me.txt��Ŀ����.Text) = "" Then
        MsgBox "��������Ŀ���ƣ�", vbInformation, gstrSysName
        Me.stbInfo.Tab = 0: Me.txt��Ŀ����.SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt��Ŀ����.Text), vbFromUnicode)) > Me.txt��Ŀ����.MaxLength Then
        MsgBox "��Ŀ���Ƴ��������" & Me.txt��Ŀ����.MaxLength & "���ַ���" & Me.txt��Ŀ����.MaxLength / 2 & "�����֣���", vbInformation, gstrSysName
        Me.stbInfo.Tab = 0: Me.txt��Ŀ����.SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt����ƴ��.Text), vbFromUnicode)) > Me.txt����ƴ��.MaxLength Then
        MsgBox "��Ŀ���Ƴ��������" & Me.txt����ƴ��.MaxLength & "���ַ�����", vbInformation, gstrSysName
        Me.stbInfo.Tab = 0: Me.txt����ƴ��.SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt�������.Text), vbFromUnicode)) > Me.txt�������.MaxLength Then
        MsgBox "��Ŀ���Ƴ��������" & Me.txt�������.MaxLength & "���ַ�����", vbInformation, gstrSysName
        Me.stbInfo.Tab = 0: Me.txt�������.SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt����ƴ��.Text), vbFromUnicode)) > Me.txt����ƴ��.MaxLength Then
        MsgBox "��Ŀ���Ƴ��������" & Me.txt����ƴ��.MaxLength & "���ַ�����", vbInformation, gstrSysName
        Me.stbInfo.Tab = 0: Me.txt����ƴ��.SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt�������.Text), vbFromUnicode)) > Me.txt�������.MaxLength Then
        MsgBox "��Ŀ���Ƴ��������" & Me.txt�������.MaxLength & "���ַ�����", vbInformation, gstrSysName
        Me.stbInfo.Tab = 0: Me.txt�������.SetFocus: Exit Sub
    End If
    If Me.cbo���㷽ʽ.ListIndex = 1 And Trim(Me.txt���㵥λ.Text) = "" Then
        MsgBox "��������Ŀ����������㵥λ��", vbInformation, gstrSysName
        Me.stbInfo.Tab = 0: Me.txt���㵥λ.SetFocus: Exit Sub
    End If
    If LenB(StrConv(Trim(Me.txt���㵥λ.Text), vbFromUnicode)) > Me.txt���㵥λ.MaxLength Then
        MsgBox "���㵥λ���������" & Me.txt���㵥λ.MaxLength & "���ַ���" & Me.txt���㵥λ.MaxLength / 2 & "�����֣���", vbInformation, gstrSysName
        Me.stbInfo.Tab = 0: Me.txt���㵥λ.SetFocus: Exit Sub
    End If
    If Left(Me.cbo���.Text, 1) = "D" And Me.opt��鲿λ(0).Value = True Then
'        If LenB(StrConv(Trim(Me.txt��鲿λ.Text), vbFromUnicode)) > 60 Then
'            MsgBox "��鲿λ���������40���ַ���20�����֣���", vbInformation, gstrSysName
'            Me.stbInfo.Tab = 0: Me.txt��鲿λ.SetFocus: Exit Sub
'        End If
'        If mbln��ϲ�λ��Ŀ And Trim(Me.txt��鲿λ.Text) = "" Then
'            MsgBox "��鲿λ����Ϊ�գ�", vbInformation, gstrSysName
'            Me.stbInfo.Tab = 0: Me.txt��鲿λ.SetFocus: Exit Sub
'        End If
    End If
    If LenB(StrConv(Trim(Me.txt��������.Text), vbFromUnicode)) > Me.txt��������.MaxLength Then
        MsgBox "�������������" & Me.txt��������.MaxLength & "���ַ���" & Int(Me.txt��������.MaxLength / 2) & "�����֣���", vbInformation, gstrSysName
        Me.stbInfo.Tab = 0: Me.txt��������.SetFocus: Exit Sub
    End If
    If Len(Trim(Me.txt¼������.Text)) > 0 Then
        If CDbl(Me.txt¼������.Text) > CDbl("99999999999.99999") Then
            MsgBox "¼���������ܴ��ڣ�99999999999.99999��ֵ��", vbInformation, gstrSysName
            Me.stbInfo.Tab = 0: Me.txt¼������.SetFocus: Exit Sub
        End If
    End If

    If Left(Me.cbo���.Text, 1) = "C" And Trim(Me.txt�걾��λ.Text) = "" Then
        MsgBox "������Ŀ��������Ĭ�ϱ걾��λ��", vbInformation, gstrSysName
        Me.stbInfo.Tab = 0: Me.txt�걾��λ.SetFocus: Exit Sub
    End If
    '10804 ��������ʱ�������������Ƿ�ɾ��
    If Left(Me.cbo���.Text, 1) = "C" Then
        If Not zlExistItem("���Ƽ�������", "����", Mid(Me.cbo��������.Text, InStr(1, Me.cbo��������, "-") + 1), "�������ͣ�" & Mid(Me.cbo��������.Text, InStr(1, Me.cbo��������, "-") + 1)) Then
            Me.cbo��������.SetFocus:  Exit Sub
        End If
    End If

    '��鲿λ
    If Left(Me.cbo���.Text, 1) = "D" Then
        '�����Ŀ �����������ȷ����֤
        
    End If

    If Me.optִ�в���(4).Value = True Then
        '����ִ�м��
        strTemp = ""
        With Me.msf����ִ��
            For intCount = 1 To .Rows - 1
                If Val(.TextMatrix(intCount, 0)) <> 0 Then
                    '���ټ���Ƿ��ظ� By ��ͮ��
                    'If InStr(1, strTemp & ";", ";" & Trim(.TextMatrix(intCount, 0)) & "-" & .TextMatrix(intCount, 1) & ";") > 0 Then
                    If InStr(1, strTemp & ";", ";" & .TextMatrix(intCount, colִ�п���) & ";") > 0 Then
                        MsgBox "�ظ�ָ����ִ�п��ҡ�" & .TextMatrix(intCount, colִ�п���) & "����", vbInformation, gstrSysName
                        Me.stbInfo.Tab = 1: .SetFocus: Exit Sub
                    Else
                        strTemp = strTemp & ";" & .TextMatrix(intCount, colִ�п���)
                    End If
'                    If Val(.TextMatrix(intCount, 2)) = 0 Then
'                        MsgBox "��" & .TextMatrix(intCount, 1) & "��δָ��ִ�п��ң�", vbInformation, gstrSysName
'                        Me.stbInfo.Tab = 1: .SetFocus: Exit Sub
'                    End If
                End If
            Next
        End With
    End If
    
    'ִ�п���Ӧ��Ϊ��Χʱ��ʾ
    If OptApp(0).Enabled = True And OptApp(0).Value = False Then
        For i = 0 To Me.OptApp.Count - 1
            If OptApp(i).Enabled = True And OptApp(i).Value = True Then
                If MsgBox("����Ŀ���õ�ִ�п��ҽ�" & OptApp(i).Caption & "��Ŀ���Ƿ񱣴棿", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                    Me.stbInfo.Tab = 1:  Exit Sub
                End If
            End If
        Next
    End If
    
    'ʹ�ÿ���Ӧ��Ϊ��Χʱ��ʾ
    If OptAppUse(0).Enabled = True And OptAppUse(0).Value = False Then
        For i = 0 To Me.OptAppUse.Count - 1
            If OptAppUse(i).Enabled = True And OptAppUse(i).Value = True Then
                If MsgBox("����Ŀ���õ�ʹ�ÿ��ҽ�" & OptAppUse(i).Caption & "��Ŀ���Ƿ񱣴棿", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
                    Me.stbInfo.Tab = 0:  Exit Sub
                End If
            End If
        Next
    End If
    
    If cmbStationNo.Text = "" Then
        strվ�� = "Null"
    Else
        strվ�� = Mid(cmbStationNo.Text, 1, InStr(1, cmbStationNo.Text, "-") - 1)
    End If
    
    '������Ŀʱ����֤�������ظ����룬������ظ��Զ���ԭ��������ϼ�1��ֱ�����ظ�
    str���� = Trim(txt��Ŀ����.Text)
    If Me.Tag = "����" Or Me.Tag = "��������" Then
        Do While True
            gstrSql = "select a.���� from ������ĿĿ¼ a,������Ŀ��� b where a.����=[1] and a.���=b.����"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "�����Ƿ��ظ�", str����)
            If rsTemp.RecordCount <> 0 Then
                str���� = zlCommFun.IncStr(str����)
            Else
                Exit Do
            End If
        Loop
    End If
        
    '���ݱ���
    If Me.Tag = "����" Or Me.Tag = "��������" Then
        lngItemID = zlDatabase.GetNextId("������ĿĿ¼")
'        If zlClinicCodeRepeat(Trim(Me.txt��Ŀ����.Text)) = True Then Exit Sub
    Else
        If zlClinicCodeRepeat(str����, lngItemID) = True Then Exit Sub
        If zlExistItem("������ĿĿ¼", "ID", lngItemID, Trim(Me.txt��Ŀ����.Text)) = False Then Exit Sub
    End If

    gcnOracle.BeginTrans
    blnBegin = True
    Do While intѭ�� <> 0
        intFirst = intFirst + 1
        gstrSql = "'" & Left(Me.cbo���.Text, 1) & "'," & Me.txt����.Tag & "," & lngItemID & ",'" & str���� & "'"
        gstrSql = gstrSql & ",'" & Trim(Me.txt��Ŀ����.Text) & "','" & Trim(Me.txt����ƴ��.Text) & "','" & Trim(Me.txt�������.Text) & "'"
        gstrSql = gstrSql & ",'" & Trim(Me.txt��������.Text) & "','" & Trim(Me.txt����ƴ��.Text) & "','" & Trim(Me.txt�������.Text) & "'"
        Select Case Left(Me.cbo���.Text, 1)
        Case "C", "D", "F", "G"     '"C-����", "D-���", "F-����", "G-����"
            gstrSql = gstrSql & ",'" & Mid(Me.cbo��������.Text, InStr(1, Me.cbo��������.Text, "-") + 1) & "'"
        Case "E", "H", "Z"              '"E-����", "H-����"
            gstrSql = gstrSql & ",'" & Mid(Me.cbo��������.Text, 1, InStr(Me.cbo��������.Text, "-") - 1) & "'"
        Case Else
            gstrSql = gstrSql & ",''"
        End Select
        gstrSql = gstrSql & "," & Mid(Me.cboִ��Ƶ��.Text, 1, 1) & "," & Me.chk����Ӧ��.Value
        gstrSql = gstrSql & "," & Me.cbo���㷽ʽ.ListIndex & ",'" & Trim(Me.txt���㵥λ.Text) & "'"
        gstrSql = gstrSql & "," & Me.cbo�����Ա�.ListIndex & "," & Me.chkִ�а���.Value
        gstrSql = gstrSql & "," & IIf(Me.chk�������(0).Value = 0, 0, 1) + IIf(Me.chk�������(1).Value = 0, 0, 2) + IIf(Me.chk�������(2).Value = 0, 0, 4)
    
        Select Case Left(Me.cbo���.Text, 1)
        Case "D"
            
            ' strSql & ",0,'" & Trim(Me.txt��鲿λ.Text) & "',null"
            '            �Ƿ������Ŀ0-����( 1-��,û�м�鲿λ)
            '            ���㷽ʽ�̶�ΪNull
                
            '�µķ�ʽ��,�����Ŀû�������Ŀ
                Dim str��鲿λ As String, lngRow As Long
                Dim str���� As String
                Dim strModusSQL() As String, arrItem() As String, lngCount As Long, lngItem As Long
                
                With vfgList
                lngCount = 0
                For lngRow = .FixedRows To .Rows - 1
                    If .RowData(lngRow) = 1 Then
                        str��鲿λ = str��鲿λ & Trim(.Cell(flexcpText, lngRow, .ColIndex("����"))) & ","
                        '�����Ŀ ���ɱ���������Ŀ��λ��SQL
                        str���� = .Cell(flexcpText, lngRow, .ColIndex("����"))
                        arrItem = Split(.Cell(flexcpText, lngRow, .ColIndex("����")), "  ")
                        For i = 0 To UBound(arrItem)
                            If Trim(arrItem(i)) <> "" Then
                                If InStr(arrItem(i), "��") > 0 Then
                                    strTemp = Mid(arrItem(i), 1, InStr(arrItem(i), "��") - 1)
                                    strItem = Mid(arrItem(i), InStr(arrItem(i), "��") + 1, InStr(arrItem(i), "��") - InStr(arrItem(i), "��") - 1)
                                    If InStr(strTemp, "��") > 0 Or InStr(strTemp, "��") > 0 Then
                                        strTemp = Replace(strTemp, "��", "")
                                        strTemp = Trim(Replace(strTemp, "��", ""))
                                        lngCount = lngCount + 1
                                        ReDim Preserve strModusSQL(lngCount) As String
                                        strModusSQL(lngCount) = "zl_������Ŀ��λ_insert(" & lngItemID & ",'" & Mid(cbo��������.Text, InStr(cbo��������, "-") + 1) & "','" & _
                                                        Trim(.Cell(flexcpText, lngRow, .ColIndex("����"))) & "','" & strTemp & "',1,'')"
                                    Else
                                        strTemp = Replace(strTemp, "��", "")
                                        strTemp = Trim(Replace(strTemp, "��", ""))
                                        lngCount = lngCount + 1
                                        ReDim Preserve strModusSQL(lngCount) As String
                                        strModusSQL(lngCount) = "zl_������Ŀ��λ_insert(" & lngItemID & ",'" & Mid(cbo��������.Text, InStr(cbo��������, "-") + 1) & "','" & _
                                                        Trim(.Cell(flexcpText, lngRow, .ColIndex("����"))) & "','" & strTemp & "','','')"
                                    End If
                                    varTemp = Split(strItem, " ")
                                    For j = 0 To UBound(varTemp)
                                        If InStr(varTemp(j), "��") > 0 Then
                                            strTmp = Trim(Replace(varTemp(j), "��", ""))
                                            lngCount = lngCount + 1
                                            ReDim Preserve strModusSQL(lngCount) As String
                                            strModusSQL(lngCount) = "zl_������Ŀ��λ_insert(" & lngItemID & ",'" & Mid(cbo��������.Text, InStr(cbo��������, "-") + 1) & "','" & _
                                                            Trim(.Cell(flexcpText, lngRow, .ColIndex("����"))) & "','" & strTmp & "',1,'" & strTemp & "')"
                                        Else
                                            strTmp = Trim(Replace(varTemp(j), "��", ""))
                                            lngCount = lngCount + 1
                                            ReDim Preserve strModusSQL(lngCount) As String
                                            strModusSQL(lngCount) = "zl_������Ŀ��λ_insert(" & lngItemID & ",'" & Mid(cbo��������.Text, InStr(cbo��������, "-") + 1) & "','" & _
                                                            Trim(.Cell(flexcpText, lngRow, .ColIndex("����"))) & "','" & strTmp & "','','" & strTemp & "')"
                                        End If
                                    Next
                                Else
                                    strTemp = arrItem(i)
                                    If InStr(strTemp, "��") > 0 Or InStr(strTemp, "��") > 0 Then
                                        strTemp = Replace(strTemp, "��", "")
                                        strTemp = Trim(Replace(strTemp, "��", ""))
                                        lngCount = lngCount + 1
                                        ReDim Preserve strModusSQL(lngCount) As String
                                        strModusSQL(lngCount) = "zl_������Ŀ��λ_insert(" & lngItemID & ",'" & Mid(cbo��������.Text, InStr(cbo��������, "-") + 1) & "','" & _
                                                        Trim(.Cell(flexcpText, lngRow, .ColIndex("����"))) & "','" & strTemp & "',1,'')"
                                    Else
                                        lngCount = lngCount + 1
                                        strTemp = Replace(strTemp, "��", "")
                                        strTemp = Trim(Replace(strTemp, "��", ""))
                                        ReDim Preserve strModusSQL(lngCount) As String
                                        strModusSQL(lngCount) = "zl_������Ŀ��λ_insert(" & lngItemID & ",'" & Mid(cbo��������.Text, InStr(cbo��������, "-") + 1) & "','" & _
                                                        Trim(.Cell(flexcpText, lngRow, .ColIndex("����"))) & "','" & strTemp & "','','')"
                                    End If
                                End If
                            End If
                        Next
                    End If
                Next
                End With
                str��鲿λ = zlCommFun.ToVarchar(str��鲿λ, 60)
                
                gstrSql = gstrSql & ",0,'" & str��鲿λ & "',Null"
                
                
                '        If Me.opt��鲿λ(0).Value = True Then
                '            gstrSql = gstrSql & ",0,'" & Trim(Me.txt��鲿λ.Text) & "',null"
                '        Else
                '            gstrSql = gstrSql & ",1,'',null"
                '        End If
        Case "F"
            If Val(Me.txt��׼����.Tag) <> 0 Then
                gstrSql = gstrSql & ",0,''," & Val(Me.txt��׼����.Tag)
            Else
                gstrSql = gstrSql & ",0,'',null"
            End If
        Case "C"
            gstrSql = gstrSql & "," & Me.chk�������.Value & ",'" & Trim(Me.txt�걾��λ.Text) & "',null"
        Case "E"
            If cbo��������.ListIndex = 1 Then
                With vsfTest
                    For i = 1 To .Rows - 1
                        If .TextMatrix(i, 2) = "��" Then
                            str���� = IIf(str���� = "", "", str���� & ",") & .TextMatrix(i, 1) & .TextMatrix(i, 0)
                        Else
                            str���� = IIf(str���� = "", "", str���� & ",") & .TextMatrix(i, 1) & .TextMatrix(i, 0)
                        End If
                    Next
                End With
                strTest = str���� & ";" & str����
                If LenB(strTest) > 60 Then
                    MsgBox "Ƥ�Խ����Ŀ���ַ���̫�࣬�������Ŀ������ַ�����", vbInformation, gstrSysName
                    Me.stbInfo.Tab = 3
                    Exit Sub
                End If
                gstrSql = gstrSql & "," & Me.chk�������.Value & ",'" & strTest & "',null"
            ElseIf cbo��������.ListIndex = 2 Then
                gstrSql = gstrSql & "," & Me.chk�������.Value & ",'" & Trim(Me.cbo����˵��.Text) & "',null"
            Else
                gstrSql = gstrSql & "," & Me.chk�������.Value & ",'',null"
            End If
        Case "Z"
            If cbo��������.ListIndex = 12 Then
                gstrSql = gstrSql & "," & Me.chk�������.Value & ",'" & IIf(chk����.Value = 1, "����", "") & "',null"
            Else
                gstrSql = gstrSql & "," & Me.chk�������.Value & ",'',null"
            End If
        Case Else
            gstrSql = gstrSql & "," & Me.chk�������.Value & ",'',null"
        End Select
    
        If Me.optִ�в���(6).Value Then
            gstrSql = gstrSql & ",6"
        ElseIf Me.optִ�в���(5).Value Then
            gstrSql = gstrSql & ",5"
        ElseIf Me.optִ�в���(4).Value Then
            gstrSql = gstrSql & ",4"
        ElseIf Me.optִ�в���(3).Value Then
            gstrSql = gstrSql & ",3"
        ElseIf Me.optִ�в���(2).Value Then
            gstrSql = gstrSql & ",2"
        ElseIf Me.optִ�в���(1).Value Then
            gstrSql = gstrSql & ",1"
        Else
            gstrSql = gstrSql & ",0"
        End If
    
        If Me.optִ�в���(4).Value Then
            If Me.txt����ִ��.Enabled Then
                gstrSql = gstrSql & "," & IIf(Val(Me.txt����ִ��.Tag) = 0 Or Me.txt����ִ��.Text = "", "null", Val(Me.txt����ִ��.Tag))
            Else
                gstrSql = gstrSql & ",null"
            End If
            If Me.txtסԺִ��.Enabled Then
                gstrSql = gstrSql & "," & IIf(Val(Me.txtסԺִ��.Tag) = 0 Or Me.txtסԺִ��.Text = "", "null", Val(Me.txtסԺִ��.Tag))
            Else
                gstrSql = gstrSql & ",null"
            End If
            strTemp = ""
            strLast = ""
            With Me.msf����ִ��
                For intCount = lngNum To .Rows - 1
                    lngSum = lngSum + 1
                    If Val(.TextMatrix(intCount, colִ�п���ID)) <> 0 Then
                        strMid = Split(.TextMatrix(intCount, col���˿���ID), ",")
                        If UBound(strMid) <> -1 Then
                            For i = LBound(strMid) To UBound(strMid)
                                strTemp = strTemp & "|" & Trim(IIf(strMid(i) = "�����в��ţ�", 0, strMid(i))) & "^" & Trim(.TextMatrix(intCount, colִ�п���ID))
                            Next
                        Else
                            strTemp = strTemp & "|" & 0 & "^" & Trim(.TextMatrix(intCount, colִ�п���ID))
                        End If
                    End If
                    
                    If intCount < .Rows - 1 Then
                        If Len(strTemp) > 4000 Then
                            lngNum = lngNum - 1
                            strTemp = strLast
                            Exit For
                        ElseIf Len(strTemp & .TextMatrix(intCount + 1, col���˿���ID)) > 4000 Then
                            Exit For
                        Else
                            strLast = strTemp
                        End If
                    End If
                Next
                
                lngNum = lngSum + 1
                If intCount = .Rows Or lngSum = 0 Then intѭ�� = 0
            End With
            If strTemp <> "" Then strTemp = Mid(strTemp, 2)
            gstrSql = gstrSql & ",'" & strTemp & "'"
        Else
            intѭ�� = 0
            gstrSql = gstrSql & ",null,null,''"
        End If
    
        For i = 0 To Me.OptApp.Count - 1
            If Me.OptApp(i).Enabled = True And Me.OptApp(i).Value = True Then
                mAppType = i
                Exit For
            End If
        Next
        If Me.OptApp(0).Enabled = False Then
            mAppType = 0
        End If
        If Len(Me.txt�ο�.Tag) = 0 Then
            If Me.Tag = "����" Or Me.Tag = "��������" Then
                gstrSql = gstrSql & ",Null" & "," & mAppType
            Else
                gstrSql = gstrSql & ",Null" & ",0," & mAppType
            End If
        Else
            If Me.Tag = "����" Or Me.Tag = "��������" Then
                gstrSql = gstrSql & "," & Me.txt�ο�.Tag & "," & mAppType
            Else
                gstrSql = gstrSql & "," & Me.txt�ο�.Tag & ",0," & mAppType
            End If
        End If
    
        '¼��������Ӧ�÷�Χ
        gstrSql = gstrSql & "," & Val(Me.txt¼������.Text) & "," & cbo¼��������Χ.ListIndex
        'ִ�б�� (���Լ��գ���ҩ;��Ϊ��ҺʱΪ��Һ����)
        If Left(Me.cbo���.Text, 1) = "E" And Me.cbo��������.ListIndex = 2 And Me.cboִ�з���.ListIndex = 1 Then
            gstrSql = gstrSql & "," & IIf(Val(Me.cbo��Һ����.ListIndex) = 0, 0, 2)
        ElseIf Left(Me.cbo���.Text, 1) = "E" And Me.cbo��������.ListIndex = 1 Then
            gstrSql = gstrSql & "," & IIf(1 = Val(Me.chkNoTMSY.Value), 2, 0)
        Else
            gstrSql = gstrSql & "," & Val(Me.chk����.Value)
        End If
        
        'ִ�з���
        If Left(Me.cbo���.Text, 1) = "E" And Me.cbo��������.ListIndex = 2 Then
            '-��Һ��ע�䣬�������ڷ�
            gstrSql = gstrSql & "," & Val(cboִ�з���.List(cboִ�з���.ListIndex))
        ElseIf Left(Me.cbo���.Text, 1) = "E" And Me.cbo��������.ListIndex = 1 Then
            'Ƥ��
            gstrSql = gstrSql & IIf(1 = Val(Me.chkYYPS.Value), ",5", ",3")
        ElseIf Left(Me.cbo���.Text, 1) = "D" And Me.cbo��������.Text = "18-����" Then
            '����
            If UBound(Split(cbo�������.Text, "-")) > 0 Then
                gstrSql = gstrSql & "," & Val(Split(cbo�������.Text, "-")(0))
            End If
        ElseIf Left(Me.cbo���.Text, 1) = "E" And Me.cbo��������.ListIndex = 8 Then
            gstrSql = gstrSql & "," & Val(cboBloodType.ListIndex)
        Else
            '����
            gstrSql = gstrSql & ",0"
        End If
        
        
        'վ��
        gstrSql = gstrSql & "," & IIf(cmbStationNo.Visible = False Or cmbStationNo.Text = "", "Null", strվ��)
        
        'Ƶ������
        If Left(Me.cbo���.Text, 1) <> "C" And stbInfo.TabVisible(4) = True Then
            With vsfFreq
                For i = 1 To .Rows - 1
                    If .TextMatrix(i, 0) = "��" Then
                        strFreq = IIf(strFreq = "", "", strFreq & "|") & .TextMatrix(i, 1)
                    End If
                Next
            End With
        End If
        gstrSql = gstrSql & "," & IIf(strFreq = "", "Null", "'" & strFreq & "'")
        
        '�������
        If Me.cbo�������.Visible = True Then
            gstrSql = gstrSql & "," & Val(cbo�������.List(cbo�������.ListIndex))
        Else
            gstrSql = gstrSql & ",Null"
        End If
        
        'ʹ�ÿ���
        gstrSql = gstrSql & ",'" & strDeptId & "'"
        'ʹ�ÿ��ҷ�Χ
        For i = 0 To Me.OptAppUse.Count - 1
            If Me.OptAppUse(i).Enabled = True And Me.OptAppUse(i).Value = True Then
                mAppType = i
                Exit For
            End If
        Next
        If Me.OptAppUse(0).Enabled = False Then
            mAppType = 0
        End If
        gstrSql = gstrSql & "," & mAppType
        
        '59964-�Ƿ��һ��ִ��
        gstrSql = gstrSql & "," & IIf(intFirst = 1, 1, 0)
        
        gstrSql = gstrSql & "," & IIf(Me.txtML.Text = "", "NULL", Val(Me.txtML.Text))
        '��Ѫ�������
        With vsfBloodLis
            For intRow = 1 To .Rows - 1
                If .TextMatrix(intRow, 0) <> "" Then
                    str��Ѫ������� = str��Ѫ������� & "|" & .TextMatrix(intRow, 0)
                End If
            Next
            str��Ѫ������� = Mid(str��Ѫ�������, 2)
        End With
        gstrSql = gstrSql & "," & IIf(str��Ѫ������� = "", "NULL", "'" & str��Ѫ������� & "'")
        
        '��������Ƶ��
        gstrSql = gstrSql & "," & IIf(Me.cboZLPL.Text = "", "NULL", "'" & Mid(Mid(Me.lblZLPL.Tag, InStr(1, Me.lblZLPL.Tag, "|" & Me.cboZLPL.Text & "-") + Len(Me.cboZLPL.Text) + 2), 1, InStr(1, Mid(Me.lblZLPL.Tag, InStr(1, Me.lblZLPL.Tag, "|" & Me.cboZLPL.Text & "-") + Len(Me.cboZLPL.Text) + 2), "|") - 1) & "'")
        
        '���������Ѫ�ɼ���ʽ�������Ӧ���Թܱ���
        str�Թܱ��� = ""
        If Left(Me.cbo���.Text, 1) = "E" And Me.cbo��������.ListIndex = 9 Then
            If cboTestTubeCode.ListIndex > 0 Then
                str�Թܱ��� = Split(cboTestTubeCode.List(cboTestTubeCode.ListIndex), "-")(0)
            End If
        End If
        
        If Me.Tag = "����" Then
            gstrSql = "zl_������Ŀ_Insert(" & gstrSql & IIf(str�Թܱ��� = "", "", ",0,'" & str�Թܱ��� & "'") & ")"
        ElseIf Me.Tag = "��������" Then
            gstrSql = "zl_������Ŀ_Insert(" & gstrSql & "," & mlngOldId & IIf(str�Թܱ��� = "", "", ",'" & str�Թܱ��� & "'") & ")"
        Else
            gstrSql = "zl_������Ŀ_Update(" & gstrSql & IIf(str�Թܱ��� = "", "", ",'" & str�Թܱ��� & "'") & ")"
        End If
    
        err = 0: On Error GoTo ErrHand
        Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    Loop
    
    '�����Ŀ ������Ŀ��λ����
    If Left(Me.cbo���.Text, 1) = "D" Then
        '����RIS�ӿڣ�����/�޸�������Ŀ
        If mblnPACSInterface = True Then
            If Not gobjRIS Is Nothing Then
                If gobjRIS.HISBasicDictTable(RISBaseItemType.ClinicItem, IIf(Me.Tag = "����", RISBaseItemOper.AddNew, RISBaseItemOper.Modify), lngItemID) <> 1 Then
                    gcnOracle.RollbackTrans
                    
                    '����ʱ��ʾ�ӿڴ�����Ϣ
                    If gobjRIS.LastErrorInfo <> "" Then
                        MsgBox gobjRIS.LastErrorInfo, vbInformation, gstrSysName
                    Else
                        MsgBox "����RIS�ӿڴ��󣬲��ܼ�����ǰ����������ϵͳ����Ա��ϵ", vbInformation, gstrSysName
                    End If
                    
                    Exit Sub
                End If
                
                '����RIS�ӿڣ�ɾ��������Ŀ��λ���ŵ�HISɾ������֮ǰ
                If Me.Tag = "�޸�" Then
                    If gobjRIS.HISBasicDictTable(RISBaseItemType.ClinicItemPart, RISBaseItemOper.Delete, lngItemID) <> 1 Then
                        gcnOracle.RollbackTrans
                        
                        '����ʱ��ʾ�ӿڴ�����Ϣ
                        If gobjRIS.LastErrorInfo <> "" Then
                            MsgBox gobjRIS.LastErrorInfo, vbInformation, gstrSysName
                        Else
                            MsgBox "����RIS�ӿڴ��󣬲��ܼ�����ǰ����������ϵͳ����Ա��ϵ", vbInformation, gstrSysName
                        End If
                        
                        Exit Sub
                    End If
                End If
                
                blnRisTrans = True
            Else
                '�ӿڲ�����Чʱ��ֹ����ʾ
                gcnOracle.RollbackTrans
                
                MsgBox "RIS�ӿڴ���ʧ�ܣ����ܼ�����ǰ�����������ǽӿ��ļ���װ��ע�᲻����������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
                
                Exit Sub
            End If
        End If
        
        'HISɾ��/�޸���Ŀ��λ
        gstrSql = "zl_������Ŀ��λ_Delete(" & lngItemID & ")"
        Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    
        If str��鲿λ <> "" Then
            For lngCount = LBound(strModusSQL) To UBound(strModusSQL) - 1
                If strModusSQL(lngCount + 1) <> "" Then
                    Call zlDatabase.ExecuteProcedure(strModusSQL(lngCount + 1), Me.Caption)
                    
                End If
            Next
            
            '����RIS�ӿڣ�����������Ŀ��λ
            '�ŵ�HIS��������֮��
            If mblnPACSInterface = True Then
                If Not gobjRIS Is Nothing Then
                    If gobjRIS.HISBasicDictTable(RISBaseItemType.ClinicItemPart, RISBaseItemOper.AddNew, lngItemID) <> 1 Then
                        gcnOracle.RollbackTrans
                        
                        '����ʱ��ʾ�ӿڴ�����Ϣ
                        If gobjRIS.LastErrorInfo <> "" Then
                            MsgBox gobjRIS.LastErrorInfo, vbInformation, gstrSysName
                        Else
                            MsgBox "����RIS�ӿڴ��󣬲��ܼ�����ǰ����������ϵͳ����Ա��ϵ", vbInformation, gstrSysName
                        End If
                        
                        Exit Sub
                    End If
                    
                    blnRisTrans = True
                Else
                    gcnOracle.RollbackTrans
                    
                    '�ӿڲ�����Чʱ��ֹ����ʾ
                    MsgBox "RIS�ӿڴ���ʧ�ܣ����ܼ�����ǰ�����������ǽӿ��ļ���װ��ע�᲻����������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
                    
                    Exit Sub
                End If
            End If
        End If
    End If
    
    '�������ݱ���
    lngVItemID0 = lngVItemID
    If Left(Me.cbo���.Text, 1) = "C" Then
        
        gstrSql = "Null,'" & str���� & "','" & Me.txt��Ŀ���� & "','" & _
            Me.txt����ƴ�� & "'," & "0,10,0,Null," & _
            "Null,0,Null,Null,Null,Null,Null,0"

        If lngVItemID0 = 0 And (Me.Tag = "����" Or Me.Tag = "��������") Then
            lngVItemID0 = zlDatabase.GetNextId("����������Ŀ")
            gstrSql = "ZL_������Ŀ_INSERT(" & lngVItemID0 & "," & gstrSql & ")"
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
        Else
            strSql = "Select id,����id,����,������,Ӣ����,����,����,С��,��λ,�ٴ�����,��ʾ��,�Ա���,��ֵ��,��ʼֵ,���ֱ���,��ֵ���� From ����������Ŀ Where id=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngVItemID0)
            Do Until rsTmp.EOF
                gstrSql = "'" & rsTmp!����id & "','" & str���� & "','" & Me.txt��Ŀ���� & "','" & _
                        Me.txtӢ�� & "'," & rsTmp!���� & "," & rsTmp!���� & "," & rsTmp!С�� & ",'" & Me.txt���㵥λ & "','" & _
                        rsTmp!�ٴ����� & "'," & rsTmp!��ʾ�� & "," & rsTmp!�Ա��� & ",'" & rsTmp!��ֵ�� & "','" & rsTmp!��ʼֵ & "','" & _
                        rsTmp!���ֱ��� & "','" & rsTmp!��ֵ���� & "'"
                rsTmp.MoveNext
                gstrSql = "ZL_������Ŀ_UPDATE(" & lngVItemID0 & "," & gstrSql & ")"
                Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
            Loop
        End If
        
        If Not (lngVItemID0 = 0 And Me.Tag <> "����") Then '������ʾ����������,�����ж�,���ִ������ݲ�����.
            gstrSql = "Select �������,��Ŀ���,�������,��λ,��ӡ����,��ӡ���,���㹫ʽ,���鷽��,�ϲ������,����쳣���� From ������Ŀ Where ������Ŀid=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngVItemID0)
            Do Until rsTmp.EOF
                gstrSql = "'" & Me.txtӢ�� & "','" & rsTmp!������� & "','" & rsTmp!��Ŀ��� & "','" & rsTmp!������� & "','" & Me.txt���㵥λ & "','" & _
                          rsTmp!��ӡ���� & "','" & rsTmp!��ӡ��� & "','" & rsTmp!���㹫ʽ & "','" & rsTmp!���鷽�� & "','" & rsTmp!�ϲ������ & "','" & _
                          rsTmp!����쳣���� & "'"
                          
                gstrSql = "ZL_������Ŀ_UPDATE(" & lngVItemID0 & "," & gstrSql & ")"
                Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
                rsTmp.MoveNext
            Loop
            gstrSql = "'^" & lngVItemID0 & "'"
            gstrSql = "ZL_���鱨����Ŀ_UPDATE(" & lngItemID & "," & gstrSql & ")"
            Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
        End If
    ElseIf lngVItemID0 > 0 Then
        'ɾ��ԭ��������Ŀ�ı�����Ŀ
        gstrSql = "ZL_���鱨����Ŀ_UPDATE(" & lngItemID & ",'')"
        Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
        
        '�����������Ŀ��Ϊ�����Ŀ����ɾ������������Ŀ
        gstrSql = "ZL_������Ŀ_DELETE(" & lngVItemID0 & ")"
        Call zlDatabase.ExecuteProcedure(gstrSql, Me.Caption)
    End If

    lngVItemID = lngVItemID0
    
    gcnOracle.CommitTrans
    
    blnBegin = False
    blnRisTrans = False

    mblnOK = True
    '�������Ӵ���
    If Me.Tag = "����" Or Me.Tag = "��������" Then
        If chkGoOn.Value Then
            
            mbln�������� = True
            lngItemID = 0
            mLast�������� = ""
            Call Form_Activate
            Me.stbInfo.Tab = 0
            Me.txt����.SetFocus
            Exit Sub
        End If
    End If
    Unload Me
    Exit Sub

ErrHand:
    If blnBegin Then gcnOracle.RollbackTrans
    
    'Ris�ӿں�HIS��ͬ��ʱ��д������־
    If blnRisTrans = True And Not gobjRIS Is Nothing Then
        MsgBox "HIS" & IIf(Me.Tag = "����", "����", "�޸�") & "������Ŀ����RIS�ӿں�HIS���ݲ�ͬ��������ϵͳ����Ա��ϵ��", vbInformation, gstrSysName
        
        On Error Resume Next
        Call gobjRIS.WriteCommLog("frmClinicItem��cmdOK_Click", "HIS" & IIf(Me.Tag = "����", "����", "�޸�") & "������Ŀ����RIS�ӿں�HIS���ݲ�ͬ��", "������ĿID=" & lngItemID, 0)
    End If
        
    'If ErrCenter() = 1 Then Resume
    Call ErrCenter
    'Resume
    Call SaveErrLog
End Sub


Private Sub IniStationNo()
    Dim strSql As String
    Dim rsRecord As ADODB.Recordset
'    lblStationNo.Visible = False
'    cmbStationNo.Visible = False
'
'    If gstrNodeNo <> "-" Then
    On Error GoTo ErrHandle
        lblStationNo.Visible = True
        cmbStationNo.Visible = True
        
        strSql = "select ���,���� from zlnodelist"
        Set rsRecord = zlDatabase.OpenSQLRecord(strSql, "վ���ѯ")
        With cmbStationNo
            .AddItem ""
            Do While Not rsRecord.EOF
                .AddItem rsRecord!��� & "-" & rsRecord!����
                rsRecord.MoveNext
            Loop
        End With
        
'        With cmbStationNo
'            .Clear
'            .AddItem ""
'            .AddItem "0"
'            .AddItem "1"
'            .AddItem "2"
'            .AddItem "3"
'            .AddItem "4"
'            .AddItem "5"
'            .AddItem "6"
'            .AddItem "7"
'            .AddItem "8"
'            .AddItem "9"
'
'            .ListIndex = 0
'        End With
'    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub
Private Sub SetStationNo(ByVal strNo As String)
    Dim n As Integer
    
'    If gstrNodeNo = "-" Then Exit Sub
    
    If strNo = "" Then
        cmbStationNo.ListIndex = 0
    Else
        For n = 1 To cmbStationNo.ListCount - 1
            If Mid(cmbStationNo.List(n), 1, InStr(1, cmbStationNo.List(n), "-") - 1) = strNo Then
                cmbStationNo.ListIndex = n
            End If
        Next
    End If
        
End Sub
Private Sub cmdTestAdd_Click()
    Dim n As Integer
    Dim str��ע As String
    Dim str���� As String
    Dim str���� As String
    
    If Trim(txtƤ�Ա�ע.Text) = "" Or Trim(txtƤ������.Text) = "" Then Exit Sub
    
    str��ע = "(" & Trim(txtƤ�Ա�ע.Text) & ")"
    str���� = Trim(txtƤ������.Text)
    str���� = IIf(chkƤ�Թ���.Value = 1, "��", "")
    
    With vsfTest
        '����Ƿ��ظ�
        For n = 1 To .Rows - 1
            If .TextMatrix(n, 0) <> "" Then
                If .TextMatrix(n, 0) = str��ע And .TextMatrix(n, 1) = str���� Then
                    MsgBox "���ظ���Ŀ�����������룡", vbExclamation, gstrSysName
                    Exit Sub
                End If
            End If
        Next
    
        '��������Ŀ
        .Rows = .Rows + 1
        .TextMatrix(.Rows - 1, 0) = str��ע
        .TextMatrix(.Rows - 1, 1) = str����
        .TextMatrix(.Rows - 1, 2) = str����
    End With
End Sub

Private Sub cmdTestDel_Click()
    Dim int�������� As Integer
    Dim int�������� As Integer
    Dim n As Integer
    Dim intCurr As Integer
    
    With vsfTest
        If .Row > 0 Then
            '��ǰ�����Ժ���������
            For n = 1 To .Rows - 1
                If .TextMatrix(n, 2) = "��" Then
                    int�������� = int�������� + 1
                Else
                    int�������� = int�������� + 1
                End If
            Next
            
            '����Ҫ��֤���Ժ�������Ŀ��ʣ��1��
            If (.TextMatrix(.Row, 2) = "��" And int�������� <= 1) Or (.TextMatrix(.Row, 2) = "" And int�������� <= 1) Then
                MsgBox "����ɾ������Ŀ������Ҫ��֤���Ժ�������Ŀ��ʣ��1��", vbInformation, gstrSysName
                Exit Sub
            End If
            
            .RemoveItem .Row
        End If
    End With
End Sub
Private Sub cmd�걾_Click()
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    Dim vPoint As POINTAPI

    On Error GoTo ErrHandle
    strSql = "Select Rownum As ID, ����, ����, ���� From ���Ƽ���걾 Order By ����"

    vPoint = zlControl.GetCoordPos(txt�걾��λ.hWnd, txt�걾��λ.Left - 165, txt�걾��λ.Top - 30)

    Set rsTmp = zlDatabase.ShowSelect(Me, strSql, 0, "�걾��λ", , , , , True, True, vPoint.x, vPoint.y)

    If rsTmp.State = 0 Then
        Me.txt�걾��λ.Text = ""
        If Trim(Me.txt�걾��λ.Text) = "" And Trim(Me.txt�걾��λ.Tag) <> "" Then
            Me.txt�걾��λ.Text = Trim(Me.txt�걾��λ.Tag)
        End If
        Exit Sub
    End If
    If Not rsTmp Is Nothing Then
        Me.txt�걾��λ.Text = rsTmp("����")
        Me.txt�걾��λ.Tag = rsTmp("����")
    Else
        If Trim(Me.txt�걾��λ.Text) = "" And Trim(Me.txt�걾��λ.Tag) <> "" Then
            Me.txt�걾��λ.Text = Trim(Me.txt�걾��λ.Tag)
        Else
            Me.txt�걾��λ.Text = ""
            Me.txt�걾��λ.SetFocus
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub cmd�ο�_Click()
    Dim rsTmp As ADODB.Recordset

    Set rsTmp = SelectRefer
    If Not rsTmp Is Nothing Then
        Me.txt�ο� = rsTmp("����"): Me.txt�ο�.Tag = rsTmp("ID"): strRefer = Me.txt�ο�
    Else
        MsgBox "û���ҵ��ɲο�����Ŀ��", vbInformation, Me.Caption
    End If
End Sub

Private Function SelectRefer(Optional ByVal strName As String = "") As ADODB.Recordset
    Dim strSql As String, strSQLItem As String
    Dim rsTmp As New ADODB.Recordset, iAttr As Integer

    strSql = "Select ���� From ���Ʒ���Ŀ¼ Where ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, lngClassId)

    If rsTmp.EOF Then
        iAttr = -1
    Else
        iAttr = rsTmp(0)
    End If
    If Len(strName) = 0 Then
        strSql = "Select 0 As ĩ��,ID,�ϼ�ID,����,����,'' As ˵�� From ���Ʋο����� a" & _
            " Where ����=" & iAttr & _
            " Start With a.�ϼ�id Is Null Connect By Prior a.id=a.�ϼ�id " & _
            " Union All" & _
            " Select 1,ID,����ID,����,����,˵�� From ���Ʋο�Ŀ¼ a Where ����=" & iAttr & " Order By ����"
    Else
        strSQLItem = " From ���Ʋο�Ŀ¼ A,���Ʋο����� B" & _
            " Where A.ID=B.�ο�Ŀ¼ID And A.����=" & iAttr & _
            " And (Upper(A.����) Like '" & UCase(strName) & "%'" & _
            " Or Upper(A.����) Like '" & mstrMatch & UCase(strName) & "%'" & _
            " Or Upper(B.����) Like '" & mstrMatch & UCase(strName) & "%'" & _
            " Or Upper(B.����) Like '" & mstrMatch & UCase(strName) & "%')"

        strSql = "Select Distinct 0 As ĩ��,ID,�ϼ�ID,����,����,'' As ˵�� From ���Ʋο����� a" & _
            " Where ����=" & iAttr & _
            " Start With ID In (Select ����ID " & strSQLItem & ") Connect By Prior a.�ϼ�id=a.id " & _
            " Union All" & _
            " Select Distinct 1,A.ID,A.����ID,A.����,A.����,A.˵�� " & strSQLItem & " Order By ����"
    End If
    Set SelectRefer = zlDatabase.ShowSelect(Me, strSql, 2, "�ο�", , , , , True)
End Function

Private Sub cmd����_Click()
    With Me.tvwClass
        .Left = Me.txt����.Left
        .Top = Me.txt����.Top + Me.txt����.Height
        .ZOrder 0: .Visible = True
        .SetFocus
    End With
End Sub

Private Sub InitVsf()
    '��ʼ����Ѫ�����ձ�
    With vsfBloodLis
        .Cols = 2
        .ColHidden(0) = True '���ص�һ��
        .ExtendLastCol = True '���һ����������
        .ColComboList(1) = "|..."
        .Editable = flexEDKbdMouse
        .AllowSelection = False '���ܶ�ѡ��Ԫ��
    End With
End Sub

Private Sub Form_Activate()
    Dim aTmp() As String
    Dim strTmp As String
    Dim i As Integer
    Dim n As Integer
        
    If mFromLoad And Not mbln�������� Then Exit Sub
    If Me.Tag = "����" Or Me.Tag = "��������" Then chkGoOn.Visible = True
    mFromLoad = True
    
    stbInfo.TabVisible(3) = False
    
    Call GetDefineSize
    Call IniStationNo

    '��ȡִ����Ŀ����Ϣ
    err = 0: On Error GoTo ErrHand

    '�����װ����ܵ����������Ըı䣬��������װ��
    gstrSql = "select ��� from ������ĿĿ¼ where ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)

    With rsTemp
        If .RecordCount > 0 Then
            For intCount = 0 To Me.cbo���.ListCount - 1
                If Left(Me.cbo���.List(intCount), 1) = IIf(IsNull(!���), "", !���) Then
                    Me.cbo���.ListIndex = intCount: Exit For
                End If
            Next
        End If
    End With

    'װ����������������
    gstrSql = "select A.����,A.����,ִ��Ƶ��,����Ӧ��,���㷽ʽ,���㵥λ,�����Ա�,ִ�а���,�������,ִ�п���,��������,�����Ŀ,�걾��λ,A.�Թܱ���," & _
            "        ����ʱ��,nvl(����ʱ��,to_date('3000-01-01','YYYY-MM-DD')) as ����ʱ��,�ο�Ŀ¼ID,B.���� As �ο�����,A.¼������,A.ִ�б��,A.ִ�з���,A.�������,A.վ��,A.����ϵ��,A.����Ƶ�ʱ��� " & _
            " from ������ĿĿ¼ A,���Ʋο�Ŀ¼ B" & _
            " where A.�ο�Ŀ¼ID=B.ID(+) And A.ID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)
    
    cboִ�з���.Visible = False
    lblִ�з���.Visible = False
    cboִ�з���.ListIndex = 0
    
    With rsTemp
        Me.txt��Ŀ����.MaxLength = .Fields("����").DefinedSize
        If .RecordCount > 0 Then
            Me.txt��Ŀ����.Text = !����
            Me.txt��Ŀ����.Text = !����
            Me.txt��Ŀ����.Tag = !����
            '�������ִ��Ƶ�ʿ�ѡΪ"0-��ѡƵ��;2-������"����Ҫ��������
            If Left(Me.cbo���.Text, 1) = "H" Then
                If Val(IIf(IsNull(!ִ��Ƶ��), 0, !ִ��Ƶ��)) = 2 Then
                    Me.cboִ��Ƶ��.ListIndex = 1
                Else
                    Me.cboִ��Ƶ��.ListIndex = 0
                End If
            Else
                Me.cboִ��Ƶ��.ListIndex = IIf(IsNull(!ִ��Ƶ��), 0, !ִ��Ƶ��)
            End If
            
            Call Init����Ƶ��(NVL(!����Ƶ�ʱ���))
            
            Me.chk����Ӧ��.Value = IIf(IsNull(!����Ӧ��), 0, !����Ӧ��)
            Me.chk����Ӧ��.Tag = IIf(IsNull(!����Ӧ��), 0, !����Ӧ��)
            Me.cbo���㷽ʽ.ListIndex = IIf(IsNull(!���㷽ʽ), 0, !���㷽ʽ)
            Me.txt���㵥λ.Text = IIf(IsNull(!���㵥λ), "", !���㵥λ)
            Me.cbo�������.ListIndex = IIf(IsNull(!�������), 0, !�������)
            Me.cbo�����Ա�.ListIndex = IIf(IsNull(!�����Ա�), 0, !�����Ա�)
            Me.chkִ�а���.Value = IIf(IsNull(!ִ�а���), 0, !ִ�а���)
            Me.chk�������.Value = IIf(IsNull(!�����Ŀ), 0, !�����Ŀ)
            Me.txt¼������.Text = IIf(IsNull(!¼������), "", !¼������)
            Me.txtML.Text = IIf(IsNull(!����ϵ��), "", !����ϵ��)
            SetStationNo IIf(IsNull(!վ��), "", !վ��)
            Select Case !�������
            Case 4
                Me.chk�������(2).Value = 1:
                Me.chk�������(0).Value = 0: Me.chk�������(1).Value = 0
            Case 3
                Me.chk�������(0).Value = 1: Me.chk�������(1).Value = 1
            Case 2
                Me.chk�������(0).Value = 0: Me.chk�������(1).Value = 1
            Case 1
                Me.chk�������(0).Value = 1: Me.chk�������(1).Value = 0
            Case Else
                Me.chk�������(0).Value = 0: Me.chk�������(1).Value = 0
            End Select
            Me.optִ�в���(IIf(IsNull(!ִ�п���), 0, !ִ�п���)).Value = True
            If Format(!����ʱ��, "YYYY-MM-DD") = "3000-01-01" Then
                Me.lblFound.Caption = "����Ŀ��" & Format(!����ʱ��, "YYYY-MM-DD") & "������"
            Else
                Me.lblFound.Caption = ""
            End If
            If Left(Me.cbo���.Text, 1) = "D" Then
                If IIf(IsNull(!�����Ŀ), 0, !�����Ŀ) = 0 Then
                    Me.opt��鲿λ(0).Value = True: Me.txt��鲿λ.Text = IIf(IsNull(!�걾��λ), "", !�걾��λ): Me.txt��鲿λ.Enabled = True
                Else
                    Me.opt��鲿λ(1).Value = True: Me.txt��鲿λ.Text = "": Me.txt��鲿λ.Enabled = False
                End If
            End If
            If Left(Me.cbo���.Text, 1) = "C" Then
                Me.txt�걾��λ.Text = IIf(IsNull(!�걾��λ), "", !�걾��λ)
                Me.txt�걾��λ.Tag = Me.txt�걾��λ.Text
            End If
            Select Case Left(Me.cbo���.Text, 1)
            Case "C", "D", "F", "G"     'C-����, D-���, F-����, G-����
                For intCount = 0 To Me.cbo��������.ListCount - 1
                    If Mid(Me.cbo��������.List(intCount), InStr(1, Me.cbo��������.List(intCount), "-") + 1) = IIf(IsNull(!��������), "", !��������) Then
                        If mLast�������� = "" Then
                            Me.cbo��������.ListIndex = intCount
                            mLast�������� = !��������
                            Exit For
                        End If
                    End If
                Next
                Me.chk������� = NVL(!�����Ŀ, 0)
                If Me.chk�������.Value = 1 Then
                    Me.chk����Ӧ��.Enabled = False
                Else
                    Me.chk����Ӧ��.Enabled = True
                End If
                If Me.chk�������(1).Value = 1 Then
                    Me.chk����.Enabled = True
                    Me.chk����.Value = NVL(!ִ�б��, 0)
                Else
                    Me.chk����.Enabled = False
                    Me.chk����.Value = 0
                End If
                
                '�ж���ǰ���õ���ʲôֵ
                If !ִ�з��� = "" Then
                    cbo�������.ListIndex = 0
                Else
                    For n = 1 To cbo�������.ListCount - 1
                        If Val(Mid(cbo�������.List(n), 1, InStr(1, cbo�������.List(n), "-") - 1)) = !ִ�з��� Then
                            cbo�������.ListIndex = n
                        End If
                    Next
                End If
                
            Case "E", "H"         'E-����, H-����
                For intCount = 0 To Me.cbo��������.ListCount - 1
                    If Val(Left(Me.cbo��������.List(intCount), 1)) = Val(IIf(IsNull(!��������), "", !��������)) Then
                        Me.cbo��������.ListIndex = intCount: Exit For
                    End If
                Next
                If Me.cbo��������.ListIndex = 2 Then
                    Me.cbo����˵��.Text = IIf(IsNull(!�걾��λ), "", !�걾��λ)
                End If
                
                If Left(Me.cbo���.Text, 1) = "E" And Me.cbo��������.ListIndex = 2 Then
                    For intCount = 0 To Me.cboִ�з���.ListCount - 1
                        If Val(Me.cboִ�з���.List(intCount)) = Val(IIf(IsNull(!ִ�з���), "", !ִ�з���)) Then
                            Me.cboִ�з���.ListIndex = intCount: Exit For
                        End If
                    Next
                    
                    If Me.cboִ�з���.ListIndex = 1 Then
                        If Val(IIf(IsNull(!ִ�б��), "", !ִ�б��)) = 2 Then
                            Me.cbo��Һ����.ListIndex = 1
                        Else
                            Me.cbo��Һ����.ListIndex = 0
                        End If
                    End If
                    
                    cboִ�з���.Visible = True
                    lblִ�з���.Visible = True
                End If
                If Left(Me.cbo���.Text, 1) = "E" And Me.cbo��������.ListIndex = 8 Then
                    cboBloodType.ListIndex = Val(!ִ�з��� & "")
                    cboBloodType.Visible = True
                    lblִ�з���.Visible = True
                End If
                If Left(Me.cbo���.Text, 1) = "E" And Me.cbo��������.ListIndex = 9 Then
                    Call Load�Թܱ���(NVL(!�Թܱ��� & ""))
                    Me.lbl�Թܱ���.Visible = True
                    Me.picTestTubeCode.Visible = True
                    Me.cboTestTubeCode.Visible = True
                End If
                If Left(Me.cbo���.Text, 1) = "E" And Me.cbo��������.ListIndex = 1 Then
                    Me.chkNoTMSY.Value = IIf(2 = Val(!ִ�б�� & ""), 1, 0)
                    Me.chkYYPS.Value = IIf(5 = Val(!ִ�з��� & ""), 1, 0)
                End If
            Case "Z"            'Z-����
                For intCount = 0 To Me.cbo��������.ListCount - 1
                    If Mid(Me.cbo��������.List(intCount), 1, InStr(1, Me.cbo��������.List(intCount), "-") - 1) = IIf(IsNull(!��������), "", !��������) Then
                        Me.cbo��������.ListIndex = intCount: Exit For
                    End If
                Next
                If Me.cbo��������.ListIndex = 12 Then
                    If IIf(IsNull(!�걾��λ), "", !�걾��λ) = "����" Then
                        chk����.Value = 1
                    End If
                End If
            Case "K"
                '����Ѫ����������ݲ�ѯ����
                If Me.Tag = "�޸�" Or Me.Tag = "����" Or Me.Tag = "��������" Then
                    gstrSql = "Select ID, ���� From ������ĿĿ¼ Where ID In (Select ������Ŀid From ��Ѫ������� Where ��Ŀid = [1])"
                    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)
                    If rsTemp.RecordCount > 0 Then
                        With vsfBloodLis
                            .Rows = rsTemp.RecordCount + 1
                            For n = 1 To rsTemp.RecordCount
                                .TextMatrix(n, 0) = rsTemp!ID
                                .TextMatrix(n, 1) = rsTemp!����
                                rsTemp.MoveNext
                            Next
                        End With
                    End If
                End If
            Case Else
            End Select

            Me.txt�ο� = NVL(!�ο�����): Me.txt�ο�.Tag = NVL(!�ο�Ŀ¼ID): strRefer = Me.txt�ο�
        End If
    End With

    If Left(Me.cbo���.Text, 1) = "F" Then
        gstrSql = "select I.ID,I.����,I.��������,I.����" & _
                " from ��������Ŀ¼ I,������϶��� R" & _
                " where I.ID=R.����ID and R.����ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)

        If rsTemp.RecordCount > 0 Then
            Me.txt��׼����.Tag = rsTemp!ID
            Me.txt��׼����.Text = IIf(IsNull(rsTemp!����), "", rsTemp!����)
            Me.lbl��׼����.Caption = IIf(IsNull(rsTemp!��������), "", "��" & rsTemp!�������� & "��") & IIf(IsNull(rsTemp!����), "", rsTemp!����)
        End If
    End If

    gstrSql = "select ����,����,����,���� from ������Ŀ���� where ������ĿID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)

    With rsTemp
        Do While Not .EOF
            If !���� = 1 And !���� = 1 Then Me.txt����ƴ��.Text = !����
            If !���� = 1 And !���� = 2 Then Me.txt�������.Text = !����
            If !���� = 9 And !���� = 1 Then Me.txt��������.Text = !����: Me.txt����ƴ��.Text = !����
            If !���� = 9 And !���� = 2 Then Me.txt��������.Text = !����: Me.txt�������.Text = !����
            .MoveNext
        Loop
    End With
    
    'ʹ�ÿ���
    If Me.Tag <> "����" Then
        gstrSql = "Select b.id,b.���� from �������ÿ��� A,���ű� B Where a.��ĿID=[1] And a.����id=b.id"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)
    
        With rsTemp
            i = 0: n = 0
            Do While Not .EOF
                vsUseDept.TextMatrix(i, n) = rsTemp!���� & ""
                vsUseDept.Cell(flexcpData, i, n) = rsTemp!���� & ""
                vsUseDept.TextMatrix(i, n + 5) = rsTemp!ID
                vsUseDept.Cell(flexcpData, i, n + 5) = rsTemp!ID & ""
                mstr��ѡʹ�ÿ��� = IIf(mstr��ѡʹ�ÿ��� = "", "", mstr��ѡʹ�ÿ��� & ";") & rsTemp!ID & "," & rsTemp!����
                If i = vsUseDept.Rows - 1 And n = 4 Then
                    vsUseDept.AddItem ""
                End If
                If n = 4 Then
                    n = 0
                    i = i + 1
                Else
                    n = n + 1
                End If
                .MoveNext
            Loop
        End With
    End If

    gstrSql = "select R.������Դ,E.ID,E.����" & _
            " from ����ִ�п��� R,���ű� E" & _
            " where R.ִ�п���ID=E.ID and R.������Դ in (1,2) and R.��������id is null and R.������ĿID=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)

    With rsTemp
        Do While Not .EOF
            If !������Դ = 1 Then Me.txt����ִ��.Text = !����: Me.txt����ִ��.Tag = !ID
            If !������Դ = 2 Then Me.txtסԺִ��.Text = !����: Me.txtסԺִ��.Tag = !ID
            .MoveNext
        Loop
    End With

    gstrSql = "select K.ID as ��������ID,K.���� as �������ұ���,K.���� as ������������," & _
            "         E.ID as ִ�в���ID,E.���� as ִ�п��ұ���,E.���� as ִ�в�������" & _
            " from ����ִ�п��� R,���ű� K,���ű� E" & _
            " where R.��������ID=K.ID(+) and R.ִ�п���ID=E.ID and nvl(R.������Դ,0)=0 and R.������ĿID=[1] " & _
            " order by e.����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)
    
    

    With rsTemp
        Me.msf����ִ��.Clear (1)

        Do While Not .EOF
            'If Me.msf����ִ��.Rows - 1 < .AbsolutePosition Then Me.msf����ִ��.Rows = Me.msf����ִ��.Rows + 1

            If strTmp <> !ִ�в������� Then
                i = i + 1
                Me.msf����ִ��.Rows = i + 1
                Me.msf����ִ��.TextMatrix(i, col���˿���ID) = IIf(IsNull(!��������ID), "�����в��ţ�", !��������ID)
                msf����ִ��.Cell(flexcpData, i, col���˿���ID) = msf����ִ��.TextMatrix(i, col���˿���ID)
                Me.msf����ִ��.TextMatrix(i, col���˿���) = IIf(IsNull(!��������ID), "�����в��ţ�", !������������)
                msf����ִ��.Cell(flexcpData, i, col���˿���) = msf����ִ��.TextMatrix(i, col���˿���)
                Me.msf����ִ��.TextMatrix(i, colִ�п���ID) = !ִ�в���ID
                msf����ִ��.Cell(flexcpData, i, colִ�п���ID) = msf����ִ��.TextMatrix(i, colִ�п���ID)
                Me.msf����ִ��.TextMatrix(i, colִ�п���) = !ִ�в�������
                msf����ִ��.Cell(flexcpData, i, colִ�п���) = msf����ִ��.TextMatrix(i, colִ�п���)
            Else
                Me.msf����ִ��.TextMatrix(i, col���˿���ID) = Me.msf����ִ��.TextMatrix(i, col���˿���ID) & "," & !��������ID
                msf����ִ��.Cell(flexcpData, i, col���˿���ID) = msf����ִ��.TextMatrix(i, col���˿���ID)
                Me.msf����ִ��.TextMatrix(i, col���˿���) = Me.msf����ִ��.TextMatrix(i, col���˿���) & "," & !������������
                msf����ִ��.Cell(flexcpData, i, col���˿���) = msf����ִ��.TextMatrix(i, col���˿���)
            End If

            strTmp = !ִ�в�������
            .MoveNext
        Loop
    End With

    '��ѯ����������Ŀ��Ӧ�ļ���ָ��
    If Left(Me.cbo���.Text, 1) = "C" And Me.chk�������.Value = 0 And Me.Tag <> "����" Then
        gstrSql = "Select A.*,B.ID,B.�ٴ�����,B.������ " & _
            "From ������Ŀ A,����������Ŀ B,���鱨����Ŀ C " & _
            "Where A.������ĿID=B.ID And B.ID=C.������ĿID And C.������ĿID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)

        With rsTemp
            If Not .EOF Then
                lngVItemID = !ID
                Me.txtӢ��.Text = "" & !��д
'                Me.txt����(0) = Nvl(!��д)
'                Me.txt����(1) = Nvl(!��λ)
'                Me.cbo��Ŀ����.ListIndex = Nvl(!��Ŀ���, 1) - 1
'                Me.cbo�������.ListIndex = Nvl(!�������, 1) - 1
'                Me.txt����(2) = TransFormula1(Nvl(!���㹫ʽ))
'                If Len(Nvl(!����쳣����)) > 0 Then
'                    aTmp = Split(!����쳣����, ";")
'                    Me.txt����(3) = aTmp(0)
'                    If UBound(aTmp) > 0 Then Me.txt����(4) = aTmp(1)
'                End If
'                Me.txt����(5) = Nvl(!�ٴ�����)
            End If
        End With
    End If

    If Me.Tag = "����" Or Me.Tag = "��������" Then
        lngItemID = 0: lngVItemID = 0
        If Val(zlDatabase.GetPara(61, glngSys)) = 0 Then '������Ŀ�������ģʽ
            gstrSql = "select nvl(max(����),'0000000') as ����" & _
                    " From ������ĿĿ¼" & _
                    " Where ��� >= 'A'"
'            If rsTemp.State = adStateOpen Then rsTemp.Close
'            Call SQLTest(App.ProductName, Me.Caption, gstrSql)
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, "Form_Activate")
'            Call SQLTest
            Me.txt��Ŀ����.Text = zlCommFun.IncStr(rsTemp!����)
        Else
            strTemp = Mid(Me.txt����.Text, 2, InStr(1, Me.txt����.Text, "]") - 2)
            
            gstrSql = "select nvl(max(����),'0000000') as ����" & _
                    " From ������ĿĿ¼" & _
                    " Where ��� >= 'A' and ���� like [1]"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, strTemp & "%")
    
            err = 0: On Error Resume Next
            If rsTemp!���� = "0000000" Then
                Me.txt��Ŀ����.Text = zlCommFun.IncStr(strTemp & "000")
            Else
                Me.txt��Ŀ����.Text = zlCommFun.IncStr(rsTemp!����)
            End If
        End If

        '���������Ϣ
        Me.txt��Ŀ����.Text = "": Me.txt����ƴ��.Text = "": Me.txt�������.Text = ""
        Me.txt��������.Text = "": Me.txt����ƴ��.Text = "": Me.txt�������.Text = "": Me.txtӢ��.Text = ""
        Me.lblFound.Visible = False
        If Me.Tag = "����" Then Me.txt�ο� = "": Me.txt�ο�.Tag = "": strRefer = ""
    End If

'        If Me.Tag = "�޸�" Then
'            Me.chk�������.Enabled = False
'        End If

    If Me.Tag = "����" Then
        Me.cmdOK.Visible = False
        Me.cmdCancel.Caption = "�ر�(&C)"
        Me.txt����.Enabled = False: Me.cmd����.Enabled = False: Me.cbo���.Enabled = False
        Me.txt��Ŀ����.Enabled = False: Me.cbo��������.Enabled = False
        Me.txt��Ŀ����.Enabled = False: Me.txt����ƴ��.Enabled = False: Me.txt�������.Enabled = False
        Me.txt��������.Enabled = False: Me.txt����ƴ��.Enabled = False: Me.txt�������.Enabled = False

        Me.cboִ��Ƶ��.Enabled = False: Me.chk����Ӧ��.Enabled = False
        Me.cbo���㷽ʽ.Enabled = False: Me.txt���㵥λ.Enabled = False
        Me.cbo�����Ա�.Enabled = False: Me.chkִ�а���.Enabled = False
        Me.fra��鲿λ.Enabled = False: Me.txtML.Enabled = False
        Me.txt�ο�.Enabled = False: Me.cmd�ο�.Enabled = False
        Me.fra�걾��λ.Enabled = False: Me.cmd�걾.Enabled = False
        Me.cboZLPL.Enabled = False


        Me.chk�������(0).Enabled = False: Me.chk�������(1).Enabled = False: Me.chk�������(2).Enabled = False
        Me.fraִ�в���.Enabled = False
        Me.fra��׼����.Enabled = False

        Me.chk�������.Enabled = False
        Me.txtӢ��.Enabled = False
'        Me.txt����(0).Enabled = False: Me.txt����(1).Enabled = False
'        Me.txt����(2).Enabled = False: Me.txt����(3).Enabled = False
'        Me.txt����(4).Enabled = False: Me.txt����(5).Enabled = False
'        Me.cbo�������.Enabled = False: Me.cbo��Ŀ����.Enabled = False
        Me.txt¼������.Enabled = False
        Me.cbo¼��������Χ.Enabled = False
        Me.cbo����˵��.Enabled = False
        Me.cmbStationNo.Enabled = False
        Me.chk����.Enabled = False
        For i = 0 To OptAppUse.Count - 1
            OptAppUse(i).Enabled = False
        Next
    End If

    '�ж��Ƿ��Ǽ������еĲ�λ��Ŀ
    mbln��ϲ�λ��Ŀ = False
    If lngItemID <> 0 And Left(Me.cbo���.Text, 1) = "D" Then
        '�����Ŀ
        gstrSql = "Select 1 From ������Ŀ��� Where �������id = [1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)

        mbln�����Ŀ = (rsTemp.RecordCount > 0)

        If Not mbln�����Ŀ Then
            '����Ŀ
            gstrSql = "Select 1 From ������ĿĿ¼ Where ID In (Select �������id From ������Ŀ��� Where ������Ŀid = [1]) And ��� = 'D'"
            Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, lngItemID)

            mbln��ϲ�λ��Ŀ = (rsTemp.RecordCount > 0)
            Me.opt��鲿λ(1).Enabled = Not mbln��ϲ�λ��Ŀ
        End If

    End If

    '���Ѿ�����Ϊ�������е�����Ŀ���Լ���������Ŀ������Ŀ������������������
    If mbln�����Ŀ = True Or mbln��ϲ�λ��Ŀ = True Then
        Me.cbo���.Enabled = False
    End If
    
    If Me.Tag = "����" Then
        Call cbo���_Click
        Call cbo��������_Click
    End If
    
    If Me.Tag = "�޸�" And Left(Me.cbo���.Text, 1) = "G" Then
        Me.chk����Ӧ��.Value = 0: Me.chk����Ӧ��.Enabled = False
    End If
    
    '�Ƿ���������
    chkGoOn.Value = Val(zlDatabase.GetPara("������Ŀ��������", glngSys, 1054, 0, Array(Me.chkGoOn), True))
    msf����ִ��.AutoSize msf����ִ��.FixedCols, msf����ִ��.Cols - 1
    Call chk�������_Click(0)
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        If Me.picDept.Visible = True Then
            Call picDept_LostFocus
            Exit Sub
        End If
        If Me.tvwClass.Visible Then
            Me.tvwClass.Visible = False: Me.txt����.SetFocus: Exit Sub
        End If
        Call cmdCancel_Click
        
    ElseIf KeyCode = vbKeyF3 Then
        If txtLocate.Enabled And txtLocate.Visible Then Call txtLocate_KeyPress(vbKeyReturn)
    ElseIf KeyCode = vbKeyF4 Then
        If txtLocate.Visible And txtLocate.Enabled Then txtLocate.SetFocus
    End If
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If InStr("'", Chr(KeyAscii)) > 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub Form_Load()
    Dim i As Long
    Dim strSql As String
    Dim rsData As ADODB.Recordset
    
    mstrFindStyle = IIf(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", 0) = "0", "%", "")
    mblnPACSInterface = (Val(zlDatabase.GetPara(255, glngSys, , "0")) = 1)
    mstrӦ�÷�Χ = zlDatabase.GetPara("��ĿӦ�÷�Χ", glngSys, 1054, "000")
    
    With Me.msf����ִ��
        .Editable = flexEDKbdMouse
        .FixedCols = 1: .Rows = 2: .Cols = 4
        .TextMatrix(0, colִ�п���ID) = "ִ�п���ID": .TextMatrix(0, colִ�п���) = "ִ�п���"
        .TextMatrix(0, col���˿���ID) = "���˿���ID": .TextMatrix(0, col���˿���) = "���˿���"
        .colData(colִ�п���ID) = 5: .colData(colִ�п���) = 1: .colData(col���˿���ID) = 5: .colData(col���˿���) = 1
        .ColWidth(colִ�п���ID) = 0: .ColWidth(colִ�п���) = 1700: .ColWidth(col���˿���ID) = 0: .ColWidth(col���˿���) = 7300
        .Row = 1: .Col = 1
        .ColHidden(colִ�п���ID) = True: .ColHidden(col���˿���ID) = True
        .ExplorerBar = flexExSortShowAndMove
        .AllowUserResizing = flexResizeBothUniform
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .RowHeightMin = 250
    End With
    vsUseDept.Editable = flexEDKbdMouse
    For i = 0 To OptAppUse.Count - 1
        If i = 0 Then
            OptAppUse(i).Enabled = True
        Else
            '���ݲ�����ȷ���Ƿ����
            OptAppUse(i).Enabled = (Val(Mid(mstrӦ�÷�Χ, i, 1)) = 1)
        End If
    Next
    
    Me.lvwItems.ListItems.Clear
    With Me.lvwItems.ColumnHeaders
        .Clear
        .Add , "����", "����", 1500
        .Add , "����", "����", 900
    End With
    With Me.lvwItems
        .ColumnHeaders("����").Position = 1
        .SortKey = .ColumnHeaders("����").Index - 1
        .SortOrder = lvwAscending
        .Width = 3000
    End With
    
    
    Me.lvwItem.ListItems.Clear
    With Me.lvwItem.ColumnHeaders
        .Clear
        .Add , "����", "����", 1500
        .Add , "����", "����", 900
        .Add , "���", "���", 0
    End With
    
    With Me.lvwItem
        .ColumnHeaders("����").Position = 1
        .SortKey = .ColumnHeaders("����").Index - 1
        .SortOrder = lvwAscending
        .Width = 3000
    End With
    
    With cbo�������
        .Clear
        .AddItem "0-����"
        .AddItem "1-����"
        .AddItem "2-ϸ��"
        .AddItem "3-����"
        .AddItem "4-ʬ��"
        .AddItem "5-����ʯ��"
        .ListIndex = 0
    End With
        
        strSql = "select ID,���� from ����������"
    Set rsData = zlDatabase.OpenSQLRecord(strSql, Me.Caption)

    If rsData.RecordCount > 0 Then
        With cbo�������
        .Clear
            rsData.MoveFirst
            Do While Not rsData.EOF
                If NVL(rsData!����, "  ") <> "  " Then
                    .AddItem NVL(rsData!ID, 0) & "-" & rsData!����
                End If
                rsData.MoveNext
            Loop
        .ListIndex = 0
        End With
    End If
    
    With Me.cbo¼��������Χ
        .Clear
        .AddItem "����Ŀ"
        .AddItem "����"
        .AddItem "������"
        .AddItem "�����"
        .AddItem "����"
        .ListIndex = 0
    End With
    
    
    With Me.cboִ�з���
        .Clear
        .AddItem "0-����"
        .AddItem "1-��Һ"
        .AddItem "2-ע��"
        .AddItem "4-�ڷ�"
        .ListIndex = 0
    End With
    
    With Me.cbo��Һ����
        .Clear
        .AddItem "0-����"
        .AddItem "2-����Ӫ��"
        .ListIndex = 0
    End With
    
    With Me.cboBloodType '��Ѫ;������ 0����Ѫ��1����Ѫ
        .Clear
        .AddItem "0-��Ѫ"
        .AddItem "1-��Ѫ"
        .ListIndex = 0
    End With
    
    Call InitVsf '��ʼ�����
    
    mstrMatch = gstrMatch
    strRefer = ""
    mLast�������� = ""
    mblnOK = False
    mlngFind = 1
    Ini���ʷ���
    Call Init����Ƶ��
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mstr��ѡʹ�ÿ��� = ""
    Call zlDatabase.SetPara("������Ŀ��������", chkGoOn.Value, glngSys, 1054)
    mFromLoad = False
End Sub

Private Sub lvwItems_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwItems.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwItems.SortOrder = IIf(Me.lvwItems.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwItems.SortKey = ColumnHeader.Index - 1
        Me.lvwItems.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwItems_DblClick()
    Dim i As Integer
    Dim m As Integer
    Dim blnBatch As Boolean
    Dim str���˿���ID As String
    Dim str���˿������� As String
    Dim strTmp As String
    Dim strArr
    Dim n As Integer
    Dim strNew As String
    Dim blnNew As Boolean
        
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    With Me.lvwItems
        Select Case .Tag
        Case "����"
            Me.txt��׼����.Tag = Mid(.SelectedItem.Key, 2)
            Me.txt��׼����.Text = .SelectedItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1)
            Me.lbl��׼����.Caption = .SelectedItem.Text
            Me.stbInfo.Tab = 1: Me.chk�������(0).SetFocus
        Case "����"
            Me.txt����ִ��.Tag = Mid(.SelectedItem.Key, 2)
            Me.txt����ִ��.Text = .SelectedItem.Text
            Me.txt����ִ��.SetFocus: Call zlCommFun.PressKey(vbKeyTab)
        Case "סԺ"
            Me.txtסԺִ��.Tag = Mid(.SelectedItem.Key, 2)
            Me.txtסԺִ��.Text = .SelectedItem.Text
            Me.txtסԺִ��.SetFocus: Call zlCommFun.PressKey(vbKeyTab)
        Case "����"
            With Me.lvwItems
                If Me.msf����ִ��.Col = col���˿��� And Me.lvwItems.Checkboxes = True Then
                    Me.msf����ִ��.Text = ""
                    For i = 1 To .ListItems.Count
                        If .ListItems(i).Checked = True Then
                            If Me.msf����ִ��.Text = "" Then
                                Me.msf����ִ��.Text = .ListItems(i).Text
                                Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, col���˿���ID) = Mid(.ListItems(i).Key, 2)
                                msf����ִ��.Cell(flexcpData, Me.msf����ִ��.Row, col���˿���ID) = msf����ִ��.TextMatrix(Me.msf����ִ��.Row, col���˿���ID)
                                Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, col���˿���) = Me.msf����ִ��.Text
                                msf����ִ��.Cell(flexcpData, Me.msf����ִ��.Row, col���˿���) = msf����ִ��.TextMatrix(Me.msf����ִ��.Row, col���˿���)
                            Else
                                Me.msf����ִ��.Text = Me.msf����ִ��.Text & "," & .ListItems(i).Text
                                Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, col���˿���ID) = Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, col���˿���ID) & "," & Mid(.ListItems(i).Key, 2)
                                msf����ִ��.Cell(flexcpData, Me.msf����ִ��.Row, col���˿���ID) = msf����ִ��.TextMatrix(Me.msf����ִ��.Row, col���˿���ID)
                                Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, col���˿���) = Me.msf����ִ��.Text
                                msf����ִ��.Cell(flexcpData, Me.msf����ִ��.Row, col���˿���) = msf����ִ��.TextMatrix(Me.msf����ִ��.Row, col���˿���)
                            End If
                            m = m + 1
                        End If
                    Next
                    If m = 0 Then
                        Me.msf����ִ��.Text = ""
                        Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, col���˿���ID) = "�����в��ţ�"
                        msf����ִ��.Cell(flexcpData, Me.msf����ִ��.Row, col���˿���ID) = msf����ִ��.TextMatrix(Me.msf����ִ��.Row, col���˿���ID)
                        Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, col���˿���) = "�����в��ţ�"
                        msf����ִ��.Cell(flexcpData, Me.msf����ִ��.Row, col���˿���) = msf����ִ��.TextMatrix(Me.msf����ִ��.Row, col���˿���)
                    End If
                Else
                    Me.msf����ִ��.Text = .SelectedItem.Text
                    Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, col���˿���ID) = Mid(.SelectedItem.Key, col���˿���ID)
                    msf����ִ��.Cell(flexcpData, Me.msf����ִ��.Row, col���˿���ID) = msf����ִ��.TextMatrix(Me.msf����ִ��.Row, col���˿���ID)
                    Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, col���˿���) = Me.msf����ִ��.Text
                    msf����ִ��.Cell(flexcpData, Me.msf����ִ��.Row, col���˿���) = msf����ִ��.TextMatrix(Me.msf����ִ��.Row, col���˿���)
                End If
            End With
            
            '���������δ����У�ѯ���Ƿ�ͬһ��������
            For i = 1 To Me.msf����ִ��.Rows - 1
                If Me.msf����ִ��.TextMatrix(i, colִ�п���ID) <> "" And Me.msf����ִ��.TextMatrix(i, col���˿���) = "" Then
                    blnBatch = True
                    Exit For
                End If
            Next
            
            If blnBatch = True Then
                If MsgBox("�Ƿ�Ӧ��������δ���õ��У�", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
                    str���˿���ID = Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, col���˿���ID)
                    str���˿������� = Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, col���˿���)
                    For i = 1 To Me.msf����ִ��.Rows - 1
                        If Me.msf����ִ��.TextMatrix(i, col���˿���) = "" Then
                            Me.msf����ִ��.TextMatrix(i, col���˿���ID) = str���˿���ID
                            msf����ִ��.Cell(flexcpData, i, col���˿���ID) = msf����ִ��.TextMatrix(i, col���˿���ID)
                            Me.msf����ִ��.TextMatrix(i, col���˿���) = str���˿�������
                            msf����ִ��.Cell(flexcpData, i, col���˿���) = msf����ִ��.TextMatrix(i, col���˿���)
                        End If
                    Next
                End If
            End If
            
            Me.msf����ִ��.SetFocus
            Call zlCommFun.PressKey(vbKeyReturn)
        Case "ִ��"
            
            If Val(Me.picDept.Tag) = 1 And lbl��������.Visible = True Then
                'ɾ��������ѡ���б��е�ִ�п���
                For i = msf����ִ��.Rows - 1 To 1 Step -1
                    If InStr(mstr��ѡִ�п���, msf����ִ��.TextMatrix(i, colִ�п���ID) & "," & msf����ִ��.TextMatrix(i, colִ�п���)) = 0 Then
                        If i > 1 Then
                            msf����ִ��.RemoveItem i
                        Else
                            msf����ִ��.TextMatrix(1, colִ�п���ID) = ""
                            msf����ִ��.Cell(flexcpData, 1, colִ�п���ID) = ""
                            msf����ִ��.TextMatrix(1, colִ�п���) = ""
                            msf����ִ��.Cell(flexcpData, 1, colִ�п���) = ""
                            msf����ִ��.TextMatrix(1, col���˿���ID) = ""
                            msf����ִ��.Cell(flexcpData, 1, col���˿���ID) = ""
                            msf����ִ��.TextMatrix(1, col���˿���) = ""
                            msf����ִ��.Cell(flexcpData, 1, col���˿���) = ""
                            
                            If msf����ִ��.Rows > 2 Then
                                msf����ִ��.RemoveItem 1
                            End If
                        End If
                    End If
                Next
                
                '������ִ�п���
                mstr��ѡִ�п��� = mstr��ѡִ�п��� & ";"
                strArr = Split(mstr��ѡִ�п���, ";")
                
                For i = 0 To UBound(strArr) - 1
                    blnNew = True
                    If strArr(i) <> "" Then
                        For n = 1 To msf����ִ��.Rows - 1
                            If strArr(i) = msf����ִ��.TextMatrix(n, colִ�п���ID) & "," & msf����ִ��.TextMatrix(n, colִ�п���) Then
                                blnNew = False
                            End If
                        Next
                        If blnNew = True Then
                            strNew = IIf(strNew = "", "", strNew & ";") & strArr(i)
                        End If
                    End If
                Next
                
                If strNew <> "" Then
                    strArr = Split(strNew & ";", ";")
                    For i = 0 To UBound(strArr) - 1
                        If strArr(i) <> "" Then
                            If msf����ִ��.TextMatrix(msf����ִ��.Rows - 1, colִ�п���) <> "" Then
                                msf����ִ��.Rows = msf����ִ��.Rows + 1
                            End If
                            msf����ִ��.TextMatrix(msf����ִ��.Rows - 1, colִ�п���ID) = Split(strArr(i), ",")(0)
                            msf����ִ��.Cell(flexcpData, msf����ִ��.Rows - 1, colִ�п���ID) = msf����ִ��.TextMatrix(msf����ִ��.Rows - 1, colִ�п���ID)
                            msf����ִ��.TextMatrix(msf����ִ��.Rows - 1, colִ�п���) = Split(strArr(i), ",")(1)
                            msf����ִ��.Cell(flexcpData, msf����ִ��.Rows - 1, colִ�п���) = msf����ִ��.TextMatrix(msf����ִ��.Rows - 1, colִ�п���)
                        End If
                    Next
                End If
                
                msf����ִ��.Row = msf����ִ��.Rows - 1
                Me.msf����ִ��.SetFocus
                Call zlCommFun.PressKey(vbKeyRight)
            Else
                Me.msf����ִ��.Text = .SelectedItem.Text
                Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, colִ�п���ID) = Mid(.SelectedItem.Key, 2)
                msf����ִ��.Cell(flexcpData, msf����ִ��.Row, colִ�п���ID) = msf����ִ��.TextMatrix(msf����ִ��.Row, colִ�п���ID)
                Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, colִ�п���) = Me.msf����ִ��.Text
                msf����ִ��.Cell(flexcpData, msf����ִ��.Row, colִ�п���) = msf����ִ��.TextMatrix(msf����ִ��.Row, colִ�п���)
                Me.msf����ִ��.SetFocus
                Call zlCommFun.PressKey(vbKeyRight)
            End If
        Case "ʹ��"
            Dim j As Long
            
            If Val(Me.picDept.Tag) = 2 And lbl��������.Visible = True Then
                'ɾ��������ѡ���б��е�ʹ�ÿ���
                For i = 0 To vsUseDept.Rows - 1
                    For j = 0 To 4
                        If InStr(mstr��ѡʹ�ÿ���, vsUseDept.TextMatrix(i, j + 5) & "," & vsUseDept.TextMatrix(i, j)) = 0 Then
                            vsUseDept.TextMatrix(i, j) = ""
                            vsUseDept.Cell(flexcpData, i, j) = ""
                            vsUseDept.TextMatrix(i, j + 5) = ""
                            vsUseDept.Cell(flexcpData, i, j + 5) = ""
                        End If
                    Next
                Next
                
                '������ִ�п���
                mstr��ѡʹ�ÿ��� = mstr��ѡʹ�ÿ��� & ";"
                strArr = Split(mstr��ѡʹ�ÿ���, ";")
                
                For i = 0 To UBound(strArr) - 1
                    blnNew = True
                    If strArr(i) <> "" Then
                        For n = 0 To vsUseDept.Rows - 1
                            For j = 0 To 4
                                If strArr(i) = vsUseDept.TextMatrix(n, j + 5) & "," & vsUseDept.TextMatrix(n, j) Then
                                    blnNew = False
                                End If
                            Next
                        Next
                        If blnNew = True Then
                            strNew = IIf(strNew = "", "", strNew & ";") & strArr(i)
                        End If
                    End If
                Next
                
                If strNew <> "" Then
                    strArr = Split(strNew & ";", ";")
                    For i = 0 To UBound(strArr) - 1
                        If strArr(i) <> "" Then
                            For n = 0 To vsUseDept.Rows - 1
                                For j = 0 To 4
                                    If n = vsUseDept.Rows - 1 And j = 4 Then vsUseDept.AddItem ""
                                    If vsUseDept.TextMatrix(n, j) = "" Then
                                        vsUseDept.TextMatrix(n, j) = Split(strArr(i), ",")(1)
                                        vsUseDept.Cell(flexcpData, n, j) = vsUseDept.TextMatrix(n, j)
                                        vsUseDept.TextMatrix(n, j + 5) = Split(strArr(i), ",")(0)
                                        vsUseDept.Cell(flexcpData, n, j + 5) = vsUseDept.TextMatrix(n, j + 5)
                                        n = vsUseDept.Rows - 1
                                        Exit For
                                    End If
                                Next
                            Next
                        End If
                    Next
                End If
                
                Me.vsUseDept.SetFocus
            Else
                If InStr(mstr��ѡʹ�ÿ���, Mid(.SelectedItem.Key, 2) & "," & .SelectedItem.Text) > 0 Then
                    MsgBox "�Ѿ�������ͬ��ʹ�ÿ����ˣ����顣", vbInformation, gstrSysName
                    Me.vsUseDept.TextMatrix(Me.vsUseDept.Row, vsUseDept.Col) = ""
                    vsUseDept.SetFocus
                Else
                    Me.vsUseDept.Text = .SelectedItem.Text
                    Me.vsUseDept.TextMatrix(Me.vsUseDept.Row, vsUseDept.Col + 5) = Mid(.SelectedItem.Key, 2)
                    vsUseDept.Cell(flexcpData, vsUseDept.Row, vsUseDept.Col + 5) = vsUseDept.TextMatrix(vsUseDept.Row, vsUseDept.Col + 5)
                    Me.vsUseDept.TextMatrix(Me.vsUseDept.Row, vsUseDept.Col) = Me.vsUseDept.Text
                    vsUseDept.Cell(flexcpData, vsUseDept.Row, vsUseDept.Col) = vsUseDept.TextMatrix(vsUseDept.Row, vsUseDept.Col)
                    Me.vsUseDept.SetFocus
                    mstr��ѡʹ�ÿ��� = IIf(mstr��ѡʹ�ÿ��� = "", "", mstr��ѡʹ�ÿ��� & ";") & Mid(.SelectedItem.Key, 2) & "," & Me.vsUseDept.Text
                    Call zlCommFun.PressKey(vbKeyRight)
                End If
            End If
        End Select
        
        DoEvents
        picDept.Visible = False
        txtFind.Text = ""
    End With
End Sub

Private Sub lvwItems_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
        If lvwItems.SelectedItem.Checked = False And KeyAscii = vbKeyReturn Then
            lvwItems.SelectedItem.Checked = Not lvwItems.SelectedItem.Checked
            Exit Sub
        End If
        If lvwItems.Checkboxes = True And KeyAscii = vbKeySpace Then Exit Sub
        Call lvwItems_DblClick
    Case vbKeyEscape
        picDept.Visible = False
        txtFind.Text = ""
    End Select
End Sub


Private Sub lvwItems_LostFocus()
    Call picDept_LostFocus
End Sub

Private Sub lvwItems_GotFocus()
    If Me.lvwItems.Tag = "����" Or Me.lvwItems.Tag = "ִ��" Or Me.lvwItems.Tag = "ʹ��" Then
        Me.lvwItems.ToolTipText = "ȫѡCtrl+A��ȫ��Ctrl+R"
    Else
        Me.lvwItems.ToolTipText = ""
    End If
End Sub


Private Sub lvwItem_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If Me.lvwItem.SortKey = ColumnHeader.Index - 1 Then
        Me.lvwItem.SortOrder = IIf(Me.lvwItem.SortOrder = lvwAscending, lvwDescending, lvwAscending)
    Else
        Me.lvwItem.SortKey = ColumnHeader.Index - 1
        Me.lvwItem.SortOrder = lvwAscending
    End If
End Sub

Private Sub lvwItem_DblClick()
    Dim i As Integer
    Dim m As Integer
    If Me.lvwItem.SelectedItem Is Nothing Then Exit Sub
    With Me.lvwItem
        Select Case .Tag
        Case "����"
            Me.txt��׼����.Tag = Mid(.SelectedItem.Key, 2)
            Me.txt��׼����.Text = .SelectedItem.SubItems(.ColumnHeaders("����").Index - 1)
'            Me.lbl��׼����.Caption = .SelectedItem.Text & vbCrLf & .SelectedItem.SubItems(.ColumnHeaders("���").Index - 1)
            Me.lbl��׼����.Caption = IIf(.SelectedItem.SubItems(.ColumnHeaders("���").Index - 1) = "", "", "��" & .SelectedItem.SubItems(.ColumnHeaders("���").Index - 1) & "��") & .SelectedItem.Text
            
            Me.stbInfo.Tab = 0: Me.chk�������(0).SetFocus
        Case "����"
            Me.txt����ִ��.Tag = Mid(.SelectedItem.Key, 2)
            Me.txt����ִ��.Text = .SelectedItem.Text
            Me.txt����ִ��.SetFocus: Call zlCommFun.PressKey(vbKeyTab)
        Case "סԺ"
            Me.txtסԺִ��.Tag = Mid(.SelectedItem.Key, 2)
            Me.txtסԺִ��.Text = .SelectedItem.Text
            Me.txtסԺִ��.SetFocus: Call zlCommFun.PressKey(vbKeyTab)
        Case "����"
            With Me.lvwItems
                If Me.msf����ִ��.Col = 3 And Me.lvwItems.Checkboxes = True Then
                    For i = 1 To .ListItems.Count
                        If .ListItems(i).Checked = True Then
                            If Me.msf����ִ��.Text = "" Then
                                Me.msf����ִ��.Text = "[" & .ListItems(i).SubItems(.ColumnHeaders("����").Index - 1) & "]" & .ListItems(i).Text
                                Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, 2) = Mid(.ListItems(i).Key, 2)
                                Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, 3) = Me.msf����ִ��.Text
                            Else
                                Me.msf����ִ��.Text = Me.msf����ִ��.Text & ",[" & .ListItems(i).SubItems(.ColumnHeaders("����").Index - 1) & "]" & .ListItems(i).Text
                                Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, 2) = Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, 2) & "," & Mid(.ListItems(i).Key, 2)
                                Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, 3) = Me.msf����ִ��.Text
                            End If
                            m = m + 1
                        End If
                    Next
                    If m = 0 Then
                        Me.msf����ִ��.Text = ""
                        Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, 2) = "�����в��ţ�"
                        Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, 3) = "�����в��ţ�"
                    End If
                Else
                    Me.msf����ִ��.Text = "[" & .SelectedItem.SubItems(.ColumnHeaders("����").Index - 1) & "]" & .SelectedItem.Text
                    Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, 0) = Mid(.SelectedItem.Key, 2)
                    Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, 1) = Me.msf����ִ��.Text
                End If
            End With
            Me.msf����ִ��.SetFocus
            Call zlCommFun.PressKey(vbKeyReturn)
        Case "ִ��"
            Me.msf����ִ��.Text = "[" & .SelectedItem.SubItems(.ColumnHeaders("����").Index - 1) & "]" & .SelectedItem.Text
            Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, 0) = Mid(.SelectedItem.Key, 2)
            Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, 1) = Me.msf����ִ��.Text
            Me.msf����ִ��.SetFocus
            Call zlCommFun.PressKey(vbKeyRight)
        End Select
    End With
End Sub

Private Sub lvwItem_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.lvwItem.SelectedItem Is Nothing Then Exit Sub
        If lvwItem.Checkboxes = True And KeyAscii = vbKeySpace Then Exit Sub
        Call lvwItem_DblClick
    Case vbKeyEscape
        Call lvwItem_LostFocus
    End Select
End Sub

Private Sub lvwItem_LostFocus()
    Me.lvwItem.Visible = False
End Sub

Private Sub msf����ִ��_GotFocus()
    If Me.lvwItems.Visible Then Me.lvwItems.SetFocus
End Sub

Private Sub optList_Click(Index As Integer)
    Dim lngRow As Long
    For lngRow = vfgList.FixedRows To vfgList.Rows - 1
        If vfgList.RowData(lngRow) = 0 Then
            If Index = 1 Then
                vfgList.RowHidden(lngRow) = True
            Else
                vfgList.RowHidden(lngRow) = False
            End If
        End If
    Next
End Sub

Private Sub opt��鲿λ_Click(Index As Integer)
    If Me.opt��鲿λ(0).Value = True Then
        Me.txt��鲿λ.Enabled = True
    Else
        Me.txt��鲿λ.Enabled = False
    End If
End Sub
Private Sub opt��鲿λ_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = 0 Then
        If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
    Else
        Me.stbInfo.Tab = 1: Me.chk�������(0).SetFocus
    End If
End Sub

Private Sub optִ�в���_Click(Index As Integer)
    If Me.optִ�в���(4).Value = True Then
        If Me.chk�������(0).Value = 1 Then
            Me.txt����ִ��.Enabled = True
            txt����ִ��.BackColor = vbWindowBackground
        Else
            Me.txt����ִ��.Enabled = False
            txt����ִ��.BackColor = vbButtonFace
        End If
        If Me.chk�������(1).Value = 1 Then
            Me.txtסԺִ��.Enabled = True
            txtסԺִ��.BackColor = vbWindowBackground
        Else
            Me.txtסԺִ��.Enabled = False
            txtסԺִ��.BackColor = vbButtonFace
        End If
        load���ʷ��� 0
        txtLocate.Enabled = True
        fraDeptFind.Enabled = True
        txtLocate.BackColor = vbWindowBackground
        optDeptKind(0).Enabled = True
        optDeptKind(1).Enabled = True
        msf����ִ��.Editable = flexEDKbdMouse
    Else
        Me.txt����ִ��.Enabled = False: Me.txtסԺִ��.Enabled = False
        txt����ִ��.BackColor = vbButtonFace: txtסԺִ��.BackColor = vbButtonFace
        txtLocate.Enabled = False
        fraDeptFind.Enabled = False
        txtLocate.BackColor = vbButtonFace
        optDeptKind(0).Enabled = False
        optDeptKind(1).Enabled = False
        msf����ִ��.Editable = flexEDNone
    End If
End Sub

Private Sub optִ�в���_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub

Private Sub stbInfo_Click(PreviousTab As Integer)
    If Me.tvwClass.Visible Then Me.stbInfo.Tab = 0: Me.tvwClass.SetFocus: Exit Sub
    If Me.lvwItems.Visible Then Me.stbInfo.Tab = 1: Me.lvwItems.SetFocus: Exit Sub
    
    Select Case Me.stbInfo.Tab
    Case 0
        If Me.txt��Ŀ����.Enabled Then Me.txt��Ŀ����.SetFocus
    Case 1
        If Me.chk�������(1).Enabled Then Me.chk�������(1).SetFocus
        If Me.chk�������(0).Enabled Then Me.chk�������(0).SetFocus
    Case 2
        '�����Ŀ ��ʾ��鲿λѡ��ҳ
        If Me.vfgList.Enabled Then
            Me.vfgList.SetFocus
        End If
    Case 3
        txtƤ�Ա�ע.SetFocus
    End Select
End Sub

Private Sub tvwClass_DblClick()
    If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
    Me.txt����.Tag = Mid(Me.tvwClass.SelectedItem.Key, 2)
    Me.txt����.Text = Me.tvwClass.SelectedItem.Text
    Me.txt����.SetFocus
End Sub

Private Sub tvwClass_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn
        If Me.tvwClass.SelectedItem Is Nothing Then Exit Sub
        If Me.tvwClass.SelectedItem.Children > 0 Then Exit Sub
        Call tvwClass_DblClick
    Case vbKeySpace
        Call tvwClass_DblClick
    Case vbKeyEscape
        Call tvwClass_LostFocus
    End Select
End Sub

Private Sub tvwClass_LostFocus()
    If Me.cmd���� Is ActiveControl Then Exit Sub
    Me.tvwClass.Visible = False
End Sub


Private Sub txtFind_GotFocus()
    Call zlControl.TxtSelAll(txtFind)
    
End Sub

Private Sub txtFind_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyEscape Then
        picDept.Visible = False
        txtFind.Text = ""
    End If
End Sub

Private Sub txtFind_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then Exit Sub
    Call cmdFind_Click
End Sub

Private Sub txtLocate_GotFocus()
    zlControl.TxtSelAll txtLocate
End Sub

Private Sub txtLocate_KeyPress(KeyAscii As Integer)
    Dim i As Long, lngStart As Long, lngCol As Long
    Dim strFind As String, BlnFind As Boolean
    Const colִ�п��� = 3, col���˿��� = 1
    
    If KeyAscii = vbKeyReturn Then
        If txtLocate.Tag <> txtLocate.Text Then
            lblLocate.Tag = ""
            txtLocate.Tag = txtLocate.Text
        End If
        
        lngStart = Val("" & lblLocate.Tag) + 1
        If lngStart > msf����ִ��.Rows - 1 Then
            MsgBox "�Ѿ����ҵ������һ�����ҡ�", vbInformation, Me.Caption
            lblLocate.Tag = 0
            Exit Sub
        End If
        strFind = IIf(gstrMatch <> "", "*", "") & UCase(txtLocate.Text) & "*"
        
        If optDeptKind(0).Value Then
            lngCol = colִ�п���
        Else
            lngCol = col���˿���
        End If
        
        For i = lngStart To msf����ִ��.Rows - 1
            If msf����ִ��.TextMatrix(i, lngCol) Like strFind Or zlCommFun.SpellCode(msf����ִ��.TextMatrix(i, lngCol)) Like strFind Then
                lblLocate.Tag = i
                msf����ִ��.Select i, lngCol
                msf����ִ��.ShowCell i, lngCol
                If msf����ִ��.Visible Then msf����ִ��.SetFocus
                BlnFind = True
                Exit For
            End If
        Next
        If Not BlnFind Then
            If Val(lblLocate.Tag & "") = 0 Then
                MsgBox "û���ҵ������ҵĿ��ҡ�", vbInformation, Me.Caption
            Else
                MsgBox "�Ѿ����ҵ������һ�����ҡ�", vbInformation, Me.Caption
                lblLocate.Tag = 0
            End If
        End If
    End If
End Sub

Private Sub txtML_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtML_LostFocus()
    Me.txtML.Text = FormatEx(Val(Me.txtML.Text), 5)
End Sub

Private Sub txt�걾��λ_GotFocus()
    Me.txt�걾��λ.SelStart = 0: Me.txt�걾��λ.SelLength = 100
End Sub

Private Sub txt�걾��λ_KeyPress(KeyAscii As Integer)
    Dim rsTmp As New ADODB.Recordset
    Dim strSql As String
    Dim vPoint As POINTAPI
    Dim strName As String
    
    On Error GoTo ErrHandle
    If KeyAscii = vbKeyReturn Then
        strName = Trim(Me.txt�걾��λ.Text)
        If strName = "" Then Exit Sub

        strSql = "Select Rownum As ID, ����, ����, ���� From ���Ƽ���걾 " & _
               " Where (���� Like '" & strName & "%'" & _
               " Or ���� Like '" & mstrMatch & strName & "%'" & _
               " Or ���� Like '" & mstrMatch & UCase(strName) & "%')"

        vPoint = zlControl.GetCoordPos(txt�걾��λ.hWnd, txt�걾��λ.Left - 165, txt�걾��λ.Top - 30)

        Set rsTmp = zlDatabase.ShowSelect(Me, strSql, 0, "�걾��λ", , , , , True, True, vPoint.x, vPoint.y)

        If rsTmp.State = 0 Then
            Me.txt�걾��λ.Text = ""
            If Trim(Me.txt�걾��λ.Text) = "" And Trim(Me.txt�걾��λ.Tag) <> "" Then
                Me.txt�걾��λ.Text = Trim(Me.txt�걾��λ.Tag)
            End If
            Exit Sub
        End If
        If Not rsTmp Is Nothing Then
            Me.txt�걾��λ.Text = rsTmp("����")
            Me.txt�걾��λ.Tag = rsTmp("����")
        Else
            If Trim(Me.txt�걾��λ.Text) = "" And Trim(Me.txt�걾��λ.Tag) <> "" Then
                Me.txt�걾��λ.Text = Trim(Me.txt�걾��λ.Tag)
            Else
                Me.txt�걾��λ.Text = ""
                Me.txt�걾��λ.SetFocus
            End If
        End If

    End If
    If InStr(" ~!@#$%^&|=`;'""?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub txt�걾��λ_Validate(Cancel As Boolean)
    txt�걾��λ.Text = txt�걾��λ.Tag
End Sub

Private Sub txt��׼����_GotFocus()
    Me.txt��׼����.SelStart = 0: Me.txt��׼����.SelLength = 100
End Sub

Private Sub txt��׼����_KeyPress(KeyAscii As Integer)
    If KeyAscii <> vbKeyReturn Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        Exit Sub
    End If
    If Trim(Me.txt��׼����.Text) = "" Then
        Me.txt��׼����.Tag = ""
        Me.txt��׼����.Text = ""
        Me.lbl��׼����.Caption = ""
        Me.stbInfo.Tab = 1: Me.chk�������(0).SetFocus
        Exit Sub
    End If

    err = 0: On Error GoTo ErrHand

    
    gstrSql = "select A.ID,A.����,A.�������� ��������,A.����,A.����" & _
            " from ��������Ŀ¼ A" & _
            " where A.���='S'" & _
            "   and (A.���� like [1] " & _
            "       OR A.���� like [2] " & _
            "       OR A.���� like [2]) and (A.����ʱ�� is null or A.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, Trim(Me.txt��׼����.Text) & "%", gstrMatch & Trim(Me.txt��׼����.Text) & "%")

    With rsTemp
        If .RecordCount = 0 Then
            MsgBox "δ�ҵ�ָ��������׼����", vbExclamation, gstrSysName
            Me.txt��׼����.SetFocus
            Exit Sub
        End If
        If .RecordCount = 1 Then
            Me.txt��׼����.Tag = !ID
            Me.txt��׼����.Text = IIf(IsNull(!����), "", !����)
            Me.lbl��׼����.Caption = IIf(IsNull(!��������), "", "��" & NVL(!��������) & "��") & IIf(IsNull(!����), "", !����)
            Me.stbInfo.Tab = 1: Me.chk�������(0).SetFocus
            Exit Sub
        End If

        Me.lvwItem.ListItems.Clear
        Do While Not .EOF
            Set objItem = Me.lvwItem.ListItems.Add(, "_" & !ID, !����, "expend", "expend")
            objItem.SubItems(Me.lvwItem.ColumnHeaders("����").Index - 1) = !����
            objItem.SubItems(Me.lvwItem.ColumnHeaders("���").Index - 1) = NVL(!��������)
            .MoveNext
        Loop
        With Me.lvwItem
            .ListItems(1).Selected = True
            .Tag = "����"
            .Left = Me.stbInfo.Left + Me.fra��׼����.Left + Me.fra��׼����.Width - .Width
            .Top = Me.stbInfo.Top + Me.fra��׼����.Top + Me.txt��׼����.Top + Me.txt��׼����.Height
            .ZOrder 0: .Visible = True
            .SetFocus
        End With
    End With
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txt����ƴ��_GotFocus()
    Me.txt����ƴ��.SelStart = 0: Me.txt����ƴ��.SelLength = 100
End Sub

Private Sub txt����ƴ��_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt�������_GotFocus()
    Me.txt�������.SelStart = 0: Me.txt�������.SelLength = 100
End Sub

Private Sub txt�������_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt�ο�_GotFocus()
    Me.txt�ο�.SelStart = 0: Me.txt�ο�.SelLength = 100
End Sub

Private Sub txt�ο�_KeyPress(KeyAscii As Integer)
    Dim rsTmp As ADODB.Recordset

    If KeyAscii = vbKeyReturn Then
        If Me.txt�ο� <> strRefer Then
            Set rsTmp = SelectRefer(Trim(Me.txt�ο�))
            If rsTmp Is Nothing Then
                Me.txt�ο� = strRefer
                MsgBox "û���ҵ��ɲο�����Ŀ��", vbInformation, Me.Caption
            Else
                Me.txt�ο� = rsTmp("����"): Me.txt�ο�.Tag = rsTmp("ID"): strRefer = Me.txt�ο�
            End If
            If Left(Me.cbo���.Text, 1) = "D" Then
                Call zlCommFun.PressKey(vbKeyTab)
            ElseIf Me.fra��׼����.Visible Then
                Call zlCommFun.PressKey(vbKeyTab)
            End If
        Else
            If Left(Me.cbo���.Text, 1) = "D" Then
                Call zlCommFun.PressKey(vbKeyTab)
            ElseIf Me.fra��׼����.Visible Then
                Call zlCommFun.PressKey(vbKeyTab)
            Else
                Me.stbInfo.Tab = 1
            End If
        End If
        Exit Sub
    End If
    If InStr(" ~!@#$%^&|=`;'""?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt�ο�_LostFocus()
    If Me.txt�ο� <> strRefer And Me.txt�ο�.Text <> "" Then
        Me.txt�ο� = strRefer
    End If
    
    If Me.txt�ο�.Text = "" Then
        Me.txt�ο�.Tag = ""
    End If
End Sub

Private Sub txt����_GotFocus()
    Me.txt����.SelStart = 0: Me.txt����.SelLength = 100
End Sub

Private Sub txt����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt���㵥λ_GotFocus()
    Me.txt���㵥λ.SelStart = 0: Me.txt���㵥λ.SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt���㵥λ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt���㵥λ_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt��鲿λ_GotFocus()
    Me.txt��鲿λ.SelStart = 0: Me.txt��鲿λ.SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt��鲿λ_KeyPress(KeyAscii As Integer)
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii = vbKeyReturn Then
        Me.stbInfo.Tab = 1: Me.chk�������(0).SetFocus
    End If
End Sub

Private Sub txt��鲿λ_LostFocus()
    Call zlCommFun.OpenIme(False)
End Sub


Private Sub txt¼������_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub


Private Sub txt¼������_LostFocus()
    Me.txt¼������.Text = FormatEx(Val(Me.txt¼������.Text), 5)
End Sub


Private Sub txt����ִ��_GotFocus()
    Me.txt����ִ��.SelStart = 0: Me.txt����ִ��.SelLength = 100
End Sub

Private Sub txt����ִ��_KeyPress(KeyAscii As Integer)
    Dim strTemp As String
    Dim rsTmp As New ADODB.Recordset
    Dim objItem As ListItem
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Trim(Me.txt����ִ��.Text) = "" Then Me.txt����ִ��.Tag = "": Me.txt����ִ��.Text = "": Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    
    strTemp = UCase(Me.txt����ִ��.Text)
    If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
    
    err = 0: On Error GoTo ErrHand
    
    gstrSql = "select distinct ID,����,����" & _
            " from ���ű� D,��������˵�� T" & _
            " where D.ID=T.����ID and T.������� in (1,2,3)" & _
            "       and (D.����ʱ�� is null or D.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))" & _
            "       and (D.���� like [1] or D.���� like [1] or D.���� like [1])" & _
            " order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, gstrMatch & strTemp & "%")
    
    With rsTmp
        If .BOF Or .EOF Then
            MsgBox "δ�ҵ�ָ�����ţ����������룡", vbExclamation, gstrSysName:  Exit Sub
        End If
        If .RecordCount = 1 Then
            Me.txt����ִ��.Tag = !ID: Me.txt����ִ��.Text = !����: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Me.lvwItems.Checkboxes = False
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !����)
            objItem.Icon = "Dept": objItem.SmallIcon = "Dept"
            objItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = !����
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    
    With Me.picDept
        .Left = Me.fraִ�в���.Left + Me.txt����ִ��.Left
        .Top = Me.fraִ�в���.Top + Me.txt����ִ��.Top + Me.txt����ִ��.Height
        
        lbl��������.Visible = False
        cboProperty.Visible = False
        ChkSelect.Visible = False
        txtFind.Visible = False
        cmdFind.Visible = False
        cmdFindOk.Visible = False
        cmdFindCancle.Visible = False
        
        .ZOrder 0: .Visible = True
    End With
    
    With Me.lvwItems
        .Tag = "����"
        .Left = 0
        .Top = 0
        .Width = Me.picDept.Width
        .Height = Me.picDept.Height
        .SetFocus
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txt����ƴ��_GotFocus()
    Me.txt����ƴ��.SelStart = 0: Me.txt����ƴ��.SelLength = 100
End Sub

Private Sub txt����ƴ��_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt�������_GotFocus()
    Me.txt�������.SelStart = 0: Me.txt�������.SelLength = 100
End Sub

Private Sub txt�������_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txtƤ�Ա�ע_KeyPress(KeyAscii As Integer)
    If InStr("();,'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub


Private Sub txtƤ������_KeyPress(KeyAscii As Integer)
    If InStr("();,'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub


Private Sub txt��������_GotFocus()
    Me.txt��������.SelStart = 0: Me.txt��������.SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt��������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txt��������.Text = MoveSpecialChar(txt��������.Text)
        Me.txt����ƴ��.Text = zlStr.GetCodeByORCL(Me.txt��������.Text, False, mlng���볤��)
        Me.txt�������.Text = zlStr.GetCodeByORCL(Me.txt��������.Text, True, mlng���볤��)
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End If
'    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt��������_LostFocus()
    Me.txt����ƴ��.Text = zlStr.GetCodeByORCL(Me.txt��������.Text, False, mlng���볤��)
    Me.txt�������.Text = zlStr.GetCodeByORCL(Me.txt��������.Text, True, mlng���볤��)
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txt��Ŀ����_GotFocus()
    Me.txt��Ŀ����.SelStart = 0: Me.txt��Ŀ����.SelLength = 100
End Sub

Private Sub txt��Ŀ����_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    Case Else
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii >= Asc("A") And KeyAscii <= Asc("Z") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt��Ŀ����_GotFocus()
    Me.txt��Ŀ����.SelStart = 0: Me.txt��Ŀ����.SelLength = 100
    Call zlCommFun.OpenIme(True)
End Sub

Private Sub txt��Ŀ����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        txt��Ŀ����.Text = MoveSpecialChar(txt��Ŀ����.Text)
        Me.txt����ƴ��.Text = zlStr.GetCodeByORCL(Me.txt��Ŀ����.Text, False, mlng���볤��)
        Me.txt�������.Text = zlStr.GetCodeByORCL(Me.txt��Ŀ����.Text, True, mlng���볤��)
        Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    End If
'    If InStr(" ~!@#$%^&*_+|=`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub txt��Ŀ����_LostFocus()
    If mbln��ϲ�λ��Ŀ And Me.txt��Ŀ����.Text <> Me.txt��Ŀ����.Tag Then
        Me.txt��Ŀ����.Text = Me.txt��Ŀ����.Tag
        Me.txt����ƴ��.Text = zlStr.GetCodeByORCL(Me.txt��Ŀ����.Text, False, mlng���볤��)
        Me.txt�������.Text = zlStr.GetCodeByORCL(Me.txt��Ŀ����.Text, True, mlng���볤��)
        MsgBox "����Ŀ�Ǽ������еĲ�λ��Ŀ�������޸����ơ�"
        Exit Sub
    End If
    Me.txt����ƴ��.Text = zlStr.GetCodeByORCL(Me.txt��Ŀ����.Text, False, mlng���볤��)
    Me.txt�������.Text = zlStr.GetCodeByORCL(Me.txt��Ŀ����.Text, True, mlng���볤��)
    Call zlCommFun.OpenIme(False)
End Sub
Private Sub txtסԺִ��_GotFocus()
    Me.txtסԺִ��.SelStart = 0: Me.txtסԺִ��.SelLength = 100
End Sub

Private Sub txtסԺִ��_KeyPress(KeyAscii As Integer)
    Dim objItem As ListItem
    Dim strTemp As String
    Dim rsTmp As New ADODB.Recordset
    
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Trim(Me.txtסԺִ��.Text) = "" Then Me.txtסԺִ��.Tag = "": Me.txtסԺִ��.Text = "": Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    
    strTemp = UCase(Me.txtסԺִ��.Text)
    If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
    
    err = 0: On Error GoTo ErrHand
   
    gstrSql = "select distinct ID,����,����" & _
            " from ���ű� D,��������˵�� T" & _
            " where D.ID=T.����ID and T.������� in (1,2,3)" & _
            "       and (D.����ʱ�� is null or D.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))" & _
            "       and (D.���� like [1] or D.���� like [1] or D.���� like [1])" & _
            " order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, gstrMatch & strTemp & "%")
        
    With rsTmp
        If .BOF Or .EOF Then
            MsgBox "δ�ҵ�ָ�����ţ����������룡", vbExclamation, gstrSysName: Exit Sub
        End If
        If .RecordCount = 1 Then
            Me.txtסԺִ��.Tag = !ID: Me.txtסԺִ��.Text = !����: Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Me.lvwItems.Checkboxes = False
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !����)
            objItem.Icon = "Dept": objItem.SmallIcon = "Dept"
            objItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = !����
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    
    With Me.picDept
        .Left = Me.fraִ�в���.Left + Me.txtסԺִ��.Left
        .Top = Me.fraִ�в���.Top + Me.txtסԺִ��.Top + Me.txtסԺִ��.Height
        
        lbl��������.Visible = False
        cboProperty.Visible = False
        ChkSelect.Visible = False
        txtFind.Visible = False
        cmdFind.Visible = False
        cmdFindOk.Visible = False
        cmdFindCancle.Visible = False
        
        .ZOrder 0: .Visible = True
    End With
    
    With Me.lvwItems
        .Tag = "סԺ"
        .Left = 0
        .Top = 0
        .Width = Me.picDept.Width
        .Height = Me.picDept.Height
        .SetFocus
    End With
    
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

'�ж��Ƿ�Ϊ�༭��
Private Function ifEditKey(ByVal KeyAscii As Integer, Optional ByVal AllowSubtract As Boolean = True) As Boolean
    If KeyAscii = vbKeyBack Or (KeyAscii = vbKeyInsert And AllowSubtract) Or KeyAscii = vbKeyDelete Or _
      KeyAscii = vbKeyHome Or KeyAscii = vbKeyEnd Or KeyAscii = vbKeyLeft Or KeyAscii = vbKeyRight Or _
      KeyAscii = vbKeyEscape Or KeyAscii = vbKeyReturn Then
        ifEditKey = True
    Else
        ifEditKey = False
    End If
End Function

Private Function TransFormula(ByVal strFormula As String, strErrorMsg As String, Optional iErrorPos As Integer = 1) As String
'��������Ŀ���㹫ʽת��Ϊ��ID��ʶ�Ĺ�ʽ

    Dim i As Integer, strTmp As String, iElementStart As Integer, strCalcForm As String
    Dim strElement As String, strSql As String, rsTmp As New ADODB.Recordset

    On Error GoTo DBError
    strErrorMsg = "": iErrorPos = 1: TransFormula = ""
    strCalcForm = ""
    For i = 1 To Len(strFormula)
        strTmp = Mid(strFormula, i, 1)
        If iElementStart > 0 Then
            '���ҵ�Ԫ�صĿ�ʼλ��
            If strTmp = "]" Then
                strElement = Trim(Mid(strFormula, iElementStart + 1, i - iElementStart - 1))
                strSql = "Select ������ĿID,nvl(��Ŀ���,1),nvl(�������,1) From ������Ŀ" & _
                    " Where ��д=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, UCase(strElement))

                If rsTmp.EOF Then
                    strErrorMsg = "���㹫ʽδ�ҵ�������Ŀ��" & strElement & "��"
                    iErrorPos = iElementStart + 1
                    TransFormula = ""
                    Exit Function
                End If
                If rsTmp(1) <> 1 Then
                    strErrorMsg = "������Ŀ��" & strElement & " ���ǻ�����Ŀ��"
                    iErrorPos = iElementStart + 1
                    TransFormula = ""
                    Exit Function
                End If
                If rsTmp(2) <> 1 Then
                    strErrorMsg = "������Ŀ��" & strElement & " ���������ͣ�"
                    iErrorPos = iElementStart + 1
                    TransFormula = ""
                    Exit Function
                End If

                TransFormula = TransFormula & "[" & rsTmp(0) & "]"
                strCalcForm = strCalcForm & "1" '���㹫ʽ��ģ����Ϊ1
                iElementStart = 0
            End If
        Else
            If strTmp = "[" Then
                iElementStart = i
            Else
                TransFormula = TransFormula & strTmp
                strCalcForm = strCalcForm & strTmp
            End If
        End If
    Next
    'У�鹫ʽ���﷨�Ƿ���ȷ
    strSql = "Select " & strCalcForm & " From Dual"
    If rsTmp.State <> adStateClosed Then rsTmp.Close
    On Error GoTo ValidError
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption)
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    TransFormula = ""
    Call SaveErrLog
    Exit Function
ValidError:
    If gcnOracle.Errors(0).NativeError <> 1476 Then '���Գ���Ϊ0
        strErrorMsg = "���㹫ʽ�﷨��" & Mid(err.Description, InStr(err.Description, ":") + 1)
        iErrorPos = 1
        TransFormula = ""
    End If
End Function

Private Function TransFormula1(ByVal strFormula As String) As String
'����ID��ʶ�Ĺ�ʽת��Ϊ����д��ʶ�Ĺ�ʽ

    Dim i As Integer, strTmp As String, iElementStart As Integer
    Dim strElement As String, strSql As String, rsTmp As New ADODB.Recordset

    On Error GoTo DBError
    TransFormula1 = ""
    For i = 1 To Len(strFormula)
        strTmp = Mid(strFormula, i, 1)
        If iElementStart > 0 Then
            '���ҵ�Ԫ�صĿ�ʼλ��
            If strTmp = "]" Then
                strElement = Trim(Mid(strFormula, iElementStart + 1, i - iElementStart - 1))
                strSql = "Select ��д From ������Ŀ" & _
                    " Where ������ĿID=[1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, strElement)

                If rsTmp.EOF Then
                    TransFormula1 = TransFormula1 & "[δ֪��Ŀ]"
                Else
                    TransFormula1 = TransFormula1 & "[" & UCase(NVL(rsTmp(0))) & "]"
                End If

                iElementStart = 0
            End If
        Else
            If strTmp = "[" Then
                iElementStart = i
            Else
                TransFormula1 = TransFormula1 & strTmp
            End If
        End If
    Next
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    TransFormula1 = ""
    Call SaveErrLog
End Function

Private Function CalcExist(ByVal lngItemID As Long) As String
'�ж�ָ������Ŀ�Ƿ�����������Ŀ���ã��������õ���Ŀ����
    Dim strSql As String, rsTmp As New ADODB.Recordset

    On Error GoTo DBError
    CalcExist = ""
    strSql = "Select a.������,b.��д From ����������Ŀ a,������Ŀ b" & _
        " Where a.id=b.������Ŀid And b.��Ŀ���=3 And b.���㹫ʽ Like [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption, "%[" & lngItemID & "]%")

    If Not rsTmp.EOF Then CalcExist = rsTmp(0)
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub initVfgList()
    Dim strSql As String
    Dim rsTmp As ADODB.Recordset, rsZlxm As ADODB.Recordset
    Dim strTemp As String
    Dim strItem As String
    Dim strName As String
    Dim varTemp As Variant
    Dim i As Long
    
    On Error GoTo ErrHandle
    With vfgList
        '��ʼ�����
        .Clear
        .FixedCols = 0: .FixedRows = 1
        .Rows = 1: .Cols = 7
        
        .MergeRow(0) = True
        .MergeCellsFixed = flexMergeRestrictColumns
        
        .MergeCol(0) = True: .MergeCol(1) = True
        .MergeCells = flexMergeRestrictColumns
        
        .RowHeightMin = 300
        .ColWidthMin = 450
        If cbo���.Text = "D-���" And cbo��������.Text = "18-����" Then
            .TextMatrix(0, 0) = "�걾����": .TextMatrix(0, 1) = "�걾����": .TextMatrix(0, 2) = "�������": .TextMatrix(0, 3) = "��ע"
        Else
            .TextMatrix(0, 0) = "��λ": .TextMatrix(0, 1) = "��λ": .TextMatrix(0, 2) = "����": .TextMatrix(0, 3) = "��ע"
        End If
        .TextMatrix(0, 4) = "Ĭ��": .TextMatrix(0, 5) = "ʹ��": .TextMatrix(0, 6) = "Ψһ��"
        .ColKey(0) = "����": .ColKey(1) = "����": .ColKey(2) = "����": .ColKey(3) = "��ע"
        .ColKey(4) = "Ĭ��": .ColKey(5) = "ʹ��": .ColKey(6) = "Ψһ��"
        
        .ColHidden(.ColIndex("����")) = False: .ColHidden(.ColIndex("����")) = False
        .ColHidden(.ColIndex("����")) = False: .ColHidden(.ColIndex("��ע")) = False
        .ColHidden(.ColIndex("Ĭ��")) = True: .ColHidden(.ColIndex("ʹ��")) = True
        .ColHidden(.ColIndex("Ψһ��")) = True
        
        .ColWidth(.ColIndex("����")) = 950: .ColWidth(.ColIndex("����")) = 450: .ColWidth(.ColIndex("����")) = 5000
        .ColWidth(.ColIndex("��ע")) = 1800: .ColWidth(.ColIndex("Ĭ��")) = 0: .ColWidth(.ColIndex("ʹ��")) = 0
        .ColWidth(.ColIndex("Ψһ��")) = 0
        
        .Cell(flexcpAlignment, 0, 0, 0, .Cols - 1) = flexAlignCenterCenter
        .WordWrap = True
        .AutoResize = True
        
        
        .Editable = flexEDKbdMouse
        
        '��ȡ����,�����
        '       2007-07-09 1.�����Ĳ�λ����; 2.�޷����Ĳ�λ����ʹ�á�
        strSql = "Select a.����, a.����, a.����, a.����, a.��ע" & vbNewLine & _
                "From ���Ƽ�鲿λ a " & vbNewLine & _
                "Where a.���� Is Not Null And a.����=[1] Order by a.����, a.����"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSql, Me.Caption & " ��ȡ��鲿λ", Mid(cbo��������.Text, InStr(1, cbo��������.Text, "-") + 1))
        
        '���������õķ���
        strSql = "Select A.��Ŀid, A.����, A.��λ, A.����, A.Ĭ��, Decode(Nvl(B.�շ���Ŀid, 0), 0, 0, 1) As ʹ��,A.�ϼ����� " & vbNewLine & _
                "From ������Ŀ��λ A, �����շѹ�ϵ B" & vbNewLine & _
                "Where A.��λ = B.��鲿λ(+) And A.���� = B.��鷽��(+) And A.��Ŀid = B.������Ŀid(+)" & vbNewLine & _
                " And instr([2],A.����)>0 And A.��ĿID=[1] order by id"

        Set rsZlxm = zlDatabase.OpenSQLRecord(strSql, Me.Caption & " ������Ŀ��λ", lngItemID, cbo��������.Text)
        
        Do Until rsTmp.EOF
            .Rows = .Rows + 1
            
            .TextMatrix(.Rows - 1, .ColIndex("����")) = "" & rsTmp.Fields("����")
            .TextMatrix(.Rows - 1, .ColIndex("����")) = rsTmp.Fields("����")
            .RowData(.Rows - 1) = 0
            Set .Cell(flexcpPicture, .Rows - 1, .ColIndex("����")) = imgList.ListImages(4).Picture
            
            .Cell(flexcpData, .Rows - 1, .ColIndex("����")) = "" & rsTmp.Fields("����")
            .TextMatrix(.Rows - 1, .ColIndex("����")) = Anslyze_MethodString("" & rsTmp.Fields("����"))
            
            If InStr(2, rsTmp.Fields("����"), vbTab) = 0 And InStr(2, rsTmp.Fields("����"), ";") = 0 Then
                .TextMatrix(.Rows - 1, .ColIndex("Ψһ��")) = "1"
            End If
            
            strName = ""
            rsZlxm.Filter = ""
            rsZlxm.Filter = " ��λ='" & "" & rsTmp.Fields("����") & "' and Ĭ��=1"
            .TextMatrix(.Rows - 1, .ColIndex("Ĭ��")) = ""
            varTemp = Split(.TextMatrix(.Rows - 1, .ColIndex("����")), "  ")
            If rsZlxm.RecordCount > 0 Then
                For i = 0 To UBound(varTemp)
                    strItem = varTemp(i)
                    rsZlxm.MoveFirst
                    Do Until rsZlxm.EOF
                        If InStr(varTemp(i), rsZlxm.Fields("����")) > 0 And varTemp(i) <> "" Then
                            If InStr(varTemp(i), "��") > 0 Then
                                strTemp = Mid(varTemp(i), 1, InStr(varTemp(i), "��") - 1)
                                If "" & rsZlxm.Fields("�ϼ�����") <> "" Then
                                    If InStr(strTemp, rsZlxm.Fields("�ϼ�����")) > 0 Then
                                        strItem = Replace(strItem, "��" & rsZlxm.Fields("����"), "��" & rsZlxm.Fields("����"))
                                    End If
                                Else
                                    If InStr(varTemp(i), "" & rsZlxm.Fields("����")) > 0 Then
                                        strItem = Replace(strItem, "��" & rsZlxm.Fields("����"), "��" & rsZlxm.Fields("����"))
                                        strItem = Replace(strItem, "��" & rsZlxm.Fields("����"), "��" & rsZlxm.Fields("����"))
                                    End If
                                End If
                            Else
                                If InStr(varTemp(i), "" & rsZlxm.Fields("����")) > 0 Then
                                    strItem = Replace(strItem, "��" & rsZlxm.Fields("����"), "��" & rsZlxm.Fields("����"))
                                    strItem = Replace(strItem, "��" & rsZlxm.Fields("����"), "��" & rsZlxm.Fields("����"))
                                End If
                            End If
                        End If
                        rsZlxm.MoveNext
                    Loop
                    If strItem <> "" Then strName = strName & "  " & strItem
                Next
                .TextMatrix(.Rows - 1, .ColIndex("����")) = strName
                rsZlxm.MoveFirst
                Do While Not rsZlxm.EOF
                    If rsZlxm.Fields("�ϼ�����") <> "" Then
                        .TextMatrix(.Rows - 1, .ColIndex("Ĭ��")) = .TextMatrix(.Rows - 1, .ColIndex("Ĭ��")) & "" & rsZlxm.Fields("�ϼ�����") & "����" & rsZlxm.Fields("����") & ","
                    Else
                        .TextMatrix(.Rows - 1, .ColIndex("Ĭ��")) = .TextMatrix(.Rows - 1, .ColIndex("Ĭ��")) & "" & rsZlxm.Fields("����") & ","
                    End If
                    rsZlxm.MoveNext
                Loop
            End If
            
            If InStr(.TextMatrix(.Rows - 1, .ColIndex("Ĭ��")), ",") > 0 Then
                .TextMatrix(.Rows - 1, .ColIndex("Ĭ��")) = Mid(.TextMatrix(.Rows - 1, .ColIndex("Ĭ��")), 1, Len(.TextMatrix(.Rows - 1, .ColIndex("Ĭ��"))) - 1)
            End If
            
            rsZlxm.Filter = ""
            rsZlxm.Filter = " ��λ='" & "" & rsTmp.Fields("����") & "'"
            If Not rsZlxm.EOF Then
                If .RowData(.Rows - 1) = 0 Then
                    .RowData(.Rows - 1) = 1
                    Set .Cell(flexcpPicture, .Rows - 1, .ColIndex("����")) = imgList.ListImages(5).Picture
                End If
                .TextMatrix(.Rows - 1, .ColIndex("ʹ��")) = rsZlxm.Fields("ʹ��")
            Else
                If optList(1).Value Then
                    .RowHidden(.Rows - 1) = True
                End If
            End If
            .TextMatrix(.Rows - 1, .ColIndex("��ע")) = "" & rsTmp.Fields("��ע")
            
            rsTmp.MoveNext
        Loop
        .AutoSizeMode = flexAutoSizeColWidth
        .AutoSize 1, .ColIndex("����")
        .AutoSize 3, .ColIndex("��ע")
        .AutoSizeMode = flexAutoSizeRowHeight
        .AutoSize 0, .ColIndex("����")
        .AutoSize 2, .ColIndex("����")
        
        .RowHeight(0) = 350
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Function Anslyze_MethodString(ByVal strMethod As String) As String
    '������鷽����
    '   strMethod:������
    '   ����,��ʽ���õķ�����
    Dim aryItem() As String, strItems As String, strTemp As String
    Dim aryChild() As String, lngChild As Long, lngCount As Long
    
    strItems = ""
    strMethod = Replace(strMethod, vbTab, ";" & vbTab)
    
    aryItem() = Split(strMethod, ";")
    For lngCount = 0 To UBound(aryItem)
        strTemp = aryItem(lngCount)
        If strTemp <> "" Then
            If InStr(strTemp, vbTab) >= 1 Then
                strTemp = Mid(aryItem(lngCount), 3)
                If InStr(1, strTemp, ",") > 0 Then
                    aryChild = Split(strTemp, ",")
                    strTemp = ""
                    For lngChild = 1 To UBound(aryChild)
                        strTemp = strTemp & " ��" & Mid(aryChild(lngChild), 2)
                    Next
                    strTemp = aryChild(0) & "��" & Trim(strTemp) & "��"
                End If
                strItems = strItems & "  ��" & strTemp '�������ո񣬷�������ȡ
            Else
                strTemp = Mid(aryItem(lngCount), 2)
                If InStr(1, strTemp, ",") > 0 Then
                    aryChild = Split(strTemp, ",")
                    strTemp = ""
                    For lngChild = 1 To UBound(aryChild)
                        strTemp = strTemp & " ��" & Mid(aryChild(lngChild), 2)
                    Next
                    strTemp = aryChild(0) & "��" & Trim(strTemp) & "��"
                End If
                strItems = strItems & "  ��" & strTemp '�������ո񣬷�������ȡ
            End If
        End If
    Next
    
    Anslyze_MethodString = strItems
End Function

Private Sub vfgList_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col <> vfgList.ColIndex("����") Then
        Cancel = True
    Else
        If vfgList.RowData(Row) = 1 Then
            vfgList.ColComboList(vfgList.ColIndex("����")) = "..."
        Else
            vfgList.ColComboList(vfgList.ColIndex("����")) = ""
            Cancel = True
        End If
    End If
End Sub

Private Sub vfgList_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim pt As POINTAPI, strDefault As String
    Dim arrItem() As String, lngCount As Long
    Dim strTemp As String, strItem As String
    Dim varTemp As Variant
    Dim i As Long
    
    pt.x = vfgList.ColPos(Col) \ Screen.TwipsPerPixelX
    pt.y = (vfgList.RowPos(Row) + vfgList.RowHeight(Row)) \ Screen.TwipsPerPixelY
    ClientToScreen vfgList.hWnd, pt
    
    If InStr(vfgList.Cell(flexcpText, 0, Col), "����") > 0 Then
        If vfgList.RowData(Row) = 1 Then
            With frmClinicDefaultModus
                 strDefault = vfgList.TextMatrix(Row, vfgList.ColIndex("Ĭ��"))
                .Move pt.x * Screen.TwipsPerPixelX, pt.y * Screen.TwipsPerPixelY
                Call .ShowModus(vfgList.Cell(flexcpData, Row, Col), strDefault)
            End With
            With vfgList
                If strDefault <> .TextMatrix(Row, .ColIndex("Ĭ��")) Then
                    '������ʾ
                    If .TextMatrix(Row, .ColIndex("Ĭ��")) <> "" Then
                        .TextMatrix(Row, .ColIndex("����")) = Replace(.TextMatrix(Row, .ColIndex("����")), "��", "��")
                        .TextMatrix(Row, .ColIndex("����")) = Replace(.TextMatrix(Row, .ColIndex("����")), "��", "��")
                    End If
                    If strDefault <> "" Then
                        arrItem = Split(strDefault, ",")
                        For lngCount = LBound(arrItem) To UBound(arrItem)
                            If InStr(arrItem(lngCount), "����") > 0 Then
                                varTemp = Split(.TextMatrix(Row, .ColIndex("����")), "  ")
                                For i = 0 To UBound(varTemp)
                                    strItem = ""
                                    If InStr(varTemp(i), Split(arrItem(lngCount), "����")(0)) > 0 And InStr(varTemp(i), Split(arrItem(lngCount), "����")(1)) > 0 Then
                                        strItem = Replace(varTemp(i), "��" & Split(arrItem(lngCount), "����")(1), "��" & Split(arrItem(lngCount), "����")(1))
                                    End If
                                    If strItem <> "" Then .TextMatrix(Row, .ColIndex("����")) = Replace(.TextMatrix(Row, .ColIndex("����")), varTemp(i), strItem)
                                Next
                            Else
                                .TextMatrix(Row, .ColIndex("����")) = Replace(.TextMatrix(Row, .ColIndex("����")) & " ", "��" & arrItem(lngCount) & " ", "��" & arrItem(lngCount) & " ")
                                .TextMatrix(Row, .ColIndex("����")) = Replace(.TextMatrix(Row, .ColIndex("����")) & " ", "��" & arrItem(lngCount) & "��", "��" & arrItem(lngCount) & "��")
                                .TextMatrix(Row, .ColIndex("����")) = Replace(.TextMatrix(Row, .ColIndex("����")) & " ", "��" & arrItem(lngCount) & "��", "��" & arrItem(lngCount) & "��")
                                .TextMatrix(Row, .ColIndex("����")) = Replace(.TextMatrix(Row, .ColIndex("����")) & " ", "��" & arrItem(lngCount) & " ", "��" & arrItem(lngCount) & " ")
                                .TextMatrix(Row, .ColIndex("����")) = Replace(.TextMatrix(Row, .ColIndex("����")) & " ", "��" & arrItem(lngCount) & "��", "��" & arrItem(lngCount) & "��")
                            End If
                        Next
                    End If
                    .TextMatrix(Row, .ColIndex("Ĭ��")) = strDefault
                End If
            End With
        End If
    End If
End Sub

Private Sub vfgList_EnterCell()
    With vfgList
        If (.Col = .ColIndex("����") Or .Col = .ColIndex("����")) And .Row > 0 Then
            On Error Resume Next
            Call .CellBorder(.GridColor, 1, 1, 2, 2, 0, 0)
        End If
    End With
End Sub

Private Sub vfgList_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        
        If vfgList.Col = vfgList.ColIndex("����") Then
            Call SwapPic(vfgList.Row, vfgList.Col)
        End If
    ElseIf KeyCode = vbKeyReturn Then
        With vfgList
            If .Col = .ColIndex("����") Then
                If .Row >= .FixedCols And .Row + 1 <= .Rows - 1 Then
                    .Select .Row + 1, .ColIndex("����")
                End If
            ElseIf .Col < .Cols - 1 Then
                
                .Select .Row, .Col + 1
            
            End If
        End With
        
    End If
End Sub


Private Sub SwapPic(ByVal lngRow As Long, ByVal lngCol As Long)
    
    If Not cbo��������.Enabled Then Exit Sub
    With vfgList
        If .Col = .ColIndex("����") And .Row > 0 And .Row < .Rows Then
            lngRow = .Row
            lngCol = .Col
            If .RowData(lngRow) = 0 Then
                .RowData(lngRow) = 1
                Set .Cell(flexcpPicture, lngRow, lngCol) = imgList.ListImages(5).Picture
                If .TextMatrix(lngRow, .ColIndex("Ψһ��")) = "1" Then
                    .TextMatrix(lngRow, .ColIndex("����")) = Replace(.TextMatrix(lngRow, .ColIndex("����")), "��", "��")
                    .TextMatrix(lngRow, .ColIndex("����")) = Replace(.TextMatrix(lngRow, .ColIndex("����")), "��", "��")
                    .TextMatrix(lngRow, .ColIndex("Ĭ��")) = Replace(.TextMatrix(lngRow, .ColIndex("����")), "��", "")
                    .TextMatrix(lngRow, .ColIndex("Ĭ��")) = Replace(.TextMatrix(lngRow, .ColIndex("Ĭ��")), "��", "")
                End If
            Else
                If Val(.TextMatrix(lngRow, .ColIndex("ʹ��"))) = 0 Then
                    .RowData(lngRow) = 0
                    Set .Cell(flexcpPicture, lngRow, lngCol) = imgList.ListImages(4).Picture
                    If .TextMatrix(lngRow, .ColIndex("Ψһ��")) = "1" Then
                        .TextMatrix(lngRow, .ColIndex("����")) = Replace(.TextMatrix(lngRow, .ColIndex("����")), "��", "��")
                        .TextMatrix(lngRow, .ColIndex("����")) = Replace(.TextMatrix(lngRow, .ColIndex("����")), "��", "��")
                        .TextMatrix(lngRow, .ColIndex("Ĭ��")) = ""
                    End If
                Else
                    MsgBox "�ò�λ��ʹ�ã�����ȡ����", vbInformation, gstrSysName
                End If
            End If
        End If
    End With
End Sub

Private Sub vfgList_LeaveCell()
    With vfgList
        If (.Col = .ColIndex("����") Or .Col = .ColIndex("����")) And .Row > 0 Then
            On Error Resume Next
            Call .CellBorder(.GridColor, 0, 0, 0, 0, 0, 0)
        End If
    End With
End Sub

Private Sub vfgList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
        If vfgList.Col = vfgList.ColIndex("����") And Button = 1 And x > vfgList.CellLeft And x < vfgList.CellLeft + 250 Then
            Call SwapPic(vfgList.Row, vfgList.Col)
        End If
    
End Sub

Private Sub txtӢ��_GotFocus()
    Me.txtӢ��.SelStart = 0: Me.txtӢ��.SelLength = 1000
    Call zlCommFun.OpenIme(False)
End Sub

Private Sub txtӢ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub


Private Sub vsfBloodLis_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    mstrOldBlood = vsfBloodLis.TextMatrix(vsfBloodLis.Row, vsfBloodLis.Col)
End Sub

Private Sub vsfBloodLis_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTemp As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim strPro As String
    Dim strSql As String
    Dim intAttr As Integer
    Dim strSQLItem As String
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim intRow As Integer
    Dim str����id As String
    
    vRect = zlControl.GetControlRect(vsfBloodLis.hWnd) '��ȡλ��
    dblLeft = vRect.Left + vsfBloodLis.CellLeft
    dblTop = vRect.Top + vsfBloodLis.CellTop + vsfBloodLis.CellHeight + 3200
    
    With vsfBloodLis
        gstrSql = "Select ID, ����id, ����, ���� From ������ĿĿ¼ Where ��� = 'C'  Order By ID"
        Set rsTemp = zlDatabase.ShowSelect(Me, gstrSql, 0, "������Ŀ", False, "", "", False, False, _
                True, dblLeft, dblTop, vsfBloodLis.Height, blnCancel, False, True)
        
        If Not rsTemp Is Nothing Then
            For intRow = 1 To .Rows - 1
                If .TextMatrix(intRow, 0) <> "" Then
                    str����id = str����id & "," & .TextMatrix(intRow, 0)
                End If
            Next
            If InStr(1, "," & str����id & ",", "," & rsTemp!ID & ",") > 0 Then
                MsgBox "�Ѿ��иü�����Ŀ�ˣ�����Ҫ����ӣ�", vbInformation, Me.Caption
            Else
                .TextMatrix(Row, 0) = rsTemp!ID
                .TextMatrix(Row, 1) = rsTemp!����
            End If
        Else
            MsgBox "û���ҵ���ѡ��ļ�����Ŀ��", vbInformation, Me.Caption
        End If
    End With
End Sub

Private Sub GetBloodLis(ByVal strInput As String)
    '�ֶ�����ʱ����ȡ��Ѫ���ձ�ļ��������Ŀ
    Dim rsTemp As ADODB.Recordset
    Dim vRect As RECT, blnCancel As Boolean
    Dim strPro As String
    Dim strSql As String
    Dim intAttr As Integer
    Dim strSQLItem As String
    Dim dblLeft As Double
    Dim dblTop As Double
    Dim intRow As Integer
    Dim intCol As Integer
    Dim str����id As String
    
    vRect = zlControl.GetControlRect(vsfBloodLis.hWnd) '��ȡλ��
    dblLeft = vRect.Left + vsfBloodLis.CellLeft
    dblTop = vRect.Top + vsfBloodLis.CellTop + vsfBloodLis.CellHeight + 3200
        
    strInput = UCase(mstrFindStyle & strInput & "%")
    With vsfBloodLis
        gstrSql = "Select distinct a.Id, a.����id, a.����, a.����" & vbNewLine & _
            "From ������ĿĿ¼ A, ������Ŀ���� B" & vbNewLine & _
            "Where a.Id = b.������Ŀid And a.��� = 'C' And (b.���� Like [1] Or b.���� Like [1] or a.���� Like [1])"
        Set rsTemp = zlDatabase.ShowSQLSelect(Me, gstrSql, 0, "������Ŀ", False, "", "", False, False, _
                True, dblLeft, dblTop, vsfBloodLis.Height, blnCancel, False, True, strInput)
        
        If Not rsTemp Is Nothing Then
            For intRow = 1 To .Rows - 1
                If .TextMatrix(intRow, 0) <> "" Then
                    str����id = str����id & "," & .TextMatrix(intRow, 0)
                End If
            Next
            If InStr(1, "," & str����id & ",", "," & rsTemp!ID & ",") > 0 Then
                MsgBox "�Ѿ��иü�����Ŀ�ˣ�����Ҫ����ӣ�", vbInformation, Me.Caption
                If .TextMatrix(.Row, 0) = rsTemp!ID Then
                    .EditText = mstrOldBlood
                    .TextMatrix(.Row, .Col) = mstrOldBlood
                Else
                    If .TextMatrix(.Row, 0) <> "" Then
                        .EditText = mstrOldBlood
                        .TextMatrix(.Row, .Col) = mstrOldBlood
                    Else
                        .EditText = ""
                        .TextMatrix(.Row, .Col) = ""
                    End If
                End If
            Else
                .TextMatrix(.Row, 0) = rsTemp!ID
                .TextMatrix(.Row, 1) = rsTemp!����
            End If
        Else
            MsgBox "û���ҵ���ѡ��ļ�����Ŀ��", vbInformation, Me.Caption
            If .TextMatrix(intRow, 0) = "" Then
                .EditText = ""
                .TextMatrix(.Row, .Col) = ""
            Else
                .EditText = mstrOldBlood
                .TextMatrix(.Row, .Col) = mstrOldBlood
            End If
        End If
    End With
End Sub

Private Sub vsfBloodLis_EnterCell()
    With vsfBloodLis
        .Editable = flexEDNone
    End With
End Sub

Private Sub vsfBloodLis_KeyDown(KeyCode As Integer, Shift As Integer)
    With vsfBloodLis
        If KeyCode = vbKeyReturn Then
            If .Row = .Rows - 1 Then
                If Me.Tag <> "����" Then
                    If .TextMatrix(.Row, 0) <> "" Then
                        .Rows = .Rows + 1
                        .Row = .Rows - 1
                    Else
                        KeyCode = 0
                    End If
                End If
            Else
                .Row = .Row + 1
            End If
        ElseIf KeyCode = vbKeyDelete And Me.Tag <> "����" Then
            If .Row = .Rows - 1 And .Row = 1 Then
                .TextMatrix(.Row, 0) = ""
                .TextMatrix(.Row, 1) = ""
            Else
                .RemoveItem .Row
            End If
        End If
    End With
End Sub

Private Sub vsfBloodLis_KeyPress(KeyAscii As Integer)
    With vsfBloodLis
        If KeyAscii <> vbKeyReturn And Me.Tag <> "����" Then
            .Editable = flexEDKbdMouse
        End If
    End With
End Sub

Private Sub vsfBloodLis_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsfBloodLis
        If .EditText <> "" And KeyAscii = vbKeyReturn Then
            Call GetBloodLis(.EditText)
        End If
    End With
End Sub


Private Sub vsfBloodLis_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        With vsfBloodLis
            .Editable = flexEDNone
            If Me.Tag <> "����" Then
                If .Col = 1 Then
                    .Editable = flexEDKbdMouse
                End If
            End If
        End With
    End If
End Sub


Private Sub vsfFreq_DblClick()
    With vsfFreq
        If .Rows = 1 Then Exit Sub
        If .Row = 0 Then Exit Sub
        
        If .TextMatrix(.Row, 0) = "" Then
            .TextMatrix(.Row, 0) = "��"
        Else
            .TextMatrix(.Row, 0) = ""
        End If
    End With
End Sub


Private Sub vsfFreq_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeySpace Then Exit Sub
    With vsfFreq
        If .Rows = 1 Then Exit Sub
        If .Row = 0 Then Exit Sub
        
        If .TextMatrix(.Row, 0) = "" Then
            .TextMatrix(.Row, 0) = "��"
        Else
            .TextMatrix(.Row, 0) = ""
        End If
    End With
End Sub

Private Sub vsUseDept_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    If vsUseDept.Editable = flexEDNone Then
        vsUseDept.FocusRect = flexFocusLight
        vsUseDept.ComboList = ""
    Else
        vsUseDept.FocusRect = flexFocusSolid
        vsUseDept.ComboList = "..."
    End If
End Sub

Private Sub vsUseDept_AfterUserResize(ByVal Row As Long, ByVal Col As Long)
    vsUseDept.AutoSize vsUseDept.FixedCols, vsUseDept.Cols - 1
End Sub

Private Sub vsUseDept_BeforeRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If vsUseDept.TextMatrix(OldRow, OldCol) <> vsUseDept.Cell(flexcpData, OldRow, OldCol) Then
        vsUseDept.TextMatrix(OldRow, OldCol) = vsUseDept.Cell(flexcpData, OldRow, OldCol)
    End If
End Sub

Private Sub vsUseDept_CellButtonClick(ByVal Row As Long, ByVal Col As Long)
    Dim rsTmp As New ADODB.Recordset
    Dim objItem As ListItem
    Dim i As Integer, j As Long
    
    mstr��ѡʹ�ÿ��� = ""
    With Me.vsUseDept
        For i = 0 To .Rows - 1
            For j = 0 To .Cols - 1
                If .TextMatrix(i, j) <> "" And .ColHidden(j) = False Then
                    If InStr(mstr��ѡʹ�ÿ���, .TextMatrix(i, j + 5) & "," & .TextMatrix(i, j)) = 0 Then
                        mstr��ѡʹ�ÿ��� = IIf(mstr��ѡʹ�ÿ��� = "", "", mstr��ѡʹ�ÿ��� & ";") & .TextMatrix(i, j + 5) & "," & .TextMatrix(i, j)
                    End If
                End If
            Next
        Next
    End With
    
    With Me.picDept
        .Tag = "2"
        Me.lvwItems.Tag = "ʹ��"
        .Left = Me.stbInfo.Left + Me.vsUseDept.Left + Me.vsUseDept.ColWidth(0) * Col
        .Width = 4000
        If .Left > Me.Width - .Width - stbInfo.Left - Me.vsUseDept.Left Then .Left = Me.Width - .Width - stbInfo.Left - Me.vsUseDept.Left
    
        .Top = 1050 + Row * vsUseDept.RowHeight(Row)
        .Height = 3850
        
        lbl��������.Visible = True
        cboProperty.Visible = lbl��������.Visible
        ChkSelect.Visible = lbl��������.Visible
        
        lbl��������.Left = 50
        ChkSelect.Left = .Width - ChkSelect.Width - 50
        cboProperty.Width = ChkSelect.Left - cboProperty.Left - 50
        
        cmdFind.Visible = True
        txtFind.Visible = True
        cmdFindOk.Visible = True
        cmdFindCancle.Visible = True
        .ZOrder 0
        .Visible = True
    End With

    With Me.lvwItems
        .Left = lbl��������.Left
        .Top = cboProperty.Top + cboProperty.Height + 50 + txtFind.Height + 50
        .Width = Me.picDept.Width - .Left - 50
        .Height = Me.picDept.Height - .Top - 10
        txtFind.Top = cboProperty.Top + cboProperty.Height + 50
        cmdFind.Top = cboProperty.Top + cboProperty.Height + 50
        cmdFindOk.Left = .Width + .Left - cmdFind.Width - 80 - cmdFindCancle.Width
        cmdFindCancle.Left = .Width + .Left - cmdFind.Width - 50
        cmdFindOk.Top = cmdFind.Top
        cmdFindCancle.Top = cmdFind.Top
        
        .SetFocus
        .Refresh
    End With
    
    load���ʷ��� 2
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Sub vsUseDept_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim rsTmp As New ADODB.Recordset
    If KeyCode > 127 Then
        '���ֱ�����뺺�ֵ�����
        Call vsUseDept_KeyPress(KeyCode)
    ElseIf KeyCode = vbKeyDelete Then
        If InStr(mstr��ѡʹ�ÿ���, vsUseDept.TextMatrix(vsUseDept.Row, vsUseDept.Col + 5) & "," & vsUseDept.TextMatrix(vsUseDept.Row, vsUseDept.Col)) > 0 Then
            mstr��ѡʹ�ÿ��� = Replace(mstr��ѡʹ�ÿ���, vsUseDept.TextMatrix(vsUseDept.Row, vsUseDept.Col + 5) & "," & vsUseDept.TextMatrix(vsUseDept.Row, vsUseDept.Col), "")
        End If
        vsUseDept.TextMatrix(vsUseDept.Row, vsUseDept.Col) = ""
        vsUseDept.Cell(flexcpData, vsUseDept.Row, vsUseDept.Col) = ""
        vsUseDept.TextMatrix(vsUseDept.Row, vsUseDept.Col + 5) = ""
        vsUseDept.Cell(flexcpData, vsUseDept.Row, vsUseDept.Col + 5) = ""
    End If
    If KeyCode <> vbKeyReturn Then Exit Sub
    With Me.vsUseDept
        If .Editable = flexEDNone Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        If .Col = 4 And .Row = .Rows - 1 Then
            .AddItem "": .Row = .Rows - 1: .Col = 0
        ElseIf .Col = 4 And .Row < .Rows - 1 Then
            .Row = .Row + 1: .Col = 0
        Else
            .Col = .Col + 1
        End If
        .ShowCell .Row, .Col
    End With
End Sub

Private Sub vsUseDept_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    Dim rsTmp As New ADODB.Recordset
    Dim strTmp As String, strվ�� As String
    
    With Me.vsUseDept
        If KeyCode <> vbKeyReturn Then Exit Sub
        If .Editable = flexEDNone Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        If Trim(.EditText) = "" Then
            .EditText = .Cell(flexcpData, Row, Col)
            Exit Sub
        End If
        strTemp = UCase(Trim(.EditText))
    End With
    If vsUseDept.EditText = vsUseDept.Cell(flexcpData, Row, Col) Then Exit Sub
    
    If strTemp = "" Then Exit Sub
    
    err = 0: On Error GoTo ErrHand
    If chk�������(1).Value = 1 Then strTmp = " T.�������=2"
    If chk�������(2).Value = 1 Or chk�������(0).Value = 1 Then strTmp = strTmp & IIf(strTmp = "", "", " Or") & " T.�������=1"
    If strTmp <> "" Then strTmp = " And (" & strTmp & " Or T.�������=3)"
    gstrSql = "select distinct ID,����,����" & _
            " from ���ű� D,��������˵�� T" & _
            " where D.ID=T.����ID " & _
            " and (D.����ʱ�� is null or D.����ʱ��=to_date('3000-01-01','YYYY-MM-DD')) And T.�������<>0" & strTmp
    If cmbStationNo.Text <> "" Then
        strվ�� = Mid(cmbStationNo.Text, 1, InStr(1, cmbStationNo.Text, "-") - 1)
        gstrSql = gstrSql & " And (D.վ��=[2] Or D.վ�� is Null)"
    End If
            
    gstrSql = gstrSql & " and �������� In('�ٴ�','���','����','����','����'" & IIf(chk�������(2).Value = 1, ",'���'", "") & ") "
            
    
        
    gstrSql = gstrSql & " and (D.���� like [1] or D.���� like [1] or D.���� like [1])"
            
    gstrSql = gstrSql & " order by ����"

    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption, gstrMatch & strTemp & "%", strվ��)

    With rsTmp
        If .BOF Or .EOF Then
            MsgBox "δ�ҵ�ָ�����ţ����������룡", vbExclamation, gstrSysName
            vsUseDept.TextMatrix(Row, Col) = vsUseDept.Cell(flexcpData, Row, Col)
            vsUseDept.EditText = vsUseDept.Cell(flexcpData, Row, Col)
            Exit Sub
        End If
        If .RecordCount = 1 Then
            Me.vsUseDept.Text = !����
            vsUseDept.EditText = Me.vsUseDept.Text
            vsUseDept.Cell(flexcpData, Row, Col) = Me.vsUseDept.Text
            vsUseDept.TextMatrix(Row, Col) = Me.vsUseDept.Text
            vsUseDept.TextMatrix(Row, Col + 5) = !ID
            vsUseDept.Cell(flexcpData, Row, Col + 5) = !ID
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Me.lvwItems.Checkboxes = False
        Do While Not .EOF
            Set objItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !����)
            objItem.Icon = "Dept": objItem.SmallIcon = "Dept"
            objItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = !����

            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    
    With Me.picDept
        .Tag = "2"
        Me.lvwItems.Tag = "ʹ��"
        .Left = Me.stbInfo.Left + Me.vsUseDept.Left + Me.vsUseDept.ColWidth(0) * Col
        .Width = 4000
        If .Left > Me.Width - .Width - stbInfo.Left - Me.vsUseDept.Left Then .Left = Me.Width - .Width - stbInfo.Left - Me.vsUseDept.Left
    
        .Top = 1050 + Row * vsUseDept.RowHeight(Row)
        .Height = 3850
        
        lbl��������.Visible = False
        cboProperty.Visible = lbl��������.Visible
        ChkSelect.Visible = lbl��������.Visible
        
        lbl��������.Left = 50
        ChkSelect.Left = .Width - ChkSelect.Width - 50
        cboProperty.Width = ChkSelect.Left - cboProperty.Left - 50
        
        cmdFind.Visible = False
        txtFind.Visible = False
        cmdFindOk.Visible = False
        cmdFindCancle.Visible = False
        .ZOrder 0
        .Visible = True
    End With
    
    With Me.lvwItems
        .Left = 0
        .Top = 0
        .Width = Me.picDept.Width
        .Height = Me.picDept.Height
        
        .SetFocus
        .Refresh
        KeyCode = 0
    End With
    Exit Sub
ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub vsUseDept_KeyPress(KeyAscii As Integer)
    If vsUseDept.Editable = flexEDNone Then Exit Sub

    With vsUseDept
        If KeyAscii = 13 Then
            KeyAscii = 0
        Else
            If KeyAscii = Asc("*") Then
                KeyAscii = 0
                Call vsUseDept_CellButtonClick(.Row, .Col)
            Else
                If KeyAscii = vbKeyBack Then Exit Sub
                .ComboList = "" 'ʹ��ť״̬��������״̬
            End If
        End If
    End With
End Sub

Private Sub vsUseDept_SetupEditWindow(ByVal Row As Long, ByVal Col As Long, ByVal EditWindow As Long, ByVal IsCombo As Boolean)
    If vsUseDept.Editable = flexEDNone Then
        vsUseDept.FocusRect = flexFocusLight
        vsUseDept.ComboList = ""
    Else
        vsUseDept.FocusRect = flexFocusSolid
        vsUseDept.ComboList = "..."
    End If
End Sub



