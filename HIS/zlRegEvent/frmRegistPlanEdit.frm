VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.9600.ocx"
Begin VB.Form frmRegistPlanEdit 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�ҺŰ��ű༭"
   ClientHeight    =   9120
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11325
   Icon            =   "frmRegistPlanEdit.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9120
   ScaleWidth      =   11325
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   9840
      TabIndex        =   32
      Top             =   1065
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   9840
      TabIndex        =   31
      Top             =   600
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   9840
      TabIndex        =   37
      Top             =   1590
      Width           =   1100
   End
   Begin XtremeSuiteControls.TabControl tbPage 
      Height          =   780
      Left            =   9240
      TabIndex        =   38
      Top             =   3720
      Width           =   1575
      _Version        =   589884
      _ExtentX        =   2778
      _ExtentY        =   1376
      _StockProps     =   64
   End
   Begin VB.PictureBox picTimeSet 
      BorderStyle     =   0  'None
      Height          =   6900
      Left            =   1380
      ScaleHeight     =   6900
      ScaleWidth      =   9525
      TabIndex        =   39
      Top             =   2235
      Width           =   9525
      Begin VB.CommandButton cmdAuto 
         Caption         =   "�Զ�����(&A)"
         Height          =   350
         Left            =   5505
         TabIndex        =   57
         ToolTipText     =   "ͨ��������޺���,�Զ�����ʱ�������м���"
         Top             =   45
         Width           =   1150
      End
      Begin VB.CommandButton cmdClearAll 
         Caption         =   "ȫ��(&D)"
         Height          =   350
         Left            =   8370
         TabIndex        =   56
         ToolTipText     =   "������¼���ʱ��"
         Top             =   45
         Width           =   1150
      End
      Begin VB.CommandButton cmdSelAll 
         Caption         =   "ȫѡ(&A)"
         Height          =   350
         Left            =   6930
         TabIndex        =   55
         ToolTipText     =   "������¼���ʱ��"
         Top             =   45
         Width           =   1150
      End
      Begin VB.Frame fraӦ���� 
         Caption         =   "Ӧ���ڡ�"
         Height          =   615
         Left            =   675
         TabIndex        =   50
         Top             =   6825
         Width           =   7755
         Begin VB.OptionButton optӦ���� 
            Caption         =   "��ҽ��(����)"
            Height          =   255
            Index           =   1
            Left            =   2100
            TabIndex        =   54
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton optӦ���� 
            Caption         =   "������"
            Height          =   255
            Index           =   0
            Left            =   795
            TabIndex        =   53
            Top             =   255
            Value           =   -1  'True
            Width           =   1185
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "������(�ڿ�)"
            Height          =   255
            Left            =   3870
            TabIndex        =   52
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "���кű�"
            Height          =   255
            Left            =   5685
            TabIndex        =   51
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.CommandButton cmdOther 
         Caption         =   "������������(&T)"
         Height          =   350
         Left            =   3690
         TabIndex        =   43
         ToolTipText     =   "������¼���ʱ��"
         Top             =   45
         Width           =   1515
      End
      Begin VB.CommandButton cmd����ʱ�� 
         Caption         =   "��������(&F)"
         Height          =   350
         Left            =   2235
         TabIndex        =   42
         ToolTipText     =   "������¼���ʱ��"
         Top             =   45
         Width           =   1150
      End
      Begin VB.TextBox txtTimeOut 
         Height          =   300
         Left            =   1185
         MaxLength       =   4
         TabIndex        =   41
         Text            =   "10"
         Top             =   75
         Width           =   465
      End
      Begin VB.PictureBox picPage 
         BorderStyle     =   0  'None
         Height          =   3540
         Index           =   0
         Left            =   690
         ScaleHeight     =   3540
         ScaleWidth      =   2535
         TabIndex        =   40
         Top             =   990
         Width           =   2535
      End
      Begin MSComCtl2.UpDown udTime 
         Height          =   300
         Left            =   1650
         TabIndex        =   44
         Top             =   75
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "Frame3"
         BuddyDispid     =   196630
         OrigLeft        =   2025
         OrigTop         =   3
         OrigRight       =   2280
         OrigBottom      =   348
         Max             =   1440
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin XtremeSuiteControls.TabControl tbSubPage 
         Height          =   4875
         Left            =   225
         TabIndex        =   45
         Top             =   1380
         Width           =   2535
         _Version        =   589884
         _ExtentX        =   4471
         _ExtentY        =   8599
         _StockProps     =   64
      End
      Begin VSFlex8Ctl.VSFlexGrid vsTime 
         Height          =   5475
         Index           =   0
         Left            =   1005
         TabIndex        =   46
         Top             =   1245
         Width           =   5100
         _cx             =   8996
         _cy             =   9657
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
         BackColorFixed  =   -2147483633
         ForeColorFixed  =   -2147483630
         BackColorSel    =   -2147483635
         ForeColorSel    =   -2147483634
         BackColorBkg    =   -2147483643
         BackColorAlternate=   -2147483643
         GridColor       =   12632256
         GridColorFixed  =   0
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   16777215
         FocusRect       =   2
         HighLight       =   1
         AllowSelection  =   -1  'True
         AllowBigSelection=   0   'False
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   1
         GridLineWidth   =   1
         Rows            =   2
         Cols            =   5
         FixedRows       =   1
         FixedCols       =   0
         RowHeightMin    =   300
         RowHeightMax    =   300
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   0   'False
         FormatString    =   $"frmRegistPlanEdit.frx":000C
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
         Begin VB.CommandButton cmdɾ�� 
            Caption         =   "ɾ"
            Height          =   255
            Index           =   0
            Left            =   4320
            TabIndex        =   48
            Top             =   840
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdԤԼ 
            Caption         =   "Ԥ"
            Height          =   255
            Index           =   0
            Left            =   2685
            TabIndex        =   47
            Top             =   2535
            Visible         =   0   'False
            Width           =   375
         End
      End
      Begin VB.Label Label4 
         AutoSize        =   -1  'True
         Caption         =   "ʱ����(��)"
         Height          =   180
         Left            =   75
         TabIndex        =   49
         Top             =   135
         Width           =   1080
      End
   End
   Begin VB.PictureBox picBaseBack 
      BorderStyle     =   0  'None
      Height          =   8865
      Left            =   120
      ScaleHeight     =   8865
      ScaleWidth      =   10125
      TabIndex        =   33
      Top             =   120
      Width           =   10125
      Begin VB.Frame Frame4 
         Caption         =   "Ӧ������:"
         Height          =   3980
         Left            =   240
         TabIndex        =   36
         Top             =   4560
         Width           =   8895
         Begin VSFlex8Ctl.VSFlexGrid vsDept 
            Height          =   3480
            Left            =   150
            TabIndex        =   30
            Top             =   300
            Width           =   8595
            _cx             =   15161
            _cy             =   6138
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
            GridColor       =   -2147483628
            GridColorFixed  =   -2147483632
            TreeColor       =   -2147483632
            FloodColor      =   192
            SheetBorder     =   -2147483634
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
            Cols            =   2
            FixedRows       =   0
            FixedCols       =   0
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
         Begin VB.OptionButton opt���� 
            Caption         =   "ƽ������"
            Height          =   180
            Index           =   3
            Left            =   4335
            TabIndex        =   29
            Top             =   0
            Width           =   1020
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "��̬����"
            Height          =   180
            Index           =   2
            Left            =   3180
            TabIndex        =   28
            Top             =   0
            Width           =   1020
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "ָ������"
            Height          =   180
            Index           =   1
            Left            =   2010
            TabIndex        =   27
            Top             =   0
            Width           =   1020
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "������"
            Height          =   180
            Index           =   0
            Left            =   1020
            TabIndex        =   26
            Top             =   0
            Value           =   -1  'True
            Width           =   1020
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Ӧ��ʱ��"
         Height          =   2550
         Left            =   240
         TabIndex        =   35
         Top             =   1800
         Width           =   8925
         Begin VB.TextBox txt�޺� 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   3045
            MaxLength       =   5
            TabIndex        =   17
            Top             =   292
            Width           =   1215
         End
         Begin VB.TextBox txt��Լ 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   4980
            MaxLength       =   5
            TabIndex        =   19
            Top             =   292
            Width           =   1215
         End
         Begin VB.CheckBox chk��Ч�� 
            Caption         =   "��Ч��"
            Height          =   195
            Left            =   255
            TabIndex        =   22
            Top             =   2115
            Width           =   855
         End
         Begin VB.ComboBox cbo�� 
            Height          =   300
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   292
            Width           =   1110
         End
         Begin VB.OptionButton opt�� 
            Caption         =   "ÿ��(&W)"
            Height          =   315
            Left            =   225
            TabIndex        =   20
            Top             =   630
            Width           =   930
         End
         Begin VB.OptionButton opt�� 
            Caption         =   "ÿ��(&D)"
            Height          =   315
            Left            =   225
            TabIndex        =   14
            Top             =   285
            Width           =   960
         End
         Begin MSComCtl2.DTPicker dtpBegin 
            Height          =   300
            Left            =   1170
            TabIndex        =   23
            Top             =   2055
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   192544771
            CurrentDate     =   38091
         End
         Begin MSComCtl2.DTPicker dtpEnd 
            Height          =   300
            Left            =   3555
            TabIndex        =   25
            Top             =   2055
            Width           =   2070
            _ExtentX        =   3651
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            CustomFormat    =   "yyyy-MM-dd HH:mm:ss"
            Format          =   192544771
            CurrentDate     =   38091
         End
         Begin VSFlex8Ctl.VSFlexGrid vsPlan 
            Height          =   1275
            Left            =   1140
            TabIndex        =   21
            Top             =   675
            Width           =   7650
            _cx             =   13494
            _cy             =   2249
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
            SheetBorder     =   -2147483643
            FocusRect       =   1
            HighLight       =   1
            AllowSelection  =   -1  'True
            AllowBigSelection=   -1  'True
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
            RowHeightMax    =   300
            ColWidthMin     =   0
            ColWidthMax     =   0
            ExtendLastCol   =   -1  'True
            FormatString    =   $"frmRegistPlanEdit.frx":0081
            ScrollTrack     =   0   'False
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
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "�޺�"
            Height          =   180
            Left            =   2610
            TabIndex        =   16
            Top             =   352
            Width           =   360
         End
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "��Լ"
            Height          =   180
            Left            =   4545
            TabIndex        =   18
            Top             =   345
            Width           =   360
         End
         Begin VB.Label Label8 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "��"
            Height          =   180
            Left            =   3315
            TabIndex        =   24
            Top             =   2115
            Width           =   180
         End
      End
      Begin VB.Frame Frame2 
         Caption         =   "������Ϣ"
         Height          =   1500
         Left            =   240
         TabIndex        =   34
         Top             =   120
         Width           =   8970
         Begin VB.TextBox txtAppLimit 
            Height          =   315
            Left            =   7125
            TabIndex        =   13
            Top             =   1058
            Width           =   765
         End
         Begin VB.CheckBox chkAppoint 
            Caption         =   "��ԤԼ          ��"
            Height          =   300
            Left            =   6240
            TabIndex        =   12
            Top             =   1065
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.CheckBox chk��ſ��� 
            Caption         =   "��ſ���"
            Height          =   255
            Left            =   2130
            TabIndex        =   2
            Top             =   285
            Width           =   1095
         End
         Begin VB.ComboBox cbo���� 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   4020
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   270
            Width           =   2115
         End
         Begin VB.CheckBox chk���� 
            Caption         =   "�Һ�ʱ���뽨����"
            Height          =   195
            Left            =   3615
            TabIndex        =   11
            Top             =   1118
            Width           =   1845
         End
         Begin VB.ComboBox cbo���� 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1050
            TabIndex        =   6
            Text            =   "cbo����"
            Top             =   660
            Width           =   2115
         End
         Begin VB.ComboBox cboDoctor 
            Height          =   300
            Left            =   1050
            TabIndex        =   10
            Top             =   1065
            Width           =   2115
         End
         Begin VB.ComboBox cboItem 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   4020
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   660
            Width           =   2115
         End
         Begin VB.TextBox txt�ű� 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   1050
            MaxLength       =   5
            TabIndex        =   1
            Top             =   270
            Width           =   960
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Left            =   3600
            TabIndex        =   3
            Top             =   330
            Width           =   360
         End
         Begin VB.Label lblҽ�� 
            Alignment       =   1  'Right Justify
            AutoSize        =   -1  'True
            Caption         =   "Ժ��ҽ����"
            Height          =   180
            Left            =   120
            TabIndex        =   9
            Top             =   1125
            Width           =   900
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "��Ŀ"
            Height          =   180
            Left            =   3615
            TabIndex        =   7
            Top             =   720
            Width           =   360
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Left            =   645
            TabIndex        =   5
            Top             =   720
            Width           =   360
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "�ű�"
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
            Left            =   615
            TabIndex        =   0
            Top             =   330
            Width           =   390
         End
      End
   End
   Begin VB.Menu mnuPopu 
      Caption         =   "�����˵�"
      Visible         =   0   'False
      Begin VB.Menu mnuViewDoctor 
         Caption         =   "Ժ��ҽ��"
         Index           =   0
      End
      Begin VB.Menu mnuViewDoctor 
         Caption         =   "����Ԯҽ��"
         Index           =   1
      End
   End
End
Attribute VB_Name = "frmRegistPlanEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private mlngModule As Long, mstrPrivs As String, mlngID As Long, mfrmMain As Form, mblnChange As Boolean
Private mrs���� As ADODB.Recordset
Private mrsDoctor As ADODB.Recordset
Private mblnFirst As Boolean
Private mblnSucces As Boolean
Private mlngȱʡ�Һſ���ID  As Long '�ڹҺŰ���ʱ��������������ѡ��Ŀ��ҽ���ȱʡ
Private mrsʱ��� As ADODB.Recordset
Private mstr�����޸� As String '��ĳһ����߶���İ������Ƹ���
Private mbln�Զ�Ĭ����Լ�� As Boolean '45519 �Զ�Ĭ����Լ��
Private mblnMinorChange As Boolean
Public Enum RegistEditType
    edt_���� = 0
    edt_�޸� = 1
    edt_���� = 2
End Enum
Private mEditType As RegistEditType
'�����ϰ�ʱ��
Private Type t_�ϰ�ʱ��
  dat_�����ϰ� As Date
  dat_�����°� As Date
  dat_�����ϰ� As Date
  dat_�����°� As Date
End Type
Private t_ʱ�� As t_�ϰ�ʱ��
Private mrs�ϰ�ʱ��� As ADODB.Recordset
Private mrs�޺�          As ADODB.Recordset

Private mPlanEditType As gPlanEditType

Private Type TimeSet
    bln��ſ��� As Boolean
    blnIsInit As Boolean
    lngSelIndex As Long
    blnChange As Boolean
    str���� As String
    strӦ��ʱ�� As String
    rsAssign As ADODB.Recordset
    rsHistory As ADODB.Recordset
    rsRegPlan As ADODB.Recordset
    blnNotBrush As Boolean
    lng�ƻ�ID As Long
    lng����ID As Long
    blnOnChange As Boolean
    strKey As String
    str�����޸� As String
End Type

Private mTimeSet As TimeSet
Private mintSysAppLimit As Integer
Private WithEvents mfrmOtherCalc As frmRegistPlanTimeOther
Attribute mfrmOtherCalc.VB_VarHelpID = -1

Private mstr����ID As String
Private mblnCboClick As Boolean     '�����cbo��keypress�¼������˵����б��API����:sendmessage,�����ͣ��cbo��,����һ���ַ�,�ƿ�����򰴻س���,
'                                    cbo��ֵ�ᱣ������,�����ᴥ��click�¼�,������Ҫ��validate�¼��е���click�¼�
Private mblnOnlyԺ��ҽ�� As Boolean '��ֻ����Ժ��ҽ��
Private Type PlanInfo               '���Ÿı���Ҫ�Աȵ���Ϣ
    strӦ��ʱ��      As String
    str�Ű�         As String       '�Ű���Ϣ
    str�޺�         As String       '�޺���Ϣ
    bln���         As Boolean      '�Ƿ���ſ���
    blnʱ���       As Boolean      '�Ƿ�������ʱ���
End Type

Private mPlanInfo     As PlanInfo 'ԭʼ�İ�����Ϣ  ��Ҫ���ڰ����޸�ʱ ��Ӧ��Ϣ�ıȽ�

Private Enum mPageIndex
    EM_���� = 0
    EM_ʱ�� = 1
End Enum

Private Enum mPgIndex
    Pg_�ƻ����� = 1
    Pg_�ƻ�ʱ�� = 2
End Enum

Private mblnChangeByCode As Boolean '�Ƿ��Ǵ�����Ƹı���tabelpage����ʾҳ
Private mrsRegOldData As ADODB.Recordset '�������ݼ�����,ԭʼ�ҺŰ���
Private mrsRegNewData As ADODB.Recordset '�������ݼ����� �������ú�İ���
Private mrsRegHistory As ADODB.Recordset '���ιҺŵ����ݼ�
Private mcllԤԼ��Ϣ  As Collection '�����Ѿ�ԤԼ��ȥ��ԤԼ��Ϣ K����_���� /K����_����
Private mblnChangeDist As Boolean


Private Function LoadCard() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��������
    '����:���سɹ�,����True,���򷵻�False
    '����:���˺�
    '����:2009-09-15 12:14:37
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL          As String
    Dim rsTemp          As New ADODB.Recordset
    Dim i               As Long
    Dim j               As Long
    Dim strTemp         As String
    Dim rs�޺�          As ADODB.Recordset
    Dim blnÿ��         As Boolean
    Dim bln�޺�         As Boolean
    Dim str�޺�         As String
    Dim bln��Լ         As Boolean
    Dim str��Լ         As String
    Dim blnExitFor      As Boolean
    Dim rsTmp           As New ADODB.Recordset
    Err = 0: On Error GoTo Errhand:

    
    If mEditType = edt_���� Then
        txt�ű�.Text = GetNext�ű�
        txt�޺�.Text = ""
        txt��Լ.Text = ""
        chk����.Value = 0

        If cbo����.ListIndex >= 0 Then
            If mlngȱʡ�Һſ���ID <> cbo����.ItemData(cbo����.ListIndex) Then
                cbo����.ListIndex = -1
                cboItem.ListIndex = -1
                cboDoctor.Text = ""
            End If
        Else
            cbo����.ListIndex = -1
            cboItem.ListIndex = -1
            cboDoctor.Text = ""
        End If
        dtpBegin.Value = Format(zlDatabase.Currentdate, "yyyy-MM-dd 00:00:00")
        dtpEnd.Value = CDate("3000-01-01")

        opt��.Value = True
        cbo��.Enabled = True
        cbo��.ListIndex = cbo.FindIndex(cbo��, "ȫ��", True)
        If cbo��.ListIndex = -1 Then cbo��.ListIndex = 0
        opt��.Value = False
        vsPlan.Enabled = False
        LoadCard = True
        opt����(0).Value = True
        mTimeSet.bln��ſ��� = False
        
        '����������ҵ�ѡ��
        For i = 0 To vsDept.Cols - 1
            For j = 0 To vsDept.Rows - 1
                If vsDept.Cell(flexcpChecked, j, i) <> 0 Then vsDept.Cell(flexcpChecked, j, i) = 2
            Next
        Next
        Exit Function
    End If
    '�޸Ļ�鿴
    strSQL = " " & _
    "   Select A.Id as ����ID,0 as �ƻ�ID,A.����,  A.����,  A.����id,  A.��Ŀid, A.ҽ������,  A.ҽ��id," & _
    "          A.����,  A.��һ,  A.�ܶ�,  A.����,  A.����,  A.����,  A.����,A.Ĭ��ʱ�μ��, " & _
    "           A.��������,  A.���﷽ʽ,  A.��ſ���,  A.��ʼʱ��,  A.��ֹʱ��,B.���� As ��Ŀ,D.���� As ����,A.ԤԼ���� " & _
    "   From �ҺŰ��� A,�շ���ĿĿ¼ B,���ű� D " & _
    "   Where A.��Ŀid=b.Id(+) And A.����id =d.Id(+) " & _
    "         And A.Id=[1]"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)

    If rsTemp.EOF Then
        ShowMsgbox "δ�ҵ�ָ���ĺű�,����!"
        Exit Function
    End If
    strSQL = "Select ������Ŀ,�޺���,��Լ�� From  �ҺŰ������� where ����ID=[1]       "
    Set rs�޺� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
    
    chkAppoint.Value = 0
    Do While Not rs�޺�.EOF
        If IsNull(rs�޺�!��Լ��) Then
            chkAppoint.Value = 1
            Exit Do
        Else
            If Val(Nvl(rs�޺�!��Լ��)) <> 0 Then
                chkAppoint.Value = 1
                Exit Do
            End If
        End If
        rs�޺�.MoveNext
    Loop
    If rs�޺�.RecordCount <> 0 Then rs�޺�.MoveFirst
    
    cbo����.ListIndex = cbo.FindIndex(cbo����, Nvl(rsTemp!����), True)
    txt�ű�.Text = Nvl(rsTemp!����)

    cbo����.ListIndex = cbo.FindIndex(cbo����, Nvl(rsTemp!����), True)
    cboItem.ListIndex = cbo.FindIndex(cboItem, Nvl(rsTemp!��Ŀ), True)

    cboDoctor.ListIndex = cbo.FindIndex(cboDoctor, Nvl(rsTemp!ҽ������), True)
    If cboDoctor.ListIndex = -1 Then cboDoctor.Text = Nvl(rsTemp!ҽ������)


    chk����.Value = IIf(Val(Nvl(rsTemp!��������)) = 1, 1, 0)

    chk��ſ���.Value = IIf(Val(Nvl(rsTemp!��ſ���)) = 1, 1, 0):     chk��ſ���.Tag = chk��ſ���.Value
    mTimeSet.bln��ſ��� = Val(rsTemp!��ſ���) = 1
    
    '��ȡ�޸�ǰ�İ����Ƿ���ſ���
    mPlanInfo.bln��� = IIf(Val(Nvl(rsTemp!��ſ���)) = 1, True, False)
    '��Чʱ�䷶Χ
    dtpBegin.Value = Format(zlDatabase.Currentdate, "yyyy-MM-dd 00:00:00")
    dtpEnd.Value = CDate("3000-01-01")
    If Not IsNull(rsTemp!��ʼʱ��) Then
        chk��Ч��.Value = 1
        dtpBegin.Value = CDate(Format(rsTemp!��ʼʱ��, "yyyy-mm-dd HH:MM:SS"))
        If Not IsNull(rsTemp!��ֹʱ��) Then
            dtpEnd.Value = CDate(Format(rsTemp!��ֹʱ��, "yyyy-mm-dd HH:MM:SS"))
        End If
    End If

     '����ԭʼ���ݵ����ݼ�
     With mrsRegOldData
        Set mrsRegOldData = New ADODB.Recordset
        mrsRegOldData.Fields.Append "ID", adBigInt, 18
        mrsRegOldData.Fields.Append "������Ŀ", adVarChar, 20
        mrsRegOldData.Fields.Append "�޺���", adBigInt, 10
        mrsRegOldData.Fields.Append "��Լ��", adBigInt, 18
        mrsRegOldData.Fields.Append "��ſ���", adBigInt, 18
        mrsRegOldData.CursorLocation = adUseClient
        mrsRegOldData.LockType = adLockOptimistic
        mrsRegOldData.CursorType = adOpenStatic
        mrsRegOldData.Open


        rs�޺�.Filter = 0
        If rs�޺�.RecordCount > 0 Then rs�޺�.MoveFirst
        Do While Not rs�޺�.EOF
            With mrsRegOldData
                .AddNew
                !id = mlngID
                !������Ŀ = Nvl(rs�޺�!������Ŀ)
                !�޺��� = Val(Nvl(rs�޺�!�޺���))
                !��Լ�� = Val(Nvl(rs�޺�!��Լ��))
                !��ſ��� = Val(Nvl(rsTemp!��ſ���))
                .Update
            End With
            rs�޺�.MoveNext
        Loop
    End With

    Call LoadRegHistory

    '---------------------------------------------------
    '�ж� ÿ�հ��� �޺��� ��Լ�� ���Ƿ�һ��
    '---------------------------------------------------
    blnÿ�� = Nvl(rsTemp!����) <> Nvl(rsTemp!��һ) Or Nvl(rsTemp!����) <> Nvl(rsTemp!�ܶ�) _
        Or Nvl(rsTemp!����) <> Nvl(rsTemp!����) Or Nvl(rsTemp!����) <> Nvl(rsTemp!����) _
        Or Nvl(rsTemp!����) <> Nvl(rsTemp!����) Or Nvl(rsTemp!����) <> Nvl(rsTemp!����)
    
    If blnÿ�� = False Then
             rs�޺�.Filter = "������Ŀ='����'"
             If Not rs�޺�.EOF Then
                str�޺� = Nvl(rs�޺�!�޺���)
                str��Լ = Nvl(rs�޺�!��Լ��)
             End If
            For i = 1 To 6
                strTemp = Switch(i = 0, "��", i = 1, "һ", i = 2, "��", i = 3, "��", i = 4, "��", i = 5, "��", True, "��")
                rs�޺�.Filter = "������Ŀ='" & "��" & strTemp & "'"
                If Not rs�޺�.EOF Then
                    bln�޺� = Nvl(rs�޺�!�޺���) = str�޺�
                    bln��Լ = Nvl(rs�޺�!��Լ��) = str��Լ
                    If bln��Լ = False Or bln�޺� = False Then Exit For
                End If
            Next
          blnÿ�� = True
         If bln�޺� And bln��Լ Then blnÿ�� = False
    End If

   If blnÿ�� Or mrsRegHistory.RecordCount > 0 Then
        'ÿ��
        opt��.Value = True
        With vsPlan
            For i = 1 To .Cols - 1
                strTemp = Switch(i - 1 = 0, "��", i - 1 = 1, "һ", i - 1 = 2, "��", i - 1 = 3, "��", i - 1 = 4, "��", i - 1 = 5, "��", True, "��")
                .TextMatrix(1, i) = Nvl(rsTemp.Fields("��" & strTemp))
                rs�޺�.Filter = "������Ŀ='" & "��" & strTemp & "'"
                If Not rs�޺�.EOF Then
                    .TextMatrix(2, i) = Nvl(rs�޺�!�޺���)
                    If IsNull(rs�޺�!��Լ��) Then
                        .TextMatrix(3, i) = ""
                    Else
                        If Val(Nvl(rs�޺�!��Լ��)) = 0 Then
                            .TextMatrix(3, i) = "0"
                        Else
                            .TextMatrix(3, i) = Nvl(rs�޺�!��Լ��)
                        End If
                    End If
                End If
                If InStr(mstr�����޸�, ";��" & strTemp & ";") > 0 Then
                    .Cell(flexcpForeColor, 2, i, 3, i) = vbBlue
                End If
            Next
        End With
        opt��.Value = False: cbo��.Enabled = False: txt�޺�.Enabled = False: txt��Լ.Enabled = False
        vsPlan.Enabled = True: chk��ſ���.Enabled = mstr�����޸� = ""
    Else
        'ÿ��
        opt��.Value = True:  cbo��.ListIndex = cbo.FindIndex(cbo��, Nvl(rsTemp!����), True)
        If cbo��.ListIndex = -1 Then cbo��.ListIndex = 0:
        opt��.Value = False: vsPlan.Enabled = False
        If rs�޺�.RecordCount <> 0 Then rs�޺�.MoveFirst
        If rs�޺�.EOF = False Then
            txt�޺�.Text = Nvl(rs�޺�!�޺���)
            If IsNull(rs�޺�!��Լ��) Then
                txt��Լ.Text = ""
            Else
                If Val(Nvl(rs�޺�!��Լ��)) = 0 Then
                    txt��Լ.Text = "0"
                Else
                    txt��Լ.Text = Nvl(rs�޺�!��Լ��)
                End If
            End If
        End If
        If chkAppoint.Value = 0 Then
            txt��Լ.Enabled = False
            txt��Լ.Text = ""
        Else
            txt��Լ.Enabled = True
        End If
    End If

    '------------------------------
    '��ȡ�޸�ǰ�� ʱ��κ� �޺���
    '�����ڱ���ʱ �Ա��޺���Լ�Լ�ʱ����Ƿ����˱仯
    '��������˱仯����Ҫ��ʾ  ����Ա��������ʱ����Ϣ
    '------------------------------
   mPlanInfo.str�Ű� = ""
   mPlanInfo.str�޺� = ""
   mPlanInfo.strӦ��ʱ�� = ""
    If blnÿ�� Or mrsRegHistory.RecordCount > 0 Then
         For i = 1 To vsPlan.Cols - 1
            mPlanInfo.str�Ű� = mPlanInfo.str�Ű� & "'" & Trim(vsPlan.TextMatrix(1, i)) & "',"
            mPlanInfo.strӦ��ʱ�� = mPlanInfo.strӦ��ʱ�� & "|" & Switch(i = 1, "����", i = 2, "��һ", i = 3, "�ܶ�", i = 4, "����", i = 5, "����", i = 6, "����", True, "����") & "-" & Trim(vsPlan.TextMatrix(1, i))
                mPlanInfo.str�޺� = mPlanInfo.str�޺� & "|" & Switch(i = 1, "����", i = 2, "��һ", i = 3, "�ܶ�", i = 4, "����", i = 5, "����", i = 6, "����", True, "����")
                If Trim(vsPlan.TextMatrix(1, i)) = "" Then
                     mPlanInfo.str�޺� = mPlanInfo.str�޺� & ",0,0"
                Else
                     mPlanInfo.str�޺� = mPlanInfo.str�޺� & "," & Val(Trim(vsPlan.TextMatrix(2, i))) & "," & Trim(vsPlan.TextMatrix(3, i))
                End If
        Next
    Else
         For i = 1 To 7
             mPlanInfo.str�Ű� = mPlanInfo.str�Ű� & "'" & Trim(cbo��.Text) & "',"
             mPlanInfo.strӦ��ʱ�� = mPlanInfo.strӦ��ʱ�� & "|" & Switch(i = 1, "����", i = 2, "��һ", i = 3, "�ܶ�", i = 4, "����", i = 5, "����", i = 6, "����", True, "����") & "-" & Trim(cbo��.Text)
             mPlanInfo.str�޺� = mPlanInfo.str�޺� & "|" & Switch(i = 1, "����", i = 2, "��һ", i = 3, "�ܶ�", i = 4, "����", i = 5, "����", i = 6, "����", True, "����")
             mPlanInfo.str�޺� = mPlanInfo.str�޺� & "," & Val(txt�޺�.Text) & "," & txt��Լ.Text
        Next
    End If
    If mPlanInfo.str�޺� <> "" Then mPlanInfo.str�޺� = Mid(mPlanInfo.str�޺�, 2)
    If mPlanInfo.strӦ��ʱ�� <> "" Then mPlanInfo.strӦ��ʱ�� = Mid(mPlanInfo.strӦ��ʱ��, 2)
    '-------------------------------

     Select Case Val(Nvl(rsTemp!���﷽ʽ))     '0-�����1-ָ�����ҡ�2-��̬���3-ƽ������,��Ӧ������������
        Case 0  '"������"
            opt����(0).Value = True
        Case 1  ' "ָ������"
            opt����(1).Value = True
        Case 2 '"��̬����"
            opt����(2).Value = True
        Case 3 ' "ƽ������"
            opt����(3).Value = True
    End Select

    strSQL = "Select �ű�ID,�������ҡ�From �ҺŰ������� Where �ű�ID=[1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
    '71253 ���ϴ� 2014-04-15 11:30:10 ��listView �滻ΪvsflexGrid
    
    With vsDept
        blnExitFor = False
        Do While Not rsTmp.EOF
            For i = 0 To .Cols - 1
                For j = 0 To .Rows - 1
                    If Nvl(rsTmp!��������) = .TextMatrix(j, i) Then
                        .Cell(flexcpChecked, j, i) = 1
                        blnExitFor = True
                        Exit For
                    End If
                Next
                If blnExitFor Then blnExitFor = False: Exit For
            Next
            rsTmp.MoveNext
        Loop
    End With
    rsTmp.Close
    
    If mstr�����޸� <> "" Then opt��.Enabled = False
    '������޸�ʱ ��ȡԭ���İ����Ƿ��Ѿ�������ʱ��
    If mEditType = edt_�޸� Then mPlanInfo.blnʱ��� = Checkʱ��
    If mrsRegHistory.RecordCount > 0 Then opt��.Enabled = False
    If chkAppoint.Value = 1 Then
        txtAppLimit.Enabled = True
        txtAppLimit.Text = Nvl(rsTemp!ԤԼ����, mintSysAppLimit)
    Else
        txtAppLimit.Enabled = False
        txtAppLimit.Text = Nvl(rsTemp!ԤԼ����, mintSysAppLimit)
    End If
    LoadCard = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Function


Private Sub chkAppoint_Click()
    Dim i As Integer
    If chkAppoint.Value = 0 Then
        If opt��.Value = True Then
            txt��Լ.Enabled = False
            txt��Լ.BackColor = &H8000000F
        End If
        txt��Լ.Text = ""
        txtAppLimit.Enabled = False
        For i = 1 To vsPlan.Cols - 1
            vsPlan.TextMatrix(3, i) = ""
        Next i
    Else
        If opt��.Value = True Then
            txt��Լ.Enabled = True
            txt��Լ.BackColor = vbWhite
        End If
        txtAppLimit.Enabled = True
        If Val(txt��Լ.Text) = 0 Then txt��Լ.Text = ""
        For i = 1 To vsPlan.Cols - 1
            If Val(vsPlan.TextMatrix(3, i)) = 0 Then vsPlan.TextMatrix(3, i) = ""
        Next i
    End If
End Sub

Private Sub cmdAuto_Click()
    If AutoAssignReapportion(tbSubPage.Item(mTimeSet.lngSelIndex).Caption) = False Then Exit Sub
    Call tbSubPage_SelectedChanged(tbSubPage.Item(mTimeSet.lngSelIndex))
End Sub

Private Function AutoAssignReapportion(ByVal str������Ŀ As String) As Boolean
    Dim strSQL As String
    Dim i As Long
    Dim j As Long
    Dim lng�޺� As Long
    Dim lng��Լ As Long
    Dim dat��ʼʱ�� As Date
    Dim dat����ʱ�� As Date
    Dim lng��� As Long
    Dim strTmp As String
    Dim strʱ�� As String
    Dim str����ʱ�� As String
    Dim lngĬ�ϼ�� As Long
    Dim lng������� As Long
    Dim lng�̶����� As Long
    Dim lngTmp As Long
    Dim blnExit As Boolean
    Dim datʱ�� As Date
    Dim lng����ʱ�� As Long
    Dim lng��ʼ��� As Long
    Dim lng����̯���� As Long
    Dim lng����̯���� As Long
    Dim lng���ʱ�� As Long
    If mrs�ϰ�ʱ��� Is Nothing Then
        Call Initʱ���
    End If

    If mrs�ϰ�ʱ��� Is Nothing Then Exit Function
    mTimeSet.rsRegPlan.Filter = "������Ŀ='" & str������Ŀ & "'"
    If mTimeSet.rsRegPlan.RecordCount = 0 Then mTimeSet.rsRegPlan.Filter = 0: Exit Function
    lng�޺� = Nvl(mTimeSet.rsRegPlan!�޺���, 0): lng��Լ = Nvl(mTimeSet.rsRegPlan!��Լ��, 0)
    If lng��Լ = 0 Then lng��Լ = lng�޺�
    If lng�޺� = 0 Then
        MsgBox "��ǰ�ű���" & str������Ŀ & ",û�жԹҺ�����������,�޷�����ʱ��,����!", vbOKOnly, Me.Caption
        Exit Function
    End If

    strʱ�� = mTimeSet.rsRegPlan!�Ű�
    mrs�ϰ�ʱ���.Filter = "ʱ���='" & strʱ�� & "'"
    If mrs�ϰ�ʱ���.RecordCount = 0 Then
        MsgBox "������ʱ��Ϊ[" & strʱ�� & "]�����°�ʱ��,����!", vbOKOnly, Me.Caption
        Exit Function
    End If
    
    mTimeSet.rsAssign.Filter = "������Ŀ='" & str������Ŀ & "' And ��ʹ��=0"
    Do While Not mTimeSet.rsAssign.EOF
        mTimeSet.rsAssign.Delete adAffectCurrent
        mTimeSet.rsAssign.MoveNext
    Loop
    mTimeSet.rsAssign.Filter = "������Ŀ='" & str������Ŀ & "'"
    If mTimeSet.rsAssign.RecordCount <> 0 Then
        lng�̶����� = mTimeSet.rsAssign.RecordCount
        lngĬ�ϼ�� = Val(Nvl(mTimeSet.rsAssign!ʱ����, lng���ʱ��))
        Do While Not mTimeSet.rsAssign.EOF
            lng������� = lng������� + Val(Nvl(mTimeSet.rsAssign!��������))
            mTimeSet.rsAssign.MoveNext
        Loop
    End If
    lng����ʱ�� = 0
    Do While Not mrs�ϰ�ʱ���.EOF
        dat��ʼʱ�� = CDate("1900-01-01 " & Format(mrs�ϰ�ʱ���!�ϰ�, "hh:mm:ss"))
        If Format(mrs�ϰ�ʱ���!�ϰ�, "hh:mm:ss") > Format(mrs�ϰ�ʱ���!�°�, "hh:mm:ss") Then
            dat����ʱ�� = CDate("1900-01-02 " & Format(mrs�ϰ�ʱ���!�°�, "hh:mm:ss"))
        Else
            dat����ʱ�� = CDate("1900-01-01 " & Format(mrs�ϰ�ʱ���!�°�, "hh:mm:ss"))
        End If
        lng����ʱ�� = lng����ʱ�� + DateDiff("n", dat��ʼʱ��, dat����ʱ��)
        mrs�ϰ�ʱ���.MoveNext
    Loop
    lng����ʱ�� = lng����ʱ�� - (lng�̶����� * lngĬ�ϼ��)
    If mTimeSet.bln��ſ��� Then
        If lng�޺� - lng������� = 0 Then Exit Function
        lng��ʼ��� = Int(lng����ʱ�� / (lng�޺� - lng�������))
        If lng��ʼ��� = 0 Then
            MsgBox "���õ��޺�������,�޷��Զ�����ʱ��!", vbInformation, gstrSysName
            Call tbSubPage_SelectedChanged(tbSubPage.Item(mTimeSet.lngSelIndex))
            Exit Function
        End If
        lng����̯���� = lng����ʱ�� - lng��ʼ��� * (lng�޺� - lng�������)
        lng����̯���� = (lng�޺� - lng�������) - lng����̯����
    Else
        If lng��Լ - lng������� = 0 Then Exit Function
        lng��ʼ��� = Int(lng����ʱ�� / (lng��Լ - lng�������))
        If lng��ʼ��� = 0 Then
            MsgBox "���õ���Լ������,�޷��Զ�����ʱ��!", vbInformation, gstrSysName
            Call tbSubPage_SelectedChanged(tbSubPage.Item(mTimeSet.lngSelIndex))
            Exit Function
        End If
        lng����̯���� = lng����ʱ�� - lng��ʼ��� * (lng��Լ - lng�������)
        lng����̯���� = (lng��Լ - lng�������) - lng����̯����
    End If
    mrs�ϰ�ʱ���.MoveFirst
    
    mTimeSet.rsAssign.Filter = 0
    j = 1: i = 1
    Do While Not mrs�ϰ�ʱ���.EOF
        dat��ʼʱ�� = CDate("1900-01-01 " & Format(mrs�ϰ�ʱ���!�ϰ�, "hh:mm:ss"))
        If Format(mrs�ϰ�ʱ���!�ϰ�, "hh:mm:ss") > Format(mrs�ϰ�ʱ���!�°�, "hh:mm:ss") Then
            dat����ʱ�� = CDate("1900-01-02 " & Format(mrs�ϰ�ʱ���!�°�, "hh:mm:ss"))
        Else
            dat����ʱ�� = CDate("1900-01-01 " & Format(mrs�ϰ�ʱ���!�°�, "hh:mm:ss"))
        End If

        If blnExit Then Exit Do
        datʱ�� = dat��ʼʱ��
        mrs�ϰ�ʱ���.MoveNext

        If mTimeSet.bln��ſ��� Then
            For i = j To lng�޺�
                If Format(datʱ��, "yyyy-MM-dd hh:mm:ss") >= Format(dat����ʱ��, "yyyy-MM-dd hh:mm:ss") Then
                    j = i
                    Exit For
                End If
                If i > lng�̶����� Then
                    With mTimeSet.rsAssign
                        .AddNew
                        !������Ŀ = str������Ŀ
                        !��ʼʱ�� = Format(datʱ��, "hh:mm:00")
                        !ʱ�� = Format(datʱ��, "hh:00:00")
                        If lng����̯���� > 0 Then
                            If Format(DateAdd("n", lng��ʼ���, datʱ��), "yyyy-MM-dd hh:mm:00") >= Format(dat����ʱ��, "yyyy-MM-dd hh:mm:ss") Then
                                !����ʱ�� = Format(dat����ʱ��, "hh:mm:ss")
                                !ʱ��� = Format(datʱ��, "hh:mm") & "-" & Format(dat����ʱ��, "hh:mm")
                            Else
                                !����ʱ�� = Format(DateAdd("n", lng��ʼ���, datʱ��), "hh:mm:00")
                                !ʱ��� = Format(datʱ��, "hh:mm") & "-" & Format(DateAdd("n", lng��ʼ���, datʱ��), "hh:mm")
                            End If
                        Else
                            If Format(DateAdd("n", lng��ʼ��� + 1, datʱ��), "yyyy-MM-dd hh:mm:00") >= Format(dat����ʱ��, "yyyy-MM-dd hh:mm:ss") Then
                                !����ʱ�� = Format(dat����ʱ��, "hh:mm:ss")
                                !ʱ��� = Format(datʱ��, "hh:mm") & "-" & Format(dat����ʱ��, "hh:mm")
                            Else
                                !����ʱ�� = Format(DateAdd("n", lng��ʼ��� + 1, datʱ��), "hh:mm:00")
                                !ʱ��� = Format(datʱ��, "hh:mm") & "-" & Format(DateAdd("n", lng��ʼ��� + 1, datʱ��), "hh:mm")
                            End If
                        End If
                        If lng����̯���� > 0 Then
                            !ʱ���� = lng��ʼ���
                        Else
                            !ʱ���� = lng��ʼ��� + 1
                        End If
                        !�������� = IIf(lng������� >= lng�޺�, 0, 1)
                        !�Ƿ�ԤԼ = 0
                        !��� = i
                        !��ʹ�� = 0
                        .Update
                        lng������� = lng������� + IIf(lng������� >= lng�޺�, 0, 1)
                    End With
                    If lng����̯���� > 0 Then
                        datʱ�� = DateAdd("n", lng��ʼ���, datʱ��)
                        lng����̯���� = lng����̯���� - 1
                    Else
                        datʱ�� = DateAdd("n", lng��ʼ��� + 1, datʱ��)
                    End If
                Else
                    mTimeSet.rsAssign.Filter = "���=" & i
                    If mTimeSet.rsAssign.RecordCount > 0 Then
                        lngĬ�ϼ�� = Nvl(mTimeSet.rsAssign!ʱ����, lngĬ�ϼ��)
                    Else
                        lngĬ�ϼ�� = lng��ʼ���
                    End If
                    datʱ�� = DateAdd("n", lngĬ�ϼ��, datʱ��)
                End If
            Next

        Else    '����ſ���

            Do While Not Format(datʱ��, "yyyy-MM-dd hh:mm:ss") >= Format(dat����ʱ��, "yyyy-MM-dd hh:mm:ss")
                ' If lngStart > lng��Լ Then blnExit = True: Exit For
                If Format(datʱ��, "yyyy-MM-dd hh:mm:ss") >= Format(dat����ʱ��, "yyyy-MM-dd hh:mm:ss") Then Exit Do

                If i > lng�̶����� Then
                    With mTimeSet.rsAssign
                        .AddNew
                        !������Ŀ = str������Ŀ
                        !��ʼʱ�� = Format(datʱ��, "hh:mm:00")
                        !ʱ�� = Format(datʱ��, "hh:00:00")
                        If lng����̯���� > 0 Then
                            If Format(DateAdd("n", lng��ʼ���, datʱ��), "yyyy-MM-dd hh:mm:00") >= Format(dat����ʱ��, "yyyy-MM-dd hh:mm:ss") Then
                                !����ʱ�� = Format(dat����ʱ��, "hh:mm:ss")
                                !ʱ��� = Format(datʱ��, "hh:mm") & "-" & Format(dat����ʱ��, "hh:mm")
                            Else
                                !����ʱ�� = Format(DateAdd("n", lng��ʼ���, datʱ��), "hh:mm:00")
                                !ʱ��� = Format(datʱ��, "hh:mm") & "-" & Format(DateAdd("n", lng��ʼ���, datʱ��), "hh:mm")
                            End If
                        Else
                            If Format(DateAdd("n", lng��ʼ��� + 1, datʱ��), "yyyy-MM-dd hh:mm:00") >= Format(dat����ʱ��, "yyyy-MM-dd hh:mm:ss") Then
                                !����ʱ�� = Format(dat����ʱ��, "hh:mm:ss")
                                !ʱ��� = Format(datʱ��, "hh:mm") & "-" & Format(dat����ʱ��, "hh:mm")
                            Else
                                !����ʱ�� = Format(DateAdd("n", lng��ʼ��� + 1, datʱ��), "hh:mm:00")
                                !ʱ��� = Format(datʱ��, "hh:mm") & "-" & Format(DateAdd("n", lng��ʼ��� + 1, datʱ��), "hh:mm")
                            End If
                        End If
                        
                        If lng����̯���� > 0 Then
                            !ʱ���� = lng��ʼ���
                        Else
                            !ʱ���� = lng��ʼ��� + 1
                        End If
                        !�������� = IIf(lng������� >= lng��Լ, 0, 1)
                        !�Ƿ�ԤԼ = 1
                        !��� = i
                        !��ʹ�� = 0
                        .Update
                        lng������� = lng������� + IIf(lng������� >= lng��Լ, 0, 1)
                    End With
                    If lng����̯���� > 0 Then
                        datʱ�� = DateAdd("n", lng��ʼ���, datʱ��)
                        lng����̯���� = lng����̯���� - 1
                    Else
                        datʱ�� = DateAdd("n", lng��ʼ��� + 1, datʱ��)
                    End If
                Else
                    mTimeSet.rsAssign.Filter = "���=" & i
                    If mTimeSet.rsAssign.RecordCount > 0 Then
                        lngĬ�ϼ�� = Nvl(mTimeSet.rsAssign!ʱ����, lngĬ�ϼ��)
                    Else
                        lngĬ�ϼ�� = lng���ʱ��
                    End If
                    datʱ�� = DateAdd("n", lngĬ�ϼ��, datʱ��)
                End If
                i = i + 1
            Loop
        End If
        If i > lng�޺� And mTimeSet.bln��ſ��� Then
            blnExit = True
        End If
    Loop
    AutoAssignReapportion = True
End Function

Private Sub cmdClearAll_Click()
    If Not mTimeSet.bln��ſ��� Or mTimeSet.lngSelIndex < 0 Then Exit Sub
    With vsTime(mTimeSet.lngSelIndex)
        If .Rows = 0 Then Exit Sub
        .Cell(flexcpForeColor, 0, 1, .Rows - 1, .Cols - 1) = &H80000008
        .Cell(flexcpFontBold, 0, 1, .Rows - 1, .Cols - 1) = False
        mTimeSet.blnChange = True
        .SetFocus
    End With
End Sub

Private Sub cmdSelAll_Click()
    If Not mTimeSet.bln��ſ��� Or mTimeSet.lngSelIndex < 0 Then Exit Sub
    With vsTime(mTimeSet.lngSelIndex)
        If .Rows = 0 Then Exit Sub
        .Cell(flexcpForeColor, 0, 1, .Rows - 1, .Cols - 1) = vbBlue
        .Cell(flexcpFontBold, 0, 1, .Rows - 1, .Cols - 1) = True
        mTimeSet.blnChange = True
        .SetFocus
    End With
End Sub

Public Function ShowEdit(ByVal frmMain As Form, ByVal EditType As RegistEditType, _
    ByVal lngModule As Long, ByVal strPrivs As String, Optional lngID As Long = 0, _
    Optional lngȱʡ����ID As Long = 0) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������
    '���:frmMain-���õ�������
    '     EditType-�༭����
    '����:
    '����:
    '����:���˺�
    '����:2009-09-15 10:25:08
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Set mfrmMain = frmMain: mlngModule = lngModule: mstrPrivs = strPrivs: mlngID = lngID: mlngȱʡ�Һſ���ID = lngȱʡ����ID
    mEditType = EditType: mblnSucces = False
    mblnChange = False
    If EditType = edt_�޸� Then
        mstr�����޸� = zl_GetԤԼ��Ϣ(lngID)
    End If
    Me.Show 1, frmMain
    ShowEdit = mblnSucces
    
End Function

Private Sub cboDoctor_KeyDown(KeyCode As Integer, Shift As Integer)
'    If KeyCode = vbKeyReturn Then
'        cboDoctor.ListIndex = GetCboIndex(cboDoctor, cboDoctor)
'    End If
End Sub

Private Sub cboDoctor_KeyPress(KeyAscii As Integer)
    Dim lngIdx As Long, lngҽ��ID As Long
    If KeyAscii <> 13 Then Exit Sub
    If cboDoctor.ListIndex <> -1 Then
        zlCommFun.PressKey vbKeyTab: Exit Sub
    End If
    If mrsDoctor Is Nothing Then Exit Sub
    If Trim(cboDoctor.Text) = "" Then zlCommFun.PressKey vbKeyTab: Exit Sub
    
    If zlPersonSelect(Me, mlngModule, cboDoctor, mrsDoctor, cboDoctor.Text, True, "") = False Then
        If mblnOnlyԺ��ҽ�� = False Then
                zlCommFun.PressKey vbKeyTab
        End If
        KeyAscii = 0: Exit Sub
    End If
    Exit Sub
End Sub

Private Sub cboDoctor_Validate(Cancel As Boolean)
      If mblnOnlyԺ��ҽ�� Then
           If cboDoctor.ListIndex < 0 Then cboDoctor.Text = ""
      End If
      
    'ָ��ҽ��ʱ����ָ���������
    If Trim(cboDoctor.Text) <> "" Then
        opt����(2).Enabled = False
        opt����(3).Enabled = False
        If opt����(2).Value Or opt����(3).Value Then opt����(0).Value = True
    Else
        opt����(2).Enabled = True
        opt����(3).Enabled = True
    End If
End Sub

Private Sub cbo����_Click()
    mblnCboClick = True
    If cbo����.ListIndex = -1 Then Exit Sub
    Call LoadDoctor
End Sub

Private Sub LoadDoctor()
    Set mrsDoctor = GetDoctor(Val(cbo����.ItemData(cbo����.ListIndex)), "")
    cboDoctor.Clear
    Do While Not mrsDoctor.EOF
        cboDoctor.AddItem mrsDoctor!����
        cboDoctor.ItemData(cboDoctor.NewIndex) = mrsDoctor!id
        mrsDoctor.MoveNext
    Loop
End Sub

Private Sub cbo����_GotFocus()
    zlControl.TxtSelAll cbo����
End Sub

Private Sub cbo����_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        If cbo����.Text = "" Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        If cbo����.ListIndex >= 0 Then Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        mblnCboClick = True
        If Select����(Me, mlngModule, mrs����, cbo����, cbo����.Text) = True Then
            mblnCboClick = False
            Call zlCommFun.PressKey(vbKeyTab): Exit Sub
        End If
        If cbo����.Enabled Then cbo����.SetFocus
        mblnCboClick = False
        zlControl.TxtSelAll cbo����
    Else
       ' Call zlControl.CboSetIndex(cbo����.hWnd, zlControl.CboMatchIndex(cbo����.hWnd, KeyAscii))
    End If
End Sub

Private Sub cbo����_Validate(Cancel As Boolean)
 '�����cbo��keypress�¼������˵����б�ĵ�API����:sendmessage,�����ͣ��cbo��,����һ���ַ�,�ƿ�����򰴻س���,
'                                    cbo��ֵ�ᱣ������,�����ᴥ��click�¼�,������Ҫ��validate�¼��е���click�¼�
    If Not mblnCboClick Then cbo����_Click
    mblnCboClick = False
End Sub

Private Sub chk��Ч��_Click()
    dtpBegin.Enabled = chk��Ч��.Value = 1
    dtpEnd.Enabled = chk��Ч��.Value = 1
    
    If Visible And dtpBegin.Enabled Then
        dtpBegin.SetFocus
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub

Private Function GetDoctorPlan(lngҽ��ID As Long, strҽ������ As String) As ADODB.Recordset
'����:����ָ��ҽ��ID�����������кű��ʱ����Ϣ
'   ���ڼ���������޸ĵĺű��Ƿ������еĺű���ʱ�����ظ�
    Dim strSQL As String
    
    On Error GoTo errH
    strSQL = "Select ����,���� D0,��һ D1,�ܶ� D2,���� D3,���� D4,���� D5,���� D6," & _
            " To_Char(��ʼʱ��,'YYYY-MM-DD HH24:MI:SS') ��ʼʱ��,To_Char(��ֹʱ��,'YYYY-MM-DD HH24:MI:SS') ��ֹʱ��" & _
            " From �ҺŰ��� Where (��ֹʱ�� is null or ��ֹʱ��>sysdate) And " & IIf(lngҽ��ID <> 0, " ҽ��ID=[1]", " ҽ������=[1]") & _
            IIf(mEditType = edt_����, "", " And ID<>[2]")
    Set GetDoctorPlan = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIf(lngҽ��ID <> 0, lngҽ��ID, strҽ������), mlngID)

    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function CheckExistsBooking() As Boolean
'����:��鵱ǰʱ���֮���Ƿ����ԤԼ�Һŵ�
    Dim rsTemp As ADODB.Recordset, rsBooking As ADODB.Recordset, strSQL As String
    Dim i As Long, strʱ��� As String
        
    On Error GoTo errH
    If opt��.Value Then
        strʱ��� = _
               "Select 1 From ʱ��� b Where b.ʱ��� = [2] And (" & _
               " ('3000-01-10 '||To_Char(a.����ʱ��,'HH24:MI:SS')" & _
               " Between" & _
               " Decode(Sign(b.��ʼʱ��-b.��ֹʱ��),1,'3000-01-09 '||To_Char(b.��ʼʱ��,'HH24:MI:SS'),'3000-01-10 '||To_Char(b.��ʼʱ��,'HH24:MI:SS'))" & _
               " And" & _
               " '3000-01-10 '||To_Char(b.��ֹʱ��,'HH24:MI:SS'))" & _
               " Or" & _
               " ('3000-01-10 '||To_Char(a.����ʱ��,'HH24:MI:SS')" & _
               " Between" & _
               " '3000-01-10 '||To_Char(b.��ʼʱ��,'HH24:MI:SS')" & _
               " And" & _
               " Decode(Sign(b.��ʼʱ��-b.��ֹʱ��),1,'3000-01-11 '||To_Char(b.��ֹʱ��,'HH24:MI:SS'),'3000-01-10 '||To_Char(b.��ֹʱ��,'HH24:MI:SS'))))"
        
        strSQL = "Select  /*+ Rule*/ Min(����ʱ��) ʱ��" & vbNewLine & _
            "From ������ü�¼ a" & vbNewLine & _
            "Where ��¼���� = 4 And ��¼״̬ In (0, 1) And ���㵥λ = [1] And ����ʱ�� > �Ǽ�ʱ��"
        If gintԤԼ���� = 0 Then
            strSQL = strSQL & " And ����ʱ�� > Sysdate"
        Else
            strSQL = strSQL & " And ����ʱ�� Between Sysdate And Sysdate+" & gintԤԼ����
        End If
        strSQL = strSQL & " And Not Exists (" & strʱ��� & ")"
        
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, txt�ű�.Text, Trim(cbo��.Text))
        CheckExistsBooking = Not IsNull(rsTemp!ʱ��)
    Else
        strSQL = "Select /*+ Rule*/ ����ʱ��,To_Char(����ʱ��,'D') ���� From ������ü�¼ a Where ��¼���� = 4 and ��¼״̬ In(0,1) And ���㵥λ = [1] And ����ʱ�� > �Ǽ�ʱ��"
        If gintԤԼ���� = 0 Then
            strSQL = strSQL & " And ����ʱ�� > Sysdate"
        Else
            strSQL = strSQL & " And ����ʱ�� Between Sysdate And Sysdate+" & gintԤԼ����
        End If
        
        Set rsBooking = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, txt�ű�.Text)
        For i = 1 To rsBooking.RecordCount
            strʱ��� = Trim(vsPlan.TextMatrix(1, rsBooking!���� - 1))
            If strʱ��� = "" Then
                CheckExistsBooking = True
            Else
               strSQL = _
                    "Select Count(*) cnt From ʱ��� b Where b.ʱ��� = [2] And (" & _
                    " ('3000-01-10 '||To_Char([1],'HH24:MI:SS')" & _
                    " Between" & _
                    " Decode(Sign(b.��ʼʱ��-b.��ֹʱ��),1,'3000-01-09 '||To_Char(b.��ʼʱ��,'HH24:MI:SS'),'3000-01-10 '||To_Char(b.��ʼʱ��,'HH24:MI:SS'))" & _
                    " And" & _
                    " '3000-01-10 '||To_Char(b.��ֹʱ��,'HH24:MI:SS'))" & _
                    " Or" & _
                    " ('3000-01-10 '||To_Char([1],'HH24:MI:SS')" & _
                    " Between" & _
                    " '3000-01-10 '||To_Char(b.��ʼʱ��,'HH24:MI:SS')" & _
                    " And" & _
                    " Decode(Sign(b.��ʼʱ��-b.��ֹʱ��),1,'3000-01-11 '||To_Char(b.��ֹʱ��,'HH24:MI:SS'),'3000-01-10 '||To_Char(b.��ֹʱ��,'HH24:MI:SS'))))"
                Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, CDate(rsBooking!����ʱ��), strʱ���)
                CheckExistsBooking = rsTemp!cnt = 0
            End If
            
            If CheckExistsBooking Then Exit Function
            rsBooking.MoveNext
        Next
    End If
    
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub SaveMinorChange()
    Dim strSQL As String, intCount As Integer
    Dim str���� As String, lngԤԼ���� As Long
    Dim i As Long
    Dim j As Long
    Dim rsTemp As ADODB.Recordset
    Dim intSync As Integer
    If CheckRegistDays(Trim(txt�ű�.Text)) = False Then Exit Sub
    '�����ж�
    If opt����(1).Value Or opt����(2).Value Or opt����(3).Value Then
        intCount = 0
        With vsDept
            For i = 0 To .Cols - 1
                For j = 0 To .Rows - 1
                    If .Cell(flexcpChecked, j, i) = 1 Then intCount = intCount + 1
                Next
            Next
        End With
        If opt����(1).Value Then
            If intCount = 0 Then
                MsgBox "ָ������ʱ����ѡ��һ����Ӧ���������ң�", vbInformation, gstrSysName
                vsDept.SetFocus: Exit Sub
            ElseIf intCount > 1 Then
                MsgBox "ָ������ʱֻ��ѡ��һ����Ӧ���������ң�", vbInformation, gstrSysName
                vsDept.SetFocus: Exit Sub
            End If
        ElseIf opt����(2).Value Or opt����(3).Value Then
            If intCount < 2 Then
                MsgBox "��̬�����ƽ������ʱ����Ҫѡ��������Ӧ���������ң�", vbInformation, gstrSysName
                vsDept.SetFocus: Exit Sub
            End If
        End If
    End If
    
    'ȡ���﷽ʽ
    intCount = 0
    For i = 0 To opt����.UBound
        If opt����(i).Value Then intCount = i: Exit For
    Next
    
    If chkAppoint.Value = 1 Then
        lngԤԼ���� = IIf(txtAppLimit.Text = "", gintԤԼ����, txtAppLimit.Text)
    Else
        lngԤԼ���� = 0
    End If
    
    With vsDept
        For i = 0 To .Cols - 1
            For j = 0 To .Rows - 1
                If .Cell(flexcpChecked, j, i) = 1 Then str���� = str���� & ";" & .TextMatrix(j, i)
            Next
        Next
    End With
    str���� = Mid(str����, 2)
    
    strSQL = "Select 1 From �ҺŰ��żƻ� Where ����ID=[1] And ʧЧʱ�� > Sysdate"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
    If Not rsTemp.EOF And mblnChangeDist Then
        If MsgBox("�޸ĵİ��Ŵ�������Ч��δ��Ч�ļƻ�,�Ƿ�ͬ�����ļƻ�����������?", vbQuestion + vbYesNo, gstrSysName) = vbYes Then
            intSync = 1
        Else
            intSync = 0
        End If
    End If
    
    strSQL = "Zl_�ҺŰ���_Modify("
    strSQL = strSQL & mlngID & ",'"
    strSQL = strSQL & str���� & "',"
    strSQL = strSQL & lngԤԼ���� & ","
    strSQL = strSQL & intCount & ","
    strSQL = strSQL & intSync & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    Unload Me
End Sub


Private Function CheckRegistDays(ByVal str�ű� As String) As Boolean
'����:���ԤԼ����
    Dim lng������� As Long
    On Error GoTo errH
    Dim strSQL As String, rsTmp As ADODB.Recordset
    If chkAppoint.Value = 1 Then
        lng������� = Val(Nvl(txtAppLimit.Text, gintԤԼ����))
    Else
        lng������� = 0
    End If
    strSQL = "Select Max(����ʱ��) As ʱ�� From ���˹Һż�¼ Where ��¼���� = 2 And ��¼״̬ = 1 And ����ʱ�� > Sysdate + [1] And �ű� = [2]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng�������, str�ű�)
    If Not rsTmp.EOF Then
        If Nvl(rsTmp!ʱ��) <> "" Then
            MsgBox "��" & Format(rsTmp!ʱ��, "YYYY-MM-DD") & "���ڳ�����ǰԤԼ������ԤԼ��¼,���ܼ���,�뽫��ԤԼ��������!", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    
    If chkAppoint.Value = 1 And txtAppLimit.Text <> "" And Val(txtAppLimit.Text) <= 0 Then
        MsgBox "����ԤԼʱ,ԤԼ��������С�ڵ���0!", vbInformation, gstrSysName
        If txtAppLimit.Visible And txtAppLimit.Enabled Then txtAppLimit.SetFocus
        Exit Function
    End If
    
    CheckRegistDays = True
    Exit Function
errH:
    If ErrCenter = 1 Then
        Resume
    End If
End Function

Private Sub cmdOK_Click()
    Dim i As Integer, intCount As Integer, j As Integer
    Dim strʱ��� As String, str���� As String, str�޺� As String
    Dim lngNextID As Long, lngҽ��ID As Long
    Dim strBegin As String, strEnd As String
    Dim strSQL As String, strInfo As String, strTmp As String, strOld As String, strNew As String
    Dim cllPro As Collection, lngԤԼ���� As Long
    Dim str�ű� As String
    Dim rsDoctorPlan As ADODB.Recordset
    Dim rsNewDate As ADODB.Recordset
    Dim rsOldDate As ADODB.Recordset
    Dim rsTemp As ADODB.Recordset
    Dim rsSNState As ADODB.Recordset
    Dim blnMulitNumPlan As Boolean  '�Ƿ��ΰ���
    Dim blnChange       As Boolean '�Ƿ�ı��� ʱ�䰲��
    Dim strMsg          As String
    
    If mblnMinorChange Then Call SaveMinorChange: Exit Sub
    If mEditType = edt_���� Then Unload Me: Exit Sub
    If Me.tbPage.Item(mPageIndex.EM_����).Selected = False Then
        mblnChangeByCode = True
        tbPage.Item(mPageIndex.EM_����).Selected = True
        mblnChangeByCode = False
    End If
    If CheckRegistDays(Trim(txt�ű�.Text)) = False Then Exit Sub
    If mblnOnlyԺ��ҽ�� Then
        If cboDoctor.ListIndex < 0 And cboDoctor.Text <> "" Then
                MsgBox "��ѡ���ҽ��������,����������ҽ��!", vbInformation + vbOKOnly + vbDefaultButton1, gstrSysName
                If cboDoctor.Enabled Then cboDoctor.SetFocus
                Exit Sub
        End If
    End If
    '�����Լ��
    If Trim(txt�ű�) = "" Then
        MsgBox "�ű���Ϊ�գ�", vbInformation, gstrSysName
        txt�ű�.SetFocus: Exit Sub
    End If
    If cbo����.ListIndex = -1 Then
        MsgBox "δ���úű�����Ӧ�Ŀ��ң�", vbInformation, gstrSysName
        cbo����.SetFocus: Exit Sub
    End If
    If cboItem.ListIndex = -1 Then
        MsgBox "δ���úű�����Ӧ�ĹҺ���Ŀ��", vbInformation, gstrSysName
        cboItem.SetFocus: Exit Sub
    End If

    If dtpBegin.Enabled And dtpEnd.Enabled Then
        If dtpBegin.Value >= dtpEnd.Value Then
            MsgBox "��ʼʱ��Ӧ��С�ڽ���ʱ�䡣", vbInformation, gstrSysName
            dtpBegin.SetFocus: Exit Sub
        End If
    End If

    If opt��.Value Then
        If cbo��.ListIndex = -1 Then
            MsgBox "�úű�ÿ���Ӧ��ʱ��δ���ã�", vbInformation, gstrSysName
            cbo��.SetFocus: Exit Sub
        End If
        If chk��ſ���.Value = 1 Then
            If Val(txt�޺�.Text) = 0 And Val(txt��Լ.Text) = 0 Then
                MsgBox "ʹ����ſ���ʱ,���������޺Ż���Լ����", vbInformation, gstrSysName
                txt�޺�.SetFocus: Exit Sub
            End If
        End If
        '�޺���Լ����
        If Trim(txt�޺�.Text) <> "" Then
            If Trim(txt��Լ.Text) <> "" And Val(txt�޺�.Text) < Val(txt��Լ.Text) Then
                MsgBox "��Լ��ӦС���޺�����", vbInformation, gstrSysName
                txt��Լ.SetFocus: Exit Sub
            End If
        ElseIf Trim(txt��Լ.Text) <> "" Then
            MsgBox "��Լ�����޺ţ�", vbInformation, gstrSysName
            txt�޺�.SetFocus: Exit Sub
        End If
    Else
        With vsPlan
            strTmp = ""
            For i = 1 To .Cols - 1
                If Trim(.TextMatrix(1, i)) <> "" Then
                    strTmp = strTmp & Trim(vsPlan.TextMatrix(1, i))
                    If chk��ſ���.Value = 1 Then
                          If Val(.TextMatrix(2, i)) = 0 And Val(.TextMatrix(3, i)) = 0 Then
                              MsgBox "ʹ����ſ���ʱ,���������޺Ż���Լ����", vbInformation, gstrSysName
                              .Row = 2: .Col = i
                              .SetFocus: Exit Sub
                          End If
                      End If
                        '�޺���Լ����
                        If Val(.TextMatrix(2, i)) <> 0 Then
                            If Trim(.TextMatrix(3, i)) <> "" And Val(.TextMatrix(2, i)) < Val(.TextMatrix(3, i)) Then
                                MsgBox "��Լ��ӦС���޺�����", vbInformation, gstrSysName
                                .Row = 2: .Col = i
                                .SetFocus: Exit Sub
                            End If
                        ElseIf Trim(.TextMatrix(3, i)) <> "" Then
                            MsgBox "��Լ�����޺ţ�", vbInformation, gstrSysName
                            .Row = 2: .Col = i
                            .SetFocus: Exit Sub
                        End If
                End If
            Next
            If strTmp = "" Then
                MsgBox "�úű�ÿ�ܵ�Ӧ��ʱ��δ���ã�", vbInformation, gstrSysName
                vsPlan.SetFocus: Exit Sub
            End If
        End With
    End If
    
    If CheckRegistDays(Trim(txt�ű�.Text)) = False Then Exit Sub
    
    '�����ж�
    If opt����(1).Value Or opt����(2).Value Or opt����(3).Value Then
        '71253 ���ϴ� 2014-04-15 11:30:10 ��listView �滻ΪvsflexGrid
        intCount = 0
        With vsDept
            For i = 0 To .Cols - 1
                For j = 0 To .Rows - 1
                    If .Cell(flexcpChecked, j, i) = 1 Then intCount = intCount + 1
                Next
            Next
        End With
        If opt����(1).Value Then
            If intCount = 0 Then
                MsgBox "ָ������ʱ����ѡ��һ����Ӧ���������ң�", vbInformation, gstrSysName
                vsDept.SetFocus: Exit Sub
            ElseIf intCount > 1 Then
                MsgBox "ָ������ʱֻ��ѡ��һ����Ӧ���������ң�", vbInformation, gstrSysName
                vsDept.SetFocus: Exit Sub
            End If
        ElseIf opt����(2).Value Or opt����(3).Value Then
            If intCount < 2 Then
                MsgBox "��̬�����ƽ������ʱ����Ҫѡ��������Ӧ���������ң�", vbInformation, gstrSysName
                vsDept.SetFocus: Exit Sub
            End If
        End If
    End If

    '��Ŀ�۸��ж�
    If ReadRegistPrice(cboItem.ItemData(cboItem.ListIndex), False, False) = 0 Then
        MsgBox "��Ŀ""" & cboItem.Text & """δ������Ч�۸�,���ȵ��շ���Ŀ���������ã�", vbInformation, gstrSysName
        cboItem.SetFocus: Exit Sub
    End If

    'ȡҽ��ID
    If cboDoctor.ListIndex <> -1 Then lngҽ��ID = cboDoctor.ItemData(cboDoctor.ListIndex)
    If lngҽ��ID = 0 And cboDoctor.Text <> "" Then
        strSQL = "Select 1 From ��Ա�� Where ���� = [1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboDoctor.Text)
        If Not rsTemp.EOF Then
            MsgBox "ҽ��""" & cboDoctor.Text & """�����ڿ���""" & cbo����.Text & """,���������øúű�Ŀ�����ҽ����Ϣ��", vbInformation, gstrSysName
            cboDoctor.SetFocus: Exit Sub
        End If
    End If
    
'    '����:����һ��ҽ�����Լ����ظ�����
'    If zlCheckPlanArrageIsValied = False Then
'        If cboDoctor.Enabled Then cboDoctor.SetFocus
'        Exit Sub
'    End If
'
'    If zlCheckRegistPlanIsValied(blnMulitNumPlan) = False Then
'        If cboDoctor.Enabled Then cboDoctor.SetFocus
'        Exit Sub
'    End If
    '�Ƿ�ͬһҽ���İ���ʱ����Ƿ��ظ��򽻲�
    If Trim(cboDoctor.Text) <> "" Then
        Set rsDoctorPlan = GetDoctorPlan(lngҽ��ID, cboDoctor.Text)
        If rsDoctorPlan.RecordCount > 0 Then
            strSQL = "Select ʱ���, ��ʼʱ��, Decode(Sign(��ֹʱ�� - ��ʼʱ��), 1, ��ֹʱ�� , ��ֹʱ��+ 1) ��ֹʱ�� From ʱ���"
            Set rsNewDate = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
            Set rsOldDate = rsNewDate.Clone
        End If

        strInfo = ""
        For j = 1 To rsDoctorPlan.RecordCount
            strTmp = ""
            For i = 0 To IIf(opt��.Value, 6, vsPlan.Cols - 2)
               strOld = "" & rsDoctorPlan.Fields("D" & i).Value
               If opt��.Value Then
                   strNew = cbo��.Text
               Else
                   strNew = Trim(vsPlan.TextMatrix(1, i + 1))
               End If

               rsNewDate.Filter = "ʱ���='" & strNew & "'"
               rsOldDate.Filter = "ʱ���='" & strOld & "'"
               If rsNewDate.RecordCount > 0 And rsOldDate.RecordCount > 0 Then
                    If rsNewDate!��ʼʱ�� >= rsOldDate!��ʼʱ�� And rsNewDate!��ʼʱ�� <= rsOldDate!��ֹʱ�� Or rsNewDate!��ֹʱ�� >= rsOldDate!��ʼʱ�� And rsNewDate!��ֹʱ�� <= rsOldDate!��ֹʱ�� Or rsNewDate!��ʼʱ�� <= rsOldDate!��ʼʱ�� And rsNewDate!��ֹʱ�� >= rsOldDate!��ֹʱ�� Then
                    'ʱ�佻��,���ж�Ч���Ƿ񽻲�
                         If chk��Ч��.Value = 0 Then
                             strTmp = strTmp & "," & "����" & Choose(i + 1, "��", "һ", "��", "��", "��", "��", "��") & ":" & strOld
                         Else
                             'Ϊ���ж�,�ٶ����ݰ��淶����,��ʼʱ��ͽ���ʱ��,Ҫô����,Ҫô��û��,���Խ��Կ�ʼʱ�����ж�����
                             If IsNull(rsDoctorPlan!��ʼʱ��) Then
                                 strTmp = strTmp & "," & "����" & Choose(i + 1, "��", "һ", "��", "��", "��", "��", "��") & ":" & strOld
                             Else
                                 If dtpBegin.Value >= CDate(rsDoctorPlan!��ʼʱ��) And dtpBegin.Value <= CDate(Nvl(rsDoctorPlan!��ֹʱ��, "3000-01-01")) Or dtpEnd.Value >= CDate(rsDoctorPlan!��ʼʱ��) And dtpEnd.Value <= CDate(Nvl(rsDoctorPlan!��ֹʱ��, "3000-01-01")) Or dtpBegin.Value <= CDate(rsDoctorPlan!��ʼʱ��) And dtpEnd.Value >= CDate(Nvl(rsDoctorPlan!��ֹʱ��, "3000-01-01")) Then
                                    strTmp = strTmp & "," & "����" & Choose(i + 1, "��", "һ", "��", "��", "��", "��", "��") & ":" & strOld
                                 End If
                             End If
                         End If
                    End If
               End If
            Next
            If strTmp <> "" Then
                strInfo = strInfo & vbCrLf & "�ںű� [" & rsDoctorPlan!���� & "] ���������°���:" & vbCrLf & "        " & Mid(strTmp, 2)
                If Not IsNull(rsDoctorPlan!��ʼʱ��) Then
                    strInfo = strInfo & vbCrLf & "        ��Ч��:" & rsDoctorPlan!��ʼʱ�� & "~" & rsDoctorPlan!��ֹʱ��
                Else
                    strInfo = strInfo & vbCrLf & "        ��Ч��:����"
                End If
            End If
            rsDoctorPlan.MoveNext
        Next
        If strInfo <> "" Then
            If blnMulitNumPlan Then
                '��ΰ���ʱ,���ܴ��ڽ���
                Call MsgBox("����" & cboDoctor.Text & "ҽ�������뵱ǰ�ű��ظ��򽻲�ĹҺŰ���" & vbCrLf & strInfo & vbCrLf & vbCrLf & "���ܰ���!", vbInformation + vbOKOnly, gstrSysName)
                Exit Sub
            Else
                If MsgBox("����" & cboDoctor.Text & "ҽ�������뵱ǰ�ű��ظ��򽻲�ĹҺŰ���" & vbCrLf & strInfo & vbCrLf & vbCrLf & "ȷʵҪ���浱ǰ�ű���?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
            End If
        End If
    End If

    If Not mEditType = edt_���� Then
        If CheckExistsBooking() Then
            If MsgBox("�úű�ǰ���ŵ�ʱ���֮�����ԤԼ�Һŵ�,�Ƿ�Ҫ����?", vbInformation + vbOKCancel + vbDefaultButton1, gstrSysName) = vbCancel Then
                Exit Sub
            End If
        End If
    End If
    '�ȼ��
    'ȡʱ���
    str�޺� = ""
    If opt��.Value Then 'ÿ��
        For i = 1 To 7
            strʱ��� = strʱ��� & "'" & Trim(cbo��.Text) & "',"
            str�޺� = str�޺� & "|" & Switch(i = 1, "����", i = 2, "��һ", i = 3, "�ܶ�", i = 4, "����", i = 5, "����", i = 6, "����", True, "����")
            str�޺� = str�޺� & "," & Val(txt�޺�.Text) & "," & IIf(chkAppoint.Value = 0, "0", txt��Լ.Text)
        Next
    Else
        For i = 1 To vsPlan.Cols - 1
            strʱ��� = strʱ��� & "'" & Trim(vsPlan.TextMatrix(1, i)) & "',"
            If Trim(vsPlan.TextMatrix(1, i)) <> "" Then
                str�޺� = str�޺� & "|" & Switch(i = 1, "����", i = 2, "��һ", i = 3, "�ܶ�", i = 4, "����", i = 5, "����", i = 6, "����", True, "����")
                If Trim(vsPlan.TextMatrix(1, i)) = "" Then
                    str�޺� = str�޺� & ",0,0"
                Else
                    str�޺� = str�޺� & "," & Val(Trim(vsPlan.TextMatrix(2, i))) & "," & IIf(chkAppoint.Value = 0, "0", Trim(vsPlan.TextMatrix(3, i)))
                End If
            End If
        Next
    End If
    If str�޺� <> "" Then str�޺� = Mid(str�޺�, 2)
    
    If chkAppoint.Value = 1 Then
        If txtAppLimit.Text <> "" And Val(txtAppLimit.Text) <= 0 Then
            MsgBox "����ԤԼ������£�ԤԼ����������Ҫ��1�죡", vbInformation, gstrSysName
            txtAppLimit.SetFocus: Exit Sub
        End If
        lngԤԼ���� = Val(IIf(txtAppLimit.Text = "", gintԤԼ����, Val(txtAppLimit.Text)))
    Else
        lngԤԼ���� = 0
    End If

    'ȡ�Һ�����
    '71253 ���ϴ� 2014-04-15 11:30:10 ��listView �滻ΪvsflexGrid
    With vsDept
        For i = 0 To .Cols - 1
            For j = 0 To .Rows - 1
                If .Cell(flexcpChecked, j, i) = 1 Then str���� = str���� & ";" & .TextMatrix(j, i)
            Next
        Next
    End With
    str���� = Mid(str����, 2)
    
    'ȡ���﷽ʽ
    intCount = 0
    For i = 0 To opt����.UBound
        If opt����(i).Value Then intCount = i: Exit For
    Next

    'ȡ��ʼʱ�䷶Χ
    strBegin = "NULL": strEnd = "NULL"
    If chk��Ч��.Value = 1 Then
        strBegin = "To_Date('" & Format(dtpBegin.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
        strEnd = "To_Date('" & Format(dtpEnd.Value, "yyyy-MM-dd HH:mm:ss") & "','YYYY-MM-DD HH24:MI:SS')"
    End If

      '�鿴�Ƿ�ı����Ű���� �ı��� �޺��� ��Լ�� ������ſ���
    blnChange = (str�޺� <> mPlanInfo.str�޺�) Or (strʱ��� <> mPlanInfo.str�Ű�)
    blnChange = blnChange Or (chk��ſ���.Value <> IIf(mPlanInfo.bln���, 1, 0))
    str�޺� = "'" & str�޺� & "',"
    Set cllPro = New Collection
    'ȡID
    If mEditType = edt_���� Then

        '����
        lngNextID = zlDatabase.GetNextId("�ҺŰ���")

        strSQL = "zl_�ҺŰ���_INSERT(" & _
            lngNextID & ",'" & Trim(txt�ű�.Text) & "','" & cbo����.Text & "'," & _
            cbo����.ItemData(cbo����.ListIndex) & "," & _
            cboItem.ItemData(cboItem.ListIndex) & ",'" & Trim(cboDoctor.Text) & "'," & _
            lngҽ��ID & "," & _
            chk����.Value & "," & strʱ��� & str�޺� & intCount & "," & _
            "'" & str���� & "'," & strBegin & "," & strEnd & ",1," & chk��ſ���.Value & ",0," & _
            5 & "," & lngԤԼ���� & ")"
    Else
'
' Zl_�ҺŰ���_Insert
'(
'  Id_In       �ҺŰ���.ID%Type,
'  ����_In     �ҺŰ���.����%Type,
'  ����_In     �ҺŰ���.����%Type,
'  ����id_In   �ҺŰ���.����id%Type,
'  ��Ŀid_In   �ҺŰ���.��Ŀid%Type,
'  ҽ��_In     �ҺŰ���.ҽ������%Type,
'  ҽ��id_In   �ҺŰ���.ҽ��id%Type,
'  ��������_In �ҺŰ���.��������%Type,
'  ����_In     �ҺŰ���.����%Type,
'  ��һ_In     �ҺŰ���.��һ%Type,
'  �ܶ�_In     �ҺŰ���.�ܶ�%Type,
'  ����_In     �ҺŰ���.����%Type,
'  ����_In     �ҺŰ���.����%Type,
'  ����_In     �ҺŰ���.����%Type,
'  ����_In     �ҺŰ���.����%Type,
'  �޺ſ���_In Varchar2,
'  ���﷽ʽ_In �ҺŰ���.���﷽ʽ%Type,
'  ����_In     Varchar2,
'  ��ʼʱ��_In �ҺŰ���.��ʼʱ��%Type,
'  ��ֹʱ��_In �ҺŰ���.��ֹʱ��%Type,
'  ����_In     Number,
'  ��ſ���_In �ҺŰ���.��ſ���%Type,
'  ��������_In Number:=0,
'  Ĭ��ʱ�μ��_In �ҺŰ���.Ĭ��ʱ�μ��%Type
') As
'  -----------------------------------------------------------
'  --������
'  --  ����_IN=��';'�ŷָ��Ķ����������
'  --  �޺ſ���_IN:|��һ,22(�޺�),13(��Լ)|�ܶ�,20(�޺�),11(��Լ)....
'  --  ��������_IN:�޸İ���ʱ ��ʱ�����ݵĴ��� 0--������ 1--ɾ��ʱ����Ϣ
        '�޸�

        lngNextID = mlngID
        strSQL = "    " & vbNewLine & "zl_�ҺŰ���_INSERT("
        strSQL = strSQL & vbNewLine & lngNextID
        strSQL = strSQL & vbNewLine & ",'" & (txt�ű�.Text) & "','" & cbo����.Text & "',"
        strSQL = strSQL & vbNewLine & cbo����.ItemData(cbo����.ListIndex) & ","
        strSQL = strSQL & vbNewLine & cboItem.ItemData(cboItem.ListIndex) & ",'" & Trim(cboDoctor.Text) & "',"
        strSQL = strSQL & vbNewLine & lngҽ��ID & "," & chk����.Value & ","
        strSQL = strSQL & vbNewLine & strʱ��� & str�޺� & intCount & ","
        strSQL = strSQL & vbNewLine & "'" & str���� & "'," & strBegin & "," & strEnd & ",0," & chk��ſ���.Value & ","
        strSQL = strSQL & vbNewLine & IIf(chk��ſ���.Value <> IIf(mPlanInfo.bln���, 1, 0), 1, 0) & ","
        strSQL = strSQL & vbNewLine & 5 & "," & lngԤԼ���� & ")"

    End If

    On Error GoTo errH
    zlAddArray cllPro, strSQL

    LoadTimePlan True
    If SaveTimeSetData(lngNextID, cllPro) = False Then Exit Sub
    
    On Error GoTo Errhand
    zlExecuteProcedureArrAy cllPro, Me.Caption
    On Error GoTo 0
    mblnSucces = True

    If mEditType <> edt_���� Then Unload Me: Exit Sub
    Call LoadCard
    mblnChangeByCode = True
    tbPage.Item(mPageIndex.EM_����).Selected = True
    mblnChangeByCode = False
    Call ClearCustomData
    Exit Sub
Errhand:
    gcnOracle.RollbackTrans
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ClearCustomData()
     mTimeSet.str���� = ""
     mTimeSet.bln��ſ��� = False
     mTimeSet.lngSelIndex = 0
     mTimeSet.blnOnChange = False
     mTimeSet.lng����ID = 0
     mTimeSet.lng�ƻ�ID = 0
     mTimeSet.blnIsInit = False
     Set mrs�޺� = Nothing
     Set mTimeSet.rsRegPlan = Nothing
     Set mTimeSet.rsAssign = Nothing
     mTimeSet.strKey = ""
     mTimeSet.blnChange = False
     mTimeSet.str�����޸� = ""
     Set mTimeSet.rsHistory = Nothing
End Sub

Private Function zl_CheckMoveAssign(Optional ByVal lngIndex As Long = -1) As Boolean
    Dim str������Ŀ As String
    If lngIndex = -1 Then lngIndex = mTimeSet.lngSelIndex
    If lngIndex = -1 Then zl_CheckMoveAssign = True: Exit Function
    If Not mTimeSet.blnChange Then zl_CheckMoveAssign = True: Exit Function
    
    If lngIndex < 0 Or lngIndex > 6 Then Exit Function
    If Not VsTimeValidate(lngIndex) Then Exit Function
    
    str������Ŀ = GetVsGridCaption(lngIndex)
    zl_CheckMoveAssign = MoveAssign(str������Ŀ)
End Function

Private Function GetVsGridIndex(ByVal str���� As String) As Integer
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ȡ�������
    '����:���˺�
    '����:2012-06-15 14:03:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    str���� = Switch(str���� = "����", 0, str���� = "��һ", 1, str���� = "�ܶ�", 2, str���� = "����", 3, str���� = "����", 4, str���� = "����", 5, str���� = "����", 6, True, 0)
    GetVsGridIndex = Val(str����)
End Function

Private Function MoveAssign(ByVal str������Ŀ As String) As Boolean
    '�����������ŵ����ݼ���
    Dim nIndex As Long
    Dim lng��� As Long
    Dim i As Long, j As Long
    Dim str��ʼʱ�� As String
    Dim str����ʱ�� As String
    Dim lng���� As Long
    Dim blnԤԼ As Boolean
    Dim str���ʱ�� As String
    If Not mTimeSet.blnChange Then MoveAssign = True: Exit Function
    
    nIndex = GetVsGridIndex(str������Ŀ)
    
    'ɾ��û��ʹ�ò���
    mTimeSet.rsAssign.Filter = "������Ŀ='" & str������Ŀ & "' and ��ʹ��=0"
    If mTimeSet.rsAssign.RecordCount > 0 Then
        Do While Not mTimeSet.rsAssign.EOF
            mTimeSet.rsAssign.Delete
            mTimeSet.rsAssign.MoveNext
        Loop
    End If
    
    If Not mTimeSet.bln��ſ��� Then
        With vsTime(nIndex)
          lng��� = 0
            For i = 1 To .Rows - 1
                For j = 0 To .Cols - 1 Step 2
                   If .TextMatrix(i, j) <> "" Then
                    
                    str��ʼʱ�� = Split(.TextMatrix(i, j), "-")(0)
                    str����ʱ�� = Split(.TextMatrix(i, j), "-")(1)
                    lng���� = Val(.TextMatrix(i, j + 1))
                    lng��� = lng��� + 1
                    blnԤԼ = True
                    
                    str���ʱ�� = ""
                    If Not mTimeSet.rsHistory Is Nothing Then
                        mTimeSet.rsHistory.Filter = "������Ŀ='" & str������Ŀ & "'"
                        If mTimeSet.rsHistory.RecordCount = 0 Then
                            str���ʱ�� = ""
                            mTimeSet.rsHistory.Filter = 0
                        Else
                            str���ʱ�� = Nvl(mTimeSet.rsHistory!����ʱ��)
                            mTimeSet.rsHistory.Filter = 0
                        End If
                    End If
                    
                    If (str���ʱ�� <> "" And str��ʼʱ�� > str���ʱ��) Or str���ʱ�� = "" Then
                        With mTimeSet.rsAssign
                            .AddNew
                            !������Ŀ = str������Ŀ
                            !��ʼʱ�� = str��ʼʱ��
                            !����ʱ�� = str����ʱ��
                            !ʱ��� = str��ʼʱ�� & "-" & str����ʱ��
                            !�������� = lng����
                            !��� = lng���
                            !��ʹ�� = 0
                            !�Ƿ�ԤԼ = 1
                            .Update
                        End With
                    End If
                   End If
                Next
            Next
        End With
        mTimeSet.blnChange = False
        MoveAssign = True
        Exit Function
    End If
    
    
    '��ſ���
    
    With vsTime(nIndex)
        For i = 0 To .Rows - 1 Step 2
            For j = 1 To .Cols - 1
                If Trim(.TextMatrix(i, j)) <> "" Then
                        str��ʼʱ�� = Split(.TextMatrix(i + 1, j) & "-", "-")(0)
                        str����ʱ�� = Split(.TextMatrix(i + 1, j) & "-", "-")(1)
                        lng��� = Val(.TextMatrix(i, j))
                        lng���� = 1
                        blnԤԼ = .Cell(flexcpForeColor, i, j) = vbBlue
                    If .Cell(flexcpFontUnderline, i, j) = False Then
                       
                        With mTimeSet.rsAssign
                            .AddNew
                            !������Ŀ = str������Ŀ
                            !��ʼʱ�� = str��ʼʱ��
                            !����ʱ�� = str����ʱ��
                            !ʱ�� = Format(str��ʼʱ��, "hh:00:00")
                            !ʱ��� = str��ʼʱ�� & "-" & str����ʱ��
                            !�������� = lng����
                            !��� = lng���
                            !��ʹ�� = 0
                            !�Ƿ�ԤԼ = IIf(blnԤԼ, 1, 0)
                            .Update
                        End With
                    ElseIf .Cell(flexcpFontUnderline, i, j) Then
                        ' �̶�����Ϣ,���ܸı��Ƿ�ԤԼ,����Ҳֻ�ɸı��Ƿ�ԤԼ
                        With mTimeSet.rsAssign
                            .Filter = "���=" & lng��� & " And ��ʼʱ��='" & Format(str��ʼʱ��, "hh:mm:00") & "'"
                            If .RecordCount > 0 Then
                                !�Ƿ�ԤԼ = IIf(blnԤԼ, 1, 0)
                                .Update
                            End If
                        End With
                    End If
                End If
            Next
        Next
    End With
    mTimeSet.blnChange = False
    MoveAssign = True
    Exit Function
End Function

 Private Function GetVsGridCaption(ByVal nIndex As Integer) As String
    '����:����������ȡ������Ŀ
    Dim str���� As String
    str���� = Switch(nIndex = 0, "����", nIndex = 1, "��һ", nIndex = 2, "�ܶ�", nIndex = 3, "����", nIndex = 4, "����", nIndex = 5, "����", nIndex = 6, "����", True, "")
    GetVsGridCaption = str����
 End Function

Private Function VsTimeValidate(ByVal lngIndex As Long) As Boolean
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��֤���õ���Լ���Ƿ����Ҫ��
    '���:lngIndex-ָ����ҳ��(���ڶ�Ӧ������):-1ʱ,��ʾ�����е�ҳ����м��
    '����:
    '����:У�Գɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-11-15 10:17:37
    '˵��:
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim lngStep As Long, i As Long, j  As Long
    Dim lngԤԼ��   As Long, lng�޺��� As Long, lng��Լ�� As Long, lng���� As Long
    Dim str����   As String, str������Ŀ As String
    Dim lngPage As Long, lngPages As Long, lngStartPage As Long
    Dim blnNotSetTime As Boolean '��������ʱ���
    Dim blnAllowNums As Boolean '�����޺�����һ��
    Dim blnAllowYYNums As Boolean '����ԤԼ�������õ�ԤԼ����һ��
    Dim strCommand As String, blnʱ�� As Boolean '�ж�������ʱ�ε�,��Ҫ�������ʱ��ҳ�Ƿ�����
    On Error GoTo errHandle
        
    lngStartPage = 0: lngPages = tbSubPage.ItemCount - 1
    If lngIndex <> -1 Then lngStartPage = lngIndex: lngPages = lngIndex
    blnʱ�� = False
    For lngPage = lngStartPage To lngPages
        If mTimeSet.bln��ſ��� Then
            With vsTime(lngPage)
                For i = 0 To .Rows - 1 Step 2
                    For j = 1 To .Cols - 1
                       If .TextMatrix(i, j) <> "" Then
                           blnʱ�� = True
                           Exit For
                       End If
                    Next
                Next
            End With
        Else
                With vsTime(lngPage)
                    For i = 1 To .Rows - 1
                        For j = 1 To .Cols - 1 Step 2
                            If .TextMatrix(i, j) <> "" Then
                               blnʱ�� = True: Exit For
                            End If
                        Next
                    Next
                End With
        End If
    Next
    'δ����ʱ��
    If blnʱ�� = False Then VsTimeValidate = True: Exit Function
    
    For lngPage = lngStartPage To lngPages
        tbSubPage(lngPage).Selected = True
        str������Ŀ = GetVsGridCaption(lngPage)
        mTimeSet.rsRegPlan.Filter = "������Ŀ='" & str������Ŀ & "'"
        If mTimeSet.rsRegPlan.RecordCount = 0 Then
            mTimeSet.rsRegPlan.Filter = 0
        Else
                lng�޺��� = Val(Nvl(mTimeSet.rsRegPlan!�޺���)): lng��Լ�� = Val(Nvl(mTimeSet.rsRegPlan!��Լ��))
                If lng��Լ�� = 0 Then lng��Լ�� = lng�޺���
                lng���� = 0: lngԤԼ�� = 0
                
                If mTimeSet.bln��ſ��� Then
                    'ר�Һż����Լ���Ƿ�����޺���
                    With vsTime(lngPage)
                        For i = 0 To .Rows - 1 Step 2
                            For j = 1 To .Cols - 1
                               If .TextMatrix(i, j) <> "" Then
                                     If .Cell(flexcpForeColor, i, j, i, j) = vbBlue Then
                                         lngԤԼ�� = lngԤԼ�� + 1
                                     End If
                                     lng���� = lng���� + 1
                               End If
                            Next
                        Next
                    End With
                    If lng���� < lng�޺��� Then
                        If lng���� = 0 Then
                           If lngIndex = -1 Then
                                If blnNotSetTime = False And blnʱ�� Then
                                        strCommand = zlCommFun.ShowMsgbox("����", "    �ڷ�ʱ��ҳ����δ���á�" & str������Ŀ & "����ʱ��,��ȷ��������ʱ���?" & vbCrLf & vbCrLf & _
                                         "���ǡ�:��ʾ��������ʱ��ν��б���" & vbCrLf & vbCrLf & _
                                         "�����ԡ�:��ʾ�������Ƶ�δ����ʱ��ε�����������,��������ʾ��" & vbCrLf & vbCrLf & _
                                         "����:��ʾ����������ʱ���,������������" & vbCrLf, "��(&O),����(&I),��(&C)", Me, vbQuestion)
                                        Select Case strCommand
                                        Case "��"
                                        Case "����"
                                             blnNotSetTime = True
                                         Case Else
                                            Call zlSaveTimePageSelected(str������Ŀ)
                                            mTimeSet.blnNotBrush = True
                                            tbSubPage.Item(lngPage).Selected = True
                                            If vsTime(lngPage).Enabled And vsTime(lngPage).Visible Then vsTime(lngPage).SetFocus
                                            mTimeSet.blnNotBrush = False
                                            Exit Function
                                         End Select
                                End If
'                           Else
'                                If MsgBox("�ڷ�ʱ��ҳ����δ���á�" & str������Ŀ & "����ʱ��,��ȷ��������ʱ���?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then
'                                    If lngIndex = -1 Then
'                                        Call zlSaveTimePageSelected(str������Ŀ)
'                                        mTimeSet.blnNotBrush = True
'                                        tbSubPage.Item(lngPage).Selected = True
'                                        If vsTime(lngPage).Enabled And vsTime(lngPage).Visible Then vsTime(lngPage).SetFocus
'                                        mTimeSet.blnNotBrush = False
'                                    End If
'                                    Exit Function
'                                End If
                            End If
                        Else
                                If lngIndex = -1 Then
                                     If blnAllowNums = False Then
                                             strCommand = zlCommFun.ShowMsgbox("����", "    �ڷ�ʱ��ҳ���еġ�" & str������Ŀ & "��������ʱ��εĺ���(" & lng���� & ")���޺���(" & lng�޺��� & ") ����,��ȷ������ǰ���õ�ʱ�α���?" & vbCrLf & vbCrLf & _
                                              "���ǡ�:��ʾ�����޺����������һ��" & vbCrLf & vbCrLf & _
                                              "�����ԡ�:��ʾ�����޺����������һ�£��������Ƶ�����,������ʾ��" & vbCrLf & vbCrLf & _
                                              "����:��ʾ�������޺����������һ��,������������" & vbCrLf, "��(&O),����(&I),��(&C)", Me, vbQuestion)
                                             Select Case strCommand
                                              Case "��"
                                              Case "����"
                                                  blnAllowNums = True
                                              Case Else
                                                 Call zlSaveTimePageSelected(str������Ŀ)
                                                 mTimeSet.blnNotBrush = True
                                                 tbSubPage.Item(lngPage).Selected = True
                                                 If vsTime(lngPage).Enabled And vsTime(lngPage).Visible Then vsTime(lngPage).SetFocus
                                                 mTimeSet.blnNotBrush = False
                                                 Exit Function
                                              End Select
                                     End If
'                                Else
'                                     If MsgBox("�ڷ�ʱ��ҳ���еġ�" & str������Ŀ & "��������ʱ��εĺ���(" & lng���� & ")���޺���(" & lng��Լ�� & ") ����,��ȷ������ǰ���õ�ʱ�α���?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                                End If
                        End If
                    ElseIf lng���� > lng�޺��� Then
                        Call MsgBox("�ڷ�ʱ��ҳ���еġ�" & str������Ŀ & "��������ʱ��εĺ���(" & lng���� & ")�������޺���(" & lng��Լ�� & ") �㲻�ܰ���ǰ���õ�ʱ�α���!", vbQuestion + vbOKOnly + vbDefaultButton2, gstrSysName)
                        If lngIndex = -1 Then
                            Call zlSaveTimePageSelected(str������Ŀ)
                            mTimeSet.blnNotBrush = True
                            tbSubPage.Item(lngPage).Selected = True
                            If vsTime(lngPage).Enabled And vsTime(lngPage).Visible Then vsTime(lngPage).SetFocus
                            mTimeSet.blnNotBrush = False
                        End If
                        Exit Function
                    End If
                Else
                     '��ͨ�ż����Լ���Ƿ�����޺���
                    With vsTime(lngPage)
                        For i = 1 To .Rows - 1
                            For j = 1 To .Cols - 1 Step 2
                                If .TextMatrix(i, j) <> "" Then
                                    lngԤԼ�� = lngԤԼ�� + Val(.TextMatrix(i, j))
                                End If
                            Next
                        Next
                    End With
                End If
                If lngԤԼ�� > lng��Լ�� Then
                   MsgBox "�ڷ�ʱ��ҳ���еġ�" & str������Ŀ & "�������õ�ԤԼ��(" & lngԤԼ�� & ")������" & IIf(lng�޺��� = lng��Լ��, "�޺���(" & lng��Լ�� & ")", "��Լ��(" & lng��Լ�� & ")") & ",�㲻�ܰ���ǰ���ñ���!", vbOKOnly, Me.Caption
                    If lngIndex = -1 Then
                        Call zlSaveTimePageSelected(str������Ŀ)
                        mTimeSet.blnNotBrush = True
                        tbSubPage.Item(lngPage).Selected = True
                        If vsTime(lngPage).Enabled And vsTime(lngPage).Visible Then vsTime(lngPage).SetFocus
                        mTimeSet.blnNotBrush = False
                    End If
                   Exit Function
                End If
                If lngԤԼ�� < lng��Լ�� And lngԤԼ�� <> 0 Then
                    If lngIndex = -1 Then
                           If blnAllowYYNums = False Then
                                   strCommand = zlCommFun.ShowMsgbox("����", "    �ڷ�ʱ��ҳ���еġ�" & str������Ŀ & "�������õ�ʵ��ԤԼ��(" & lngԤԼ�� & ") ����Լ��(" & lng��Լ�� & ") ����,��ȷ������ǰ���õ�ʱ�α���?" & vbCrLf & vbCrLf & _
                                    "���ǡ�:��ʾ������Լ����ԤԼ����һ��" & vbCrLf & vbCrLf & _
                                    "�����ԡ�:��ʾ������Լ����ԤԼ����һ�£��������Ƶ�����,������ʾ��" & vbCrLf & vbCrLf & _
                                    "����:��ʾ��������Լ����ԤԼ����һ��,������������" & vbCrLf, "��(&O),����(&I),��(&C)", Me, vbQuestion)
                                    Select Case strCommand
                                    Case "��"
                                    Case "����"
                                        blnAllowYYNums = True
                                    Case Else
                                       Call zlSaveTimePageSelected(str������Ŀ)
                                       mTimeSet.blnNotBrush = True
                                       tbSubPage.Item(lngPage).Selected = True
                                       If vsTime(lngPage).Enabled And vsTime(lngPage).Visible Then vsTime(lngPage).SetFocus
                                       mTimeSet.blnNotBrush = False
                                        Exit Function
                                    End Select
                           End If
'                      Else
'                            If MsgBox("�ڷ�ʱ��ҳ���еġ�" & str������Ŀ & "�������õ�ʵ��ԤԼ��(" & lngԤԼ�� & ") ����Լ��(" & lng��Լ�� & ") ����,��ȷ������ǰ���õ�ʱ�α���?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
                    End If
                End If
        End If
    Next
    VsTimeValidate = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub zlSaveTimePageSelected(ByVal str���� As String)
    If tbPage.Selected Is Nothing Then Exit Sub
    If tbPage.Selected.index <> mPageIndex.EM_ʱ�� Then
         tbPage.Item(mPageIndex.EM_ʱ��).Selected = True
    End If
End Sub


Private Sub txtAppLimit_Validate(Cancel As Boolean)
    If chkAppoint.Value = 1 And txtAppLimit.Text <> "" And Val(txtAppLimit.Text) <= 0 Then
        MsgBox "����ԤԼʱ,ԤԼ��������С�ڵ���0!", vbInformation, gstrSysName
        If txtAppLimit.Visible And txtAppLimit.Enabled Then txtAppLimit.SetFocus
        Cancel = True
    End If
End Sub

Private Sub vsTime_ValidateEdit(index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strʱ��() As String
     If mTimeSet.bln��ſ��� Then
        strʱ�� = Split(vsTime(index).EditText, "-")
        If UBound(strʱ��) <> 1 Then
           MsgBox "�����ʱ���ʽ����!����!", vbOKOnly, gstrSysName
           Cancel = True: Exit Sub
        End If
        If Not IsDate(strʱ��(0)) Then
           MsgBox "�����ʱ���ʽ����!����!", vbOKOnly, gstrSysName
           Cancel = True: Exit Sub
        End If
        If Not IsDate(strʱ��(1)) Then
           MsgBox "�����ʱ���ʽ����!����!", vbOKOnly, gstrSysName
           Cancel = True: Exit Sub
        End If
        If CDate(strʱ��(0)) >= CDate(strʱ��(1)) Then
           MsgBox "��ʼʱ�����С�ڽ���ʱ��!����!", vbOKOnly, gstrSysName
           Cancel = True
        End If
     End If
    mTimeSet.blnChange = True
End Sub

Private Sub picTimeSet_Resize()
    Err = 0: On Error Resume Next
    With fraӦ����
        .Top = picTimeSet.ScaleHeight - .Height - 50
        .Width = picTimeSet.ScaleWidth
        .Left = picTimeSet.ScaleLeft
        .Visible = True
    End With
    With tbSubPage
        .Top = txtTimeOut.Top + txtTimeOut.Height + 50
        .Left = picTimeSet.ScaleLeft
        .Width = picTimeSet.ScaleWidth
        .Height = fraӦ����.Top - .Top - 100
    End With
End Sub

Private Sub cmdOther_Click()
    Dim str���� As String
    If Not mTimeSet.bln��ſ��� Then Exit Sub
    Set mfrmOtherCalc = New frmRegistPlanTimeOther
    Call mfrmOtherCalc.zlShowMe(Me, Nvl(mTimeSet.rsRegPlan!�Ű�), Val(txtTimeOut.Text))
    If Not mfrmOtherCalc Is Nothing Then Unload mfrmOtherCalc
    Set mfrmOtherCalc = Nothing '
End Sub

Private Sub vsTime_BeforeRowColChange(index As Integer, ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long, Cancel As Boolean)
    If Not mTimeSet.bln��ſ��� Then
        vsTime(index).Editable = IIf(NewCol Mod 2 = 1, flexEDKbd, flexEDNone)
          cmdԤԼ(index).Visible = False: Exit Sub
    End If
    If NewRow < 0 Or NewCol < 0 Then Exit Sub
    
    SetCtrlMove index, NewRow - (NewRow) Mod 2, NewCol
    If mTimeSet.bln��ſ��� Then
        If vsTime(index).Cell(flexcpFontUnderline, NewRow, NewCol) = False And vsTime(index).Cell(flexcpBackColor, NewRow, NewCol) = 0 Then
            vsTime(index).Editable = flexEDKbdMouse
        Else
            vsTime(index).Editable = flexEDNone
        End If
        Exit Sub
    End If
    
    With vsTime(index)
        .Editable = IIf(NewCol Mod 2 = 1, flexEDKbd, flexEDNone)
    End With
End Sub

Private Sub vsTime_KeyDown(index As Integer, KeyCode As Integer, Shift As Integer)
    If Not mTimeSet.bln��ſ��� Then Exit Sub
     
    With vsTime(index)
           
        If (.Row < 0 Or .Col < 1) Or (.Row > .Rows - 1 Or .Col > .Cols - 1) Then Exit Sub 'û����Ч��Ԫ����
        If Trim(.TextMatrix(.Row, .Col)) = "" Then Exit Sub
        If KeyCode = 13 Then
            Call cmdԤԼ_Click(index)
            Exit Sub
        End If
        
        If KeyCode = 46 Then
            If cmdɾ��(index).Visible = False Then Exit Sub
            If Trim(.TextMatrix(.Row, .Col)) = "" Then Exit Sub
            Call cmdɾ��_Click(index)
        End If
     End With
End Sub

Private Sub vsTime_KeyPressEdit(index As Integer, ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
     If Not ((KeyAscii >= Asc(0) And KeyAscii <= Asc(9)) Or KeyAscii = 8 _
               Or KeyAscii = 13 Or KeyAscii = Asc("-") Or KeyAscii = Asc(":")) Then KeyAscii = 0: Exit Sub
End Sub

Private Sub vsTime_LostFocus(index As Integer)
 If Trim(vsTime(index).EditText) <> "" Then
    With vsTime(index)
        .TextMatrix(.Row, .Col) = .EditText
        mTimeSet.blnChange = True
    End With
 End If
End Sub

Private Sub SetCtrlMove(ByVal index As Integer, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim blnDel As Boolean
    With vsTime(index)
        If mTimeSet.bln��ſ��� Then
            If Trim(.TextMatrix(NewRow, NewCol)) = "" Then
                cmdɾ��(index).Visible = False
                cmdԤԼ(index).Visible = False
                Exit Sub
            End If
            cmdɾ��(index).Left = .Cell(flexcpLeft, NewRow, NewCol) + .Cell(flexcpWidth, NewRow, NewCol) - cmdɾ��(index).Width
            If .Row Mod 2 <> 0 Then
                cmdɾ��(index).Top = .Cell(flexcpTop, NewRow, NewCol)
            Else
                cmdɾ��(index).Top = .Cell(flexcpTop, NewRow, NewCol)
            End If
            cmdԤԼ(index).Left = .Cell(flexcpLeft, NewRow, NewCol)
            cmdԤԼ(index).Top = cmdɾ��(index).Top
            If NewCol < .Cols - 1 Then
                blnDel = Trim(.TextMatrix(NewRow, NewCol + 1)) = ""
            Else
                blnDel = True
            End If
             
            blnDel = blnDel And Trim(.TextMatrix(NewRow, NewCol)) <> "" And Not .Cell(flexcpFontUnderline, NewRow, NewCol)
            cmdɾ��(index).Visible = blnDel And mTimeSet.bln��ſ���
            cmdԤԼ(index).Visible = True 'Val(txt��Լ.Text) <> 0
        Else
            cmdԤԼ(index).Left = .Cell(flexcpTop, NewRow, NewCol)
            cmdԤԼ(index).Top = .Cell(flexcpLeft, NewRow, NewCol)
            cmdԤԼ(index).Visible = False
        End If
    End With
End Sub

Private Sub cmdԤԼ_Click(index As Integer)
    Dim i As Integer, j As Integer
    Dim intStartRow As Integer, intEndRow As Integer, intStartCol As Integer, intEndCol As Integer
    If Not mTimeSet.bln��ſ��� Or mTimeSet.lngSelIndex < 0 Then Exit Sub
    If chkAppoint.Value = 0 Then Exit Sub
    If mTimeSet.lngSelIndex <> index Then Exit Sub
    With vsTime(mTimeSet.lngSelIndex)
'        If .MouseRow < 0 Or .MouseCol < 0 Then Exit Sub
        If .Row < 0 Or .Col < 0 Then Exit Sub
        If .Row > .RowSel Then
            intStartRow = .RowSel
            intEndRow = .Row
        Else
            intStartRow = .Row
            intEndRow = .RowSel
        End If
        If .Col > .ColSel Then
            intStartCol = .ColSel
            intEndCol = .Col
        Else
            intStartCol = .Col
            intEndCol = .ColSel
        End If
        For i = intStartRow To intEndRow Step 2
            For j = intStartCol To intEndCol
                If i <= .Rows - 1 And j <= .Cols - 1 Then
                    If .Cell(flexcpForeColor, i, j) = vbBlue Then
                       .Cell(flexcpForeColor, i - (i Mod 2), j, i + (i + 1) Mod 2, j) = &H80000008
                        .Cell(flexcpFontBold, i - (i Mod 2), j, i + (i + 1) Mod 2, j) = False
                    Else
                        .Cell(flexcpForeColor, i - (i Mod 2), j, i + (i + 1) Mod 2, j) = vbBlue
                        .Cell(flexcpFontBold, i - (i Mod 2), j, i + (i + 1) Mod 2, j) = True
                    End If
                End If
            Next j
        Next i
        mTimeSet.blnChange = True
        .SetFocus
    End With
End Sub

Private Sub cmdɾ��_Click(index As Integer)
    Dim blnDel As Boolean
    Dim lngSelX As Long
    Dim lngSelY As Long
    Dim i As Long, j As Long
    Dim lngCurrSn As Long
    Dim lngStartCol As Long
    With vsTime(index)
        If .Col < .Cols - 1 Then
                blnDel = Trim(.TextMatrix(.Row, .Col + 1)) = ""
        Else
                blnDel = True
        End If
        blnDel = blnDel And Trim(.TextMatrix(.Row, .Col)) <> "" And Not .Cell(flexcpFontUnderline, .Row, .Col)
        If Not blnDel Then Exit Sub
        If mTimeSet.bln��ſ��� Then
          lngSelX = .Row - (.Row Mod 2): lngSelY = .Col
          lngCurrSn = Val(.TextMatrix(lngSelX, lngSelY))
          .TextMatrix(lngSelX, lngSelY) = ""
          .TextMatrix(lngSelX + 1, lngSelY) = ""
          
          For i = lngSelX To .Rows - 1 Step 2
            lngStartCol = 1
            If i = lngSelX Then lngStartCol = lngSelY
            For j = lngStartCol To .Cols - 1
                If .TextMatrix(i, j) <> "" Then
                    .TextMatrix(i, j) = lngCurrSn
                     lngCurrSn = lngCurrSn + 1
                End If
            Next
         Next
        End If
        cmdɾ��(index).Visible = False
        cmdԤԼ(index).Visible = False
        mTimeSet.blnChange = True
        .SetFocus
    End With
End Sub

Private Sub picPage_Resize(index As Integer)
    Err = 0: On Error Resume Next
    With picPage(index)
        vsTime(index).Left = .ScaleLeft
        vsTime(index).Top = .ScaleTop
        vsTime(index).Width = .ScaleWidth
        vsTime(index).Height = .ScaleHeight
    End With
End Sub

Private Sub Initʱ���()
  '--------------------------------
  '����:��ȡ���°�ʱ���
  '--------------------------------
    Dim strTmp      As String
    Dim strSQL      As String
    Dim rsTmp       As ADODB.Recordset
    Dim strDat      As String
    On Error GoTo Hd
    strTmp = zlDatabase.GetPara("�������°�ʱ��", glngSys, , "07:00:00 AND 12:00:00")
    strDat = Split(strTmp, "AND")(0)
    If IsDate(strDat) Then
        t_ʱ��.dat_�����ϰ� = CDate("1900-01-01 " & Format(CDate(strDat), "hh:mm:ss"))
    Else
        t_ʱ��.dat_�����ϰ� = CDate("08:00:00")
    End If
    
    strDat = Split(strTmp, "AND")(1)
    If IsDate(strDat) Then
        t_ʱ��.dat_�����°� = CDate("1900-01-01 " & Format(CDate(strDat), "hh:mm:ss"))
    Else
        t_ʱ��.dat_�����°� = CDate("1900-01-01 12:00:00")
    End If
    strTmp = zlDatabase.GetPara("�������°�ʱ��", glngSys, , "14:00:00 AND 18:00:00")
    
     strDat = Split(strTmp, "AND")(0)
    If IsDate(strDat) Then
        t_ʱ��.dat_�����ϰ� = CDate("1900-01-01 " & Format(CDate(strDat), "hh:mm:ss"))
    Else
        t_ʱ��.dat_�����ϰ� = CDate("1900-01-01 14:00:00")
    End If
    strDat = Split(strTmp, "AND")(1)
    If IsDate(strDat) Then
        t_ʱ��.dat_�����°� = CDate("1900-01-01 " & Format(CDate(strDat), "hh:mm:ss"))
    Else
        t_ʱ��.dat_�����°� = CDate("1900-01-01 18:00:00")
    End If
    With t_ʱ��
         If .dat_�����ϰ� > .dat_�����°� Then
            .dat_�����°� = DateAdd("d", 1, .dat_�����°�)
         End If
         If .dat_�����ϰ� > .dat_�����°� Then
            .dat_�����°� = DateAdd("d", 1, .dat_�����°�)
         End If
    End With
    strSQL = _
    "       Select ʱ���, �ϰ�, �°� " & vbNewLine & _
    "       From (" & vbNewLine & _
    "           With Tb As (Select ʱ���,To_Date('1900-01-01 ' || To_Char(��ʼʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd HH24:mi:ss') As ��ʼʱ��," & vbNewLine & _
    "                               To_Date(Decode(Sign(��ʼʱ�� - ��ֹʱ��), -1, '1900-01-01 ', '1900-01-02 ') ||To_Char(��ֹʱ��, 'hh24:mi:ss'), 'yyyy-mm-dd HH24:mi:ss') As ��ֹʱ��," & _
    "                               Sign(��ʼʱ�� - ��ֹʱ��) As ����, " & vbNewLine & _
    "                                To_Date('" & Format(t_ʱ��.dat_�����ϰ�, "yyyy-MM-dd hh:mm:ss") & "', 'yyyy-mm-dd HH24:mi:ss') As �����ϰ�ʱ��, " & vbNewLine & _
    "                                To_Date('" & Format(t_ʱ��.dat_�����°�, "yyyy-MM-dd hh:mm:ss") & "', 'yyyy-mm-dd HH24:mi:ss') As �����°�ʱ��, " & vbNewLine & _
    "                                 To_Date('" & Format(t_ʱ��.dat_�����ϰ�, "yyyy-MM-dd hh:mm:ss") & "', 'yyyy-mm-dd HH24:mi:ss') As �����ϰ�ʱ��," & vbNewLine & _
    "                                 To_Date('" & Format(t_ʱ��.dat_�����°�, "yyyy-MM-dd hh:mm:ss") & "', 'yyyy-mm-dd HH24:mi:ss') As �����°�ʱ��,����Ԥ��ʱ�� As Ԥ��ʱ�� "
    strSQL = strSQL & vbNewLine & _
    "                       From ʱ��� )" & vbNewLine & _
    "           Select ʱ���, '��' As ��ǩ, 0 As ��־, ��ʼʱ�� As �ϰ�, ��ֹʱ�� - Nvl(Ԥ��ʱ��, 0) / 24 / 60 As �°�, ��ʼʱ��, ��ֹʱ��," & _
    "                  �����ϰ�ʱ�� As �ϰ�ʱ��, �����°�ʱ�� As �°�ʱ��" & vbNewLine & _
    "            From Tb  Where (��ʼʱ�� >= �����°�ʱ�� Or ��ֹʱ�� <= �����ϰ�ʱ��) And " & _
    "                      (��ʼʱ�� >= �����°�ʱ�� Or ��ֹʱ�� <= �����ϰ�ʱ��) " & vbNewLine & _
    "           Union All" & vbNewLine & _
    "           Select ʱ���, '��-����' As ��ǩ, 1 As ��־, Decode(Sign(�����ϰ�ʱ�� - ��ʼʱ��), 1, �����ϰ�ʱ��, ��ʼʱ��) As �ϰ�, " & vbNewLine & _
    "                        Decode(Sign(��ֹʱ�� - �����°�ʱ��), 1, �����°�ʱ��, ��ֹʱ��) - Nvl(Ԥ��ʱ��, 0) / 24 / 60 As �°�, ��ʼʱ��, ��ֹʱ��, " & _
    "                        �����ϰ�ʱ�� As �ϰ�ʱ��, �����°�ʱ�� As �°�ʱ�� " & vbNewLine & _
    "           From Tb a Where ʱ��� Not In (Select ʱ��� From Tb Where ��ʼʱ�� >= �����°�ʱ�� Or ��ֹʱ�� <= �����ϰ�ʱ��) " & vbNewLine & _
    "           Union All " & vbNewLine & _
    "            Select ʱ���, '��-����' As ��ǩ, 1 As ��־, Decode(Sign(�����ϰ�ʱ�� - ��ʼʱ��), 1, �����ϰ�ʱ��, ��ʼʱ��) As �ϰ�, " & _
    "                   Decode(Sign(��ֹʱ�� - �����°�ʱ��), 1, �����°�ʱ��, ��ֹʱ��) - Nvl(Ԥ��ʱ��, 0) / 24 / 60 As �°�, ��ʼʱ��, ��ֹʱ��, �����ϰ�ʱ�� As �ϰ�ʱ��, �����°�ʱ�� As �°�ʱ�� " & vbNewLine & _
    "         From Tb a   Where ʱ��� Not In (Select ʱ��� From Tb Where ��ʼʱ�� >= �����°�ʱ�� Or ��ֹʱ�� <= �����ϰ�ʱ��)" & vbNewLine & _
    "            ) b" & vbNewLine & _
    "         Order By ʱ���,�ϰ�"
     Set mrs�ϰ�ʱ��� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    Exit Sub
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Sub


Private Function AssignReapportion(ByVal lng���ʱ�� As Long, ByVal str������Ŀ As String) As Boolean
    Dim strSQL As String
    Dim i As Long
    Dim j As Long
    Dim lng�޺� As Long
    Dim lng��Լ As Long
    Dim dat��ʼʱ�� As Date
    Dim dat����ʱ�� As Date
    Dim lng��� As Long
    Dim strTmp As String
    Dim strʱ�� As String
    Dim str����ʱ�� As String
    Dim lngĬ�ϼ�� As Long
    Dim lng������� As Long
    Dim lng�̶����� As Long
    Dim lngTmp As Long
    Dim blnExit As Boolean
    Dim datʱ�� As Date
    If mrs�ϰ�ʱ��� Is Nothing Then
        Call Initʱ���
    End If

    If mrs�ϰ�ʱ��� Is Nothing Then Exit Function
    mTimeSet.rsRegPlan.Filter = "������Ŀ='" & str������Ŀ & "'"
    If mTimeSet.rsRegPlan.RecordCount = 0 Then mTimeSet.rsRegPlan.Filter = 0: Exit Function
    lng�޺� = Nvl(mTimeSet.rsRegPlan!�޺���, 0): lng��Լ = Nvl(mTimeSet.rsRegPlan!��Լ��, 0)
    If lng��Լ = 0 Then lng��Լ = lng�޺�
    If lng�޺� = 0 Then
        MsgBox "��ǰ�ű���" & str������Ŀ & ",û�жԹҺ�����������,�޷�����ʱ��,����!", vbOKOnly, Me.Caption
        Exit Function
    End If


    strʱ�� = mTimeSet.rsRegPlan!�Ű�
    mrs�ϰ�ʱ���.Filter = "ʱ���='" & strʱ�� & "'"
    If mrs�ϰ�ʱ���.RecordCount = 0 Then
        MsgBox "������ʱ��Ϊ[" & strʱ�� & "]�����°�ʱ��,����!", vbOKOnly, Me.Caption
        Exit Function
    End If
    
    mTimeSet.rsAssign.Filter = "������Ŀ='" & str������Ŀ & "' And ��ʹ��=0"
    Do While Not mTimeSet.rsAssign.EOF
        mTimeSet.rsAssign.Delete adAffectCurrent
        mTimeSet.rsAssign.MoveNext
    Loop
    mTimeSet.rsAssign.Filter = "������Ŀ='" & str������Ŀ & "'"
    If mTimeSet.rsAssign.RecordCount <> 0 Then
        lng�̶����� = mTimeSet.rsAssign.RecordCount
        lngĬ�ϼ�� = Val(Nvl(mTimeSet.rsAssign!ʱ����, lng���ʱ��))
        Do While Not mTimeSet.rsAssign.EOF
            lng������� = lng������� + Val(Nvl(mTimeSet.rsAssign!��������))
            mTimeSet.rsAssign.MoveNext
        Loop
    End If
    mTimeSet.rsAssign.Filter = 0
    j = 1: i = 1
    Do While Not mrs�ϰ�ʱ���.EOF
        dat��ʼʱ�� = CDate("1900-01-01 " & Format(mrs�ϰ�ʱ���!�ϰ�, "hh:mm:ss"))
        If Format(mrs�ϰ�ʱ���!�ϰ�, "hh:mm:ss") > Format(mrs�ϰ�ʱ���!�°�, "hh:mm:ss") Then
            dat����ʱ�� = CDate("1900-01-02 " & Format(mrs�ϰ�ʱ���!�°�, "hh:mm:ss"))
        Else
            dat����ʱ�� = CDate("1900-01-01 " & Format(mrs�ϰ�ʱ���!�°�, "hh:mm:ss"))
        End If

        If blnExit Then Exit Do
        datʱ�� = dat��ʼʱ��
        mrs�ϰ�ʱ���.MoveNext

        If mTimeSet.bln��ſ��� Then
            For i = j To lng�޺�
                If Format(datʱ��, "yyyy-MM-dd hh:mm:ss") >= Format(dat����ʱ��, "yyyy-MM-dd hh:mm:ss") Then
                    j = i
                    Exit For
                End If
                If i > lng�̶����� Then
                    With mTimeSet.rsAssign
                        .AddNew
                        !������Ŀ = str������Ŀ
                        !��ʼʱ�� = Format(datʱ��, "hh:mm:00")
                        !ʱ�� = Format(datʱ��, "hh:00:00")
                        If Format(DateAdd("n", lng���ʱ��, datʱ��), "yyyy-MM-dd hh:mm:ss") >= Format(dat����ʱ��, "yyyy-MM-dd hh:mm:ss") Then
                            !����ʱ�� = Format(dat����ʱ��, "hh:mm:ss")
                            !ʱ��� = Format(datʱ��, "hh:mm") & "-" & Format(dat����ʱ��, "hh:mm")
                        Else
                            !����ʱ�� = Format(DateAdd("n", lng���ʱ��, datʱ��), "hh:mm:00")
                            !ʱ��� = Format(datʱ��, "hh:mm") & "-" & Format(DateAdd("n", lng���ʱ��, datʱ��), "hh:mm")
                        End If
                        !ʱ���� = lng���ʱ��
                        !�������� = IIf(lng������� >= lng�޺�, 0, 1)
                        !�Ƿ�ԤԼ = 0
                        !��� = i
                        !��ʹ�� = 0
                        .Update
                        lng������� = lng������� + IIf(lng������� >= lng�޺�, 0, 1)
                    End With
                Else
                    mTimeSet.rsAssign.Filter = "���=" & i
                    If mTimeSet.rsAssign.RecordCount > 0 Then
                        lngĬ�ϼ�� = Nvl(mTimeSet.rsAssign!ʱ����, lngĬ�ϼ��)
                    Else
                        lngĬ�ϼ�� = lng���ʱ��
                    End If
                End If
                datʱ�� = DateAdd("n", IIf(i > lng�̶�����, lng���ʱ��, lngĬ�ϼ��), datʱ��)
            Next

        Else    '����ſ���

            Do While Not Format(datʱ��, "yyyy-MM-dd hh:mm:ss") >= Format(dat����ʱ��, "yyyy-MM-dd hh:mm:ss")

                ' If lngStart > lng��Լ Then blnExit = True: Exit For
                If Format(datʱ��, "yyyy-MM-dd hh:mm:ss") >= Format(dat����ʱ��, "yyyy-MM-dd hh:mm:ss") Then Exit Do

                If i > lng�̶����� Then
                    With mTimeSet.rsAssign
                        .AddNew
                        !������Ŀ = str������Ŀ
                        !��ʼʱ�� = Format(datʱ��, "hh:mm:00")
                        !ʱ�� = Format(datʱ��, "hh:00:00")
                        If Format(DateAdd("n", lng���ʱ��, datʱ��), "yyyy-MM-dd hh:mm:ss") >= Format(dat����ʱ��, "yyyy-MM-dd hh:mm:ss") Then
                            !����ʱ�� = Format(dat����ʱ��, "hh:mm:ss")
                            !ʱ��� = Format(datʱ��, "hh:mm") & "-" & Format(dat����ʱ��, "hh:mm")
                        Else
                            !����ʱ�� = Format(DateAdd("n", lng���ʱ��, datʱ��), "hh:mm:00")
                            !ʱ��� = Format(datʱ��, "hh:mm") & "-" & Format(DateAdd("n", lng���ʱ��, datʱ��), "hh:mm")
                        End If
                        !ʱ���� = lng���ʱ��
                        !�������� = IIf(lng������� >= lng��Լ, 0, 1)
                        !�Ƿ�ԤԼ = 1
                        !��� = i
                        !��ʹ�� = 0
                        .Update
                        lng������� = lng������� + IIf(lng������� >= lng��Լ, 0, 1)
                    End With
                Else
                    mTimeSet.rsAssign.Filter = "���=" & i
                    If mTimeSet.rsAssign.RecordCount > 0 Then
                        lngĬ�ϼ�� = Nvl(mTimeSet.rsAssign!ʱ����, lngĬ�ϼ��)
                    Else
                        lngĬ�ϼ�� = lng���ʱ��
                    End If
                End If
                datʱ�� = DateAdd("n", IIf(i > lng�̶�����, lng���ʱ��, lngĬ�ϼ��), datʱ��)
                i = i + 1
            Loop
        End If
        If i > lng�޺� And mTimeSet.bln��ſ��� Then
            blnExit = True
        End If
    Loop
    AssignReapportion = True
End Function

Private Sub cmd����ʱ��_Click()
    If AssignReapportion(Val(txtTimeOut.Text), tbSubPage.Item(mTimeSet.lngSelIndex).Caption) = False Then Exit Sub
    Call tbSubPage_SelectedChanged(tbSubPage.Item(mTimeSet.lngSelIndex))
End Sub

Private Sub tbSubPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
   Dim str������Ŀ As String
   If Not mTimeSet.blnIsInit Then Exit Sub

   If Item.index <> mTimeSet.lngSelIndex And mTimeSet.lngSelIndex <> -1 Then '
'     If mTimeSet.lngSelIndex <> -1 And mTimeSet.blnChange Then
'        If VsTimeValidate(mTimeSet.lngSelIndex) = False Then
'            mTimeSet.blnOnChange = True
'            tbSubPage.Item(mTimeSet.lngSelIndex).Selected = True
'            mTimeSet.blnOnChange = False
'            Exit Sub
'        End If
'     End If
     
     str������Ŀ = GetVsGridCaption(mTimeSet.lngSelIndex)
     If MoveAssign(str������Ŀ) = False Then
        If mTimeSet.lngSelIndex <> -1 Then tbSubPage.Item(mTimeSet.lngSelIndex).Selected = True
        Exit Sub
     End If
   End If

   If mTimeSet.blnOnChange Then Exit Sub
   mTimeSet.lngSelIndex = Item.index
   SetStyle mTimeSet.bln��ſ���, Item.index

   LoadTimeSetPlan Item.Caption
   setVsGridSNStyle Item.index
End Sub

Private Sub setVsGridSNStyle(ByVal lngIndex As Long)
 '�����ʱ����vsFex����������ݺ���Ҫ�������ñ����ʽ
 '****************************************
'�Ա����ʽ��������
'****************************************
    Dim i           As Long
    Dim lngWidth    As Long
    Dim X           As Long
    Dim Y           As Long
    Dim j           As Long
    Dim lngHeight   As Long
   
    If vsTime(lngIndex).Cols <= 1 Then Exit Sub
    If mTimeSet.bln��ſ��� Then
        With vsTime(lngIndex)
            For i = 1 To .Cols - 1
                .ColAlignment(i) = flexAlignCenterCenter
                .ColWidth(i) = 1200
             Next
             .ColWidth(0) = 1200
             .FixedAlignment(0) = flexAlignRightTop
             .ColAlignment(0) = flexAlignRightTop
             If .Rows > 0 Then
                .Cell(flexcpFontBold, 0, 0, .Rows - 1, 0) = True
                .Cell(flexcpFontSize, 0, 0, .Rows - 1, 0) = 16
             End If
    '��ʱ������ü������
         End With
    Else
    
    End If
    With vsTime(lngIndex)
         If (mTimeSet.bln��ſ��� And .Rows = 0) Or (mTimeSet.bln��ſ��� = False And .Rows = 1) Then Exit Sub
         For i = IIf(mTimeSet.bln��ſ���, 0, 1) To .Rows - 1 Step 2
             .Cell(flexcpBackColor, i, IIf(mTimeSet.bln��ſ���, 1, 0), i, .Cols - 1) = &HE0E0D3
         Next
    End With

End Sub

Private Function LoadTimeSetPlan(ByVal str������Ŀ As String) As Boolean
    Dim nIndex As Integer
    Dim i As Long, r As Long
    Dim strTime As String
    Dim lngRow As Long
    Dim lngCol As Long
    Dim strʱ�� As String
    Dim strData As String
    If mTimeSet.rsAssign Is Nothing Then Exit Function
    nIndex = GetVsGridIndex(str������Ŀ)
    cmdԤԼ(nIndex).Visible = False
    cmdɾ��(nIndex).Visible = False
    If Not mTimeSet.bln��ſ��� Then
        With vsTime(nIndex)
            mTimeSet.rsAssign.Filter = "������Ŀ='" & str������Ŀ & "'"
            mTimeSet.rsAssign.Sort = "��� asc "
               r = 1: i = -1
            Do While Not mTimeSet.rsAssign.EOF
                i = i + 1
                If i * 2 > .Cols - 2 Then r = r + 1: i = 0
                strData = Val(Nvl(mTimeSet.rsAssign!��������))
                strTime = mTimeSet.rsAssign!ʱ���
                If r > .Rows - 1 Then .Rows = .Rows + 1
                .TextMatrix(r, i * 2) = strTime
                .TextMatrix(r, i * 2 + 1) = strData
                If Val(Nvl(mTimeSet.rsAssign!��ʹ��)) = 1 Then
                    .Cell(flexcpFontUnderline, r, i * 2, r, i * 2 + 1) = True
                Else
                   '������ɫ����
                End If
                mTimeSet.rsAssign.MoveNext
            Loop
             mTimeSet.rsAssign.Filter = 0
        End With
        LoadTimeSetPlan = True
        Exit Function
    End If
    '-��ſ���
    With vsTime(nIndex)
        .Cols = 1: .FixedCols = 1
        .Rows = 0: .FixedRows = 0
        .Cols = 2: .Clear
        lngRow = -1: lngCol = 0
        mTimeSet.rsAssign.Filter = "������Ŀ='" & str������Ŀ & "'"
        If mTimeSet.rsAssign.RecordCount = 0 Then mTimeSet.rsAssign.Filter = 0: Exit Function
        i = 1
        mTimeSet.rsAssign.Sort = "��� asc "
        Do While Not mTimeSet.rsAssign.EOF
             lngCol = lngCol + 1
             If strʱ�� <> Nvl(mTimeSet.rsAssign!ʱ��) Then lngRow = lngRow + 2: lngCol = 1
             If lngCol = 1 Then
                strʱ�� = Nvl(mTimeSet.rsAssign!ʱ��)
                If lngRow > .Rows - 1 Then .Rows = .Rows + 2
                 .TextMatrix(lngRow - 1, 0) = Format(strʱ��, "hh:mm")
                 .TextMatrix(lngRow, 0) = Format(strʱ��, "hh:mm")
             End If
             strData = mTimeSet.rsAssign!���
             strTime = mTimeSet.rsAssign!ʱ���
            If lngCol > .Cols - 1 Then .Cols = .Cols + 1
            If lngRow > .Rows - 1 Then .Rows = .Rows + 2
             .TextMatrix(lngRow - 1, lngCol) = strData
             .TextMatrix(lngRow, lngCol) = strTime
            If Val(Nvl(mTimeSet.rsAssign!�Ƿ�ԤԼ)) = 1 Then
                .Cell(flexcpForeColor, lngRow - 1, lngCol, lngRow, lngCol) = vbBlue
                .Cell(flexcpFontBold, lngRow - 1, lngCol, lngRow, lngCol) = True
            End If
            If Val(Nvl(mTimeSet.rsAssign!��ʹ��)) = 1 Then
                    .Cell(flexcpFontUnderline, lngRow - 1, lngCol, lngRow, lngCol) = True
            Else
               '������ɫ����
            End If
            mTimeSet.rsAssign.MoveNext
        Loop
        If .Rows = 0 Then .Rows = 1
    End With
End Function

Private Sub SetStyle(ByVal bln��ſ��� As Boolean, ByVal lngIndex As Long)
    '����
    Dim i As Long
    Dim lngWidth As Long
    Dim lngHeight As Long
    If lngIndex > vsTime.UBound Then Exit Sub
    If Not mTimeSet.blnIsInit Then Exit Sub
    With vsTime(lngIndex)
        If bln��ſ��� Then
            If .Cols <= 1 Then Exit Sub
            .Rows = 0
            .FixedCols = 1
            .MergeCellsFixed = flexMergeFree
            .MergeCol(0) = True
            .FixedAlignment(0) = flexAlignRightTop
            .ColAlignment(0) = flexAlignRightTop
            lngWidth = 1275
        Else
             .Clear
             .Cols = 8: .Rows = 1
             .MergeCol(0) = False
            .FixedCols = 0
            .FixedAlignment(0) = flexAlignCenterCenter
            .FixedRows = 1
            
            .RowHeightMax = 400: .RowHeightMin = 400
            For i = 0 To .Cols - 1 Step 2
              .TextMatrix(0, i) = "ʱ���"
            Next
            For i = 1 To .Cols - 1 Step 2
              .TextMatrix(0, i) = "ԤԼ����"
            Next
            For i = 0 To .Cols - 1
               .ColAlignment(i) = flexAlignCenterCenter
               .ColWidth(i) = 1200
            Next
        End If
    End With
End Sub
 
 Private Sub InitRs(Optional ByVal blnInitRs As Boolean = True)
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    If Not mTimeSet.rsAssign Is Nothing Then Exit Sub
    With mTimeSet.rsAssign
        Set mTimeSet.rsAssign = New ADODB.Recordset
        mTimeSet.rsAssign.Fields.Append "������Ŀ", adVarChar, 20
        mTimeSet.rsAssign.Fields.Append "��ʼʱ��", adVarChar, 20
        mTimeSet.rsAssign.Fields.Append "ʱ��", adVarChar, 20
        mTimeSet.rsAssign.Fields.Append "����ʱ��", adVarChar, 20
        mTimeSet.rsAssign.Fields.Append "ʱ���", adVarChar, 50
        mTimeSet.rsAssign.Fields.Append "ʱ����", adBigInt, 4
        mTimeSet.rsAssign.Fields.Append "��������", adBigInt, 10
        mTimeSet.rsAssign.Fields.Append "�Ƿ�ԤԼ", adBigInt, 18
        mTimeSet.rsAssign.Fields.Append "���", adBigInt, 18
        mTimeSet.rsAssign.Fields.Append "��ʹ��", adBigInt, 2
        mTimeSet.rsAssign.CursorLocation = adUseClient
        mTimeSet.rsAssign.LockType = adLockOptimistic
        mTimeSet.rsAssign.CursorType = adOpenStatic
        mTimeSet.rsAssign.Open
    End With
    If blnInitRs Then Call InitAssignRs
End Sub
 
 Private Function InitAssignRs() As Boolean
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim lng�̶� As Long  '�̶�����Ų��������
    Dim i As Long
    '��ʼ���ѷ������ݼ���
    If mPlanEditType = EM_����_���� Then Exit Function
     On Error GoTo Hd
    If mPlanEditType = EM_����_���� Or mPlanEditType = EM_����_�޸� Or mPlanEditType = EM_�ƻ�_���� Then
        strSQL = "Select ���, ���� As ������Ŀ, To_Char(��ʼʱ��, 'hh24:mi:ss') As ��ʼʱ��, To_Char(����ʱ��, 'hh24:mi:ss') As ����ʱ��,"
        strSQL = strSQL & vbCrLf & "         �Ƿ�ԤԼ , ��������,To_Char(��ʼʱ��, 'hh24') || ':00:00' As ʱ��,To_Char(��ʼʱ��, 'hh24:mi') || '-' || To_Char(����ʱ��, 'hh24:mi') As ʱ���"
        strSQL = strSQL & vbCrLf & " From �ҺŰ���ʱ�� Where ����ID=[1] "
        strSQL = strSQL & vbCrLf & " Order By ����"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mTimeSet.lng����ID)
    ElseIf mPlanEditType = EM_�ƻ�_���� Or mPlanEditType = EM_�ƻ�_�޸� Then
        strSQL = "Select ���, ���� As ������Ŀ, To_Char(��ʼʱ��, 'hh24:mi:ss') As ��ʼʱ��, To_Char(����ʱ��, 'hh24:mi:ss') As ����ʱ��,"
        strSQL = strSQL & vbCrLf & "         �Ƿ�ԤԼ , ��������, To_Char(��ʼʱ��, 'hh24') || ':00:00' As ʱ��,To_Char(��ʼʱ��, 'hh24:mi') || '-' || To_Char(����ʱ��, 'hh24:mi') As ʱ���"
        strSQL = strSQL & vbCrLf & " From �Һżƻ�ʱ�� Where �ƻ�ID=[1] "
        strSQL = strSQL & vbCrLf & " Order By ����"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mTimeSet.lng�ƻ�ID)
    End If
    Do While Not rsTmp.EOF
            With mTimeSet.rsAssign
                .AddNew
                !������Ŀ = Nvl(rsTmp!������Ŀ)
                !��ʼʱ�� = Nvl(rsTmp!��ʼʱ��, "00:00:00")
                !����ʱ�� = Nvl(rsTmp!����ʱ��, "00:00:00")
                !ʱ��� = Nvl(rsTmp!ʱ���, "__:__-__:__")
                !ʱ���� = DateDiff("n", CDate(!��ʼʱ��), CDate(!����ʱ��))
                !�������� = Val(Nvl(rsTmp!��������))
                !�Ƿ�ԤԼ = Val(Nvl(rsTmp!�Ƿ�ԤԼ))
                !ʱ�� = Nvl(rsTmp!ʱ��, "00:00:00")
                !��� = Val(Nvl(rsTmp!���))
                lng�̶� = 0
                If Not mTimeSet.rsHistory Is Nothing Then
                mTimeSet.rsHistory.Filter = "������Ŀ='" & Nvl(rsTmp!������Ŀ) & "'"
                    If mTimeSet.rsHistory.RecordCount > 0 Then
                        If CStr(mTimeSet.rsHistory!����ʱ��) >= CStr(Nvl(rsTmp!��ʼʱ��, "00:00:00")) Then
                            lng�̶� = 1
                        End If
                    End If
                End If
                !��ʹ�� = lng�̶�
                .Update
            End With
        rsTmp.MoveNext
    Loop
    Call AssignManage
    InitAssignRs = True
Exit Function
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Function AssignManage() As Boolean
    Dim varData As Variant, varTemp As Variant, i As Long
    Dim j As Long, lngIndex As Long, p As Long, strTemp As String
    Dim lng�޺��� As Long, lng��Լ�� As Long, lng�������� As Long
    Dim lng����ԤԼ As Long, lngTmp  As Long, lngTemp As Long
    Dim str���ʱ�� As String, blnChange As Boolean
     
    varData = Split(mTimeSet.str����, "|")
    lngIndex = -1
    For i = 0 To 6
        strTemp = Switch(i = 0, "����", i = 1, "��һ", i = 2, "�ܶ�", i = 3, "����", i = 4, "����", i = 5, "����", i = 6, "����")
        '�������Ӧ��ʱ��ı�
        If InStr("|" & mTimeSet.str����, "|" & strTemp & ",") = 0 Or InStr("|" & mTimeSet.strӦ��ʱ�� & "|", "|" & strTemp & "|") = 0 Then
            mTimeSet.rsAssign.Filter = "������Ŀ='" & strTemp & "'"
            Do While Not mTimeSet.rsAssign.EOF
                mTimeSet.rsAssign.Delete adAffectCurrent
                mTimeSet.rsAssign.Update
                mTimeSet.rsAssign.MoveNext
            Loop
        End If
    Next
    For i = 0 To UBound(varData)
        ''��һ,�޺���,��Լ��|�ܶ�,�޺���,��Լ��|....
        varTemp = Split(varData(i) & ",,,,", ",")
        If varTemp(0) <> "" Then
            lng�޺��� = Val(varTemp(1)): lng��Լ�� = Val(varTemp(2))
            If lng��Լ�� = 0 Then lng��Լ�� = lng�޺���
            str���ʱ�� = ""
            If Not mTimeSet.rsHistory Is Nothing Then
                mTimeSet.rsHistory.Filter = "������Ŀ='" & varTemp(0) & "'"
                If mTimeSet.rsHistory.RecordCount = 0 Then
                   str���ʱ�� = ""
                Else
                   str���ʱ�� = Nvl(mTimeSet.rsHistory!����ʱ��)
                End If
            End If
            mTimeSet.rsAssign.Filter = "������Ŀ='" & varTemp(0) & "'"
            mTimeSet.rsAssign.Sort = "���"

              lng�������� = 0
              blnChange = False
             Do While Not mTimeSet.rsAssign.EOF
                If lng�������� + Val(Nvl(mTimeSet.rsAssign!��������)) > IIf(mTimeSet.bln��ſ���, lng�޺���, lng��Լ��) Then
                    blnChange = True
                    If Val(Nvl(mTimeSet.rsAssign!��ʹ��)) = 0 Then
                        lngTmp = Val(mTimeSet.rsAssign!��������)
                        lngTemp = lng�������� + lngTmp - IIf(mTimeSet.bln��ſ���, lng�޺���, lng��Լ��)
                        If lngTmp <= lngTemp Then
                            lngTmp = 0
                        Else
                            lngTmp = lngTmp - lngTemp
                            lng�������� = lng�޺���
                        End If
                        mTimeSet.rsAssign!�������� = lngTmp
                        mTimeSet.rsAssign.Update
                        If mTimeSet.bln��ſ��� Then
                            mTimeSet.rsAssign.Delete adAffectCurrent
                        End If
                    End If
                Else
                    lng�������� = lng�������� + Val(Nvl(mTimeSet.rsAssign!��������))
                End If
                mTimeSet.rsAssign.MoveNext
             Loop
             If blnChange Then
                mTimeSet.rsAssign.Filter = "������Ŀ='" & varTemp(0) & "' And ��������>0"
                lng�������� = 0
                If mTimeSet.rsAssign.RecordCount = 0 Then mTimeSet.rsAssign.Filter = 0: AssignManage = True: Exit Function
                mTimeSet.rsAssign.Sort = "��� desc"
                mTimeSet.rsAssign.MoveFirst
                'lng��������
                Do While Not mTimeSet.rsAssign.EOF
                   lng�������� = lng�������� + Val(Nvl(mTimeSet.rsAssign!��������))
                   mTimeSet.rsAssign.MoveNext
                Loop
                mTimeSet.rsAssign.MoveFirst
                If lng�������� > IIf(mTimeSet.bln��ſ���, lng�޺���, lng��Լ��) Then
                   Do While Not mTimeSet.rsAssign.EOF
                      If Val(Nvl(mTimeSet.rsAssign!��ʹ��)) = 0 Then
                           lngTmp = Val(Nvl(mTimeSet.rsAssign!��������))
                           lngTemp = lng�������� - lng�޺���
                           If lngTemp >= lngTmp Then
                               mTimeSet.rsAssign!�������� = 0
                               mTimeSet.rsAssign.Update
                               lng�������� = lng�������� - lngTmp
                           Else
                               lngTmp = lngTmp - lngTemp
                               mTimeSet.rsAssign!�������� = lngTmp
                               mTimeSet.rsAssign.Update
                               lng�������� = lng�������� - lngTemp
                           End If
                      End If
                      If lng�������� <= lng�޺��� Then Exit Do
                      mTimeSet.rsAssign.MoveNext
                   Loop
                End If
             End If
        End If
    Next
    mTimeSet.rsAssign.Filter = 0
    If Not mTimeSet.rsHistory Is Nothing Then mTimeSet.rsHistory.Filter = 0
    AssignManage = True
End Function

Private Sub mfrmOtherCalc_zlRefreshCon(ByVal varTimes As Variant)
    Dim strSQL As String
    Dim i As Long
    Dim j As Long
    Dim lng�޺� As Long
    Dim lng��Լ As Long
    Dim dat��ʼʱ�� As Date
    Dim dat����ʱ�� As Date
    Dim lng��� As Long
    Dim strTmp As String
    Dim strʱ�� As String
    Dim str����ʱ�� As String
    Dim lngĬ�ϼ�� As Long
    Dim lng������� As Long
    Dim lng�̶����� As Long
    Dim lngTmp As Long
    Dim blnExit As Boolean
    Dim datʱ�� As Date
    Dim str�ֶμ�� As String
    Dim str������Ŀ As String
    Dim cllPro As Collection
    Dim varTemp As Variant
    Dim strStart As String
    Dim strEnd As String
    Dim int���� As Integer
    Dim strʱ�� As String
    Dim lngʱ���� As Long
    Dim varData As Variant
    Dim str��ʱ�� As String
    
    If Not mTimeSet.bln��ſ��� Then Exit Sub
    If varTimes Is Nothing Then Exit Sub
    If varTimes("ʱ����") <> "" Then
        txtTimeOut.Text = Val(varTimes("ʱ����"))
        Call cmd����ʱ��_Click
        Exit Sub
    End If

    str�ֶμ�� = varTimes("�ֶμ��")
    If Trim(str�ֶμ��) = "" Then Exit Sub


    If mrs�ϰ�ʱ��� Is Nothing Then
        Call Initʱ���
    End If
    str������Ŀ = GetVsGridCaption(mTimeSet.lngSelIndex)


    If mrs�ϰ�ʱ��� Is Nothing Then Exit Sub
    mTimeSet.rsRegPlan.Filter = "������Ŀ='" & str������Ŀ & "'"
    If mTimeSet.rsRegPlan.RecordCount = 0 Then mTimeSet.rsRegPlan.Filter = 0: Exit Sub
    lng�޺� = Nvl(mTimeSet.rsRegPlan!�޺���, 0): lng��Լ = Nvl(mTimeSet.rsRegPlan!��Լ��, 0)
    If lng��Լ = 0 Then lng��Լ = lng�޺�
    If lng�޺� = 0 Then
        MsgBox "��ǰ�ű���" & str������Ŀ & ",û�жԹҺ�����������,�޷�����ʱ��,����!", vbOKOnly, Me.Caption
        Exit Sub
    End If


    strʱ�� = mTimeSet.rsRegPlan!�Ű�
    mrs�ϰ�ʱ���.Filter = "ʱ���='" & strʱ�� & "'"
    If mrs�ϰ�ʱ���.RecordCount = 0 Then
        MsgBox "������ʱ��Ϊ[" & strʱ�� & "]�����°�ʱ��,����!", vbOKOnly, Me.Caption
        Exit Sub

    End If

    Set cllPro = New Collection
    varData = Split(str�ֶμ��, ";")

    For i = 0 To UBound(varData)
        varTemp = Split(varData(i), ",")
        int���� = Val(varTemp(1))
        varTemp = Split(varTemp(0), "��")
        strStart = varTemp(0)
        strEnd = varTemp(1)
        cllPro.Add int����, "K" & Replace(strStart, ":", "_")
        cllPro.Add strStart, "K" & Replace(strStart, ":", "_") & "_Start"
        cllPro.Add strEnd, "K" & Replace(strStart, ":", "_") & "_End"
    Next

    mTimeSet.rsAssign.Filter = "������Ŀ='" & str������Ŀ & "' And ��ʹ��=0"
    Do While Not mTimeSet.rsAssign.EOF
        mTimeSet.rsAssign.Delete adAffectCurrent
        mTimeSet.rsAssign.MoveNext
    Loop
    mTimeSet.rsAssign.Filter = "������Ŀ='" & str������Ŀ & "'"
    If mTimeSet.rsAssign.RecordCount <> 0 Then
        lng�̶����� = mTimeSet.rsAssign.RecordCount
        lngĬ�ϼ�� = Val(Nvl(mTimeSet.rsAssign!ʱ����, lngʱ����))
        lngʱ���� = lngĬ�ϼ��
        Do While Not mTimeSet.rsAssign.EOF
            lng������� = lng������� + Val(Nvl(mTimeSet.rsAssign!��������))
            mTimeSet.rsAssign.MoveNext
        Loop
    End If
    mTimeSet.rsAssign.Filter = 0
    j = 1: i = 1
    Do While Not mrs�ϰ�ʱ���.EOF
        dat��ʼʱ�� = CDate("1900-01-01 " & Format(mrs�ϰ�ʱ���!�ϰ�, "hh:mm:ss"))
        If Format(mrs�ϰ�ʱ���!�ϰ�, "hh:mm:ss") > Format(mrs�ϰ�ʱ���!�°�, "hh:mm:ss") Then
            dat����ʱ�� = CDate("1900-01-02 " & Format(mrs�ϰ�ʱ���!�°�, "hh:mm:ss"))
        Else
            dat����ʱ�� = CDate("1900-01-01 " & Format(mrs�ϰ�ʱ���!�°�, "hh:mm:ss"))
        End If

        If blnExit Then Exit Do
        datʱ�� = dat��ʼʱ��
        mrs�ϰ�ʱ���.MoveNext

        For i = j To lng�޺�
            If Format(datʱ��, "yyyy-MM-dd hh:mm:ss") >= Format(dat����ʱ��, "yyyy-MM-dd hh:mm:ss") Then
                j = i
                Exit For
            End If
            If strʱ�� <> Format(datʱ��, "HH:00") Then
                If strʱ�� = "" Or Format(str��ʱ��, "HH:00") = Format(datʱ��, "HH:00") Then
                    If strʱ�� = "" Then str��ʱ�� = Format(datʱ��, "HH:MM")
                    strʱ�� = Format(datʱ��, "HH:MM")
    
                    If InStr("," & str�ֶμ��, str��ʱ�� & "��") > 0 Then
                        lngʱ���� = Val(cllPro("K" & Replace(str��ʱ��, ":", "_")))
                    Else
                        lngʱ���� = lngĬ�ϼ��
                    End If
                Else
                    strʱ�� = Format(datʱ��, "HH:00")
                    '�����:115865,����,2017/10/31,����ҺŰ��Ź���ʱ������ʱ�޸�ʱ����Ϊ"0"ʱ����
                    If InStr("," & str�ֶμ��, strʱ�� & "��") > 0 Then
                        lngʱ���� = Val(cllPro("K" & Replace(strʱ��, ":", "_")))
                    Else
                        lngʱ���� = lngĬ�ϼ��
                    End If
                End If
            End If
            If lngʱ���� = 0 Then
                datʱ�� = DateAdd("h", 1, datʱ��)
                i = i - 1
            Else
                If i > lng�̶����� Then
                    With mTimeSet.rsAssign
                        .AddNew
                        !������Ŀ = str������Ŀ
                        !��ʼʱ�� = Format(datʱ��, "hh:mm:00")
                        !ʱ�� = Format(datʱ��, "hh:00:00")
                        If Format(DateAdd("n", lngʱ����, datʱ��), "hh:mm:00") >= Format(dat����ʱ��, "yyyy-MM-dd hh:mm:ss") Then
                            !����ʱ�� = Format(dat����ʱ��, "hh:mm:ss")
                            !ʱ��� = Format(datʱ��, "hh:mm") & "-" & Format(dat����ʱ��, "hh:mm")
                        Else
                            !����ʱ�� = Format(DateAdd("n", lngʱ����, datʱ��), "hh:mm:00")
                            !ʱ��� = Format(datʱ��, "hh:mm") & "-" & Format(DateAdd("n", lngʱ����, datʱ��), "hh:mm")
                        End If
                        !ʱ���� = lngʱ����
                        !�������� = IIf(lng������� >= lng�޺�, 0, 1)
                        !�Ƿ�ԤԼ = 0
                        !��� = i
                        !��ʹ�� = 0
                        .Update
                        lng������� = lng������� + IIf(lng������� >= lng�޺�, 0, 1)
                    End With
                Else
                    mTimeSet.rsAssign.Filter = "���=" & i
                    If mTimeSet.rsAssign.RecordCount > 0 Then
                        lngĬ�ϼ�� = Nvl(mTimeSet.rsAssign!ʱ����, lngĬ�ϼ��)
                    Else
                        lngĬ�ϼ�� = lngʱ����
                    End If
                End If
                datʱ�� = DateAdd("n", IIf(i > lng�̶�����, lngʱ����, lngĬ�ϼ��), datʱ��)
            End If
        Next
        If i > lng�޺� And mTimeSet.bln��ſ��� Then
            blnExit = True
        End If
    Loop
    Call tbSubPage_SelectedChanged(tbSubPage(mTimeSet.lngSelIndex))
End Sub

Private Function SaveTimeSetData(ByVal lngID As Long, ByRef cllPro As Collection) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ݱ���
    '���:
    '����:cllPro-������ر������ݵ�SQL
    '����:���˺�
    '����:2012-06-15 13:18:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    If Not mTimeSet.blnIsInit Then Exit Function
    If zl_CheckMoveAssign() = False Then Exit Function
    If VsTimeValidate(-1) = False Then
        Exit Function
    End If
    
    If mPlanEditType = EM_����_�޸� Or mPlanEditType = EM_����_���� Then
        If SaveSetData(lngID, cllPro) = False Then Exit Function
    Else
        If SavePlanData(lngID, cllPro) = False Then Exit Function
    End If
    
    SaveTimeSetData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Function SavePlanData(ByVal lngID As Long, ByRef cllPro As Collection) As Boolean
    Dim i As Long, str���� As String, lng��� As String, strSQL As String
    Dim str���s As String, bytType As Byte 'Ӧ����
    Dim bytRowStep As Byte, bytStepCol As Byte
    Dim intPage As Integer, cllPage As Collection
    Dim strʱ�� As String
    Dim strProc As String
    Dim strTmp As String
    Dim strTemp As String
    Dim p As Integer, j As Long
   
    On Error GoTo errHandle
    
    Call AssignManage  '��ŷ��䴦��
    If cllPro Is Nothing Then
        Set cllPro = New Collection
    End If
    strSQL = "Zl_�Һżƻ�ʱ��_Delete(" & lngID & ")"
    zlAddArray cllPro, strSQL
    For i = 0 To 6
        strTemp = Switch(i = 0, "����", i = 1, "��һ", i = 2, "�ܶ�", i = 3, "����", i = 4, "����", i = 5, "����", i = 6, "����")
        mTimeSet.rsAssign.Filter = "������Ŀ='" & strTemp & "'"
        If mTimeSet.rsAssign.RecordCount > 0 Then
            Do While Not mTimeSet.rsAssign.EOF
    '            ���,��ʼʱ��,����ʱ��,��������,ԤԼ��־|...
                strTmp = mTimeSet.rsAssign!���
                strTmp = strTmp & "," & mTimeSet.rsAssign!��ʼʱ�� & "," & mTimeSet.rsAssign!����ʱ�� & "," & mTimeSet.rsAssign!�������� & "," & mTimeSet.rsAssign!�Ƿ�ԤԼ
                If Len(strʱ�� & "|" & strTmp) > 4000 Then
                    strʱ�� = Mid(strʱ��, 2)
                    strSQL = "  Zl_�Һżƻ�ʱ��_Insert("
                    '  ����id_In �ҺŰ���ʱ��.����id%Type,
                    strSQL = strSQL & lngID & ","
                    '  ����_In   �ҺŰ���ʱ��.����%Type,
                    strSQL = strSQL & "'" & strTemp & "',"
                    '  ʱ��_In   Varchar2,
                    strSQL = strSQL & "'" & strʱ�� & "'"
                    strSQL = strSQL & "" & ")"
                    zlAddArray cllPro, strSQL
                    strʱ�� = ""
                End If
                strʱ�� = strʱ�� & "|" & strTmp
                mTimeSet.rsAssign.MoveNext
            Loop
            If strʱ�� <> "" Then
                 
                strʱ�� = Mid(strʱ��, 2)
                strSQL = "  Zl_�Һżƻ�ʱ��_Insert("
                '  ����id_In �ҺŰ���ʱ��.����id%Type,
                strSQL = strSQL & lngID & ","
                '  ����_In   �ҺŰ���ʱ��.����%Type,
                strSQL = strSQL & "'" & strTemp & "',"
                '  ʱ��_In   Varchar2,
                strSQL = strSQL & "'" & strʱ�� & "'"
                strSQL = strSQL & "" & ")"
                zlAddArray cllPro, strSQL
                strʱ�� = ""
            End If
        
        End If
    Next
    SavePlanData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function

Private Function SaveSetData(ByVal lngID As Long, ByRef cllPro As Collection) As Boolean
  '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ʱ�����ݱ���
    '���:lngID-����ID
    '����:cllPro-������ر������ݵ�SQL
    '����:���˺�
    '����:2012-06-15 13:18:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, str���� As String, lng��� As String, strSQL As String
    Dim str���s As String, bytType As Byte 'Ӧ����
    Dim bytRowStep As Byte, bytStepCol As Byte
    Dim intPage As Integer
    Dim strʱ�� As String
    Dim strProc As String
    Dim strTmp As String
    Dim strTemp As String
    Dim p As Integer, j As Long
     
   
    On Error GoTo errHandle
      
    Call AssignManage  '��ŷ��䴦��
    If cllPro Is Nothing Then
        Set cllPro = New Collection
    End If
    strSQL = "Zl_�ҺŰ���ʱ��_Delete(" & lngID & ")"
    zlAddArray cllPro, strSQL
    For i = 0 To 6
        strTemp = Switch(i = 0, "����", i = 1, "��һ", i = 2, "�ܶ�", i = 3, "����", i = 4, "����", i = 5, "����", i = 6, "����")
        mTimeSet.rsAssign.Filter = "������Ŀ='" & strTemp & "'"
        If mTimeSet.rsAssign.RecordCount > 0 Then
            Do While Not mTimeSet.rsAssign.EOF
    '            ���,��ʼʱ��,����ʱ��,��������,ԤԼ��־|...
                strTmp = mTimeSet.rsAssign!���
                strTmp = strTmp & "," & mTimeSet.rsAssign!��ʼʱ�� & "," & mTimeSet.rsAssign!����ʱ�� & "," & mTimeSet.rsAssign!�������� & "," & mTimeSet.rsAssign!�Ƿ�ԤԼ
                If Len(strʱ�� & "|" & strTmp) > 4000 Then
                    strʱ�� = Mid(strʱ��, 2)
                    strSQL = "  Zl_�ҺŰ���ʱ��_Insert("
                    '  ����id_In �ҺŰ���ʱ��.����id%Type,
                    strSQL = strSQL & lngID & ","
                    '  ����_In   �ҺŰ���ʱ��.����%Type,
                    strSQL = strSQL & "'" & strTemp & "',"
                    '  ʱ��_In   Varchar2,
                    strSQL = strSQL & "'" & strʱ�� & "'"
                    strSQL = strSQL & "" & ")"
                    zlAddArray cllPro, strSQL
                    strʱ�� = ""
                End If
                strʱ�� = strʱ�� & "|" & strTmp
                mTimeSet.rsAssign.MoveNext
            Loop
            If strʱ�� <> "" Then
                 
                strʱ�� = Mid(strʱ��, 2)
                strSQL = "  Zl_�ҺŰ���ʱ��_Insert("
                '  ����id_In �ҺŰ���ʱ��.����id%Type,
                strSQL = strSQL & lngID & ","
                '  ����_In   �ҺŰ���ʱ��.����%Type,
                strSQL = strSQL & "'" & strTemp & "',"
                '  ʱ��_In   Varchar2,
                strSQL = strSQL & "'" & strʱ�� & "'"
                strSQL = strSQL & "" & ")"
                zlAddArray cllPro, strSQL
                strʱ�� = ""
            End If
        
        End If
    Next
    SaveSetData = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function

Private Function InitData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ������
    '����:�ɹ�,����true,���򷵻�false
    '����:���˺�
    '����:2009-09-15 13:14:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, i As Long, rsTemp As ADODB.Recordset
    Dim bln�������� As Boolean
    Dim lngColsWidth As Long
    Dim intRow As Integer
    
    Err = 0: On Error GoTo Errhand:
    gint�ų� = GetMaxLen
    mintSysAppLimit = Val(zlDatabase.GetPara("�Һ�����ԤԼ����", glngSys))
    
    strSQL = "" & _
    "   Select '    ' ʱ��� From dual Union All  " & _
    "   Select ʱ��� From ʱ���"
    Set mrsʱ��� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    With vsPlan
        .Clear 1
        .Tag = .BuildComboList(mrsʱ���, "ʱ���")
        
        .ColComboList(1) = .BuildComboList(mrsʱ���, "ʱ���")
        For i = 2 To .Cols - 1
            .ColComboList(i) = .ColComboList(0)
        Next
    End With
    With cbo��
        Do While Not mrsʱ���.EOF
            cbo��.AddItem Nvl(mrsʱ���!ʱ���)
            mrsʱ���.MoveNext
        Loop
        .ListIndex = 0
    End With
    
   'ȡ�������ٴ�����
    Set mrs���� = GetDepartments("'�ٴ�'", "1,3", Not zlStr.IsHavePrivs(mstrPrivs, "���п���"))
    If mrs����.RecordCount = 0 Then
        MsgBox "�㲻�߱����õ��ٴ�������Ϣ����Ȩ�޲���,���ȵ����Ź����н������û���ϵͳ����Ա����Ȩ�ޣ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    cbo����.Clear
    Do While Not mrs����.EOF
        cbo����.AddItem mrs����!����
        cbo����.ItemData(cbo����.NewIndex) = Val(Nvl(mrs����!id))
        If mlngȱʡ�Һſ���ID = Val(Nvl(mrs����!id)) Then cbo����.ListIndex = cbo����.NewIndex  '���˺�:���Ӵ��������д���Ŀ���
        mrs����.MoveNext
    Loop
        
    '�Һ���Ŀ
    strSQL = "Select ID as ���,���� From �շ���ĿĿ¼ " & _
        " Where ���='1' And (Sysdate Between ����ʱ�� And ����ʱ�� Or ����ʱ��<Sysdate And ����ʱ�� Is Null)" & _
        " And (վ��='" & gstrNodeNo & "' Or վ�� is Null) " & _
        " Order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    If rsTemp.RecordCount = 0 Then
        MsgBox "û�п��õĹҺ���Ŀ��Ϣ,���ȵ��Һ���Ŀ�����г�ʼ��", vbInformation, gstrSysName
        Exit Function
    End If
    cboItem.Clear
    Do While Not rsTemp.EOF
        cboItem.AddItem rsTemp!����
        cboItem.ItemData(cboItem.NewIndex) = rsTemp!���
        rsTemp.MoveNext
    Loop
    
    '����
    strSQL = "Select ����,����,ȱʡ��־ From ���� Order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    cbo����.Clear
    Do While Not rsTemp.EOF
        cbo����.AddItem rsTemp!����
        If IIf(IsNull(rsTemp!ȱʡ��־), 0, rsTemp!ȱʡ��־) = 1 Then
            cbo����.ListIndex = cbo����.NewIndex
        End If
        rsTemp.MoveNext
    Loop
    
    '��������
    strSQL = "Select ����,���ơ�From �������� Where (վ��='" & gstrNodeNo & "' Or վ�� is Null) Order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
    '71253 ���ϴ� 2014-04-14 16:05:10 ����������ʾ��ȫ
    vsDept.Clear
    If rsTemp.RecordCount <> 0 Then
        With vsDept
            Do While Not rsTemp.EOF
                If intRow = .Rows Then .Cols = .Cols + 1: intRow = 0
                .Cell(flexcpChecked, intRow, .Cols - 1) = 2
                .TextMatrix(intRow, .Cols - 1) = Nvl(rsTemp!����)
                intRow = intRow + 1
                rsTemp.MoveNext
            Loop
            .AutoSize 0, .Cols - 1
            'vsDept�·��϶�����ʾ����
            For i = 0 To .Cols - 1 '�п�֮��
                lngColsWidth = lngColsWidth + .Cell(flexcpWidth, 0, i)
            Next
            If lngColsWidth > .ClientWidth Then .Height = .Height + 130: Frame4.Height = Frame4.Height + 30
            .Editable = flexEDKbdMouse
        End With
    End If
    
    InitData = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
End Function
Private Sub Form_Load()
    Dim intType As Integer
    Dim strSQL As String, rsTmp As ADODB.Recordset
    '1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
    mblnFirst = True
    Call InitPage
    mblnOnlyԺ��ҽ�� = Val(zlDatabase.GetPara("ֻ����ѡԺ��ҽ��", glngSys, mlngModule, "0", , InStr(1, mstrPrivs, ";��������;") > 0, intType)) = 1
    If mblnOnlyԺ��ҽ�� Then
        mnuViewDoctor(0).Checked = True
        mnuViewDoctor(1).Checked = False
    Else
        mnuViewDoctor(0).Checked = False
        mnuViewDoctor(1).Checked = True
    End If
    Call LoadTimeSetControl
    Call LoadvsDept
    lblҽ��.Tag = IIf(mblnOnlyԺ��ҽ��, "0", "1")
    lblҽ��.Caption = IIf(mblnOnlyԺ��ҽ��, "Ժ��ҽ��", "ҽ��") & IIf(lblҽ��.Tag = "1", "��", "")
    lblҽ��.ToolTipText = IIf(mblnOnlyԺ��ҽ��, "ֻ��ѡԺ�ڽ���ҽ��", "����Ԯҽ��(���˿���ѡ��Ժ��ҽ���⣬������������Ԯҽ��)")
    
End Sub

Private Sub LoadvsDept()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:vsDept��������
    '����:���ϴ�
    '����:2014-04-14 16:05:10
    '---------------------------------------------------------------------------------------------------------------------------------------------
    On Error GoTo errHandle
    With vsDept
        .Cols = 1
        .Rows = 11
        .FixedCols = 0
        .FixedRows = 0
        .RowHeight(-1) = 300
        .AllowSelection = False
        .BackColorSel = &HE0E0E0
        .BackColorBkg = &H80000005
        .SheetBorder = &H80000005 '������ɫ
        .GridColor = &H80000005 '��������ɫ
        .ColWidthMin = 1200
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadTimeSetControl()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ҳ�ؼ�
    '����:���˺�
    '����:2012-06-15 13:33:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    On Error GoTo errHandle
    For i = 1 To 6
        Load picPage(i): Load vsTime(i)
        Load cmdԤԼ(i): Load cmdɾ��(i)
       ' cmdԤԼ(i).Visible = True
        Set cmdԤԼ(i).Container = vsTime(i)
        Set cmdɾ��(i).Container = vsTime(i)
        'cmdɾ��(i).Visible = True
        picPage(i).Visible = True: vsTime(i).Visible = True
        Set vsTime(i).Container = picPage(i)
    Next
    Set vsTime(0).Container = picPage(0)
    Call LoadTimeSet
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Sub

Private Function LoadTimeSet() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ҳ
    '����:���˺�
    '����:2012-06-15 13:37:14
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, ObjItem As TabControlItem
    Dim strTemp As String
    On Error GoTo errHandle
    
    tbSubPage.RemoveAll
    For i = 0 To 6
        strTemp = Switch(i = 0, "����", i = 1, "��һ", i = 2, "�ܶ�", i = 3, "����", i = 4, "����", i = 5, "����", i = 6, "����")
        Set ObjItem = tbSubPage.InsertItem(i + 1, strTemp, picPage(i).Hwnd, 0)
        ObjItem.Tag = strTemp
    Next
     With tbSubPage
        tbSubPage.Item(0).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = False
        .PaintManager.ClientFrame = xtpTabFrameNone
        .PaintManager.Position = xtpTabPositionBottom
    End With
    LoadTimeSet = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Function

Private Sub Form_Activate()
    Dim i As Integer, intIndex As Integer
    Dim strSQL As String, rsTmp As ADODB.Recordset
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    If InitData = False Then Unload Me: Exit Sub
    If LoadCard = False Then Unload Me: Exit Sub
    Call cboDoctor_Validate(False)
    For i = 0 To opt����.UBound
        If opt����(i).Value Then Call opt����_Click(i): Exit For
    Next
    If txt�ű�.Enabled And txt�ű�.Visible Then txt�ű�.SetFocus
    strSQL = "Select 1 From �ҺŰ��żƻ� Where ����id = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
    If Not rsTmp.EOF Then
        mblnMinorChange = True
        If txtAppLimit.Enabled And txtAppLimit.Visible Then
            txtAppLimit.SetFocus
            zlControl.TxtSelAll txtAppLimit
        End If
        tbPage.Item(1).Visible = False
        txt�ű�.Enabled = False
        chkAppoint.Enabled = False
        chk��ſ���.Enabled = False
        chk����.Enabled = False
        Frame3.Enabled = False
        cbo����.Enabled = False
        cboDoctor.Enabled = False
        cbo����.Enabled = False
        cboItem.Enabled = False
        opt����.Item(0).Enabled = False
        opt����.Item(1).Enabled = False
        opt����.Item(2).Enabled = False
        opt����.Item(3).Enabled = False
        vsPlan.HighLight = flexHighlightNever
        If zlStr.IsHavePrivs(mstrPrivs, "�޸�����") = False Then
            For i = 0 To 3
                If opt����(i).Value = True Then intIndex = i
            Next i
            opt����(0).Enabled = False
            opt����(1).Enabled = False
            opt����(2).Enabled = False
            opt����(3).Enabled = False
            opt����(intIndex).Value = True
            vsDept.Enabled = False
        End If
    Else
        mblnMinorChange = False
    End If
    mblnChangeDist = False
End Sub
Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode <> vbKeyReturn Then Exit Sub
    If Me.ActiveControl Is cbo���� Then Exit Sub
    If Me.ActiveControl Is cboDoctor Then Exit Sub
    If Me.ActiveControl Is vsPlan Then Exit Sub
    Call zlCommFun.PressKey(vbKeyTab)
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Err = 0: On Error Resume Next
    mstr�����޸� = ""
    Set mcllԤԼ��Ϣ = Nothing
    If Not mrs�ϰ�ʱ��� Is Nothing Then
        Set mrs�ϰ�ʱ��� = Nothing
    End If
    If Not mrs�޺� Is Nothing Then
        Set mrs�޺� = Nothing
    End If
    If Not mrsRegHistory Is Nothing Then
        Set mrsRegHistory = Nothing
    End If
    If Not mrsRegNewData Is Nothing Then
        Set mrsRegNewData = Nothing
    End If
    If Not mrsRegOldData Is Nothing Then
        Set mrsRegOldData = Nothing
    End If
    '72729:������,2014-05-06,��һ���޸ĵ��ȡ���������Ͻǵ�X��ť���ٴ��޸�ʱ�����ʾ����ȷ������
    Call ClearCustomData
End Sub


Private Sub Form_Resize()
    Err = 0: On Error Resume Next
    With cmdOK
        .Left = ScaleWidth - .Width - 100
        cmdCancel.Left = .Left
        cmdHelp.Left = .Left
    End With

    With tbPage
        .Top = 50
        .Height = ScaleHeight - 100
        .Left = 50
        .Width = cmdOK.Left - .Left - 100
    End With

End Sub
 

Private Sub lblҽ��_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
        If Button = 0 Then Exit Sub
        If Val(lblҽ��.Tag) = 0 Then Exit Sub
        
        PopupMenu mnuPopu, 2
End Sub

Private Sub ClearVsGridCheckValue()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ؼ��ĸ�ѡ��ֵ
    '����:���ϴ�
    '����:2014-04-14 18:19:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i  As Integer
    Dim intRow As Integer
    On Error GoTo errHandle

    With vsDept
        .Redraw = flexRDNone
        intRow = -1
        For i = 0 To .Rows - 1
            If .Cell(flexcpChecked, i, .Cols - 1) = 0 Then intRow = i: Exit For
        Next
        .Cell(flexcpChecked, 0, 0, .Rows - 1, .Cols - 1) = 2
        If intRow <> -1 Then .Cell(flexcpChecked, intRow, .Cols - 1, .Rows - 1, .Cols - 1) = 0
        .Redraw = flexRDBuffered
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub mfrmTime_zlSaveTimePageSelected(ByVal str���� As String)
       If tbPage.Selected Is Nothing Then Exit Sub
       If tbPage.Selected.index <> mPageIndex.EM_ʱ�� Then
            tbPage.Item(mPageIndex.EM_ʱ��).Selected = True
       End If
End Sub
Private Sub mnuViewDoctor_Click(index As Integer)
        mnuViewDoctor(index).Checked = True
        If index = 0 Then
            mnuViewDoctor(1).Checked = False: mblnOnlyԺ��ҽ�� = True
        Else
            mnuViewDoctor(0).Checked = False: mblnOnlyԺ��ҽ�� = False
        End If
 
        lblҽ��.Caption = IIf(mblnOnlyԺ��ҽ��, "Ժ��ҽ��", "ҽ��") & "��"
        lblҽ��.ToolTipText = IIf(mblnOnlyԺ��ҽ��, "ֻ��ѡ��Ժ�ڽ���ҽ��", "����Ԯҽ��(���˿���ѡ��Ժ��ҽ���⣬������������Ԯҽ��)")
End Sub
Private Sub opt����_Click(index As Integer)
    Dim intRow As Integer
    Dim intCol As Integer
    
    '71253 ���ϴ� 2014-04-15 11:30:10 ��listView �滻ΪvsflexGrid
    If index <> 1 Then Exit Sub
    With vsDept
        For intCol = 0 To .Cols - 1
            For intRow = 0 To .Rows - 1
                If .Cell(flexcpChecked, intRow, intCol) = 1 Then
                    Call ClearVsGridCheckValue
                    .Row = intRow: .Col = intCol
                    .Cell(flexcpChecked, intRow, intCol) = 1
                    Exit Sub
                End If
            Next
        Next

    End With
    mblnChangeDist = True
End Sub

Private Sub opt��_Click()
    Dim i As Integer
    Dim strPlan As String
    
    For i = 0 To vsPlan.Cols - 1
        If Trim(vsPlan.TextMatrix(1, i)) <> "" Then
            If strPlan = "" Then
                strPlan = vsPlan.TextMatrix(1, i)
            Else
                If vsPlan.TextMatrix(1, i) <> strPlan Then
                    strPlan = "": Exit For
                End If
            End If
        End If
    Next
    
    opt��.Value = -True: txt�޺�.Enabled = True: txt��Լ.Enabled = (chkAppoint.Value = 1)
    cbo��.Enabled = True
    
    opt��.Value = False
    With vsPlan
        .Enabled = False: .TabStop = False
        For i = 1 To 7
             .TextMatrix(1, i) = ""
             .TextMatrix(2, i) = ""
             .TextMatrix(3, i) = ""
        Next
    End With
    
    cbo��.ListIndex = cbo.FindIndex(cbo��, strPlan, True)
    cbo��.SetFocus
End Sub

Private Sub opt��_Click()
    Dim i As Integer
    
    If Trim(cbo��.Text) <> "" Then
        For i = 1 To vsPlan.Cols - 1
            vsPlan.TextMatrix(1, i) = cbo��.Text
            vsPlan.TextMatrix(2, i) = txt�޺�.Text
            vsPlan.TextMatrix(3, i) = txt��Լ.Text
        Next
    End If
    
    opt��.Value = False
    cbo��.Enabled = False: txt�޺�.Enabled = False: txt��Լ.Enabled = False
    cbo��.ListIndex = -1

    opt��.Value = True
    vsPlan.Enabled = True: vsPlan.TabStop = True
    vsPlan.Col = 1: vsPlan.SetFocus
End Sub

Private Sub txt�ű�_GotFocus()
    Call zlControl.TxtSelAll(txt�ű�)
End Sub

Private Sub txt�ű�_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt�޺�_GotFocus()
    Call zlControl.TxtSelAll(txt�޺�)
End Sub

Private Sub txt�޺�_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt�޺�_Validate(Cancel As Boolean)
    If Trim(txt�޺�.Text) = "" And Trim(txt��Լ.Text) <> "" Then
        MsgBox "��Լ�����޺�!", vbInformation, gstrSysName
        Cancel = True: Exit Sub
    End If
    
    If Trim(txt�޺�.Text) <> "" And Trim(txt��Լ.Text) <> "" And Val(txt�޺�.Text) < Val(txt��Լ.Text) Then
        MsgBox "�޺���Ӧ������Լ��!", vbInformation, gstrSysName
        Cancel = True: Exit Sub
    End If
End Sub

Private Sub txt��Լ_GotFocus()
    Call zlControl.TxtSelAll(txt��Լ)
End Sub

Private Sub txt��Լ_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    If Val(txt�޺�.Text) = 0 Then KeyAscii = 0
End Sub

Private Sub txt��Լ_Validate(Cancel As Boolean)
    If Val(txt�޺�.Text) < Val(txt��Լ.Text) And _
        Trim(txt�޺�.Text) <> "" And Trim(txt��Լ.Text) <> "" Then
        MsgBox "��Լ��ӦС���޺���!", vbInformation, gstrSysName
        Cancel = True: Exit Sub
    End If
End Sub
Private Function zlCheckRegistPlanIsValied(ByRef blnMulitNumPlan As Boolean) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��鵱ǰ������ĺ����Ƿ�Ϸ�
    '����:blnMulitNumPlan-�����Ƿ��ж����ͬ(ͬһ��Ŀ,ͬһ����,ͬһ��,��ͬ��)�İ���
    '����:�Ϸ�����,�򷵻�true,���򷵻�False
    '����:���˺�
    '����:2010-12-29 10:26:45
    '������ͬһ��Ŀ,ͬһ����,ͬһ��,��ͬ�ţ�:
    '     1.ͬ���ڲ����н���İ���
    '����Ŀ:35057
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, strҽ�� As String
    Dim lng��Ŀid As Long, lng����ID As Long, lngҽ��ID As Long
    Dim str�ű� As String, strTemp As String, strTemp1 As String
    Dim i As Long
    On Error GoTo errHandle
    lng����ID = cbo����.ItemData(cbo����.ListIndex)
    lng��Ŀid = cboItem.ItemData(cboItem.ListIndex)
    lngҽ��ID = 0: strҽ�� = Trim(cboDoctor.Text)
    If cboDoctor.ListIndex <> -1 Then lngҽ��ID = cboDoctor.ItemData(cboDoctor.ListIndex)
    strSQL = "" & _
        "   Select ����,���,���� D0,��һ D1,�ܶ� D2,���� D3,���� D4,���� D5,���� D6," & _
        "           To_Char(��ʼʱ��,'YYYY-MM-DD HH24:MI:SS') ��ʼʱ��,To_Char(��ֹʱ��,'YYYY-MM-DD HH24:MI:SS') ��ֹʱ��" & _
        "   From �ҺŰ���  "

    If lngҽ��ID = 0 Then
        strSQL = strSQL & _
            "   Where ����id=[1] and  ��ĿID =[2] and ҽ������=[3] and nvl(ҽ��ID,0)=0 and ID<>" & mlngID & " Order by ���"
    Else
        strSQL = strSQL & _
        "   Where ����id=[1] and  ��ĿID =[2] and  ҽ��ID=[4] and ID<>" & mlngID & " Order by ���"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��Ŀid, strҽ��, lngҽ��ID)
    blnMulitNumPlan = Not rsTemp.EOF
    If blnMulitNumPlan = False Then zlCheckRegistPlanIsValied = True: Exit Function
    str�ű� = ""
    Do While Not rsTemp.EOF
        str�ű� = str�ű� & "," & Nvl(rsTemp!����)
        If opt��.Value Then
            If Trim(Nvl(rsTemp!D0)) <> "" Then strTemp = strTemp & vbCrLf & " ����:" & Nvl(rsTemp!D0)
            If Trim(Nvl(rsTemp!D1)) <> "" Then strTemp = strTemp & vbCrLf & " ��һ:" & Nvl(rsTemp!D1)
            If Trim(Nvl(rsTemp!D2)) <> "" Then strTemp = strTemp & vbCrLf & " �ܶ�:" & Nvl(rsTemp!D2)
            If Trim(Nvl(rsTemp!D3)) <> "" Then strTemp = strTemp & vbCrLf & " ����:" & Nvl(rsTemp!D3)
            If Trim(Nvl(rsTemp!D4)) <> "" Then strTemp = strTemp & vbCrLf & " ����:" & Nvl(rsTemp!D4)
            If Trim(Nvl(rsTemp!D5)) <> "" Then strTemp = strTemp & vbCrLf & " ����:" & Nvl(rsTemp!D5)
            If Trim(Nvl(rsTemp!D6)) <> "" Then strTemp = strTemp & vbCrLf & " ����:" & Nvl(rsTemp!D6)
            If strTemp <> "" Then
                strTemp = vbCrLf & "�ںű� [" & rsTemp!���� & "] ���������°���:" & vbCrLf & "        " & Mid(strTemp, 2)
                Call MsgBox("���֡�" & cboDoctor.Text & "��ҽ�������뵱ǰ�ű��ظ��򽻲�ĹҺŰ��� " & vbCrLf & strTemp & vbCrLf & vbCrLf & "���޸Ĵ˰���.", vbInformation + vbOKOnly + vbDefaultButton2, gstrSysName)
                zlCheckRegistPlanIsValied = False: Exit Function
            End If
        Else
            With vsPlan
                For i = 0 To 6
                    strTemp1 = "��" & Switch(i = 0, "��", i = 1, "һ", i = 2, "��", i = 3, "��", i = 4, "��", i = 5, "��", True, "��")
                    If Trim(Nvl(rsTemp.Fields("D" & i).Value)) <> "" And Trim(.TextMatrix(1, i)) <> "" Then
                        '����,�϶��ظ���
                        strTemp = strTemp & vbCrLf & strTemp1 & ":" & Trim(Nvl(rsTemp.Fields("D" & i).Value))
                    End If
                Next
            End With
            If strTemp <> "" Then
                strTemp = vbCrLf & "�ںű� [" & rsTemp!���� & "] ���������°���:" & vbCrLf & "        " & Mid(strTemp, 2)
                Call MsgBox("���֡�" & cboDoctor.Text & "��ҽ�������뵱ǰ�ű��ظ��򽻲�ĹҺŰ��� " & vbCrLf & strTemp & vbCrLf & vbCrLf & "���޸Ĵ˰���.", vbInformation + vbOKOnly + vbDefaultButton2, gstrSysName)
                zlCheckRegistPlanIsValied = False: Exit Function
            End If
        End If
        rsTemp.MoveNext
    Loop
    If str�ű� <> "" Then str�ű� = Mid(str�ű�, 2)
    If MsgBox("ע��:" & vbCrLf & "   ���֡�" & cboDoctor.Text & "��ҽ���Ѿ��������°���:" & vbCrLf & "    " & str�ű� & vbCrLf & "   �Ƿ�����Ը�ҽ�����а���?", vbQuestion + vbYesNo + vbDefaultButton2) = vbYes Then
        zlCheckRegistPlanIsValied = True: Exit Function
    End If
    zlCheckRegistPlanIsValied = False
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function
Private Function zlCheckPlanArrageIsValied() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ƻ������Ƿ���Ч
    '����:���ƻ������Ƿ������صİ���,�������صİ���,�򷵻�False,���򷵻�true
    '����:���˺�
    '����:2010-12-29 19:53:56
    '����Ŀ:35057
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As ADODB.Recordset, strҽ�� As String
    Dim lng��Ŀid As Long, lng����ID As Long, lngҽ��ID As Long
    Dim str�ű� As String, strTemp As String, strTemp1 As String
    Dim blnCheck As Boolean
    Dim i As Long
    On Error GoTo errHandle
    lng����ID = cbo����.ItemData(cbo����.ListIndex)
    lng��Ŀid = cboItem.ItemData(cboItem.ListIndex)
    lngҽ��ID = 0: strҽ�� = Trim(cboDoctor.Text)
    If cboDoctor.ListIndex <> -1 Then lngҽ��ID = cboDoctor.ItemData(cboDoctor.ListIndex)
    
    On Error GoTo errHandle
    strSQL = "" & _
    "   Select  distinct A.����,A.���� D0,A.��һ D1,A.�ܶ� D2,A.���� D3,A.���� D4,A.���� D5,A.���� D6," & _
    "           To_Char(��Чʱ��,'YYYY-MM-DD HH24:MI:SS') ��Чʱ��,To_Char(ʧЧʱ��,'YYYY-MM-DD HH24:MI:SS') ʧЧʱ��" & _
    "   From �ҺŰ��żƻ� A, �ҺŰ��� B " & _
    "   Where A.����ID=B.ID    " & _
    "      and   B.����id=[1] and  B.��ĿID =[2] and B.ҽ������=[3] and nvl(B.ҽ��ID,0)=[4] and B.ID<>" & mlngID & _
    "   Order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID, lng��Ŀid, strҽ��, lngҽ��ID)
    If rsTemp.EOF Then
        zlCheckPlanArrageIsValied = True: Exit Function
    End If
    Do While Not rsTemp.EOF
        str�ű� = str�ű� & "," & Nvl(rsTemp!����)
        blnCheck = chk��Ч��.Value = 0
        If chk��Ч��.Value = 1 Then
            blnCheck = Nvl(rsTemp!��Чʱ��) >= Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS") And Nvl(rsTemp!��Чʱ��) < Format(dtpEnd.Value, "yyyy-mm-dd HH:MM:SS")
            blnCheck = blnCheck Or Nvl(rsTemp!ʧЧʱ��) >= Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS") And Nvl(rsTemp!ʧЧʱ��) < Format(dtpEnd.Value, "yyyy-mm-dd HH:MM:SS")
            blnCheck = blnCheck Or Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS") >= Nvl(rsTemp!��Чʱ��) And Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS") < Nvl(rsTemp!ʧЧʱ��)
            blnCheck = blnCheck Or Format(dtpEnd.Value, "yyyy-mm-dd HH:MM:SS") >= Nvl(rsTemp!��Чʱ��) And Format(dtpEnd.Value, "yyyy-mm-dd HH:MM:SS") < Nvl(rsTemp!ʧЧʱ��)
             
        End If
        If blnCheck Then
            If opt��.Value Then
                If Trim(Nvl(rsTemp!D0)) <> "" Then strTemp = strTemp & vbCrLf & " ����:" & Nvl(rsTemp!D0)
                If Trim(Nvl(rsTemp!D1)) <> "" Then strTemp = strTemp & vbCrLf & " ��һ:" & Nvl(rsTemp!D1)
                If Trim(Nvl(rsTemp!D2)) <> "" Then strTemp = strTemp & vbCrLf & " �ܶ�:" & Nvl(rsTemp!D2)
                If Trim(Nvl(rsTemp!D3)) <> "" Then strTemp = strTemp & vbCrLf & " ����:" & Nvl(rsTemp!D3)
                If Trim(Nvl(rsTemp!D4)) <> "" Then strTemp = strTemp & vbCrLf & " ����:" & Nvl(rsTemp!D4)
                If Trim(Nvl(rsTemp!D5)) <> "" Then strTemp = strTemp & vbCrLf & " ����:" & Nvl(rsTemp!D5)
                If Trim(Nvl(rsTemp!D6)) <> "" Then strTemp = strTemp & vbCrLf & " ����:" & Nvl(rsTemp!D6)
                If strTemp <> "" Then
                    strTemp = vbCrLf & "�ںű� [" & rsTemp!���� & "] ���������¼ƻ�����:" & vbCrLf & "        " & Mid(strTemp, 2)
                    Call MsgBox("���֡�" & cboDoctor.Text & "��ҽ�������뵱ǰ�ű��ظ��򽻲�ĹҺŰ��� " & vbCrLf & strTemp & vbCrLf & vbCrLf & "���޸Ĵ˰���.", vbInformation + vbOKOnly + vbDefaultButton2, gstrSysName)
                    zlCheckPlanArrageIsValied = False: Exit Function
                End If
            Else
                With vsPlan
                    For i = 0 To 6
                        strTemp1 = "��" & Switch(i = 0, "��", i = 1, "һ", i = 2, "��", i = 3, "��", i = 4, "��", i = 5, "��", True, "��")
                        If Trim(Nvl(rsTemp.Fields("D" & i).Value)) <> "" And Trim(.TextMatrix(1, i)) <> "" Then
                            '����,�϶��ظ���
                            strTemp = strTemp & vbCrLf & strTemp1 & ":" & Trim(Nvl(rsTemp.Fields("D" & i).Value))
                        End If
                    Next
                End With
                If strTemp <> "" Then
                    strTemp = vbCrLf & "�ںű� [" & rsTemp!���� & "] ���������¼ƻ�����:" & vbCrLf & "        " & Mid(strTemp, 2) & vbCrLf & "  ��Чʱ��:" & IIf(Nvl(rsTemp!��Чʱ��) = "1901-01-01", "����", Nvl(rsTemp!��Чʱ��) & "-" & Nvl(rsTemp!ʧЧʱ��)) & vbCrLf
                    Call MsgBox("���֡�" & cboDoctor.Text & "��ҽ�������뵱ǰ�ű��ظ��򽻲�ĹҺŰ��� " & vbCrLf & strTemp & vbCrLf & vbCrLf & "���޸Ĵ˰���.", vbInformation + vbOKOnly + vbDefaultButton2, gstrSysName)
                    zlCheckPlanArrageIsValied = False: Exit Function
                End If
            End If
        End If
        rsTemp.MoveNext
    Loop
    zlCheckPlanArrageIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function
Private Sub vsPlan_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsPlan
        If mEditType = edt_���� Then Cancel = True: Exit Sub
        If Not opt��.Value = True Then Cancel = True: Exit Sub
        If Row = 3 And chkAppoint.Value = 0 Then Cancel = True
    End With
End Sub

 
Private Sub vsPlan_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:������صĸ�ʽ
    '����:���˺�
    '����:2011-11-11 11:33:11
    '---------------------------------------------------------------------------------------------------------------------------------------------
    With vsPlan
       If Row = 1 Then
              If Trim(.EditText) = "" Then
               .TextMatrix(2, Col) = ""
               .TextMatrix(3, Col) = ""
            End If
            Exit Sub
        End If
        If Val(.TextMatrix(Row, Col)) <> 0 Then
            .TextMatrix(Row, Col) = Format(Val(.TextMatrix(Row, Col)), "###;;;")
        End If
    End With
    Exit Sub
End Sub
Private Sub vsPlan_AfterRowColChange(ByVal OldRow As Long, ByVal OldCol As Long, ByVal NewRow As Long, ByVal NewCol As Long)
    Dim strTmp As String
    Call zl_VsGridRowChange(vsPlan, OldRow, NewRow, OldCol, NewCol)
    vsPlan.ColComboList(NewCol) = ""
     
    If mstr�����޸� <> "" Then
        strTmp = ";��" & vsPlan.TextMatrix(0, NewCol) & ";"
        vsPlan.Editable = flexEDKbdMouse
        If InStr(mstr�����޸�, strTmp) > 0 And NewRow = 1 Then vsPlan.Editable = flexEDNone
    End If
    If OldRow = 2 And Trim(vsPlan.TextMatrix(3, OldCol)) = "" And mbln�Զ�Ĭ����Լ�� Then
        vsPlan.TextMatrix(3, OldCol) = vsPlan.TextMatrix(2, OldCol)
    End If
    
    If OldRow = 1 And Trim(vsPlan.TextMatrix(1, OldCol)) = "" Then
        vsPlan.TextMatrix(2, OldCol) = ""
        vsPlan.TextMatrix(3, OldCol) = ""
    End If
    If NewRow <> 1 Then Exit Sub
    vsPlan.ColComboList(NewCol) = vsPlan.Tag
End Sub
Private Sub vsPlan_GotFocus()
    Call zl_VsGridGotFocus(vsPlan)
End Sub
Private Sub vsPlan_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim lngCol As Long, blnCancel As Boolean, lngRow As Long
    With vsPlan
        If KeyCode = vbKeyDelete Then
            .TextMatrix(.Row, .Col) = ""
        End If
    End With
    If KeyCode <> vbKeyReturn Then Exit Sub
    
    With vsPlan
        If .Row = 3 And .Col = .Cols - 1 Then zlCommFun.PressKey vbKeyTab: Exit Sub
        If .Row < 3 Then
            .Row = .Row + 1
        Else
            .Row = 1
            If .Col + 1 <= .Cols - 1 Then .Col = .Col + 1
         End If
    End With
End Sub

Private Sub vsPlan_KeyDownEdit(ByVal Row As Long, ByVal Col As Long, KeyCode As Integer, ByVal Shift As Integer)
    '�༭����
    Dim intCol As Integer, strKey As String, lngRow As Long
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    With vsPlan
            If .Row = 3 And .Col = .Cols - 1 Then zlCommFun.PressKey vbKeyTab: Exit Sub
        If .Row < 3 Then
            .Row = .Row + 1
        Else
            .Row = 1
            If .Col + 1 <= .Cols - 1 Then .Col = .Col + 1
         End If
    End With
End Sub
Private Sub vsPlan_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub
Private Sub vsPlan_KeyPressEdit(ByVal Row As Long, ByVal Col As Long, KeyAscii As Integer)
    With vsPlan
        If Row <= 1 Then Exit Sub
        VsFlxGridCheckKeyPress vsPlan, Row, Col, KeyAscii, m����ʽ
    End With
End Sub
Private Sub vsPlan_LostFocus()
    zlCommFun.OpenIme False
    Call zl_VsGridLOSTFOCUS(vsPlan)
End Sub

Private Sub vsPlan_ValidateEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strKey As String, intCol As Integer, strTemp As String, strTmp As String
    '������֤
    With vsPlan
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "")
        If .Row <= 1 Then Exit Sub
        If zlCommFun.DblIsValid(strKey, 5, True, False, 0, .ColKey(Col)) = False Then
            Cancel = True: Exit Sub
        End If
        If Val(strKey) <> 0 Then
            strKey = Format(Abs(Val(strKey)), "####;;;")
        End If
         If mstr�����޸� <> "" Then
               strTmp = "��" & vsPlan.TextMatrix(0, Col)
               'vsPlan.Editable = flexEDKbdMouse
               If InStr(mstr�����޸�, ";" & strTmp & ";") > 0 Then
                   If mcllԤԼ��Ϣ Is Nothing Then
                        Cancel = Val(strKey) < Val(.TextMatrix(Row, Col))
                   Else
                        Cancel = Val(mcllԤԼ��Ϣ("K" & strTmp & "_����")) > Val(strKey)
                        If Cancel Then Exit Sub
                        If chk��ſ���.Value = 1 Then
                            Cancel = Val(mcllԤԼ��Ϣ("K" & strTmp & "_���")) > Val(strKey)
                        End If
                   End If
               End If
        End If
        If Cancel Then Exit Sub
        If Row = 2 Then
            If Val(strKey) < Val(.TextMatrix(3, Col)) Then
                If MsgBox("�޺���С������Լ��,�Ƿ������Լ��?", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Cancel = True: Exit Sub
                .TextMatrix(3, Col) = ""
            End If
        ElseIf Row = 3 Then
            If Val(strKey) > Val(.TextMatrix(2, Col)) Then
                Call MsgBox("�޺���С������Լ��,���ܼ���", vbOKOnly, gstrSysName)
                Cancel = True: Exit Sub
            End If
        End If

        .EditText = strKey
    End With
End Sub


Private Function Checkʱ��() As Boolean
    '----------------------------------
    '�ж��Ƿ��ʱ��
    '----------------------------------
    Dim strSQL      As String
    Dim rsTmp       As ADODB.Recordset

    If mEditType = edt_���� Or mEditType = edt_���� Then Exit Function

    On Error GoTo Hd
    strSQL = _
    "   Select 1 As Hdata From �ҺŰ���ʱ�� Where ����id =[1] And Rownum=1"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
     Checkʱ�� = Not rsTmp.EOF
    Set rsTmp = Nothing
    Exit Function
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Function zl_GetԤԼ��Ϣ(ByVal lng����ID As Long) As String
    Dim strSQL As String
    Dim rsTmp As ADODB.Recordset
    Dim strTmp As String
    Dim cllԤԼ��Ϣ As Collection
    strSQL = "    " & vbCrLf & " Select ����, Max(ԤԼ����) As ԤԼ����, Max(����) As ����,���"
    strSQL = strSQL & vbCrLf & " From ("
    strSQL = strSQL & vbCrLf & "    Select Decode(To_Char(A.����ʱ��, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����',"
    strSQL = strSQL & vbCrLf & "                    '7', '����') As ����, To_Char(A.����ʱ��, 'yyyy-mm-dd') As ԤԼ����, Count(Rownum) As ����, B.ID,Max(Nvl(A.����,0)) as ��� "
    strSQL = strSQL & vbCrLf & "    From ���˹Һż�¼ A, �ҺŰ��š�b"
    strSQL = strSQL & vbCrLf & "    Where A.�ű� = B.���� And A.��¼״̬ = 1 And b.ID = [1] And"
    strSQL = strSQL & vbCrLf & "          A.����ʱ�� > A.�Ǽ�ʱ��"
'    If gintԤԼ���� = 0 Then
    strSQL = strSQL & " And A.����ʱ�� > Sysdate "
'    Else
'        strSQL = strSQL & " And A.����ʱ�� Between Sysdate And Sysdate+" & gintԤԼ����
'    End If
    strSQL = strSQL & vbCrLf & "    Group By To_Char(A.����ʱ��, 'yyyy-mm-dd'),"
    strSQL = strSQL & vbCrLf & "              Decode(To_Char(A.����ʱ��, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6',"
    strSQL = strSQL & vbCrLf & "                      '����', '7', '����'), B.ID)"
    strSQL = strSQL & vbCrLf & " Group By ����,���"
  On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lng����ID)
    If rsTmp.EOF Then Exit Function
    Set cllԤԼ��Ϣ = New Collection
    Do While Not rsTmp.EOF
        If InStr(strTmp, Nvl(rsTmp!����)) <= 0 Or strTmp = "" Then
            strTmp = strTmp & ";" & Nvl(rsTmp!����)
            cllԤԼ��Ϣ.Add Nvl(rsTmp!����), "K" & Nvl(rsTmp!����) & "_����"
            cllԤԼ��Ϣ.Add Nvl(rsTmp!ԤԼ����), "K" & Nvl(rsTmp!����) & "_����"
            cllԤԼ��Ϣ.Add Nvl(rsTmp!���), "K" & Nvl(rsTmp!����) & "_���"
        End If
        rsTmp.MoveNext
    Loop
    If strTmp <> "" Then strTmp = strTmp & ";"
    Set mcllԤԼ��Ϣ = cllԤԼ��Ϣ
    zl_GetԤԼ��Ϣ = strTmp
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Public Property Let �Զ�Ĭ����Լ��(ByVal vNewValue As Boolean)
    mbln�Զ�Ĭ����Լ�� = vNewValue
End Property


Private Sub InitPage()
     '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ҳ��ؼ�
    '����:���˺�
    '����:2009-09-09 11:01:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, ObjItem As TabControlItem, objForm As Object
    Err = 0: On Error GoTo Errhand:

    Set ObjItem = tbPage.InsertItem(mPgIndex.Pg_�ƻ�����, "�ƻ�����", picBaseBack.Hwnd, 0)
    ObjItem.Tag = mPgIndex.Pg_�ƻ�����

    Set ObjItem = tbPage.InsertItem(mPgIndex.Pg_�ƻ�ʱ��, "ʱ������", picTimeSet.Hwnd, 0)
    ObjItem.Tag = mPgIndex.Pg_�ƻ�ʱ��
     With tbPage
        tbPage.Item(0).Selected = True
        .PaintManager.Appearance = xtpTabAppearancePropertyPage2003
        .PaintManager.BoldSelected = True
        .PaintManager.Layout = xtpTabLayoutAutoSize
        .PaintManager.StaticFrame = False
        .PaintManager.ClientFrame = xtpTabFrameBorder
    End With
    Exit Sub
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
End Sub

Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If mblnChangeByCode Then Exit Sub
    PageChange Item
End Sub

Private Sub PageChange(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If mblnChangeByCode Then Exit Sub
    If Item.index = mPageIndex.EM_ʱ�� Then
       mblnChangeByCode = True
       tbPage.Item(mPageIndex.EM_����).Selected = True
        If IsValied() = False Then
            mblnChangeByCode = False
            Exit Sub
        End If
        tbPage.Item(mPageIndex.EM_ʱ��).Selected = True
        mblnChangeByCode = False
        Call LoadTimePlan
        If mTimeSet.bln��ſ��� = True Then
            cmdSelAll.Enabled = True
            cmdSelAll.Visible = True
            cmdClearAll.Enabled = True
            cmdClearAll.Visible = True
        Else
            cmdSelAll.Enabled = False
            cmdSelAll.Visible = False
            cmdClearAll.Enabled = False
            cmdClearAll.Visible = False
        End If
        mTimeSet.lngSelIndex = tbSubPage.Selected.index
    Else
        If mTimeSet.blnChange = False Then Exit Sub
        If zl_CheckMoveAssign() = False Then
             mblnChangeByCode = True
            tbPage.Item(mPageIndex.EM_ʱ��).Selected = True
             mblnChangeByCode = False
        End If
    End If
End Sub

Private Sub LoadTimePlan(Optional ByVal blnSaveBeforCheck As Boolean = False)
    Dim i As Long
    Dim lng�޺��� As Long
    Dim lng��Լ�� As Long
    Dim strTemp As String
    Dim str���� As String
    Dim str�Ű� As String
    Dim strӦ��ʱ�� As String
    Dim strӦ��     As String
     
    If Not mrsRegNewData Is Nothing Then Set mrsRegNewData = Nothing

    If mrsRegNewData Is Nothing Then
        Set mrsRegNewData = New ADODB.Recordset
        mrsRegNewData.Fields.Append "ID", adBigInt, 18
        mrsRegNewData.Fields.Append "������Ŀ", adVarChar, 20
        mrsRegNewData.Fields.Append "�Ű�", adVarChar, 20
        mrsRegNewData.Fields.Append "�޺���", adBigInt, 10
        mrsRegNewData.Fields.Append "��Լ��", adBigInt, 18
        mrsRegNewData.Fields.Append "��ſ���", adBigInt, 18
        mrsRegNewData.CursorLocation = adUseClient
        mrsRegNewData.LockType = adLockOptimistic
        mrsRegNewData.CursorType = adOpenStatic
        mrsRegNewData.Open
     End If
      
     If opt��.Value = True Then
          lng�޺��� = Val(txt�޺�.Text)
          lng��Լ�� = Val(txt��Լ.Text)
          str�Ű� = Me.cbo��.Text
          For i = 0 To 6
            strTemp = Switch(i = 0, "����", i = 1, "��һ", i = 2, "�ܶ�", i = 3, "����", i = 4, "����", i = 5, "����", i = 6, "����")
            '��һ,�޺���,��Լ��|�ܶ�,�޺���,��Լ��|....
            str���� = str���� & "|" & strTemp & "," & lng�޺��� & "," & lng��Լ��
             With mrsRegNewData
                .AddNew
                !id = 0
                !������Ŀ = strTemp
                !�Ű� = str�Ű�
                !�޺��� = lng�޺���
                !��Լ�� = lng��Լ��
                !��ſ��� = Me.chk��ſ���.Value
                .Update
            End With
            If InStr("|" & mPlanInfo.strӦ��ʱ�� & "|", "|" & strTemp & "-" & Trim(str�Ű�) & "|") > 0 Then
                '���û�иı䵱����Ű���Ϣ,�򱣳�ԭ��ʱ�β���,
                strӦ��ʱ�� = strӦ��ʱ�� & "|" & strTemp
            End If
            
            strӦ�� = strӦ�� & "|" & strTemp & "-" & str�Ű�
             
          Next

        Else

           With vsPlan
            For i = 1 To .Cols - 1
                If Trim(.TextMatrix(1, i)) <> "" Then
                    strTemp = Switch(i = 1, "����", i = 2, "��һ", i = 3, "�ܶ�", i = 4, "����", i = 5, "����", i = 6, "����", True, "����")
                    lng�޺��� = Val(Trim(vsPlan.TextMatrix(2, i)))
                    lng��Լ�� = Val(Trim(vsPlan.TextMatrix(3, i)))
                    str�Ű� = Trim(vsPlan.TextMatrix(1, i))
                    str���� = str���� & "|" & strTemp & "," & lng�޺��� & "," & lng��Լ��
                    With mrsRegNewData
                        .AddNew
                        !id = Val(mlngID)
                        !������Ŀ = strTemp
                        !�Ű� = str�Ű�
                        !�޺��� = lng�޺���
                        !��Լ�� = lng��Լ��
                        !��ſ��� = Me.chk��ſ���.Value
                        .Update
                    End With
                    If InStr("|" & mPlanInfo.strӦ��ʱ�� & "|", "|" & strTemp & "-" & Trim(str�Ű�) & "|") > 0 Then
                        '���û�иı䵱����Ű���Ϣ,�򱣳�ԭ��ʱ�β���,
                        strӦ��ʱ�� = strӦ��ʱ�� & "|" & strTemp
                    End If
                     
                    strӦ�� = strӦ�� & "|" & strTemp & "-" & str�Ű�
                     
                End If
            Next
        End With
     End If
     If str���� <> "" Then str���� = Mid(str����, 2)
     If strӦ��ʱ�� <> "" Then strӦ��ʱ�� = Mid(strӦ��ʱ��, 2)
     If strӦ�� <> "" Then strӦ�� = Mid(strӦ��, 2)
     mPlanInfo.strӦ��ʱ�� = strӦ��

     zlShowPagePlan str����, mrsRegNewData, mrsRegHistory, chk��ſ���.Value = 1, Switch(mEditType = ed_�ƻ�����, EM_����_����, mEditType = Ed_�����޸�, EM_����_�޸�, True, EM_����_����), mlngID, Val(0), blnSaveBeforCheck, strӦ��ʱ��
End Sub

Private Sub ShowTimeSetPage()
   '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾҳ��
    '����:�ɹ�����true,���򷵻�False
    '����:���˺�
    '����:2013-11-26 15:21:18
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim varData As Variant, varTemp As Variant, i As Long
    Dim j As Long, lngIndex As Long, p As Long, strTemp As String
    
    For j = 0 To tbSubPage.ItemCount - 1
         tbSubPage(j).Visible = False: tbSubPage(j).Enabled = False
         tbSubPage(j).Selected = False
    Next
    
    On Error GoTo errHandle
    varData = Split(mTimeSet.str����, "|")
    lngIndex = -1: mTimeSet.lngSelIndex = -1
    For i = 0 To UBound(varData)
        ''��һ,�޺���,��Լ��|�ܶ�,�޺���,��Լ��|....
        varTemp = Split(varData(i) & ",,,,", ",")
        If varTemp(0) <> "" Then
            For j = 0 To tbSubPage.ItemCount - 1
                If tbSubPage(j).Tag = varTemp(0) Then
                    If lngIndex < 0 Then lngIndex = j
                    tbSubPage(j).Visible = True: tbSubPage(j).Enabled = True
                    p = GetVsGridIndex(varTemp(0))
                    vsTime(p).Tag = varTemp(1) & "," & varTemp(2)
                    If mTimeSet.lngSelIndex = -1 Then mTimeSet.lngSelIndex = j: tbSubPage(j).Selected = True
                End If
            Next
        End If
    Next
    If mTimeSet.lngSelIndex = -1 Then mTimeSet.lngSelIndex = 0
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
 End Sub

Private Sub zlShowPagePlan(ByVal str������� As String, ByVal rsRegPlan As ADODB.Recordset, ByRef rsHistory As ADODB.Recordset, _
                        ByVal bln��ſ��� As Boolean, ByVal bytType As gPlanEditType, Optional ByVal lng����ID As Long, _
                        Optional ByVal lng�ƻ�ID As Long, Optional ByVal blnBeforCheck As Boolean = False, Optional ByVal strӦ��ʱ�� As String)
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾҳ��
    '����:���˺�
    '����:2012-06-15 13:49:20
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long
    mTimeSet.str���� = str�������
    mTimeSet.strӦ��ʱ�� = strӦ��ʱ��
                                                                          
    Set mTimeSet.rsRegPlan = rsRegPlan
    If bln��ſ��� <> mTimeSet.bln��ſ��� And Not mTimeSet.rsAssign Is Nothing Then
         mTimeSet.rsAssign.Filter = 0
         Do While Not mTimeSet.rsAssign.EOF
            mTimeSet.rsAssign.Delete
            mTimeSet.rsAssign.MoveNext
         Loop
         If blnBeforCheck Then Exit Sub
    End If
    mPlanEditType = bytType: mTimeSet.lng����ID = lng����ID: mTimeSet.lng�ƻ�ID = lng�ƻ�ID
    Set mTimeSet.rsHistory = rsHistory
    If Not blnBeforCheck Then Call ShowTimeSetPage
    If mTimeSet.blnIsInit Then
        Call AssignManage
    End If
    mTimeSet.blnIsInit = True
    Call InitRs(mTimeSet.bln��ſ��� = bln��ſ���)
    mTimeSet.bln��ſ��� = bln��ſ���
    If blnBeforCheck Then Exit Sub
    For i = 0 To 6
       If tbSubPage.Item(i).Selected Then
            Call tbSubPage_SelectedChanged(tbSubPage.Item(i))
            Exit For
       End If
    Next
 End Sub

Private Function LoadRegHistory() As Boolean
    Dim strSQL As String
    strSQL = "Select ������Ŀ, Max(������) As ������, Max(ͳ��) As ͳ��, Max(����ʱ��) As ����ʱ��" & vbNewLine & _
            " From (Select Decode(To_Char(a.����ʱ��, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����') As ������Ŀ," & vbNewLine & _
            "              Max(Nvl(a.����, 0)) As ������, Count(1) As ͳ��, To_Char(Max(����ʱ��), 'hh24:mi:ss') As ����ʱ��," & vbNewLine & _
            "              To_Char(����ʱ��, 'YYYY-MM-DD') As ��������" & vbNewLine & _
            "       From ���˹Һż�¼ A, �ҺŰ��� B" & vbNewLine & _
            "       Where a.��¼״̬ = 1 And a.����ʱ�� Between Sysdate And Sysdate + Nvl(b.ԤԼ����, " & IIf(gintԤԼ���� = 0, 15, gintԤԼ����) & ") And a.�ű� = b.���� And b.Id = [1] " & vbNewLine & _
            "       Group By Decode(To_Char(a.����ʱ��, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����')," & vbNewLine & _
            "                To_Char(����ʱ��, 'YYYY-MM-DD'))" & vbNewLine & _
            " Group By ������Ŀ"
                    
    On Error GoTo Hd:
    Set mrsRegHistory = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlngID)
    LoadRegHistory = True
Exit Function
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function


Private Function IsValied() As Boolean
     Dim i As Integer, intCount As Integer, j As Integer
    Dim strʱ��� As String, str���� As String, str�޺� As String
    Dim lngNextID As Long, lngҽ��ID As Long
    Dim strBegin As String, strEnd As String
    Dim strSQL As String, strInfo As String, strTmp As String, strOld As String, strNew As String
    Dim str�ű� As String
    Dim rsDoctorPlan As ADODB.Recordset
    Dim rsNewDate As ADODB.Recordset
    Dim rsOldDate As ADODB.Recordset
    Dim rsSNState As ADODB.Recordset
    Dim blnMulitNumPlan As Boolean  '�Ƿ��ΰ���
    Dim blnChange       As Boolean '�Ƿ�ı��� ʱ�䰲��
    Dim strMsg          As String

    If opt��.Value Then
        If cbo��.ListIndex = -1 Then
            MsgBox "�úű�ÿ���Ӧ��ʱ��δ���ã�", vbInformation, gstrSysName
            cbo��.SetFocus: Exit Function
        End If

        If Val(txt�޺�.Text) = 0 And Val(txt��Լ.Text) = 0 Then
            MsgBox "��������ʱ��ʱ,���������޺Ż���Լ����", vbInformation, gstrSysName
            If txt�޺�.Visible And txt�޺�.Enabled Then txt�޺�.SetFocus
            Exit Function
        End If
        If (chkAppoint.Value = 0 And chk��ſ���.Value = 0) Or (chkAppoint.Value = 1 And txt��Լ.Text <> "" And Val(txt��Լ.Text) = 0 And chk��ſ���.Value = 0) Then
            MsgBox "����ſ��Ƶİ�������ʱ��ʱ,�����ǿ�ԤԼ�İ��ţ�", vbInformation, gstrSysName
            If txt�޺�.Visible And txt�޺�.Enabled Then txt�޺�.SetFocus
            Exit Function
        End If
        '�޺���Լ����
        If Trim(txt�޺�.Text) <> "" Then
            If Trim(txt��Լ.Text) <> "" And Val(txt�޺�.Text) < Val(txt��Լ.Text) Then
                MsgBox "��Լ��ӦС���޺�����", vbInformation, gstrSysName
                txt��Լ.SetFocus: Exit Function
            End If
        ElseIf Trim(txt��Լ.Text) <> "" Then
            MsgBox "��Լ�����޺ţ�", vbInformation, gstrSysName
            txt�޺�.SetFocus: Exit Function
        End If
    Else
        If chkAppoint.Value = 0 And chk��ſ���.Value = 0 Then
            MsgBox "����ſ��Ƶİ�������ʱ��ʱ,�����ǿ�ԤԼ�İ��ţ�", vbInformation, gstrSysName
            Exit Function
        End If
        With vsPlan
            strTmp = ""
            For i = 1 To .Cols - 1
                If Trim(.TextMatrix(1, i)) <> "" Then
                    strTmp = strTmp & Trim(vsPlan.TextMatrix(1, i))

                        If Val(.TextMatrix(2, i)) = 0 And Val(.TextMatrix(3, i)) = 0 Then
                            MsgBox "��������ʱ��ʱ,���������޺Ż���Լ����", vbInformation, gstrSysName
                            .Row = 2: .Col = i
                            .SetFocus: Exit Function
                        End If

                        '�޺���Լ����
                        If Val(.TextMatrix(2, i)) <> 0 Then
                            If Trim(.TextMatrix(3, i)) <> "" And Val(.TextMatrix(2, i)) < Val(.TextMatrix(3, i)) Then
                                MsgBox "��Լ��ӦС���޺�����", vbInformation, gstrSysName
                                .Row = 2: .Col = i
                                .SetFocus: Exit Function
                            End If
                        ElseIf Trim(.TextMatrix(3, i)) <> "" Then
                            
                            MsgBox "��Լ�����޺ţ�", vbInformation, gstrSysName
                            .Row = 2: .Col = i
                            .SetFocus: Exit Function
                        End If
                End If
            Next
            If strTmp = "" Then
                MsgBox "�úű�ÿ�ܵ�Ӧ��ʱ��δ���ã�", vbInformation, gstrSysName
                vsPlan.SetFocus: Exit Function
            End If
        End With
    End If
    IsValied = True
End Function

Private Sub vsDept_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If vsDept.Cell(flexcpChecked, Row, Col) = 0 Then Cancel = True
End Sub

Private Sub vsDept_AfterEdit(ByVal Row As Long, ByVal Col As Long)
    Dim intType As Integer
    If opt����(1).Value Then
        intType = vsDept.Cell(flexcpChecked, Row, Col)
        Call ClearVsGridCheckValue
        vsDept.Cell(flexcpChecked, Row, Col) = intType
    End If
    mblnChangeDist = True
End Sub

Private Sub vsDept_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub vsDept_GotFocus()
    Dim intRow As Integer
    Dim intCol As Integer
    On Error GoTo errHandle
    
    With vsDept
        If .Row >= 0 And .Col >= 0 Then Exit Sub
        For intCol = 0 To .Cols - 1
            For intRow = 0 To .Rows - 1
                If .Cell(flexcpChecked, intRow, intCol) = 1 Then
                    .Row = intRow: .Col = intCol
                    Exit Sub
                End If
            Next
        Next
        If .Rows >= 0 And .Cols >= 0 Then .Row = 0: .Col = 0
    End With
    Exit Sub
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Sub
