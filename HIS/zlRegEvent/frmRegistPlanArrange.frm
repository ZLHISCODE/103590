VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "Codejock.SuiteCtrls.Unicode.9600.ocx"
Begin VB.Form frmRegistPlanArrange 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "�Һżƻ�����"
   ClientHeight    =   9810
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10950
   Icon            =   "frmRegistPlanArrange.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9810
   ScaleWidth      =   10950
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox picBack 
      BorderStyle     =   0  'None
      Height          =   9540
      Left            =   540
      ScaleHeight     =   9540
      ScaleWidth      =   8985
      TabIndex        =   32
      Top             =   120
      Width           =   8985
      Begin VB.OptionButton opt��Чʱ�� 
         Caption         =   "����ִ��"
         Height          =   360
         Index           =   0
         Left            =   1110
         TabIndex        =   27
         Top             =   7860
         Width           =   1170
      End
      Begin VB.OptionButton opt��Чʱ�� 
         Caption         =   "ָ��ʱ��"
         Height          =   180
         Index           =   1
         Left            =   2280
         TabIndex        =   28
         Top             =   7935
         Value           =   -1  'True
         Width           =   1035
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   3
         Left            =   6345
         TabIndex        =   39
         Top             =   8715
         Width           =   2370
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   2
         Left            =   6345
         TabIndex        =   38
         Top             =   8325
         Width           =   2370
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   1
         Left            =   1110
         TabIndex        =   37
         Top             =   8715
         Width           =   2370
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         IMEMode         =   3  'DISABLE
         Index           =   0
         Left            =   1110
         TabIndex        =   36
         Top             =   8265
         Width           =   2370
      End
      Begin VB.Frame Frame2 
         Caption         =   "Ӧ��ʱ��"
         Height          =   2010
         Left            =   60
         TabIndex        =   35
         Top             =   1635
         Width           =   8685
         Begin VB.TextBox txt��Լ 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   5145
            MaxLength       =   5
            TabIndex        =   18
            Top             =   270
            Width           =   1215
         End
         Begin VB.TextBox txt�޺� 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   3030
            MaxLength       =   5
            TabIndex        =   16
            Top             =   270
            Width           =   1215
         End
         Begin VB.ComboBox cbo�� 
            Height          =   300
            Left            =   1170
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   270
            Width           =   1365
         End
         Begin VB.OptionButton opt�� 
            Caption         =   "ÿ��(&W)"
            Height          =   315
            Left            =   225
            TabIndex        =   19
            Top             =   630
            Width           =   930
         End
         Begin VB.OptionButton opt�� 
            Caption         =   "ÿ��(&D)"
            Height          =   315
            Left            =   225
            TabIndex        =   13
            Top             =   285
            Width           =   960
         End
         Begin VSFlex8Ctl.VSFlexGrid vsPlan 
            Height          =   1275
            Left            =   1200
            TabIndex        =   20
            Top             =   600
            Width           =   7440
            _cx             =   13123
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
            ExtendLastCol   =   0   'False
            FormatString    =   $"frmRegistPlanArrange.frx":06EA
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
         Begin VB.Label Label6 
            AutoSize        =   -1  'True
            Caption         =   "��Լ"
            Height          =   180
            Left            =   4710
            TabIndex        =   17
            Top             =   330
            Width           =   360
         End
         Begin VB.Label Label5 
            AutoSize        =   -1  'True
            Caption         =   "�޺�"
            Height          =   180
            Left            =   2595
            TabIndex        =   15
            Top             =   330
            Width           =   360
         End
      End
      Begin VB.Frame Frame3 
         Caption         =   "Ӧ������:"
         Height          =   3850
         Left            =   60
         TabIndex        =   34
         Top             =   3840
         Width           =   8670
         Begin VSFlex8Ctl.VSFlexGrid vsDept 
            Height          =   3390
            Left            =   150
            TabIndex        =   25
            Top             =   300
            Width           =   8415
            _cx             =   14843
            _cy             =   5980
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
            BackColorBkg    =   -2147483643
            BackColorAlternate=   -2147483643
            GridColor       =   -2147483643
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
            Rows            =   50
            Cols            =   1
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
            TabIndex        =   24
            Top             =   0
            Width           =   1020
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "��̬����"
            Height          =   180
            Index           =   2
            Left            =   3180
            TabIndex        =   23
            Top             =   0
            Width           =   1020
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "ָ������"
            Height          =   180
            Index           =   1
            Left            =   2010
            TabIndex        =   22
            Top             =   0
            Width           =   1020
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "������"
            Height          =   180
            Index           =   0
            Left            =   1020
            TabIndex        =   21
            Top             =   0
            Value           =   -1  'True
            Width           =   1020
         End
      End
      Begin MSComCtl2.DTPicker dtpEndDate 
         Height          =   300
         Left            =   6345
         TabIndex        =   30
         Top             =   7890
         Width           =   2370
         _ExtentX        =   4180
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   134545411
         CurrentDate     =   401769
      End
      Begin MSComCtl2.DTPicker dtpBegin 
         Height          =   300
         Left            =   3360
         TabIndex        =   29
         Top             =   7890
         Width           =   2220
         _ExtentX        =   3916
         _ExtentY        =   529
         _Version        =   393216
         CustomFormat    =   "yyyy-MM-dd HH:mm"
         Format          =   134545411
         CurrentDate     =   38091
      End
      Begin VB.Frame Frame1 
         Caption         =   "������Ϣ"
         Height          =   1455
         Left            =   60
         TabIndex        =   33
         Top             =   105
         Width           =   8670
         Begin VB.CheckBox chkAppoint 
            Caption         =   "����ԤԼ"
            Height          =   300
            Left            =   6195
            TabIndex        =   12
            Top             =   1027
            Value           =   1  'Checked
            Width           =   1080
         End
         Begin VB.CheckBox chk��ſ��� 
            Caption         =   "��ſ���"
            Height          =   255
            Left            =   1750
            TabIndex        =   2
            Top             =   293
            Width           =   1095
         End
         Begin VB.ComboBox cbo���� 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   3840
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   270
            Width           =   2595
         End
         Begin VB.CheckBox chk���� 
            Caption         =   "�Һ�ʱ���뽨����"
            Height          =   195
            Left            =   3420
            TabIndex        =   11
            Top             =   1080
            Width           =   1845
         End
         Begin VB.ComboBox cbo���� 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   660
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   660
            Width           =   2400
         End
         Begin VB.ComboBox cboDoctor 
            Height          =   300
            Left            =   660
            TabIndex        =   10
            Top             =   1035
            Width           =   2400
         End
         Begin VB.ComboBox cboItem 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   3840
            Style           =   2  'Dropdown List
            TabIndex        =   8
            Top             =   675
            Width           =   2580
         End
         Begin VB.TextBox txt�ű� 
            Height          =   300
            IMEMode         =   3  'DISABLE
            Left            =   660
            MaxLength       =   5
            TabIndex        =   1
            Top             =   270
            Width           =   960
         End
         Begin VB.Label Label7 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Left            =   3405
            TabIndex        =   3
            Top             =   330
            Width           =   360
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "ҽ��"
            Height          =   180
            Left            =   240
            TabIndex        =   9
            Top             =   1080
            Width           =   360
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "��Ŀ"
            Height          =   180
            Left            =   3420
            TabIndex        =   7
            Top             =   750
            Width           =   360
         End
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Left            =   240
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
            Left            =   210
            TabIndex        =   0
            Top             =   330
            Width           =   390
         End
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��"
         Height          =   180
         Index           =   7
         Left            =   5715
         TabIndex        =   47
         Top             =   7950
         Width           =   180
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�ƻ�ʱ��"
         Height          =   180
         Index           =   6
         Left            =   240
         TabIndex        =   26
         Top             =   7935
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "���ʱ��"
         Height          =   180
         Index           =   3
         Left            =   5535
         TabIndex        =   43
         Top             =   8775
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�����"
         Height          =   180
         Index           =   2
         Left            =   5715
         TabIndex        =   42
         Top             =   8385
         Width           =   540
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ��"
         Height          =   180
         Index           =   1
         Left            =   345
         TabIndex        =   41
         Top             =   8775
         Width           =   720
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "������"
         Height          =   180
         Index           =   0
         Left            =   540
         TabIndex        =   40
         Top             =   8385
         Width           =   540
      End
   End
   Begin VB.PictureBox picTimeSet 
      BorderStyle     =   0  'None
      Height          =   7320
      Left            =   60
      ScaleHeight     =   7320
      ScaleWidth      =   8580
      TabIndex        =   50
      Top             =   1500
      Width           =   8580
      Begin VB.CommandButton cmdAuto 
         Caption         =   "�Զ�����(&A)"
         Height          =   350
         Left            =   5415
         TabIndex        =   56
         ToolTipText     =   "ͨ��������޺���,�Զ�����ʱ�������м���"
         Top             =   30
         Width           =   1150
      End
      Begin VB.CommandButton cmdSelAll 
         Caption         =   "ȫѡ(&A)"
         Height          =   350
         Left            =   6810
         TabIndex        =   68
         ToolTipText     =   "������¼���ʱ��"
         Top             =   30
         Width           =   1150
      End
      Begin VB.CommandButton cmdClearAll 
         Caption         =   "ȫ��(&D)"
         Height          =   350
         Left            =   8250
         TabIndex        =   67
         ToolTipText     =   "������¼���ʱ��"
         Top             =   30
         Width           =   1150
      End
      Begin VB.PictureBox picPage 
         BorderStyle     =   0  'None
         Height          =   3540
         Index           =   0
         Left            =   3420
         ScaleHeight     =   3540
         ScaleWidth      =   2535
         TabIndex        =   63
         Top             =   1800
         Width           =   2535
      End
      Begin VB.Frame fraӦ���� 
         Caption         =   "Ӧ���ڡ�"
         Height          =   615
         Left            =   0
         TabIndex        =   58
         Top             =   6720
         Width           =   7755
         Begin VB.OptionButton opt���� 
            Caption         =   "���кű�"
            Height          =   255
            Left            =   5685
            TabIndex        =   62
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "������(�ڿ�)"
            Height          =   255
            Left            =   3870
            TabIndex        =   61
            Top             =   240
            Width           =   1575
         End
         Begin VB.OptionButton optӦ���� 
            Caption         =   "������"
            Height          =   255
            Index           =   0
            Left            =   795
            TabIndex        =   59
            Top             =   255
            Value           =   -1  'True
            Width           =   1185
         End
         Begin VB.OptionButton optӦ���� 
            Caption         =   "��ҽ��(����)"
            Height          =   255
            Index           =   1
            Left            =   2100
            TabIndex        =   60
            Top             =   240
            Width           =   1575
         End
      End
      Begin VB.TextBox txtTimeOut 
         Height          =   300
         Left            =   1170
         MaxLength       =   4
         TabIndex        =   52
         Text            =   "10"
         Top             =   60
         Width           =   465
      End
      Begin VB.CommandButton cmd����ʱ�� 
         Caption         =   "��������(&F)"
         Height          =   350
         Left            =   2220
         TabIndex        =   53
         ToolTipText     =   "������¼���ʱ��"
         Top             =   30
         Width           =   1150
      End
      Begin VB.CommandButton cmdOther 
         Caption         =   "������������(&T)"
         Height          =   350
         Left            =   3675
         TabIndex        =   55
         ToolTipText     =   "������¼���ʱ��"
         Top             =   30
         Width           =   1515
      End
      Begin MSComCtl2.UpDown udTime 
         Height          =   300
         Left            =   1635
         TabIndex        =   51
         Top             =   60
         Width           =   255
         _ExtentX        =   450
         _ExtentY        =   529
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtTimeOut"
         BuddyDispid     =   196650
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
         Left            =   285
         TabIndex        =   64
         Top             =   1410
         Width           =   2535
         _Version        =   589884
         _ExtentX        =   4471
         _ExtentY        =   8599
         _StockProps     =   64
      End
      Begin VSFlex8Ctl.VSFlexGrid vsTime 
         Height          =   5475
         Index           =   0
         Left            =   2175
         TabIndex        =   57
         Top             =   1185
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
         FormatString    =   $"frmRegistPlanArrange.frx":07D0
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
         Begin VB.CommandButton cmdԤԼ 
            Caption         =   "Ԥ"
            Height          =   255
            Index           =   0
            Left            =   2685
            TabIndex        =   66
            Top             =   2535
            Visible         =   0   'False
            Width           =   375
         End
         Begin VB.CommandButton cmdɾ�� 
            Caption         =   "ɾ"
            Height          =   255
            Index           =   0
            Left            =   4320
            TabIndex        =   65
            Top             =   840
            Visible         =   0   'False
            Width           =   375
         End
      End
      Begin VB.Label Label8 
         AutoSize        =   -1  'True
         Caption         =   "ʱ����(��)"
         Height          =   180
         Left            =   60
         TabIndex        =   54
         Top             =   120
         Width           =   1080
      End
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   9360
      TabIndex        =   45
      Top             =   1425
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   9360
      TabIndex        =   31
      Top             =   960
      Width           =   1100
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   9360
      TabIndex        =   44
      Top             =   1950
      Width           =   1100
   End
   Begin VB.CheckBox chk������Ч 
      Caption         =   "������Ч"
      Enabled         =   0   'False
      Height          =   285
      Left            =   6600
      TabIndex        =   49
      Top             =   120
      Width           =   1650
   End
   Begin VB.CheckBox chk������� 
      Caption         =   "������������"
      Enabled         =   0   'False
      Height          =   285
      Left            =   6600
      TabIndex        =   48
      Top             =   80
      Width           =   1650
   End
   Begin XtremeSuiteControls.TabControl tbPage 
      Height          =   780
      Left            =   -120
      TabIndex        =   46
      Top             =   0
      Width           =   9015
      _Version        =   589884
      _ExtentX        =   15901
      _ExtentY        =   1376
      _StockProps     =   64
   End
End
Attribute VB_Name = "frmRegistPlanArrange"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit 'Ҫ���������
Private mstr�ƻ�ID As String, mlng����ID As Long, mblnSucces As Boolean, mblnFirst As Boolean
Private mlngModule As Long, mstrPrivs As String
Private mblnActive As Boolean
Private Enum mPageIndex
    EM_�ƻ� = 0
    EM_ʱ�� = 1
End Enum
Private mrsRegOldData As ADODB.Recordset '�������ݼ�����,ԭʼ�ҺŰ���
Private mrsRegNewData As ADODB.Recordset '�������ݼ����� �������ú�İ���
Private mrsRegHistory As ADODB.Recordset '���ιҺŵ����ݼ�
Private mrs�ϰ�ʱ��� As ADODB.Recordset
Private mrsLongPlan As ADODB.Recordset '���ڼƻ�
Private mdatBegin As Date
Private mdatEnd As Date
Private mdatOriBegin As Date
Private mdatOriEnd As Date

Public Enum gPlanEditType
    EM_����_���� = 0
    EM_����_�޸�
    EM_����_����
    EM_�ƻ�_���� = 11
    EM_�ƻ�_�޸�
    EM_�ƻ�_����
End Enum
Private mPlanEditType As gPlanEditType

Private mblnChangeByCode As Boolean
Public Enum mRegEditType
    ed_�ƻ����� = 0
    Ed_�����޸� = 1
    Ed_����ɾ�� = 2
    Ed_������� = 3
    Ed_����ȡ�� = 4
    ed_���Ų��� = 5
End Enum
Private Enum midxTxt
    idx_������ = 0
    idx_����ʱ�� = 1
    idx_����� = 2
    idx_���ʱ�� = 3
End Enum
'�����ϰ�ʱ��
Private Type t_�ϰ�ʱ��
  dat_�����ϰ� As Date
  dat_�����°� As Date
  dat_�����ϰ� As Date
  dat_�����°� As Date
End Type
Private t_ʱ�� As t_�ϰ�ʱ��
Private mEditType As mRegEditType
Private mstrԭ�Ű� As String '"��һ,����;�ܶ�,����;..."
Private mstr����ID As String
Private mblnCboClick As Boolean     '�����cbo��keypress�¼������˵����б��API����:sendmessage,�����ͣ��cbo��,����һ���ַ�,�ƿ�����򰴻س���,
'                                    cbo��ֵ�ᱣ������,�����ᴥ��click�¼�,������Ҫ��validate�¼��е���click�¼�
Private mrsDoctor As ADODB.Recordset

Private Type PlanInfo               '���Ÿı���Ҫ�Աȵ���Ϣ
    strӦ��ʱ��     As String
    str�Ű�         As String       '�Ű���Ϣ
    str�޺�         As String       '�޺���Ϣ
    bln���         As Boolean      '�Ƿ���ſ���
    blnʱ���       As Boolean      '�Ƿ�������ʱ���
End Type

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
End Type

Private mTimeSet As TimeSet
Private WithEvents mfrmOtherCalc As frmRegistPlanTimeOther
Attribute mfrmOtherCalc.VB_VarHelpID = -1
Private mblnSaveMinorChange As Boolean

Private mPlanInfo As PlanInfo '����ʱ���ڱ���ԭʼ������Ϣ  �޸�ʱ ����ԭʼ�ļƻ���Ϣ �ڱ���ʱ �Ƚ���Ӧ��Ϣ
Private Enum mPgIndex
    Pg_�ƻ����� = 1
    Pg_�ƻ�ʱ�� = 2
End Enum
Private mbln�Զ�Ĭ����Լ�� As Boolean '45519 �Զ�Ĭ����Լ������
Private mbln�����޸� As Boolean '�Ƿ������޸�
Private mstr��Լ���� As String '������Щ�Ű����Ƹ���
Private mdtMinCustom As Date '�������ԤԼ��,��С��ʱ��

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

Private Sub InitPage()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʼ��ҳ��ؼ�
    '����:���˺�
    '����:2009-09-09 11:01:36
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim i As Long, ObjItem As TabControlItem, objForm As Object
    Err = 0: On Error GoTo Errhand:

    Set ObjItem = tbPage.InsertItem(mPgIndex.Pg_�ƻ�����, "�ƻ�����", picBack.Hwnd, 0)
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

Public Function ShowCard(ByVal mfrmMain As Form, ByVal lngModule As Long, ByVal strPrivs As String, _
    ByVal EditType As mRegEditType, Optional lng����ID As Long, Optional ByVal str�ƻ�Id As String = "") As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��ʾ��Ҫ�޸ĵļƻ�����
    '���:mfrmMain-���õ�������
    '     lngModule-ģ���
    '     strPrivs-Ȩ�޴�
    '     EditType-�༭������
    '     lng����ID-�ҺŰ���ID.
    '     str�ƻ�Id-����ʱΪ��,����,����Ϊָ���ļƻ�ID
    '����:
    '����:�ɹ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-09-14 14:31:59
    '---------------------------------------------------------------------------------------------------------------------------------------------
    mEditType = EditType: mlngModule = lngModule: mstrPrivs = strPrivs: mstr�ƻ�ID = str�ƻ�Id: mblnSucces = False: mlng����ID = lng����ID
    Me.Show 1, mfrmMain
    ShowCard = mblnSucces
End Function

Private Function LoadData(Optional blnNoChangeTime As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ؼƻ�����������Ϣ
    '����:���˺�
    '����:2009-09-14 14:40:46
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp          As New ADODB.Recordset
    Dim rsDept          As New ADODB.Recordset
    Dim strSQL          As String
    Dim i               As Long
    Dim j               As Long
    Dim rs�޺�          As ADODB.Recordset
    Dim strTemp         As String
    Dim blnÿ��         As Boolean
    Dim bln�޺�         As Boolean
    Dim str�޺�         As String
    Dim bln��Լ         As Boolean
    Dim str��Լ         As String
    Dim dtSys           As Date
    Dim dtTmp           As Date
    Dim blnExitFor      As Boolean
    Err = 0: On Error GoTo Errhand:
    
    '���ذ���
    If mEditType = ed_�ƻ����� Then
       '��������
        strSQL = " " & _
        "   Select A.Id as ����ID,0 as �ƻ�ID,A.����,A.��ĿID as �ƻ���ĿID,   A.����,  A.����id,  A.��Ŀid, A.ҽ������,  A.ҽ��id ,   " & _
        "          A.����,  A.��һ,  A.�ܶ�,  A.����,  A.����,  A.����,  A.����,A.Ĭ��ʱ�μ�� As Ĭ��ʱ�μ��, " & _
        "           A.��������,  A.���﷽ʽ,  A.��ſ���,  A.��ʼʱ��,  A.��ֹʱ��,B.���� As ��Ŀ,D.���� As ����,NULL��as ��Чʱ��,'3000-01-01 00:00:00' as ʧЧʱ�� ," & _
        "           NULL as ������,NULL as ����ʱ��,NULL �����,NULL ���ʱ��" & _
        "   From �ҺŰ��� A,�շ���ĿĿ¼ B,�ҺŰ��żƻ� C,���ű� D " & _
        "   Where A.Id=C.����ID(+) And A.��Ŀid=b.Id(+) And A.����id =d.Id(+) " & _
        "         And A.Id=[1]"
    Else
         '������
        strSQL = " " & _
        "Select a.����id, a.Id As �ƻ�id, a.����, �ƻ���Ŀid, a.����, a.����id, a.��Ŀid, a.ҽ������, a.ҽ��id,   a.����, a.��һ, a.�ܶ�, a.����," & _
        "  a.����, a.����, a.����, a.��������, a.���﷽ʽ, a.��ſ���, a.��ʼʱ��, a.��ֹʱ��, b.���� As ��Ŀ, d.���� As ����, ��Чʱ��, a.ʧЧʱ��, a.������, a.����ʱ��," & _
        " a.����� , ���ʱ��,A.Ĭ��ʱ�μ�� As Ĭ��ʱ�μ��" & _
        " From (Select c.����id, c.Id, a.����, Nvl(c.��Ŀid, a.��Ŀid) As �ƻ���Ŀid, c.����, a.����id, Nvl(c.��Ŀid, a.��Ŀid) As ��Ŀid, C.ҽ������, C.ҽ��id," & _
        "       c.����, c.��һ, c.�ܶ�, c.����, c.����, c.����, c.����, a.��������, c.���﷽ʽ, c.��ſ���, a.��ʼʱ��, a.��ֹʱ��, Nvl(C.Ĭ��ʱ�μ��,5) as Ĭ��ʱ�μ��," & _
        "      To_Char(c.��Чʱ��, 'yyyy-mm-dd hh24:mi:ss') As ��Чʱ��, To_Char(c.ʧЧʱ��, 'yyyy-mm-dd hh24:mi:ss') As ʧЧʱ��, c.������," & _
        "      To_Char(c.����ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ����ʱ��, c.�����, To_Char(c.���ʱ��, 'yyyy-mm-dd hh24:mi:ss') As ���ʱ��" & _
        " From �ҺŰ��� A, �ҺŰ��żƻ� C " & _
        " Where a.Id = c.����id) A, �շ���ĿĿ¼ B, ���ű� D " & _
        " Where a.��Ŀid = b.Id(+) And a.����id = d.Id(+) " & _
        "  and a.id=[2]"
    End If
    
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, Val(mstr�ƻ�ID))
    If rsTemp.EOF Then
        If mEditType = ed_�ƻ����� Then
            MsgBox "ע��:" & vbCrLf & _
                   "    �ҺŰ��ſ����Ѿ�������ɾ��,�����ٽ��мƻ�����", vbInformation + vbOKOnly, gstrSysName
        Else
            MsgBox "ע��:" & vbCrLf & _
                   "    �Һżƻ����ſ����Ѿ�������ɾ��,����!", vbInformation + vbOKOnly, gstrSysName
        End If
        Exit Function
    End If
    If mEditType = ed_�ƻ����� Then
        strSQL = "Select ������Ŀ,�޺���,  ��Լ�� From  �ҺŰ������� where ����ID=[1]       "
    Else
        strSQL = "Select ������Ŀ,�޺���,  ��Լ�� From  �Һżƻ����� where �ƻ�ID=[2]       "
    End If
    Set rs�޺� = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, Val(mstr�ƻ�ID))
    
    chkAppoint.Value = 0
    Do While Not rs�޺�.EOF
        If IsNull(rs�޺�!��Լ��) Then
            mblnChangeByCode = True
            chkAppoint.Value = 1
            mblnChangeByCode = False
            Exit Do
        Else
            If Val(Nvl(rs�޺�!��Լ��)) <> 0 Then
                mblnChangeByCode = True
                chkAppoint.Value = 1
                mblnChangeByCode = False
                Exit Do
            End If
        End If
        rs�޺�.MoveNext
    Loop
    If rs�޺�.RecordCount <> 0 Then rs�޺�.MoveFirst
    
    '�������һЩ����
    If mEditType = Ed_�����޸� And Nvl(rsTemp!���ʱ��) <> "" Then
        mblnSaveMinorChange = True
    Else
        mblnSaveMinorChange = False
    End If
    If mEditType = Ed_����ɾ�� And Nvl(rsTemp!���ʱ��) <> "" Then
            MsgBox "ע��:" & vbCrLf & _
                   "    �Һżƻ������Ѿ����������,�����ٽ��мƻ�ɾ����", vbInformation + vbOKOnly, gstrSysName
            Exit Function
    End If
    
    If mEditType = Ed_������� And Nvl(rsTemp!���ʱ��) <> "" Then
            MsgBox "ע��:" & vbCrLf & _
                   "    �Һżƻ������Ѿ����������,�����ٽ��мƻ���ˣ�", vbInformation + vbOKOnly, gstrSysName
            Exit Function
    End If

    If mEditType = Ed_����ȡ�� And Nvl(rsTemp!���ʱ��) = "" Then
            MsgBox "ע��:" & vbCrLf & _
                   "    �Һżƻ������Ѿ�������ȡ�����,�����ٽ��мƻ����ȡ����", vbInformation + vbOKOnly, gstrSysName
            Exit Function
    End If
    
    '�������ݵ��ؼ���
    txt�ű�.Text = Nvl(rsTemp!����)
    cbo����.AddItem Nvl(rsTemp!����): cbo����.ListIndex = cbo����.NewIndex
    chk��ſ���.Value = IIf(Val(Nvl(rsTemp!��ſ���)) = 1, 1, 0)
    
    mTimeSet.bln��ſ��� = Val(rsTemp!��ſ���) = 1
    
    '��ȡ�İ��Ż��߼ƻ��Ƿ���ſ���
    mPlanInfo.bln��� = IIf(Val(Nvl(rsTemp!��ſ���)) = 1, True, False)
    
    chk����.Value = IIf(Val(Nvl(rsTemp!��������)) = 1, 1, 0)
    
    
    txtEdit(midxTxt.idx_������).Text = Nvl(rsTemp!������)
    txtEdit(midxTxt.idx_����ʱ��).Text = Nvl(rsTemp!����ʱ��)
    If mEditType = ed_�ƻ����� Then
        txtEdit(midxTxt.idx_������) = UserInfo.����
        txtEdit(midxTxt.idx_����ʱ��) = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    End If
    txtEdit(midxTxt.idx_�����) = Nvl(rsTemp!�����)
    txtEdit(midxTxt.idx_���ʱ��) = Nvl(rsTemp!���ʱ��)
    If mEditType = Ed_������� Then
        txtEdit(midxTxt.idx_�����) = UserInfo.����
        txtEdit(midxTxt.idx_���ʱ��) = Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS")
    End If
    
    With cbo����
        .AddItem Nvl(rsTemp!����): .ItemData(.NewIndex) = Val(Nvl(rsTemp!����ID)): .ListIndex = .NewIndex
    End With
    With cboItem
         If mEditType = Ed_�����޸� Or mEditType = ed_�ƻ����� Then
            zlControl.CboSetText cboItem, rsTemp!��Ŀ
        Else
            .AddItem Nvl(rsTemp!��Ŀ): .ItemData(.NewIndex) = Val(Nvl(rsTemp!��ĿID)): .ListIndex = .NewIndex
        End If
         
    End With
    With cboDoctor
       If mEditType = ed_�ƻ����� Or mEditType = Ed_�����޸� Then
          LoadDoctor
          zlControl.CboLocate cboDoctor, Nvl(rsTemp!ҽ������)
'          cboDoctor.Text = Nvl(rsTemp!ҽ������)
        Else
            .AddItem Nvl(rsTemp!ҽ������): .ItemData(.NewIndex) = Val(Nvl(rsTemp!ҽ��ID)): .ListIndex = .NewIndex
        End If
    End With
   ' mstr��Լ���� = Get��Լ����(mlng����ID)
    'mbln�����޸� = CheckExistsBooking(Nvl(rsTemp!����), mdtMinCustom)
    
 
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
                !ID = Val(mstr�ƻ�ID)
                !������Ŀ = Nvl(rs�޺�!������Ŀ)
                !�޺��� = Val(Nvl(rs�޺�!�޺���))
                !��Լ�� = Val(Nvl(rs�޺�!��Լ��))
                !��ſ��� = Val(Nvl(rsTemp!��ſ���))
                .Update
            End With
            rs�޺�.MoveNext
        Loop
    End With
    
    If blnNoChangeTime = False Then
        '-------------------------------
        dtSys = zlDatabase.Currentdate
        If mEditType = Ed_�����޸� Or mEditType = ed_�ƻ����� Then
           dtpBegin.MinDate = dtSys
             'Ĭ����һ����Ч
           dtSys = DateAdd("d", 1, Format(dtSys, "yyyy-mm-dd"))
        End If
        If IsNull(rsTemp!��Чʱ��) Then
            dtpBegin.Value = Format(zlGetNextWeekDate, "yyyy-mm-dd HH:MM:SS")
        Else
            If mEditType = Ed_�����޸� Or mEditType = ed_�ƻ����� Then
                '59754
                dtpBegin.Value = IIf(Format(dtSys, "yyyy-mm-dd HH:MM:SS") > Format(CDate(Nvl(rsTemp!��Чʱ��, "1900-01-01")), "yyyy-mm-dd HH:MM:SS"), dtSys, CDate(Nvl(rsTemp!��Чʱ��)))
            Else
                dtpBegin.Value = CDate(Nvl(rsTemp!��Чʱ��))
            End If
        End If
        dtpEndDate.Value = CDate(Nvl(rsTemp!ʧЧʱ��, "3000-01-01"))
        mdatOriBegin = CDate(Nvl(rsTemp!��Чʱ��, "2000-01-01"))
        mdatOriEnd = CDate(Nvl(rsTemp!ʧЧʱ��, "3000-01-01"))
        
        If mEditType = ed_�ƻ����� Then
            strSQL = "Select nvl(��Чʱ��,Sysdate) as ��Чʱ�� ,nvl(ʧЧʱ��,to_date('3000-01-01','yyyy-mm-dd')) as ʧЧʱ�� From �ҺŰ��żƻ� where ID=(Select Max(ID) From �ҺŰ��żƻ� where ����ID=[1]) "
            Set rsDept = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
            If Not rsDept.EOF Then
                If Format(rsDept!ʧЧʱ��, "yyyy-mm-dd") < "3000-01-01" Then
                    '��һ���ƻ�����ֹ����,���Ǳ�������Чʱ��
                    dtTmp = CDate(Format(rsDept!ʧЧʱ��, "yyyy-mm-dd HH:MM:SS"))
                    '59754
                    dtpBegin.Value = IIf(Format(dtSys, "yyyy-mm-dd HH:MM:SS") > Format(dtTmp, "yyyy-mm-dd HH:MM:SS"), dtSys, dtTmp)
                Else '����һ������Чʱ�����һ��Ϊ׼
                    dtTmp = zlGetNextWeekDate(Format(rsDept!��Чʱ��, "yyyy-mm-dd HH:MM:SS"))
                    '�����Ӽƻ�ʱ,����һ�����ڿ�ʼ����
                    dtSys = zlGetNextWeekDate(Format(DateAdd("d", -1, dtSys), "yyyy-mm-dd"))
                     '59754
                    dtpBegin.Value = IIf(Format(dtSys, "yyyy-mm-dd HH:MM:SS") > Format(dtTmp, "yyyy-mm-dd HH:MM:SS"), dtSys, dtTmp)
                End If
            End If
            
            strSQL = "Select �ű�ID as ID,�������ҡ�From �ҺŰ������� Where �ű�ID=[1]"
            Set rsDept = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
        Else
            strSQL = "Select �ƻ�ID as ID,�������ҡ�From �Һżƻ����� Where �ƻ�ID=[2]"
            Set rsDept = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, Val(mstr�ƻ�ID))
        End If
    End If
    
    Call LoadLongPlan
    Call LoadRegHistory
    '---------------------------------------------------
    '�ж� ÿ�հ��� �޺��� ��Լ�� ���Ƿ�һ��
    '---------------------------------------------------
    rs�޺�.Filter = 0
    If rs�޺�.RecordCount > 0 Then rs�޺�.MoveFirst
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
    mstrԭ�Ű� = ""
    If blnÿ�� Or mrsRegHistory.RecordCount > 0 Then
        'ÿ��
        opt��.Value = True:
        txt�޺�.Enabled = False: txt��Լ.Enabled = False
        With vsPlan
            For i = 1 To 7
                '��֪ʲôԭ��,��.colkey(i)����,Ҫ���ĳ�������.
                strTemp = "��" & Replace(.ColKey(i), "����", "��")
                .TextMatrix(1, i) = Nvl(rsTemp.Fields(strTemp))
                mstrԭ�Ű� = mstrԭ�Ű� & ";" & strTemp & "," & Nvl(rsTemp.Fields(strTemp)) '"��һ,����;�ܶ�,����;..."
                rs�޺�.Filter = "������Ŀ='" & strTemp & "'"
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
            Next
            If mstrԭ�Ű� <> "" Then mstrԭ�Ű� = Mid(mstrԭ�Ű�, 2)
        End With
    Else
        'ÿ��
        opt��.Value = True:  cbo��.ListIndex = cbo.FindIndex(cbo��, Nvl(rsTemp!����), True): cbo��.Enabled = True
        mstrԭ�Ű� = Nvl(rsTemp!����)
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
    '��ȡ�޸Ļ�������ǰ�� ʱ��κ� �޺���
    '�����ڱ���ʱ �Ա��޺���Լ����ſ����Լ�ʱ����Ƿ����˱仯
    '��������˱仯����Ҫ��ʾ  ����Ա��������ʱ����Ϣ
    '------------------------------
   mPlanInfo.str�Ű� = ""
   mPlanInfo.str�޺� = ""
   mPlanInfo.strӦ��ʱ�� = ""
    If blnÿ�� = False Or mrsRegHistory.RecordCount > 0 Then
        For i = 1 To 7
             mPlanInfo.str�Ű� = mPlanInfo.str�Ű� & ",'" & Trim(cbo��.Text) & "'"
             mPlanInfo.strӦ��ʱ�� = mPlanInfo.strӦ��ʱ�� & "|" & Switch(i = 1, "����", i = 2, "��һ", i = 3, "�ܶ�", i = 4, "����", i = 5, "����", i = 6, "����", True, "����") & "-" & Trim(cbo��.Text)
             mPlanInfo.str�޺� = mPlanInfo.str�޺� & "|" & Switch(i = 1, "����", i = 2, "��һ", i = 3, "�ܶ�", i = 4, "����", i = 5, "����", i = 6, "����", True, "����")
             mPlanInfo.str�޺� = mPlanInfo.str�޺� & "," & Val(txt�޺�.Text) & "," & txt��Լ.Text
             mPlanInfo.strӦ��ʱ�� = mPlanInfo.strӦ��ʱ�� & "|" & Switch(i = 1, "����", i = 2, "��һ", i = 3, "�ܶ�", i = 4, "����", i = 5, "����", i = 6, "����", True, "����") & "-" & Trim(vsPlan.TextMatrix(1, i))
             
        Next
    Else
        For i = 1 To vsPlan.Cols - 1
            mPlanInfo.str�Ű� = mPlanInfo.str�Ű� & ",'" & Trim(vsPlan.TextMatrix(1, i)) & "'"
            If Trim(vsPlan.TextMatrix(1, i)) <> "" Then
                mPlanInfo.strӦ��ʱ�� = mPlanInfo.strӦ��ʱ�� & "|" & Switch(i = 1, "����", i = 2, "��һ", i = 3, "�ܶ�", i = 4, "����", i = 5, "����", i = 6, "����", True, "����") & "-" & Trim(vsPlan.TextMatrix(1, i))
                mPlanInfo.str�޺� = mPlanInfo.str�޺� & "|" & Switch(i = 1, "����", i = 2, "��һ", i = 3, "�ܶ�", i = 4, "����", i = 5, "����", i = 6, "����", True, "����")
                If Trim(vsPlan.TextMatrix(1, i)) = "" Then
                     mPlanInfo.str�޺� = mPlanInfo.str�޺� & ",0,0"
                Else
                     mPlanInfo.str�޺� = mPlanInfo.str�޺� & "," & Val(Trim(vsPlan.TextMatrix(2, i))) & "," & Trim(vsPlan.TextMatrix(3, i))
                End If
            End If
        Next
    End If
    If mPlanInfo.str�޺� <> "" Then mPlanInfo.str�޺� = Mid(mPlanInfo.str�޺�, 2)
    If mPlanInfo.strӦ��ʱ�� <> "" Then mPlanInfo.strӦ��ʱ�� = Mid(mPlanInfo.strӦ��ʱ��, 2)
    
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
    
    '71253 ���ϴ� 2014-04-15 14:23:10 ��listView �滻ΪvsflexGrid
    If blnNoChangeTime = False Then
    With vsDept
        blnExitFor = False
        Do While Not rsDept.EOF
            For i = 0 To .Cols - 1
                For j = 0 To .Rows - 1
                    If Nvl(rsDept!��������) = .TextMatrix(j, i) Then
                        .Cell(flexcpChecked, j, i) = 1
                        blnExitFor = True
                        Exit For
                    End If
                Next
                If blnExitFor Then blnExitFor = False: Exit For
            Next
            rsDept.MoveNext
        Loop
    End With
    rsDept.Close
    End If
    
    If mEditType = ed_�ƻ����� Or mEditType = Ed_�����޸� Then mPlanInfo.blnʱ��� = Checkʱ��()
    If mrsRegHistory.RecordCount > 0 Then opt��.Enabled = False
    If mEditType = Ed_����ɾ�� Then
        picTimeSet.Enabled = False
    End If
    mdatBegin = dtpBegin
    mdatEnd = dtpEndDate
    LoadData = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Sub chkAppoint_Click()
    Dim i As Integer
    If mblnChangeByCode Then Exit Sub
    If chkAppoint.Value = 0 Then
        If opt��.Value = True Then
            txt��Լ.Enabled = False
            txt��Լ.BackColor = &H8000000F
        End If
        txt��Լ.Text = ""
        For i = 1 To vsPlan.Cols - 1
            vsPlan.TextMatrix(3, i) = ""
        Next i
    Else
        If opt��.Value = True Then
            txt��Լ.Enabled = True
            txt��Լ.BackColor = vbWhite
        End If
        If Val(txt��Լ.Text) = 0 Then txt��Լ.Text = ""
        For i = 1 To vsPlan.Cols - 1
            If Val(vsPlan.TextMatrix(3, i)) = 0 Then vsPlan.TextMatrix(3, i) = ""
        Next i
    End If
End Sub

Private Function InitData() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���س�ʼ������
    '����:���˺�
    '����:2009-09-14 15:50:31
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsTemp As New ADODB.Recordset, i As Long
    Dim intRow As Integer
    Dim lngColsWidth As Long
    
    Err = 0: On Error GoTo Errhand:

    strSQL = "Select '    ' ʱ��� From dual Union All  " & _
             " Select ʱ��� From ʱ���"
        
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    If rsTemp.RecordCount <> 0 Then rsTemp.MoveFirst
    Do While Not rsTemp.EOF
        cbo��.AddItem rsTemp!ʱ���
        rsTemp.MoveNext
    Loop
    
    With vsPlan
        .ColComboList(1) = .BuildComboList(rsTemp, "ʱ���")
        For i = 2 To .Cols - 1
            .ColComboList(i) = .ColComboList(1)
        Next
        .Tag = .ColComboList(1)
    End With
 
    
    '��������
    strSQL = "Select ����,���ơ�From �������� Where (վ��='" & gstrNodeNo & "' Or վ�� is Null) Order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)

    '71253 ���ϴ� 2014-04-15 14:23:10 ����������ʾ��ȫ
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
            If lngColsWidth > .ClientWidth Then .Height = .Height + 230: Frame3.Height = Frame3.Height + 150
            .Editable = flexEDKbdMouse
        End With
    End If
 
    '�Һ���Ŀ
    If mEditType = Ed_�����޸� Or mEditType = ed_�ƻ����� Then
        strSQL = "Select ID as ���,���� From �շ���ĿĿ¼ " & _
            " Where ���='1' And (Sysdate Between ����ʱ�� And ����ʱ�� Or ����ʱ��<Sysdate And ����ʱ�� Is Null)" & _
            " And (վ��='" & gstrNodeNo & "' Or վ�� is Null)" & _
            " Order by ����"
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption)
    
        If rsTemp.EOF Then
            MsgBox "û�п��õĹҺ���Ŀ��Ϣ,���ȵ��Һ���Ŀ�����г�ʼ��", vbInformation, gstrSysName
            Exit Function
        End If
    
        cboItem.Clear
        For i = 1 To rsTemp.RecordCount
            cboItem.AddItem rsTemp!����
            cboItem.ItemData(cboItem.NewIndex) = rsTemp!���
            rsTemp.MoveNext
        Next
    End If
    
    'cmdCancel.Caption = "�˳�(&X)"
    If mEditType = Ed_������� Then
        Me.Caption = Me.Caption & "�������"
    ElseIf mEditType = Ed_����ɾ�� Then
        Me.Caption = Me.Caption & "����ɾ��"
        'cmdOK.Caption = "ɾ��(&D)"
    ElseIf mEditType = Ed_����ȡ�� Then
        Me.Caption = Me.Caption & "����ȡ�����"
    ElseIf mEditType = ed_���Ų��� Then
        cmdOK.Visible = False
        cmdCancel.Top = cmdOK.Top
    End If
    
    InitData = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
    SaveErrLog
End Function

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
        MsgBox "��ǰ�ű���" & str������Ŀ & ",û�жԹҺ�����������,�޷�����ʱ��,����!", vbInformation, Me.Caption
        Exit Function
    End If

    strʱ�� = mTimeSet.rsRegPlan!�Ű�
    mrs�ϰ�ʱ���.Filter = "ʱ���='" & strʱ�� & "'"
    If mrs�ϰ�ʱ���.RecordCount = 0 Then
        MsgBox "������ʱ��Ϊ[" & strʱ�� & "]�����°�ʱ��,����!", vbInformation, Me.Caption
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

Private Sub dtpBegin_Validate(Cancel As Boolean)
    Dim strStartTime As String
    
    If Format(dtpEndDate.Value, "YYYY-MM-DD") <> "3000-01-01" Then Exit Sub
    
    If chk������Ч.Value = 1 Then
        strStartTime = Format(zlDatabase.Currentdate, "yyyy-MM-dd hh:mm:ss")
    Else
        strStartTime = Format(dtpBegin.Value, "yyyy-MM-dd hh:mm:ss")
    End If
    
    If mEditType = ed_�ƻ����� Or mEditType = Ed_�����޸� Then
        If Not mrsLongPlan Is Nothing Then
            If mrsLongPlan.RecordCount > 0 Then
                If Format(Nvl(mrsLongPlan!��Чʱ��), "yyyy-MM-dd hh:mm:ss") > strStartTime Then
                    dtpEndDate.Value = CDate(Nvl(mrsLongPlan!��Чʱ��, "3000-01-01"))
                End If
            End If
        End If
    End If
End Sub

Private Sub Form_Load()
    Call InitPage
    opt��Чʱ��(0).Enabled = True: opt��Чʱ��(1).Enabled = True
    mblnFirst = True
    Call LoadTimeSetControl
End Sub

Private Sub Form_Activate()
    If mblnFirst = False Then Exit Sub
    mblnFirst = False
    Call LoadvsDept
    If InitData = False Then Unload Me: Exit Sub
    If LoadData = False Then Unload Me: Exit Sub
    Call SetCtrlEnabled
    If IsValidation() = False Then Unload Me: Exit Sub
    If mEditType = ed_�ƻ����� Or mEditType = Ed_�����޸� Then
        zlControl.ControlSetFocus chk��ſ���
    Else
        zlControl.ControlSetFocus cmdOK
    End If
    If mblnSaveMinorChange Then
        tbPage.Item(1).Visible = False
        txt�ű�.Enabled = False
        chk��ſ���.Enabled = False
        chk����.Enabled = False
        Frame2.Enabled = False
        Frame1.Enabled = False
        opt��.Enabled = False
        opt��.Enabled = False
        cbo��.Enabled = False
        txt�޺�.Enabled = False
        txt��Լ.Enabled = False
        cbo����.Enabled = False
        cboDoctor.Enabled = False
        cbo����.Enabled = False
        vsPlan.Enabled = False
        cboItem.Enabled = False
        chk������Ч.Visible = False
        vsPlan.HighLight = flexHighlightNever
        opt��Чʱ��(0).Enabled = False
        opt��Чʱ��(1).Enabled = False
        dtpBegin.Enabled = False
        dtpEndDate.Enabled = False
        chk�������.Visible = False
        dtpBegin.MinDate = mdatOriBegin
        dtpBegin.Value = mdatOriBegin
        dtpEndDate.MaxDate = mdatOriEnd
        dtpEndDate.Value = mdatOriEnd
    End If
End Sub

Private Sub SaveMinorChange()
    Dim strSQL As String, intCount As Integer
    Dim str���� As String
    Dim i As Long
    Dim j As Long
    Dim rsTemp As ADODB.Recordset
    Dim intSync As Integer
    Dim lngִ�мƻ�ID As Long
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
    
    With vsDept
        For i = 0 To .Cols - 1
            For j = 0 To .Rows - 1
                If .Cell(flexcpChecked, j, i) = 1 Then str���� = str���� & ";" & .TextMatrix(j, i)
            Next
        Next
    End With
    str���� = Mid(str����, 2)
    
    strSQL = "Zl_�ҺŰ��żƻ�_Modify("
    strSQL = strSQL & mstr�ƻ�ID & ",'"
    strSQL = strSQL & str���� & "',"
    strSQL = strSQL & intCount & ")"
    Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
    
    '112585��л�٣��޸�����˵ļƻ���δ��ʱ���ĹҺŰ��ŵ�������Ϣ
    If txtEdit(3).Text <> "" Then
        '��⵱ǰ�޸ļƻ���ID�Ƿ�Ϊ�ҺŰ��ŵ�ִ�мƻ�ID
        gstrSQL = "Select ִ�мƻ�ID From �ҺŰ��� Where ID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng����ID)
        lngִ�мƻ�ID = Val("" & rsTemp!ִ�мƻ�ID)
        If lngִ�мƻ�ID = Val(mstr�ƻ�ID) Then
            '�����������Ч�����¹ҺŰ��ŵ�������Ϣ
            strSQL = "Zl_�ҺŰ���_Modify("
            strSQL = strSQL & mlng����ID & ",'"
            strSQL = strSQL & str���� & "',"
            strSQL = strSQL & "Null,"
            strSQL = strSQL & intCount & ")"
            Call zlDatabase.ExecuteProcedure(strSQL, Me.Caption)
        End If
    End If
    Unload Me
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then Call zlCommFun.PressKey(vbKeyTab)
End Sub
Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = Asc("'") Then KeyAscii = 0
End Sub

Private Sub SetCtrlEnabled()
    '���ÿؼ���Enabled����
    Dim ctl As Control
    For Each ctl In Me.Controls
        Select Case UCase(TypeName(ctl))
        Case "TEXTBOX"
            ctl.Enabled = False
            '�޸Ļ��������ƻ�ʱ �����޺š���Լ�ı��� ���޸�
            If ctl Is Me.txt�޺� Or ctl Is txt��Լ Or ctl Is txtTimeOut Then
               ctl.Enabled = mEditType = Ed_�����޸� Or mEditType = ed_�ƻ�����
            End If
        Case UCase("ComboBox")
            If ctl Is cbo�� And mEditType = ed_�ƻ����� Then
                   ctl.Enabled = opt��.Value = 1
              ElseIf ctl Is cboItem Or ctl Is cboDoctor Then
                 '-----------------------------------------------------
                 'Ϊ�޸Ļ��� ����ģʽʱ ���Ŷ� ��Ŀ��ҽ���ĸ���

                 '------------------------------------------------------
                   If mEditType = ed_�ƻ����� Or mEditType = Ed_�����޸� Then
                       ctl.Enabled = True
                   Else
                       ctl.Enabled = False
                   End If
               Else:
                   ctl.Enabled = False
               End If
        Case UCase("ListView")
            ctl.Enabled = False
        Case UCase("DTPicker")
            ctl.Enabled = False
        Case UCase("optionbutton"), UCase("CheckBox")
            ctl.Enabled = False
            If (ctl.Name = "opt��Чʱ��" Or ctl.Name = "optӦ����" Or ctl.Name = "opt����" Or ctl.Name = "opt����" Or ctl Is chkAppoint) And (mEditType = Ed_�����޸� Or mEditType = ed_�ƻ�����) Then
               ctl.Enabled = True: ctl.Visible = True
            End If
        Case Else
        End Select
    Next
    
    Select Case mEditType
    Case ed_�ƻ�����, Ed_�����޸�
        chk��ſ���.Enabled = True
        txt�޺�.Enabled = IIf(opt��.Value = True, True, False): txt��Լ.Enabled = IIf(opt��.Value = True And chkAppoint.Value = 1, True, False)
        cbo��.Enabled = IIf(opt��.Value = True, True, False)
        dtpBegin.Enabled = IIf(opt��Чʱ��(1).Value = 1, True, False)
        dtpEndDate.Enabled = True
        vsDept.Enabled = True
        chkAppoint.Enabled = True
        opt����(0).Enabled = True: opt����(1).Enabled = True: opt����(2).Enabled = True: opt����(3).Enabled = True
        opt��.Enabled = True: opt��.Enabled = True
        dtpBegin.Enabled = True:
        
        '�Է����������:
        '   ָ��ҽ��ʱ���������ó�,��̬�����ƽ������
        If Trim(cboDoctor.Text) <> "" Then
            opt����(2).Enabled = False: opt����(3).Enabled = False
            If opt����(2).Value Or opt����(3).Value Then opt����(0).Value = True
        Else
            opt����(2).Enabled = True: opt����(3).Enabled = True
        End If
        If opt��.Value = True Then cbo��.Enabled = True
        chk������Ч.Enabled = False: chk������Ч.Visible = False
        chk�������.Enabled = True
    Case Ed_�������
        chk�������.Enabled = False: chk�������.Visible = False
        chk������Ч.Enabled = True: chk������Ч.Visible = True
    Case Else
    End Select
    
    '���ñ༭����ɫ
    For Each ctl In Me.Controls
        Select Case UCase(TypeName(ctl))
        Case "TEXTBOX", UCase("ComboBox")
            Call zlSetCtrolBackColor(ctl)
        Case UCase("ListView")
        Case UCase("DTPicker")
        Case Else
        End Select
    Next
    
End Sub
 
Private Sub chk������Ч_Click()
'    dtpBegin.Enabled = chk������Ч.Value = 0
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    ShowHelp App.ProductName, Me.Hwnd, Me.Name
End Sub
Private Function CheckPlanValied() As Boolean
    '------------------------------------------------------------------------------------------------------------------------
    '���ܣ����ƻ��ĺϷ���
    '���أ��ƻ����źϷ�,����True,���򷵻�False
    '���ƣ����˺�
    '���ڣ�2010-07-21 17:49:30
    '˵����
    '------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As New ADODB.Recordset
    Dim strSQL As String
    If mEditType <> Ed_�����޸� And mEditType <> ed_�ƻ����� Then
        CheckPlanValied = True: Exit Function
    End If
    
    If dtpBegin.Value > dtpEndDate.Value Then
        ShowMsgbox "ע��:" & vbCrLf & "    ��Чʱ��С����ʧЧʱ��,����!"
        If dtpEndDate.Enabled And dtpEndDate.Visible Then dtpEndDate.SetFocus
        Exit Function
    End If
    If zlDatabase.Currentdate > dtpBegin.Value Then
        ShowMsgbox "ע��:" & vbCrLf & "    ��Чʱ��С���˵�ǰϵͳʱ��,����!"
        If dtpBegin.Enabled And dtpBegin.Visible Then dtpBegin.SetFocus
        Exit Function
    End If
    Set rsTemp = Nothing
     CheckPlanValied = True: Exit Function
End Function

Private Function IsValied(Optional ByVal blnSave As Boolean = False) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�����������ݵĺϷ���
    '����:���ݺϷ�,����true,���򷵻�False
    '����:���˺�
    '����:2009-09-14 16:31:50
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim rsTemp As ADODB.Recordset, strSQL As String, i As Long, intCount As Integer
    Dim strTmp As String, lngҽ��ID As Long, lng��Լ�� As Long
    Dim j As Integer
    Dim str������Ŀ As String
    
    Err = 0: On Error GoTo Errhand:
    If Trim(txt�ű�) = "" Then
        MsgBox "�ű���Ϊ�գ�", vbInformation, gstrSysName
        txt�ű�.SetFocus: Exit Function
    End If
    If cbo����.ListIndex = -1 Then
        MsgBox "δ���úű�����Ӧ�Ŀ��ң�", vbInformation, gstrSysName
        cbo����.SetFocus: Exit Function
    End If
    If cboItem.ListIndex = -1 Then
        MsgBox "δ���úű�����Ӧ�ĹҺ���Ŀ��", vbInformation, gstrSysName
        cboItem.SetFocus: Exit Function
    End If
    
    If opt��.Value Then
        If cbo��.ListIndex = -1 Then
            MsgBox "�úű�ÿ���Ӧ��ʱ��δ���ã�", vbInformation, gstrSysName
            If txt�޺�.Enabled Then txt�޺�.SetFocus
            Exit Function
        End If
        If chk��ſ���.Value = 1 Then
            If Val(txt�޺�.Text) = 0 And Val(txt��Լ.Text) = 0 Then
                MsgBox "ʹ����ſ���ʱ,���������޺Ż���Լ����", vbInformation, gstrSysName
                If txt�޺�.Enabled Then txt�޺�.SetFocus
                Exit Function
            End If
        Else
            If Not blnSave Then
                If chkAppoint.Value = 0 Or (chkAppoint.Value = 1 And txt��Լ.Text <> "" And Val(txt��Լ.Text) = 0) Then
                    MsgBox "����ſ��Ƶİ�������ʱ��ʱ,�����ǿ�ԤԼ�İ��ţ�", vbInformation, gstrSysName
                    If txt�޺�.Enabled Then txt�޺�.SetFocus: Exit Function
                End If
            End If
        End If
        '�޺���Լ����
        If Trim(txt�޺�.Text) <> "" Then
            If Trim(txt��Լ.Text) <> "" And Val(txt�޺�.Text) < Val(txt��Լ.Text) Then
                MsgBox "��Լ��ӦС���޺�����", vbInformation, gstrSysName
               If txt��Լ.Enabled Then txt��Լ.SetFocus
                Exit Function
            End If
        ElseIf Trim(txt��Լ.Text) <> "" Then
            MsgBox "��Լ�����޺ţ�", vbInformation, gstrSysName
            If txt�޺�.Enabled Then txt�޺�.SetFocus
            Exit Function
        End If
    Else
        If Not blnSave Then
            If chkAppoint.Value = 0 And chk��ſ���.Value = 0 Then
                MsgBox "����ſ��Ƶİ�������ʱ��ʱ,�����ǿ�ԤԼ�İ��ţ�", vbInformation, gstrSysName
                Exit Function
            End If
        End If
     With vsPlan
            strTmp = ""
            For i = 1 To .Cols - 1
                If Trim(.TextMatrix(1, i)) <> "" Then
                    strTmp = strTmp & Trim(vsPlan.TextMatrix(1, i))
                    If chk��ſ���.Value = 1 Then
                          If Val(.TextMatrix(2, i)) = 0 And Val(.TextMatrix(3, i)) = 0 Then
                              MsgBox "ʹ����ſ���ʱ,���������޺Ż���Լ����", vbInformation, gstrSysName
                              .Row = 2: .Col = i
                              .SetFocus: Exit Function
                          End If
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
    
    If cboDoctor.ListIndex <> -1 Then lngҽ��ID = cboDoctor.ItemData(cboDoctor.ListIndex)
    If lngҽ��ID = 0 And cboDoctor.Text <> "" Then
        strSQL = "Select 1 From ��Ա�� Where ���� = [1] "
        Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, cboDoctor.Text)
        If Not rsTemp.EOF Then
            MsgBox "ҽ��""" & cboDoctor.Text & """�����ڿ���""" & cbo����.Text & """,���������øúű�Ŀ�����ҽ����Ϣ��", vbInformation, gstrSysName
            cboDoctor.SetFocus: Exit Function
        End If
    End If
    
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
                vsDept.SetFocus: Exit Function
            ElseIf intCount > 1 Then
                MsgBox "ָ������ʱֻ��ѡ��һ����Ӧ���������ң�", vbInformation, gstrSysName
                vsDept.SetFocus: Exit Function
            End If
        ElseIf opt����(2).Value Or opt����(3).Value Then
            If intCount < 2 Then
                MsgBox "��̬�����ƽ������ʱ����Ҫѡ��������Ӧ���������ң�", vbInformation, gstrSysName
                vsDept.SetFocus: Exit Function
            End If
        End If
    End If
     
    '��Ŀ�۸��ж�
    If ReadRegistPrice(cboItem.ItemData(cboItem.ListIndex), False, False) = 0 Then
        MsgBox "��Ŀ""" & cboItem.Text & """δ������Ч�۸�,���ȵ��շ���Ŀ���������ã�", vbInformation, gstrSysName
        cboItem.SetFocus: Exit Function
    End If
    If opt��Чʱ��(1).Value = 0 Then
        If Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS") < Format(zlDatabase.Currentdate, "yyyy-mm-dd HH:MM:SS") Then
            ShowMsgbox "��Чʱ�䲻��С�ڵ�ǰϵͳʱ��,����!"
            Exit Function
        End If
    End If
    '�����صļƻ�
    If CheckPlanValied = False Then Exit Function
    If mEditType = ed_�ƻ����� Then
        '�����Ӽƻ�ʱ,���
        If Format(dtpBegin.Value, "yyyy-mm-dd hh:mm:ss") < Format(mdtMinCustom, "yyyy-mm-dd hh:mm:ss") Then
            If MsgBox("�üƻ�����Ч���ں��Ѵ���ԤԼ��,�Ƿ����?", vbYesNo + vbDefaultButton1 + vbInformation, Me.Caption) = vbNo Then
                Exit Function
            End If
        End If
    End If
    
    If CheckUsedCount() = False Then Exit Function
    IsValied = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
    If 1 = 2 Then
        Resume
    End If
End Function

Private Function LoadLongPlan() As Boolean
    Dim strSQL As String
    On Error GoTo errH
    strSQL = "Select ��Чʱ�� From �ҺŰ��żƻ�" & _
            " Where ����id = [1] And ʧЧʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd') And ���ʱ�� Is Not Null"
    Set mrsLongPlan = zlDatabase.OpenSQLRecord(strSQL, "���ڼƻ�", mlng����ID)
    LoadLongPlan = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function LongPlanIsValied(ByRef lng�ϴμƻ�ID As Long) As Boolean
    '---------------------------------------------------------------------------------------
    ' ���� : ��鳤�ڼƻ�����Ч��
    ' ��� :
    ' ���� : lng�ϴμƻ�ID-�������ĳ��ڼƻ�
    ' ���� :
    ' ���� : ���ϴ�
    ' ���� : 2018/11/9 11:29
    ' ���� :133584
    '---------------------------------------------------------------------------------------
    Dim strSQL As String, rsPlan As ADODB.Recordset
    Dim strStartTime As String
    
    On Error GoTo errH
    If Format(dtpEndDate.Value, "YYYY-MM-DD") <> "3000-01-01" Then LongPlanIsValied = True: Exit Function
    
    If chk������Ч.Value = 1 Then
        strStartTime = Format(zlDatabase.Currentdate, "yyyy-MM-dd hh:mm:ss")
    Else
        strStartTime = Format(dtpBegin.Value, "yyyy-MM-dd hh:mm:ss")
    End If
    
    If mEditType = ed_�ƻ����� Or mEditType = Ed_�����޸� Then
        If Not mrsLongPlan Is Nothing Then
            If mrsLongPlan.RecordCount > 0 Then
                If Format(Nvl(mrsLongPlan!��Чʱ��), "yyyy-MM-dd hh:mm:ss") > strStartTime Then
                    MsgBox "���ڼƻ�(" & Format(Nvl(mrsLongPlan!��Чʱ��), "yyyy-MM-dd hh:mm:ss") & "~" & "3000-01-01)����Чʱ��ȱ��ε���Чʱ��������ֻ����Ϊ���ڼƻ���", vbInformation, gstrSysName
                    Exit Function
                End If
            End If
        End If
    End If
    
    If chk�������.Value = 0 And Not mEditType = Ed_������� Then LongPlanIsValied = True: Exit Function
    
    strSQL = "Select 0 as �Ƿ�����, ID, ��Чʱ�� From �ҺŰ��żƻ�" & _
            " Where ����id = [1] And ʧЧʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd') And ��Чʱ�� < To_Date([2],'YYYY-MM-DD hh24:mi:ss') " & _
            " And ���ʱ�� Is Null" & vbNewLine & _
            " Union All" & vbNewLine & _
            " Select 1 as �Ƿ�����, ID, ��Чʱ�� From �ҺŰ��żƻ�" & _
            " Where ����id = [1] And ʧЧʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd') And ���ʱ�� Is Not Null"
    Set rsPlan = zlDatabase.OpenSQLRecord(strSQL, "��鳤�ڼƻ�", mlng����ID, strStartTime)
    If rsPlan.RecordCount = 0 Then LongPlanIsValied = True: Exit Function
    rsPlan.Filter = "�Ƿ����� = " & 0
    If rsPlan.RecordCount > 0 Then
        MsgBox "����δ��˵ĳ��ڼƻ�(" & Format(Nvl(rsPlan!��Чʱ��), "yyyy-MM-dd hh:mm:ss") & "~" & "3000-01-01)����������˻�ɾ����", vbInformation, gstrSysName
        Exit Function
    End If
    rsPlan.Filter = "�Ƿ����� = 1"
    If rsPlan.RecordCount = 0 Then LongPlanIsValied = True: Exit Function
    If Format(Nvl(rsPlan!��Чʱ��), "yyyy-MM-dd hh:mm:ss") < strStartTime Then
        If MsgBox("���ڳ��ڼƻ�(" & Format(Nvl(rsPlan!��Чʱ��), "yyyy-MM-dd hh:mm:ss") & "~" & "3000-01-01), �Ƿ���ʧЧʱ�����Ϊ���ε���Чʱ�䣿", vbQuestion + vbYesNo, gstrSysName) = vbNo Then
            Exit Function
        End If
        lng�ϴμƻ�ID = Val(Nvl(rsPlan!ID))
    Else
        MsgBox "���ڼƻ�(" & Format(Nvl(rsPlan!��Чʱ��), "yyyy-MM-dd hh:mm:ss") & "~" & "3000-01-01)����Чʱ��ȱ��ε���Чʱ�������Ƚ������޸�Ϊ���ڼƻ���", vbInformation, gstrSysName
        Exit Function
    End If
    LongPlanIsValied = True
    Exit Function
errH:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SavePlan(ByVal lng�ϴμƻ�ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:����ƻ�����
    '����:����ɹ�������true,���򷵻�False
    '����:���˺�
    '����:2009-09-14 16:41:22
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, strʱ��� As String, str���� As String, i As Long, int���� As Integer
    Dim lng�ƻ�ID As Long, str�޺� As String
    Dim strҽ������         As String
    Dim strҽ��ID           As String
    Dim blnChange           As Boolean
    Dim bytType             As Byte
    Dim vMsgResult          As VbMsgBoxResult
    Dim strMsg              As String
    Dim colPro              As Collection
    Dim blnTrans            As Boolean
    Dim j                   As Integer
    'bytType 0-����ʱ ��ʱ�β����д��� �޸�ʱ ��ʱ��ֻɾ���Ѿ�ȥ�����Ű���Ϣ
    '        1-����ʱ ��ȡԭ���ŵ�ʱ����Ϣ  �޸�ʱ �Լƻ���ʱ�ν���ɾ��
    
    Err = 0: On Error GoTo Errhand:
    
    strʱ��� = "": str�޺� = ""
    If opt��.Value Then
        For i = 1 To 7
            strʱ��� = strʱ��� & ",'" & Trim(cbo��.Text) & "'"
            str�޺� = str�޺� & "|" & Switch(i = 1, "����", i = 2, "��һ", i = 3, "�ܶ�", i = 4, "����", i = 5, "����", i = 6, "����", True, "����")
            str�޺� = str�޺� & "," & Val(txt�޺�.Text) & "," & IIf(chkAppoint.Value = 0, "0", txt��Լ.Text)
        Next
    Else
        With vsPlan
            For i = 1 To .Cols - 1
                strʱ��� = strʱ��� & ",'" & Trim(.TextMatrix(1, i)) & "'"
                If Trim(.TextMatrix(1, i)) <> "" Then
                    str�޺� = str�޺� & "|" & Switch(i = 1, "����", i = 2, "��һ", i = 3, "�ܶ�", i = 4, "����", i = 5, "����", i = 6, "����", True, "����")
                    str�޺� = str�޺� & "," & Val(Trim(vsPlan.TextMatrix(2, i))) & "," & IIf(chkAppoint.Value = 0, "0", Trim(vsPlan.TextMatrix(3, i)))
                End If
            Next
        End With
    End If
    If str�޺� <> "" Then str�޺� = Mid(str�޺�, 2)
            
    If mPlanInfo.blnʱ��� Then
      '�ж����Ѿ��ı� �ƻ���Ϣ
        'blnChange = (mPlanInfo.str�Ű� <> strʱ���) Or (mPlanInfo.str�޺� <> str�޺�) Or (IIf(mPlanInfo.bln���, 1, 0) <> chk��ſ���.Value)
         blnChange = True
    End If
    '71253 ���ϴ� 2014-04-15 14:23:10 ��listView �滻ΪvsflexGrid
    With vsDept
        For i = 0 To .Cols - 1
            For j = 0 To .Rows - 1
                If .Cell(flexcpChecked, j, i) = 1 Then str���� = str���� & ";" & .TextMatrix(j, i)
            Next
        Next
    End With
    str���� = Mid(str����, 2)
    
    'ȡ���﷽ʽ
    int���� = 0
    For i = 0 To opt����.UBound
        If opt����(i).Value Then int���� = i: Exit For
    Next
     '�����:52275
    '�ڼƻ����߰���������ʱ��ʱ ��ʱ�δ���Ĵ�������
'    If mPlanInfo.blnʱ��� And mEditType = ed_�ƻ����� And blnChange = False Then
'        '���ԭ�ƻ����߰���ʱ ������ʱ�� ��ʾ����ԭ���д���
'        strMsg = "������������ʱ��,�Ƿ���ȡ���ŵ�ʱ����Ϊ�ƻ���ʱ����Ϣ? " & vbCrLf
'        strMsg = strMsg & "[��(Y)]��ȡ���ŵ�ʱ����Ϣ��Ϊ�ƻ���ʱ��" & vbCrLf
'        strMsg = strMsg & "[��(N)]����ȡ���ŵ�ʱ��,��������ʱ��" & vbCrLf
'        vMsgResult = MsgBox(strMsg, vbYesNo + vbQuestion + vbDefaultButton1, Me.Caption)
'        bytType = IIf(vMsgResult = vbYes, 1, 0)
'    End If
    If mEditType = Ed_�����޸� Then
      bytType = IIf(IIf(mPlanInfo.bln���, 1, 0) <> chk��ſ���.Value, 1, 0)
    End If
    'ȡʱ�䷶Χ
    If mEditType = ed_�ƻ����� Then
        lng�ƻ�ID = zlDatabase.GetNextId("�ҺŰ��żƻ�")
    Else
        lng�ƻ�ID = Val(mstr�ƻ�ID)
    End If
     If cboDoctor.ListIndex = -1 Then
        strҽ������ = ""
        strҽ��ID = "0"
     Else
        strҽ������ = cboDoctor.Text
        strҽ��ID = Val(cboDoctor.ItemData(cboDoctor.ListIndex))
     End If
    'Zl_�ҺŰ��żƻ�_Insert
    strSQL = "Zl_�ҺŰ��żƻ�_Insert("
    '  Id_In       In �ҺŰ��żƻ�.ID%Type,
    strSQL = strSQL & "" & lng�ƻ�ID & ","
    '  ����id_In   In �ҺŰ��żƻ�.����id%Type,
    strSQL = strSQL & "" & mlng����ID & ","
    '  ����_In     In �ҺŰ��żƻ�.����%Type,
    strSQL = strSQL & "'" & txt�ű�.Text & "',"
    '  ��Чʱ��_In In �ҺŰ��żƻ�.��Чʱ��%Type,
    If opt��Чʱ��(0).Value = True Then
        strSQL = strSQL & "Sysdate ,"
    Else
        strSQL = strSQL & "to_date('" & dtpBegin.Value & "','yyyy-mm-dd hh24:mi:ss'),"
    End If
    '  ʧЧʱ��_In In �ҺŰ��żƻ�.ʧЧʱ��%Type
    strSQL = strSQL & "to_date('" & dtpEndDate.Value & "','yyyy-mm-dd hh24:mi:ss') "
    '  ����_In     In �ҺŰ��żƻ�.����%Type,
    '  ��һ_In     In �ҺŰ��żƻ�.��һ%Type,
    '  �ܶ�_In     In �ҺŰ��żƻ�.�ܶ�%Type,
    '  ����_In     In �ҺŰ��żƻ�.����%Type,
    '  ����_In     In �ҺŰ��żƻ�.����%Type,
    '  ����_In     In �ҺŰ��żƻ�.����%Type,
    '  ����_In     In �ҺŰ��żƻ�.����%Type,
    strSQL = strSQL & strʱ��� & ","
    '   �޺ſ���_In In Varchar2,
    strSQL = strSQL & "'" & str�޺� & "',"
    '  ���﷽ʽ_In In �ҺŰ��żƻ�.���﷽ʽ%Type,
    strSQL = strSQL & "" & int���� & ","
    '  ��ſ���_In In �ҺŰ��żƻ�.��ſ���%Type,
    strSQL = strSQL & "" & IIf(chk��ſ���.Value = 1, 1, 0) & ","
    '  ��ĿID_In   In �ҺŰ��żƻ�.��ĿID%Type,
    strSQL = strSQL & Me.cboItem.ItemData(cboItem.ListIndex) & ","
    'ҽ������_In In �ҺŰ��żƻ�.ҽ������%Type,
    strSQL = strSQL & IIf(strҽ������ = "", "NULL,", "'" & strҽ������ & "',")
    'ҽ��id_In   In �ҺŰ��żƻ�.ҽ��id%Type,
    strSQL = strSQL & strҽ��ID & ","
    '  ����_In     Varchar2,
    strSQL = strSQL & "'" & str���� & "',"
    '  ����_In Number:=1,��������
    strSQL = strSQL & "" & IIf(mEditType = ed_�ƻ�����, 1, 0) & "," & bytType & ")"
     

    Set colPro = New Collection
    zlAddArray colPro, strSQL
    If Not mTimeSet.blnIsInit Then
         Call LoadTimePlan
    End If
    If SaveTimeSetData(lng�ƻ�ID, colPro) = False Then Exit Function
    '�������������Ч
    If chk�������.Value = 1 Then
        strSQL = "Zl_�ҺŰ��żƻ�_Verify(" & lng�ƻ�ID & "," & IIf(opt��Чʱ��(0).Value, 1, 0) & "," & ZVal(lng�ϴμƻ�ID) & ")"
        zlAddArray colPro, strSQL
    End If
    gcnOracle.BeginTrans: blnTrans = True
    zlExecuteProcedureArrAy colPro, Me.Caption, True, True
    gcnOracle.CommitTrans: blnTrans = False
    SavePlan = True
    Exit Function
Errhand:
    If blnTrans Then gcnOracle.RollbackTrans
    If ErrCenter = 1 Then Resume
    SaveErrLog
End Function

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
    If VsTimeValidate(-1) = False Then Exit Function
    
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
             
'             If str���ʱ�� <> "" Then
'                Do While Not mTimeSet.rsAssign.EOF
'                   If str���ʱ�� > Nvl(mTimeSet.rsAssign!��ʼʱ��) Then mTimeSet.rsAssign.Delete adAffectCurrent
'                   mTimeSet.rsAssign.MoveNext
'                Loop
'             End If

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
    '68499,������,2014-1-8,����û��ʱ����Ϣ�ƻ���ʱ����Ϣʱʱ����Ϣû�б���Ĵ���
    For i = 0 To 6
        strTemp = Switch(i = 0, "����", i = 1, "��һ", i = 2, "�ܶ�", i = 3, "����", i = 4, "����", i = 5, "����", i = 6, "����")
        mTimeSet.blnChange = True
        Call MoveAssign(strTemp)
    Next i
    
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
                           blnʱ�� = True: Exit For
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
'                                   Else
'                                        If MsgBox("�ڷ�ʱ��ҳ���еġ�" & str������Ŀ & "��������ʱ��εĺ���(" & lng���� & ")���޺���(" & lng��Լ�� & ") ����,��ȷ������ǰ���õ�ʱ�α���?", vbQuestion + vbYesNo + vbDefaultButton2, gstrSysName) = vbNo Then Exit Function
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
                   MsgBox "�ڷ�ʱ��ҳ���еġ�" & str������Ŀ & "�������õ�ԤԼ��(" & lngԤԼ�� & ")������" & IIf(lng�޺��� = lng��Լ��, "�޺���(" & lng��Լ�� & ")", "��Լ��(" & lng��Լ�� & ")") & ",�㲻�ܰ���ǰ���ñ���!", vbInformation, Me.Caption
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

Private Function CheckUsedCount() As Boolean
    '������ԤԼ��¼�İ��� �޺š���Լ��ҽ���ϰ�ʱ��
    Dim var������Ŀ As Variant, str������Ŀ As String
    Dim lng��Լ�� As Long, lng������ As Long
    Dim varԭ�Ű� As Variant, var���� As Variant
    Dim i As Long, k As Long
    Dim lng��Լ�� As Long, lng�޺��� As String, str�ϰ�ʱ�� As String
    
    On Error GoTo ErrHandler
    Call LoadRegHistory
    If mrsRegHistory.RecordCount = 0 Then CheckUsedCount = True: Exit Function
    
    var������Ŀ = Array("", "����", "��һ", "�ܶ�", "����", "����", "����", "����")
    lng�޺��� = Val(txt�޺�.Text)
    lng��Լ�� = Val(txt��Լ.Text)
    str�ϰ�ʱ�� = cbo��.Text

    For i = 1 To 7
        If opt��.Value = False Then
            lng�޺��� = Val(vsPlan.TextMatrix(2, i))
            lng��Լ�� = Val(vsPlan.TextMatrix(3, i))
            str�ϰ�ʱ�� = vsPlan.TextMatrix(1, i)
        End If
        lng��Լ�� = 0: lng������ = 0
    
        str������Ŀ = var������Ŀ(i)
        mrsRegHistory.Filter = "������Ŀ='" & str������Ŀ & "'"
        If mrsRegHistory.RecordCount <> 0 Then
            lng��Լ�� = Val(Nvl(mrsRegHistory!ͳ��))
            If lng��Լ�� > lng�޺��� Then
               Call MsgBox(IIf(opt��.Value, "", str������Ŀ) & "�޺���С����" & IIf(opt��.Value, str������Ŀ, "") & "�Ѿ�ԤԼ��ȥ������[" & lng��Լ�� & "]�����ܼ�����", vbInformation, gstrSysName)
               Exit Function
            End If
            If lng��Լ�� > lng��Լ�� Then
               Call MsgBox(IIf(opt��.Value, "", str������Ŀ) & "��Լ��С����" & IIf(opt��.Value, str������Ŀ, "") & "�Ѿ�ԤԼ��ȥ������[" & lng��Լ�� & "]�����ܼ�����", vbInformation, gstrSysName)
               Exit Function
            End If
            lng������ = Val(Nvl(mrsRegHistory!������))
            If lng������ > lng�޺��� Then
               Call MsgBox(IIf(opt��.Value, "", str������Ŀ) & "�޺���С����" & IIf(opt��.Value, str������Ŀ, "") & "�Ѿ�ԤԼ��ȥ��������[" & lng������ & "]�����ܼ�����", vbInformation, gstrSysName)
               Exit Function
            End If
            
            If lng��Լ�� > 0 Then
                If InStr(mstrԭ�Ű�, ",") = 0 Then '"��һ,����;�ܶ�,����;..."
                    'ԭ���ǡ����족
                    If str�ϰ�ʱ�� <> mstrԭ�Ű� Then
                        Call MsgBox("��ǰ�ƻ���Чʱ����" & str������Ŀ & "�����Ѿ�ԤԼ��ȥ�ĹҺż�¼�������޸��Ű࣡", vbInformation, gstrSysName)
                        Exit Function
                    End If
                Else
                    'ԭ���ǡ����ܡ�
                    varԭ�Ű� = Split(mstrԭ�Ű�, ";")
                    For k = 0 To UBound(varԭ�Ű�)
                        var���� = Split(varԭ�Ű�(k), ",")
                        If str������Ŀ = var����(0) Then
                            If var����(1) <> "" And str�ϰ�ʱ�� <> var����(1) Then
                                Call MsgBox("��ǰ�ƻ���Чʱ����" & str������Ŀ & "�����Ѿ�ԤԼ��ȥ�ĹҺż�¼�������޸��Ű࣡", vbInformation, gstrSysName)
                                Exit Function
                            End If
                            Exit For
                        End If
                    Next
                End If
            End If
        End If
    Next
    CheckUsedCount = True
    Exit Function
ErrHandler:
    If ErrCenter = 1 Then
        Resume
    End If
    SaveErrLog
End Function

Private Function SaveVerify(ByVal lng�ϴμƻ�ID As Long) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:��˹ҺŰ��żƻ�
    '����:��˳ɹ�,����true, ���򷵻�False
    '����:���˺�
    '����:2009-09-14 17:11:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    
    Err = 0: On Error GoTo Errhand
    If CheckUsedCount() = False Then Exit Function
    
    'Zl_�ҺŰ��żƻ�_Verify(Id_In In �ҺŰ��żƻ�.ID%Type,������Ч_in Number:=0)
    strSQL = "Zl_�ҺŰ��żƻ�_Verify(" & Val(mstr�ƻ�ID) & "," & chk������Ч.Value & "," & ZVal(lng�ϴμƻ�ID) & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    SaveVerify = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
    SaveErrLog
End Function
Private Function SaveCancel() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ȡ����˹ҺŰ��żƻ�
    '����:ȡ����˳ɹ�,����true, ���򷵻�False
    '����:���˺�
    '����:2009-09-14 17:11:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String, rsPlan As ADODB.Recordset
    Err = 0: On Error GoTo Errhand:
    
    strSQL = "Select 1 From �ҺŰ��żƻ� Where �ϴμƻ�ID = [1]"
    Set rsPlan = zlDatabase.OpenSQLRecord(strSQL, "�������ƻ�", Val(mstr�ƻ�ID))
    If rsPlan.RecordCount > 0 Then
        MsgBox "��ǰ�ƻ��б���ƻ�������ȡ����ˡ�", vbInformation, gstrSysName
        Exit Function
    End If
    
    'Zl_�ҺŰ��żƻ�_Cancel(Id_In In �ҺŰ��żƻ�.ID%Type) Is
    strSQL = "Zl_�ҺŰ��żƻ�_Cancel(" & Val(mstr�ƻ�ID) & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    SaveCancel = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
     SaveErrLog
End Function
Private Function SaveDelete() As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:ȡ����˹ҺŰ��żƻ�
    '����:ȡ����˳ɹ�,����true, ���򷵻�False
    '����:���˺�
    '����:2009-09-14 17:11:24
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim strSQL As String
    Err = 0: On Error GoTo Errhand:
    'Zl_�ҺŰ��żƻ�_Delete(Id_In In �ҺŰ��żƻ�.ID%Type) Is
    strSQL = "Zl_�ҺŰ��żƻ�_Delete(" & Val(mstr�ƻ�ID) & ")"
    zlDatabase.ExecuteProcedure strSQL, Me.Caption
    SaveDelete = True
    Exit Function
Errhand:
    If ErrCenter = 1 Then Resume
     SaveErrLog
End Function

Private Sub cmdOK_Click()
    Dim lng�ϴμƻ�ID As Long
    If mblnSaveMinorChange Then Call SaveMinorChange: Exit Sub
    If mEditType = ed_���Ų��� Then Unload Me: Exit Sub
    If mEditType = Ed_����ɾ�� Then
        If SaveDelete = False Then Exit Sub
        mblnSucces = True
        Unload Me: Exit Sub
    End If
    
    If mEditType = Ed_������� Then
        If LongPlanIsValied(lng�ϴμƻ�ID) = False Then Exit Sub
        If SaveVerify(lng�ϴμƻ�ID) = False Then Exit Sub
        mblnSucces = True
        Unload Me: Exit Sub
    End If
    
    If mEditType = Ed_����ȡ�� Then
        If SaveCancel = False Then Exit Sub
        mblnSucces = True
        Unload Me: Exit Sub
    End If
    
    If IsValied(True) = False Then Exit Sub
    If LongPlanIsValied(lng�ϴμƻ�ID) = False Then Exit Sub
    If SavePlan(lng�ϴμƻ�ID) = False Then Exit Sub
    mblnSucces = True
    Unload Me
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

Private Sub zlSaveTimePageSelected(ByVal str���� As String)
    If tbPage.Selected Is Nothing Then Exit Sub
    If tbPage.Selected.index <> mPageIndex.EM_ʱ�� Then
         tbPage.Item(mPageIndex.EM_ʱ��).Selected = True
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call ClearCustomData
End Sub

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
        MsgBox "��ǰ�ű���" & str������Ŀ & ",û�жԹҺ�����������,�޷�����ʱ��,����!", vbInformation, Me.Caption
        Exit Sub
    End If


    strʱ�� = mTimeSet.rsRegPlan!�Ű�
    mrs�ϰ�ʱ���.Filter = "ʱ���='" & strʱ�� & "'"
    If mrs�ϰ�ʱ���.RecordCount = 0 Then
        MsgBox "������ʱ��Ϊ[" & strʱ�� & "]�����°�ʱ��,����!", vbInformation, Me.Caption
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
                '�����:119348,����,2018/1/9,�������ţ�ʱ������ʹ�������������㣬����ĳ��ʱ��̶ȵ�ʱ����Ϊ0�����������ʱ����Ϣ����
                datʱ�� = DateAdd("h", 1, datʱ��)
                i = i - 1
            Else
                If i > lng�̶����� Then
                    With mTimeSet.rsAssign
                        .AddNew
                        !������Ŀ = str������Ŀ
                        !��ʼʱ�� = Format(datʱ��, "hh:mm:00")
                        !ʱ�� = Format(datʱ��, "hh:00:00")
                        If Format(DateAdd("n", lngʱ����, datʱ��), "yyyy-MM-dd hh:mm:00") >= Format(dat����ʱ��, "yyyy-MM-dd hh:mm:ss") Then
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

Private Sub vsTime_ValidateEdit(index As Integer, ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    Dim strʱ��() As String
     If mTimeSet.bln��ſ��� Then
        strʱ�� = Split(vsTime(index).EditText, "-")
        If UBound(strʱ��) <> 1 Then
           MsgBox "�����ʱ���ʽ����!����!", vbInformation, gstrSysName
           Cancel = True: Exit Sub
        End If
        If Not IsDate(strʱ��(0)) Then
           MsgBox "�����ʱ���ʽ����!����!", vbInformation, gstrSysName
           Cancel = True: Exit Sub
        End If
        If Not IsDate(strʱ��(1)) Then
           MsgBox "�����ʱ���ʽ����!����!", vbInformation, gstrSysName
           Cancel = True: Exit Sub
        End If
        If CDate(strʱ��(0)) >= CDate(strʱ��(1)) Then
           MsgBox "��ʼʱ�����С�ڽ���ʱ��!����!", vbInformation, gstrSysName
           Cancel = True
        End If
     End If
    mTimeSet.blnChange = True
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

Private Sub opt����_Click(index As Integer)
    Dim intRow As Integer
    Dim intCol As Integer
    '71253 ���ϴ� 2014-04-15 14:23:10 ��listView �滻ΪvsflexGrid
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
End Sub

Private Sub vsPlan_ComboCloseUp(ByVal Row As Long, ByVal Col As Long, FinishEdit As Boolean)
    If Row = 1 Then FinishEdit = True
End Sub

Private Sub opt��Чʱ��_Click(index As Integer)
     dtpBegin.Enabled = opt��Чʱ��(0).Value = False
     
     If opt��Чʱ��(0).Value = True Then
        chk�������.Value = 1
     End If
End Sub
Private Sub opt��_Click()
    Dim i As Integer
    Dim strPlan As String
    Dim ctl As Control
    
    With vsPlan
        For i = 1 To .Cols - 1
            If Trim(.TextMatrix(1, i)) <> "" Then
                If strPlan = "" Then
                    strPlan = .TextMatrix(1, i)
                Else
                    If .TextMatrix(1, i) <> strPlan Then
                        strPlan = "": Exit For
                    End If
                End If
            End If
        Next
        For i = 1 To .Cols - 1
            .TextMatrix(1, i) = ""
            .TextMatrix(2, i) = ""
            .TextMatrix(3, i) = ""
        Next
        .Enabled = False: .TabStop = False
    End With
    opt��.Value = -True: txt�޺�.Enabled = True: txt��Լ.Enabled = (chkAppoint.Value = 1)
    cbo��.Enabled = True
    opt��.Value = False
    cbo��.ListIndex = cbo.FindIndex(cbo��, strPlan, True)
    cbo��.SetFocus

    '���ñ༭����ɫ
    For Each ctl In Me.Controls
        Select Case UCase(TypeName(ctl))
        Case "TEXTBOX", UCase("ComboBox")
            Call zlSetCtrolBackColor(ctl)
        Case UCase("ListView")
        Case UCase("DTPicker")
        Case Else
        End Select
    Next
End Sub

Private Sub opt��_Click()
    Dim i As Integer
    Dim ctl As Control
    
    If Trim(cbo��.Text) <> "" Then
        With vsPlan
            For i = 1 To .Cols - 1
                .TextMatrix(1, i) = cbo��.Text
                .TextMatrix(2, i) = txt�޺�.Text
                .TextMatrix(3, i) = txt��Լ.Text
            Next
            .Enabled = True: .TabStop = True
            .Col = 1: .SetFocus
        End With
    End If
    opt��.Value = False: txt�޺�.Enabled = False: txt��Լ.Enabled = False
    cbo��.Enabled = False: cbo��.ListIndex = -1
    opt��.Value = True: vsPlan.Enabled = True

    '���ñ༭����ɫ
    For Each ctl In Me.Controls
        Select Case UCase(TypeName(ctl))
        Case "TEXTBOX", UCase("ComboBox")
            Call zlSetCtrolBackColor(ctl)
        Case UCase("ListView")
        Case UCase("DTPicker")
        Case Else
        End Select
    Next
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

Private Sub tbPage_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If mblnChangeByCode Then Exit Sub
    PageChange Item
End Sub

Private Sub PageChange(ByVal Item As XtremeSuiteControls.ITabControlItem)
    If mblnChangeByCode Then Exit Sub
    If Item.index = mPageIndex.EM_ʱ�� Then
       mblnChangeByCode = True
       tbPage.Item(mPageIndex.EM_�ƻ�).Selected = True
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



Private Sub LoadTimePlan()
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
                !ID = Val(mstr�ƻ�ID)
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
                        !ID = Val(mstr�ƻ�ID)
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
'Public Enum mRegEditType
'Ed_�ƻ����� = 0
'Ed_�����޸� = 1
'Ed_����ɾ�� = 2
'Ed_������� = 3
'Ed_����ȡ�� = 4
'Ed_���Ų��� = 5
'End Enum
     
     '���Ӽƻ�,��ʱ�����Ѿ�ԤԼ����mfrmTime.zlShowPagePlan str����, mrsRegNewData, mrsRegHistory, chk��ſ���.Value = 1, Switch(mEditType = ed_�ƻ�����, EM_�ƻ�_����, mEditType = Ed_�����޸�, EM_�ƻ�_�޸�, True, EM_�ƻ�_����), mlng����ID, Val(mstr�ƻ�ID)
     zlShowPagePlan str����, mrsRegNewData, Nothing, chk��ſ���.Value = 1, Switch(mEditType = ed_�ƻ�����, EM_�ƻ�_����, mEditType = Ed_�����޸�, EM_�ƻ�_�޸�, True, EM_�ƻ�_����), mlng����ID, Val(mstr�ƻ�ID), , strӦ��ʱ��
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
        Call tbSubPage_SelectedChanged(tbSubPage.Item(i))
    Next
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
 
Private Sub ClearCustomData()
     mTimeSet.str���� = ""
     mTimeSet.bln��ſ��� = False
     mTimeSet.lngSelIndex = 0
     mTimeSet.blnOnChange = False
     mTimeSet.lng����ID = 0
     mTimeSet.lng�ƻ�ID = 0
     mTimeSet.blnIsInit = False
     Set mTimeSet.rsRegPlan = Nothing
     Set mTimeSet.rsAssign = Nothing
     mTimeSet.blnChange = False
     Set mTimeSet.rsHistory = Nothing
End Sub
 
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
 
Private Sub txt�ű�_GotFocus()
    zlControl.TxtSelAll txt�ű�
End Sub
Private Sub txt�ű�_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
End Sub

Private Sub txt�޺�_GotFocus()
    zlControl.TxtSelAll txt�޺�
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
        MsgBox "��Լ������С���޺���!", vbInformation, gstrSysName
        Cancel = True: Exit Sub
    End If
End Sub
Private Sub txt��Լ_GotFocus()
    zlControl.TxtSelAll txt��Լ
End Sub

Private Sub txt��Լ_KeyPress(KeyAscii As Integer)
    If InStr("0123456789" & Chr(8), Chr(KeyAscii)) = 0 Then KeyAscii = 0
    If Val(txt�޺�.Text) = 0 Then KeyAscii = 0
End Sub

Private Sub txt��Լ_Validate(Cancel As Boolean)
    If Val(txt�޺�.Text) < Val(txt��Լ.Text) And _
        Trim(txt�޺�.Text) <> "" And Trim(txt��Լ.Text) <> "" Then
        MsgBox "��Լ������С���޺���!", vbInformation, gstrSysName
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
    Dim i As Long, bytCheckType As Byte '0-���ƻ��Ƿ�Ϸ�;1-��鰲��������ִ����Ŀ�Ƿ�Ϸ�.
    Dim strTittle As String
    
    On Error GoTo errHandle
    lng����ID = cbo����.ItemData(cbo����.ListIndex)
    lng��Ŀid = cboItem.ItemData(cboItem.ListIndex)
    lngҽ��ID = 0: strҽ�� = Trim(cboDoctor.Text)
    If cboDoctor.ListIndex <> -1 Then lngҽ��ID = cboDoctor.ItemData(cboDoctor.ListIndex)
    
    '���ƻ����Ƿ�����ظ�
    bytCheckType = 0
goReCheck:
    If bytCheckType <> 0 Then

        strSQL = "" & _
        "   Select Distinct A.����, A.���� D0, A.��һ D1, A.�ܶ� D2, A.���� D3, A.���� D4, A.���� D5, A.���� D6, " & _
        "                 Nvl(To_Char(a.��ʼʱ��, 'YYYY-MM-DD HH24:MI:SS'), '1901-01-01') ��Чʱ��, " & _
        "                 Nvl(To_Char(a.��ֹʱ��, 'YYYY-MM-DD HH24:MI:SS'), '3000-01-01 00:00:00') ʧЧʱ�� " & _
        "   From �ҺŰ��� A,�ҺŰ��� B " & _
        "   Where A.����id = b.����id And A.ҽ������ = b.ҽ������ And Nvl(A.ҽ��id, 0) = nvl(b.ҽ��id,0) " & _
        "               And a.ID + 0 <> [1]   And B.ID = [1]  " & _
        "   Order By ����"
            strTittle = "����"
    Else
        strSQL = "" & _
            "   Select  distinct A.����,A.���� D0,A.��һ D1,A.�ܶ� D2,A.���� D3,A.���� D4,A.���� D5,A.���� D6," & _
            "           To_Char(A.��Чʱ��,'YYYY-MM-DD HH24:MI:SS') ��Чʱ��,To_Char(A.ʧЧʱ��,'YYYY-MM-DD HH24:MI:SS') ʧЧʱ��" & _
            "   From �ҺŰ��żƻ� A, �ҺŰ��� B,�ҺŰ��� C " & _
            "   Where A.����ID=B.ID and B.����ID=C.����ID and B.ҽ������=C.ҽ������ and nvl(B.ҽ��ID,0)=nvl(C.ҽ��ID,0) " & _
            "           And B.ID+0<>[1] and C.ID=[1]  " & _
            "   Order by ����"
            strTittle = "�ƻ�����"
    End If
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID)
    blnMulitNumPlan = Not rsTemp.EOF
    If blnMulitNumPlan = False And bytCheckType = 0 Then
        bytCheckType = bytCheckType + 1
        GoTo goReCheck:
    End If
    If blnMulitNumPlan = False Then zlCheckRegistPlanIsValied = True: Exit Function
    str�ű� = ""
    Do While Not rsTemp.EOF
        str�ű� = str�ű� & "," & Nvl(rsTemp!����)
        If (Nvl(rsTemp!��Чʱ��) >= Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS") And Nvl(rsTemp!��Чʱ��) < Format(dtpEndDate.Value, "yyyy-mm-dd HH:MM:SS")) Or _
           (Nvl(rsTemp!ʧЧʱ��) >= Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS") And Nvl(rsTemp!ʧЧʱ��) < Format(dtpEndDate.Value, "yyyy-mm-dd HH:MM:SS")) Or _
           (Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS") >= Nvl(rsTemp!��Чʱ��) And Format(dtpBegin.Value, "yyyy-mm-dd HH:MM:SS") < Nvl(rsTemp!ʧЧʱ��)) Or _
           (Format(dtpEndDate.Value, "yyyy-mm-dd HH:MM:SS") >= Nvl(rsTemp!��Чʱ��) And Format(dtpEndDate.Value, "yyyy-mm-dd HH:MM:SS") < Nvl(rsTemp!ʧЧʱ��)) Then
           'ʱ���ڲ��ܽ���
            If opt��.Value Then
                If Trim(Nvl(rsTemp!D0)) <> "" Then strTemp = strTemp & vbCrLf & "  ����:" & Nvl(rsTemp!D0)
                If Trim(Nvl(rsTemp!D1)) <> "" Then strTemp = strTemp & vbCrLf & "  ��һ:" & Nvl(rsTemp!D1)
                If Trim(Nvl(rsTemp!D2)) <> "" Then strTemp = strTemp & vbCrLf & "  �ܶ�:" & Nvl(rsTemp!D2)
                If Trim(Nvl(rsTemp!D3)) <> "" Then strTemp = strTemp & vbCrLf & "  ����:" & Nvl(rsTemp!D3)
                If Trim(Nvl(rsTemp!D4)) <> "" Then strTemp = strTemp & vbCrLf & "  ����:" & Nvl(rsTemp!D4)
                If Trim(Nvl(rsTemp!D5)) <> "" Then strTemp = strTemp & vbCrLf & "  ����:" & Nvl(rsTemp!D5)
                If Trim(Nvl(rsTemp!D6)) <> "" Then strTemp = strTemp & vbCrLf & "  ����:" & Nvl(rsTemp!D6)
                If strTemp <> "" Then
                    strTemp = vbCrLf & "�ںű� [" & rsTemp!���� & "] ����������" & strTittle & ":" & vbCrLf & "        " & Mid(strTemp, 2) & vbCrLf & vbCrLf & "  ��Чʱ��:" & IIf(Nvl(rsTemp!��Чʱ��) = "1901-01-01", "����", Nvl(rsTemp!��Чʱ��) & "-" & Nvl(rsTemp!ʧЧʱ��)) & vbCrLf
                    Call MsgBox("���֡�" & cboDoctor.Text & "��ҽ�������뵱ǰ�ű��ظ��򽻲�ĹҺżƻ����� " & vbCrLf & strTemp & vbCrLf & vbCrLf & "���޸Ĵ˼ƻ�����.", vbInformation + vbOKOnly + vbDefaultButton2, gstrSysName)
                    zlCheckRegistPlanIsValied = False: Exit Function
                End If
            Else
                With vsPlan
                    For i = 0 To 6
                        strTemp1 = "  ��" & Switch(i = 0, "��", i = 1, "һ", i = 2, "��", i = 3, "��", i = 4, "��", i = 5, "��", True, "��")
                        If Trim(Nvl(rsTemp.Fields("D" & i).Value)) <> "" And Trim(.TextMatrix(1, i)) <> "" Then
                            '����,�϶��ظ���
                            strTemp = strTemp & vbCrLf & strTemp1 & ":" & Trim(Nvl(rsTemp.Fields("D" & i).Value))
                        End If
                    Next
                End With
                If strTemp <> "" Then
                    strTemp = vbCrLf & "�ںű� [" & rsTemp!���� & "] ����������" & strTittle & ":" & vbCrLf & "        " & Mid(strTemp, 2) & vbCrLf & "  ��Чʱ��:" & IIf(Nvl(rsTemp!��Чʱ��) = "1901-01-01", "����", Nvl(rsTemp!��Чʱ��) & "-" & Nvl(rsTemp!ʧЧʱ��)) & vbCrLf
                    Call MsgBox("���֡�" & cboDoctor.Text & "��ҽ�������뵱ǰ�ű��ظ��򽻲�ĹҺŰ��� " & vbCrLf & strTemp & vbCrLf & vbCrLf & "���޸Ĵ˼ƻ�����.", vbInformation + vbOKOnly + vbDefaultButton2, gstrSysName)
                    zlCheckRegistPlanIsValied = False: Exit Function
                End If
            End If
        End If
        rsTemp.MoveNext
    Loop
    If bytCheckType = 0 Then
        bytCheckType = bytCheckType + 1
        GoTo goReCheck:
    End If
    zlCheckRegistPlanIsValied = True
    Exit Function
errHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
     SaveErrLog
End Function

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
        MsgBox "��ǰ�ű���" & str������Ŀ & ",û�жԹҺ�����������,�޷�����ʱ��,����!", vbInformation, Me.Caption
        Exit Function
    End If


    strʱ�� = mTimeSet.rsRegPlan!�Ű�
    mrs�ϰ�ʱ���.Filter = "ʱ���='" & strʱ�� & "'"
    If mrs�ϰ�ʱ���.RecordCount = 0 Then
        MsgBox "������ʱ��Ϊ[" & strʱ�� & "]�����°�ʱ��,����!", vbInformation, Me.Caption
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

Private Sub txtTimeOut_Change()
    If Val(txtTimeOut.Text) > 1440 Then txtTimeOut.Text = 1440
End Sub

Private Sub cmd����ʱ��_Click()
    If AssignReapportion(Val(txtTimeOut.Text), tbSubPage.Item(mTimeSet.lngSelIndex).Caption) = False Then Exit Sub
    Call tbSubPage_SelectedChanged(tbSubPage.Item(mTimeSet.lngSelIndex))
End Sub

Private Sub vsPlan_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsPlan
        If mEditType <> ed_�ƻ����� And mEditType <> Ed_�����޸� Then Cancel = True: Exit Sub
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
   Call zl_VsGridRowChange(vsPlan, OldRow, NewRow, OldCol, NewCol)
    vsPlan.ColComboList(NewCol) = ""
    If OldRow = 1 And Trim(vsPlan.TextMatrix(1, OldCol)) = "" Then
        vsPlan.TextMatrix(2, OldCol) = ""
        vsPlan.TextMatrix(3, OldCol) = ""
    End If
    If OldRow = 2 And Trim(vsPlan.TextMatrix(3, OldCol)) = "" And mbln�Զ�Ĭ����Լ�� Then
        vsPlan.TextMatrix(3, OldCol) = vsPlan.TextMatrix(2, OldCol)
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
    Dim strKey As String, intCol As Integer, strTemp As String
    Dim str������Ŀ As String
    Dim lng��Լ��  As Long
    '������֤
    With vsPlan
        str������Ŀ = Switch(Col = 1, "����", Col = 2, "��һ", Col = 3, "�ܶ�", Col = 4, "����", Col = 5, "����", Col = 6, "����", True, "����")
        strKey = Trim(.EditText): strKey = Replace(strKey, Chr(vbKeyReturn), ""): strKey = Replace(strKey, Chr(10), "")
        If .Row <= 1 Then Exit Sub
        If zlCommFun.DblIsValid(strKey, 5, True, False, 0, .ColKey(Col)) = False Then
            Cancel = True: Exit Sub
        End If
        If Val(strKey) <> 0 Then
            strKey = Format(Abs(Val(strKey)), "####;;;")
        End If
        If Row = 2 Then
            If Val(strKey) < Val(.TextMatrix(3, Col)) Then
                If MsgBox("�޺���С������Լ��,�Ƿ������Լ��?", vbYesNo + vbDefaultButton2 + vbQuestion, gstrSysName) = vbNo Then Cancel = True: Exit Sub
                .TextMatrix(3, Col) = ""
            End If
        ElseIf Row = 3 Then
            If Val(strKey) > Val(.TextMatrix(2, Col)) Then
                Call MsgBox("�޺���С������Լ��,���ܼ���", vbInformation, gstrSysName)
                Cancel = True: Exit Sub
            End If
        End If
        .EditText = strKey
    End With
End Sub




Private Sub cboDoctor_Validate(Cancel As Boolean)
       
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

Private Sub LoadDoctor()
    Set mrsDoctor = GetDoctor(Val(cbo����.ItemData(cbo����.ListIndex)), "")
    cboDoctor.Clear
    Do While Not mrsDoctor.EOF
        cboDoctor.AddItem mrsDoctor!����
        cboDoctor.ItemData(cboDoctor.NewIndex) = mrsDoctor!ID
        mrsDoctor.MoveNext
    Loop
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
        KeyAscii = 0: Exit Sub
    End If
    Exit Sub
End Sub

Private Function Checkʱ��() As Boolean
    '�����Ӽƻ�ʱ ��ȡԭ�еİ����Ƿ����ʱ��
    '�޸ļƻ�ʱ ��ȡԭ�ƻ��Ƿ����ʱ��
   Dim strSQL           As String
   Dim rsTmp            As ADODB.Recordset
   If mEditType <> Ed_�����޸� And mEditType <> ed_�ƻ����� Then Exit Function
    On Error GoTo Hd
    If mEditType = ed_�ƻ����� Then
        strSQL = " Select 1 As Hdata From �ҺŰ���ʱ�� Where ����id =[1] And Rownum=1"
    Else
        strSQL = "Select 1  as haveData From �Һżƻ�ʱ�� Where �ƻ�ID=[2] and Rownum=1"
    End If
     Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, Val(mstr�ƻ�ID))
     Checkʱ�� = Not rsTmp.EOF
    Set rsTmp = Nothing
   
   Exit Function
Hd:
   If ErrCenter() = 1 Then
        Resume
   End If
   SaveErrLog
End Function


Private Function IsValidation() As Boolean
    '��� ��Чʱ���Ƿ�Ϸ�
     If mbln�����޸� Then
      Select Case mEditType
        
                     ' dtpBegin.MinDate = mdtMinCustom
          Case Ed_�������
             If Format(dtpBegin.Value, "yyyy-mm-dd hh:mm:ss") < Format(mdtMinCustom, "yyyy-mm-dd hh:mm:ss") Then
                If MsgBox("�üƻ�����Ч���ں��Ѵ���ԤԼ��,�Ƿ����?", vbYesNo + vbDefaultButton1 + vbInformation, Me.Caption) = vbNo Then
                    Exit Function
                End If
            End If
          Case Ed_�����޸�
                'dtpBegin.MinDate = mdtMinCustom
      End Select
    End If
    IsValidation = True
 End Function
 Private Function CheckExistsBooking(str�ű� As String, Optional dtCustom As Date) As Boolean
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:���ָ���ű��Ƿ����ԤԼ�Һŵ�
    '���:str�ű�-�ű�
    '����:����,����true,���򷵻�False
    '����:
    '����:2009-09-15 10:32:02
    '---------------------------------------------------------------------------------------------------------------------------------------------

    Dim rsTmp As ADODB.Recordset, strSQL As String
 
    strSQL = "Select /*+ Rule*/ Max(����ʱ��) ʱ��" & vbNewLine & _
            "From ������ü�¼" & vbNewLine & _
            "Where ��¼���� = 4 And ��¼״̬ In (0, 1) And ���㵥λ = [1] And ����ʱ�� > �Ǽ�ʱ��"
    If gintԤԼ���� = 0 Then
        strSQL = strSQL & " And ����ʱ�� > Sysdate"
    Else
        strSQL = strSQL & " And ����ʱ�� Between Sysdate And Sysdate+" & gintԤԼ����
    End If
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, str�ű�)
    
    CheckExistsBooking = Not IsNull(rsTmp!ʱ��)
    dtCustom = IIf(CheckExistsBooking, rsTmp!ʱ��, zlDatabase.Currentdate)
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Function LoadRegHistory() As Boolean
    Dim strSQL As String
    strSQL = "Select ������Ŀ, Max(������) As ������, Max(ͳ��) As ͳ��, Max(����ʱ��) As ����ʱ��" & vbNewLine & _
            " From (Select Decode(To_Char(a.����ʱ��, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����') As ������Ŀ," & vbNewLine & _
            "              Max(Nvl(a.����, 0)) As ������, Count(1) As ͳ��, To_Char(Max(����ʱ��), 'hh24:mi:ss') As ����ʱ��," & vbNewLine & _
            "              To_Char(����ʱ��, 'YYYY-MM-DD') As ��������" & vbNewLine & _
            "       From ���˹Һż�¼ A, �ҺŰ��� B" & vbNewLine & _
            "       Where a.��¼״̬ = 1 And a.����ʱ�� Between [2] And [3] And a.�ű� = b.���� And b.Id = [1] " & vbNewLine & _
            "       Group By Decode(To_Char(a.����ʱ��, 'D'), '1', '����', '2', '��һ', '3', '�ܶ�', '4', '����', '5', '����', '6', '����', '7', '����')," & vbNewLine & _
            "                To_Char(����ʱ��, 'YYYY-MM-DD'))" & vbNewLine & _
            " Group By ������Ŀ"

    On Error GoTo Hd:
    Set mrsRegHistory = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mlng����ID, dtpBegin, dtpEndDate)
    LoadRegHistory = True
    
    Exit Function
Hd:
    If ErrCenter() = 1 Then
        Resume
    End If
    SaveErrLog
End Function
 
Public Property Let �Զ�Ĭ����Լ��(ByVal vNewValue As Boolean)
    mbln�Զ�Ĭ����Լ�� = vNewValue
End Property

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

Private Sub ClearVsGridCheckValue()
    '---------------------------------------------------------------------------------------------------------------------------------------------
    '����:�������ؼ��ĸ�ѡ��ֵ
    '����:���ϴ�
    '����:2014-04-14 18:19:32
    '---------------------------------------------------------------------------------------------------------------------------------------------
    Dim intType As Integer
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
