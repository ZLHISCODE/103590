VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.4#0"; "ZL9BillEdit.ocx"
Begin VB.Form frmDrugstorePara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҩ�����в���"
   ClientHeight    =   5265
   ClientLeft      =   2565
   ClientTop       =   1500
   ClientWidth     =   7950
   Icon            =   "frmDrugstorePara.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5265
   ScaleWidth      =   7950
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  '��Ļ����
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   4065
      Index           =   3
      Left            =   210
      TabIndex        =   56
      Top             =   510
      Width           =   7425
      Begin VB.CommandButton cmdOperate 
         Caption         =   "����(&A)"
         CausesValidation=   0   'False
         Height          =   350
         Index           =   0
         Left            =   6240
         TabIndex        =   63
         Top             =   510
         Width           =   1100
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "�޸�(&M)"
         CausesValidation=   0   'False
         Height          =   350
         Index           =   1
         Left            =   6240
         TabIndex        =   62
         Top             =   990
         Width           =   1100
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "ɾ��(&D)"
         CausesValidation=   0   'False
         Height          =   350
         Index           =   2
         Left            =   6240
         TabIndex        =   61
         Top             =   1470
         Width           =   1100
      End
      Begin VB.CommandButton cmdOperate 
         Caption         =   "���(&L)"
         CausesValidation=   0   'False
         Height          =   350
         Index           =   3
         Left            =   6240
         TabIndex        =   57
         Top             =   1950
         Width           =   1100
      End
      Begin MSComctlLib.ImageList ils16 
         Left            =   6600
         Top             =   2880
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   1
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "frmDrugstorePara.frx":000C
               Key             =   "Limit"
            EndProperty
         EndProperty
      End
      Begin MSComctlLib.ListView lvw 
         Height          =   3465
         Index           =   1
         Left            =   300
         TabIndex        =   64
         Top             =   480
         Width           =   5865
         _ExtentX        =   10345
         _ExtentY        =   6112
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         SmallIcons      =   "ils16"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "������"
            Object.Width           =   2646
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "��������"
            Object.Width           =   2117
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   2
            Text            =   "��ʷ����"
            Object.Width           =   1411
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   3
            Text            =   "�����޸����˵���"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   4
            Text            =   "�������"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "������Ա�Բ�ͬ���ݵĲ���Ȩ�ޣ���Ե��ݵ���ʷ��������������˽�������"
         Height          =   180
         Index           =   0
         Left            =   300
         TabIndex        =   65
         Top             =   180
         Width           =   6120
      End
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   4080
      Index           =   4
      Left            =   150
      TabIndex        =   58
      Top             =   480
      Width           =   7500
      Begin ZL9BillEdit.BillEdit bill 
         Height          =   3585
         Index           =   0
         Left            =   210
         TabIndex        =   60
         Top             =   330
         Width           =   7185
         _ExtentX        =   12674
         _ExtentY        =   6324
         CellAlignment   =   9
         Text            =   ""
         TextMatrix0     =   ""
         MaxDate         =   2958465
         MinDate         =   -53688
         Value           =   36395
         Cols            =   3
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
         BackStyle       =   0  'Transparent
         Caption         =   "����ҩƷ�ڲ�ͬ�ⷿ�����ͨ����"
         Height          =   180
         Index           =   23
         Left            =   240
         TabIndex        =   59
         Top             =   60
         Width           =   2700
      End
   End
   Begin VB.Frame fraMain 
      BackColor       =   &H8000000A&
      BorderStyle     =   0  'None
      Height          =   4065
      Index           =   5
      Left            =   150
      TabIndex        =   67
      Top             =   405
      Width           =   7530
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid msf�ⷿ��λ 
         Height          =   3900
         Left            =   180
         TabIndex        =   68
         Top             =   105
         Width           =   7275
         _ExtentX        =   12832
         _ExtentY        =   6879
         _Version        =   393216
         Cols            =   5
         FixedCols       =   0
         GridColor       =   -2147483636
         GridColorFixed  =   -2147483631
         AllowBigSelection=   0   'False
         GridLinesFixed  =   1
         ScrollBars      =   2
         AllowUserResizing=   1
         FormatString    =   "ҩƷ�ⷿ|�ۼ۵�λ|���ﵥλ|סԺ��λ|ҩ�ⵥλ"
         _NumberOfBands  =   1
         _Band(0).Cols   =   5
      End
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   4065
      Index           =   1
      Left            =   210
      TabIndex        =   34
      Top             =   510
      Width           =   7425
      Begin VB.Frame fra 
         Caption         =   "������ʾ"
         Height          =   1095
         Index           =   11
         Left            =   4020
         TabIndex        =   39
         Top             =   180
         Width           =   2595
         Begin VB.CheckBox chk 
            Caption         =   "��Ա������������ʾ"
            Height          =   285
            Index           =   14
            Left            =   420
            TabIndex        =   40
            ToolTipText     =   "��ʾ����������￨���봦�Ƿ�Ϊ������ʾ"
            Top             =   450
            Width           =   1920
         End
      End
      Begin VB.Frame fra 
         Caption         =   "�վ��д�"
         Height          =   1095
         Index           =   10
         Left            =   300
         TabIndex        =   35
         Top             =   150
         Width           =   2685
         Begin VB.TextBox txtUD 
            Alignment       =   1  'Right Justify
            Height          =   300
            Index           =   2
            Left            =   1575
            Locked          =   -1  'True
            TabIndex        =   37
            Text            =   "3"
            Top             =   480
            Width           =   435
         End
         Begin MSComCtl2.UpDown ud 
            Height          =   300
            Index           =   2
            Left            =   2010
            TabIndex        =   38
            TabStop         =   0   'False
            Top             =   480
            Width           =   240
            _ExtentX        =   423
            _ExtentY        =   529
            _Version        =   393216
            Value           =   3
            BuddyControl    =   "txtUD(2)"
            BuddyDispid     =   196618
            BuddyIndex      =   2
            OrigLeft        =   1965
            OrigTop         =   390
            OrigRight       =   2205
            OrigBottom      =   690
            Max             =   100
            SyncBuddy       =   -1  'True
            BuddyProperty   =   65547
            Enabled         =   -1  'True
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "�շ��վ��д�"
            Height          =   180
            Index           =   18
            Left            =   420
            TabIndex        =   36
            Top             =   540
            Width           =   1080
         End
      End
      Begin VB.Frame fra 
         Height          =   75
         Index           =   9
         Left            =   1230
         TabIndex        =   42
         Top             =   1620
         Width           =   5415
      End
      Begin VB.CheckBox chk 
         Caption         =   "Ʊ���ϸ����"
         Height          =   285
         Index           =   13
         Left            =   5295
         TabIndex        =   47
         ToolTipText     =   "��ʾ����������￨���봦�Ƿ�Ϊ������ʾ"
         Top             =   3045
         Width           =   1380
      End
      Begin VB.TextBox txtUD 
         Alignment       =   1  'Right Justify
         Enabled         =   0   'False
         Height          =   300
         Index           =   4
         Left            =   4005
         Locked          =   -1  'True
         TabIndex        =   45
         Text            =   "7"
         Top             =   3060
         Width           =   390
      End
      Begin MSComCtl2.UpDown ud 
         Height          =   300
         Index           =   4
         Left            =   4395
         TabIndex        =   46
         TabStop         =   0   'False
         Top             =   3060
         Width           =   240
         _ExtentX        =   423
         _ExtentY        =   529
         _Version        =   393216
         Value           =   7
         BuddyControl    =   "txtUD(4)"
         BuddyDispid     =   196618
         BuddyIndex      =   4
         OrigLeft        =   3795
         OrigTop         =   3630
         OrigRight       =   4035
         OrigBottom      =   3915
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   65547
         Enabled         =   -1  'True
      End
      Begin MSComctlLib.ListView lvw 
         Height          =   1125
         Index           =   0
         Left            =   300
         TabIndex        =   43
         Top             =   1800
         Width           =   6315
         _ExtentX        =   11139
         _ExtentY        =   1984
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   3
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Ʊ������"
            Object.Width           =   4410
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "���볤��"
            Object.Width           =   2540
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   2
            SubItemIndex    =   2
            Text            =   "Ʊ���ϸ����"
            Object.Width           =   2540
         EndProperty
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "���볤��"
         Height          =   180
         Index           =   19
         Left            =   3195
         TabIndex        =   44
         Top             =   3120
         Width           =   720
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Ʊ�ݻ���"
         Height          =   180
         Index           =   9
         Left            =   300
         TabIndex        =   41
         Top             =   1560
         Width           =   720
      End
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   4065
      Index           =   0
      Left            =   210
      TabIndex        =   1
      Top             =   510
      Width           =   7425
      Begin VB.Frame fra��Ա�� 
         Caption         =   "��Ա���۸�"
         Height          =   1155
         Left            =   3135
         TabIndex        =   27
         Top             =   2640
         Width           =   4275
         Begin VB.CommandButton cmdSelect 
            Caption         =   "��"
            Height          =   255
            Left            =   2175
            TabIndex        =   33
            TabStop         =   0   'False
            Top             =   720
            Width           =   255
         End
         Begin VB.CheckBox chk��� 
            Caption         =   "�۸���Ա䶯"
            Height          =   285
            Left            =   2820
            TabIndex        =   30
            Top             =   330
            Width           =   1395
         End
         Begin VB.TextBox txt�۸� 
            Alignment       =   1  'Right Justify
            Height          =   300
            Left            =   1200
            TabIndex        =   29
            Top             =   300
            Width           =   1245
         End
         Begin VB.TextBox txt������Ŀ 
            Height          =   300
            Left            =   1200
            Locked          =   -1  'True
            TabIndex        =   32
            Top             =   690
            Width           =   1245
         End
         Begin VB.Label lbl������Ŀ 
            AutoSize        =   -1  'True
            Caption         =   "����������Ŀ"
            Height          =   180
            Left            =   90
            TabIndex        =   31
            Top             =   750
            Width           =   1080
         End
         Begin VB.Label lbl�۸� 
            AutoSize        =   -1  'True
            Caption         =   "��ǰ�۸�"
            Height          =   180
            Left            =   420
            TabIndex        =   28
            Top             =   360
            Width           =   720
         End
      End
      Begin VB.Frame fra 
         Caption         =   "����"
         Height          =   3750
         Index           =   4
         Left            =   210
         TabIndex        =   2
         Top             =   60
         Width           =   2775
         Begin VB.CheckBox chk 
            Caption         =   "ʱ��ҩƷ�Լӳ������"
            Height          =   285
            Index           =   21
            Left            =   120
            TabIndex        =   66
            Top             =   3360
            Width           =   2160
         End
         Begin VB.CheckBox chk 
            Caption         =   "�շ���ɺ��Ƿ��Զ���ҩ"
            Height          =   285
            Index           =   17
            Left            =   120
            TabIndex        =   9
            Top             =   3000
            Width           =   2370
         End
         Begin VB.ComboBox cmb 
            Height          =   300
            Index           =   3
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   4
            Top             =   630
            Width           =   2235
         End
         Begin VB.CheckBox chk 
            Caption         =   "δ��ҩ������ҩ"
            Height          =   285
            Index           =   16
            Left            =   120
            TabIndex        =   8
            Top             =   2670
            Width           =   1680
         End
         Begin VB.ComboBox cmb 
            Height          =   300
            Index           =   1
            Left            =   120
            Style           =   2  'Dropdown List
            TabIndex        =   6
            Top             =   1260
            Width           =   2235
         End
         Begin VB.CheckBox chk 
            Caption         =   "ָ��ҩ��ʱ�޶�ҩƷ�Ŀ��"
            Height          =   285
            Index           =   2
            Left            =   120
            TabIndex        =   7
            Top             =   2355
            Value           =   1  'Checked
            Width           =   2460
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "ָ�������۶��۵�λ"
            Height          =   180
            Index           =   11
            Left            =   120
            TabIndex        =   3
            Top             =   360
            Width           =   1620
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "�ֱҴ���"
            Height          =   180
            Index           =   15
            Left            =   120
            TabIndex        =   5
            Top             =   1020
            Width           =   720
         End
      End
      Begin VB.Frame fra 
         Caption         =   "ҩƷ����"
         Height          =   1215
         Index           =   8
         Left            =   4950
         TabIndex        =   14
         Top             =   75
         Width           =   2475
         Begin VB.OptionButton opt 
            Caption         =   "��飬�����ֹ"
            Height          =   195
            Index           =   6
            Left            =   375
            TabIndex        =   17
            Top             =   915
            Width           =   1560
         End
         Begin VB.OptionButton opt 
            Caption         =   "��飬��������"
            Height          =   195
            Index           =   5
            Left            =   375
            TabIndex        =   16
            Top             =   600
            Width           =   1560
         End
         Begin VB.OptionButton opt 
            Caption         =   "�����п����"
            Height          =   195
            Index           =   4
            Left            =   375
            TabIndex        =   15
            Top             =   315
            Value           =   -1  'True
            Width           =   1560
         End
      End
      Begin VB.Frame fra 
         Caption         =   "�շ�ʱ��Ա����"
         Height          =   1215
         Index           =   6
         Left            =   3135
         TabIndex        =   10
         Top             =   75
         Width           =   1650
         Begin VB.CheckBox chk 
            Caption         =   "����"
            Height          =   210
            Index           =   7
            Left            =   315
            TabIndex        =   11
            Top             =   330
            Value           =   1  'Checked
            Width           =   1020
         End
         Begin VB.CheckBox chk 
            Caption         =   "��ԱID"
            Height          =   225
            Index           =   8
            Left            =   315
            TabIndex        =   12
            Top             =   615
            Value           =   1  'Checked
            Width           =   960
         End
         Begin VB.CheckBox chk 
            Caption         =   "ˢ��Ա��"
            Height          =   210
            Index           =   9
            Left            =   315
            TabIndex        =   13
            Top             =   930
            Value           =   1  'Checked
            Width           =   1020
         End
      End
      Begin VB.Frame fra 
         Caption         =   "�������°�ʱ��"
         Height          =   1125
         Index           =   1
         Left            =   3120
         TabIndex        =   18
         Top             =   1410
         Width           =   4305
         Begin MSComCtl2.DTPicker dtp 
            Height          =   315
            Index           =   0
            Left            =   825
            TabIndex        =   20
            Top             =   270
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   43778051
            UpDown          =   -1  'True
            CurrentDate     =   36526.3541666667
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   315
            Index           =   1
            Left            =   2040
            TabIndex        =   22
            Top             =   270
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   43778051
            UpDown          =   -1  'True
            CurrentDate     =   36526.5
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   315
            Index           =   2
            Left            =   825
            TabIndex        =   24
            Top             =   675
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   43778051
            UpDown          =   -1  'True
            CurrentDate     =   36526.5625
         End
         Begin MSComCtl2.DTPicker dtp 
            Height          =   315
            Index           =   3
            Left            =   2040
            TabIndex        =   26
            Top             =   675
            Width           =   870
            _ExtentX        =   1535
            _ExtentY        =   556
            _Version        =   393216
            CustomFormat    =   "HH:mm"
            Format          =   43778051
            UpDown          =   -1  'True
            CurrentDate     =   36526.75
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Index           =   2
            Left            =   330
            TabIndex        =   19
            Top             =   330
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   4
            Left            =   1785
            TabIndex        =   21
            Top             =   345
            Width           =   180
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "����"
            Height          =   180
            Index           =   3
            Left            =   330
            TabIndex        =   23
            Top             =   735
            Width           =   360
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "��"
            Height          =   180
            Index           =   5
            Left            =   1785
            TabIndex        =   25
            Top             =   750
            Width           =   180
         End
      End
   End
   Begin VB.Frame fraMain 
      BorderStyle     =   0  'None
      Height          =   4065
      Index           =   2
      Left            =   210
      TabIndex        =   48
      Top             =   510
      Width           =   7425
      Begin VB.ListBox lst 
         Height          =   3420
         Index           =   1
         Left            =   2430
         Style           =   1  'Checkbox
         TabIndex        =   52
         Top             =   390
         Width           =   1935
      End
      Begin VB.ListBox lst 
         Height          =   3420
         Index           =   0
         Left            =   240
         Style           =   1  'Checkbox
         TabIndex        =   50
         Top             =   390
         Width           =   1935
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "���Ѳ������÷�������"
         Height          =   180
         Index           =   21
         Left            =   2430
         TabIndex        =   51
         Top             =   150
         Width           =   1800
      End
      Begin VB.Label lbl 
         AutoSize        =   -1  'True
         Caption         =   "ҽ���������÷�������"
         Height          =   180
         Index           =   20
         Left            =   270
         TabIndex        =   49
         Top             =   150
         Width           =   1800
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      CausesValidation=   0   'False
      Height          =   350
      Left            =   300
      TabIndex        =   55
      Top             =   4785
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   6510
      TabIndex        =   54
      Top             =   4785
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   5190
      TabIndex        =   53
      Top             =   4785
      Width           =   1100
   End
   Begin MSComctlLib.TabStrip tabMain 
      Height          =   4530
      Left            =   135
      TabIndex        =   0
      Top             =   75
      Width           =   7650
      _ExtentX        =   13494
      _ExtentY        =   7990
      TabWidthStyle   =   1
      MultiRow        =   -1  'True
      ShowTips        =   0   'False
      HotTracking     =   -1  'True
      TabMinWidth     =   2117
      TabStyle        =   1
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   6
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "����"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Ʊ�ݹ���"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "Ȩ��"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "���ݲ���"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab5 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ҩƷ����"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab6 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ҩƷ�ⷿ��λ"
            ImageVarType    =   2
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
End
Attribute VB_Name = "frmDrugstorePara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum const��
    ud_��ʷ���� = 0
    ud_�շ��վ� = 2
    ud_���볤�� = 4
End Enum

Private Enum constChk
    chk_�޶�ҩƷ�Ŀ�� = 2
    chk_�������� = 7
    chk_����ID = 8
    chk_ˢ���￨ = 9
    chk_Ʊ�ſ��� = 13
    chk_������ʾ = 14
    chk_δ��ҩ������ҩ = 16
    chk_�շ�ͬʱ��ҩ = 17
    chk_ʱ��ҩƷ��� = 21
End Enum

Private Enum const����
    dtp_�����ϰ� = 0
    dtp_�����°� = 1
    dtp_�����ϰ� = 2
    dtp_�����°� = 3
End Enum

Private Enum constCmb
    cmb_�ֱҴ��� = 1
    cmb_���۵�λ = 3
End Enum

Private Enum constBill
    bill_ҩƷ���� = 0
End Enum

Private Enum constLvw
    lvw_Ʊ�� = 0
    lvw_���� = 1
End Enum

Private Enum constListBox
    lst_ҽ������ = 0
    lst_���Ѳ��� = 1
End Enum

Private Enum constOpt
    opt_�����п���� = 4
    opt_�������� = 5
    opt_�����ֹ = 6
End Enum

'��������
Dim mblnChange As Boolean     '�Ƿ�ı���
Dim mblnInit As Boolean       '�Ƿ��ʼ��ʧ��
Dim mblnLoad As Boolean
Dim mintColumn As Integer '

'���ڻ�Ա���۸����ö��ر����ӵı���
Dim mlng��Ա��ID  As Long
Dim mstr��Ա������  As String
Dim mlng��ĿID As Long

Private Sub cmdHelp_Click()
     ShowHelp App.ProductName, Me.hwnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdOperate_Click(Index As Integer)
    Dim str���� As String, str��ԱID As String, str���� As String
    Dim lng���� As Long, lng���� As Long, bln�޸����� As Boolean
    Dim dbl������� As Double
    Dim lst As ListItem
    
    
    Select Case Index
        Case 0 '����
            If frmBillPrivilege.�༭Ȩ��(str����, str��ԱID, str����, lng����, lng����, bln�޸�����, dbl�������, Me) = False Then
                Exit Sub
            End If
                
            For Each lst In lvw(lvw_����).ListItems
                If lst.Tag = str��ԱID And lst.ListSubItems(1).Tag = lng���� Then
                    MsgBox "���������Ĳ��������Ѿ����ڡ�", vbInformation, gstrSysName
                    Exit Sub
                End If
            Next
        Case 1 '�޸�
            If lvw(lvw_����).SelectedItem Is Nothing Then Exit Sub
            
            With lvw(lvw_����).SelectedItem
                str���� = .Text
                str���� = .SubItems(1)
                lng���� = Val(.SubItems(2))
                bln�޸����� = (.SubItems(3) = "��")
                dbl������� = Val(.SubItems(4))
                str��ԱID = .Tag
                lng���� = .ListSubItems(1).Tag
            End With
            If frmBillPrivilege.�༭Ȩ��(str����, str��ԱID, str����, lng����, lng����, bln�޸�����, dbl�������, Me) = False Then
                Exit Sub
            End If
                
            For Each lst In lvw(lvw_����).ListItems
                If Not lst Is lvw(lvw_����).SelectedItem Then
                    If lst.Tag = str��ԱID And lst.ListSubItems(1).Tag = lng���� Then
                        MsgBox "���θı�Ĳ��������Ѿ����ڡ�", vbInformation, gstrSysName
                        Exit Sub
                    End If
                End If
            Next
            
        Case 2 'ɾ��
            If lvw(lvw_����).SelectedItem Is Nothing Then Exit Sub
            
            With lvw(lvw_����).SelectedItem
                If MsgBox("��ȷʵҪɾ����" & .Text & "���ԡ�" & .SubItems(1) & "���Ĳ������ƣ�", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                    Exit Sub
                End If
                
                lvw(lvw_����).ListItems.Remove .Index
            End With
        Case 3 '���
            If MsgBox("��ȷʵҪɾ�����еĲ������ƣ�", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                Exit Sub
            End If
            
            lvw(lvw_����).ListItems.Clear
    End Select
    
    If Index = 0 Or Index = 1 Then
        If Index = 0 Then
            Set lst = lvw(lvw_����).ListItems.Add(, , str����, , "Limit")
            lst.Selected = True
            lst.EnsureVisible
        Else
            Set lst = lvw(lvw_����).SelectedItem
            lst.Text = str����
        End If
        lst.SubItems(1) = str����
        lst.SubItems(2) = lng����
        lst.SubItems(3) = IIF(bln�޸����� = True, "��", "��")
        lst.SubItems(4) = Format(dbl�������, "0.00")
        lst.Tag = str��ԱID
        lst.ListSubItems(1).Tag = lng����
    End If
    mblnChange = True
End Sub

Private Sub cmdSelect_Click()
'ѡ��������Ŀ
    Dim strSQL As String
    Dim blnRe As Boolean
    Dim strTemp As String
    Dim strID As String
    Dim lngRow As Long
    
    strTemp = txt������Ŀ.Text
    strID = txt������Ŀ.Tag
    strSQL = "select ID,�ϼ�ID,����,ĩ��  from ������Ŀ where " & Where����ʱ��() & _
        "  start with �ϼ�ID is null  connect by prior ID =�ϼ�ID"
    blnRe = frmTreeLeafSel.ShowTree(strSQL, strID, strTemp, "������Ŀ")
    If blnRe Then
        On Error Resume Next
        txt������Ŀ.Tag = strID
        txt������Ŀ.Text = strTemp
        mblnChange = True
    End If
End Sub

Private Sub Form_Activate()
    If mblnLoad = False Then Exit Sub
    '���²���ֻ����һ��
    mblnLoad = False
    If mblnInit = False Then Unload Me
    Call tabMain_Click
End Sub

Private Sub Form_Load()
    On Error GoTo ErrHandle
    
    mblnLoad = True
    '���г�ʼ��
    Call InitEnv
    Call LoadPara
    Call Load��Ա��
    Call LoadҩƷ����
    Call LoadҩƷ�ⷿ��λ
    
    RestoreFlexState Bill(bill_ҩƷ����), App.ProductName & "\" & Me.Name & bill_ҩƷ����
    '��ʼ���ɹ�
    mblnChange = False
    mblnInit = True
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub InitEnv()
    '��ʼ�����ڣ����ǵ��ǲ���Ҫ�����ݿ��
    cmb(cmb_�ֱҴ���).AddItem "0-������"
    cmb(cmb_�ֱҴ���).AddItem "1-��������"
    cmb(cmb_�ֱҴ���).AddItem "2-������ȡ"
    cmb(cmb_�ֱҴ���).AddItem "3-�����ȡ"
    cmb(cmb_�ֱҴ���).ListIndex = 0
    
    cmb(cmb_���۵�λ).AddItem "0-�ۼ۵�λ"
    cmb(cmb_���۵�λ).AddItem "1-�ɹ���λ"
    cmb(cmb_���۵�λ).ListIndex = 0
    
    lvw(lvw_Ʊ��).ListItems.Add , "C1", "�շ��վ�"
    lvw(lvw_Ʊ��).ListItems.Add , "C5", "��Ա��"
    
    With Bill(bill_ҩƷ����)
        .Cols = 4 '����һ��������
        .ColAlignment(0) = 1
        .ColAlignment(1) = 1
        .ColAlignment(2) = 1
        .ColAlignment(3) = 1
        .TextMatrix(0, 0) = "���ڿⷿ"
        .TextMatrix(0, 1) = "�Է��ⷿ"
        .TextMatrix(0, 2) = "�Է��ⷿID"
        .TextMatrix(0, 3) = "����"
        .ColWidth(0) = 1700
        .ColWidth(1) = 1700
        .ColWidth(2) = 0
        .ColWidth(3) = 3600
        .ColData(0) = 3
        .ColData(1) = 3
        .ColData(2) = 5
        .ColData(3) = 0
        .PrimaryCol = 0
        .Active = True
    End With
    '�ⷿ��λ
    msf�ⷿ��λ.AllowUserResizing = flexResizeNone
    msf�ⷿ��λ.Cols = 3
    msf�ⷿ��λ.FormatString = "ҩƷ�ⷿ|�ۼ۵�λ|ҩ�ⵥλ"
    msf�ⷿ��λ.ColWidth(1) = 900
    msf�ⷿ��λ.ColWidth(2) = 900
    msf�ⷿ��λ.ColAlignment(1) = 4
    msf�ⷿ��λ.ColAlignment(2) = 4
    msf�ⷿ��λ.ColWidth(0) = msf�ⷿ��λ.Width - 900 * 2 - 27 * Screen.TwipsPerPixelX
End Sub

Private Sub LoadPara()
'ϵͳ������
    
    Dim rsTemp As New ADODB.Recordset
    Dim strTemp As String
    Dim i As Integer
    
    '���ȶԷ������ͽ��г�ʼ��
    Call Load��������
    
    On Error GoTo ErrHandle
    gstrSQL = "select ������,����ֵ from Zlparameters Where ϵͳ = " & glngSys & " And Nvl(˽��, 0) = 0 And ģ�� Is Null Order By ������"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)

    Do Until rsTemp.EOF
        Select Case rsTemp("������")
            Case 1 '�������°�ʱ��
                i = InStr(UCase(rsTemp("����ֵ")), "AND")
                strTemp = Mid(rsTemp("����ֵ"), 1, i - 2)
                dtp(dtp_�����ϰ�).Value = CDate(strTemp)
                strTemp = Mid(rsTemp("����ֵ"), i + 4)
                dtp(dtp_�����°�).Value = CDate(strTemp)
            Case 2 '�������°�ʱ��
                i = InStr(UCase(rsTemp("����ֵ")), "AND")
                strTemp = Mid(rsTemp("����ֵ"), 1, i - 2)
                dtp(dtp_�����ϰ�).Value = CDate(strTemp)
                strTemp = Mid(rsTemp("����ֵ"), i + 4)
                dtp(dtp_�����°�).Value = CDate(strTemp)
            Case 4 '�շ��վ����д�
                If Not IsNull(rsTemp("����ֵ")) Then
                    ud(ud_�շ��վ�).Value = rsTemp("����ֵ")
                End If
            Case 8 'δ��ҩ������ҩ
                chk(chk_δ��ҩ������ҩ) = IIF(rsTemp("����ֵ") <> 0, 1, 0)
            Case 9 'ҩƷ��������
                '�����һ���ؼ���Indexֵ��4
                opt(CInt(IIF(IsNull(rsTemp("����ֵ")), "0", rsTemp("����ֵ"))) + 4).Value = True
            Case 12 '���￨��������ʾ
                chk(chk_������ʾ).Value = IIF(rsTemp("����ֵ") <> 0, 1, 0)
            Case 14 '�շѷֱҴ���
                cmb(cmb_�ֱҴ���).ListIndex = IIF(IsNull(rsTemp("����ֵ")), 0, rsTemp("����ֵ"))
            Case 17 '�������뷽ʽ���ֱ�Ϊ���������￨���Һŵ�������ID
                strTemp = IIF(IsNull(rsTemp("����ֵ")), "1111", rsTemp("����ֵ"))
                chk(chk_��������).Value = Val(Mid(strTemp, 1, 1))
                chk(chk_ˢ���￨).Value = Val(Mid(strTemp, 2, 1))
                chk(chk_����ID).Value = Val(Mid(strTemp, 4, 1))
            Case 18 'ָ��ҩ��ʱ���ƿ��
                chk(chk_�޶�ҩƷ�Ŀ��).Value = IIF(rsTemp("����ֵ") <> 0, 1, 0)
            Case 45 '�շ�ͬʱ��ҩ
                chk(chk_�շ�ͬʱ��ҩ).Value = IIF(rsTemp("����ֵ") <> 0, 1, 0)
            Case 54 'ʱ��ҩƷ�ԼӼ������
                chk(chk_ʱ��ҩƷ���).Value = IIF(rsTemp("����ֵ") <> 0, 1, 0)
            Case 20 '��ʾ����Ʊ�ݵĺ��볤�ȣ���λ�ֱ�Ϊ1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨
                strTemp = IIF(IsNull(rsTemp("����ֵ")), "77777", rsTemp("����ֵ"))
                lvw(lvw_Ʊ��).ListItems("C1").SubItems(1) = IIF(CLng(Mid(strTemp, 1, 1)) = 0, 10, CLng(Mid(strTemp, 1, 1)))
                lvw(lvw_Ʊ��).ListItems("C5").SubItems(1) = IIF(CLng(Mid(strTemp, 5, 1)) = 0, 10, CLng(Mid(strTemp, 5, 1)))
            Case 22 '�ձ�ʱ������
'                chk(chk_�ձ�ʱ��).Value = IIf(rsTemp("����ֵ") <> 0, 1, 0)
            Case 24 '��ʾ�Ƿ��ϸ���ƹ����Ʊ�ݵ�ʹ�ã���λ�ֱ�Ϊ1-�շ�,2-Ԥ��,3-����,4-�Һ�,5-���￨
                strTemp = IIF(IsNull(rsTemp("����ֵ")), "11111", rsTemp("����ֵ"))
                lvw(lvw_Ʊ��).ListItems("C1").SubItems(2) = IIF(Mid(strTemp, 1, 1) = "1", "��", "")
                lvw(lvw_Ʊ��).ListItems("C5").SubItems(2) = IIF(Mid(strTemp, 5, 1) = "1", "��", "")
            Case 29 'ָ�������۶��۵�λ
                cmb(cmb_���۵�λ).ListIndex = IIF(rsTemp("����ֵ") = "1", 1, 0)
            Case 41 'ҽ���������÷�������
                SetListByText lst(lst_ҽ������), Replace(IIF(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ")), "|", ",")
            Case 42 '���Ѳ������÷�������
                SetListByText lst(lst_���Ѳ���), Replace(IIF(IsNull(rsTemp("����ֵ")), "", rsTemp("����ֵ")), "|", ",")
        End Select
        rsTemp.MoveNext
    Loop
    '��ʾ��ǰƱ�ݵ����
    lvw(lvw_Ʊ��).ListItems("C1").Selected = True
    lvw_ItemClick lvw_Ʊ��, lvw(lvw_Ʊ��).SelectedItem
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Load���ݲ���()
    Dim rsTemp As New ADODB.Recordset
    Dim lst As ListItem, str���� As String
    
    On Error GoTo ErrHandle
    gstrSQL = "select A.��ԱID,B.����,A.����,A.ʱ������,A.���˵���,A.������� from ���ݲ������� A,��Ա�� B where A.��ԱID=B.ID"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)

    lvw(lvw_����).ListItems.Clear
    Do Until rsTemp.EOF
        Set lst = lvw(lvw_����).ListItems.Add(, , rsTemp("����"), , "Limit")
        
        str���� = Switch(rsTemp("����") = 2, "�շѵ�", rsTemp("����") = 8, "��Ա��")
        lst.SubItems(1) = str����
        lst.SubItems(2) = rsTemp("ʱ������")
        lst.SubItems(3) = IIF(rsTemp("���˵���") = 1, "��", "��")
        lst.SubItems(4) = IIF(IsNull(rsTemp("�������")), "", Format(rsTemp("���˵���"), "0.00"))
        lst.Tag = rsTemp("��ԱID")
        lst.ListSubItems(1).Tag = rsTemp("����")
        
        rsTemp.MoveNext
    Loop
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Load��������()
'���ܣ���ʼ����������
    Dim rsTemp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    rsTemp.CursorLocation = adUseClient
    gstrSQL = "Select ����,���� From �������� Order by ����"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)

    lst(lst_ҽ������).Clear
    lst(lst_���Ѳ���).Clear
    Do Until rsTemp.EOF
        lst(lst_ҽ������).AddItem rsTemp("����") & "." & rsTemp("����")
        lst(lst_���Ѳ���).AddItem rsTemp("����") & "." & rsTemp("����")
        
        rsTemp.MoveNext
    Loop
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadҩƷ����()
'����:װ��ҩƷ��������
    Dim rsTemp As New ADODB.Recordset
    Dim lngRow As Long
    
    On Error GoTo ErrHandle
    With Bill(bill_ҩƷ����)
        '����װ��ⷿ
        rsTemp.CursorLocation = adUseClient
        gstrSQL = "select distinct A.ID,A.����,A.���� " & _
                   " from  ��������˵�� b,���ű� a " & _
                   " where B.�������� in ('��ҩ��','��ҩ��','��ҩ��','�Ƽ���','��ҩ��','��ҩ��','��ҩ��') " & _
                   " and  b.����ID=a.ID and " & Where����ʱ��("A") & " order by ����"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)

        .Clear
        Do Until rsTemp.EOF
            .AddItem rsTemp("����") & "-" & rsTemp("����")
            .ItemData(.NewIndex) = rsTemp("ID")
            
            rsTemp.MoveNext
        Loop
        
        'װ�������������
        gstrSQL = "select A.���ڿⷿID,A.�Է��ⷿID,A.����" & _
                ",B.���� as ���ڱ���,B.���� as ��������,C.���� as �Է�����,C.���� as �Է����� " & _
                " from ҩƷ������� A,���ű� B,���ű� C " & _
                " where A.���ڿⷿID= B.ID and A.�Է��ⷿID=C.ID"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)
        
        lngRow = 1
        Do Until rsTemp.EOF
            .Rows = lngRow + 1
            .RowData(lngRow) = rsTemp("���ڿⷿID")
            .TextMatrix(lngRow, 0) = rsTemp("���ڱ���") & "-" & rsTemp("��������")
            .TextMatrix(lngRow, 1) = rsTemp("�Է�����") & "-" & rsTemp("�Է�����")
            .TextMatrix(lngRow, 2) = rsTemp("�Է��ⷿID")
            .TextMatrix(lngRow, 3) = Switch(rsTemp("����") = 1, "1-���ڿⷿ������Է��ⷿ", _
                                            rsTemp("����") = 2, "2-�Է��ⷿ���������ڿⷿ", _
                                                          True, "3-���ⷿ���˫����ͨ")
            lngRow = lngRow + 1
            rsTemp.MoveNext
        Loop
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub LoadҩƷ�ⷿ��λ()
    Dim rsTmp As New ADODB.Recordset
    Dim lngRow As Long, lngTemp As Long, lng��λ As Long, i As Long, lngMaxRow As Long
    Dim strobjTemp As String, strWorkTemp As String
    Dim blnHave As Boolean
    
    '����ⷿ��λ
    On Error GoTo ErrHandle
    gstrSQL = "" & vbCrLf & _
            "   SELECT b.id,nvl(b.����,'') ����,nvl(b.����,'') ����,a.�������,a.��������" & vbCrLf & _
            "          FROM ��������˵�� A, ���ű� B" & vbCrLf & _
            " WHERE B.ID=A.����ID AND A.�������� IN ('��ҩ��', '��ҩ��', '��ҩ��', '�Ƽ���', '��ҩ��', '��ҩ��', '��ҩ��')  "
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        msf�ⷿ��λ.Rows = 2
        lngTemp = 0
        lngMaxRow = rsTmp.RecordCount
        For lngRow = 1 To lngMaxRow
            rsTmp.Filter = "id=" & rsTmp!ID
            strobjTemp = "": strWorkTemp = ""
            blnHave = False
            For i = 0 To msf�ⷿ��λ.Rows - 1
                If msf�ⷿ��λ.RowData(i) = rsTmp!ID Then
                    blnHave = True
                    Exit For
                End If
            Next
            If blnHave = False Then
                For i = 1 To rsTmp.RecordCount
                    strobjTemp = strobjTemp & rsTmp!�������
                    strWorkTemp = strWorkTemp & rsTmp!��������
                    rsTmp.MoveNext
                Next
                '1-��;2-��;3-ס;4-��
                If InStr(strobjTemp, "2") <> 0 Or InStr(strobjTemp, "3") <> 0 Then
                    'סԺ��λ
                    lng��λ = 1
                ElseIf InStr(strobjTemp, "1") <> 0 Then
                    '���ﵥλ
                    lng��λ = 1
                ElseIf InStr(strWorkTemp, "ҩ��") <> 0 Then
                    'ҩ�ⵥλ
                    lng��λ = 2
                Else
                    '�ۼ۵�λ����Ҫ���Ƽ���
                    lng��λ = 1
                End If
                If lngTemp > 0 Then
                    msf�ⷿ��λ.AddItem ""
                End If
                rsTmp.MoveFirst
                msf�ⷿ��λ.TextMatrix(msf�ⷿ��λ.Rows - 1, 0) = "[" & rsTmp!���� & "]" & rsTmp!����
                msf�ⷿ��λ.TextMatrix(msf�ⷿ��λ.Rows - 1, 1) = ""
                msf�ⷿ��λ.TextMatrix(msf�ⷿ��λ.Rows - 1, 2) = ""
                msf�ⷿ��λ.TextMatrix(msf�ⷿ��λ.Rows - 1, lng��λ) = "��"
                msf�ⷿ��λ.RowData(msf�ⷿ��λ.Rows - 1) = rsTmp!ID
                lngTemp = lngTemp + 1
            End If
            rsTmp.Filter = ""
            rsTmp.MoveFirst
            rsTmp.Move lngRow, adBookmarkFirst
        Next
        gstrSQL = "select * from ҩƷ�ⷿ��λ"
        Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
        If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            lngMaxRow = rsTmp.RecordCount
            For lngRow = 1 To lngMaxRow
                For i = 1 To msf�ⷿ��λ.Rows - 1
                    If rsTmp!�ⷿid = msf�ⷿ��λ.RowData(i) Then
                        msf�ⷿ��λ.TextMatrix(i, 1) = ""
                        msf�ⷿ��λ.TextMatrix(i, 2) = ""
                        msf�ⷿ��λ.TextMatrix(i, IIF(rsTmp!���� = 1, 1, 2)) = "��"
                        Exit For
                    End If
                Next
                rsTmp.MoveNext
            Next
        End If
    Else
        msf�ⷿ��λ.Rows = 2
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Load��Ա��()
'���ܣ��õ���Ա���۸�
    Dim rsTemp As New ADODB.Recordset
    
    mlng��Ա��ID = 0
    mlng��ĿID = 0
    
    On Error GoTo ErrHandle
    rsTemp.CursorLocation = adUseClient
    gstrSQL = "Select ID,����,�Ƿ��� From �շ�ϸĿ where ĩ��=1 and ���='Z' and ����='��Ա��'"
    Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption)

    If rsTemp.RecordCount > 0 Then
        mlng��Ա��ID = rsTemp("ID")
        mstr��Ա������ = rsTemp("����")
        chk���.Value = IIF(rsTemp("�Ƿ���") = 1, 1, 0)
        
        '��ü۸���Ϣ
        rsTemp.Close
        
        gstrSQL = "Select A.ID,A.������ĿID,A.�ּ�,B.���� From �շѼ�Ŀ A,������Ŀ B " & _
                  "where A.������ĿID=B.ID and A.��ֹ����=to_date('3000-01-01','yyyy-MM-dd') and A.�շ�ϸĿID=[1]"
        Set rsTemp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, mlng��Ա��ID)
                
        If rsTemp.RecordCount > 0 Then
            mlng��ĿID = rsTemp("ID")
            txt�۸�.Text = Format(rsTemp("�ּ�"), "###########0.000;-##########0.000;0.000;0.000")
            txt������Ŀ.Text = rsTemp("����")
            txt������Ŀ.Tag = rsTemp("������ĿID")
        End If
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Form_Unload(Cancel As Integer)
    '�����п�
    SaveFlexState Bill(bill_ҩƷ����), App.ProductName & "\" & Me.Name & bill_ҩƷ����
    
    If mblnChange = False Then Exit Sub
    If MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
        Cancel = 1
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If IsValid() = False Then Exit Sub
    If Save����() = False Then Exit Sub
    mblnChange = False
    Unload Me
End Sub

Private Function IsValid() As Boolean
    Dim lngRow As Long, lngTemp As Long
    
    If IsNumeric(txt�۸�.Text) = False Then
        MsgBox "��������ȷ�Ļ�Ա���۸�", vbInformation, gstrSysName
        Call ShowTab(1)
        txt�۸�.SetFocus
        Exit Function
    End If
    
    If Val(txt�۸�.Text) < 0 Or Val(txt�۸�.Text) > 10000 Then
        MsgBox "��Ա���۸񲻺���", vbInformation, gstrSysName
        Call ShowTab(1)
        txt�۸�.SetFocus
        Exit Function
    End If
    
    If txt������Ŀ.Tag = "" Then
        MsgBox "��Ϊ��Ա��ѡ��������Ŀ��", vbInformation, gstrSysName
        Call ShowTab(1)
        txt������Ŀ.SetFocus
        Exit Function
    End If
    
    With Bill(bill_ҩƷ����)
        For lngRow = 1 To .Rows - 1
            If .TextMatrix(lngRow, 0) = "" And .TextMatrix(lngRow, 1) <> "" Or .TextMatrix(lngRow, 0) <> "" And .TextMatrix(lngRow, 1) = "" Then
                MsgBox "��" & lngRow & "����Ϣ��������", vbInformation, gstrSysName
                .Row = lngRow
                .Col = 0
                Call ShowTab(5)
                Exit Function
            End If
            If .RowData(lngRow) > 0 And .RowData(lngRow) = Val(.TextMatrix(lngRow, 2)) Then
                MsgBox "��" & lngRow & "�������ڿⷿ��Է��ⷿ��ͬ��", vbInformation, gstrSysName
                .Row = lngRow
                .Col = 0
                Call ShowTab(5)
                Exit Function
            End If
            
            For lngTemp = lngRow + 1 To .Rows - 1
                If .RowData(lngRow) = .RowData(lngTemp) And Val(.TextMatrix(lngRow, 2)) = Val(.TextMatrix(lngTemp, 2)) Then
                    MsgBox "��" & lngRow & "�����" & lngTemp & "����Ϣ�ⷿ��ͬ�ˡ�", vbInformation, gstrSysName
                    .Row = lngTemp
                    .Col = 0
                    Call ShowTab(5)
                    Exit Function
                End If
            Next
        Next
    End With
    
    IsValid = True
End Function

Private Function Save����() As Boolean
    On Error GoTo ErrHandle
    gcnOracle.BeginTrans
    
    Call SavePara
    Call SaveҩƷ����
    Call Save�ⷿ��λ
    
    If Save��Ա�� = False Then
        '���ڸù��̵�SQL���Ƚ϶࣬���Ե����Ĵ�����
        gcnOracle.RollbackTrans
        Exit Function
    End If
    
    '������ϣ������ύ
    gcnOracle.CommitTrans
    Save���� = True
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub Save�ⷿ��λ()
    '����ⷿ��λ����
    Dim i As Long
    Dim lngTmp As Long
    
    On Error GoTo ErrHandle
    If msf�ⷿ��λ.Rows > 1 And Trim(msf�ⷿ��λ.TextMatrix(1, 0)) <> "" Then
        gstrSQL = ""
        For i = 1 To msf�ⷿ��λ.Rows - 1
            gstrSQL = gstrSQL & msf�ⷿ��λ.RowData(i) & ","
            lngTmp = 1
            Select Case True
                Case msf�ⷿ��λ.TextMatrix(i, 1) = "��"
                    lngTmp = 1
                Case msf�ⷿ��λ.TextMatrix(i, 2) = "��"
                    lngTmp = 4
            End Select
            gstrSQL = gstrSQL & lngTmp & ","
        Next
        gstrSQL = "ZL_ҩƷ�ⷿ��λ_INSERT('" & gstrSQL & "')"
        Call gcnOracle.Execute(gstrSQL)
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub SavePara()
    Dim strTemp As String
    Dim lngTemp As Long
    
    '����Բ������б���
    On Error GoTo ErrHandle
    strTemp = "1," & Format(dtp(dtp_�����ϰ�).Value, "HH:mm") & " AND " & Format(dtp(dtp_�����°�).Value, "HH:mm") & ","
    strTemp = strTemp & "2," & Format(dtp(dtp_�����ϰ�).Value, "HH:mm") & " AND " & Format(dtp(dtp_�����°�).Value, "HH:mm") & ","
    strTemp = strTemp & "4," & ud(ud_�շ��վ�).Value & ","
    strTemp = strTemp & "8," & chk(chk_δ��ҩ������ҩ).Value & ","
    If opt(opt_�����п����).Value = True Then
        lngTemp = 0
    ElseIf opt(opt_��������).Value = True Then
        lngTemp = 1
    Else
        lngTemp = 2
    End If
    strTemp = strTemp & "9," & lngTemp & ","
    strTemp = strTemp & "12," & chk(chk_������ʾ).Value & ","
    strTemp = strTemp & "14," & cmb(cmb_�ֱҴ���).ListIndex & ","
    strTemp = strTemp & "17," & chk(chk_��������).Value & chk(chk_ˢ���￨).Value & "0" & chk(chk_����ID).Value & ","
    strTemp = strTemp & "18," & chk(chk_�޶�ҩƷ�Ŀ��).Value & ","
    strTemp = strTemp & "45," & chk(chk_�շ�ͬʱ��ҩ).Value & ","
    strTemp = strTemp & "54," & chk(chk_ʱ��ҩƷ���).Value & ","
    strTemp = strTemp & "20,"
    strTemp = strTemp & IIF(lvw(lvw_Ʊ��).ListItems("C1").SubItems(1) = 10, "0", lvw(lvw_Ʊ��).ListItems("C1").SubItems(1))
    strTemp = strTemp & "777" '�м������Ʊ��ûʹ��
    strTemp = strTemp & IIF(lvw(lvw_Ʊ��).ListItems("C5").SubItems(1) = 10, "0", lvw(lvw_Ʊ��).ListItems("C5").SubItems(1)) & ","
    strTemp = strTemp & "24,"
    strTemp = strTemp & IIF(lvw(lvw_Ʊ��).ListItems("C1").SubItems(2) = "��", "1", "0")
    strTemp = strTemp & "000" '�м������Ʊ��ûʹ��
    strTemp = strTemp & IIF(lvw(lvw_Ʊ��).ListItems("C5").SubItems(2) = "��", "1", "0") & ","
    strTemp = strTemp & "29," & cmb(cmb_���۵�λ).ListIndex & ","
    'ע�ⷵ��ֵ����,�ָ��������������š�����ʱҪת��һ��
    strTemp = strTemp & "41," & Replace(Replace(GetTextFromList(lst(lst_ҽ������)), "'", ""), ",", "|") & ","
    strTemp = strTemp & "42," & Replace(Replace(GetTextFromList(lst(lst_���Ѳ���)), "'", ""), ",", "|") & ","
    
    gstrSQL = "zl_Parameters_Update_Batch(" & glngSys & ",'" & strTemp & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub Save���ݲ���()
    Dim lst As ListItem
    
    '����ɾ����ǰ�����е��ݲ���
    On Error GoTo ErrHandle
    gstrSQL = "zl_���ݲ�������_Delete"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '�������µ�
    For Each lst In lvw(lvw_����).ListItems
        gstrSQL = "zl_���ݲ�������_Insert(" & lst.Tag & "," & lst.ListSubItems(1).Tag & _
                    "," & lst.SubItems(2) & "," & IIF(lst.SubItems(3) = "��", 1, 0) & "," & IIF(lst.SubItems(4) = "", "NULL", lst.SubItems(4)) & " )"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Next
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub
    
Private Sub SaveҩƷ����()
    Dim strTemp As String
    Dim lngRow As Long
    Dim str���� As String
    
    On Error GoTo ErrHandle
    With Bill(bill_ҩƷ����)
        For lngRow = 1 To .Rows - 1
            If .RowData(lngRow) > 0 Then
                str���� = Left(.TextMatrix(lngRow, 3), 1)
                If str���� = "" Then str���� = "3"
                
                strTemp = strTemp & .RowData(lngRow) & "," & Val(.TextMatrix(lngRow, 2)) & "," & str���� & ","
            End If
        Next
    End With
    
    gstrSQL = "zl_ҩƷ�������_Modify('" & strTemp & "')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function Save��Ա��() As Boolean
'���ܣ�����Ի�Ա��������
    Dim lngϸĿID As Long
    Dim lng��ĿID As Long
    Dim str���� As String
    Dim oldlng�ϼ� As Long
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    If mlng��Ա��ID = 0 Then
        '�����շ�ϸĿ
        lngϸĿID = zlDatabase.GetNextId("�շ�ϸĿ")
        str���� = GetMaxLocalCode("", "�շ�ϸĿ", " and ���='Z' ")
        
        gstrSQL = "zl_�շ�ϸĿ_insert(" & lngϸĿID & ",'Z','" & str���� & "','','','��Ա��','HYK',1" & _
            ",'','','��','',0," & chk���.Value & ",0,null,null,0,'')"
    Else
        '�޸��շ�ϸĿ
        lngϸĿID = mlng��Ա��ID
        gstrSQL = "select * from �շ�ϸĿ where id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, lngϸĿID)
        
        If rsTmp.RecordCount < 1 Then
            MsgBox "��Ա����Ŀ�����ڣ�", vbInformation, gstrSysName
            Exit Function
        End If
        oldlng�ϼ� = zlCommFun.Nvl(rsTmp!�ϼ�id, 0)
        gstrSQL = "zl_�շ�ϸĿ_update(" & lngϸĿID & ",'" & mstr��Ա������ & "','','','��Ա��','HYK'" & _
            IIF(oldlng�ϼ� = 0, ", Null", "," & oldlng�ϼ�) & ",'','','��','',0," & chk���.Value & ",0,null,0,'')"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    If mlng��ĿID = 0 Then
        '�����۸�
        lng��ĿID = zlDatabase.GetNextId("�շѼ�Ŀ")
        gstrSQL = "zl_�շѼ�Ŀ_insert(" & _
           lng��ĿID & ",null," & lngϸĿID & "," & txt������Ŀ.Tag & ",0," & txt�۸�.Text & _
           ",0,0,''," & lng��ĿID & ",'" & gstrUserName & "',sysdate)"
    Else
        '�޸ļ۸�
        gstrSQL = "zl_�շѼ�Ŀ_update(" & lngϸĿID & "," & txt������Ŀ.Tag & ",0," & txt�۸�.Text & _
             ",0,0,''," & mlng��ĿID & ",'" & gstrUserName & "')"
    End If
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    '�����ض��շ���Ŀ
    gstrSQL = "zl_�շ��ض���Ŀ_Modify('���￨," & lngϸĿID & ",')"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    
    Save��Ա�� = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
End Function

Private Sub chk_Click(Index As Integer)
    mblnChange = True
    If Index = chk_Ʊ�ſ��� Then
        lvw(lvw_Ʊ��).SelectedItem.SubItems(2) = IIF(chk(Index).Value = 1, "��", "")
    End If
End Sub

Private Sub chk_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub lst_ItemCheck(Index As Integer, Item As Integer)
    mblnChange = True
End Sub

Private Sub lst_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub cmb_Click(Index As Integer)
    mblnChange = True
End Sub

Private Sub cmb_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub dtp_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub dtp_Change(Index As Integer)
    Dim intNext As Integer

    mblnChange = True
    If Index < dtp_�����°� Then
        intNext = Index + 1
        
        dtp(intNext).MinDate = dtp(Index).Value
        If dtp(intNext).Value < dtp(intNext).MinDate Then
            dtp(intNext).Value = dtp(intNext).MinDate
            dtp_Change intNext
        End If
    End If
End Sub

Private Sub lvw_ColumnClick(Index As Integer, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
     If Index = lvw_���� Then
        If mintColumn = ColumnHeader.Index - 1 Then '���Ǹղ�����
            lvw(lvw_����).SortOrder = IIF(lvw(lvw_����).SortOrder = lvwAscending, lvwDescending, lvwAscending)
        Else
            mintColumn = ColumnHeader.Index - 1
            lvw(lvw_����).SortKey = mintColumn
            lvw(lvw_����).SortOrder = lvwAscending
        End If
     End If
End Sub

Private Sub lvw_DblClick(Index As Integer)
    If Index = lvw_���� Then
        Call cmdOperate_Click(1)
    End If
End Sub

Private Sub lvw_ItemCheck(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    mblnChange = True
    
    Dim itemTemp As MSComctlLib.ListItem
    For Each itemTemp In lvw(Index).ListItems
        If Not itemTemp Is Item Then
            itemTemp.Checked = False
        End If
    Next
End Sub

Private Sub lvw_ItemClick(Index As Integer, ByVal Item As MSComctlLib.ListItem)
    Dim lngԭֵ As Long
    
    If Index = lvw_Ʊ�� Then
        lngԭֵ = Val(Item.SubItems(1))
        
        If Item.Text = "���￨" Then
            ud(ud_���볤��).Max = 8
        Else
            ud(ud_���볤��).Max = 10
        End If
        '�������ֵʱ�������Ѿ��������б��е�ֵ
        ud(ud_���볤��).Value = lngԭֵ
        chk(chk_Ʊ�ſ���).Value = IIF(Item.SubItems(2) = "��", 1, 0)
    End If
End Sub

Private Sub lvw_KeyPress(Index As Integer, KeyAscii As Integer)
    If Index = lvw_Ʊ�� Then
        If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
    Else
        If KeyAscii = vbKeyReturn Then cmdOperate_Click (1)
    End If
End Sub

Private Sub opt_Click(Index As Integer)
    mblnChange = True
End Sub

Private Sub opt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txtUD_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txt�۸�_GotFocus()
    zlControl.TxtSelAll txt�۸�
End Sub

Private Sub txt�۸�_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub txt������Ŀ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        SendKeys "{TAB}"
    Else
        If KeyAscii = Asc("*") Then
            Call cmdSelect_Click
        End If
    End If
End Sub

Private Sub chk���_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then SendKeys "{TAB}"
End Sub

Private Sub ud_Change(Index As Integer)
    mblnChange = True
    '��̬�ı�Ʊ�ų���
    If Index = ud_���볤�� Then
        lvw(lvw_Ʊ��).SelectedItem.SubItems(1) = ud(ud_���볤��).Value
    End If
End Sub

Private Sub bill_cboClick(Index As Integer, ListIndex As Long)
    If Index <> bill_ҩƷ���� Then Exit Sub
    
    With Bill(bill_ҩƷ����)
        If ListIndex < 0 Then Exit Sub
        If .Col = 0 Then
            .RowData(.Row) = .ItemData(ListIndex)
        Else
            .TextMatrix(.Row, 2) = .ItemData(ListIndex)
        End If
        .TextMatrix(.Row, .Col) = .CboText
        
        If .TextMatrix(.Row, 3) = "" Then .TextMatrix(.Row, 3) = "3-���ⷿ���˫����ͨ"
    End With
End Sub

Private Sub bill_cboKeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    With Bill(Index)
        If .ListIndex < 0 Then Exit Sub
        If KeyCode = vbKeyReturn Then
            If Index = bill_ҩƷ���� And .Col = 1 Then
                .TextMatrix(.Row, 2) = .ItemData(.ListIndex)
            Else
                .RowData(.Row) = .ItemData(.ListIndex)
            End If
            
            If Index = bill_ҩƷ���� Then
                If .TextMatrix(.Row, 3) = "" Then .TextMatrix(.Row, 3) = "3-���ⷿ���˫����ͨ"
            End If
        End If
    End With
End Sub

Private Sub bill_DblClick(Index As Integer, Cancel As Boolean)
'�������һ�еı仯
With Bill(Index)
    If .MouseRow = 0 Then Exit Sub
    
    If Index = bill_ҩƷ���� Then
        If .MouseCol <> .Cols - 1 Then Exit Sub
        Select Case Left(.TextMatrix(.Row, .Col), 1)
            Case "1"
                .TextMatrix(.Row, .Col) = "2-�Է��ⷿ���������ڿⷿ"
            Case "2"
                .TextMatrix(.Row, .Col) = "3-���ⷿ���˫����ͨ"
            Case Else
                .TextMatrix(.Row, .Col) = "1-���ڿⷿ������Է��ⷿ"
        End Select
    End If
    mblnChange = True
End With
    
End Sub

Private Sub bill_KeyPress(Index As Integer, KeyAscii As Integer)
With Bill(Index)
    If Index = bill_ҩƷ���� Then
        If .Col = 3 Then
            Select Case KeyAscii
                Case Asc(" ")
                    '�л������־
                    Select Case Left(.TextMatrix(.Row, .Col), 1)
                        Case "1"
                            .TextMatrix(.Row, .Col) = "2-�Է��ⷿ���������ڿⷿ"
                        Case "2"
                            .TextMatrix(.Row, .Col) = "3-���ⷿ���˫����ͨ"
                        Case Else
                            .TextMatrix(.Row, .Col) = "1-���ڿⷿ������Է��ⷿ"
                    End Select
                    mblnChange = True
                Case vbKey1
                    .TextMatrix(.Row, .Col) = "1-���ڿⷿ������Է��ⷿ"
                    mblnChange = True
                Case vbKey2
                    .TextMatrix(.Row, .Col) = "2-�Է��ⷿ���������ڿⷿ"
                    mblnChange = True
                Case vbKey3
                    .TextMatrix(.Row, .Col) = "3-���ⷿ���˫����ͨ"
                    mblnChange = True
            End Select
        End If
    End If
End With

End Sub

Private Sub tabMain_Click()
    Dim i As Integer
    
    For i = fraMain.LBound To fraMain.UBound
        fraMain(i).Visible = False
    Next
    
    i = tabMain.SelectedItem.Index - 1
    fraMain(i).Visible = True
    Select Case tabMain.SelectedItem.Index
        Case 1 '����
            cmb(cmb_���۵�λ).SetFocus
        Case 2 'Ʊ�ݹ���
            txtUD(ud_�շ��վ�).SetFocus
        Case 3 'Ȩ��
            lst(lst_ҽ������).SetFocus
        Case 4 '����
            lvw(lvw_����).SetFocus
        Case 5 'ҩƷ����
            Bill(bill_ҩƷ����).SetFocus
        Case 6  'ҩƷ�ⷿ��λ
            msf�ⷿ��λ.SetFocus
    End Select
End Sub

Private Sub ShowTab(ByVal intTab As Integer)
    tabMain.Tabs(intTab).Selected = True
    tabMain_Click
End Sub

Private Function NumIsValid(ByVal strNumber As String) As Boolean
'����:�������������Ƿ���һ����Ч������
'����:strNumber  ��������
'����ֵ:��Ч����True,����ΪFalse
    NumIsValid = False
    If Not IsNumeric(strNumber) Then
        MsgBox "������һ����ֵ��", vbInformation, gstrSysName
        Exit Function
    End If
    If Val(strNumber) > 9999999999.999 Then
        MsgBox "�����̫���ˡ�", vbInformation, gstrSysName
        Exit Function
    End If
    If Val(strNumber) < 0 Then
        MsgBox "����Ϊ������", vbInformation, gstrSysName
        Exit Function
    End If
    NumIsValid = True
End Function


