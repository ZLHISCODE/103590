VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{876E3FF4-6E21-11D5-AF7D-0080C8EC27A9}#1.5#0"; "ZL9BillEdit.ocx"
Object = "{A8E5842E-102B-4289-9D57-3B3F5B5E15D3}#9.60#0"; "CODEJO~1.OCX"
Begin VB.Form frmChargeItem 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�շ���Ŀ����"
   ClientHeight    =   7650
   ClientLeft      =   1155
   ClientTop       =   2520
   ClientWidth     =   7260
   Icon            =   "frmChargeItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7650
   ScaleWidth      =   7260
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   WhatsThisHelp   =   -1  'True
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   5850
      Index           =   2
      Left            =   60
      TabIndex        =   78
      Top             =   960
      Visible         =   0   'False
      Width           =   6840
      Begin VB.PictureBox pic�۸�ȼ� 
         BorderStyle     =   0  'None
         Height          =   4545
         Index           =   0
         Left            =   240
         ScaleHeight     =   4545
         ScaleWidth      =   6345
         TabIndex        =   113
         TabStop         =   0   'False
         Top             =   660
         Width           =   6345
         Begin ZL9BillEdit.BillEdit msh��Ŀ 
            Height          =   3420
            Index           =   0
            Left            =   0
            TabIndex        =   115
            Top             =   100
            Width           =   6240
            _ExtentX        =   11007
            _ExtentY        =   6033
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
         Begin VB.TextBox txt����˵�� 
            Height          =   300
            Index           =   0
            Left            =   1155
            MaxLength       =   100
            TabIndex        =   84
            Top             =   3960
            Width           =   5070
         End
         Begin VB.CheckBox chkNow 
            Caption         =   "����ִ��(&N)"
            Height          =   225
            Index           =   0
            Left            =   4170
            TabIndex        =   82
            Top             =   3600
            Width           =   1785
         End
         Begin MSComCtl2.DTPicker dtpBegin 
            Height          =   285
            Index           =   0
            Left            =   1170
            TabIndex        =   81
            Top             =   3570
            Width           =   2535
            _ExtentX        =   4471
            _ExtentY        =   503
            _Version        =   393216
            CustomFormat    =   "yyyy��MM��dd�� HH:mm:ss"
            Format          =   111214595
            CurrentDate     =   36444
            MaxDate         =   401768
         End
         Begin VB.Label lbl����ִ��ʱ�� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "ִ������(&B)"
            Height          =   180
            Index           =   0
            Left            =   75
            TabIndex        =   80
            Top             =   3615
            Width           =   1050
         End
         Begin VB.Label lbl����˵�� 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "����˵��(&X)"
            Height          =   180
            Index           =   0
            Left            =   60
            TabIndex        =   83
            Top             =   4020
            Width           =   990
         End
      End
      Begin XtremeSuiteControls.TabControl tbPriceGrade 
         Height          =   4800
         Left            =   240
         TabIndex        =   114
         TabStop         =   0   'False
         Top             =   870
         Width           =   6375
         _Version        =   589884
         _ExtentX        =   11245
         _ExtentY        =   8467
         _StockProps     =   64
      End
      Begin VB.Image img��Ŀ 
         Height          =   600
         Left            =   270
         Picture         =   "frmChargeItem.frx":000C
         Stretch         =   -1  'True
         Top             =   210
         Width           =   600
      End
      Begin VB.Label lblEdit 
         Caption         =   "    �˴������շ���Ŀ�ļ۸񣬵����Ǳ��ʱ��ֻ��ѡ��һ��������Ŀ��"
         Height          =   435
         Index           =   14
         Left            =   1170
         TabIndex        =   79
         Top             =   300
         Width           =   3795
      End
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   6555
      Index           =   4
      Left            =   1455
      TabIndex        =   87
      Top             =   6165
      Visible         =   0   'False
      Width           =   6840
      Begin VB.CommandButton cmdFind 
         Caption         =   "����(&F)"
         Enabled         =   0   'False
         Height          =   300
         Left            =   5040
         TabIndex        =   107
         Top             =   3405
         Width           =   855
      End
      Begin VB.TextBox txtFind 
         BackColor       =   &H8000000F&
         Enabled         =   0   'False
         Height          =   300
         Left            =   4080
         TabIndex        =   106
         Top             =   3405
         Width           =   975
      End
      Begin VB.OptionButton optʹ�ÿ��� 
         Caption         =   "ȫԺ"
         Height          =   180
         Index           =   1
         Left            =   3240
         TabIndex        =   99
         Top             =   3480
         Value           =   -1  'True
         Width           =   1095
      End
      Begin VB.OptionButton optʹ�ÿ��� 
         Caption         =   "ָ������"
         Height          =   180
         Index           =   0
         Left            =   2040
         TabIndex        =   98
         Top             =   3480
         Width           =   1095
      End
      Begin ZL9BillEdit.BillEdit msh���� 
         Height          =   2175
         Left            =   240
         TabIndex        =   75
         Top             =   840
         Width           =   6270
         _ExtentX        =   11060
         _ExtentY        =   3836
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
      Begin MSComctlLib.ListView Lvw���� 
         Height          =   1980
         Left            =   240
         TabIndex        =   96
         Top             =   3840
         Width           =   6270
         _ExtentX        =   11060
         _ExtentY        =   3493
         View            =   2
         Arrange         =   1
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         Enabled         =   0   'False
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "����"
            Object.Width           =   5292
         EndProperty
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "��ǰ�����ʹ�÷�Χ"
         Height          =   180
         Left            =   240
         TabIndex        =   97
         Top             =   3480
         Width           =   1620
      End
      Begin VB.Label lbl�����ϼ� 
         Alignment       =   1  'Right Justify
         Caption         =   "�ϼ�:##.##"
         Height          =   180
         Left            =   4680
         TabIndex        =   90
         Top             =   3060
         Width           =   1695
      End
      Begin VB.Image img���� 
         Height          =   600
         Left            =   270
         Picture         =   "frmChargeItem.frx":0246
         Stretch         =   -1  'True
         Top             =   120
         Width           =   600
      End
      Begin VB.Label lblEdit 
         Caption         =   "    ������Ŀ��ָ�û��ڽ��е���¼���У����������շ���Ŀ�����Ӷ��Զ����ӵ��շ���Ŀ��"
         Height          =   435
         Index           =   13
         Left            =   1140
         TabIndex        =   88
         Top             =   240
         Width           =   5370
      End
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   6660
      Index           =   3
      Left            =   0
      TabIndex        =   85
      Top             =   450
      Visible         =   0   'False
      Width           =   6870
      Begin VB.Frame Frame1 
         Height          =   4785
         Left            =   150
         TabIndex        =   91
         Top             =   0
         Width           =   6585
         Begin VB.Frame Frame2 
            Height          =   120
            Left            =   195
            TabIndex        =   92
            Top             =   660
            Width           =   6135
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "���������ڿ���(&F)"
            Height          =   195
            Index           =   6
            Left            =   4380
            TabIndex        =   95
            Top             =   450
            Width           =   1860
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "Ժ��ִ��(&E)"
            Height          =   195
            Index           =   5
            Left            =   4395
            TabIndex        =   94
            Top             =   825
            Visible         =   0   'False
            Width           =   1860
         End
         Begin VB.TextBox txt����ִ�� 
            Enabled         =   0   'False
            Height          =   300
            Left            =   1305
            MaxLength       =   30
            TabIndex        =   7
            Top             =   1065
            Width           =   1785
         End
         Begin VB.TextBox txtסԺִ�� 
            Enabled         =   0   'False
            Height          =   300
            Left            =   4065
            MaxLength       =   30
            TabIndex        =   9
            Top             =   1065
            Width           =   1905
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "����ȷִ�п���(&N)"
            Height          =   180
            Index           =   0
            Left            =   210
            TabIndex        =   1
            Top             =   210
            Value           =   -1  'True
            Width           =   1905
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "�������ڲ���(&B)"
            Height          =   195
            Index           =   2
            Left            =   4380
            TabIndex        =   3
            Top             =   195
            Width           =   1755
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "����Ա���ڿ���(&R)"
            Height          =   195
            Index           =   3
            Left            =   2265
            TabIndex        =   5
            Top             =   450
            Width           =   1920
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "�������ڿ���(&K)"
            Height          =   180
            Index           =   1
            Left            =   2280
            TabIndex        =   2
            Top             =   210
            Width           =   1725
         End
         Begin VB.OptionButton opt���� 
            Caption         =   "ָ������(&D)"
            ForeColor       =   &H00C00000&
            Height          =   195
            Index           =   4
            Left            =   210
            TabIndex        =   4
            Top             =   450
            Width           =   2265
         End
         Begin ZL9BillEdit.BillEdit msf����ִ�� 
            Height          =   3000
            Left            =   405
            TabIndex        =   11
            Top             =   1680
            Width           =   5940
            _ExtentX        =   10478
            _ExtentY        =   5292
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
         Begin MSComctlLib.ImageList imgList 
            Left            =   -210
            Top             =   2640
            _ExtentX        =   1005
            _ExtentY        =   1005
            BackColor       =   -2147483643
            ImageWidth      =   16
            ImageHeight     =   16
            MaskColor       =   12632256
            _Version        =   393216
            BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
               NumListImages   =   3
               BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmChargeItem.frx":0688
                  Key             =   "close"
               EndProperty
               BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmChargeItem.frx":0C22
                  Key             =   "expend"
               EndProperty
               BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
                  Picture         =   "frmChargeItem.frx":11BC
                  Key             =   "Dept"
               EndProperty
            EndProperty
         End
         Begin VB.Label lbl����ִ�� 
            AutoSize        =   -1  'True
            Caption         =   "����(&O)"
            Height          =   180
            Left            =   645
            TabIndex        =   6
            Top             =   1125
            Width           =   630
         End
         Begin VB.Label lbl����ִ�� 
            AutoSize        =   -1  'True
            Caption         =   "2��ָ�����˿��ң�"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   225
            TabIndex        =   10
            Top             =   1455
            Width           =   1530
         End
         Begin VB.Label lblһ����� 
            AutoSize        =   -1  'True
            Caption         =   "1����ָ�����˿����⣺"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   195
            TabIndex        =   93
            Top             =   855
            Width           =   1890
         End
         Begin VB.Label lblסԺִ�� 
            AutoSize        =   -1  'True
            Caption         =   "סԺ(&I)"
            Height          =   180
            Left            =   3405
            TabIndex        =   8
            Top             =   1125
            Width           =   630
         End
      End
      Begin VB.PictureBox picDept 
         BackColor       =   &H00C0C0C0&
         BorderStyle     =   0  'None
         Height          =   2655
         Left            =   3600
         ScaleHeight     =   2655
         ScaleWidth      =   3000
         TabIndex        =   101
         Top             =   1920
         Visible         =   0   'False
         Width           =   3000
         Begin VB.CheckBox ChkSelect 
            Appearance      =   0  'Flat
            BackColor       =   &H00C0C0C0&
            Caption         =   "ȫѡ"
            ForeColor       =   &H80000008&
            Height          =   225
            Left            =   2115
            TabIndex        =   103
            TabStop         =   0   'False
            Top             =   88
            Width           =   675
         End
         Begin VB.ComboBox cboProperty 
            Height          =   300
            Left            =   795
            Style           =   2  'Dropdown List
            TabIndex        =   102
            Top             =   50
            Width           =   1215
         End
         Begin MSComctlLib.ListView lvwItems 
            Height          =   2040
            Left            =   50
            TabIndex        =   104
            Top             =   380
            Width           =   2715
            _ExtentX        =   4789
            _ExtentY        =   3598
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
            BackColor       =   &H00C0C0C0&
            Caption         =   "��������"
            Height          =   180
            Left            =   50
            TabIndex        =   105
            Top             =   110
            Width           =   720
         End
      End
      Begin VB.Frame fra���� 
         Caption         =   "Ӧ�÷�Χ"
         Height          =   1530
         Left            =   150
         TabIndex        =   86
         Top             =   4920
         Visible         =   0   'False
         Width           =   6585
         Begin VB.OptionButton optApply 
            Caption         =   "Ӧ���ڸ÷�����������Ŀ(&L)"
            Height          =   285
            Index           =   2
            Left            =   210
            TabIndex        =   14
            Top             =   885
            Width           =   6270
         End
         Begin VB.OptionButton optApply 
            Caption         =   "Ӧ���ڸ������������Ŀ(&U)"
            Height          =   225
            Index           =   3
            Left            =   210
            TabIndex        =   15
            Top             =   1215
            Width           =   6315
         End
         Begin VB.OptionButton optApply 
            Caption         =   "Ӧ����ͬ����������Ŀ(&G)"
            Height          =   285
            Index           =   1
            Left            =   210
            TabIndex        =   13
            Top             =   555
            Width           =   6285
         End
         Begin VB.OptionButton optApply 
            Caption         =   "���Ա���Ŀ������(&W)"
            Height          =   285
            Index           =   0
            Left            =   225
            TabIndex        =   12
            Top             =   225
            Value           =   -1  'True
            Width           =   6240
         End
      End
   End
   Begin VB.CheckBox chk���� 
      Caption         =   "�½���һ��ʱ���沿����Ϣ"
      Height          =   255
      Left            =   4665
      TabIndex        =   89
      TabStop         =   0   'False
      ToolTipText     =   "��������ʱ�Ƿ������������"
      Top             =   75
      Visible         =   0   'False
      Width           =   2565
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   350
      Left            =   345
      TabIndex        =   71
      Tag             =   "����"
      Top             =   7200
      Width           =   1100
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   350
      Left            =   4170
      TabIndex        =   69
      Tag             =   "����"
      Top             =   7200
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��(&C)"
      Height          =   350
      Left            =   5490
      TabIndex        =   70
      Tag             =   "����"
      Top             =   7200
      Width           =   1100
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeItem.frx":1756
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmChargeItem.frx":1A70
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Frame fra 
      BorderStyle     =   0  'None
      Height          =   6405
      Index           =   1
      Left            =   210
      TabIndex        =   76
      Top             =   360
      Visible         =   0   'False
      Width           =   6795
      Begin VB.PictureBox picTwo 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1035
         Left            =   155
         ScaleHeight     =   1035
         ScaleWidth      =   6735
         TabIndex        =   59
         Top             =   5400
         Width           =   6735
         Begin VB.CommandButton cmd���� 
            Caption         =   "��"
            Height          =   240
            Left            =   6240
            TabIndex        =   108
            TabStop         =   0   'False
            Tag             =   "����"
            ToolTipText     =   "��*��ѡ����"
            Top             =   750
            Width           =   255
         End
         Begin VB.ComboBox cbo¼��������Χ 
            Height          =   300
            Left            =   4350
            Style           =   2  'Dropdown List
            TabIndex        =   63
            Top             =   0
            Width           =   2205
         End
         Begin VB.TextBox txt¼������ 
            Height          =   300
            Left            =   1065
            MaxLength       =   13
            TabIndex        =   61
            Top             =   15
            Width           =   1170
         End
         Begin VB.ComboBox cmb����ȷ�� 
            Height          =   300
            Left            =   1065
            Style           =   2  'Dropdown List
            TabIndex        =   65
            Top             =   360
            Width           =   1875
         End
         Begin VB.CheckBox chk����ȷ�Ϸ�Χ 
            Caption         =   "����ȷ��Ӧ���ڵ�ǰ����������Ŀ"
            Height          =   255
            Left            =   3120
            TabIndex        =   66
            Top             =   360
            Width           =   3495
         End
         Begin VB.ComboBox cmbStationNo 
            Height          =   300
            Left            =   1065
            Style           =   2  'Dropdown List
            TabIndex        =   68
            Top             =   690
            Visible         =   0   'False
            Width           =   1890
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   15
            Left            =   4365
            MaxLength       =   40
            TabIndex        =   109
            ToolTipText     =   "��*��ѡ����"
            Top             =   720
            Width           =   2205
         End
         Begin VB.Label lbl������Ŀ 
            Caption         =   "������Ŀ(&F)"
            Height          =   255
            Left            =   3360
            TabIndex        =   110
            Top             =   750
            Width           =   990
         End
         Begin VB.Label Label4 
            AutoSize        =   -1  'True
            Caption         =   "¼������Ӧ����"
            Height          =   180
            Left            =   3015
            TabIndex        =   62
            Top             =   75
            Width           =   1260
         End
         Begin VB.Label Label3 
            AutoSize        =   -1  'True
            Caption         =   "¼������(&P)"
            Height          =   180
            Left            =   0
            TabIndex        =   60
            Top             =   75
            Width           =   990
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            Caption         =   "����ȷ��(&Q)"
            Height          =   180
            Left            =   0
            TabIndex        =   64
            Top             =   425
            Width           =   990
         End
         Begin VB.Label lblStationNo 
            AutoSize        =   -1  'True
            Caption         =   "Ժ��(&Z)"
            Height          =   180
            Left            =   345
            TabIndex        =   67
            Top             =   780
            Visible         =   0   'False
            Width           =   630
         End
      End
      Begin VB.PictureBox picOne 
         Appearance      =   0  'Flat
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   1455
         Left            =   160
         ScaleHeight     =   1455
         ScaleWidth      =   6615
         TabIndex        =   44
         Top             =   3480
         Width           =   6615
         Begin VB.ComboBox cmb��Ŀ���� 
            Height          =   300
            Left            =   4335
            Style           =   2  'Dropdown List
            TabIndex        =   111
            Top             =   0
            Visible         =   0   'False
            Width           =   2205
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   4
            Left            =   1050
            MaxLength       =   100
            TabIndex        =   56
            Top             =   1125
            Width           =   5505
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   2
            Left            =   1050
            MaxLength       =   72
            TabIndex        =   46
            Top             =   0
            Width           =   1770
         End
         Begin VB.ComboBox cmb������� 
            Height          =   300
            Left            =   4335
            Style           =   2  'Dropdown List
            TabIndex        =   50
            Top             =   375
            Width           =   2205
         End
         Begin VB.ComboBox cmb���㵥λ 
            Height          =   300
            Left            =   1050
            TabIndex        =   48
            Top             =   375
            Width           =   1755
         End
         Begin VB.ComboBox cmb�������� 
            Height          =   300
            Left            =   4335
            Style           =   2  'Dropdown List
            TabIndex        =   54
            Top             =   750
            Width           =   2205
         End
         Begin VB.TextBox txtEdit 
            Height          =   300
            Index           =   6
            Left            =   1050
            Locked          =   -1  'True
            TabIndex        =   52
            Tag             =   "����"
            Top             =   750
            Width           =   1755
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "��Ŀ����(&B)"
            Height          =   180
            Index           =   10
            Left            =   3285
            TabIndex        =   112
            Top             =   60
            Visible         =   0   'False
            Width           =   990
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "��������(&B)"
            Height          =   180
            Index           =   12
            Left            =   0
            TabIndex        =   51
            Tag             =   "����"
            Top             =   810
            Width           =   990
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "˵��(&X)"
            ForeColor       =   &H80000007&
            Height          =   180
            Index           =   8
            Left            =   360
            TabIndex        =   55
            Top             =   1185
            Width           =   630
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "���(&R)"
            ForeColor       =   &H80000007&
            Height          =   180
            Index           =   4
            Left            =   350
            TabIndex        =   45
            Top             =   60
            Width           =   630
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "���㵥λ(&L)"
            Height          =   180
            Index           =   5
            Left            =   0
            TabIndex        =   47
            Top             =   435
            Width           =   990
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "�������(&J)"
            Height          =   180
            Index           =   6
            Left            =   3285
            TabIndex        =   49
            Top             =   435
            Width           =   990
         End
         Begin VB.Label lblEdit 
            AutoSize        =   -1  'True
            Caption         =   "��������(&F)"
            Height          =   180
            Index           =   7
            Left            =   3290
            TabIndex        =   53
            Top             =   810
            Width           =   990
         End
      End
      Begin VB.CheckBox chk�Զ����� 
         Caption         =   "�������Զ�����(&A)"
         Height          =   210
         Left            =   4965
         TabIndex        =   38
         ToolTipText     =   "��¼����ü�¼ʱ���Ը���Ŀ����ժҪ"
         Top             =   2400
         Width           =   1890
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   13
         Left            =   1210
         TabIndex        =   41
         Top             =   3105
         Width           =   1785
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   14
         Left            =   4485
         TabIndex        =   43
         Top             =   3105
         Width           =   2205
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   12
         Left            =   4695
         MaxLength       =   40
         TabIndex        =   28
         Top             =   885
         Width           =   2025
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   11
         Left            =   1210
         MaxLength       =   100
         TabIndex        =   58
         Top             =   4995
         Width           =   5505
      End
      Begin VB.ComboBox cmbClass 
         Height          =   300
         Left            =   840
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   150
         Width           =   1635
      End
      Begin MSComctlLib.ListView lvwSel 
         Height          =   1635
         Left            =   825
         TabIndex        =   30
         Top             =   -1500
         Visible         =   0   'False
         Width           =   2415
         _ExtentX        =   4260
         _ExtentY        =   2884
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   0
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   9
         Left            =   3570
         MaxLength       =   40
         TabIndex        =   24
         Top             =   525
         Width           =   1605
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   10
         Left            =   6240
         MaxLength       =   40
         TabIndex        =   26
         Top             =   525
         Width           =   465
      End
      Begin VB.CheckBox chk���� 
         Caption         =   "����(&Z)"
         Height          =   210
         Left            =   4965
         TabIndex        =   100
         Top             =   2745
         Width           =   1305
      End
      Begin VB.ComboBox cmb���� 
         Height          =   300
         ItemData        =   "frmChargeItem.frx":1D8A
         Left            =   4965
         List            =   "frmChargeItem.frx":1D8C
         Style           =   2  'Dropdown List
         TabIndex        =   74
         Top             =   2700
         Visible         =   0   'False
         Width           =   1740
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   8
         Left            =   3120
         MaxLength       =   12
         TabIndex        =   33
         Tag             =   "����"
         Top             =   1275
         Width           =   1605
      End
      Begin VB.CheckBox chkժҪ 
         Caption         =   "����ժҪ(&A)"
         Height          =   210
         Left            =   4965
         TabIndex        =   39
         ToolTipText     =   "��¼����ü�¼ʱ���Ը���Ŀ����ժҪ"
         Top             =   2475
         Width           =   1305
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   3
         Left            =   840
         MaxLength       =   12
         TabIndex        =   31
         Tag             =   "����"
         Top             =   1275
         Width           =   1620
      End
      Begin VB.TextBox txtEdit 
         BorderStyle     =   0  'None
         Height          =   180
         Index           =   0
         Left            =   960
         TabIndex        =   73
         Tag             =   "����"
         Text            =   "111111"
         Top             =   570
         Width           =   1485
      End
      Begin ZL9BillEdit.BillEdit mshAlias 
         Height          =   1335
         Left            =   180
         TabIndex        =   34
         Top             =   1665
         Width           =   4590
         _ExtentX        =   8096
         _ExtentY        =   2355
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
      Begin VB.CheckBox chk�Ӱ�Ӽ� 
         Caption         =   "�Ӱ�Ӽ�(&D)"
         Height          =   210
         Left            =   4965
         TabIndex        =   37
         Top             =   2190
         Width           =   1305
      End
      Begin VB.CheckBox chk���ηѱ� 
         Caption         =   "���ηѱ�(&M)"
         Height          =   240
         Left            =   4965
         TabIndex        =   36
         Top             =   1920
         Width           =   1305
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   1
         Left            =   840
         MaxLength       =   40
         TabIndex        =   22
         Top             =   885
         Width           =   2790
      End
      Begin VB.CheckBox chk��� 
         Caption         =   "���(&G)"
         Height          =   210
         Left            =   4965
         TabIndex        =   35
         Top             =   1665
         Width           =   945
      End
      Begin VB.CommandButton cmd�ϼ� 
         Caption         =   "��"
         Height          =   240
         Left            =   6435
         TabIndex        =   20
         TabStop         =   0   'False
         Tag             =   "����"
         ToolTipText     =   "��*��ѡ����"
         Top             =   180
         Width           =   255
      End
      Begin VB.TextBox txtTemp 
         Enabled         =   0   'False
         ForeColor       =   &H000000FF&
         Height          =   300
         Left            =   840
         TabIndex        =   77
         TabStop         =   0   'False
         Tag             =   "����"
         Text            =   "11"
         Top             =   525
         Width           =   1635
      End
      Begin VB.TextBox txtEdit 
         Height          =   300
         Index           =   5
         Left            =   3570
         MaxLength       =   40
         TabIndex        =   19
         ToolTipText     =   "��*��ѡ����"
         Top             =   150
         Width           =   3150
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����޼�(&M)"
         Height          =   180
         Index           =   20
         Left            =   150
         TabIndex        =   40
         Top             =   3180
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����޼�(&N)"
         Height          =   180
         Index           =   21
         Left            =   3450
         TabIndex        =   42
         Top             =   3165
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��ѡ��(&B)"
         Height          =   180
         Index           =   19
         Left            =   3810
         TabIndex        =   27
         Top             =   945
         Width           =   810
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&T)"
         ForeColor       =   &H80000007&
         Height          =   180
         Index           =   18
         Left            =   520
         TabIndex        =   57
         Top             =   5055
         Width           =   630
      End
      Begin VB.Label lblClass 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "���(&C)"
         Height          =   180
         Left            =   180
         TabIndex        =   16
         Top             =   225
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��ʶ����(&P)"
         Height          =   180
         Index           =   17
         Left            =   2550
         TabIndex        =   23
         Top             =   585
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��ʶ����(&I)"
         Height          =   180
         Index           =   16
         Left            =   5235
         TabIndex        =   25
         Top             =   600
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "��ʼ���(&W)"
         Height          =   180
         Index           =   9
         Left            =   6840
         TabIndex        =   32
         Top             =   -90
         Visible         =   0   'False
         Width           =   990
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&S)                   (ƴ��)                   (���)"
         Height          =   180
         Index           =   2
         Left            =   180
         TabIndex        =   29
         Top             =   1335
         Width           =   5130
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "�ϼ�(&V)"
         Height          =   180
         Index           =   3
         Left            =   2895
         TabIndex        =   18
         Tag             =   "����"
         Top             =   210
         Width           =   600
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&U)"
         Height          =   180
         Index           =   0
         Left            =   180
         TabIndex        =   72
         Tag             =   "����"
         Top             =   585
         Width           =   630
      End
      Begin VB.Label lblEdit 
         AutoSize        =   -1  'True
         Caption         =   "����(&N)"
         Height          =   180
         Index           =   1
         Left            =   180
         TabIndex        =   21
         Tag             =   "����"
         Top             =   945
         Width           =   630
      End
   End
   Begin MSComctlLib.TabStrip TabMain 
      Height          =   6990
      Left            =   120
      TabIndex        =   0
      Tag             =   "����"
      Top             =   45
      Width           =   7005
      _ExtentX        =   12356
      _ExtentY        =   12330
      MultiRow        =   -1  'True
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   4
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "������Ϣ"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "�շѼ�Ŀ"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "������Ŀ"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab4 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "ִ�п���"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmChargeItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Enum enum�༭
    text���� = 0
    Text���� = 1
    Text��� = 2
    text���� = 3
    Text˵�� = 4
    Text���� = 5
    Text����ʱ�� = 6
    text����˵�� = 7
    text��� = 8
    text��ʶ���� = 9
    Text��ʶ���� = 10
    text���� = 11
    Text��ѡ�� = 12
    text����޼� = 13
    text����޼� = 14
    text������Ŀ = 15
End Enum

Private mlng���볤�� As Long
Private mlng��λ���� As Long
Private mlng�������� As Long
Private mlng���볤�� As Long

Private mstr��� As String  '������,ֻ��һ����ĸ
Private mstr������� As String    'ԭʼ���ϼ������ֵ
Private mstr���� As String        'ԭʼ�ı��������ֵ
Private mdblҽ�ۼ۸� As Double    'ҽ�۽ӿڱ�׼�۸�
Private mdbl����޼� As Double
Private mdbl����޼� As Double
Private mblnOk As Boolean
Private mlngFind As Long
Private mblnVerifyPris As Boolean   '��˵��۵�Ȩ�� true-��Ȩ�ޣ�false-��Ȩ��
Private mblnVerifyFlow As Boolean   '�����Ƿ�������������̣�true-���ã�false-δ����

'������Ŀ�����޸�
Private mstr����ID As String
Private mstrID As String
Private mint���� As Integer       '�޸�ǰ�����¼����ڵı�����ĳ���

Dim mcol��Ŀ() As Collection    '��������շѼ�Ŀ��ID����������Ŀ��ID��Key�����ͬһ������Ŀʧȥԭ�м�ĿID
Dim mblnNew() As Boolean  '�¼۸�
Dim mblnChanged��Ŀ() As Boolean  '��Ŀ�Ƿ�ı�
Dim mlngĩ�� As Long    'ĩ��
Dim medit��ʽ As EditMode   '0��������1���޸ģ�2�����ۣ�3��ִ�п��ҡ�4��������Ŀ��5����������
Dim mblnChange As Boolean     '�Ƿ�ı���
Dim mstr�б�(1 To 4) As String '����һЩ�б�ֵ3
Dim mblnCancel As Boolean
Dim mblnEditCancel As Boolean   'ȡ������
Dim mstrSel  As String  'ѡ��Ŀ������
Dim mblnShow�շѼ�Ŀ As Boolean '�ж��Ƿ��Ѿ���ʾ���շѼ�Ŀҳ������ҽ��ϵͳ��
Private mstrServerObj As String  '�������

'�Ƿ���  ͨ���ؼ�chk����ж�
'�Ӱ�Ӽ�  ͨ���ؼ�chk����ж�
Private strInputed As String

Private mblnIsSpecialItem As Boolean                '�Ƿ���������Ŀ(������Ŀָ���ǣ���λ�ͻ�������Ŀ�Լ�������"�Զ��Ƽ���Ŀ"�е������Զ�������Ŀ(�����־Ϊ6,7,8));�����Ǵ�λ������Ŀ�Ĵ�����Ŀ
Private mstrCurrentDateFormat As String             '��ǰʹ�õ����ڸ�ʽ

Private mrs���ʷ��� As ADODB.Recordset
Private mrs���� As ADODB.Recordset

Private mstr��ѡִ�п��� As String
Private mblnRefresh As Boolean

'�շѼ�Ŀ�б�
Private Const mcstCol�շ���Ŀ As Integer = 0
Private Const mcstColԭ�� As Integer = 1
Private Const mcstCol�ּ� As Integer = 2
Private Const mcstColȱʡ�۸� As Integer = 3
Private Const mcstCol���������շ��� As Integer = 4
Private Const mcstCol�Ӱ�Ӽ��� As Integer = 5
Private Const mcstCols As Integer = 6
Private mstrPrivs As String
Private mblnNotClick As Boolean
Private mblnCanUpdateAll As Boolean '�Ƿ��������������Ŀ��δ���ü۸�ȼ��������˼۸�ȼ��С�����Ժ����Ȩ��

Private Sub Ini���ʷ���()
    'ȡ�������ʷ��࣬����Ѿ���ȡ�����˳�
    On Error GoTo ErrHandle
    If Not mrs���ʷ��� Is Nothing Then
        mrs���ʷ���.Filter = ""
        If Not mrs���ʷ���.EOF Then
            Exit Sub
        End If
    End If
    
    gstrSQL = "Select ����,������ From �������ʷ���"
    Set mrs���ʷ��� = zlDatabase.OpenSQLRecord(gstrSQL, "ȡ�������ʷ���")
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Load����(ByVal intType As Integer, ByVal str�������� As String)
    'intType:0-ִ�п��ң��������ʣ������ڲ��ˣ���1-���˿��ң��ٴ����ʣ�
    Dim rsData As ADODB.Recordset
    Dim ObjItem As ListItem
    
    On Error GoTo ErrHandle
    
    If intType = 1 Then
        gstrSQL = "select distinct ID,����,����" & _
                " from ���ű� D,��������˵�� T" & _
                " where D.ID=T.����ID and ��������=[1] " & _
                "       and (D.����ʱ�� is null or D.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))" & _
                " order by ����"
    Else
        gstrSQL = "select distinct ID,����,����" & _
                " from ���ű� D,��������˵�� T" & _
                " where D.ID=T.����ID and T.������� in (1,2,3) " & _
                " and (D.����ʱ�� is null or D.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))"
                
        If str�������� <> "��������" Then
            gstrSQL = gstrSQL & " and ��������=[1] "
        End If
                
        gstrSQL = gstrSQL & " order by ����"
    End If
    
    Set rsData = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, str��������)
    
    Me.lvwItems.ListItems.Clear
    
    Me.lvwItems.Checkboxes = True
   
    Do Until rsData.EOF
        
        If Me.lvwItems.Tag = "����" Then
            Set ObjItem = Me.lvwItems.ListItems.Add(, "_" & rsData!ID, rsData!����)
            ObjItem.Icon = "Dept": ObjItem.SmallIcon = "Dept"
            ObjItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = rsData!����
            ObjItem.Checked = False
        
            If InStr(Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, 2) & ",", rsData!ID & ",") > 0 Then
                ObjItem.Checked = True
            End If
        End If
        
        If Me.lvwItems.Tag = "ִ��" Then
            If InStr(mstr��ѡִ�п���, rsData!ID & "," & "[" & rsData!���� & "]" & rsData!����) = 0 Then
                Set ObjItem = Me.lvwItems.ListItems.Add(, "_" & rsData!ID, rsData!����)
                ObjItem.Icon = "Dept": ObjItem.SmallIcon = "Dept"
                ObjItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = rsData!����
                ObjItem.Checked = False
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
Private Sub Ini�շ����ÿ���(ByVal str��ĿID As String)
    Dim rsTmp As ADODB.Recordset
    Dim n As Integer
    
    '�����ٴ���ҽ�����ҺͲ���
'    gstrSQL = " Select Distinct ����||'-'||���� ����,ID From ���ű� " & _
'         " Where ID in (Select ����ID From ��������˵�� Where �������� In ('�ٴ�', '����', '���', '����', '����', '����') And ������� IN(2,3))" & _
'         " And (����ʱ�� Is Null Or ����ʱ��=To_Date('3000-01-01','yyyy-MM-dd')) " & _
'         " Order By ����||'-'||���� "
    'Oracle11g ������ظ�id���ʸ�Ϊ����SQL
    On Error GoTo ErrHandle
    gstrSQL = _
        "Select Distinct a.���� || '-' || a.���� ����, a.Id " & vbNewLine & _
        "From ���ű� A, ��������˵�� B " & vbNewLine & _
        "Where a.Id = b.����id And b.�������� In ('�ٴ�', '����', '���', '����', '����', '����') And b.������� In (2, 3) And " & vbNewLine & _
        "      (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'yyyy-MM-dd')) " & vbNewLine & _
        "Order By ���� || '-' || ���� "

    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "�����ٴ���ҽ�����ҺͲ���")
    
    Lvw����.ListItems.Clear
    With rsTmp
        Do While Not .EOF
            Lvw����.ListItems.Add , "_" & !ID, !����, 1, 1
            .MoveNext
        Loop
    End With
    
    If str��ĿID = "" Then Exit Sub
    
    '�շ����ÿ���
    gstrSQL = "Select ����ID From �շ����ÿ��� Where ��Ŀid = [1]"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "�շ����ÿ���", Val(str��ĿID))
    
    With rsTmp
        If .RecordCount > 0 Then
            optʹ�ÿ���(0).value = True
            Lvw����.Enabled = True
            Do While Not .EOF
                For n = 1 To Lvw����.ListItems.Count
                    If Val(Mid(Lvw����.ListItems(n).Key, 2)) = !����ID Then
                        Lvw����.ListItems(n).Checked = True
                    End If
                Next
                .MoveNext
            Loop
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Function IsValid������Ŀ�ʹ�����ϵ() As Boolean
    '�������ļ۸�������ڶ��������Ŀʱ�Ͳ��������ô�����Ŀ������д�����Ŀ�Ͳ������ö��������Ŀ
    Dim rs As New ADODB.Recordset
    Dim blnIs���ڶ��������Ŀ As Boolean
    Dim blnIs���ڴ�����Ŀ As Boolean
    Dim i As Integer
    
    '�Ƿ��Ѵ��ڶ��������Ŀ,���۸�ȼ������ж�
    On Error GoTo ErrHandle
    If mstrID <> "" Then
        gstrSQL = "Select 1 From �շѼ�Ŀ" & vbNewLine & _
                " Where �շ�ϸĿid=[1] And ִ������ <= SYSDATE AND (��ֹ���� > SYSDATE OR ��ֹ���� IS NULL) " & vbNewLine & _
                " Group By �۸�ȼ�" & vbNewLine & _
                " Having Count(1) > 1"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstrID))
        If Not rs.EOF Then
            blnIs���ڶ��������Ŀ = True
        End If
        rs.Close
    End If
    
    '�༭���Ƿ���ڶ��������Ŀ
    'ֻҪһ���۸�ȼ����ڶ��������Ŀ������Ϊ�Ǵ��ڶ��������Ŀ
    If medit��ʽ = EditNew Or medit��ʽ = EditCopy Or medit��ʽ = EditRaise Then
        For i = msh��Ŀ.LBound To msh��Ŀ.UBound
            If Me.msh��Ŀ(i).Rows > 2 Then
                If Me.msh��Ŀ(i).TextMatrix(2, mcstColԭ��) <> "" Then
                    blnIs���ڶ��������Ŀ = True
                    Exit For
                Else
                    blnIs���ڶ��������Ŀ = False
                End If
            Else
                blnIs���ڶ��������Ŀ = False
            End If
        Next
    End If
            
    '�Ƿ��Ѵ��ڴ���
    If mstrID <> "" Then
        gstrSQL = "select ����id from �շѴ�����Ŀ where ����id=[1] "
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstrID))
        
        If rs.RecordCount > 0 Then
            blnIs���ڴ�����Ŀ = True
        End If
        rs.Close
    End If
    
    '�༭���Ƿ���ڴ���
    If medit��ʽ = EditNew Or medit��ʽ = EditCopy Or medit��ʽ = EditSlave Then
        If Me.msh����.Rows > 1 Then
            If Me.msh����.TextMatrix(1, 1) <> "" Then
                blnIs���ڴ�����Ŀ = True
            Else
                blnIs���ڴ�����Ŀ = False
            End If
        Else
            blnIs���ڴ�����Ŀ = False
        End If
    End If
    
    '������ڶ��������Ŀ�ʹ�����ϵ�Ļ��⣬����ʾ
    If blnIs���ڶ��������Ŀ And blnIs���ڴ�����Ŀ Then
         '���ݱ༭״̬��ʾ��ʾ����
        Select Case medit��ʽ
        Case EditNew, EditCopy
            MsgBox "�������ļ۸������˶��������Ŀ���Ͳ��������ô�����Ŀ����������˴������۸�Ͳ����ж��������Ŀ��", vbExclamation, gstrSysName
        Case EditRaise
            MsgBox "�����Ѿ��д�����Ŀ���������ö��������Ŀ��", vbExclamation, gstrSysName
        Case EditSlave
            MsgBox "����ļ۸��ж��������Ŀ�����������ô�����Ŀ��", vbExclamation, gstrSysName
        End Select
        IsValid������Ŀ�ʹ�����ϵ = False
        Exit Function
    End If
    
    IsValid������Ŀ�ʹ�����ϵ = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function GetSpecialItem(ByVal strID As String) As Boolean
    '�ж��Ƿ��ǵ�����Ҫ���⴦�����Ŀ
    '������Ŀָ���ǣ�1����λ�ͻ�������Ŀ�Լ�������"�Զ��Ƽ���Ŀ"�е������Զ�������Ŀ(�����־Ϊ6,7)
    '                2����ǰ��Ŀ�Ƿ���������λ���߻�������Ŀ�Ĵ�����Ŀ
    '����True����������Ŀ
    '����False������������Ŀ
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    Dim blnTmp As Boolean
    
    On Error GoTo ErrHandle
    strSQL = "Select Id From �շ���ĿĿ¼ " & _
        " Where Id=[1] And (���='J' Or ���='H')" & _
        " Or Id= (Select Distinct �շ�ϸĿid From �Զ��Ƽ���Ŀ Where �����־ In(6,7) And �շ�ϸĿid=[1])"
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption & "-�ж��Ƿ���������Ŀ��", Val(strID))
    
    blnTmp = (rs.RecordCount > 0)
    
    If Not blnTmp And Val(strID) <> 0 Then
        gstrSQL = "Select ID From �շ���ĿĿ¼ Where ID In (Select ����id From �շѴ�����Ŀ Where ����id = [1]) And ��� In ('J', 'H') And Rownum = 1"
        Set rs = zlDatabase.OpenSQLRecord(gstrSQL, "������Ŀ", Val(strID))
        
        blnTmp = (rs.RecordCount > 0)
    End If
    GetSpecialItem = blnTmp
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function


Private Sub load���ʷ���(ByVal intType As Integer)
    'intType:0-ִ�п��ң��������ʣ������ڲ��ˣ���1-���˿��ң��ٴ����ʣ�
    
    mblnRefresh = True
    
    With cboProperty
        .Clear
        
        If mrs���ʷ��� Is Nothing Then Exit Sub
        
        If intType = 0 Then
            mrs���ʷ���.Filter = "������=1 Or ������=2 Or ������=3"
        Else
            mrs���ʷ���.Filter = "����='�ٴ�'"
        End If
        
        If mrs���ʷ���.RecordCount = 0 Then Exit Sub
        
        If intType = 0 Then
            .AddItem "��������"
            
            Do While Not mrs���ʷ���.EOF
                .AddItem mrs���ʷ���!����
                
                mrs���ʷ���.MoveNext
            Loop
        Else
            .AddItem "�ٴ�"
        End If
        
        If .ListCount > 0 Then .ListIndex = 0
    End With
    
    DoEvents
    
    mblnRefresh = False
End Sub

Private Function TabExist(ByVal strTabName As String) As Boolean
    Dim i As Integer
    
    For i = 1 To TabMain.Tabs.Count
        If TabMain.Tabs(i).Key = "_" & strTabName Then
            TabExist = True
            Exit Function
        End If
    Next
End Function

Public Function �༭��Ŀ(ByVal strPrivs As String, ByVal blnCanUpdateAll As Boolean, _
    ByVal str����ID As String, Optional strID As String = "", _
    Optional ByVal lngĩ����Ŀ As Long = 1, Optional ByVal edit��ʽ As EditMode = EditNew, _
    Optional ByVal PriceImp As Boolean = False) As Boolean
    '����:��������õ��շ�ϸĿ�����ڽ���ͨѶ�ĳ���
    '����:
    '     strPrivs Ȩ�޴�
    '     str����ID   �շ���Ŀ�ķ���ID   'Ϊ���ֱ�ʾID������Ϊ�����
    '     strID           ���շ���Ŀ�ĵ�ID
    '     blnĩ����Ŀ     ���շ���Ŀ�Ƿ�ĩ��
    '     edit��ʽ  ȡֵΪ��0��������1���޸ģ�2�����ۣ�3��ִ�п��ҡ�4��������Ŀ��5����������
    '     PriceImp  =True��ʾʹ��ҽ�� =False����ʹ��ҽ�� Ĭ��Ϊ��ʹ��ҽ��
    '����ֵ:�༭�ɹ�����True,����ΪFalse
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    
    On Error GoTo ErrHandle
    mstrPrivs = strPrivs
    mblnCanUpdateAll = blnCanUpdateAll
    mblnShow�շѼ�Ŀ = False
    
    mblnVerifyPris = IIF(InStr(1, ";" & strPrivs & ";", ";�շѼ�Ŀ�������;") > 0, True, False)
    mblnVerifyFlow = IIF(Val(zlDatabase.GetPara("������Ҫ���", glngSys, 1009, 0)) = 0, False, True)
        
    '��ʹ��ҽ��ʱ���Σ���ʶ��������룩
'    If PriceImp = False Then
'        Me.txtEdit(9).Enabled = False
'        Me.txtEdit(9).BackColor = &H80000004
'        Me.txtEdit(10).Enabled = False
'        Me.txtEdit(10).BackColor = &H80000004
'    End If
    
    medit��ʽ = edit��ʽ
    mstrID = strID
    Call GetPriceGrade(gstrҩƷ�۸�ȼ�, gstr���ļ۸�ȼ�, gstr��ͨ�۸�ȼ�)
    '�����շ���Ŀ��ID�������
    If medit��ʽ <> EditNew Then
        If IsNumeric(mstrID) Then
            strSQL = "select ��� from �շ���ĿĿ¼ where ���<>'5' and ���<>'6' and ���<>'7' and id=[1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mstrID))
            
            If rsTmp.RecordCount < 1 Then
                MsgBox "�����ڵ���ĿID��", vbExclamation, gstrSysName
                Exit Function
            End If
            mstr��� = rsTmp!���
        Else
            MsgBox "��Ч����ĿID��", vbExclamation, gstrSysName
            Exit Function
        End If
        '�ж��Ƿ���������Ŀ
        mblnIsSpecialItem = GetSpecialItem(strID)
    Else
        '����������ֻ�д����洫�����
        mstr��� = Mid(str����ID, 2, 1)
        strSQL = "select 1 from �շ���Ŀ��� where ����<>'4' And ����<>'5' and ����<>'6' and ����<>'7' and Upper(����)=[1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UCase(Trim(mstr���)))
        
        If rsTmp.RecordCount < 1 Then
            mstr��� = ""
        End If
    End If
    If edit��ʽ <> EditNew Then
        '�жϸ��շ���Ŀ�Ƿ����,��������ĿID�������ID
        strSQL = "select ����ID from �շ���ĿĿ¼ where ���<>'5' and ���<>'6' and ���<>'7' and  id=[1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mstrID))
        
        If rsTmp.RecordCount > 0 Then
            mstr����ID = Nvl(rsTmp!����id)
        Else
            MsgBox "ѡ���շ���Ŀ�����ڣ�", vbInformation, gstrSysName
            Exit Function
        End If
    ElseIf edit��ʽ = EditNew Then
        If Len(str����ID) > 2 Then
            If IsNumeric(Mid(str����ID, 3)) Then
                '�жϸ��շ���Ŀ�Ƿ����,��������ĿID�������ID
                strSQL = "select ID from �շѷ���Ŀ¼ where id=[1] "
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(Mid(str����ID, 3)))
                
                If rsTmp.RecordCount < 1 Then
                    mstr����ID = ""
                Else
                    mstr����ID = CStr(rsTmp!ID)
                End If
            Else
                mstr����ID = ""
            End If
        Else
            mstr����ID = ""
        End If
    End If
    
    mstr���� = ""
    
    If Trim(mstr����ID) = "0" Then
        mstr����ID = ""
    End If
    frmChargeItem.Caption = "�շ���Ŀ����"
    Call GetDefineSize
    If edit��ʽ <> EditNew And edit��ʽ <> EditCopy Then chk����.Visible = False
    msh����.Cols = 3
    msh��Ŀ(0).Cols = mcstCols
    TabMain.Tabs.Clear
    Select Case edit��ʽ
    Case EditNew, EditCopy
        TabMain.Tabs.Add , "_������Ϣ", "������Ϣ"
        If init���� = False Then
            Exit Function
        End If
        TabMain.Tabs.Add , "_�շѼ�Ŀ", "�շѼ�Ŀ"
        TabMain.Tabs.Add , "_ִ�п���", "ִ�п���"
        If InStr(strPrivs, "��Ŀ�������") > 0 Then
            TabMain.Tabs.Add , "_������Ŀ", "������Ŀ"
            init����
            Call Ini�շ����ÿ���(mstrID)
        End If
        If init��Ŀ = False Then Exit Function
        initִ��
        chk����.Visible = True
        '�����Ǹ�������������Ҫ���һ������
        If medit��ʽ = EditCopy Then
            ClearContext False
        End If
    Case EditModify
        TabMain.Tabs.Add , "_������Ϣ", "������Ϣ"
        init����
    Case EditRaise
        TabMain.Tabs.Add , "_�շѼ�Ŀ", "�շѼ�Ŀ"
        If init��Ŀ = False Then Exit Function
    Case EditDept
        TabMain.Tabs.Add , "_ִ�п���", "ִ�п���"
        initִ��
    Case EditSlave
        TabMain.Tabs.Add , "_������Ŀ", "������Ŀ"
        init����
        Call Ini�շ����ÿ���(mstrID)
    End Select
    Call tabMain_Click
    mblnChange = False
    frmChargeItem.Show vbModal
    �༭��Ŀ = mblnOk
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
    mblnChange = False
End Function

Private Function IsValidִ��() As Boolean
    '����Ƿ���Ҫ����ִ�п���
    On Error GoTo ErrHandle
    Dim i As Long
    Dim j As Long
    Dim blnEmpt As Boolean
    Dim strTemp As String

    If opt����(4).value = True Then 'ָ������
        Select Case mstrServerObj
            Case "1"
                If txt����ִ��.Tag = "" Then
                    blnEmpt = True
                    strTemp = ",�շ���Ŀ�������Ϊ�����ʱ��Ӧ����������һ��ִ�п��ң�"
                End If
            Case "2"
                If txtסԺִ��.Tag = "" Then
                    blnEmpt = True
                    strTemp = ",�շ���Ŀ�������Ϊ��סԺ��ʱ��Ӧ����������һ��ִ�п��ң�"
                End If
            Case "3"
                If txt����ִ��.Tag = "" And txtסԺִ��.Tag = "" Then
                    blnEmpt = True
                    strTemp = ",�շ���Ŀ�������Ϊ�������סԺ��ʱ��Ӧ����������һ��ִ�п��ң�"
                End If
        End Select
        
        If blnEmpt = True Then
            If msf����ִ��.TextMatrix(1, 0) <> "" And msf����ִ��.TextMatrix(1, 2) <> "" Then
                IsValidִ�� = True
            Else
                MsgBox "ָ������" & strTemp, vbInformation, gstrSysName
                If medit��ʽ = EditNew Or medit��ʽ = EditCopy Then '��������������
                    TabMain.Tabs(3).Selected = True
                End If
                Select Case mstrServerObj
                    Case "1"
                        txt����ִ��.SetFocus
                    Case "2"
                        txtסԺִ��.SetFocus
                    Case "3"
                        txt����ִ��.SetFocus
                End Select
                IsValidִ�� = False
            End If
        Else
            IsValidִ�� = True
        End If
    Else
        IsValidִ�� = True
    End If
'    If Trim(mstr���) <> "1" And Trim(mstr���) <> "H" And Trim(mstr���) <> "J" Then
'        If sstAdmin.Enabled = True Then
'            txtOutIn.Visible = False
'            cmdSel��������(0).Visible = False
'            cmdSelִ�п���(0).Visible = False
'            cmdSel��������(1).Visible = False
'            cmdSelִ�п���(1).Visible = False
'ReOut:
'            For i = 2 To msfOut.Rows - 1
'                If Trim(msfOut.TextMatrix(i, 0)) = "" And Trim(msfOut.TextMatrix(i, 2)) = "" Then
'                    msfOut.RemoveItem i
'                    GoTo ReOut
'                End If
'            Next
'ReIn:
'            For i = 2 To msfIn.Rows - 1
'                If Trim(msfIn.TextMatrix(i, 0)) = "" And Trim(msfIn.TextMatrix(i, 2)) = "" Then
'                    msfIn.RemoveItem i
'                    GoTo ReIn
'                End If
'            Next
'            For i = 0 To msfOut.Rows - 1
'                If Trim(msfOut.TextMatrix(i, 0)) = "" And Trim(msfOut.TextMatrix(i, 2)) <> "" Then
'                    msfOut.Row = i: msfOut.Col = 0
'                    sstAdmin.Tab = 0
''                    msfOut_RowColChange
'                    MsgBox "�������Ҳ���Ϊ�գ�", vbExclamation, gstrSysName
'                    If msfOut.Enabled And msfOut.Visible Then
'                        msfOut.SetFocus
'                        txtOutIn.Visible = True
'                    End If
'                    Exit Function
'                End If
'                If Trim(msfOut.TextMatrix(i, 1)) = "" And Trim(msfOut.TextMatrix(i, 2)) <> "" Then
'                    msfOut.Row = i: msfOut.Col = 1
'                    sstAdmin.Tab = 0
''                    msfOut_RowColChange
'                    MsgBox "ִ�п��Ҳ���Ϊ�գ�", vbExclamation, gstrSysName
'                    If msfOut.Enabled And msfOut.Visible Then
'                        msfOut.SetFocus
'                        txtOutIn.Visible = True
'                    End If
'                    Exit Function
'                End If
'                For j = 1 To msfOut.Rows - 1
'                    If msfOut.TextMatrix(i, 0) = msfOut.TextMatrix(j, 0) And i <> j Then
'                        msfOut.Row = j: msfOut.Col = 0
'                        sstAdmin.Tab = 0
''                        msfOut_RowColChange
'                        MsgBox "�������� " & msfOut.Text & " �����ظ���", vbExclamation, gstrSysName
'                        If msfOut.Enabled And msfOut.Visible Then
'                            msfOut.SetFocus
'                            txtOutIn.Visible = True
'                        End If
'                        Exit Function
'                    End If
'                Next
'            Next
'            For i = 0 To msfIn.Rows - 1
'                If Trim(msfIn.TextMatrix(i, 0)) = "" And Trim(msfIn.TextMatrix(i, 2)) <> "" Then
'                    msfIn.Row = i: msfIn.Col = 0
'                    sstAdmin.Tab = 1
''                    msfIn_RowColChange
'                    MsgBox "�������Ҳ���Ϊ�գ�", vbExclamation, gstrSysName
'                    If msfIn.Enabled And msfIn.Visible Then
'                        msfIn.SetFocus
'                        txtOutIn.Visible = True
'                    End If
'                    Exit Function
'                End If
'                If Trim(msfIn.TextMatrix(i, 1)) = "" And Trim(msfIn.TextMatrix(i, 2)) <> "" Then
'                    msfIn.Row = i: msfIn.Col = 1
'                    sstAdmin.Tab = 1
''                    msfIn_RowColChange
'                    MsgBox "ִ�п��Ҳ���Ϊ�գ�", vbExclamation, gstrSysName
'                    If msfIn.Enabled And msfIn.Visible Then
'                        msfIn.SetFocus
'                        txtOutIn.Visible = True
'                    End If
'                    Exit Function
'                End If
'                For j = 1 To msfIn.Rows - 1
'                    If Trim(msfIn.TextMatrix(i, 0)) = Trim(msfIn.TextMatrix(j, 0)) And i <> j Then
'                        msfIn.Row = j: msfIn.Col = 0
'                        sstAdmin.Tab = 1
''                        msfIn_RowColChange
'                        MsgBox "�������� " & msfIn.Text & " �����ظ���", vbExclamation, gstrSysName
'                        If msfIn.Enabled And msfIn.Visible Then
'                            msfIn.SetFocus
'                            txtOutIn.Visible = True
'                        End If
'                        Exit Function
'                    End If
'                Next
'            Next
'        End If
'    End If
'    IsValidִ�� = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function IsValid() As Boolean
    '���Ƿ���
    On Error GoTo ErrHandle
    Dim i As Long
    Select Case medit��ʽ
    Case EditNew, EditCopy
        If IsValid���� = False Then Exit Function
        If IsValidִ�� = False Then Exit Function
        
        If IsValid������Ŀ�ʹ�����ϵ = False Then Exit Function
        If IsValid��Ŀ = False Then Exit Function
        If InStr(frmChargeManage.mstrPrivs, "��Ŀ�������") > 0 Then
            If IsValid���� = False Then Exit Function
        End If
    Case EditModify
        If IsValid���� = False Then Exit Function
        '�����ʾ�˵��۽��棬��Ҫ����Ŀ
        If mblnShow�շѼ�Ŀ Then
            If IsValid��Ŀ = False Then Exit Function
        End If
    Case EditRaise
        If IsValid������Ŀ�ʹ�����ϵ = False Then Exit Function
        If IsValid��Ŀ = False Then Exit Function
    Case EditDept
        If IsValidִ�� = False Then Exit Function
        If optApply(0).value = False Then
            For i = 1 To 3
                If optApply(i).value = True Then
                    If MsgBox("��ѡ���ˡ�" & Mid(optApply(i).Caption, 1, InStr(optApply(i).Caption, "(") - 1) & "��Ӧ��ģʽ��" & vbCrLf & _
                        "���Ӱ�쵽������Ŀ���Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                        Exit Function
                    End If
                End If
            Next
        End If
    Case EditSlave
        If IsValid������Ŀ�ʹ�����ϵ = False Then Exit Function
        If IsValid���� = False Then Exit Function
    End Select
    IsValid = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function SaveϸĿ() As Boolean
    '���ݵ�ǰģʽ����ϸĿ
    On Error GoTo errSave
    gcnOracle.BeginTrans
    Select Case medit��ʽ
    Case EditNew, EditCopy
        If Save����() = False Then Exit Function
        If Save��Ŀ = False Then Exit Function
        If Saveִ�� = False Then Exit Function
        If InStr(frmChargeManage.mstrPrivs, "��Ŀ�������") > 0 Then
            If Save���� = False Then Exit Function
        End If
    Case EditModify
        If Save���� = False Then Exit Function
        '��������˵��۽��棬�����Ҫ���±����Ŀ
        If mblnShow�շѼ�Ŀ Then
            If Save��Ŀ = False Then Exit Function
        End If
    Case EditRaise
        If Save��Ŀ = False Then Exit Function
    Case EditDept
        If Saveִ�� = False Then Exit Function
    Case EditSlave
        If Save���� = False Then Exit Function
    End Select
    gcnOracle.CommitTrans
    SaveϸĿ = True
    Exit Function
errSave:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Save����() As Boolean
    Dim str���� As String
    Dim i  As Integer
    Dim intDept As Integer
    Dim int������Ŀ As Integer
    Dim strվ�� As String
    
    On Error GoTo ErrHandle
    With mshAlias
        If Trim(txtEdit(text����).Text) <> "" Then
            str���� = "1''" & txtEdit(Text����).Text & "''1''" & txtEdit(text����).Text & "''"
        End If
        If Trim(txtEdit(text���).Text) <> "" Then
            str���� = str���� & "1''" & txtEdit(Text����).Text & "''2''" & txtEdit(text���).Text & "''"
        End If
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, 0)) <> "" Then
                If Trim(.TextMatrix(i, 1)) <> "" Then
                    str���� = str���� & "9''" & Trim(.TextMatrix(i, 0)) & "''1''" & Trim(.TextMatrix(i, 1)) & "''"
                End If
                If Trim(.TextMatrix(i, 2)) <> "" Then
                    str���� = str���� & "9''" & Trim(.TextMatrix(i, 0)) & "''2''" & Trim(.TextMatrix(i, 2)) & "''"
                End If
            End If
        Next
    End With
    If mstr��� = "1" Then
        If chk����.value = 0 Then
            int������Ŀ = 1
        Else
            int������Ŀ = 2
        End If
    ElseIf mstr��� = "H" Then
        int������Ŀ = cmb����.ListIndex + 3
    Else
        int������Ŀ = 0
    End If
    strվ�� = zlStr.NeedCode(cmbStationNo.Text, "-")
    
    If medit��ʽ <> EditModify Then
        '����
        For i = 0 To 6
            If opt����(i).value = True Then
                intDept = i
                Exit For
            End If
        Next
        mstrID = sys.NextId("�շ���ĿĿ¼")
        gstrSQL = "zl_�շ�ϸĿ_insert(" & int������Ŀ & "," & mstrID & ",'" & mstr��� & "','" & txtEdit(text����).Text & "','" & txtEdit(text��ʶ����).Text & "','" & txtEdit(Text��ʶ����).Text & "','" & txtEdit(Text��ѡ��).Text & "','" & txtEdit(Text����).Text & _
            "'," & IIF(mstr����ID = "", "Null", mstr����ID) & ",'" & Replace(txtEdit(Text���).Text, "'", "''") & "','" & Replace(txtEdit(Text˵��).Text, "'", "''") & _
            "','" & cmb���㵥λ.Text & "'," & GetTextFromCombo(cmb��������, True) & "," & chk���ηѱ�.value & "," & chk���.value & "," & chk�Ӱ�Ӽ�.value & "," & intDept & "," & _
            Left(cmb�������.Text, 1) & "," & chkժҪ.value & "," & txtEdit(text����޼�).Text & "," & txtEdit(text����޼�).Text & ",'" & str���� & "'," & Val(Me.txt¼������.Text) & "," & cbo¼��������Χ.ListIndex & "," & cmb����ȷ��.ListIndex & "," & chk����ȷ�Ϸ�Χ.value & "," & chk�Զ�����.value & _
            ",'" & strվ�� & "','" & txtEdit(text������Ŀ).Text & "'," & cmb��Ŀ����.ListIndex & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        'Ϊ���ϸ��²���
        If mstr��� = "M" Then
            gstrSQL = "ZL_�շ�ϸĿ_���ϲ���(" & mstrID & ",'" & Replace(Me.txtEdit(text����).Text, "'", "''") & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        End If
    Else
        '�޸�
        gstrSQL = "zl_�շ�ϸĿ_update(" & int������Ŀ & "," & mstrID & ",'" & mstr��� & "','" & txtEdit(text����).Text & "','" & txtEdit(text��ʶ����).Text & "','" & txtEdit(Text��ʶ����).Text & "','" & txtEdit(Text��ѡ��).Text & "','" & txtEdit(Text����).Text & _
            "'," & IIF(mstr����ID = "", "Null", mstr����ID) & _
            ",'" & Replace(txtEdit(Text���).Text, "'", "''") & "','" & Replace(txtEdit(Text˵��).Text, "'", "''") & "','" & cmb���㵥λ.Text & "'," & GetTextFromCombo(cmb��������, True) & _
            "," & chk���ηѱ�.value & "," & chk���.value & "," & chk�Ӱ�Ӽ�.value & "," & _
            Left(cmb�������.Text, 1) & "," & chkժҪ.value & "," & txtEdit(text����޼�).Text & "," & txtEdit(text����޼�).Text & ",'" & str���� & "'," & Val(Me.txt¼������.Text) & "," & cbo¼��������Χ.ListIndex & "," & cmb����ȷ��.ListIndex & "," & chk����ȷ�Ϸ�Χ.value & "," & chk�Զ�����.value & _
            ",'" & strվ�� & "','" & txtEdit(text������Ŀ).Text & "'," & cmb��Ŀ����.ListIndex & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        'Ϊ���ϸ��²���
        If mstr��� = "M" Then
            gstrSQL = "ZL_�շ�ϸĿ_���ϲ���(" & mstrID & ",'" & Replace(Me.txtEdit(text����).Text, "'", "''") & "')"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        End If
    End If
    Save���� = True
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Save��Ŀ() As Boolean
    Dim intRow As Integer
    Dim i As Integer, k As Integer
    Dim lng����ID As Long
    Dim lng��ĿID As Long
    Dim dateExec As Date
    Dim str��ʼʱ�� As String
    Dim str��ֹʱ�� As String
    Dim strNo As String
    Dim str�������� As String
    Dim dtTemp As Date
    
    On Error GoTo ErrHandle
    dtTemp = sys.Currentdate
    For k = 0 To tbPriceGrade.ItemCount - 1
        If tbPriceGrade(k).Visible _
            And Not (mblnChanged��Ŀ(k) = False And (medit��ʽ = EditModify Or medit��ʽ = EditRaise)) Then    '���صĺ�δ�ı�Ĳ�����
            str��ʼʱ�� = Format(IIF(Me.chkNow(k).value = 1, dtTemp, dtpBegin(k).value), mstrCurrentDateFormat)
            str��ֹʱ�� = Format(DateAdd("s", -1, str��ʼʱ��), "yyyy-MM-dd hh:mm:ss")
            str�������� = Format(dtTemp, "yyyy-MM-dd hh:mm:ss")
            intRow = 0
        
            '1�������������ڱ��в������ݣ����򲻲�������
            '2����Ȩ������ˣ�û��Ȩ�������
            With msh��Ŀ(k)
                If Trim(.TextMatrix(1, mcstCol�շ���Ŀ)) <> "" Then
                    lng����ID = sys.NextId("�շѼ�Ŀ")
                    strNo = sys.GetNextNo(9)
            
                    If medit��ʽ = EditRaise Or (medit��ʽ = EditModify And mblnShow�շѼ�Ŀ) Then
                        If mblnVerifyFlow = False And mblnVerifyPris = False Then
                            gcnOracle.RollbackTrans
                            MsgBox "��û�����õ������ģʽ�£�����Ա����Ҫ�����Ȩ�޲��ܵ��ۣ�", vbInformation, gstrSysName
                            Exit Function
                        End If
                        '����
                        If mblnVerifyFlow = True Then
                            For i = 1 To .Rows - 1
                                If .RowData(i) > 0 Then
                                    If intRow = 0 Then
                                        lng��ĿID = lng����ID
                                    Else
                                        lng��ĿID = sys.NextId("�շѼ�Ŀ")
                                    End If
                                    'Zl_�շѵ��ۼ�¼_Insert(
                                    gstrSQL = "Zl_�շѵ��ۼ�¼_Insert("
                                    '  Id_In         In �շѼ�Ŀ.Id%Type,
                                    gstrSQL = gstrSQL & "" & lng��ĿID & ","
                                    '  ԭ��id_In     In �շѼ�Ŀ.ԭ��id%Type := Null,
                                    gstrSQL = gstrSQL & "" & IIF(mcol��Ŀ(k)("C" & .RowData(i)) = 0, "null", mcol��Ŀ(k)("C" & .RowData(i))) & ","
                                    '  �շ�ϸĿid_In In �շѼ�Ŀ.�շ�ϸĿid%Type := Null,
                                    gstrSQL = gstrSQL & "" & mstrID & ","
                                    '  ������Ŀid_In In �շѼ�Ŀ.������Ŀid%Type := Null,
                                    gstrSQL = gstrSQL & "" & .RowData(i) & ","
                                    '  ԭ��_In       In �շѼ�Ŀ.ԭ��%Type := Null,
                                    gstrSQL = gstrSQL & "" & Val(.TextMatrix(i, mcstColԭ��)) & ","
                                    '  �ּ�_In       In �շѼ�Ŀ.�ּ�%Type := Null,
                                    gstrSQL = gstrSQL & "" & Val(.TextMatrix(i, mcstCol�ּ�)) & ","
                                    '  ȱʡ�۸�_In   In �շѼ�Ŀ.ȱʡ�۸�%Type := Null,
                                    gstrSQL = gstrSQL & "" & IIF(Val(.TextMatrix(1, mcstColȱʡ�۸�)) = 0, "Null", Val(.TextMatrix(1, mcstColȱʡ�۸�))) & ","
                                    '  �����շ���_In In �շѼ�Ŀ.�����շ���%Type := Null,
                                    gstrSQL = gstrSQL & "" & IIF(.TextMatrix(i, mcstCol���������շ���) = "", 0, .TextMatrix(i, mcstCol���������շ���)) & ","
                                    '  �Ӱ�Ӽ���_In In �շѼ�Ŀ.�Ӱ�Ӽ���%Type := Null,
                                    gstrSQL = gstrSQL & "" & IIF(.TextMatrix(i, mcstCol�Ӱ�Ӽ���) = "", 0, .TextMatrix(i, mcstCol�Ӱ�Ӽ���)) & ","
                                    '  ����˵��_In   In �շѼ�Ŀ.����˵��%Type := Null,
                                    gstrSQL = gstrSQL & "'" & txt����˵��(k).Text & "',"
                                    '  ����id_In     In �շѼ�Ŀ.����id%Type := Null,
                                    gstrSQL = gstrSQL & "" & lng����ID & ","
                                    '  ������_In     In �շѵ��ۼ�¼.������%Type := Null,
                                    gstrSQL = gstrSQL & "'" & gstrUserName & "',"
                                    '  ��������_In   In �շѵ��ۼ�¼.��������%Type := Null,
                                    gstrSQL = gstrSQL & "" & "to_date('" & str�������� & "','YYYY-MM-DD HH24:MI:SS')" & ","
                                    '  ִ������_In   In �շѼ�Ŀ.ִ������%Type := Null,
                                    gstrSQL = gstrSQL & "" & "to_date('" & str��ʼʱ�� & "','YYYY-MM-DD HH24:MI:SS')" & ","
                                    '  �䶯ԭ��_In   In �շѼ�Ŀ.�䶯ԭ��%Type := 1,
                                    gstrSQL = gstrSQL & "" & "1" & ","
                                    '  No_In         In �շѼ�Ŀ.No%Type := Null,
                                    gstrSQL = gstrSQL & "'" & strNo & "',"
                                    '  ���_In       In �շѼ�Ŀ.���%Type := 1,
                                    gstrSQL = gstrSQL & "" & intRow + 1 & ","
                                    '  �۸�ȼ�_In   In �շѼ�Ŀ.�۸�ȼ�%Type := Null
                                    gstrSQL = gstrSQL & "" & IIF(tbPriceGrade(k).Caption = "ȱʡ", "NULL", "'" & tbPriceGrade(k).Caption & "'") & ")"
                                    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            
                                    If mblnVerifyPris = True Then
                                        'Zl_�շѵ��ۼ�¼_Verify(
                                        gstrSQL = "Zl_�շѵ��ۼ�¼_Verify("
                                        '  Id_In       In �շѵ��ۼ�¼.Id%Type,
                                        gstrSQL = gstrSQL & "" & lng��ĿID & ","
                                        '  ��˱�־_In In �շѵ��ۼ�¼.��˱�־%Type := 1,
                                        gstrSQL = gstrSQL & "" & "1" & ","
                                        '  �����_In   In �շѵ��ۼ�¼.�����%Type := Null,
                                        gstrSQL = gstrSQL & "'" & gstrUserName & "',"
                                        '  �������_In In �շѵ��ۼ�¼.�������%Type := Null,
                                        gstrSQL = gstrSQL & "" & "to_date('" & str�������� & "','YYYY-MM-DD HH24:MI:SS')" & ")"
                                        '  ˵��_In     In �շѵ��ۼ�¼.˵��%Type := Null
                                        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                                    End If
                                    intRow = intRow + 1
                                End If
                            Next
                        Else
                            '�����Ȩ�ޣ�δ�������ģʽ��ֱ�����շѵ��ۼ�¼���в����Ѿ���Ч���������
                            If mblnVerifyPris = True Then
                                '����������ʾʱ����Ƕ�����Ŀ��Ҫ�����շѼ�Ŀ��������ǰ��Ŀ��ͣ��ʱ�䣩
                                If (medit��ʽ = EditRaise Or (medit��ʽ = EditModify And mblnShow�շѼ�Ŀ)) _
                                    And chk���.value = 0 Then   '����
                                    '��д��ǰ��Ŀ����ֹ����
                                    'ZL_�շѼ�Ŀ_STOP(
                                    gstrSQL = "zl_�շѼ�Ŀ_stop("
                                    '  �շ�ϸĿid_In In �շѼ�Ŀ.�շ�ϸĿid%Type,
                                    gstrSQL = gstrSQL & mstrID & ","
                                    '  ��ֹ����_In   In �շѼ�Ŀ.��ֹ����%Type := Null,
                                    gstrSQL = gstrSQL & "To_Date('" & str��ֹʱ�� & "','YYYY-MM-DD HH24:MI:SS')" & ","
                                    '  �۸�ȼ�_In   In �շѼ�Ŀ.�۸�ȼ�%Type := Null
                                    gstrSQL = gstrSQL & IIF(tbPriceGrade(k).Caption = "ȱʡ", "NULL", "'" & tbPriceGrade(k).Caption & "'") & ")"
                                    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                                End If
            
                                If chk���.value = 0 Or medit��ʽ <> EditRaise Or mblnNew(k) Then
                                    If chk���.value = 1 And medit��ʽ = EditModify And mblnShow�շѼ�Ŀ Then
                                        'ZL_�շѼ�Ŀ_UPDATE(
                                        gstrSQL = "ZL_�շѼ�Ŀ_UPDATE("
                                        '  �շ�ϸĿid_In In �շѼ�Ŀ.�շ�ϸĿid%Type := Null,
                                        gstrSQL = gstrSQL & "" & mstrID & ","
                                        '  ������Ŀid_In In �շѼ�Ŀ.������Ŀid%Type := Null,
                                        gstrSQL = gstrSQL & "" & .RowData(1) & ","
                                        '  ԭ��_In       In �շѼ�Ŀ.ԭ��%Type := Null,
                                        gstrSQL = gstrSQL & "" & Val(.TextMatrix(1, mcstColԭ��)) & ","
                                        '  �ּ�_In       In �շѼ�Ŀ.�ּ�%Type := Null,
                                        gstrSQL = gstrSQL & "" & Val(.TextMatrix(1, mcstCol�ּ�)) & ","
                                        '  �����շ���_In In �շѼ�Ŀ.�����շ���%Type := Null,
                                        gstrSQL = gstrSQL & "" & IIF(.TextMatrix(1, mcstCol���������շ���) = "", 0, .TextMatrix(1, mcstCol���������շ���)) & ","
                                        '  �Ӱ�Ӽ���_In In �շѼ�Ŀ.�Ӱ�Ӽ���%Type := Null,
                                        gstrSQL = gstrSQL & "" & IIF(.TextMatrix(1, mcstCol�Ӱ�Ӽ���) = "", 0, .TextMatrix(1, mcstCol�Ӱ�Ӽ���)) & ","
                                        '  ����˵��_In   In �շѼ�Ŀ.����˵��%Type := Null,
                                        gstrSQL = gstrSQL & "" & "'" & txt����˵��(k).Text & "',"
                                        '  ����id_In     In �շѼ�Ŀ.����id%Type := Null,
                                        gstrSQL = gstrSQL & "" & lng����ID & ","
                                        '  ������_In     In �շѼ�Ŀ.������%Type := Null,
                                        gstrSQL = gstrSQL & "'" & gstrUserName & "',"
                                        '  ȱʡ�۸�_In   In �շѼ�Ŀ.ȱʡ�۸�%Type := Null,
                                        gstrSQL = gstrSQL & "" & IIF(Val(.TextMatrix(1, mcstColȱʡ�۸�)) = 0, "Null", Val(.TextMatrix(1, mcstColȱʡ�۸�))) & ","
                                        '  �۸�ȼ�_In   In �շѼ�Ŀ.�۸�ȼ�%Type := Null
                                        gstrSQL = gstrSQL & IIF(tbPriceGrade(k).Caption = "ȱʡ", "NULL", "'" & tbPriceGrade(k).Caption & "'") & ")"
                                        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                                        Exit Function
                                    End If
                                    For i = 1 To .Rows - 1
                                        If .RowData(i) > 0 Then
                                            If intRow = 0 Then
                                                lng��ĿID = lng����ID
                                            Else
                                                lng��ĿID = sys.NextId("�շѼ�Ŀ")
                                            End If
                                            'Zl_�շѼ�Ŀ_Insert(
                                            gstrSQL = "Zl_�շѼ�Ŀ_Insert("
                                            '  Id_In         In �շѼ�Ŀ.Id%Type,
                                            gstrSQL = gstrSQL & "" & lng��ĿID & ","
                                            '  ԭ��id_In     In �շѼ�Ŀ.ԭ��id%Type := Null,
                                            gstrSQL = gstrSQL & "" & IIF(mcol��Ŀ(k)("C" & .RowData(i)) = 0, "null", mcol��Ŀ(k)("C" & .RowData(i))) & ","
                                            '  �շ�ϸĿid_In In �շѼ�Ŀ.�շ�ϸĿid%Type := Null,
                                            gstrSQL = gstrSQL & "" & mstrID & ","
                                            '  ������Ŀid_In In �շѼ�Ŀ.������Ŀid%Type := Null,
                                            gstrSQL = gstrSQL & "" & .RowData(i) & ","
                                            '  ԭ��_In       In �շѼ�Ŀ.ԭ��%Type := Null,
                                            gstrSQL = gstrSQL & "" & Val(.TextMatrix(i, mcstColԭ��)) & ","
                                            '  �ּ�_In       In �շѼ�Ŀ.�ּ�%Type := Null,
                                            gstrSQL = gstrSQL & "" & Val(.TextMatrix(i, mcstCol�ּ�)) & ","
                                            '  �����շ���_In In �շѼ�Ŀ.�����շ���%Type := Null,
                                            gstrSQL = gstrSQL & "" & IIF(.TextMatrix(i, mcstCol���������շ���) = "", 0, .TextMatrix(i, mcstCol���������շ���)) & ","
                                            '  �Ӱ�Ӽ���_In In �շѼ�Ŀ.�Ӱ�Ӽ���%Type := Null,
                                            gstrSQL = gstrSQL & "" & IIF(.TextMatrix(i, mcstCol�Ӱ�Ӽ���) = "", 0, .TextMatrix(i, mcstCol�Ӱ�Ӽ���)) & ","
                                            '  ����˵��_In   In �շѼ�Ŀ.����˵��%Type := Null,
                                            gstrSQL = gstrSQL & "'" & txt����˵��(k).Text & "',"
                                            '  ����id_In     In �շѼ�Ŀ.����id%Type := Null,
                                            gstrSQL = gstrSQL & "" & lng����ID & ","
                                            '  ������_In     In �շѼ�Ŀ.������%Type := Null,
                                            gstrSQL = gstrSQL & "'" & gstrUserName & "',"
                                            '  ִ������_In   In �շѼ�Ŀ.ִ������%Type := Null,
                                            gstrSQL = gstrSQL & "" & "to_date('" & str��ʼʱ�� & "','YYYY-MM-DD HH24:MI:SS')" & ","
                                            '  �䶯ԭ��_In   In �շѼ�Ŀ.�䶯ԭ��%Type := 1,
                                            gstrSQL = gstrSQL & "" & "1" & ","
                                            '  No_In         In �շѼ�Ŀ.No%Type := Null,
                                            gstrSQL = gstrSQL & "'" & strNo & "',"
                                            '  ���_In       In �շѼ�Ŀ.���%Type := 1,
                                            gstrSQL = gstrSQL & "" & intRow + 1 & ","
                                            '  ȱʡ�۸�_In   In �շѼ�Ŀ.ȱʡ�۸�%Type := Null,
                                            gstrSQL = gstrSQL & "" & IIF(Val(.TextMatrix(1, mcstColȱʡ�۸�)) = 0, "Null", Val(.TextMatrix(1, mcstColȱʡ�۸�))) & ","
                                            '  ���ۻ��ܺ�_In In �շѼ�Ŀ.���ۻ��ܺ�%Type := Null,
                                            gstrSQL = gstrSQL & "" & "NULL" & ","
                                            '  �۸�ȼ�_In   In �շѼ�Ŀ.�۸�ȼ�%Type := Null
                                            gstrSQL = gstrSQL & "" & IIF(tbPriceGrade(k).Caption = "ȱʡ", "NULL", "'" & tbPriceGrade(k).Caption & "'") & ")"
                                            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                                            intRow = intRow + 1
                                        End If
                                    Next
                                Else
                                    '���ֱ���޸�
                                    'ZL_�շѼ�Ŀ_UPDATE(
                                    gstrSQL = "ZL_�շѼ�Ŀ_UPDATE("
                                    '  �շ�ϸĿid_In In �շѼ�Ŀ.�շ�ϸĿid%Type := Null,
                                    gstrSQL = gstrSQL & "" & mstrID & ","
                                    '  ������Ŀid_In In �շѼ�Ŀ.������Ŀid%Type := Null,
                                    gstrSQL = gstrSQL & "" & .RowData(1) & ","
                                    '  ԭ��_In       In �շѼ�Ŀ.ԭ��%Type := Null,
                                    gstrSQL = gstrSQL & "" & Val(.TextMatrix(1, mcstColԭ��)) & ","
                                    '  �ּ�_In       In �շѼ�Ŀ.�ּ�%Type := Null,
                                    gstrSQL = gstrSQL & "" & Val(.TextMatrix(1, mcstCol�ּ�)) & ","
                                    '  �����շ���_In In �շѼ�Ŀ.�����շ���%Type := Null,
                                    gstrSQL = gstrSQL & "" & IIF(.TextMatrix(1, mcstCol���������շ���) = "", 0, .TextMatrix(1, mcstCol���������շ���)) & ","
                                    '  �Ӱ�Ӽ���_In In �շѼ�Ŀ.�Ӱ�Ӽ���%Type := Null,
                                    gstrSQL = gstrSQL & "" & IIF(.TextMatrix(1, mcstCol�Ӱ�Ӽ���) = "", 0, .TextMatrix(1, mcstCol�Ӱ�Ӽ���)) & ","
                                    '  ����˵��_In   In �շѼ�Ŀ.����˵��%Type := Null,
                                    gstrSQL = gstrSQL & "" & "'" & txt����˵��(k).Text & "',"
                                    '  ����id_In     In �շѼ�Ŀ.����id%Type := Null,
                                    gstrSQL = gstrSQL & "" & lng����ID & ","
                                    '  ������_In     In �շѼ�Ŀ.������%Type := Null,
                                    gstrSQL = gstrSQL & "'" & gstrUserName & "',"
                                    '  ȱʡ�۸�_In   In �շѼ�Ŀ.ȱʡ�۸�%Type := Null,
                                    gstrSQL = gstrSQL & "" & IIF(Val(.TextMatrix(1, mcstColȱʡ�۸�)) = 0, "Null", Val(.TextMatrix(1, mcstColȱʡ�۸�))) & ","
                                    '  �۸�ȼ�_In   In �շѼ�Ŀ.�۸�ȼ�%Type := Null
                                    gstrSQL = gstrSQL & IIF(tbPriceGrade(k).Caption = "ȱʡ", "NULL", "'" & tbPriceGrade(k).Caption & "'") & ")"
                                    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                                End If
                            End If
                        End If
                    Else
                        '���������շ�ϸĿֱ�Ӳ��뵽�շѼ�Ŀ���м���
                        If medit��ʽ = EditNew Or medit��ʽ = EditCopy Then
                            For i = 1 To .Rows - 1
                                If .RowData(i) > 0 Then
                                    If intRow = 0 Then
                                        lng��ĿID = lng����ID
                                    Else
                                        lng��ĿID = sys.NextId("�շѼ�Ŀ")
                                    End If
                                    'Zl_�շѼ�Ŀ_Insert(
                                    gstrSQL = "Zl_�շѼ�Ŀ_Insert("
                                    '  Id_In         In �շѼ�Ŀ.Id%Type,
                                    gstrSQL = gstrSQL & "" & lng��ĿID & ","
                                    '  ԭ��id_In     In �շѼ�Ŀ.ԭ��id%Type := Null,
                                    gstrSQL = gstrSQL & "" & IIF(mcol��Ŀ(k)("C" & .RowData(i)) = 0, "null", mcol��Ŀ(k)("C" & .RowData(i))) & ","
                                    '  �շ�ϸĿid_In In �շѼ�Ŀ.�շ�ϸĿid%Type := Null,
                                    gstrSQL = gstrSQL & "" & mstrID & ","
                                    '  ������Ŀid_In In �շѼ�Ŀ.������Ŀid%Type := Null,
                                    gstrSQL = gstrSQL & "" & .RowData(i) & ","
                                    '  ԭ��_In       In �շѼ�Ŀ.ԭ��%Type := Null,
                                    gstrSQL = gstrSQL & "" & Val(.TextMatrix(i, mcstColԭ��)) & ","
                                    '  �ּ�_In       In �շѼ�Ŀ.�ּ�%Type := Null,
                                    gstrSQL = gstrSQL & "" & Val(.TextMatrix(i, mcstCol�ּ�)) & ","
                                    '  �����շ���_In In �շѼ�Ŀ.�����շ���%Type := Null,
                                    gstrSQL = gstrSQL & "" & IIF(.TextMatrix(i, mcstCol���������շ���) = "", 0, .TextMatrix(i, mcstCol���������շ���)) & ","
                                    '  �Ӱ�Ӽ���_In In �շѼ�Ŀ.�Ӱ�Ӽ���%Type := Null,
                                    gstrSQL = gstrSQL & "" & IIF(.TextMatrix(i, mcstCol�Ӱ�Ӽ���) = "", 0, .TextMatrix(i, mcstCol�Ӱ�Ӽ���)) & ","
                                    '  ����˵��_In   In �շѼ�Ŀ.����˵��%Type := Null,
                                    gstrSQL = gstrSQL & "'" & txt����˵��(k).Text & "',"
                                    '  ����id_In     In �շѼ�Ŀ.����id%Type := Null,
                                    gstrSQL = gstrSQL & "" & lng����ID & ","
                                    '  ������_In     In �շѼ�Ŀ.������%Type := Null,
                                    gstrSQL = gstrSQL & "'" & gstrUserName & "',"
                                    '  ִ������_In   In �շѼ�Ŀ.ִ������%Type := Null,
                                    gstrSQL = gstrSQL & "" & "to_date('" & str��ʼʱ�� & "','YYYY-MM-DD HH24:MI:SS')" & ","
                                    '  �䶯ԭ��_In   In �շѼ�Ŀ.�䶯ԭ��%Type := 1,
                                    gstrSQL = gstrSQL & "" & "1" & ","
                                    '  No_In         In �շѼ�Ŀ.No%Type := Null,
                                    gstrSQL = gstrSQL & "'" & strNo & "',"
                                    '  ���_In       In �շѼ�Ŀ.���%Type := 1,
                                    gstrSQL = gstrSQL & "" & intRow + 1 & ","
                                    '  ȱʡ�۸�_In   In �շѼ�Ŀ.ȱʡ�۸�%Type := Null,
                                    gstrSQL = gstrSQL & "" & IIF(Val(.TextMatrix(1, mcstColȱʡ�۸�)) = 0, "Null", Val(.TextMatrix(1, mcstColȱʡ�۸�))) & ","
                                    '  ���ۻ��ܺ�_In In �շѼ�Ŀ.���ۻ��ܺ�%Type := Null,
                                    gstrSQL = gstrSQL & "" & "NULL" & ","
                                    '  �۸�ȼ�_In   In �շѼ�Ŀ.�۸�ȼ�%Type := Null
                                    gstrSQL = gstrSQL & "" & IIF(tbPriceGrade(k).Caption = "ȱʡ", "NULL", "'" & tbPriceGrade(k).Caption & "'") & ")"
                                    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                                    intRow = intRow + 1
                                End If
                            Next
                        End If
                    End If
                ElseIf medit��ʽ = EditModify Or medit��ʽ = EditRaise Then
                    'Zl_�շѼ�Ŀ_Delete(
                    gstrSQL = "Zl_�շѼ�Ŀ_Delete("
                    '  ϸĿid_In   In �շѼ�Ŀ.�շ�ϸĿid%Type,
                    gstrSQL = gstrSQL & "" & mstrID & ","
                    '  վ��_In     In �շ���ĿĿ¼.վ��%Type := Null,
                    gstrSQL = gstrSQL & "" & "NULL" & ","
                    '  �۸�ȼ�_In In �շѼ�Ŀ.�۸�ȼ�%Type := Null
                    gstrSQL = gstrSQL & "" & IIF(tbPriceGrade(k).Caption = "ȱʡ", "NULL", "'" & tbPriceGrade(k).Caption & "'") & ")"
                    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
                End If
            End With
        ElseIf tbPriceGrade(k).Visible = False And medit��ʽ = EditModify Then
            'ɾ������ļ۸�ȼ�����
            'Zl_�շѼ�Ŀ_Delete(
            gstrSQL = "Zl_�շѼ�Ŀ_Delete("
            '  ϸĿid_In   In �շѼ�Ŀ.�շ�ϸĿid%Type,
            gstrSQL = gstrSQL & "" & mstrID & ","
            '  վ��_In     In �շ���ĿĿ¼.վ��%Type := Null,
            gstrSQL = gstrSQL & "" & "NULL" & ","
            '  �۸�ȼ�_In In �շѼ�Ŀ.�۸�ȼ�%Type := Null
            gstrSQL = gstrSQL & "" & IIF(tbPriceGrade(k).Caption = "ȱʡ", "NULL", "'" & tbPriceGrade(k).Caption & "'") & ")"
            Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
        End If
    Next
    Save��Ŀ = True
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Saveִ��() As Boolean
    Dim i As Long
    Dim str���� As String
    Dim lng���� As Long
    Dim lngӦ�� As Long
    Dim strTemp As String
    Dim strMid As Variant
    Dim intCount As Integer
    Dim strIn As String
    Dim strOut As String
    
    If medit��ʽ <> EditDept And opt����(4).value = False And Not (opt����(0).value And mstrServerObj <> "1") Then
        Saveִ�� = True: Exit Function
    End If
    
    '����ִ�м��
    On Error GoTo ErrHandle
    With Me.msf����ִ��
        strTemp = ""
        For intCount = 1 To .Rows - 1
            If Val(.TextMatrix(intCount, 0)) <> 0 Then
                '���ټ���Ƿ��ظ� By ��ͮ��
                'If InStr(1, strTemp & ";", ";" & Trim(.TextMatrix(intCount, 0)) & "-" & .TextMatrix(intCount, 1) & ";") > 0 Then
                If InStr(1, strTemp & ";", ";" & .TextMatrix(intCount, 0) & ";") > 0 Then
                    MsgBox "�ظ�ָ����ִ�п��ҡ�" & .TextMatrix(intCount, 1) & "����", vbInformation, gstrSysName
                    .SetFocus: Exit Function
                Else
                    strTemp = strTemp & ";" & .TextMatrix(intCount, 0)
                End If
'                    If Val(.TextMatrix(intCount, 2)) = 0 Then
'                        MsgBox "��" & .TextMatrix(intCount, 1) & "��δָ��ִ�п��ң�", vbInformation, gstrSysName
'                        Me.stbInfo.Tab = 1: .SetFocus: Exit Sub
'                    End If
            End If
        Next
        
        strTemp = ""
        
        For intCount = 1 To .Rows - 1
            If Val(.TextMatrix(intCount, 0)) <> 0 Then
                strMid = Split(.TextMatrix(intCount, 2), ",")
                For i = LBound(strMid) To UBound(strMid)
                    strTemp = strTemp & "|" & Trim(IIF(strMid(i) = "�����в��ţ�", 0, strMid(i))) & "^" & Trim(.TextMatrix(intCount, 0))
                Next
            End If
        Next
        If strTemp <> "" Then strTemp = Mid(strTemp, 2)
        str���� = strTemp
        
    End With
    
    If Len(Me.txt����ִ��.Tag) > 0 And opt����(4).value Then
        strOut = Me.txt����ִ��.Tag
    End If
    
    If Len(Me.txtסԺִ��.Tag) > 0 And (opt����(4).value Or opt����(0).value And mstrServerObj <> "1") Then
        strIn = Me.txtסԺִ��.Tag
    End If
    
    For i = 0 To 6
        If opt����(i).value = True Then lng���� = i: Exit For
    Next
    For i = 0 To 3
        If optApply(i).value = True Then lngӦ�� = i: Exit For
    Next
    
    'Zl_�շ�ϸĿ_Dept(
    gstrSQL = "Zl_�շ�ϸĿ_Dept("
    '  �շ�ϸĿid_In In �շ���ĿĿ¼.Id%Type,
    gstrSQL = gstrSQL & "" & mstrID & ","
    '  ִ�п���_In   In Number,
    gstrSQL = gstrSQL & "" & lng���� & ","
    '  Ӧ�÷�Χ_In   In Number,
    gstrSQL = gstrSQL & "" & lngӦ�� & ","
    '  ����id_In     In �շ���ĿĿ¼.����id%Type,
    gstrSQL = gstrSQL & "" & IIF(mstr����ID = "", "Null", mstr����ID) & ","
    '  ���_In       In �շ���ĿĿ¼.���%Type,
    gstrSQL = gstrSQL & "'" & mstr��� & "',"
    '  �����б�_In   In Varchar2, --�������Ҷ���ִ�е�˵��������|�ָÿ�����򰴿�������id^ִ�п���id��ʽ��֯
    gstrSQL = gstrSQL & "'" & str���� & "',"
    '  ����ִ��_In   In ����ִ�п���.ִ�п���id%Type := Null,
    gstrSQL = gstrSQL & "'" & strOut & "',"
    '  סԺִ��_In   In ����ִ�п���.ִ�п���id%Type := Null,
    gstrSQL = gstrSQL & "'" & strIn & "',"
    '  վ��_In       In �շ���ĿĿ¼.վ��%Type := Null
    gstrSQL = gstrSQL & "" & IIF(mblnCanUpdateAll, "NULL", "'" & gstrNodeNo & "'") & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Saveִ�� = True
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function Save����() As Boolean
    Dim i As Integer
    Dim str����id As String
    
    On Error GoTo ErrHandle
    If medit��ʽ = EditSlave Then
        gstrSQL = "zl_�շѴ�����Ŀ_delete(" & mstrID & ")"
        Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    End If
    With msh����
        For i = 1 To .Rows - 1
            If .RowData(i) > 0 Then
                gstrSQL = "zl_�շѴ�����Ŀ_insert(" & _
                mstrID & "," & .RowData(i) & "," & .TextMatrix(i, 1) & "," & Left(.TextMatrix(i, 2), 1) & ")"
                Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
            End If
        Next
    End With
    
    If optʹ�ÿ���(0).value = True Then
        With Lvw����
            For i = 1 To .ListItems.Count
                If .ListItems(i).Checked = True Then
                    str����id = IIF(str����id = "", "", str����id & ",") & Mid(.ListItems(i).Key, 2)
                End If
            Next
        End With
    Else
        str����id = ""
    End If
    gstrSQL = "Zl_�շ����ÿ���_Update(" & mstrID & "," & IIF(str����id = "", "Null", "'" & str����id & "'") & ")"
    Call zlDatabase.ExecuteProcedure(gstrSQL, Me.Caption)
    Save���� = True
    Exit Function
ErrHandle:
    gcnOracle.RollbackTrans
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cboProperty_Click()
    If Me.msf����ִ��.Col = 3 Then
        Load���� 1, cboProperty.Text
    Else
        Load���� 0, cboProperty.Text
    End If
    
    ChkSelect.value = 0
End Sub


Private Sub cboProperty_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyEscape
         picDept.Visible = False
    End Select
End Sub


Private Sub cboProperty_LostFocus()
    Call picDept_LostFocus
End Sub



Private Sub chkNow_Click(Index As Integer)
  '��ǰ�Ƿ���������Ч
    If Me.chkNow(Index).value = 1 Then
        Me.dtpBegin(Index).Enabled = False
        '������ǰʱ�䲻��������Ч
        If Me.dtpBegin(Index).MinDate > sys.Currentdate Then
            MsgBox "�ϴ�ִ��ʱ���ѳ�����ǰʱ�䲻��ʹ��������Ч�����ֶ�����ʱ�䣡", vbInformation
            Me.chkNow(Index).value = 0
        End If
    ElseIf medit��ʽ = EditModify And txtEdit(text��ʶ����).Text <> txtEdit(text��ʶ����).Tag Then
        MsgBox "���Ѿ��ı���ҽ����Ŀ����Ӧ�ļ۸�ֻ��ѡ��������Ч��", vbInformation
        Me.chkNow(Index).value = 1
    Else
        Me.dtpBegin(Index).Enabled = True
    End If
End Sub

Private Sub chkNow_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub ChkSelect_Click()
    Dim i As Integer
    Dim str���� As String
    
    If mblnRefresh = True Then Exit Sub
    
    If ChkSelect.value = 2 Then Exit Sub
    Call SetSelect(lvwItems, ChkSelect.value)
    
    If cboProperty.Text = "��������" Then
        mstr��ѡִ�п��� = ""
    End If
    
    If ChkSelect.value = 1 Then
        '��ǰ����ȫѡ
        For i = 1 To lvwItems.ListItems.Count
            str���� = Mid(lvwItems.ListItems(i).Key, 2) & "," & "[" & lvwItems.ListItems(i).SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) & "]" & lvwItems.ListItems(i).Text
            
            If InStr(mstr��ѡִ�п���, str����) = 0 Or cboProperty.Text = "��������" Then
                mstr��ѡִ�п��� = IIF(mstr��ѡִ�п��� = "", "", mstr��ѡִ�п��� & ";") & str����
            End If
        Next
    ElseIf cboProperty.Text <> "��������" Then
        '��ǰ����ȫ��

        For i = 1 To lvwItems.ListItems.Count
            str���� = Mid(lvwItems.ListItems(i).Key, 2) & "," & "[" & lvwItems.ListItems(i).SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) & "]" & lvwItems.ListItems(i).Text
               
            If InStr(mstr��ѡִ�п���, str����) > 0 Then
                mstr��ѡִ�п��� = Replace(mstr��ѡִ�п���, str����, "")
            End If
        Next
    End If
End Sub

Private Sub SetSelect(ByVal lvwObj As Object, Optional ByVal BlnSelect As Boolean = True)
    Dim intSelect As Integer
    With lvwObj
        For intSelect = 1 To .ListItems.Count
            .ListItems(intSelect).Checked = BlnSelect
        Next
    End With
End Sub

Private Sub ChkSelect_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyEscape
         picDept.Visible = False
    End Select
End Sub

Private Sub chk����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub cmbClass_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub cmbStationNo_Change()
    Call init��Ŀ(True)
End Sub

Private Sub cmbStationNo_Click()
    Dim strStationNo As String
    
    On Error GoTo ErrHandler
    If mblnNotClick Then Exit Sub
    If Val(mstrID) <> 0 Then
        strStationNo = zlStr.NeedCode(cmbStationNo.Text, "-")
        If strStationNo <> "" And cmbStationNo.Tag <> strStationNo Then
            If CanChangeStation(mstrID) = False Then
                mblnNotClick = True
                cbo.SeekIndex cmbStationNo, cmbStationNo.Tag
                mblnNotClick = False
            End If
        End If
    End If
    
    Call init��Ŀ(True)
    Exit Sub
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function CanChangeStation(ByVal lngId As Long)
    '�ж��Ƿ��ܹ�����վ��
    '����ţ�110164
    Dim strSQL As String, rsTemp As ADODB.Recordset
    
    On Error GoTo ErrHandler
    strSQL = _
        " Select 1" & vbNewLine & _
        " From �շѼ�Ŀ A" & vbNewLine & _
        " Where a.�շ�ϸĿid = [1] And a.�۸�ȼ� Is Not Null And Rownum < 2"
    Set rsTemp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lngId)
    If Not rsTemp.EOF Then
        MsgBox "��ǰ��Ŀ�����˼۸�ȼ������������Ϊ����Ժ����", vbInformation, gstrSysName
        Exit Function
    End If
    CanChangeStation = True
    Exit Function
ErrHandler:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub cmb����_Click()
    Me.chk�Զ�����.Visible = (Me.cmb����.ListIndex <> 0)
End Sub
Private Sub cmb����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdFind_Click()
    Dim strFind As String
    Dim i As Long
    Dim blnIsFind As Boolean
    
    strFind = UCase(Trim(txtFind.Text))
    If strFind = "" Then Exit Sub
    For i = mlngFind To Lvw����.ListItems.Count
        If zlStr.GetCodeByVB(Mid(Lvw����.ListItems(i).Text, InStr(Lvw����.ListItems(i).Text, "-") + 1)) Like UCase(IIF(gstrLike <> "", "*", "") & strFind & "*") Or _
                UCase(Lvw����.ListItems(i).Text) Like UCase(IIF(gstrLike <> "", "*", "") & strFind & "*") Then
            Lvw����.ListItems(i).Selected = True
            Lvw����.ListItems(i).EnsureVisible
            Lvw����.SetFocus
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

Private Sub cmdHelp_Click()
    If Me.Caption = "�շ���Ŀ����" Then
        ShowHelp App.ProductName, Me.hwnd, "frmChargeItem", Int((glngSys) / 100)
'    ElseIf Me.Caption = "�շѷ�������" Then
'        ShowHelp App.ProductName, Me.hwnd, "frm�շ���Ŀ����1", Int((glngSys) / 100)
    End If
End Sub

Private Sub cmdOK_Click()
    If IsValid() = False Then Exit Sub
    If SaveϸĿ() = False Then Exit Sub
    'ˢ�������ڵ���ʾ
    Call frmChargeManage.FillTree
    If medit��ʽ <> EditNew And medit��ʽ <> EditCopy Then
        mblnChange = False
        Unload Me
        Exit Sub
    End If
    
    '��������
    ClearContext (chk����.value = 0)
    ShowTab "������Ϣ"
    txtEdit(text����).SetFocus
    mblnChange = False
    mblnOk = True
End Sub

Private Sub ChangeCode(nod As Node, ByVal strOldCode As String, ByVal strNewCode As String)
    '����:�ı��¼��ı�������
    Dim nodChild As Node
    
    Set nodChild = nod.Child
    Do Until nodChild Is Nothing
        nodChild.Text = strNewCode & Mid(nodChild.Text, Len(strOldCode))
        ChangeCode nodChild, strOldCode, strNewCode
        Set nodChild = nodChild.Next
    Loop
End Sub

Private Sub chk���_Click()
    Dim i As Integer
    
    For i = 0 To tbPriceGrade.ItemCount - 1
        If msh��Ŀ(i).Rows > 2 Then
            chk���.value = 0
            Exit Sub
        End If
    Next
    
    '���޸���Ŀ�ı��ˡ����/���ۡ�����ʱ����������
    If medit��ʽ = EditModify Then
        If chk���.value <> chk���.Tag Then
            If Not mblnShow�շѼ�Ŀ Then
                TabMain.Tabs.Add , "_�շѼ�Ŀ", "�շѼ�Ŀ"
                mblnShow�շѼ�Ŀ = True
                Call init��Ŀ: Call init��Ŀ(True)
                MsgBox "������ȷ���շѼ�Ŀ��", vbInformation, gstrSysName
            End If
        Else
            If mblnShow�շѼ�Ŀ Then
                TabMain.Tabs.Remove "_�շѼ�Ŀ"
                mblnShow�շѼ�Ŀ = False
            End If
        End If
    End If
    For i = 0 To tbPriceGrade.ItemCount - 1
        With msh��Ŀ(i)
            If chk���.value = 1 Then
                .Rows = 2
                .TextMatrix(0, mcstColԭ��) = "����޼�"
                .TextMatrix(0, mcstCol�ּ�) = "����޼�"
                .ColData(mcstColԭ��) = IIF(gstrҽ�۽ӿڱ�� <> "" And gbln����ҽ���շ���Ŀ = True, 5, 4)
                .ColData(mcstCol�ּ�) = IIF(gstrҽ�۽ӿڱ�� <> "" And gbln����ҽ���շ���Ŀ = True, 5, 4)
                .TextMatrix(1, mcstColԭ��) = txtEdit(text����޼�).Text
                .TextMatrix(1, mcstCol�ּ�) = txtEdit(text����޼�).Text
                .ColWidth(mcstColȱʡ�۸�) = 1000
            Else
                .TextMatrix(0, mcstColԭ��) = "ԭ��"
                .TextMatrix(0, mcstCol�ּ�) = "�ּ�"
                .TextMatrix(1, mcstColԭ��) = "0.000"
                .ColData(mcstColԭ��) = 5
                .ColData(mcstCol�ּ�) = 4
                .ColWidth(mcstColȱʡ�۸�) = 0
            End If
        End With
    Next
    mblnChange = True
End Sub

Private Sub chk���_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub chk�Ӱ�Ӽ�_Click()
    Dim i As Integer
    
    For i = 0 To tbPriceGrade.ItemCount - 1
        With msh��Ŀ(i)
            If chk�Ӱ�Ӽ�.value = 1 Then
                .ColWidth(mcstCol�Ӱ�Ӽ���) = 1500
                .TextMatrix(0, mcstCol�Ӱ�Ӽ���) = "�Ӱ�Ӽ���"
            Else
                .ColWidth(mcstCol�Ӱ�Ӽ���) = 0
            End If
        End With
    Next
    mblnChange = True
End Sub

Private Sub chk�Ӱ�Ӽ�_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub chk���ηѱ�_Click()
    mblnChange = True
End Sub

Private Sub chk���ηѱ�_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub cmb��������_KeyPress(KeyAscii As Integer)
    On Error GoTo ErrHandle
    Dim lngIdx As Long
    If KeyAscii = vbKeyReturn Then
        OS.PressKey vbKeyTab
    Else
        lngIdx = cbo.MatchIndex(cmb��������.hwnd, KeyAscii)
        If lngIdx <> -2 Then cmb��������.ListIndex = lngIdx
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub chkժҪ_Click()
    mblnChange = True
End Sub

Private Sub chkժҪ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub cmb�������_Click()
    mblnChange = True
End Sub

Private Sub cmb��������_Click()
    mblnChange = True
End Sub

Private Sub cmb�������_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub cmb���㵥λ_Change()
    mblnChange = True
End Sub

Private Sub cmb���㵥λ_KeyPress(KeyAscii As Integer)
    If InStr("'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub








Private Sub cmdOK_GotFocus()
    ''
End Sub

Private Sub cmdOkDept_Click()
    Dim i As Integer
    Dim strTmp As String
    Dim strArr
    Dim n As Integer
    Dim strNew As String
    Dim blnNew As Boolean
    
    With Me.lvwItems
        Select Case .Tag
            Case "ִ��"
                'ɾ��������ѡ���б��е�ִ�п���
                For i = msf����ִ��.Rows - 1 To 1 Step -1
                    If InStr(mstr��ѡִ�п���, msf����ִ��.TextMatrix(i, 0) & "," & msf����ִ��.TextMatrix(i, 1)) = 0 Then
                        If i > 1 Then
                            msf����ִ��.MsfObj.RemoveItem i
                        Else
                            msf����ִ��.TextMatrix(1, 0) = ""
                            msf����ִ��.TextMatrix(1, 1) = ""
                            msf����ִ��.TextMatrix(1, 2) = ""
                            msf����ִ��.TextMatrix(1, 3) = ""
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
                            If strArr(i) = msf����ִ��.TextMatrix(n, 0) & "," & msf����ִ��.TextMatrix(n, 1) Then
                                blnNew = False
                            End If
                        Next
                        If blnNew = True Then
                            strNew = IIF(strNew = "", "", strNew & ";") & strArr(i)
                        End If
                    End If
                Next
                
                If strNew <> "" Then
                    strArr = Split(strNew & ";", ";")
                    For i = 0 To UBound(strArr) - 1
                        If strArr(i) <> "" Then
                            If msf����ִ��.TextMatrix(msf����ִ��.Rows - 1, 1) <> "" Then
                                msf����ִ��.Rows = msf����ִ��.Rows + 1
                            End If
                            msf����ִ��.TextMatrix(msf����ִ��.Rows - 1, 0) = Split(strArr(i), ",")(0)
                            msf����ִ��.TextMatrix(msf����ִ��.Rows - 1, 1) = Split(strArr(i), ",")(1)
                        End If
                    Next
                End If
        End Select
    End With
    
    picDept.Visible = False
End Sub

Private Sub cmd����_Click()
    On Error GoTo ErrHandle
    Dim strSQL As String
    Dim blnRe As Boolean
    Dim str���� As String
    Dim strID As String
    
    strSQL = "Select ���� id,�ϼ� as �ϼ�id, ����, ����, ĩ�� From ������Ŀ Start With �ϼ� Is Null Connect By Prior ���� = �ϼ�"
    blnRe = frmTreeLeafSel.ShowTree(strSQL, strID, str����, "������Ŀ")
    '�ɹ�����
    If blnRe Then
        '�µı����Ŀ��
        lbl������Ŀ.Tag = strID
        txtEdit(text������Ŀ).Text = str����
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmd�ϼ�_Click()
    On Error GoTo ErrHandle
    Dim strSQL As String
    Dim blnRe As Boolean
    Dim str���� As String
    Dim strID As String
    Dim str���� As String
    
    strSQL = "select ID,�ϼ�ID,����,����,���� from �շѷ���Ŀ¼ " & _
        " start with �ϼ�ID is null   connect by prior ID =�ϼ�ID"
    strID = mstr����ID
    str���� = txtEdit(Text����).Text
    str���� = txtTemp.Text
    blnRe = frmTreeSel.ShowTree(strSQL, strID, str����, str����, mstrID, "�շ���Ŀѡ��", "�����շ���Ŀ����", , mstr����, 3, 4, 4, False)
    '�ɹ�����
    If blnRe Then
        '�µı����Ŀ��
        mstr����ID = strID
        txtEdit(Text����).Text = str����
        mstr������� = str����
        Call SetCodeNO
        txtEdit(text����).MaxLength = mlng���볤��
        mblnChange = True
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function IsRaiseByDate(ByVal strID As String) As Boolean
    '�жϸ��շ���Ŀ�Ƿ��ǰ��յ���
    '����True-�ǰ�������
    '����False-���ǰ������
    Dim rs As New ADODB.Recordset
    Dim strSQL As String
    
    On Error GoTo ErrHandle
    strSQL = "Select a.Id" & _
            " From �շ���ĿĿ¼ A, �շѼ�Ŀ D" & _
            " Where a.ID = d.�շ�ϸĿID And Nvl(d.��ֹ����, To_Date('3000-01-01', 'yyyy-mm-dd')) > Sysdate " & _
            "       And d.ִ������<>trunc(d.ִ������,'dd') And d.�շ�ϸĿid=[1] "
            
    Set rs = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(strID))
    IsRaiseByDate = Not (rs.RecordCount > 0)
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub dtpBegin_Change(Index As Integer)
    mblnChange = True
    If Mid(cmbClass.Text, 1, 1) <> "J" And Mid(cmbClass.Text, 1, 1) <> "H" Then
        If DateDiff("s", Me.dtpBegin(Index).value, Format(sys.Currentdate, "yyyy-mm-dd h:m:s")) > 0 Then
            MsgBox "����ִ��ʱ�䲻��С�ڵ�ǰʱ�䣡", vbInformation, gstrSysName
            Me.dtpBegin(Index).value = DateAdd("n", 1, sys.Currentdate)
        End If
    End If
End Sub

Private Sub dtpBegin_GotFocus(Index As Integer)
    msh��Ŀ(Index).CmdVisible = False
End Sub

Private Sub dtpBegin_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    Dim i As Integer
    
    picOne.Top = mshAlias.Top + mshAlias.Height + IIF(txtEdit(text����޼�).Visible, txtEdit(text����޼�).Height + 100, 0) + 100
    picTwo.Top = picOne.Top + picOne.Height + IIF(txtEdit(text����).Visible, txtEdit(text����).Height + 50, 0) + 50
    lblEdit(18).Top = picOne.Top + picOne.Height + 85
    txtEdit(text����).Top = picOne.Top + picOne.Height + 50
    fra(1).Height = 6660 - (IIF(Not txtEdit(text����޼�).Visible, txtEdit(text����޼�).Height + 50, 0) + 50) - (IIF(Not txtEdit(text����).Visible, txtEdit(text����).Height + 50, 0) + 50) - 100
    For i = 2 To 4
        fra(i).Left = fra(1).Left
        fra(i).Top = fra(1).Top
        fra(i).Height = fra(1).Height
    Next
    TabMain.Height = fra(1).Height + 350
    Me.Height = TabMain.Height + 1080
    cmdOK.Top = TabMain.Top + TabMain.Height + 100
    cmdCancel.Top = cmdOK.Top
    cmdHelp.Top = cmdOK.Top
    tbPriceGrade.Height = fra(2).Height - tbPriceGrade.Top - 100
    Frame1.Height = fra(3).Height - IIF(fra����.Visible, fra����.Height, 0) - 200
    msf����ִ��.Height = Frame1.Height - (lbl����ִ��.Top + lbl����ִ��.Height) - 150
    fra����.Top = Frame1.Top + Frame1.Height + 100
    Lvw����.Height = fra(4).Height - Label2.Top - Label2.Height - 250
End Sub

Private Sub lvwItems_DblClick()
    Dim i As Integer
    Dim m As Integer
    Dim blnBatch As Boolean
    Dim str���˿���ID As String
    Dim str���˿������� As String
    
    If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
    With Me.lvwItems
        Select Case .Tag
        Case "����"
            Me.txt����ִ��.Tag = Mid(.SelectedItem.Key, 2)
            Me.txt����ִ��.Text = .SelectedItem.Text
            Me.txt����ִ��.SetFocus: Call OS.PressKey(vbKeyTab)
        Case "סԺ"
            Me.txtסԺִ��.Tag = Mid(.SelectedItem.Key, 2)
            Me.txtסԺִ��.Text = .SelectedItem.Text
            Me.txtסԺִ��.SetFocus: Call OS.PressKey(vbKeyTab)
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
                    Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, 2) = Mid(.SelectedItem.Key, 2)
                    Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, 3) = Me.msf����ִ��.Text
                End If
            End With
            
            '���������δ����У�ѯ���Ƿ�ͬһ��������
            For i = 1 To Me.msf����ִ��.Rows - 1
                If Me.msf����ִ��.TextMatrix(i, 0) <> "" And Me.msf����ִ��.TextMatrix(i, 3) = "" Then
                    blnBatch = True
                    Exit For
                End If
            Next
            
            If blnBatch = True Then
                If MsgBox("�Ƿ�Ӧ��������δ���õ��У�", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbYes Then
                    str���˿���ID = Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, 2)
                    str���˿������� = Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, 3)
                    For i = 1 To Me.msf����ִ��.Rows - 1
                        If Me.msf����ִ��.TextMatrix(i, 3) = "" Then
                            Me.msf����ִ��.TextMatrix(i, 2) = str���˿���ID
                            Me.msf����ִ��.TextMatrix(i, 3) = str���˿�������
                        End If
                    Next
                End If
            End If
            
            Me.msf����ִ��.SetFocus
            Call OS.PressKey(vbKeyReturn)
        Case "ִ��"
            Dim strTmp As String
            Dim strArr
            Dim n As Integer
            Dim strNew As String
            Dim blnNew As Boolean
            
            If Val(Me.picDept.Tag) = 1 And lbl��������.Visible = True Then
                'ɾ��������ѡ���б��е�ִ�п���
                For i = msf����ִ��.Rows - 1 To 1 Step -1
                    If InStr(mstr��ѡִ�п���, msf����ִ��.TextMatrix(i, 0) & "," & msf����ִ��.TextMatrix(i, 1)) = 0 Then
                        If i > 1 Then
                            msf����ִ��.MsfObj.RemoveItem i
                        Else
                            msf����ִ��.TextMatrix(1, 0) = ""
                            msf����ִ��.TextMatrix(1, 1) = ""
                            msf����ִ��.TextMatrix(1, 2) = ""
                            msf����ִ��.TextMatrix(1, 3) = ""
                            
                            If msf����ִ��.Rows > 2 Then
                                msf����ִ��.MsfObj.RemoveItem 1
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
                            If strArr(i) = msf����ִ��.TextMatrix(n, 0) & "," & msf����ִ��.TextMatrix(n, 1) Then
                                blnNew = False
                            End If
                        Next
                        If blnNew = True Then
                            strNew = IIF(strNew = "", "", strNew & ";") & strArr(i)
                        End If
                    End If
                Next
                
                If strNew <> "" Then
                    strArr = Split(strNew & ";", ";")
                    For i = 0 To UBound(strArr) - 1
                        If strArr(i) <> "" Then
                            If msf����ִ��.TextMatrix(msf����ִ��.Rows - 1, 1) <> "" Then
                                msf����ִ��.Rows = msf����ִ��.Rows + 1
                            End If
                            msf����ִ��.TextMatrix(msf����ִ��.Rows - 1, 0) = Split(strArr(i), ",")(0)
                            msf����ִ��.TextMatrix(msf����ִ��.Rows - 1, 1) = Split(strArr(i), ",")(1)
                        End If
                    Next
                End If
                
                msf����ִ��.Row = msf����ִ��.Rows - 1
                Me.msf����ִ��.SetFocus
                Call OS.PressKey(vbKeyRight)
            Else
                Me.msf����ִ��.Text = "[" & .SelectedItem.SubItems(.ColumnHeaders("����").Index - 1) & "]" & .SelectedItem.Text
                Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, 0) = Mid(.SelectedItem.Key, 2)
                Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, 1) = Me.msf����ִ��.Text
                Me.msf����ִ��.SetFocus
                Call OS.PressKey(vbKeyRight)
            End If

            picDept.Visible = False
        End Select
    End With
End Sub

Private Sub lvwItems_ItemCheck(ByVal Item As MSComctlLib.ListItem)
    Dim str���� As String
    
    If Me.lvwItems.Tag = "ִ��" Then
        str���� = Mid(Item.Key, 2) & "," & "[" & Item.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) & "]" & Item.Text
        
        If Item.Checked = True Then
            If InStr(mstr��ѡִ�п���, str����) = 0 Then
                mstr��ѡִ�п��� = IIF(mstr��ѡִ�п��� = "", "", mstr��ѡִ�п��� & ";") & str����
            End If
        Else
            If InStr(mstr��ѡִ�п���, str����) > 0 Then
                mstr��ѡִ�п��� = Replace(mstr��ѡִ�п���, str����, "")
            End If
        End If
    End If
End Sub

Private Sub lvwItems_KeyDown(KeyCode As Integer, Shift As Integer)
    If Me.lvwItems.Tag = "����" Or Me.lvwItems.Tag = "ִ��" Then
        If KeyCode = vbKeyA And Shift = vbCtrlMask Then 'ȫѡ Ctrl+A
            If Me.lvwItems.Tag = "ִ��" Then
                If Me.ChkSelect.value = 0 Then
                    Me.ChkSelect.value = 1
                    Call SetSelect(lvwItems, True)
                End If
            Else
                Call SetSelect(lvwItems, True)
            End If
        End If
        
        If KeyCode = vbKeyR And Shift = vbCtrlMask Then     'ȫ�� Ctrl+R
            If Me.lvwItems.Tag = "ִ��" Then
                If Me.ChkSelect.value = 1 Then
                    Me.ChkSelect.value = 0
                    Call SetSelect(lvwItems, False)
                End If
            Else
                Call SetSelect(lvwItems, False)
            End If
        End If
    End If
End Sub

Private Sub lvwItems_GotFocus()
    If Me.lvwItems.Tag = "����" Or Me.lvwItems.Tag = "ִ��" Then
        Me.lvwItems.ToolTipText = "ȫѡCtrl+A��ȫ��Ctrl+R"
    Else
        Me.lvwItems.ToolTipText = ""
    End If
End Sub
Private Sub lvwItems_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyReturn, vbKeySpace
        If Me.lvwItems.SelectedItem Is Nothing Then Exit Sub
        If lvwItems.Checkboxes = True And KeyAscii = vbKeySpace Then Exit Sub
        Call lvwItems_DblClick
    Case vbKeyEscape
         picDept.Visible = False
    End Select

End Sub

Private Sub lvwItems_LostFocus()
    Call picDept_LostFocus
End Sub

Private Sub Lvw����_ItemClick(ByVal Item As MSComctlLib.ListItem)
    mlngFind = Item.Index + 1
End Sub

Private Sub Lvw����_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyF3 Then Call cmdFind_Click
End Sub

Private Sub Lvw����_KeyPress(KeyAscii As Integer)
    Dim i As Long
    
    If KeyAscii = vbKeyBack Then Exit Sub
    
    For i = 1 To Lvw����.ListItems.Count
        If zlStr.GetCodeByVB(Mid(Lvw����.ListItems(i).Text, InStr(Lvw����.ListItems(i).Text, "-") + 1)) Like UCase(Chr(KeyAscii)) & "*" Then
            Lvw����.ListItems(i).Selected = True: Exit For
        End If
    Next
End Sub

Private Sub msf����ִ��_CommandClick()
    Dim i As Integer
    
    mstr��ѡִ�п��� = ""
    If Me.msf����ִ��.Col = 1 Then
        With Me.msf����ִ��
            For i = 1 To .Rows - 1
                If .TextMatrix(i, 1) <> "" Then
                    mstr��ѡִ�п��� = IIF(mstr��ѡִ�п��� = "", "", mstr��ѡִ�п��� & ";") & .TextMatrix(i, 0) & "," & .TextMatrix(i, 1)
                End If
            Next
        End With
    End If
    
    With Me.picDept
        If Me.msf����ִ��.Col = 3 Then
            .Tag = ""
            Me.lvwItems.Tag = "����"
            .Left = Me.fra(3).Left + Me.msf����ִ��.Left + Me.msf����ִ��.ColWidth(0) + Me.msf����ִ��.ColWidth(1) + Me.msf����ִ��.ColWidth(2)
            .Width = IIF(Me.msf����ִ��.ColWidth(3) < 3000, 3000, Me.msf����ִ��.ColWidth(3))
        Else
            .Tag = "1"
            Me.lvwItems.Tag = "ִ��"
            .Left = Me.fra(3).Left + Me.msf����ִ��.Left + Me.msf����ִ��.ColWidth(0)
            .Width = IIF(Me.msf����ִ��.ColWidth(2) < 3000, 3000, Me.msf����ִ��.ColWidth(2))
        End If
        
        .Top = Me.fra(3).Top + Me.Frame1.Top + Me.msf����ִ��.Top + (Me.msf����ִ��.Row - Me.msf����ִ��.MsfObj.TopRow + 1) * Me.msf����ִ��.RowHeight(0) - 50
        
        If fra����.Top + fra����.Height - .Top - 50 > 0 Then
            .Height = fra����.Top + fra����.Height - .Top - 50
        Else
            .Height = (fra����.Top - Frame1.Top - Frame1.Height) + fra����.Height
        End If
        
        lbl��������.Visible = (Me.msf����ִ��.Col = 1)
        cboProperty.Visible = lbl��������.Visible
        ChkSelect.Visible = lbl��������.Visible
        
        If Me.lvwItems.Tag = "ִ��" Then
            lbl��������.Left = 50
            ChkSelect.Left = .Width - ChkSelect.Width - 50
            cboProperty.Width = ChkSelect.Left - cboProperty.Left - 50
        End If
        
        .ZOrder 0
        .Visible = True
    End With
    
    With Me.lvwItems
        If .Tag = "ִ��" Then
            .Left = lbl��������.Left
            .Top = cboProperty.Top + cboProperty.Height + 50
            .Width = Me.picDept.Width - .Left - 50
            .Height = Me.picDept.Height - .Top - 50
        Else
            .Left = 0
            .Top = 0
            .Width = Me.picDept.Width
            .Height = Me.picDept.Height
        End If
        
        .SetFocus
        .Refresh
    End With
    
    If Me.msf����ִ��.Col = 3 Then
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

Private Sub msf����ִ��_EnterCell(Row As Long, Col As Long)
    strInputed = Me.msf����ִ��.TextMatrix(Row, Col)
End Sub

Private Sub msf����ִ��_GotFocus()
    If Me.lvwItems.Visible Then Me.lvwItems.SetFocus
End Sub

Private Sub msf����ִ��_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strTemp As String
    Dim rsTmp As New ADODB.Recordset
    Dim ObjItem As ListItem
    If KeyCode <> vbKeyReturn Then Exit Sub
    With Me.msf����ִ��
        If .Active = False Then Call OS.PressKey(vbKeyTab): Exit Sub
        If .TxtVisible = False Then
            If .Col = 1 And .TextMatrix(.Row, 1) = "" Then
                ShowTab "������Ŀ"
                '----------------------------------
                'û���ҵ���������,��������ķ��洦��
                OS.PressKey (vbKeyTab)
                OS.PressKey (vbKeyTab)
                OS.PressKey (vbKeyTab)
                If .Row = 1 Then
                    OS.PressKey (vbKeyTab)
                End If
                '-----------------------------------
                Exit Sub
            End If
            If .Col = 3 And (.TextMatrix(.Row, 3) = "") Then
                .TextMatrix(.Row, 3) = "�����в��ţ�"
                .TextMatrix(.Row, 2) = "�����в��ţ�"
                Exit Sub
            End If
            strTemp = UCase(Trim(.TextMatrix(.Row, .Col)))
        Else
            If .Col = 1 And Trim(.Text) = "" Then
                If .Row = 1 Then .SetFocus: Call OS.PressKey(vbKeyTab)
                Exit Sub
            End If
            
            If .Col = 3 And Trim(.Text) = "" Then
                .TextMatrix(.Row, 3) = ""
                .TextMatrix(.Row, 2) = ""
                Exit Sub
            End If
            strTemp = UCase(Trim(.Text))
        End If
    End With
    If strTemp = strInputed Then Exit Sub
    If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
    
    Err = 0: On Error GoTo ErrHand
    
    If Me.msf����ִ��.Col = 3 Then
        gstrSQL = "select distinct ID,����,����" & _
                " from ���ű� D,��������˵�� T" & _
                " where D.ID=T.����ID and ��������='�ٴ�'" & _
                "       and (D.����ʱ�� is null or D.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))" & _
                "       and (D.���� like [1] or D.���� like [1] or D.���� like [1])" & _
                " order by ����"
    Else
        gstrSQL = "select distinct ID,����,����" & _
                " from ���ű� D,��������˵�� T" & _
                " where D.ID=T.����ID and T.������� in (1,2,3)" & _
                "       and (D.����ʱ�� is null or D.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))" & _
                "       and (D.���� like [1] or D.���� like [1] or D.���� like [1])" & _
                " order by ����"
    End If
    
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strTemp & "%")
    
    With rsTmp
        If .BOF Or .EOF Then
            MsgBox "δ�ҵ�ָ�����ţ����������룡", vbExclamation, gstrSysName: Cancel = True: Exit Sub
        End If
        If .RecordCount = 1 Then
            Me.msf����ִ��.Text = "[" & !���� & "]" & !����
            If Me.msf����ִ��.Col = 1 Then
                Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, 0) = !ID
                Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, 1) = Me.msf����ִ��.Text
            Else
                Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, 2) = !ID
                Me.msf����ִ��.TextMatrix(Me.msf����ִ��.Row, 3) = Me.msf����ִ��.Text
            End If
            Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Me.lvwItems.Checkboxes = False
        Do While Not .EOF
            Set ObjItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !����)
            ObjItem.Icon = "Dept": ObjItem.SmallIcon = "Dept"
            ObjItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = !����
            
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    
    With Me.picDept
        If Me.msf����ִ��.Col = 3 Then
            .Tag = ""
            Me.lvwItems.Tag = "����"
            .Left = Me.fra(3).Left + Me.msf����ִ��.Left + Me.msf����ִ��.ColWidth(0) + Me.msf����ִ��.ColWidth(1) + Me.msf����ִ��.ColWidth(2)
            .Width = IIF(Me.msf����ִ��.ColWidth(3) < 3000, 3000, Me.msf����ִ��.ColWidth(3))
        Else
            .Tag = "1"
            Me.lvwItems.Tag = "ִ��"
            .Left = Me.fra(3).Left + Me.msf����ִ��.Left
            .Width = IIF(Me.msf����ִ��.ColWidth(2) < 3000, 3000, Me.msf����ִ��.ColWidth(2))
        End If
        
        .Top = Me.fra(3).Top + Me.Frame1.Top + Me.msf����ִ��.Top + (Me.msf����ִ��.Row - Me.msf����ִ��.MsfObj.TopRow + 1) * Me.msf����ִ��.RowHeight(0) - 50
        
        If fra����.Top + fra����.Height - .Top - 50 > 0 Then
            .Height = fra����.Top + fra����.Height - .Top - 50
        Else
            .Height = (fra����.Top - Frame1.Top - Frame1.Height) + fra����.Height
        End If
        
        lbl��������.Visible = False
        cboProperty.Visible = lbl��������.Visible
        ChkSelect.Visible = lbl��������.Visible
        
        If Me.msf����ִ��.Col = 1 Then
            lbl��������.Left = 50
            ChkSelect.Left = .Width - ChkSelect.Width - 50
            cboProperty.Width = ChkSelect.Left - cboProperty.Left - 50
        End If
        
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
    Cancel = True
    Exit Sub

ErrHand:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub mshAlias_EditKeyPress(KeyAscii As Integer)
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub mshAlias_KeyPress(KeyAscii As Integer)
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 Then KeyAscii = 0
End Sub

Private Sub msh��Ŀ_AfterAddRow(Index As Integer, Row As Long)
    If chk���.value = 1 Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub optʹ�ÿ���_Click(Index As Integer)
    If Index = 1 Then
        Lvw����.Enabled = False
        txtFind.Enabled = False
        cmdFind.Enabled = False
        txtFind.BackColor = &H8000000F
    Else
        Lvw����.Enabled = True
        txtFind.Enabled = True
        cmdFind.Enabled = True
        txtFind.BackColor = &H80000005
    End If
End Sub

Private Sub picDept_LostFocus()
    Dim strActive As String
    
    strActive = UCase(Me.ActiveControl.Name)
    
    If InStr(1, "CMDOKDEPT,CMDCANCELDEPT,LVWITEMS,CBOPROPERTY,PICDEPT,CHKSELECT", strActive) <> 0 Then
        Exit Sub
    End If

    picDept.Visible = False
End Sub

Private Sub pic�۸�ȼ�_Resize(Index As Integer)
    On Error Resume Next
    txt����˵��(Index).Top = pic�۸�ȼ�(Index).ScaleHeight - txt����˵��(Index).Height - 100
    lbl����˵��(Index).Top = txt����˵��(Index).Top + (txt����˵��(Index).Height - lbl����˵��(Index).Height) \ 2
    dtpBegin(Index).Top = txt����˵��(Index).Top - dtpBegin(Index).Height - 100
    lbl����ִ��ʱ��(Index).Top = dtpBegin(Index).Top + (dtpBegin(Index).Height - lbl����ִ��ʱ��(Index).Height) \ 2
    chkNow(Index).Top = dtpBegin(Index).Top + (dtpBegin(Index).Height - chkNow(Index).Height) \ 2
    msh��Ŀ(Index).Height = dtpBegin(Index).Top - msh��Ŀ(Index).Top - 100
End Sub

Private Sub tbPriceGrade_SelectedChanged(ByVal Item As XtremeSuiteControls.ITabControlItem)
'    msh��Ŀ(Val(tbPriceGrade.Tag)).
    tbPriceGrade.Tag = Item.Index
End Sub

Private Sub txt¼������_KeyPress(KeyAscii As Integer)
    Select Case KeyAscii
    Case vbKeyBack, vbKeyEscape, 3, 22
        Exit Sub
    Case vbKeyReturn
        Call OS.PressKey(vbKeyTab): Exit Sub
    Case Else
        If KeyAscii >= Asc("0") And KeyAscii <= Asc("9") Or KeyAscii = Asc(".") Then Exit Sub
    End Select
    KeyAscii = 0
End Sub

Private Sub txt¼������_LostFocus()
    Me.txt¼������.Text = FormatEx(Val(Me.txt¼������.Text), 5)
End Sub


Private Sub txt����ִ��_Change()
    If Trim(txt����ִ��.Text) = "" Then
        txt����ִ��.Tag = ""
    End If
End Sub

Private Sub txt����ִ��_GotFocus()
     Me.txt����ִ��.SelStart = 0: Me.txt����ִ��.SelLength = 100
End Sub
Private Sub mshAlias_AfterDeleteRow()
    mblnChange = True
End Sub
Private Sub mshAlias_EnterCell(Row As Long, Col As Long)
    If Col = 0 Then
        OS.OpenIme True
        mshAlias.MaxLength = mlng��������
    Else
        OS.OpenIme False
        mshAlias.MaxLength = mlng���볤��
    End If
End Sub
Private Sub mshAlias_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim strTemp As String
    
    If KeyCode = vbKeyReturn Then
        If mshAlias.TxtVisible = False Then
            If mshAlias.Col = 0 And mshAlias.Row = 1 Then OS.PressKey vbKeyTab
            Exit Sub
        End If
        strTemp = mshAlias.Text
        If mshAlias.Col = 0 Then
            If zlCommFun.StrIsValid(strTemp, mlng��������) = False Then
                Cancel = True
                If mshAlias.Active And mshAlias.Visible Then
                    mshAlias.TxtSetFocus
                End If
            Else
                mshAlias.TextMatrix(mshAlias.Row, 1) = zlStr.GetCodeByORCL(strTemp, False, mlng���볤��)
                mshAlias.TextMatrix(mshAlias.Row, 2) = zlStr.GetCodeByORCL(strTemp, True, mlng���볤��)
                
                If mshAlias.TextMatrix(mshAlias.Row, 1) = "" Then mshAlias.TextMatrix(mshAlias.Row, 1) = " "
                If mshAlias.TextMatrix(mshAlias.Row, 2) = "" Then mshAlias.TextMatrix(mshAlias.Row, 2) = " "
            End If
        Else
            Cancel = Not zlCommFun.StrIsValid(strTemp, mlng���볤��)
            If Cancel = True Then
                If mshAlias.Active And mshAlias.Visible Then
                    mshAlias.TxtSetFocus
                End If
            Else
                If strTemp = "" Then mshAlias.Text = " "
            End If
        End If
    End If
    If Cancel = False Then mblnChange = True
End Sub

Private Sub msh��Ŀ_BeforeDeleteRow(Index As Integer, Row As Long, Cancel As Boolean)
    If msh��Ŀ(Index).RowData(Row) <> 0 Then
        mcol��Ŀ(Index).Remove "C" & msh��Ŀ(Index).RowData(Row)
        mblnChange = True: mblnChanged��Ŀ(Index) = True
    End If
End Sub

Private Sub msh��Ŀ_LostFocus(Index As Integer)
    If chk���.value = 1 Then
        msh��Ŀ(Index).Rows = 2
    End If
    msh��Ŀ(Index).CmdVisible = False
End Sub

Private Sub msh��Ŀ_CommandClick(Index As Integer)
    On Error GoTo ErrHandle
    Dim strSQL As String
    Dim blnRe As Boolean
    Dim strTemp As String
    Dim strID As String
    Dim lngRow As Long
    
    With msh��Ŀ(Index)
        lngRow = .Row
        strTemp = .TextMatrix(lngRow, mcstCol�շ���Ŀ)
        strID = .RowData(lngRow)
        strSQL = "select ID,�ϼ�ID,����,ĩ��  from ������Ŀ where " & Where����ʱ��() & _
            "  start with �ϼ�ID is null  connect by prior ID =�ϼ�ID"
        blnRe = frmTreeLeafSel.ShowTree(strSQL, strID, strTemp, "������Ŀ")
        If blnRe Then
            On Error Resume Next
            If .RowData(lngRow) <> strID Then
                mcol��Ŀ(Index).Add 0, "C" & strID
                If Err <> 0 Then
                    MsgBox "��������Ŀ�������˼�Ŀ��", vbExclamation, gstrSysName
                    Exit Sub
                End If
                If .RowData(lngRow) > 0 Then mcol��Ŀ(Index).Remove "C" & .RowData(lngRow)
                .RowData(lngRow) = strID
            End If
            .TextMatrix(lngRow, mcstCol�շ���Ŀ) = strTemp
            If .TextMatrix(lngRow, mcstCol���������շ���) = "" Then .TextMatrix(lngRow, mcstCol���������շ���) = "100.0"
            If .TextMatrix(lngRow, mcstColԭ��) = "" Then .TextMatrix(lngRow, mcstColԭ��) = "0.000"
            mblnChange = True
            mblnChanged��Ŀ(Index) = True
        End If
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub msh����_CommandClick()
    On Error GoTo ErrHandle
    Dim strSQL As String
    Dim strTemp As String
    Dim strID As String
    Dim i As Integer
    Dim lngRow As Long
    Dim rsTmp As New ADODB.Recordset
    Dim strReturn As String
    Dim strHyID As Long
    Dim strWherePriceGrade As String
    
    With msh����
        'û�����շ����Ͳ�����
        lngRow = .Row '�ñ�������
        strTemp = .TextMatrix(lngRow, 0)
        strID = .RowData(lngRow)
        
        If gstr��ͨ�۸�ȼ� = "" And gstrҩƷ�۸�ȼ� = "" And gstr���ļ۸�ȼ� = "" Then
           strWherePriceGrade = " And d.�۸�ȼ� Is Null"
        Else
           strWherePriceGrade = "" & _
               " And ((Instr(';5;6;7;', ';' || a.��� || ';') > 0 And d.�۸�ȼ� = [1])" & vbNewLine & _
               "      Or (Instr(';4;', ';' || a.��� || ';') > 0 And d.�۸�ȼ� = [2])" & vbNewLine & _
               "      Or (Instr(';4;5;6;7;', ';' || a.��� || ';') = 0 And d.�۸�ȼ� = [3])" & vbNewLine & _
               "      Or (d.�۸�ȼ� Is Null" & vbNewLine & _
               "          And Not Exists (Select 1" & vbNewLine & _
               "                          From �շѼ�Ŀ" & vbNewLine & _
               "                          Where d.�շ�ϸĿid = �շ�ϸĿid And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
               "                                And ((Instr(';5;6;7;', ';' || a.��� || ';') > 0 And �۸�ȼ� = [1])" & vbNewLine & _
               "                                      Or (Instr(';4;', ';' || a.��� || ';') > 0 And �۸�ȼ� = [2])" & vbNewLine & _
               "                                      Or (Instr(';4;5;6;7;', ';' || a.��� || ';') = 0 And �۸�ȼ� = [3])))))"
        End If
        strSQL = _
            "SELECT A.����,A.����,A.���,A.���㵥λ," & _
            "       ltrim(rtrim(to_char(Sum(nvl(D.�ּ�,0)),'9999999990.00'))) �۸�,A.ID" & _
            " FROM �շ���ĿĿ¼ A,�շѼ�Ŀ D" & _
            " WHERE A.ID=D.�շ�ϸĿID(+) And a.ID>0" & _
            "       And (A.����ʱ��=to_date('3000-01-01','yyyy-mm-dd') or A.����ʱ�� is null)" & _
            "       And D.ִ������ <= SYSDATE AND (D.��ֹ���� > SYSDATE OR D.��ֹ���� IS NULL)" & _
                    strWherePriceGrade & vbNewLine & _
            " Group By A.����,A.����,A.���,A.���㵥λ,A.ID"

'            "SELECT DISTINCT C.���� ���,B.���� ����,A.����,A.����," & vbCrLf & _
'            "      A.���,A.����,A.���㵥λ,ltrim(rtrim(to_char(nvl(D.�۸�,0),'9999999990.00'))) �۸�,A.ID" & vbCrLf & _
'            "  FROM �շ���ĿĿ¼ A, �շѷ���Ŀ¼ B, �շ���Ŀ��� C,(SELECT �շ�ϸĿID,SUM(�۸�) AS �۸�  FROM (" & vbCrLf & _
'            "        SELECT �շ�ϸĿID,SUM(�ּ�) AS �۸� FROM �շѼ�Ŀ " & vbCrLf & _
'            "          WHERE ִ������ <= SYSDATE AND (��ֹ���� > SYSDATE OR ��ֹ���� IS NULL) " & vbCrLf & _
'            "          GROUP BY  �շ�ϸĿID " & vbCrLf & _
'            "          UNION All " & vbCrLf & _
'            "        SELECT m.����ID �շ�ϸĿID,SUM(n.�ּ�) AS �۸� FROM �շѼ�Ŀ n ,�շѴ�����Ŀ m " & vbCrLf & _
'            "         WHERE m.����id = n.�շ�ϸĿId " & vbCrLf & _
'            "          AND  n.ִ������<=SYSDATE AND (n.��ֹ����> SYSDATE OR n.��ֹ���� IS null) " & vbCrLf & _
'            "          GROUP BY m.����ID) GROUP BY �շ�ϸĿID  ) D" & vbCrLf & _
'            " WHERE A.����ID = B.ID(+) AND A.��� = C.���� AND  (A.����ʱ��=to_date('3000-01-01','yyyy-mm-dd') or A.����ʱ�� is null)  " & vbCrLf & _
'            "   AND A.ID = D.�շ�ϸĿID(+) AND " & Where����ʱ��("A")
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, gstrҩƷ�۸�ȼ�, gstr���ļ۸�ȼ�, gstr��ͨ�۸�ȼ�)
        If rsTmp.RecordCount < 1 Then Exit Sub
        
        strReturn = frmSelCur.ShowCurrSel(Me, rsTmp, "����,1000,0,2;����,2500,0,2;���,1500,0,2;��λ,1500,0,2;�۸�,1000,1,2;ID,0,0,2", _
            "��Ŀѡ����", True, strTemp, 3, 1000 + 2500 + 1500 + 1500 + 1000 + 1800)
        If Trim(strReturn) = "" Then Exit Sub
        For i = 1 To .Rows - 1
            If .RowData(i) > 0 And .RowData(i) = Split(strReturn, ",")(UBound(Split(strReturn, ","))) Then
                MsgBox "���շ���Ŀ�ѱ���Ϊ�������ˡ�", vbExclamation, gstrSysName
                Exit Sub
            End If
            
        Next
        If Val(Split(strReturn, ",")(UBound(Split(strReturn, ",")))) = Val(mstrID) And Val(mstrID) > 0 Then
            MsgBox "�շ���Ŀ��������Ϊ�Լ��Ĵ�����Ŀ��", vbExclamation, gstrSysName
            Exit Sub
        End If
        '�ݹ���
        strHyID = Split(strReturn, ",")(UBound(Split(strReturn, ",")))
        If CheckHypotaxis(strHyID) = True Then
            MsgBox "���շ���Ŀ�Ѵ��ڴ���������������Ϊ���ӹ�����", vbExclamation, gstrSysName
            Exit Sub
        End If
        
        '�����������Ŀ���������Ŀ�ļ۸�ִ������ֻ�ܰ��յ���
        If mblnIsSpecialItem Then
            If Not IsRaiseByDate(Val(strHyID)) Then
                 MsgBox "[" & Split(strReturn, ",")(0) & "]" & Split(strReturn, ",")(1) & "�ļ۸�������ǰ�����ִ�еģ�������Ϊ������Ŀ��", vbOKOnly + vbInformation, gstrSysName
                 Exit Sub
            End If
        End If
        
        .RowData(lngRow) = Split(strReturn, ",")(UBound(Split(strReturn, ",")))
        .TextMatrix(lngRow, 0) = "[" & Split(strReturn, ",")(0) & "]" & Split(strReturn, ",")(1)
        If .TextMatrix(lngRow, 1) = "" Then .TextMatrix(lngRow, 1) = "0"
        
        If gstr��ͨ�۸�ȼ� = "" And gstrҩƷ�۸�ȼ� = "" And gstr���ļ۸�ȼ� = "" Then
           strWherePriceGrade = " And b.�۸�ȼ� Is Null"
        Else
           strWherePriceGrade = "" & _
               " And ((Instr(';5;6;7;', ';' || a.��� || ';') > 0 And b.�۸�ȼ� = [2])" & vbNewLine & _
               "      Or (Instr(';4;', ';' || a.��� || ';') > 0 And b.�۸�ȼ� = [3])" & vbNewLine & _
               "      Or (Instr(';4;5;6;7;', ';' || a.��� || ';') = 0 And b.�۸�ȼ� = [4])" & vbNewLine & _
               "      Or (b.�۸�ȼ� Is Null" & vbNewLine & _
               "          And Not Exists (Select 1" & vbNewLine & _
               "                          From �շѼ�Ŀ" & vbNewLine & _
               "                          Where b.�շ�ϸĿid = �շ�ϸĿid And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
               "                                And ((Instr(';5;6;7;', ';' || a.��� || ';') > 0 And �۸�ȼ� = [2])" & vbNewLine & _
               "                                      Or (Instr(';4;', ';' || a.��� || ';') > 0 And �۸�ȼ� = [3])" & vbNewLine & _
               "                                      Or (Instr(';4;5;6;7;', ';' || a.��� || ';') = 0 And �۸�ȼ� = [4])))))"
        End If
        strSQL = "SELECT a.id,a.�Ƿ���,sum(b.ԭ��) ԭ��,sum(b.�ּ�) �ּ�," & vbCrLf & _
                "       decode(nvl(a.�Ƿ���,0),1,ltrim(rtrim(to_char(sum(b.ԭ��),'9999999990.00')))||'��'||ltrim(rtrim(to_char(sum(b.�ּ�),'9999999990.00'))),ltrim(rtrim(to_char(sum(b.�ּ�),'9999999990.00'))))  AS  �۸� " & vbCrLf & _
                " FROM �շ���ĿĿ¼ a,�շѼ�Ŀ b " & vbCrLf & _
                " WHERE a.id=b.�շ�ϸĿid AND  a.id=[1] " & vbCrLf & _
                "       And b.ִ������ <= SYSDATE AND (b.��ֹ���� > SYSDATE OR b.��ֹ���� IS NULL)" & _
                        strWherePriceGrade & vbNewLine & _
                "GROUP BY a.id,a.�Ƿ���"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(.RowData(lngRow)), gstrҩƷ�۸�ȼ�, gstr���ļ۸�ȼ�, gstr��ͨ�۸�ȼ�)
        
        If rsTmp.RecordCount > 0 Then
            .TextMatrix(lngRow, 3) = Trim(rsTmp!�۸�)
        End If
        mblnChange = True
    End With
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub msh��Ŀ_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    With msh��Ŀ(Index)
        If .TxtVisible = False Then Exit Sub
        Select Case .Col
        Case mcstCol�շ���Ŀ
            If IsRecord("������Ŀ", .Text, Index) = False Then
                Cancel = True
                .TxtSetFocus
                Exit Sub
            End If
            .Text = .TextMatrix(.Row, mcstCol�շ���Ŀ)
            If .TextMatrix(.Row, mcstCol���������շ���) = "" Then .TextMatrix(.Row, mcstCol���������շ���) = "100.0"
        Case mcstColԭ��, mcstCol�ּ�, mcstColȱʡ�۸�
            If chk���.value = 1 And gstrҽ�۽ӿڱ�� <> "" And gbln����ҽ���շ���Ŀ = True Then
                Cancel = True
                Exit Sub
            End If
            If NumIsValid(.Text) = False Then
                Cancel = True
                .TxtSetFocus
                Exit Sub
            End If
            .Text = Format(Val(.Text), "###########0.000;-##########0.000;0.000;0.000")
            If .Col = mcstColԭ�� Then
                If .Text = .TextMatrix(.Row, mcstCol�ּ�) Then
                    Cancel = True
                    .TxtSetFocus
                End If
                If chk���.value = 1 And Val(.TextMatrix(.Row, mcstCol�ּ�)) <> 0 Then
                    If Val(.Text) > Val(.TextMatrix(.Row, mcstCol�ּ�)) Then
                        Cancel = True
                        .TxtSetFocus
                    End If
                End If
                If chk���.value = 1 And Val(.TextMatrix(.Row, mcstColȱʡ�۸�)) <> 0 Then
                    If Val(.Text) > Val(.TextMatrix(.Row, mcstColȱʡ�۸�)) Then
                        .TextMatrix(.Row, mcstColȱʡ�۸�) = .Text
                    End If
                End If
            ElseIf .Col = mcstCol�ּ� Then
                If .Text = .TextMatrix(.Row, mcstColԭ��) Then
                    If MsgBox("�����۸���ͬ�ˣ��Ƿ������", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName) = vbNo Then
                        Cancel = True
                        .TxtSetFocus
                        Exit Sub
                    End If
                End If
                If chk���.value = 1 And .TextMatrix(.Row, mcstColԭ��) <> "" Then
                    If Val(.Text) < Val(.TextMatrix(.Row, mcstColԭ��)) Then
                        Cancel = True
                        .TxtSetFocus
                    End If
                End If
                If chk���.value = 1 And Val(.TextMatrix(.Row, mcstColȱʡ�۸�)) <> 0 Then
                    If Val(.Text) < Val(.TextMatrix(.Row, mcstColȱʡ�۸�)) Then
                        .TextMatrix(.Row, mcstColȱʡ�۸�) = .Text
                    End If
                End If
            ElseIf .Col = mcstColȱʡ�۸� Then
                If Val(.Text) <> 0 Then
                    If Val(.Text) < Val(.TextMatrix(.Row, mcstColԭ��)) Or Val(.Text) > Val(.TextMatrix(.Row, mcstCol�ּ�)) Then
                        Cancel = True
                        .TxtSetFocus
                    End If
                End If
            End If
'            If chk���.value = 1 And Not Cancel Then
'                If .Col = mcstColԭ�� Then
'                    Me.txtEdit(text����޼�) = .Text
'                ElseIf .Col = mcstCol�ּ� Then
'                    Me.txtEdit(text����޼�) = .Text
'                End If
'            End If
        Case mcstCol���������շ���, mcstCol�Ӱ�Ӽ���
            If NumIsValid(.Text) = False Then
                Cancel = True
                Exit Sub
            End If
            .Text = Format(Val(.Text), "###########0.0;-##########0.0;0.0;0.0")
        End Select
    End With
    If Cancel = False Then mblnChange = True: mblnChanged��Ŀ(Index) = True
End Sub

Private Sub msh����_EnterCell(Row As Long, Col As Long)
    Dim var�б� As Variant
    Dim lngCount As Long
    Dim i As Long
    
    On Error Resume Next
    '��ʾ�ϼ�
    Me.lbl�����ϼ�.Tag = 0
    For i = 0 To msh����.Rows - 1
        Me.lbl�����ϼ�.Tag = Me.lbl�����ϼ�.Tag + Val(msh����.TextMatrix(i, 1)) * Val(msh����.TextMatrix(i, 3))
    Next
    Me.lbl�����ϼ�.Caption = "�ϼ�:" & Format(Me.lbl�����ϼ�.Tag, "0.00")
    On Error GoTo 0
    '���ù̶���ϵ
    var�б� = Split(mstr�б�(Col + 1), ";")
    msh����.Clear
    For lngCount = LBound(var�б�) To UBound(var�б�)
        msh����.AddItem var�б�(lngCount)
    Next
    If msh����.ListCount = 0 Or Row = 0 Then Exit Sub
    If Row > 1 And msh����.TextMatrix(Row - 1, Col) <> "" Then
        If msh����.TextMatrix(Row, Col) = "" Then msh����.TextMatrix(Row, Col) = msh����.TextMatrix(Row - 1, Col)
    Else
        msh����.ListIndex = 0
    End If
End Sub

Private Sub msh����_KeyDown(KeyCode As Integer, Shift As Integer, Cancel As Boolean)
    Dim i As Long
    Dim strTmp As String
    
    If KeyCode <> vbKeyReturn Then Exit Sub
    With msh����
        If msh����.TxtVisible = False And .CboVisible = False Then
            If msh����.Col = 0 And msh����.TextMatrix(msh����.Row, 0) = "" Then cmdOK.SetFocus
            Exit Sub
        End If
        Select Case msh����.Col
        Case 0
            If IsRecord("�շ���ĿĿ¼", .Text) = False Then
                Cancel = True
                .TxtSetFocus
                Exit Sub
            End If
            .Text = .TextMatrix(.Row, 0)
        Case 1
            If NumIsValid(.Text) = False Then
                Cancel = True
                .TxtSetFocus
                Exit Sub
            End If
            strTmp = .TextMatrix(.Row, 1)
            If .TextMatrix(.Row, 2) <> "0-���̶�" And Val(.Text) = 0 Then
                .Text = strTmp
                Exit Sub
            End If
            
            .Text = Val(.Text)
            Me.lbl�����ϼ�.Tag = 0
            For i = 0 To msh����.Rows - 1
                If IsNumeric(msh����.TextMatrix(i, 1)) And IsNumeric(msh����.TextMatrix(i, 3)) Then
                    Me.lbl�����ϼ�.Tag = Me.lbl�����ϼ�.Tag + Val(msh����.TextMatrix(i, 1)) * Val(msh����.TextMatrix(i, 3))
                End If
            Next
            Me.lbl�����ϼ�.Caption = "�ϼ�:" & Format(Me.lbl�����ϼ�.Tag, "0.00")
        Case 2
            If .TextMatrix(.Row, 2) <> "0-���̶�" And Val(.TextMatrix(.Row, 1)) = 0 Then
                .TextMatrix(.Row, 1) = "1"
            End If
        End Select
    End With
    If Cancel = False Then mblnChange = True
End Sub

Private Sub optApply_Click(Index As Integer)
    Dim i As Integer
    
    mblnChange = True
    For i = 1 To optApply.UBound
        If i = Index Then
            optApply(i).FontBold = True
        Else
            optApply(i).FontBold = False
        End If
    Next
End Sub

Private Sub optApply_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub opt����_Click(Index As Integer)
    Dim sngLeft As Single
    
    '101736,����������סԺ�����������ֹ�����ȱʡִ�п���
    lblһ�����.Caption = "1����ָ�����˿����⣺"
    txt����ִ��.Enabled = False: txt����ִ��.BackColor = &H8000000F: txt����ִ��.Text = "": txt����ִ��.Tag = ""
    txtסԺִ��.Enabled = False: txtסԺִ��.BackColor = &H8000000F: txtסԺִ��.Text = "": txtסԺִ��.Tag = ""
    msf����ִ��.Active = False: msf����ִ��.Enabled = False
    msf����ִ��.BackColorBkg = &H8000000F: msf����ִ��.ClearBill
    
    If Index = 4 Then
        txt����ִ��.Enabled = True: txt����ִ��.BackColor = &H80000005
        txtסԺִ��.Enabled = True: txtסԺִ��.BackColor = &H80000005
        Select Case mstrServerObj
            Case "1"
                txtסԺִ��.Enabled = False: txt����ִ��.BackColor = &H8000000F
            Case "2"
                txt����ִ��.Enabled = False: txtסԺִ��.BackColor = &H8000000F
        End Select
        msf����ִ��.Active = True: msf����ִ��.Enabled = True
        msf����ִ��.BackColorBkg = &H80000005
        
        '2010-05-10 �����ִ�п�����ʾ
        Ini���ʷ���
        load���ʷ��� 0
    ElseIf Index = 0 Then
        '����ȷִ�п���
        If mstrServerObj <> "1" Then
            lblһ�����.Caption = "1���ֹ�����ȱʡִ�п������ã�"
            txtסԺִ��.Enabled = True: txtסԺִ��.BackColor = &H80000005
        End If
    End If
End Sub

Private Sub opt����_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        KeyAscii = 0
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub SetCodeNO()
    '���ñ���
    On Error GoTo ErrHandle
    Dim strSQL  As String
    Dim rsTmp As New ADODB.Recordset
    Dim lngMaxLen As Long
    Dim strTmp As String
    Dim strTmp1 As String
    
    '����������Ҫ����������Ŀ����
    If medit��ʽ = EditNew Or medit��ʽ = EditCopy Then
        '�ȵõ�����������󳤶�
        lngMaxLen = 2
        
        If mstr������� = "" Then
            strSQL = "select max(length(����)) from �շ���ĿĿ¼ where " & IIF(Trim(mstr����ID) = "" Or Trim(mstr����ID) = "0", " ����id is null ", "  ����id=[1] ")
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mstr����ID))
        
            If rsTmp.RecordCount > 0 Then
                lngMaxLen = Nvl(rsTmp(0), 2)
            End If
            
            '��ͨ��GetMax�õ��������ı���+1
            strTmp = sys.MaxCode("�շ���ĿĿ¼", "����", lngMaxLen, " where ����id=" & mstr����ID)
        
            strTmp1 = String(lngMaxLen, "0")
            RSet strTmp1 = strTmp
            strTmp = Replace(strTmp1, " ", "0")
            '�жϸ÷�����û��û��Ŀ�����û�о�Ӧ��ʼ���Ϸ������
            strSQL = "select count(*) ��Ŀ�� from �շ���ĿĿ¼ where ����id=[1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mstr����ID))
            
            If Nvl(rsTmp!��Ŀ��, 0) > 0 Then
                txtEdit(text����).Text = strTmp
            Else
                txtEdit(text����).Text = mstr������� & strTmp
            End If
        Else
            '�������������루������������
            strSQL = "select max(����) as ������ from �շ���ĿĿ¼ where ����id=[1] And ���� Like [2] And Length(����) > Length([3])"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mstr����ID), mstr������� & "%", mstr�������)
            
            If Nvl(rsTmp!������, "") = "" Then
                strTmp = mstr������� & "01"
            Else
                strTmp = zlStr.Increase(rsTmp!������)
            End If
            
            '��������������Ƿ���ڱ���������Ʊ��루��Ҫ��������Ŀ�ı������ɵģ�
            strSQL = "select max(����) as ������ from �շ���ĿĿ¼ where ����id<>[1] And ���� Like [2] And Length(����) > Length([3])"
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mstr����ID), mstr������� & "%", mstr�������)
            
            If Nvl(rsTmp!������, "") <> "" Then
                '����������ڱȱ��������ı���
                If strTmp <= rsTmp!������ Then
                    strTmp = zlStr.Increase(rsTmp!������)
                End If
            End If
            
            txtEdit(text����).Text = strTmp
        End If
        
        mstr���� = txtEdit(text����).Text
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub cmbClass_Click()
    On Error GoTo ErrHandle
    Dim strClass As String
    Dim i As Integer
    
    If Trim(cmbClass.Text) = "" Or InStr(cmbClass.Text, "-") < 1 Then Exit Sub
    Me.chk���.Visible = True
    
    txtEdit(Text���).BackColor = RGB(255, 255, 255)
    Me.chk���ηѱ�.Visible = True
    Me.chk�Ӱ�Ӽ�.Visible = True
    Me.chkժҪ.Visible = True
    Me.chk����.Visible = False
    Me.cmb����.Visible = False
    Me.chk�Զ�����.Visible = False
    lblEdit(10).Visible = False
    cmb��Ŀ����.Visible = False
    
    '���
    Me.lblEdit(4).Enabled = True
    Me.txtEdit(2).Enabled = True
    '���÷������
    '�������
    Me.lblEdit(6).Enabled = True
    Me.cmb�������.Enabled = True
    If Me.cmb�������.ListCount > 3 Then
        Me.cmb�������.ListIndex = 3
    End If
    '��������
    Me.lblEdit(7).Enabled = True
    If InStr(1, mstrPrivs, ";ҽ������;") = 0 Then
        cmb��������.Enabled = False
    Else
        cmb��������.Enabled = True
    End If
    
    '�õ����ͱ���
    If mblnEditCancel = False Then
        mstr��� = Left(Me.cmbClass.Text, 1)
    End If
    
    If TabExist("�շѼ�Ŀ") = True Then
        For i = 0 To tbPriceGrade.ItemCount - 1
            If mstr��� = "F" Then
                msh��Ŀ(i).TextMatrix(0, mcstCol���������շ���) = "���������շ���"
                If msh��Ŀ(i).ColWidth(mcstCol���������շ���) = 0 Then
                   msh��Ŀ(i).ColWidth(mcstCol���������շ���) = 1500
                End If
            Else
                msh��Ŀ(i).ColWidth(mcstCol���������շ���) = 0
            End If
        Next
    End If
    
    '���ñ���
    Call SetCodeNO
    '�ϸ�������ǲ�����ȷ
    strClass = cmbClass.Text
    strClass = Trim(zlStr.NeedName(strClass))
    
    '��ʾ��ǰӦ���ĸ����
    Me.optApply(3).Caption = "Ӧ����" & IIF(mblnCanUpdateAll, "", "��Ժ��") & " " & strClass & " �����������Ŀ(&U)"
    '����
    If strClass = "����" Then
        Me.lblEdit(18).Visible = True
        Me.lblEdit(18).Enabled = True
        Me.txtEdit(text����).Visible = True
        Me.txtEdit(text����).Enabled = True
        Call Form_Resize
    Else
        Me.lblEdit(18).Visible = False
        Me.lblEdit(18).Enabled = False
        Me.txtEdit(text����).Visible = False
        Me.txtEdit(text����).Enabled = False
        Call Form_Resize
    End If
    If strClass = "��Ѫ" Then
        lblEdit(10).Visible = True
        cmb��Ŀ����.Visible = True
        If cmb�������.ListCount > 0 Then
            cmb�������.ListIndex = 0
        End If
    End If
    If Not (strClass = "�Һ�" Or strClass = "����" Or strClass = "��λ") Then
        Exit Sub
    End If
    '���ý�ֹ�������Ŀ
    Me.chk���.value = 0
    Me.chk���.Visible = False
    Me.chk���ηѱ�.Visible = False
    Me.chk�Ӱ�Ӽ�.value = 0
    Me.chk�Ӱ�Ӽ�.Visible = False
    Me.chkժҪ.Visible = False
    '���
    Me.lblEdit(4).Enabled = False
    Me.txtEdit(2).Enabled = False
    Me.txtEdit(Text���).BackColor = Me.BackColor
    '�������
    Me.lblEdit(6).Enabled = False
    Me.cmb�������.Enabled = False
    If Me.cmb�������.ListCount > 3 Then
        Me.cmb�������.ListIndex = 3
    End If
    Select Case strClass
    Case "�Һ�"
        '��������
        Me.lblEdit(7).Enabled = False
        Me.cmb��������.Enabled = False
        If Me.cmb��������.ListCount > 0 Then
            Me.cmb��������.ListIndex = 0
        End If
        Me.chk����.Visible = True
        If Me.cmb�������.ListCount > 1 Then
            Me.cmb�������.ListIndex = 1
        End If
        Exit Sub
    Case "����"
        Me.cmb����.Visible = True
        Me.chk�Զ�����.Visible = (Me.cmb����.ListIndex <> 0)
        If InStr(1, mstrPrivs, ";ҽ������;") = 0 Then
            cmb��������.Enabled = False
        End If
        Exit Sub
    Case "��λ"
        '���
        Me.lblEdit(4).Enabled = True
        Me.txtEdit(Text���).Enabled = True
        txtEdit(Text���).BackColor = RGB(255, 255, 255)
        If InStr(1, mstrPrivs, ";ҽ������;") = 0 Then
            cmb��������.Enabled = False
        End If
        Exit Sub
    End Select
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub


Private Sub txtEdit_Change(Index As Integer)
    Dim i As Integer
    
On Error GoTo ErrHandle
    mblnChange = True
    
    Select Case Index
    Case Text����
        Dim strTmp As String
        '���¼�����ƣ���ȥ �������ַ�
        strTmp = MoveSpecialChar(txtEdit(Text����).Text)
        If txtEdit(Text����).Text <> strTmp Then
            txtEdit(Text����).Text = strTmp
            Me.txtEdit(text����).Text = zlStr.GetCodeByORCL(strTmp, False, mlng���볤��)
            Me.txtEdit(text���).Text = zlStr.GetCodeByORCL(strTmp, True, mlng���볤��)
        End If
        txtEdit(text����).Text = zlStr.GetCodeByORCL(txtEdit(Text����).Text, False, mlng���볤��)
        txtEdit(text���).Text = zlStr.GetCodeByORCL(txtEdit(Text����).Text, True, mlng���볤��)
    Case text��ʶ����, Text��ʶ����
        txtEdit(Index).Text = UCase(txtEdit(Index).Text)
        txtEdit(Index).SelStart = Len(txtEdit(Index).Text)
    Case Text��ѡ��
        txtEdit(Index).SelStart = Len(txtEdit(Index).Text)
    Case Text����
        '����������Ҫ����������Ŀ����
        Call SetCodeNO
    Case text����޼�, text����޼�
        If IsNumeric(txtEdit(Index).Text) Then
            If text����޼� = Index Then
                mdbl����޼� = Val(txtEdit(Index).Text)
                If chk���.value = 1 Then
                    For i = 0 To tbPriceGrade.ItemCount - 1
                        msh��Ŀ(i).TextMatrix(1, mcstCol�ּ�) = Format(mdbl����޼�, "0.000")
                    Next
                End If
            Else
                mdbl����޼� = Val(txtEdit(Index).Text)
                If chk���.value = 1 Then
                    For i = 0 To tbPriceGrade.ItemCount - 1
                        msh��Ŀ(i).TextMatrix(1, mcstColԭ��) = Format(mdbl����޼�, "0.000")
                    Next
                End If
            End If
        End If
    End Select
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtEdit_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
    mstrSel = 0
    Select Case Index
    Case Text����, Text˵��
        OS.OpenIme True
    Case text����, text����, Text����ʱ��, text����޼�, text����޼�
        OS.OpenIme False
    Case Text����
        mstrSel = 1
    Case text��ʶ����
        OS.OpenIme False
        mstrSel = 2
    End Select
End Sub

Private Sub InitLvwSel()
    '��ʼ��lvwSel�ؼ�
    lvwSel.View = lvwReport
    lvwSel.Visible = False
    lvwSel.GridLines = True
    lvwSel.FullRowSelect = True
    lvwSel.Width = 5000
    zlControl.LvwSelectColumns lvwSel, "����,1000,0,2;����,1500,0,2", True
    Select Case True
        Case mstrSel = 1
            lvwSel.Top = txtEdit(Text����).Top + txtEdit(Text����).Height + Screen.TwipsPerPixelY * 1
            lvwSel.Left = txtEdit(Text����).Left
            lvwSel.Height = 1635
            lvwSel.Width = txtEdit(Text����).Width
        Case mstrSel = 2
            lvwSel.Top = txtEdit(text��ʶ����).Top + txtEdit(text��ʶ����).Height + Screen.TwipsPerPixelY * 1
            lvwSel.Left = txtEdit(text��ʶ����).Left
            lvwSel.Width = 3200
            lvwSel.Height = 2500
    End Select
    lvwSel.Tag = False
    zlControl.LvwFlatColumnHeader lvwSel
End Sub

Private Sub lvwSel_LostFocus()
    lvwSel.Visible = False
    If mstrSel = 1 Then txtEdit(Text����).SetFocus
    If mstrSel = 2 Then txtEdit(text��ʶ����).SetFocus
End Sub

Private Sub lvwSel_DblClick()
    lvwSel_KeyPress 13
End Sub

Private Sub lvwSel_KeyPress(KeyAscii As Integer)
    Dim strSQL As String
    Dim i As Long
    Dim rsTmp As New ADODB.Recordset
    
    '�ؼ�ѡ����
    On Error GoTo ErrHandle
    Select Case KeyAscii
    Case 13, Asc(" ")
        Select Case True
            Case mstrSel = 1
                If Not lvwSel.SelectedItem Is Nothing Then
                    mstr����ID = lvwSel.SelectedItem.Tag
                    mstr������� = lvwSel.SelectedItem.Text
                    txtEdit(Text����).Text = lvwSel.SelectedItem.SubItems(1)
                    lvwSel.Visible = False
                    txtEdit(Text����).SetFocus
                End If
                OS.PressKey vbKeyTab
            Case mstrSel = 2
                If Not lvwSel.SelectedItem Is Nothing Then
                    '�ȼ���ǲ������ظ�
                    If medit��ʽ <> EditNew And IsNumeric(mstrID) Then
                        strSQL = " SELECT ����,���� FROM  �շ���ĿĿ¼ WHERE UPPER(��ʶ����) = [1] AND ID<>[2] "
                    Else
                        strSQL = " SELECT ����,���� FROM  �շ���ĿĿ¼ WHERE UPPER(��ʶ����) = [1] "
                    End If
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, lvwSel.SelectedItem.Text, Val(mstrID))
                    
                    If rsTmp.RecordCount > 0 Then
                        strSQL = ""
                        rsTmp.MoveFirst
                        For i = 1 To rsTmp.RecordCount
                            If i = rsTmp.RecordCount Then
                                strSQL = strSQL & "[" & Nvl(rsTmp!����) & "]" & Nvl(rsTmp!����)
                            Else
                                strSQL = strSQL & "[" & Nvl(rsTmp!����) & "]" & Nvl(rsTmp!����) & vbCrLf
                            End If
                            rsTmp.MoveNext
                        Next
                        MsgBox "��Ŀ����" & strSQL & "���Ѿ�ʹ�øñ�׼�۸�", vbInformation, gstrSysName
                        Exit Sub
                    End If
                    '��ʼȡ�Ǹ��۸���Ŀ
                    txtEdit(text��ʶ����).Text = lvwSel.SelectedItem.Text
                    strSQL = "select ��Ŀ����, ��Ŀ����, ƴ����, ��Ŀ����, �Ƽ۵�λ, ��Ŀ�ں�, ��������, ��Ŀ˵��, ��Ŀ�۸�, �ظ���־, ҽԺ�ȼ�, ע����־, �������, ����޼�, ����޼�, �������� from ��׼ҽ�۹淶 where nvl(ע����־,0) =0 and  ��Ŀ���� = [1] "
                    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, txtEdit(text��ʶ����).Text)
                    
                    If rsTmp.RecordCount = 1 Then
                        txtEdit(text��ʶ����).Text = Nvl(rsTmp!��Ŀ����)
                        If medit��ʽ = EditNew Then
                            '����
                            txtEdit(Text����).Text = Nvl(rsTmp!��Ŀ����)
                            txtEdit(text����).Text = zlStr.GetCodeByORCL(txtEdit(Text����).Text, False, mlng���볤��)
                            txtEdit(text���).Text = zlStr.GetCodeByORCL(txtEdit(Text����).Text, True, mlng���볤��)
                            '��λ
                            If Nvl(rsTmp!�Ƽ۵�λ) <> "" Then
                                cmb���㵥λ.Text = Nvl(rsTmp!�Ƽ۵�λ)
                            End If
                            '����
                            If mshAlias.Rows > 2 And Trim(mshAlias.TextMatrix(mshAlias.Rows - 1, 0)) <> "" Then
                                mshAlias.Rows = mshAlias.Rows + 1
                            End If
                            mshAlias.TextMatrix(mshAlias.Rows - 1, 0) = Nvl(rsTmp!��Ŀ����)
                            mshAlias.TextMatrix(mshAlias.Rows - 1, 1) = zlStr.GetCodeByORCL(Nvl(rsTmp!��Ŀ����), False, mlng���볤��)
                            mshAlias.TextMatrix(mshAlias.Rows - 1, 2) = zlStr.GetCodeByORCL(Nvl(rsTmp!��Ŀ����), True, mlng���볤��)
                            '���������޼�
                            txtEdit(text����޼�).Text = Format(Nvl(rsTmp!����޼�, 0), "0.00")
                            txtEdit(text����޼�).Text = Format(Nvl(rsTmp!����޼�, 0), "0.00")
                            If chk���.value = 1 Then
                                For i = 0 To tbPriceGrade.ItemCount - 1
                                    msh��Ŀ(i).Rows = 2
                                    msh��Ŀ(i).TextMatrix(1, mcstCol�ּ�) = txtEdit(text����޼�).Text
                                    msh��Ŀ(i).TextMatrix(1, mcstColԭ��) = txtEdit(text����޼�).Text
                                Next
                            End If
                            mdbl����޼� = Format(Nvl(rsTmp!����޼�, 0), "0.00")
                            mdbl����޼� = Format(Nvl(rsTmp!����޼�, 0), "0.00")
                            '��Ŀ�۸�
                            mdblҽ�ۼ۸� = Nvl(rsTmp!��Ŀ�۸�, 0)
                        ElseIf medit��ʽ = EditModify Then
                            '���������޼�
                            txtEdit(text����޼�).Text = Format(Nvl(rsTmp!����޼�, 0), "0.00")
                            txtEdit(text����޼�).Text = Format(Nvl(rsTmp!����޼�, 0), "0.00")
                            mdbl����޼� = Format(Nvl(rsTmp!����޼�, 0), "0.00")
                            mdbl����޼� = Format(Nvl(rsTmp!����޼�, 0), "0.00")
                            
                            If Not mblnShow�շѼ�Ŀ Then
                                TabMain.Tabs.Add , "_�շѼ�Ŀ", "�շѼ�Ŀ"
                                mblnShow�շѼ�Ŀ = True
                            End If
                            Call init��Ŀ
                            MsgBox "������ȷ���շѼ�Ŀ��", vbInformation, gstrSysName
                        End If
                        OS.PressKey vbKeyTab
                        txtEdit(Text��ʶ����).SetFocus
                    End If
                End If
        End Select
    Case vbKeyEscape
        lvwSel.Visible = False
    End Select
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub


Private Sub txtEdit_KeyDown(Index As Integer, KeyCode As Integer, Shift As Integer)
    On Error GoTo ErrHandle
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim blnMatching As Boolean
    Dim i As Long
    Dim ObjItem As ListItem
    
    
    blnMatching = IIF(GetSetting("ZLSOFT", "����ģ��\����", "����ƥ��", "0") = "0", True, False)
    Select Case Index
    Case Text����   '����
        If KeyCode = 13 Then
            KeyCode = 0
            strSQL = "Select ID,����,���� From �շѷ���Ŀ¼ Where Upper(����) Like [1] or  Upper(����) Like [2] Or Upper(Zlspellcode(����)) Like [1] "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, IIF(blnMatching = True, "%", "") & UCase(txtEdit(Text����).Text) & "%", UCase(txtEdit(Text����).Text) & "%")
            
            If rsTmp.RecordCount = 1 Then
                txtEdit(Text����).Text = Nvl(rsTmp!����)
                txtEdit(text����).Text = Nvl(rsTmp!����)
                mstr������� = Nvl(rsTmp!����)
                mstr����ID = rsTmp!ID
                Call SetCodeNO
                txtEdit(text����).MaxLength = mlng���볤��
                OS.PressKey vbKeyTab
            ElseIf rsTmp.RecordCount > 0 Then
                rsTmp.MoveFirst
                lvwSel.ListItems.Clear
                '��ʼ��ѡ����
                Call InitLvwSel
                For i = 1 To rsTmp.RecordCount
                    Set ObjItem = lvwSel.ListItems.Add(, , Nvl(rsTmp!����))
                    ObjItem.SubItems(1) = Nvl(rsTmp!����)
                    ObjItem.Tag = rsTmp!ID
                    rsTmp.MoveNext
                Next
                lvwSel.ListItems(1).Selected = True
                lvwSel.SelectedItem.EnsureVisible
                lvwSel.Visible = True
                lvwSel.Enabled = True
                lvwSel.ZOrder
                lvwSel.SetFocus
            ElseIf Trim(txtEdit(Text����).Text) = "��" Then
                txtEdit(Text����).Text = "��"
                mstr����ID = "0"
            Else
                strSQL = "Select ID,����,���� From �շѷ���Ŀ¼ Where Nvl(����,'') = [1]"
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, txtEdit(Text����).Text)
                
                If rsTmp.RecordCount > 0 Then
                    txtEdit(Text����).Text = Nvl(rsTmp!����)
                    txtEdit(text����).Text = Nvl(rsTmp!����)
                    mstr����ID = rsTmp!ID
                    Call SetCodeNO
                    txtEdit(text����).MaxLength = mlng���볤��
                    OS.PressKey vbKeyTab
                Else
                    mstr����ID = 0
                    txtEdit(Text����).Text = ""
                End If
            End If
        End If
    Case text��ʶ����    '��ʶ����
        If KeyCode = 13 And txtEdit(text��ʶ����).Text = txtEdit(text��ʶ����).Tag Then
            OS.PressKey vbKeyTab
            Exit Sub
        End If
        If KeyCode = 13 And gstrҽ�۽ӿڱ�� <> "" And gbln����ҽ���շ���Ŀ = True Then
            txtEdit(text����޼�).Enabled = False
            txtEdit(text����޼�).Enabled = False
            '�ȼ���ǲ������ظ�
            If medit��ʽ <> EditNew And IsNumeric(mstrID) Then
                strSQL = " SELECT ����,���� FROM  �շ���ĿĿ¼ WHERE UPPER(��ʶ����) = [1] AND ID<>[2] "
            Else
                strSQL = " SELECT ����,���� FROM  �շ���ĿĿ¼ WHERE UPPER(��ʶ����) = [1] "
            End If
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, txtEdit(Index).Text, Val(mstrID))
            
            If rsTmp.RecordCount > 0 Then
                strSQL = ""
                rsTmp.MoveFirst
                For i = 1 To rsTmp.RecordCount
                    If i = rsTmp.RecordCount Then
                        strSQL = strSQL & "[" & Nvl(rsTmp!����) & "]" & Nvl(rsTmp!����)
                    Else
                        strSQL = strSQL & "[" & Nvl(rsTmp!����) & "]" & Nvl(rsTmp!����) & vbCrLf
                    End If
                    rsTmp.MoveNext
                Next
                MsgBox "��Ŀ����" & strSQL & "���Ѿ�ʹ�øñ�׼�۸�", vbInformation, gstrSysName
                zlControl.TxtSelAll txtEdit(Index)
                Exit Sub
            End If
            'ȡ�Ǹ��۸���Ŀ
            strSQL = "select ��Ŀ����, ��Ŀ����, ƴ����, ��Ŀ����, �Ƽ۵�λ, ��Ŀ�ں�, ��������," & vbNewLine & _
                    "        ��Ŀ˵��, ��Ŀ�۸�, �ظ���־, ҽԺ�ȼ�, ע����־, �������, ����޼�, ����޼�, ��������" & vbNewLine & _
                    " from ��׼ҽ�۹淶" & vbNewLine & _
                    " where nvl(ע����־,0) =0 and  upper(��Ŀ����) like [1] or upper(��Ŀ����) LIKE [2] or upper(ƴ����) LIKE [2] "
            Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Trim(UCase(txtEdit(Index))) & "%", IIF(blnMatching = True, "%", "") & Trim(UCase(txtEdit(Index))) & "%")
            
            If rsTmp.RecordCount = 1 Then
                txtEdit(Index).Text = Nvl(rsTmp!��Ŀ����)
                If medit��ʽ = EditNew Then
                    '����
                    txtEdit(Text����).Text = Nvl(rsTmp!��Ŀ����)
                    txtEdit(text����).Text = zlStr.GetCodeByORCL(txtEdit(Text����).Text, False, mlng���볤��)
                    txtEdit(text���).Text = zlStr.GetCodeByORCL(txtEdit(Text����).Text, True, mlng���볤��)
                    '��λ
                    If Nvl(rsTmp!�Ƽ۵�λ) <> "" Then
                        cmb���㵥λ.Text = Nvl(rsTmp!�Ƽ۵�λ)
                    End If
                    '����
                    If mshAlias.Rows > 2 And Trim(mshAlias.TextMatrix(mshAlias.Rows - 1, 0)) <> "" Then
                        mshAlias.Rows = mshAlias.Rows + 1
                    End If
                    mshAlias.TextMatrix(mshAlias.Rows - 1, 0) = Nvl(rsTmp!��Ŀ����)
                    mshAlias.TextMatrix(mshAlias.Rows - 1, 1) = zlStr.GetCodeByORCL(Nvl(rsTmp!��Ŀ����), False, mlng���볤��)
                    mshAlias.TextMatrix(mshAlias.Rows - 1, 2) = zlStr.GetCodeByORCL(Nvl(rsTmp!��Ŀ����), True, mlng���볤��)
                    '���������޼�
                    txtEdit(text����޼�).Text = Format(Nvl(rsTmp!����޼�, 0), "0.00")
                    txtEdit(text����޼�).Text = Format(Nvl(rsTmp!����޼�, 0), "0.00")
                    If chk���.value = 1 Then
                        For i = 0 To tbPriceGrade.ItemCount - 1
                            msh��Ŀ(i).Rows = 2
                            msh��Ŀ(i).TextMatrix(1, mcstCol�ּ�) = txtEdit(text����޼�).Text
                            msh��Ŀ(i).TextMatrix(1, mcstColԭ��) = txtEdit(text����޼�).Text
                        Next
                    End If
                    mdbl����޼� = Format(Nvl(rsTmp!����޼�, 0), "0.00")
                    mdbl����޼� = Format(Nvl(rsTmp!����޼�, 0), "0.00")
                    '��Ŀ�۸�
                    mdblҽ�ۼ۸� = Nvl(rsTmp!��Ŀ�۸�, 0)
                ElseIf medit��ʽ = EditModify Then
                    '���������޼�
                    txtEdit(text����޼�).Text = Format(Nvl(rsTmp!����޼�, 0), "0.00")
                    txtEdit(text����޼�).Text = Format(Nvl(rsTmp!����޼�, 0), "0.00")
                    If chk���.value = 1 Then
                        For i = 0 To tbPriceGrade.ItemCount - 1
                            msh��Ŀ(i).Rows = 2
                            msh��Ŀ(i).TextMatrix(1, mcstCol�ּ�) = txtEdit(text����޼�).Text
                            msh��Ŀ(i).TextMatrix(1, mcstColԭ��) = txtEdit(text����޼�).Text
                        Next
                    End If
                    mdbl����޼� = Format(Nvl(rsTmp!����޼�, 0), "0.00")
                    mdbl����޼� = Format(Nvl(rsTmp!����޼�, 0), "0.00")
                End If
                If medit��ʽ = EditModify Then
                    If Not mblnShow�շѼ�Ŀ Then
                        TabMain.Tabs.Add , "_�շѼ�Ŀ", "�շѼ�Ŀ"
                        mblnShow�շѼ�Ŀ = True
                    End If
                    Call init��Ŀ
                    MsgBox "������ȷ���շѼ�Ŀ��", vbInformation, gstrSysName
                End If
                
                OS.PressKey vbKeyTab
            ElseIf rsTmp.RecordCount > 1 Then
                KeyCode = 0
                lvwSel.ListItems.Clear
                '��ʼ��ѡ����
                Call InitLvwSel
                For i = 1 To rsTmp.RecordCount
                    Set ObjItem = lvwSel.ListItems.Add(, , Nvl(rsTmp!��Ŀ����))
                    ObjItem.SubItems(1) = Nvl(rsTmp!��Ŀ����)
                    ObjItem.Tag = Nvl(rsTmp!��Ŀ����)
                    rsTmp.MoveNext
                Next
                lvwSel.ListItems(1).Selected = True
                lvwSel.SelectedItem.EnsureVisible
                lvwSel.Visible = True
                lvwSel.Enabled = True
                lvwSel.ZOrder
                lvwSel.SetFocus
            Else
                KeyCode = 0
                MsgBox "�����ڵġ���ʶ���롱��", vbInformation, gstrSysName
                zlControl.TxtSelAll txtEdit(Index)
            End If
        End If
    End Select
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txtEdit_LostFocus(Index As Integer)
    Select Case Index
    Case 1, 4, 8
        OS.OpenIme False
    End Select
End Sub

Private Sub txtEdit_Validate(Index As Integer, Cancel As Boolean)
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo ErrHandle
    mblnEditCancel = True
    Select Case True
    Case Index = text���� Or Index = text���� Or Index = text��� Or Index = Text���� Or _
        Index = Text���� Or Index = Text����ʱ�� Or Index = Text��� Or _
        Index = Text��ѡ�� Or Index = Text��ʶ���� Or Index = text����޼� Or Index = text����޼�
'        ShowTab "������Ϣ"
        If Index = text����޼� Or Index = text����޼� Then
            If Trim(txtEdit(Index).Text) = "" Then txtEdit(Index).Text = 0
            If IsNumeric(txtEdit(Index).Text) = False Then
                Cancel = True
                mblnEditCancel = False
                MsgBox "������һ���Ϸ��ļ۸�", vbInformation, gstrSysName
                Exit Sub
            End If
            If Val(txtEdit(text����޼�).Text) <> 0 And Val(txtEdit(text����޼�).Text) < Val(txtEdit(text����޼�).Text) Then
                MsgBox "����޼۱�����ڻ��������޼ۣ�", vbInformation, gstrSysName
                Cancel = True
                mblnEditCancel = False
                Exit Sub
            End If
            '����ּ��Ƿ����޼۳�ͻ
            If Len(Trim(mstrID)) > 0 And (Val(txtEdit(text����޼�).Text) <> 0 Or Val(txtEdit(text����޼�).Text) <> 0) Then
                strSQL = "Select Max(�ּ�) As ����ּ�,Min(�ּ�) As ����ּ� From �շѼ�Ŀ Where" & _
                    " Decode(��ֹ����,to_date('3000-01-01','YYYY-MM-DD'),Null,��ֹ����) is Null And �շ�ϸĿID =" & mstrID
                Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
                If Not rsTmp.EOF Then
                    If Val(txtEdit(text����޼�).Text) <> 0 And Val(Nvl(rsTmp!����ּ�)) > Val(txtEdit(text����޼�).Text) Then Cancel = True
                    If Val(txtEdit(text����޼�).Text) <> 0 And Val(Nvl(rsTmp!����ּ�)) < Val(txtEdit(text����޼�).Text) Then Cancel = True
                    If Cancel Then
                        MsgBox "���м۸��ڵ�ǰ���õ��޼��ڣ�����ۻ����������޼ۣ�", vbInformation, gstrSysName
                        mblnEditCancel = False
                        Exit Sub
                    End If
                End If
            ElseIf Val(txtEdit(text����޼�).Text) <> 0 Or Val(txtEdit(text����޼�).Text) <> 0 Then
                For i = 0 To tbPriceGrade.ItemCount - 1
                    If Len(Trim(Me.msh��Ŀ(i).TextMatrix(1, mcstCol�ּ�))) > 0 Then
                        If Val(txtEdit(text����޼�).Text) <> 0 And Val(Me.msh��Ŀ(i).TextMatrix(1, mcstCol�ּ�)) > Val(txtEdit(text����޼�).Text) Then Cancel = True
                        If Val(txtEdit(text����޼�).Text) <> 0 And Val(Me.msh��Ŀ(i).TextMatrix(1, mcstCol�ּ�)) < Val(txtEdit(text����޼�).Text) Then Cancel = True
                        If Cancel Then
                            MsgBox tbPriceGrade(i).Caption & "���м۸�" & Format(Me.msh��Ŀ(i).TextMatrix(1, mcstCol�ּ�), "#0.000") & "�����ڵ�ǰ���õ��޼��ڣ�����ۻ����������޼ۣ�", vbInformation, gstrSysName
                            mblnEditCancel = False
                            Exit Sub
                        End If
                    End If
                Next
            End If
        End If
    Case Index = Text˵��
'        ShowTab "�շѼ�Ŀ"
    Case Index = text��ʶ����
        ShowTab "������Ϣ"
        Cancel = Not zlCommFun.StrIsValid(txtEdit(Index).Text, txtEdit(Index).MaxLength)
        If Not Cancel And (gstrҽ�۽ӿڱ�� <> "" And gbln����ҽ���շ���Ŀ) Then
            '��鲻�ǰ����зǷ��ַ�
            If Trim(txtEdit(Index)) = "" Then
                Cancel = True
                MsgBox "����ʶ���롱����Ϊ�գ�", vbInformation, gstrSysName
            Else
                strSQL = "select 1 from ��׼ҽ�۹淶 where ��Ŀ����= [1] "
                Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, txtEdit(Index).Text)
                
                If rsTmp.RecordCount < 1 Then
                    Cancel = True
                    MsgBox "�����ڵġ���ʶ���롱��", vbInformation, gstrSysName
                    txtEdit(text��ʶ����).Text = txtEdit(text��ʶ����).Tag
'                    zlControl.TxtSelAll txtEdit(Index)
                End If
            End If
        End If
        mblnEditCancel = False
        Exit Sub
    End Select
    mblnEditCancel = False
    If Index <> Text��� And Index <> Text˵�� And Index <> text���� Then Cancel = Not zlCommFun.StrIsValid(txtEdit(Index).Text, txtEdit(Index).MaxLength)
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub txtEdit_KeyPress(Index As Integer, KeyAscii As Integer)

    If Index = text������Ŀ Then
         If KeyAscii = vbKeyDelete Then
            txtEdit(text������Ŀ).Text = ""
            Exit Sub
        Else
            KeyAscii = 0
            Exit Sub
         End If
    End If
'    If InStr("~@%^&_|`'""/?", Chr(KeyAscii)) > 0 And _
'        index <> Text��� And index <> Text˵�� And index <> Text���� Then KeyAscii = 0: Exit Sub
    If InStr("'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If InStr(GCST_INVALIDCHAR, Chr(KeyAscii)) > 0 And _
        Index <> Text��� And Index <> Text˵�� And Index <> text���� Then KeyAscii = 0: Exit Sub
    If (Index = 5) And KeyAscii = Asc("*") Then
        KeyAscii = 0
        cmd�ϼ�_Click
        Exit Sub
    ElseIf KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        KeyAscii = 0
        If (mstr��� = "M" And Index = text����) Or (mstr��� <> "M" And Index = Text˵��) Then
            If TabMain.Tabs.Count > 1 Then
                ShowTab "ִ�п���"
            Else
                cmdOK.SetFocus
            End If
        ElseIf Not (Index = Text���� Or Index = text��ʶ����) _
            Or (Index = text��ʶ���� And gstrҽ�۽ӿڱ�� = "") Then
            OS.PressKey vbKeyTab
        End If
    ElseIf Index = Text���� Then
        Select Case KeyAscii
        Case Asc("?")
            KeyAscii = Asc("��")
        Case Asc("%")
            KeyAscii = Asc("��")
        Case Asc("_")
            KeyAscii = Asc("��")
        End Select
        txtEdit(text����).Text = zlStr.GetCodeByORCL(txtEdit(Text����).Text, False, mlng���볤��)
        txtEdit(text���).Text = zlStr.GetCodeByORCL(txtEdit(Text����).Text, True, mlng���볤��)
    ElseIf Index = text���� Or Index = text��� Then
        KeyAscii = Asc(UCase(Chr(KeyAscii)))
    ElseIf Index = Text��ʶ���� Then
        If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZ", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack Then
            KeyAscii = 0
        End If
    ElseIf Index = Text��ѡ�� Then
        If InStr("0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ", Chr(KeyAscii)) < 1 And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack Then
            KeyAscii = 0
        End If
    ElseIf Index = text����޼� Or Index = text����޼� Then
        If InStr("0123456789.", UCase(Chr(KeyAscii))) < 1 And KeyAscii <> vbKeyDelete And KeyAscii <> vbKeyBack Then
            KeyAscii = 0
        End If
    End If
End Sub

Private Sub txtTemp_Change()
    txtEdit(text����).Width = txtTemp.Width - TextWidth(txtTemp.Text) - 120
    txtEdit(text����).Left = txtTemp.Left + TextWidth(txtTemp.Text) + 60
End Sub

Private Sub Form_Activate()
    
    On Error Resume Next
    
    If lblStationNo.Visible = False Then
        lbl������Ŀ.Left = Label1.Left
        txtEdit(text������Ŀ).Left = lbl������Ŀ.Left + lbl������Ŀ.Width + 50
        cmd����.Left = txtEdit(text������Ŀ).Left + txtEdit(text������Ŀ).Width - cmd����.Width
    Else
        lbl������Ŀ.Left = lblEdit(7).Left
        txtEdit(text������Ŀ).Left = lbl������Ŀ.Left + lbl������Ŀ.Width + 50
        cmd����.Left = txtEdit(text������Ŀ).Left + txtEdit(text������Ŀ).Width - cmd����.Width
    End If
    
    Select Case TabMain.SelectedItem.Caption
    Case "������Ϣ"
        If txtEdit(Text����).Enabled And txtEdit(Text����).Visible Then
            txtEdit(Text����).SetFocus
        End If
    Case "�շѼ�Ŀ"
        If msh��Ŀ(0).Visible And msh��Ŀ(0).Active Then
            msh��Ŀ(0).SetFocus
        End If
    Case "������Ŀ"
        If msh����.Visible And msh����.Active Then
            msh����.SetFocus
        End If
    Case "ִ�п���"
        Dim i As Integer
        For i = 0 To 3
            If opt����(i).value = True Then
                If opt����(i).Visible And opt����(i).Enabled Then
                    opt����(i).SetFocus
                End If
                Exit Sub
            End If
        Next
    End Select
End Sub

Private Sub Form_Load()
    '���Ի�����
    chk����.value = Val(GetSetting("ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "����״̬", "0"))
    With Me.msf����ִ��
        .Active = True
        .MsfObj.ScrollBars = flexScrollBarVertical
        .MsfObj.FixedCols = 1: .Rows = 2: .Cols = 4
        .TextMatrix(0, 0) = "ִ�п���ID": .TextMatrix(0, 1) = "ִ�п���"
        .TextMatrix(0, 2) = "���˿���ID": .TextMatrix(0, 3) = "���˿���"
        .ColData(0) = 5: .ColData(1) = 1: .ColData(2) = 5: .ColData(3) = 1
        .ColWidth(0) = 0: .ColWidth(1) = 1550: .ColWidth(2) = 0: .ColWidth(3) = 3600
        .PrimaryCol = 1: .LocateCol = 1
        .Row = 1: .Col = 1
    End With
    
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
    mlngFind = 1
    mblnOk = False
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Dim i As Long
    
    mblnVerifyFlow = False
    mblnVerifyPris = False
    If mblnChange = False Then
        Exit Sub
    End If
    i = MsgBox("�����������˳��Ļ������е��޸Ķ�������Ч��" & vbCrLf & "�Ƿ�ȷ���˳���", vbQuestion Or vbYesNo Or vbDefaultButton2, gstrSysName)
    If i = vbNo Then
        Cancel = 1
    End If
    SaveSetting "ZLSOFT", "˽��ģ��\" & gstrDbUser & "\" & App.ProductName & "\" & Me.Name, "����״̬", chk����.value
End Sub

Private Sub tabMain_Click()
    Dim rsTmp As ADODB.Recordset
    Dim i As Integer
    
    On Error GoTo ErrHandle
    If mblnEditCancel = True Then Exit Sub
    If Val(TabMain.Tag) = TabMain.SelectedItem.Index Then Exit Sub
    TabMain.Tag = TabMain.SelectedItem.Index
    
    fra(1).Visible = False
    fra(2).Visible = False
    fra(3).Visible = False
    fra(4).Visible = False
    On Error Resume Next
    Select Case TabMain.SelectedItem.Caption
    Case "�շѼ�Ŀ"
        If mstrID = "" Then
            If Mid(cmbClass.Text, 1, 1) = "J" Or Mid(cmbClass.Text, 1, 1) = "H" Then
                For i = 0 To tbPriceGrade.ItemCount - 1
                    dtpBegin(i).CustomFormat = "yyyy��MM��dd��"
                    dtpBegin(i).Width = 1600
                    dtpBegin(i).value = DateAdd("d", 1, sys.Currentdate)
                    dtpBegin(i).MinDate = sys.Currentdate
                Next
                mstrCurrentDateFormat = "yyyy-mm-dd"
            Else
                For i = 0 To tbPriceGrade.ItemCount - 1
                    dtpBegin(i).CustomFormat = "yyyy��MM��dd�� HH:mm:ss"
                    dtpBegin(i).Width = 2535
                    dtpBegin(i).value = DateAdd("d", 1, sys.Currentdate)
                    dtpBegin(i).MinDate = sys.Currentdate
                Next
                mstrCurrentDateFormat = "yyyy-mm-dd hh:mm:ss"
            End If
        End If
        fra(2).Visible = True
        fra(2).ZOrder
        If msh��Ŀ(0).Active And msh��Ŀ(0).Visible Then
            msh��Ŀ(0).SetFocus
        End If
        
        '�ڱ༭״̬���������ҽ��ϵͳ����ѡ�����µ�ҽ����Ŀ�����շѼ�Ŀҳ������۸�ֻ��ѡ������ִ�С�
        If medit��ʽ = EditModify And (gstrҽ�۽ӿڱ�� <> "" And gbln����ҽ���շ���Ŀ) Then
            If txtEdit(text��ʶ����).Text <> txtEdit(text��ʶ����).Tag Then
                For i = 0 To tbPriceGrade.ItemCount - 1
                    dtpBegin(i).Enabled = False
                    chkNow(i).value = 1
                Next
            End If
        End If
    Case "ִ�п���"
        fra(3).Visible = True
        fra(3).ZOrder
        
        mstrServerObj = ""
        If medit��ʽ = EditDept Then '����ִ�п���ʱ����ֱ��ͨ���ؼ���ȡ��ǰ�ķ������
            gstrSQL = "select nvl(�������,0) as ������� from �շ���ĿĿ¼ where id=[1]"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, "�������", mstrID)
            If rsTmp.RecordCount > 0 Then
                mstrServerObj = rsTmp!�������
            End If
        Else
            mstrServerObj = Mid(cmb�������.Text, 1, 1)
        End If
        
        For i = 0 To 3
            If opt����(i).value = True And opt����(i).Enabled And opt����(i).Visible Then
                opt����(i).SetFocus
                Exit Sub
            End If
        Next
    Case "������Ŀ"
        fra(4).Visible = True
        fra(4).ZOrder
        '����û���ҵ�������(��ʱ��ô����)
'        If msh����.Active And msh����.Visible Then
'            msh����.SetFocus
'        End If
    Case Else
        fra(1).Visible = True
        fra(1).ZOrder
        If txtEdit(text����).Enabled And txtEdit(text����).Visible Then
            txtEdit(text����).SetFocus
        End If
    End Select
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ClearContext(Optional ByVal bln��ȫ As Boolean = True)
    On Error GoTo ErrHandle
    Dim lngCol As Long, i As Integer
    
    mstrID = ""
    If Trim(mstr����ID) = "0" Then mstr����ID = ""
    '���ñ���
    Call SetCodeNO
    
    If txtEdit(text����).Text = "" Then txtEdit(text����).Text = 1
    mstr���� = txtEdit(text����).Text
    txtEdit(Text����).Text = ""
    txtEdit(text��ʶ����).Text = ""
    txtEdit(Text��ʶ����).Text = ""
    txtEdit(text����޼�).Text = ""
    txtEdit(text����޼�).Text = ""
    txtEdit(text����).Text = ""
    txtEdit(text���).Text = ""
    
    txtEdit(Text���).Text = ""
    mshAlias.Rows = 2
    mshAlias.TextMatrix(1, 0) = "": mshAlias.TextMatrix(1, 1) = "": mshAlias.TextMatrix(1, 2) = ""
    
    For i = 0 To tbPriceGrade.ItemCount - 1
        For lngCol = 1 To mcol��Ŀ(i).Count
            mcol��Ŀ(i).Remove 1
        Next
        For lngCol = 1 To msh��Ŀ(i).Rows - 1
            If msh��Ŀ(i).RowData(lngCol) > 0 Then
                mcol��Ŀ(i).Add 0, "C" & msh��Ŀ(i).RowData(lngCol)
                msh��Ŀ(i).TextMatrix(lngCol, 1) = "0.000"
            End If
        Next
    Next
    
    mshAlias.Col = 0
    msh����.Col = 0
    For i = 0 To tbPriceGrade.ItemCount - 1
        msh��Ŀ(i).Col = 0
    Next
    
    If bln��ȫ = False Then Exit Sub
    txtEdit(Text˵��).Text = ""
    txtEdit(text����).Text = ""
    cmb���㵥λ.Text = ""
    chk���.value = 0
    chk�Ӱ�Ӽ�.value = 0
    chk���ηѱ�.value = 0
    chkժҪ.value = 0
    chk����.value = 0
    
    For i = 0 To tbPriceGrade.ItemCount - 1
        For lngCol = 1 To mcol��Ŀ(i).Count
            mcol��Ŀ(i).Remove 1
        Next
        
        msh��Ŀ(i).ClearBill
        msh��Ŀ(i).Rows = 2
        msh��Ŀ(i).RowData(1) = 0
        For lngCol = 0 To msh��Ŀ(i).Cols - 1
            msh��Ŀ(i).TextMatrix(1, lngCol) = ""
        Next
        txt����˵��(i).Text = ""
    Next
    
    msh����.ClearBill
    msh����.Rows = 2
    msh����.RowData(1) = 0
    For lngCol = 0 To msh����.Cols - 1
        msh����.TextMatrix(1, lngCol) = ""
    Next
    
    opt����(0).value = True
    optApply(0).value = True
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function init����() As Boolean
    On Error GoTo ErrHandle
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long
    Dim strSQL As String
    
    rsTmp.CursorLocation = adUseClient
    rsTmp.CursorType = adOpenKeyset
    rsTmp.LockType = adLockReadOnly
    rsTmp.CursorLocation = adUseClient
    
    mdbl����޼� = 0
    mdbl����޼� = 0
    
    mblnNotClick = True
    Call IniStationNo
    mblnNotClick = False
    
    mshAlias.Cols = 3
    mshAlias.ColAlignment(0) = 1
    mshAlias.ColAlignment(1) = 1
    mshAlias.ColAlignment(2) = 1
    mshAlias.ColWidth(0) = 1800
    mshAlias.ColWidth(1) = 1200
    mshAlias.ColWidth(2) = 1200
    mshAlias.TextMatrix(0, 0) = "����"
    mshAlias.TextMatrix(0, 1) = "ƴ������"
    mshAlias.TextMatrix(0, 2) = "��ʼ���"
    mshAlias.PrimaryCol = 0
    mshAlias.ColData(0) = 4 '�ı���
    mshAlias.ColData(1) = 4 '�ı���
    mshAlias.ColData(2) = 4 '�ı���
    mshAlias.Rows = 2
    mshAlias.Active = True
    
    '��ʼ�����
    strSQL = "select ����,���� from �շ���Ŀ��� where ����<>'4' And ����<>'5' and ����<>'6' and ����<>'7'"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    mblnEditCancel = True
    Me.cmbClass.Clear
    If rsTmp.RecordCount > 0 Then
        rsTmp.MoveFirst
        For i = 1 To rsTmp.RecordCount
            cmbClass.AddItem Nvl(rsTmp!����) & "-" & Nvl(rsTmp!����)
            If i = 1 Then
                cmbClass.ListIndex = 0
            ElseIf Nvl(rsTmp!����) = "E" Then
                cmbClass.ListIndex = cmbClass.NewIndex
            End If
            
            rsTmp.MoveNext
        Next
    End If
    mblnEditCancel = False
    
    txtTemp.Text = ""
    If Trim(mstr����ID) = "0" Or Trim(mstr����ID) = "" Then
        '�����ڵ㣬���������Ϣ
        mstr������� = ""
        txtEdit(Text����).Text = "��"
    Else
        'һ��ڵ㣬ֱ�Ӵ������ж�ȡ
        strSQL = "select ����,���� from �շѷ���Ŀ¼ where ID=[1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mstr����ID))
                
        mstr������� = rsTmp("����")
        txtEdit(Text����).Text = rsTmp("����")
    End If
    
    'ȡ���ϼ����룬�������볤�ȵ�ֵ
    txtTemp.MaxLength = 0
    
    '������Ŀ����
    cmb��Ŀ����.Clear
    cmb��Ŀ����.AddItem "0-ѪҺ"
    cmb��Ŀ����.AddItem "1-����"
    cmb��Ŀ����.ListIndex = 0
    
    cmb�������.Clear
    cmb�������.AddItem "0-��"
    cmb�������.AddItem "1-����"
    cmb�������.AddItem "2-סԺ"
    cmb�������.AddItem "3-������סԺ"
    cmb�������.ListIndex = 3
    
    '���û����������Ŀ
    cmb����.Clear
    cmb����.AddItem "0-һ����Ŀ"
    cmb����.AddItem "1-����ȼ�"
    cmb����.AddItem "2-��������ȼ�"
    cmb����.ListIndex = 0
    
    '���÷���ȷ����Ŀ
    cmb����ȷ��.Clear
    cmb����ȷ��.AddItem "0-����Ҫȷ�ϻ���"
    cmb����ȷ��.AddItem "1-��Ҫȷ�ϻ���"
    cmb����ȷ��.ListIndex = 0
    
    '����¼��������Ӧ�÷�Χ
    With Me.cbo¼��������Χ
        .Clear
        .AddItem "����Ŀ"
        .AddItem "����"
        .AddItem "������"
        .AddItem "�����"
        .AddItem "����"
        .ListIndex = 0
    End With
    
    '���÷�������
    strSQL = "select ����,����,ȱʡ��־ from �������� where ����<>'1' order by ����"
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    
    cmb��������.Clear
    cmb��������.AddItem ""
    Do Until rsTmp.EOF
        cmb��������.AddItem rsTmp("����") & "-" & rsTmp("����")
        If cmb��������.ListIndex = -1 And rsTmp("ȱʡ��־") = 1 Then
            cmb��������.ListIndex = cmb��������.NewIndex
        End If
        rsTmp.MoveNext
    Loop
    If InStr(1, mstrPrivs, "ҽ������") = 0 Then
        cmb��������.Enabled = False
    End If
    'ȡ���ù��ļ��㵥λ
    strSQL = "select distinct ���㵥λ from �շ���ĿĿ¼ where ���=[1] and rownum<500"
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, mstr���)
        
    cmb���㵥λ.Clear
    Do Until rsTmp.EOF
        If Not IsNull(rsTmp("���㵥λ")) Then
            cmb���㵥λ.AddItem rsTmp("���㵥λ")
        End If
        rsTmp.MoveNext
    Loop
    
    chk����.Visible = False
    cmb����.Visible = False
    'txtTemp.MaxLengthΪ0��ʾ�ø��ڵ㻹û���ӽڵ㣬Ҫ��೤�����
    msh����.RowData(1) = 0
    If medit��ʽ <> EditNew And IsNumeric(mstrID) Then
        strSQL = "select A.���,b.���� �������,A.����,A.��ʶ����,A.��ʶ����,a.��ѡ��, A.����,A.���,A.���㵥λ,A.��������,A.��Ŀ����,A.�������" & _
        "    ,A.����ժҪ,A.˵��,A.����,A.���ηѱ�,A.�Ƿ���,A.�Ӱ�Ӽ�,A.����ʱ��,A.����޼�,A.����޼�,A.¼������,A.����ȷ��,A.վ��,A.���㷽ʽ,a.������Ŀ " & _
            " From �շ���ĿĿ¼ A,�շ���Ŀ��� B  " & _
            " Where A.���=B.���� and  A.ID=[1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mstrID))
        
        mblnEditCancel = True
        If Me.cmbClass.ListCount > 0 Then
            For i = 0 To Me.cmbClass.ListCount
                If Me.cmbClass.List(i) = Nvl(rsTmp!���) & "-" & Nvl(rsTmp!�������) Then
                    Me.cmbClass.ListIndex = i
                    Exit For
                End If
            Next
        End If
        mblnEditCancel = False
        If medit��ʽ <> EditCopy Then
            txtEdit(text����).Text = Mid(rsTmp("����"), Len(txtTemp.Text) + 1)
            mstr���� = rsTmp("����")
        Else
            Call SetCodeNO
        End If
        '
        If gstrҽ�۽ӿڱ�� <> "" And gbln����ҽ���շ���Ŀ = True Then
            txtEdit(text����޼�).Enabled = False
            txtEdit(text����޼�).Enabled = False
        Else
            txtEdit(text����޼�).Visible = False
            txtEdit(text����޼�).Visible = False
            lblEdit(20).Visible = False
            lblEdit(21).Visible = False
        End If
        txtEdit(text����޼�).Text = Format(Nvl(rsTmp("����޼�"), 0), "0.00")
        txtEdit(text����޼�).Text = Format(Nvl(rsTmp("����޼�"), 0), "0.00")
        mdbl����޼� = Format(Nvl(rsTmp("����޼�"), 0), "0.00")
        mdbl����޼� = Format(Nvl(rsTmp("����޼�"), 0), "0.00")
        
        '��������ӽڵ����ڵ������
        txtEdit(text��ʶ����).Text = Nvl(rsTmp("��ʶ����"))
        txtEdit(text��ʶ����).Tag = Nvl(rsTmp("��ʶ����"))
        txtEdit(Text��ʶ����).Text = Nvl(rsTmp("��ʶ����"))
        txtEdit(Text��ѡ��).Text = Nvl(rsTmp("��ѡ��"))
        
        txtEdit(Text����).Text = rsTmp("����")
        txtEdit(Text���).Text = Nvl(rsTmp("���"))
        txtEdit(Text˵��).Text = Nvl(rsTmp("˵��"))
        txtEdit(text����).Text = Nvl(rsTmp("����"))
        txtEdit(Text����ʱ��).Text = Format(rsTmp("����ʱ��"), "yyyy-MM-dd")
        
        chk���ηѱ�.value = IIF(rsTmp("���ηѱ�") = 1, 1, 0)
        chk�Ӱ�Ӽ�.value = IIF(rsTmp("�Ӱ�Ӽ�") = 1, 1, 0)
        chk���.Tag = IIF(rsTmp("�Ƿ���") = 1, 1, 0)
        chk���.value = IIF(rsTmp("�Ƿ���") = 1, 1, 0)
        chkժҪ.value = IIF(rsTmp("����ժҪ") = 1, 1, 0)
        txt¼������.Text = IIF(IsNull(rsTmp("¼������")), "", rsTmp("¼������"))
        
        mblnNotClick = True
        cbo.SeekIndex cmbStationNo, Nvl(rsTmp("վ��"))
        cmbStationNo.Tag = Nvl(rsTmp("վ��"))
        mblnNotClick = False
        
        chk�Զ�����.value = IIF(rsTmp("���㷽ʽ") = 1, 1, 0)
        txtEdit(text������Ŀ).Text = IIF(IsNull(rsTmp!������Ŀ), "", rsTmp!������Ŀ)
        
        Select Case rsTmp!���
        Case "1"   '�Һ�
            chk����.value = IIF(rsTmp("��Ŀ����") = 1, 1, 0)
            chk����.Visible = True
            cmb����.Visible = False
            
            chk���.Visible = False
            chk���ηѱ�.Visible = chk���.Visible
            chk�Ӱ�Ӽ�.Visible = chk���.Visible
            chkժҪ.Visible = chk���.Visible
            txtEdit(Text���).Enabled = chk���.Visible
            txtEdit(Text���).BackColor = Me.BackColor
            cmb�������.ListIndex = 1
            cmb�������.Enabled = chk���.Visible
        Case "H"    '����
            If IsNull(rsTmp!��Ŀ����) = False Then
                cmb����.ListIndex = rsTmp!��Ŀ����
            End If
            cmb����.Visible = True
            chk����.Visible = False
            
            chk���.Visible = False
            chk���ηѱ�.Visible = chk���.Visible
            chk�Ӱ�Ӽ�.Visible = chk���.Visible
            chkժҪ.Visible = chk���.Visible
            txtEdit(Text���).Enabled = chk���.Visible
            txtEdit(Text���).BackColor = Me.BackColor
            cmb�������.Enabled = chk���.Visible
        Case "J"    '��λ
            chk���.Visible = False
            chk���ηѱ�.Visible = chk���.Visible
            chk�Ӱ�Ӽ�.Visible = chk���.Visible
            chkժҪ.Visible = chk���.Visible
            txtEdit(Text���).Enabled = True
            cmb�������.Enabled = chk���.Visible
        Case "K" '��Ѫ
            If IsNull(rsTmp!��Ŀ����) = True Then
                cmb��Ŀ����.ListIndex = 0
            Else
                cmb��Ŀ����.ListIndex = rsTmp!��Ŀ����
            End If
        End Select
        
        cmb���㵥λ.Text = IIF(IsNull(rsTmp("���㵥λ")), "", rsTmp("���㵥λ"))
        cmb�������.ListIndex = IIF(IsNull(rsTmp("�������")), 0, rsTmp("�������"))
        cmb����ȷ��.ListIndex = IIF(IsNull(rsTmp("����ȷ��")), 0, rsTmp("����ȷ��"))
        
        Call SetComboByText(cmb��������, IIF(IsNull(rsTmp("��������")), "", rsTmp("��������")), True)
        '�õ�������
        strSQL = "select ����,nvl(����,1) ����,nvl(����,'') ���� From �շ���Ŀ���� where ���� in (1,9) And �շ�ϸĿID=[1] order by ����"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mstrID))
                
        Dim blnYes As Boolean
        Do Until rsTmp.EOF
            If rsTmp("����") = txtEdit(Text����).Text Then
                If rsTmp!���� = 1 Then
                    txtEdit(text����).Text = IIF(IsNull(rsTmp("����")), "", rsTmp("����"))
                Else
                    txtEdit(text���).Text = IIF(IsNull(rsTmp("����")), "", rsTmp("����"))
                End If
            Else
                blnYes = False
                For i = 1 To mshAlias.Rows - 1
                    If mshAlias.TextMatrix(i, 0) = rsTmp!���� Then
                        If rsTmp!���� = 1 Or rsTmp!���� = 2 Then
                            mshAlias.TextMatrix(i, rsTmp!����) = rsTmp!����
                        End If
                        blnYes = True
                    End If
                Next
                If blnYes = False Then
                    If Not (mshAlias.Rows = 2 And mshAlias.TextMatrix(1, 0) = "") Then
                        mshAlias.Rows = mshAlias.Rows + 1
                    End If
                    mshAlias.TextMatrix(mshAlias.Rows - 1, 0) = rsTmp!����
                    If rsTmp!���� = 1 Or rsTmp!���� = 2 Then
                        mshAlias.TextMatrix(mshAlias.Rows - 1, rsTmp!����) = rsTmp!����
                    End If
                End If
            End If
            rsTmp.MoveNext
        Loop
        '�õ���ǰ�շѼ�Ŀ����
        
'        strSQL = "select a.ID " & _
'            " from �շѼ�Ŀ A  Where decode(a.��ֹ����,to_date('3000-01-01','YYYY-MM-DD'),null,a.��ֹ����) is null And a.�շ�ϸĿID = [1] "
'        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mstrID))
'
'        msh��Ŀ.Rows = IIF(rsTmp.RecordCount = 0, 2, rsTmp.RecordCount + 1)
        '�޸�
    ElseIf Trim(mstrID) = "" Then    '����
        If gstrҽ�۽ӿڱ�� <> "" And gbln����ҽ���շ���Ŀ = True Then
            txtEdit(text����޼�).Enabled = False
            txtEdit(text����޼�).Enabled = False
        Else
            txtEdit(text����޼�).Visible = False
            txtEdit(text����޼�).Visible = False
            lblEdit(20).Visible = False
            lblEdit(21).Visible = False
        End If
        
        strSQL = "select ������Ŀ from �շ���ĿĿ¼ a,(select max(����ʱ��) as ����ʱ�� from �շ���ĿĿ¼ where ����id=[1]) b where a.����ʱ��=b.����ʱ�� and a.����id=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "������Ŀ", mstr����ID)
        If rsTmp.RecordCount > 0 Then
            txtEdit(15).Text = IIF(IsNull(rsTmp!������Ŀ), "", rsTmp!������Ŀ)
        End If
        
        strSQL = "select ����,���� from �շ���Ŀ��� where Upper(����)=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UCase(Trim(mstr���)))
        
        mblnEditCancel = True
        If rsTmp.RecordCount > 0 Then
            If Me.cmbClass.ListCount > 0 Then
                For i = 0 To Me.cmbClass.ListCount
                    If Me.cmbClass.List(i) = Nvl(rsTmp!����) & "-" & Nvl(rsTmp!����) Then
                        Me.cmbClass.ListIndex = i
                        Exit For
                    End If
                Next
            End If
        End If
        mblnEditCancel = False
        
        '���ñ���
        Call SetCodeNO
        mstr���� = txtEdit(text����).Text
        txtEdit(Text����ʱ��).Text = Format(sys.Currentdate, "yyyy-MM-dd")
        txtEdit(text����).Text = ""
        txtEdit(text���).Text = ""
'        dtpBegin.value = sys.Currentdate
        Select Case mstr���
        Case "1"    '�Һ�
            chk����.Visible = True
            cmb����.Visible = False
            
            chk���.Visible = False
            chk���ηѱ�.Visible = chk���.Visible
            chk�Ӱ�Ӽ�.Visible = chk���.Visible
            chkժҪ.Visible = chk���.Visible
            txtEdit(Text���).Enabled = chk���.Visible
            txtEdit(Text���).BackColor = Me.BackColor
            cmb�������.ListIndex = 1
            cmb�������.Enabled = chk���.Visible
        Case "H"    '����
            cmb����.Visible = True
            chk����.Visible = False
            
            chk���.Visible = False
            chk���ηѱ�.Visible = chk���.Visible
            chk�Ӱ�Ӽ�.Visible = chk���.Visible
            chkժҪ.Visible = chk���.Visible
            txtEdit(Text���).Enabled = chk���.Enabled
            txtEdit(Text���).BackColor = Me.BackColor
            cmb�������.Enabled = chk���.Visible
        Case "J"    '��λ
            chk���.Visible = False
            chk���ηѱ�.Visible = chk���.Visible
            chk�Ӱ�Ӽ�.Visible = chk���.Visible
            chkժҪ.Visible = chk���.Visible
            txtEdit(Text���).Enabled = True
            cmb�������.Enabled = chk���.Visible
        End Select
    End If
    init���� = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub IniStationNo()
    Dim strSQL As String
    Dim rsRecord As ADODB.Recordset
'    lblStationNo.Visible = False
'    cmbStationNo.Visible = False
'
'    If gstrNodeNo <> "-" Then
        lblStationNo.Visible = True
        cmbStationNo.Visible = True
        
        On Error GoTo ErrHandle
        cmbStationNo.Clear
        cmbStationNo.Tag = ""
        strSQL = "select ���,���� from zlnodelist"
        Set rsRecord = zlDatabase.OpenSQLRecord(strSQL, "վ���ѯ")
        If mblnCanUpdateAll Then
            With cmbStationNo
                .AddItem ""
                Do While Not rsRecord.EOF
                    .AddItem rsRecord!��� & "-" & rsRecord!����
                    rsRecord.MoveNext
                Loop
            End With
        Else
            rsRecord.Filter = "���='" & gstrNodeNo & "'"
            With cmbStationNo
                Do While Not rsRecord.EOF
                    .AddItem rsRecord!��� & "-" & rsRecord!����
                    rsRecord.MoveNext
                Loop
                If .ListCount > 0 And .ListIndex = -1 Then .ListIndex = .NewIndex
            End With
        End If
        
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

Private Function init��Ŀ(Optional ByVal blnStationChanged As Boolean) As Boolean
    On Error GoTo ErrHandle
    '����:��ʼ���շѼ�Ŀ��ʹ�����Ŀ��
    Dim strSQL As String, rsTmp As New ADODB.Recordset
    Dim str���� As String
    Dim lngCol As Long
    Dim lngԭ��ID As Long
    Dim lngRow As Long, i As Integer
    Dim strWhere As String, objTabItem As TabControlItem
    Dim blnFind As Boolean
    
    If blnStationChanged = False Then
        With rsTmp
            strSQL = "select ID,����,���� from ������Ŀ where ĩ��=1 and rownum<2"
            Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
            
            If .RecordCount = 0 Then
                MsgBox "���ڡ�������Ŀ���������������Ŀ����ʹ�ñ����ܡ�", vbExclamation, gstrSysName
                .Close
                Exit Function
            End If
            .Close
        End With
    End If
    
    '��ʼ��ҳǩ
    If mblnCanUpdateAll Then
        '������շ���Ŀ��������վ��ģ���϶�ֻ��Ҫ�������վ��ļ۸�ȼ�����
        If Not (medit��ʽ = EditNew Or medit��ʽ = EditCopy Or medit��ʽ = EditModify) Then
            strWhere = "       And Exists(Select 1 From �շ���ĿĿ¼ Where ID = [1] And (վ�� Is Null Or վ�� = b.վ��))"
        ElseIf cmbStationNo.Text <> "" And blnStationChanged Then
            strWhere = "       And b.վ�� =[3]"
        End If
        strSQL = "Select '00' As ����, 'ȱʡ' As �۸�ȼ� From Dual" & vbNewLine & _
                " Union All" & vbNewLine & _
                " Select Distinct a.����, a.���� As �۸�ȼ�" & vbNewLine & _
                " From �շѼ۸�ȼ� A, �շѼ۸�ȼ�Ӧ�� B" & vbNewLine & _
                " Where a.���� = b.�۸�ȼ� And Nvl(a.�Ƿ�������ͨ��Ŀ, 0) = 1" & vbNewLine & _
                        strWhere & vbNewLine & _
                "       And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd'))" & vbNewLine & _
                " Order By ����"
    Else
        strSQL = "Select '00' As ����, 'ȱʡ' As �۸�ȼ� From Dual" & vbNewLine & _
                " Union All" & vbNewLine & _
                " Select Distinct a.����, a.���� As �۸�ȼ�" & vbNewLine & _
                " From �շѼ۸�ȼ� A, �շѼ۸�ȼ�Ӧ�� B" & vbNewLine & _
                " Where a.���� = b.�۸�ȼ� And b.վ�� = [2] And Nvl(a.�Ƿ�������ͨ��Ŀ, 0) = 1" & vbNewLine & _
                "       And (a.����ʱ�� Is Null Or a.����ʱ�� = To_Date('3000-01-01', 'yyyy-mm-dd'))" & vbNewLine & _
                " Order By ����"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "��ȡ�۸�ȼ�", Val(mstrID), gstrNodeNo, zlStr.NeedCode(cmbStationNo.Text, "-"))
    If rsTmp.RecordCount = 0 Then
        MsgBox "��û�е������շ���Ŀ�۸��Ȩ�ޣ�", vbInformation + vbOKOnly, gstrSysName
        Unload Me: Exit Function
    End If
        
    tbPriceGrade.RemoveAll
    lngRow = 0
    Do While Not rsTmp.EOF
        If lngRow > pic�۸�ȼ�.UBound Then
            Load pic�۸�ȼ�(lngRow): pic�۸�ȼ�(lngRow).Visible = True
            Load msh��Ŀ(lngRow): Set msh��Ŀ(lngRow).Container = pic�۸�ȼ�(lngRow): msh��Ŀ(lngRow).Visible = True
            Load lbl����ִ��ʱ��(lngRow): Set lbl����ִ��ʱ��(lngRow).Container = pic�۸�ȼ�(lngRow): lbl����ִ��ʱ��(lngRow).Visible = True
            Load dtpBegin(lngRow): Set dtpBegin(lngRow).Container = pic�۸�ȼ�(lngRow): dtpBegin(lngRow).Visible = True
            Load chkNow(lngRow): Set chkNow(lngRow).Container = pic�۸�ȼ�(lngRow): chkNow(lngRow).Visible = True
            Load lbl����˵��(lngRow): Set lbl����˵��(lngRow).Container = pic�۸�ȼ�(lngRow): lbl����˵��(lngRow).Visible = True
            Load txt����˵��(lngRow): Set txt����˵��(lngRow).Container = pic�۸�ȼ�(lngRow): txt����˵��(lngRow).Visible = True
        End If
        Set objTabItem = tbPriceGrade.InsertItem(lngRow, Nvl(rsTmp!�۸�ȼ�), pic�۸�ȼ�(lngRow).hwnd, 0)
        lngRow = lngRow + 1
        rsTmp.MoveNext
    Loop
    If blnStationChanged Then Exit Function
    
    With tbPriceGrade.PaintManager
        .Appearance = xtpTabAppearancePropertyPage2003 '��ʾ���
        .BoldSelected = True '��ʾҳ��������Ӵ�
        .ClientFrame = xtpTabFrameSingleLine 'ҳ��߿�
        .Layout = xtpTabLayoutAutoSize
        .Position = xtpTabPositionBottom
        .StaticFrame = True
    End With
    If tbPriceGrade.ItemCount > 0 Then tbPriceGrade(0).Selected = True
    
    ReDim mcol��Ŀ(tbPriceGrade.ItemCount - 1)
    ReDim mblnNew(tbPriceGrade.ItemCount - 1)
    ReDim mblnChanged��Ŀ(tbPriceGrade.ItemCount - 1)
    For i = 0 To tbPriceGrade.ItemCount - 1
        Set mcol��Ŀ(i) = New Collection
        With msh��Ŀ(i)
            .Cols = mcstCols
            .ColWidth(mcstCol�շ���Ŀ) = 1500
            .ColWidth(mcstColԭ��) = 1000
            .ColWidth(mcstCol�ּ�) = 1000
            .ColWidth(mcstColȱʡ�۸�) = IIF(chk���.value = 1, 1000, 0)
            .TextMatrix(0, mcstCol�շ���Ŀ) = "������Ŀ"
            .TextMatrix(0, mcstColԭ��) = "ԭ��"
            .TextMatrix(0, mcstCol�ּ�) = "�ּ�"
            .TextMatrix(0, mcstColȱʡ�۸�) = "ȱʡ�۸�"
            If mstr��� = "F" Then
                .TextMatrix(0, mcstCol���������շ���) = "���������շ���"
                .ColWidth(mcstCol���������շ���) = 1500
            Else
                .ColWidth(mcstCol���������շ���) = 0
            End If
            .TextMatrix(0, mcstCol�Ӱ�Ӽ���) = "�Ӱ�Ӽ���"
            .ColWidth(mcstCol�Ӱ�Ӽ���) = 0
            '���뷽ʽ
            .ColAlignment(mcstCol�շ���Ŀ) = 1
            .ColAlignment(mcstColԭ��) = 7
            .ColAlignment(mcstCol�ּ�) = 7
            .ColAlignment(mcstColȱʡ�۸�) = 7
            .ColAlignment(mcstCol���������շ���) = 7
            .ColAlignment(mcstCol�Ӱ�Ӽ���) = 7
            'ʵ�ַ�ʽ
            .ColData(mcstCol�շ���Ŀ) = 1 '�������룬����һ����ť
            .ColData(mcstColԭ��) = 5 '������ѡ��
            .ColData(mcstCol�ּ�) = 4 'ֱ������
            .ColData(mcstColȱʡ�۸�) = 4 'ֱ������
            .ColData(mcstCol���������շ���) = 4
            .ColData(mcstCol�Ӱ�Ӽ���) = 4
            
            .PrimaryCol = 0
            .Active = True
        End With
        Me.dtpBegin(i).value = DateAdd("d", 1, Now)
    Next
    If mstrID = "" Then
        init��Ŀ = True
        Exit Function
    End If
        
    'װ������
    strSQL = "select ����,�Ƿ���,�Ӱ�Ӽ�,����޼�,����޼�  from �շ���ĿĿ¼ where ID=[1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mstrID))
        
    If Not mblnShow�շѼ�Ŀ Then
        chk���.Tag = IIF(rsTmp("�Ƿ���") = 1, 1, 0)
        chk���.value = IIF(rsTmp("�Ƿ���") = 1, 1, 0)
        For i = 0 To tbPriceGrade.ItemCount - 1
            msh��Ŀ(i).ColWidth(mcstColȱʡ�۸�) = IIF(rsTmp("�Ƿ���") = 1, 1000, 0)
        Next
    End If
    chk�Ӱ�Ӽ�.value = IIF(rsTmp("�Ӱ�Ӽ�") = 1, 1, 0)
    If Not mblnShow�շѼ�Ŀ Then
        mdbl����޼� = Nvl(rsTmp!����޼�, 0)
        mdbl����޼� = Nvl(rsTmp!����޼�, 0)
    Else
        mdbl����޼� = Val(txtEdit(text����޼�).Text)
        mdbl����޼� = Val(txtEdit(text����޼�).Text)
    End If
    '���ݾ������ݸı���ͷ
    Call chk���_Click
    Call chk�Ӱ�Ӽ�_Click
    
    For i = 0 To tbPriceGrade.ItemCount - 1
        '��ʾ�շѼ�Ŀ
        strSQL = "Select a.ID,a.ԭ��ID,a.�շ�ϸĿID,Nvl(a.ԭ��,0) As ԭ��,Nvl(a.�ּ�,0) As �ּ�," & vbNewLine & _
                "        Nvl(a.ȱʡ�۸�,0) As ȱʡ�۸�,a.������ĿID,b.����,a.�Ӱ�Ӽ���,a.�����շ���," & vbNewLine & _
                "        a.�䶯ԭ��,a.����˵��,a.ִ������,a.��ֹ���� " & vbNewLine & _
                " From �շѼ�Ŀ A,������Ŀ B" & vbNewLine & _
                " Where a.������ĿID=b.ID And (a.��ֹ���� Is Null Or a.��ֹ���� = To_Date('3000-01-01','YYYY-MM-DD'))" & vbNewLine & _
                "       And a.�շ�ϸĿID = [1] " & IIF(tbPriceGrade(i).Caption = "ȱʡ", " And a.�۸�ȼ� Is Null", " And a.�۸�ȼ�=[2]")
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Val(mstrID), tbPriceGrade(i).Caption)
            
        msh��Ŀ(i).Rows = IIF(rsTmp.RecordCount = 0, 2, rsTmp.RecordCount + 1)
        msh��Ŀ(i).Tag = msh��Ŀ(i).Rows
        
        mblnNew(i) = rsTmp.RecordCount = 0 '�¼۸�
        If rsTmp.RecordCount = 0 Then
            For lngCol = 0 To mcstCols - 1
                msh��Ŀ(i).TextMatrix(1, lngCol) = ""
            Next
            dtpBegin(i).value = sys.Currentdate
            If medit��ʽ = EditCopy Or medit��ʽ = EditNew Then
                txt����˵��(i).Text = "��ʼ�۸�"
            Else
                txt����˵��(i).Text = ""
            End If
        Else
            lngCol = 1
            If medit��ʽ = EditCopy Then
                dtpBegin(i).value = Format(sys.Currentdate, "yyyy-mm-dd h:m:s")
            Else
                If mblnIsSpecialItem Then
                    dtpBegin(i).CustomFormat = "yyyy��MM��dd��"
                    dtpBegin(i).Width = 1600
                    mstrCurrentDateFormat = "yyyy-mm-dd"
                Else
                    dtpBegin(i).CustomFormat = "yyyy��MM��dd�� HH:mm:ss"
                    dtpBegin(i).Width = 2535
                    mstrCurrentDateFormat = "yyyy-mm-dd hh:mm:ss"
                End If
                
                If chk���.value = 0 Then '1 ���Ǳ����Ŀ
                    If mblnIsSpecialItem Then       '1.1 ��������Ŀ
                        If DateDiff("s", rsTmp("ִ������"), sys.Currentdate) > 0 Then        '1.1.1 �ϴο�ʼʱ��С�ڵ�ǰʱ��
                            If DateDiff("s", rsTmp("ִ������"), Format(sys.Currentdate, "yyyy-mm-dd 00:00:00")) > 0 Then     '1.1.1.1 �ϴο�ʼʱ����ڵ������ʱ��
                                chkNow(i).Visible = True
                                dtpBegin(i).MinDate = sys.Currentdate
                            Else        '1.1.1.2 �ϴο�ʼʱ��С�ڵ������ʱ��
                                chkNow(i).Visible = False
                                dtpBegin(i).MinDate = DateAdd("d", 1, Format(sys.Currentdate, "yyyy-mm-dd h:m:s"))
                            End If
                            dtpBegin(i).value = DateAdd("d", 1, Format(sys.Currentdate, "yyyy-mm-dd h:m:s"))
                        Else        '1.1.2 �ϴο�ʼʱ����ڵ�ǰʱ��
                            dtpBegin(i).value = DateAdd("d", 1, Format(rsTmp("ִ������"), "yyyy-mm-dd h:m:s"))
                            dtpBegin(i).MinDate = DateAdd("d", 1, Format(rsTmp("ִ������"), "yyyy-mm-dd h:m:s"))
                            chkNow(i).Visible = False
                        End If
                    Else        '1.2 ����������Ŀ
                        If DateDiff("s", rsTmp("ִ������"), sys.Currentdate) > 0 Then        '1.2.1 �ϴο�ʼʱ��С�ڵ�ǰʱ��
                            dtpBegin(i).value = Format(DateAdd("d", 1, Format(sys.Currentdate, "yyyy-mm-dd h:m:s")), "yyyy-mm-dd 00:00:00")
                            dtpBegin(i).MinDate = DateAdd("s", 1, Format(sys.Currentdate, "yyyy-mm-dd h:m:s"))
                        Else    '1.2.2 �ϴο�ʼʱ����ڵ�ǰʱ��
                            dtpBegin(i).value = Format(DateAdd("d", 1, Format(rsTmp("ִ������"), "yyyy-mm-dd h:m:s")), "yyyy-mm-dd 00:00:00")
                            dtpBegin(i).MinDate = DateAdd("s", 1, Format(rsTmp("ִ������"), "yyyy-mm-dd h:m:s"))
                        End If
                        chkNow(i).Visible = True
                    End If
                    txt����˵��(i).Text = ""
                Else    '2 �Ǳ����Ŀ
                    dtpBegin(i).value = Format(rsTmp("ִ������"), "yyyy-mm-dd h:m:s")
                    dtpBegin(i).Enabled = False
                End If
            End If
            Do Until rsTmp.EOF
                msh��Ŀ(i).TextMatrix(lngCol, mcstCol�շ���Ŀ) = rsTmp("����")
                If chk���.value = 1 Then '���
                    msh��Ŀ(i).TextMatrix(lngCol, mcstColԭ��) = Format(rsTmp("ԭ��"), "###########0.000;-##########0.000;0.000;0.000")
                    msh��Ŀ(i).TextMatrix(lngCol, mcstCol�ּ�) = Format(rsTmp("�ּ�"), "###########0.000;-##########0.000;0.000;0.000")
                    msh��Ŀ(i).TextMatrix(lngCol, mcstColȱʡ�۸�) = Format(rsTmp("ȱʡ�۸�"), "###########0.000;-##########0.000;0.000;0.000")
                Else
                    msh��Ŀ(i).TextMatrix(lngCol, mcstColԭ��) = Format(rsTmp("�ּ�"), "###########0.000;-##########0.000;0.000;0.000")
                    If medit��ʽ = EditCopy Then msh��Ŀ(i).TextMatrix(lngCol, mcstCol�ּ�) = Format(rsTmp("�ּ�"), "###########0.000;-##########0.000;0.000;0.000")
                End If
                msh��Ŀ(i).TextMatrix(lngCol, mcstCol���������շ���) = IIF(IsNull(rsTmp("�����շ���")), 0, rsTmp("�����շ���"))
                msh��Ŀ(i).TextMatrix(lngCol, mcstCol�Ӱ�Ӽ���) = IIF(IsNull(rsTmp("�Ӱ�Ӽ���")), 0, rsTmp("�Ӱ�Ӽ���"))
                msh��Ŀ(i).RowData(lngCol) = rsTmp("������ĿID")
                lngԭ��ID = rsTmp("ID")
                mcol��Ŀ(i).Add lngԭ��ID, "C" & rsTmp("������ĿID")
                lngCol = lngCol + 1
                rsTmp.MoveNext
            Loop
        End If
        If medit��ʽ = EditRaise Then
            msh��Ŀ(i).Col = 2
        End If
    Next
    
    If mblnCanUpdateAll = False Then
        strSQL = "Select 1 From �շ���ĿĿ¼ Where ID = [1] And վ�� = [2] And Rownum < 2"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, "�жϸ���Ŀ�Ƿ����ڵ�ǰվ��", Val(mstrID), gstrNodeNo)
        If rsTmp.EOF Then
            '����ȫԺ��Ŀ�����ܵ���ȱʡ�۸�
            msh��Ŀ(0).MsfObj.Enabled = False
            msh��Ŀ(0).BackColor = vbButtonFace
            msh��Ŀ(0).BackColorBkg = vbButtonFace
            dtpBegin(0).Enabled = False
            chkNow(0).Enabled = False
            txt����˵��(0).Enabled = False
            txt����˵��(0).BackColor = vbButtonFace
            If tbPriceGrade.ItemCount > 1 Then tbPriceGrade.Item(1).Selected = True
        Else
            msh��Ŀ(0).MsfObj.Enabled = True
            msh��Ŀ(0).BackColor = vbWindowBackground
            msh��Ŀ(0).BackColorBkg = vbWindowBackground
            'dtpBegin(0).Enabled = True
            chkNow(0).Enabled = True
            txt����˵��(0).Enabled = True
            txt����˵��(0).BackColor = vbWindowBackground
        End If
    End If
    
    init��Ŀ = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub init����()
    On Error GoTo ErrHandle
    Dim rsTmp As New ADODB.Recordset
    Dim i As Integer
    Dim strWherePriceGrade As String
    
    With msh����
        .Cols = 4
        .ColWidth(0) = 2000
        .ColWidth(1) = 800
        .ColWidth(2) = 1400
        .ColWidth(3) = 1400
        .ColAlignment(0) = 1
        .TextMatrix(0, 0) = "�շ���Ŀ"
        .TextMatrix(0, 1) = "����"
        .TextMatrix(0, 2) = "�̶���ϵ"
        .TextMatrix(0, 3) = "����"
        .ColAlignment(2) = 1
        'ʵ�ַ�ʽ
        .ColData(0) = 1 '��ʾ���п������룬�ⲿ��ʾΪ��ťѡ��
        .ColData(1) = 4 'ֱ������
        .ColData(2) = 3
        
        .PrimaryCol = 1
        .Active = True
    End With
    Me.lbl�����ϼ�.Caption = ""
    Me.lbl�����ϼ�.Tag = 0
    
    mstr�б�(3) = "0-���̶�;1-�̶�;2-����������"
    '��ʾ������Ŀ
    If mstrID = "" Then Exit Sub
    
    If gstr��ͨ�۸�ȼ� = "" And gstrҩƷ�۸�ȼ� = "" And gstr���ļ۸�ȼ� = "" Then
       strWherePriceGrade = " And d.�۸�ȼ� Is Null"
    Else
       strWherePriceGrade = "" & _
           " And ((Instr(';5;6;7;', ';' || b.��� || ';') > 0 And d.�۸�ȼ� = [2])" & vbNewLine & _
           "      Or (Instr(';4;', ';' || b.��� || ';') > 0 And d.�۸�ȼ� = [3])" & vbNewLine & _
           "      Or (Instr(';4;5;6;7;', ';' || b.��� || ';') = 0 And d.�۸�ȼ� = [4])" & vbNewLine & _
           "      Or (d.�۸�ȼ� Is Null" & vbNewLine & _
           "          And Not Exists (Select 1" & vbNewLine & _
           "                          From �շѼ�Ŀ" & vbNewLine & _
           "                          Where d.�շ�ϸĿid = �շ�ϸĿid And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
           "                                And ((Instr(';5;6;7;', ';' || b.��� || ';') > 0 And �۸�ȼ� = [2])" & vbNewLine & _
           "                                      Or (Instr(';4;', ';' || b.��� || ';') > 0 And �۸�ȼ� = [3])" & vbNewLine & _
           "                                      Or (Instr(';4;5;6;7;', ';' || b.��� || ';') = 0 And �۸�ȼ� = [4])))))"
    End If
    gstrSQL = "Select a.����ID,a.����ID,a.���д���,a.��������,b.����,b.���� ��Ŀ����,c.���� ,c.���� ���, " & vbCrLf & _
            "           decode(nvl(b.�Ƿ���,0),1,ltrim(rtrim(to_char(sum(d.ԭ��),'9999999990.00')))||'��'||ltrim(rtrim(to_char(sum(d.�ּ�),'9999999990.00'))),ltrim(rtrim(to_char(sum(d.�ּ�),'9999999990.00'))))  AS  �۸� " & vbCrLf & _
            " From �շѴ�����Ŀ a,�շ���ĿĿ¼ b,�շ���Ŀ��� c ,�շѼ�Ŀ d " & vbCrLf & _
            " Where c.����=b.��� and  a.����ID=b.id  and b.id=d.�շ�ϸĿid  and ����ID=[1] " & vbCrLf & _
            "       AND NVL (D.��ֹ����, TO_DATE ('3000-01-01', 'YYYY-MM-DD')) = TO_DATE ('3000-01-01', 'YYYY-MM-DD') " & _
                    strWherePriceGrade & vbNewLine & _
            " GROUP BY a.ROWID,a.����ID,b.�Ƿ���,a.����ID,a.���д���,a.��������,b.����,b.���� ,c.���� ,c.���� " & _
            " ORDER BY a.ROWID "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstrID), gstrҩƷ�۸�ȼ�, gstr���ļ۸�ȼ�, gstr��ͨ�۸�ȼ�)
    
    msh����.Rows = IIF(rsTmp.RecordCount = 0, 2, rsTmp.RecordCount + 1)
    If rsTmp.RecordCount = 0 Then
        For i = 0 To 3
            msh����.TextMatrix(1, i) = ""
        Next
    Else
        i = 1
        Do Until rsTmp.EOF
            msh����.TextMatrix(i, 0) = "[" & rsTmp("��Ŀ����") & "]" & rsTmp("����")
            msh����.TextMatrix(i, 1) = rsTmp("��������")
            
            If rsTmp("���д���") = 0 Then
                msh����.TextMatrix(i, 2) = "0-���̶�"
            ElseIf rsTmp("���д���") = 2 Then
                msh����.TextMatrix(i, 2) = "2-����������"
            Else
                msh����.TextMatrix(i, 2) = "1-�̶�"
            End If
            msh����.TextMatrix(i, 3) = rsTmp("�۸�")
            msh����.RowData(i) = rsTmp("����ID")
            i = i + 1
            rsTmp.MoveNext
        Loop
        msh����_EnterCell 1, 0
    End If
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub initִ��()
    On Error GoTo ErrHandle
    Dim rsTmp As New ADODB.Recordset
    Dim i As Long
    Dim lngSel As Long
    Dim strTmp As String
    
    opt����_Click 0
    If mstrID = "" Then Exit Sub
    '��ʾ����
    gstrSQL = "select A.���,A.ID,A.����ID,A.ִ�п���,B.����,C.����  ��� from �շ���ĿĿ¼ A,�շѷ���Ŀ¼ B,�շ���Ŀ��� C where A.����ID=B.ID(+) and A.���=C.���� and A.ID=[1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstrID))
        
    '��ʾ��ǰӦ���ĸ����
    Me.optApply(1).Caption = "Ӧ����" & IIF(mblnCanUpdateAll, "", "��Ժ��") & " " & Nvl(rsTmp("����")) & " ����ͬ����������Ŀ(&G)"
    Me.optApply(2).Caption = "Ӧ����" & IIF(mblnCanUpdateAll, "", "��Ժ��") & " " & Nvl(rsTmp("����")) & " ������������Ŀ(&L)"
    Me.optApply(3).Caption = "Ӧ����" & IIF(mblnCanUpdateAll, "", "��Ժ��") & " " & Nvl(rsTmp("���")) & " �����������Ŀ(&U)"
    lngSel = IIF(rsTmp("ִ�п���") < 7, rsTmp("ִ�п���"), 0)
    opt����(lngSel).value = True
    opt����_Click IIF(rsTmp("ִ�п���") < 7, rsTmp("ִ�п���"), 0)
    
    If opt����(4).value = True Or opt����(0).value = True Then
        '����סԺִ�п���
        gstrSQL = "select R.������Դ,E.ID,E.����" & _
                " from �շ�ִ�п��� R,���ű� E" & _
                " where R.ִ�п���ID=E.ID and R.������Դ in (1,2) and R.��������id is null and R.�շ�ϸĿID=[1] "
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstrID))
                
        With rsTmp
            Do While Not .EOF
                If !������Դ = 1 Then Me.txt����ִ��.Text = !����: Me.txt����ִ��.Tag = !ID
                If !������Դ = 2 Then Me.txtסԺִ��.Text = !����: Me.txtסԺִ��.Tag = !ID
                .MoveNext
            Loop
        End With
        
        If opt����(4).value = True Then
            gstrSQL = _
            "select a.�շ�ϸĿid �շ�ϸĿId,a.������Դ," & vbCrLf & _
                "       b.id ����id,b.���� ��������,b.���� ��������," & vbCrLf & _
                "       c.id ִ��id,c.���� ִ�б���,c.���� ִ������  " & vbCrLf & _
                "  from �շ�ִ�п��� a,���ű� b,���ű� c" & vbCrLf & _
                " where a.ִ�п���ID=c.id(+) And a.��������ID=b.id(+)  and a.������Դ is null and a.�շ�ϸĿID=[1] and " & Where����ʱ��("B") & vbCrLf & _
                " Order By c.����"
            Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstrID))
            
            Me.msf����ִ��.ClearBill
            
            With rsTmp
                Do While Not .EOF
                    If strTmp <> !ִ������ Then
                        i = i + 1
                        Me.msf����ִ��.Rows = i + 1
                        Me.msf����ִ��.TextMatrix(i, 2) = IIF(IsNull(!����ID), "�����в��ţ�", !����ID)
                        Me.msf����ִ��.TextMatrix(i, 3) = IIF(IsNull(!����ID), "�����в��ţ�", "[" & !�������� & "]" & !��������)
                        Me.msf����ִ��.TextMatrix(i, 0) = !ִ��ID
                        Me.msf����ִ��.TextMatrix(i, 1) = "[" & !ִ�б��� & "]" & !ִ������
                    Else
                        Me.msf����ִ��.TextMatrix(i, 2) = Me.msf����ִ��.TextMatrix(i, 2) & "," & !����ID
                        Me.msf����ִ��.TextMatrix(i, 3) = Me.msf����ִ��.TextMatrix(i, 3) & ",[" & !�������� & "]" & !��������
                    End If
                    strTmp = !ִ������
                    .MoveNext
                Loop
            End With
        End If
    End If
    
    '�����޸�
    If medit��ʽ = 3 Then
        fra����.Visible = True
    End If
    
    'ȡ�������ʷ���
    Ini���ʷ���
    
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub ShowTab(ByVal strTab As String)
    '����:��ʾָ��ҳ
    '����:strTab ҳ��
    On Error Resume Next
    TabMain.Tabs("_" & strTab).Selected = True
    tabMain_Click
End Sub

Private Sub ShowItem(lst As ListItem)
    On Error GoTo ErrHandle
    '������ʾĳһ��,����ˢ��
    Dim rsTmp As New ADODB.Recordset
    Dim lngCol  As Long
    Dim varValue As Variant
    
    rsTmp.CursorLocation = adUseClient
    gstrSQL = "Select A.ID,A.���,A.����,A.����,A.���,A.���㵥λ,A.��������," & vbCrLf & _
        " decode(A.�������,1,'����',2,'סԺ',3,'������סԺ','��') as �������,decode(A.����ժҪ,1,'��','') as ����ժҪ," & vbCrLf & _
        " decode(A.���,'1',decode(A.��Ŀ����,1,'����',''),'H',decode(A.��Ŀ����,1,'����ȼ�',2,'��������', '')) ��Ŀ����," & vbCrLf & _
        " A.˵��,decode(A.���ηѱ�,1,'��','') as ���ηѱ�,decode(A.�Ƿ���,1,'��','') as �Ƿ���,decode(A.�Ӱ�Ӽ�,1,'��','') as �Ӱ�Ӽ�,A.ִ�п���," & vbCrLf & _
        " to_char(A.����ʱ��,'yyyy-mm-dd') as ����ʱ��,to_char(A.����ʱ��,'yyyy-mm-dd') as ����ʱ��," & vbCrLf & _
        " '" & txtEdit(Text����).Text & "' as �������� From �շ���ĿĿ¼ A " & vbCrLf & _
        " Where A.ID=[1] "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(mstrID))
        
    '����ListView�����������ݿ�ȡ��
    lst.Text = rsTmp("����")
    For lngCol = 2 To frmChargeManage.lvwMain_S.ColumnHeaders.Count
        varValue = rsTmp(frmChargeManage.lvwMain_S.ColumnHeaders(lngCol).Text).value
        lst.SubItems(lngCol - 1) = IIF(IsNull(varValue), "", varValue)
        lst.Tag = rsTmp!ID
    Next
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Function IsValid����() As Boolean
    '����:����������Ϣҳ�����������Ƿ���Ч
    '����:intTab ҳ��
    '����ֵ:��Ч����True,����ΪFalse
    On Error GoTo ErrHandle
    Dim i As Integer
    Dim strTemp As String
    Dim strSQL As String
    Dim rsTmp As New ADODB.Recordset
    Dim j As Long
    Dim str���� As String
    
    IsValid���� = False
    For i = 0 To 5
        strTemp = Trim(txtEdit(i).Text)
        If i <> Text��� And i <> text���� Then
        If zlCommFun.StrIsValid(txtEdit(i).Text, txtEdit(i).MaxLength) = False Then
            ShowTab "������Ϣ"
            If txtEdit(i).Enabled And txtEdit(i).Visible Then
                txtEdit(i).SetFocus
                zlControl.TxtSelAll txtEdit(i)
            End If
            Exit Function
        End If
        End If
    Next
    '�����
    If InStr(cmbClass.Text, "-") > 0 Then
        strTemp = Left(cmbClass.Text, 1)
        strSQL = "select ���� from �շ���Ŀ��� where ����<>'4' And ����<>'5' and ����<>'6' and ����<>'7' and upper(����) =[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, UCase(Trim(strTemp)))
        
        If rsTmp.RecordCount < 1 Then
            ShowTab "������Ϣ"
            MsgBox "����������ȷ�����������룡", vbExclamation, gstrSysName
            cmbClass.SetFocus
            Exit Function
        Else
            mstr��� = Nvl(rsTmp!����)
        End If
    Else
        ShowTab "������Ϣ"
        If Trim(cmbClass.Text) = "" Then
            MsgBox "�����Ϊ�գ����������룡", vbExclamation, gstrSysName
        Else
            MsgBox "�����ȷ�����������룡", vbExclamation, gstrSysName
        End If
        If cmbClass.Visible And cmbClass.Enabled Then
            cmbClass.SetFocus
        End If
        Exit Function
    End If
    '������
    If Trim(txtEdit(Text����).Text) = "��" Or Trim(txtEdit(Text����).Text) = "" Then
        txtEdit(Text����).Text = "��"
        mstr����ID = "0"
        MsgBox "���಻��Ϊ�գ����������룡", vbExclamation, gstrSysName
        If txtEdit(Text����).Visible And txtEdit(Text����).Enabled Then
            txtEdit(Text����).SetFocus
        End If
        Exit Function
    Else
        strSQL = "Select 1 From �շѷ���Ŀ¼ Where ID " & IIF(Trim(mstr����ID) = "" Or Trim(mstr����ID) = "0", " is null ", " = [2]") & " And ����=[1]"
        Set rsTmp = zlDatabase.OpenSQLRecord(strSQL, Me.Caption, Trim(txtEdit(Text����).Text), Val(mstr����ID))

        If rsTmp.RecordCount < 1 Then
            ShowTab "������Ϣ"
            MsgBox "��������������������룡", vbExclamation, gstrSysName
            If txtEdit(Text����).Visible And txtEdit(Text����).Enabled Then
                txtEdit(Text����).SetFocus
            End If
            Exit Function
        End If
    End If
    
    txtEdit(text����).Text = Trim(txtEdit(text����).Text)
    '���㵥λ
    If zlCommFun.StrIsValid(cmb���㵥λ.Text, mlng��λ����, , "���㵥λ") = False Then
        ShowTab "������Ϣ"
        If cmb���㵥λ.Enabled And cmb���㵥λ.Visible Then
            cmb���㵥λ.SetFocus
        End If
        Exit Function
    End If
    
    If txtTemp.MaxLength = 0 Then
        If Len(txtEdit(text����).Text) = 0 Then
            ShowTab "������Ϣ"
            txtEdit(text����).SetFocus
            MsgBox "���벻��Ϊ�ա�", vbExclamation, gstrSysName
            Exit Function
        End If
    Else
        If Len(txtEdit(text����).Text) < txtEdit(text����).MaxLength Then
            ShowTab "������Ϣ"
            txtEdit(text����).SetFocus
            MsgBox "����ĳ��Ȳ�����", vbExclamation, gstrSysName
            Exit Function
        End If
    End If
    If medit��ʽ = EditCopy Or medit��ʽ = EditNew Or medit��ʽ = EditModify Then
        gstrSQL = "select ���,����,���� from �շ���ĿĿ¼ where ����=[1] " & IIF(medit��ʽ = EditCopy Or medit��ʽ = EditNew, "", " And ID <> [2] ")
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Trim(txtTemp.Text) & txtEdit(text����).Text, Val(mstrID))
        
        If rsTmp.RecordCount > 0 Then
            rsTmp.MoveFirst
            strTemp = ""
            For j = 0 To rsTmp.RecordCount - 1
                strTemp = strTemp & "   [" & rsTmp!��� & rsTmp!���� & "]" & rsTmp!���� & IIF(j = rsTmp.RecordCount - 1, "", vbCrLf)
                rsTmp.MoveNext
            Next
            ShowTab "������Ϣ"
            txtEdit(text����).SetFocus
            MsgBox "������������Ŀ�����ظ��� " & vbCrLf & strTemp & vbCrLf & " �����������������룡", vbInformation, gstrSysName
            Exit Function
        End If
    End If
    '����ʶ����Ϊ��д
    txtEdit(text��ʶ����).Text = UCase(txtEdit(text��ʶ����).Text)
    txtEdit(Text��ʶ����).Text = UCase(txtEdit(Text��ʶ����).Text)
'    txtEdit(Text��ѡ��).Text = UCase(txtEdit(Text��ѡ��).Text)
    If Len(Trim(txtEdit(text��ʶ����).Text)) < 1 And (gstrҽ�۽ӿڱ�� <> "" And gbln����ҽ���շ���Ŀ) Then
        MsgBox "����ʶ���롱������Ϊ�գ�", vbExclamation, gstrSysName
        ShowTab "������Ϣ"
        If txtEdit(text��ʶ����).Enabled And txtEdit(text��ʶ����).Visible Then
            txtEdit(text��ʶ����).SetFocus
        End If
        Exit Function
    End If
    If Len(Trim(txtEdit(Text��ѡ��).Text)) > 0 Then
        For i = 1 To Len(Trim(txtEdit(Text��ѡ��).Text))
            If InStr("0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ", Mid(Trim(txtEdit(Text��ѡ��).Text), i, 1)) < 1 Then
                MsgBox "��ѡ�����������ĸ��������ɡ�", vbExclamation, gstrSysName
                ShowTab "������Ϣ"
                If txtEdit(Text��ѡ��).Enabled And txtEdit(Text��ѡ��).Visible Then
                    txtEdit(Text��ѡ��).SetFocus
                End If
                Exit Function
            End If
        Next
    End If
    If Len(Trim(txtEdit(Text����).Text)) = 0 Then
        ShowTab "������Ϣ"
        txtEdit(Text����).SetFocus
        MsgBox "���Ʋ���Ϊ�ա�", vbExclamation, gstrSysName
        txtEdit(Text����).Text = ""
        Exit Function
    End If
    If Len(Trim(txtEdit(text����޼�).Text)) = 0 Then txtEdit(text����޼�).Text = 0
    If Len(Trim(txtEdit(text����޼�).Text)) = 0 Then txtEdit(text����޼�).Text = 0
    For i = 1 To mshAlias.Rows - 1
        If Trim(mshAlias.TextMatrix(i, 0)) = Trim(txtEdit(Text����).Text) Then
            ShowTab "������Ϣ"
            mshAlias.Row = i
            mshAlias.Col = 0
            If mshAlias.Active And mshAlias.Visible Then
                mshAlias.SetFocus
            End If
            MsgBox "������������ͬ�ˡ�", vbExclamation, gstrSysName
            Exit Function
        End If
        For j = 1 To mshAlias.Rows - 1
            If Trim(mshAlias.TextMatrix(i, 0)) = Trim(mshAlias.TextMatrix(j, 0)) And i <> j Then
                ShowTab "������Ϣ"
                mshAlias.Row = j
                mshAlias.Col = 0
                If mshAlias.Active And mshAlias.Visible Then
                    mshAlias.SetFocus
                End If
                MsgBox "�����ظ���", vbExclamation, gstrSysName
                Exit Function
            End If
        Next
    Next
    
    '�������ַ�������
    If Trim(txtEdit(text����).Text) = "" Then
        txtEdit(text����).Text = zlStr.GetCodeByORCL(txtEdit(Text����).Text, False, mlng���볤��)
    End If

    If Trim(txtEdit(text���).Text) = "" Then
        txtEdit(text���).Text = zlStr.GetCodeByORCL(txtEdit(Text����).Text, True, mlng���볤��)
    End If
    
    With mshAlias
        If Trim(txtEdit(text����).Text) <> "" Then
            str���� = "1''" & txtEdit(Text����).Text & "''1''" & txtEdit(text����).Text & "''"
        End If
        If Trim(txtEdit(text���).Text) <> "" Then
            str���� = str���� & "1''" & txtEdit(Text����).Text & "''2''" & txtEdit(text���).Text & "''"
        End If
        For i = 1 To .Rows - 1
            If Trim(.TextMatrix(i, 0)) <> "" Then
                If Trim(.TextMatrix(i, 1)) <> "" Then
                    str���� = str���� & "9''" & Trim(.TextMatrix(i, 0)) & "''1''" & Trim(.TextMatrix(i, 1)) & "''"
                End If
                If Trim(.TextMatrix(i, 2)) <> "" Then
                    str���� = str���� & "9''" & Trim(.TextMatrix(i, 0)) & "''2''" & Trim(.TextMatrix(i, 2)) & "''"
                End If
            End If
        Next
    End With
    If LenB(str����) > 4000 Then
        ShowTab "������Ϣ"
        If mshAlias.Active And mshAlias.Visible Then
            mshAlias.SetFocus
        End If
        MsgBox "�����ַ���̫��������ٱ����������߱������ȡ�", vbExclamation, gstrSysName
        Exit Function
    End If
    
    IsValid���� = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function IsValid��Ŀ() As Boolean
    '����:�����շѼ�Ŀҳ�����������Ƿ���Ч
    '����:intTab ҳ��
    '����ֵ:��Ч����True,����ΪFalse
    On Error GoTo ErrHandle
    Dim i As Integer
    Dim j As Integer, k As Integer
    Dim dbl�ϼƼ۸� As Double
    Dim blnNothing As Boolean
    Dim blnHaveData As Boolean
    
    IsValid��Ŀ = False
    For k = 0 To tbPriceGrade.ItemCount - 1
        If tbPriceGrade(k).Visible _
            And Not (mblnChanged��Ŀ(k) = False And (medit��ʽ = EditModify Or medit��ʽ = EditRaise)) Then   '���صĺ�δ�ı�Ĳ����
            blnNothing = False
            With msh��Ŀ(k)
                If Trim(.TextMatrix(1, mcstCol�շ���Ŀ)) = "" Then
                    If tbPriceGrade(k).Caption = "ȱʡ" Or tbPriceGrade.ItemCount = 1 Then
                        ShowTab "�շѼ�Ŀ"
                        tbPriceGrade.Item(k).Selected = True
                        If .Active And .Visible Then
                            .SetFocus
                        End If
                        .Row = 1
                        MsgBox "��Ϊ���շ���Ŀ���ü۸�", vbExclamation, gstrSysName
                        Exit Function
                    Else
                        '��ȱʡ�������۸�ȼ����ɲ������շ���Ŀ
                        blnNothing = True
                    End If
                End If
                If blnNothing = False Then
                    For i = 1 To .Rows - 1
                        If .RowData(i) > 0 Then
                            For j = 1 To .Cols - 1
                                If Not IsNumeric(.TextMatrix(i, j)) And .ColWidth(j) > 0 Then
                                    ShowTab "�շѼ�Ŀ"
                                    tbPriceGrade.Item(k).Selected = True
                                    If .Active And .Visible Then
                                        .SetFocus
                                    End If
                                    .Row = i: .Col = j
                                    MsgBox "�շѼ�Ŀ��" & i & "��" & j + 1 & "��Ӧ������ֵ��", vbExclamation, gstrSysName
                                    Exit Function
                                End If
                            Next
                            If Val(.TextMatrix(i, mcstCol�ּ�)) < 0 Then
                                ShowTab "�շѼ�Ŀ"
                                tbPriceGrade.Item(k).Selected = True
                                MsgBox "�۸�����Ϊ���������ڵ� " & i & " ��������ȷ�ļ۸�", vbExclamation, gstrSysName
                                Exit Function
                            End If
                            
                            '�����Ŀ���ȱʡ�۸�
                            If Me.chk���.value = 1 Then
                                If Val(.TextMatrix(i, mcstColȱʡ�۸�)) > 0 Then
                                    If Val(.TextMatrix(i, mcstColȱʡ�۸�)) < Val(.TextMatrix(i, mcstColԭ��)) Or Val(.TextMatrix(i, mcstColȱʡ�۸�)) > Val(.TextMatrix(i, mcstCol�ּ�)) Then
                                        ShowTab "�շѼ�Ŀ"
                                        tbPriceGrade.Item(k).Selected = True
                                        MsgBox "ȱʡ�۸�Ӧ������ͼۺ���߼�֮�䣬���ڵ� " & i & " ��������ȷ��ȱʡ�۸�", vbExclamation, gstrSysName
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    Next
                    
                    If chk���.value = 0 And gstrҽ�۽ӿڱ�� <> "" And gbln����ҽ���շ���Ŀ Then
                        For i = 1 To .Rows - 1
                            If .RowData(i) > 0 Then
                                dbl�ϼƼ۸� = dbl�ϼƼ۸� + Val(.TextMatrix(i, mcstCol�ּ�))
                            End If
                        Next
                        
                        If dbl�ϼƼ۸� > mdbl����޼� Or dbl�ϼƼ۸� < mdbl����޼� Then
                            ShowTab "�շѼ�Ŀ"
                            tbPriceGrade.Item(k).Selected = True
                            MsgBox "�۸�����趨������޼�(" & Format(mdbl����޼�, "0.00") & ")������޼�(" & Format(mdbl����޼�, "0.00") & ")֮�䡣", vbExclamation, gstrSysName
                            Exit Function
                        End If
                    End If
                End If
            End With
            
            If blnNothing = False Then
                If Me.chkNow(k).value = 0 Then
                    If DateDiff("s", sys.Currentdate, Me.dtpBegin(k)) < 0 Then
                        MsgBox "����ִ��ʱ�䲻��С�ڵ�ǰʱ�䣡", vbInformation, gstrSysName
                        Me.dtpBegin(k).value = DateAdd("n", 1, sys.Currentdate)
                        If TabMain.Tabs.Count > 1 Then
                            TabMain.Tabs(2).Selected = True
                        End If
                        tbPriceGrade.Item(k).Selected = True
                        If Me.dtpBegin(k).Enabled = True Then
                            Me.dtpBegin(k).SetFocus
                        End If
                        tabMain_Click
                        Exit Function
                    End If
                End If
                If zlCommFun.StrIsValid(txt����˵��(k).Text, txt����˵��(k).MaxLength) = False Then
                    ShowTab "�շѼ�Ŀ"
                    tbPriceGrade.Item(k).Selected = True
                    If txt����˵��(k).Enabled And txt����˵��(k).Visible Then
                        txt����˵��(k).SetFocus
                        zlControl.TxtSelAll txt����˵��(k)
                    End If
                    Exit Function
                End If
            End If
        End If
    Next
    IsValid��Ŀ = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function IsValid����() As Boolean
    '����:����������Ŀҳ�����������Ƿ���Ч
    '����:intTab ҳ��
    '����ֵ:��Ч����True,����ΪFalse
    On Error GoTo ErrHandle
    Dim i As Integer
    
    IsValid���� = False
    With msh����
        For i = 1 To .Rows - 1
            If .RowData(i) > 0 Then
                If .TextMatrix(i, 1) = "" Then
                    ShowTab "������Ŀ"
                    If .Enabled And .Visible Then
                        .SetFocus
                    End If
                    .Row = i: .Col = 1
                    MsgBox "�����������", vbExclamation, gstrSysName
                    Exit Function
                End If
                If .TextMatrix(i, 2) = "" Then
                    ShowTab "������Ŀ"
                    If .Enabled And .Visible Then
                        .SetFocus
                    End If
                    .Row = i: .Col = 2
                    MsgBox "��ѡ�������ϵ��", vbExclamation, gstrSysName
                    Exit Function
                End If
                If .TextMatrix(i, 2) <> "0-���̶�" And Val(.TextMatrix(i, 1)) = 0 Then
                    ShowTab "������Ŀ"
                    If .Enabled And .Visible Then
                        .SetFocus
                    End If
                    .Row = i: .Col = 1
                    MsgBox "���ڹ̶���ϵ�����������Ϊ0��", vbExclamation, gstrSysName
                    Exit Function
                End If
            End If
        Next
    End With
    IsValid���� = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Function NumIsValid(ByVal strNumber As String) As Boolean
    '����:�������������Ƿ�����Ч������
    '����:strNumber  ��������
    '����ֵ:��Ч����True,����ΪFalse
    NumIsValid = False
    If Not IsNumeric(strNumber) Then
        MsgBox "������һ����ֵ��", vbExclamation, gstrSysName
        Exit Function
    End If
    If Val(strNumber) > 9999999999.999 Then
        MsgBox "�����̫���ˡ�", vbExclamation, gstrSysName
        Exit Function
    End If
    If Val(strNumber) < 0 Then
        MsgBox "����Ϊ������", vbExclamation, gstrSysName
        Exit Function
    End If
    NumIsValid = True
End Function

Private Function IsRecord(ByVal strTable As String, ByVal strWhere As String, _
    Optional ByVal Index As Long) As Boolean
    '����:�������������Ƿ�����Ч�����ݿ��б�ļ�¼
    '����:strTable ����;
    '     strWhere SQL��������
    '     index  ѡ��������Ŀʱ����
    '����ֵ:��Ч����True,����ΪFalse
    On Error GoTo ErrHandle
    Dim rsTmp As New ADODB.Recordset
    Dim strTemp As String
    Dim i As Integer
    Dim strReturn As String 'ѡ���������ַ���
    Dim strHyID As Long
    Dim strWherePriceGrade As String
    
    rsTmp.CursorLocation = adUseClient
    IsRecord = False
    If InStr(strWhere, "'") > 0 Then
        MsgBox "�����˷Ƿ��ַ���", vbExclamation, gstrSysName
        Exit Function
    End If
    If strTable = "������Ŀ" Then
        gstrSQL = "select ����,����,�վݷ�Ŀ,������Ŀ,id from ������Ŀ where ĩ��=1 and ( ���� like [1] or ���� like [1] or ���� like [2] ) and " & Where����ʱ��
    Else
        If gstr��ͨ�۸�ȼ� = "" And gstrҩƷ�۸�ȼ� = "" And gstr���ļ۸�ȼ� = "" Then
            strWherePriceGrade = " And d.�۸�ȼ� Is Null"
        Else
            strWherePriceGrade = "" & _
                " And ((Instr(';5;6;7;', ';' || a.��� || ';') > 0 And d.�۸�ȼ� = [4])" & vbNewLine & _
                "      Or (Instr(';4;', ';' || a.��� || ';') > 0 And d.�۸�ȼ� = [5])" & vbNewLine & _
                "      Or (Instr(';4;5;6;7;', ';' || a.��� || ';') = 0 And d.�۸�ȼ� = [6])" & vbNewLine & _
                "      Or (d.�۸�ȼ� Is Null" & vbNewLine & _
                "          And Not Exists (Select 1" & vbNewLine & _
                "                          From �շѼ�Ŀ" & vbNewLine & _
                "                          Where d.�շ�ϸĿid = �շ�ϸĿid And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
                "                                And ((Instr(';5;6;7;', ';' || a.��� || ';') > 0 And �۸�ȼ� = [4])" & vbNewLine & _
                "                                      Or (Instr(';4;', ';' || a.��� || ';') > 0 And �۸�ȼ� = [5])" & vbNewLine & _
                "                                      Or (Instr(';4;5;6;7;', ';' || a.��� || ';') = 0 And �۸�ȼ� = [6])))))"
        End If
        
        gstrSQL = _
            "SELECT A.����,A.����," & _
            "       A.���,A.���㵥λ,ltrim(rtrim(to_char(Sum(nvl(D.�ּ�,0)),'9999999990.00'))) �۸�,A.ID" & _
            " FROM(Select Distinct a.���,A.ID,A.����,A.����,A.���,A.���㵥λ" & _
            "       From �շ���ĿĿ¼ A,�շ���Ŀ���� B" & _
            "       WHERE A.ID = B.�շ�ϸĿID" & _
            "             And (A.����ʱ��=to_date('3000-01-01','yyyy-mm-dd') or A.����ʱ�� is null)" & _
            "             And (A.���� like [1] or A.���� like [1] or  ('['||A.����||']'||A.����  =[3])  or  B.���� like [2])" & _
            "   ) A,�շѼ�Ŀ D" & _
            " Where A.ID=D.�շ�ϸĿID(+)" & _
            "       And D.ִ������ <= SYSDATE AND (D.��ֹ���� > SYSDATE OR D.��ֹ���� IS NULL)" & _
                    strWherePriceGrade & vbNewLine & _
            " Group By A.����,A.����,A.���,A.���㵥λ,A.ID"
    End If
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strWhere & "%", "%" & UCase(strWhere) & "%", strWhere, _
        gstrҩƷ�۸�ȼ�, gstr���ļ۸�ȼ�, gstr��ͨ�۸�ȼ�)
    
    If rsTmp.RecordCount < 1 Then MsgBox "û���ҵ������ҵ��շ���Ŀ��", vbInformation, Me.Caption: Exit Function
    If rsTmp.RecordCount > 1 Then
        If strTable = "������Ŀ" Then
            strReturn = frmSelCur.ShowCurrSel(Me, rsTmp, "����,800,0,2;����,1500,0,2;�վݷ�Ŀ,1200,0,2;������Ŀ,1200,0,2;ID,0,1,2", "������Ŀѡ����", True, , , 800 + 1500 + 1200 + 1200 + 800)
        Else
            strReturn = frmSelCur.ShowCurrSel(Me, rsTmp, "����,1000,0,2;����,1500,0,2;���,1500,0,2;���㵥λ,800,0,2;�۸�,1000,1,2;ID,0,1,2", "�շ���Ŀѡ����", True, , , 1000 + 1500 + 1500 + 800 + 800 + 2000)
        End If
        If Trim(strReturn) = "" Then
            Exit Function
        End If
    Else
        If strTable = "������Ŀ" Then
            strReturn = Nvl(rsTmp!����) & "," & Nvl(rsTmp!����) & "," & Nvl(rsTmp!�վݷ�Ŀ) & "," & Nvl(rsTmp!������Ŀ) & "," & Nvl(rsTmp!ID, 0)
        Else
            strReturn = Nvl(rsTmp!����) & "," & Nvl(rsTmp!����) & "," & Nvl(rsTmp!���) & "," & Nvl(rsTmp!���㵥λ) & "," & Nvl(rsTmp!�۸�) & "," & Nvl(rsTmp!ID, 0)
        End If
    End If
    If strTable = "������Ŀ" Then
        On Error Resume Next
        With msh��Ŀ(Index)
            strTemp = Split(strReturn, ",")(UBound(Split(strReturn, ",")))
            If .RowData(.Row) <> strTemp Then
                mcol��Ŀ(Index).Add 0, "C" & strTemp
                If Err <> 0 Then
                    MsgBox "������Ŀ��" & Split(strReturn, ",")(1) & "���������˼�Ŀ��", vbExclamation, gstrSysName
                    Exit Function
                End If
                If .RowData(.Row) > 0 Then mcol��Ŀ(Index).Remove "C" & .RowData(.Row)
                .RowData(.Row) = CLng(strTemp)
            End If
            .TextMatrix(.Row, mcstCol�շ���Ŀ) = Split(strReturn, ",")(1)
            If .TextMatrix(.Row, mcstColԭ��) = "" Then .TextMatrix(.Row, mcstColԭ��) = "0.000"
        End With
    Else
        For i = 0 To msh����.Rows - 1
            If msh����.RowData(i) > 0 And msh����.RowData(i) = CLng(Split(strReturn, ",")(UBound(Split(strReturn, ",")))) And i <> msh����.Row Then
                MsgBox "�շ���Ŀ��" & Split(strReturn, ",")(1) & "���ѱ���Ϊ�������ˡ�", vbExclamation, gstrSysName
                Exit Function
            End If
        Next
        If Val(Split(strReturn, ",")(UBound(Split(strReturn, ",")))) = Val(mstrID) And Val(mstrID) > 0 Then
            MsgBox "�շ���Ŀ��������Ϊ�Լ��Ĵ�����Ŀ��", vbExclamation, gstrSysName
            Exit Function
        End If
        '�ݹ���
        strHyID = Split(strReturn, ",")(UBound(Split(strReturn, ",")))
        If CheckHypotaxis(strHyID) = True Then
            MsgBox "���շ���Ŀ�Ѵ��ڴ���������������Ϊ���ӹ�����", vbExclamation, gstrSysName
            Exit Function
        End If
        
        '�����������Ŀ���������Ŀ�ļ۸�ִ������ֻ�ܰ��յ���
        If mblnIsSpecialItem Then
            If Not IsRaiseByDate(Val(strHyID)) Then
                 MsgBox "[" & Split(strReturn, ",")(0) & "]" & Split(strReturn, ",")(1) & "�ļ۸�������ǰ�����ִ�еģ�������Ϊ������Ŀ��", vbOKOnly + vbInformation, gstrSysName
                 Exit Function
            End If
        End If
        
        msh����.RowData(msh����.Row) = CLng(Split(strReturn, ",")(UBound(Split(strReturn, ","))))
        msh����.TextMatrix(msh����.Row, 0) = "[" & Split(strReturn, ",")(0) & "]" & Split(strReturn, ",")(1)
        If msh����.TextMatrix(msh����.Row, 1) = "" Then
            msh����.TextMatrix(msh����.Row, 1) = "0"
            msh����.TextMatrix(msh����.Row, 2) = "0-���̶�"
        End If
        If gstr��ͨ�۸�ȼ� = "" And gstrҩƷ�۸�ȼ� = "" And gstr���ļ۸�ȼ� = "" Then
            strWherePriceGrade = " And b.�۸�ȼ� Is Null"
        Else
            strWherePriceGrade = "" & _
                " And ((Instr(';5;6;7;', ';' || a.��� || ';') > 0 And b.�۸�ȼ� = [2])" & vbNewLine & _
                "      Or (Instr(';4;', ';' || a.��� || ';') > 0 And b.�۸�ȼ� = [3])" & vbNewLine & _
                "      Or (Instr(';4;5;6;7;', ';' || a.��� || ';') = 0 And b.�۸�ȼ� = [4])" & vbNewLine & _
                "      Or (b.�۸�ȼ� Is Null" & vbNewLine & _
                "          And Not Exists (Select 1" & vbNewLine & _
                "                          From �շѼ�Ŀ" & vbNewLine & _
                "                          Where b.�շ�ϸĿid = �շ�ϸĿid And Sysdate Between ִ������ And Nvl(��ֹ����, To_Date('3000-01-01', 'YYYY-MM-DD'))" & vbNewLine & _
                "                                And ((Instr(';5;6;7;', ';' || a.��� || ';') > 0 And �۸�ȼ� = [2])" & vbNewLine & _
                "                                      Or (Instr(';4;', ';' || a.��� || ';') > 0 And �۸�ȼ� = [3])" & vbNewLine & _
                "                                      Or (Instr(';4;5;6;7;', ';' || a.��� || ';') = 0 And �۸�ȼ� = [4])))))"
        End If
        gstrSQL = "SELECT a.id,a.�Ƿ���,sum(b.ԭ��) ԭ��,sum(b.�ּ�) �ּ�," & vbCrLf & _
                "         Decode(nvl(a.�Ƿ���,0),1,ltrim(rtrim(to_char(sum(b.ԭ��),'9999999990.00')))||'��'||ltrim(rtrim(to_char(sum(b.�ּ�),'9999999990.00'))),ltrim(rtrim(to_char(sum(b.�ּ�),'9999999990.00'))))  AS  �۸� " & vbCrLf & _
                " FROM �շ���ĿĿ¼ a,�շѼ�Ŀ b " & vbCrLf & _
                " WHERE a.id=b.�շ�ϸĿid AND  a.id=[1] " & vbCrLf & _
                "       And b.ִ������ <= SYSDATE AND (b.��ֹ���� > SYSDATE OR b.��ֹ���� IS NULL)" & _
                        strWherePriceGrade & vbNewLine & _
                " GROUP BY a.id,a.�Ƿ���"
        Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, Val(msh����.RowData(msh����.Row)), _
            gstrҩƷ�۸�ȼ�, gstr���ļ۸�ȼ�, gstr��ͨ�۸�ȼ�)
        
        If rsTmp.RecordCount > 0 Then
             msh����.TextMatrix(msh����.Row, 3) = Trim(rsTmp!�۸�)
        End If
    End If
    IsRecord = True
    Exit Function
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Function

Private Sub GetDefineSize()
    '���ܣ��õ����ݿ�ı��ֶεĳ���
    On Error GoTo ErrHandle
    Dim rsTmp As New ADODB.Recordset
    gstrSQL = "Select A.����, A.���㵥λ, A.����, A.��ʶ����, A.��ʶ����, A.����޼�, A.����޼�, A.���, A.˵��, A.����, A.��ѡ��, B.���� ����, B.���� " & _
            " From �շ���ĿĿ¼ A, �շ���Ŀ���� B " & _
            " Where A.ID = B.�շ�ϸĿid And A.ID = 0 And B.���� = 1 "
    Call zlDatabase.OpenRecordset(rsTmp, gstrSQL, Me.Caption)
    
    mlng���볤�� = rsTmp.Fields("����").DefinedSize
    mlng��λ���� = rsTmp.Fields("���㵥λ").DefinedSize
    mlng���볤�� = rsTmp.Fields("����").DefinedSize
    mlng�������� = rsTmp.Fields("����").DefinedSize
    
    txtEdit(text����).MaxLength = mlng���볤��
    txtEdit(Text����).MaxLength = rsTmp.Fields("����").DefinedSize
    txtEdit(text��ʶ����).MaxLength = rsTmp.Fields("��ʶ����").DefinedSize
    txtEdit(Text��ʶ����).MaxLength = rsTmp.Fields("��ʶ����").DefinedSize
    txtEdit(text����޼�).MaxLength = rsTmp.Fields("����޼�").DefinedSize - 2
    txtEdit(text����޼�).MaxLength = rsTmp.Fields("����޼�").DefinedSize - 2
    txtEdit(Text���).MaxLength = rsTmp.Fields("���").DefinedSize
    txtEdit(Text˵��).MaxLength = rsTmp.Fields("˵��").DefinedSize
    txtEdit(text����).MaxLength = rsTmp.Fields("����").DefinedSize
    txtEdit(Text��ѡ��).MaxLength = rsTmp.Fields("��ѡ��").DefinedSize
    txtEdit(text����).MaxLength = mlng���볤��
    txtEdit(text���).MaxLength = mlng���볤��
    Exit Sub
ErrHandle:
    If ErrCenter() = 1 Then
        Resume
    End If
    Call SaveErrLog
End Sub

Private Sub txt����ִ��_KeyPress(KeyAscii As Integer)
    Dim strTemp As String
    Dim rsTmp As New ADODB.Recordset
    Dim ObjItem As ListItem
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Trim(Me.txt����ִ��.Text) = "" Then Me.txt����ִ��.Tag = "": Me.txt����ִ��.Text = "": Call OS.PressKey(vbKeyTab): Exit Sub
    
    strTemp = UCase(Me.txt����ִ��.Text)
    If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
    
    Err = 0: On Error GoTo ErrHand
    
    gstrSQL = "select distinct ID,����,����" & _
            " from ���ű� D,��������˵�� T" & _
            " where D.ID=T.����ID and T.������� in (1,2,3)" & _
            "       and (D.����ʱ�� is null or D.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))" & _
            "       and (D.���� like [1] or D.���� like [1] or D.���� like [1])" & _
            " order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strTemp & "%")
    
    With rsTmp
        If .BOF Or .EOF Then
            MsgBox "δ�ҵ�ָ�����ţ����������룡", vbExclamation, gstrSysName:  Exit Sub
        End If
        If .RecordCount = 1 Then
            Me.txt����ִ��.Tag = !ID: Me.txt����ִ��.Text = !����: Call OS.PressKey(vbKeyTab): Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Me.lvwItems.Checkboxes = False
        Do While Not .EOF
            Set ObjItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !����)
            ObjItem.Icon = "Dept": ObjItem.SmallIcon = "Dept"
            ObjItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = !����
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    
    With Me.picDept
        .Left = Me.TabMain.Left + Me.fra(3).Left + Me.txt����ִ��.Left - 130
        .Top = Me.TabMain.Top + Me.fra(3).Top + Me.txt����ִ��.Top + Me.txt����ִ��.Height - Me.Frame2.Top + 160
        
        lbl��������.Visible = False
        cboProperty.Visible = False
        ChkSelect.Visible = False
        
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

Private Sub txt����˵��_GotFocus(Index As Integer)
    zlControl.TxtSelAll txtEdit(Index)
End Sub

Private Sub txt����˵��_KeyPress(Index As Integer, KeyAscii As Integer)
    If InStr("'", Chr(KeyAscii)) > 0 Then KeyAscii = 0
    If KeyAscii = vbKeyReturn Or KeyAscii = vbKeyTab Then
        OS.PressKey vbKeyTab
    End If
End Sub

Private Sub txtסԺִ��_Change()
    If Trim(txtסԺִ��.Text) = "" Then
        txtסԺִ��.Tag = ""
    End If
End Sub

Private Sub txtסԺִ��_GotFocus()
    Me.txtסԺִ��.SelStart = 0: Me.txtסԺִ��.SelLength = 100
End Sub

Private Sub txtסԺִ��_KeyPress(KeyAscii As Integer)
    Dim ObjItem As ListItem
    Dim strTemp As String
    Dim rsTmp As New ADODB.Recordset
    
    If InStr(" ~!@#$%^&*_+|=-`;'""/?", Chr(KeyAscii)) > 0 Then KeyAscii = 0: Exit Sub
    If KeyAscii <> vbKeyReturn Then Exit Sub
    If Trim(Me.txtסԺִ��.Text) = "" Then Me.txtסԺִ��.Tag = "": Me.txtסԺִ��.Text = "": Call OS.PressKey(vbKeyTab): Exit Sub
    
    strTemp = UCase(Me.txtסԺִ��.Text)
    If InStr(1, strTemp, "[") <> 0 And InStr(1, strTemp, "]") <> 0 Then strTemp = Mid(strTemp, 2, InStr(1, strTemp, "]") - 2)
    
    Err = 0: On Error GoTo ErrHand
   
    gstrSQL = "select distinct ID,����,����" & _
            " from ���ű� D,��������˵�� T" & _
            " where D.ID=T.����ID and T.������� in (1,2,3)" & _
            "       and (D.����ʱ�� is null or D.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))" & _
            "       and (D.���� like [1] or D.���� like [1] or D.���� like [1])" & _
            " order by ����"
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, strTemp & "%")
        
    With rsTmp
        If .BOF Or .EOF Then
            MsgBox "δ�ҵ�ָ�����ţ����������룡", vbExclamation, gstrSysName: Exit Sub
        End If
        If .RecordCount = 1 Then
            Me.txtסԺִ��.Tag = !ID: Me.txtסԺִ��.Text = !����: Call OS.PressKey(vbKeyTab): Exit Sub
        End If
        Me.lvwItems.ListItems.Clear
        Me.lvwItems.Checkboxes = False
        Do While Not .EOF
            Set ObjItem = Me.lvwItems.ListItems.Add(, "_" & !ID, !����)
            ObjItem.Icon = "Dept": ObjItem.SmallIcon = "Dept"
            ObjItem.SubItems(Me.lvwItems.ColumnHeaders("����").Index - 1) = !����
            .MoveNext
        Loop
        Me.lvwItems.ListItems(1).Selected = True
    End With
    
    With Me.picDept
        .Left = Me.TabMain.Left + Me.fra(3).Left + Me.txtסԺִ��.Left - 1300
        .Top = Me.TabMain.Top + Me.fra(3).Top + Me.txtסԺִ��.Top + Me.txtסԺִ��.Height - Me.Frame2.Top + 130
        
        lbl��������.Visible = False
        cboProperty.Visible = False
        ChkSelect.Visible = False
        
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
Private Function CheckHypotaxis(HypotaxisID As Long) As Boolean
    ''''''''''''''''''''''''''''''''''''''''''
    '����           ��������Ŀ�Ƿ�ݹ�
    '����
    '               hypotaxisID������ĿID
    '����           Flase=û���ظ� True=���ظ�
    '''''''''''''''''''''''''''''''''''''''''
    Dim rsTmp As New ADODB.Recordset
    
    gstrSQL = "select 1 from �շѴ�����Ŀ where ����ID= [1] "
    
    On Error GoTo errH
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSQL, Me.Caption, HypotaxisID)
    
    If rsTmp.EOF = True Then
        CheckHypotaxis = False
    Else
        CheckHypotaxis = True
    End If
    Exit Function
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function
