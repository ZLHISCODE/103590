VERSION 5.00
Object = "{BEEECC20-4D5F-4F8B-BFDC-5D9B6FBDE09D}#1.0#0"; "vsflex8.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmInAdviceSetup 
   AutoRedraw      =   -1  'True
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "סԺҽ��ѡ��"
   ClientHeight    =   9030
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9675
   Icon            =   "frmInAdviceSetup.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9030
   ScaleWidth      =   9675
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin TabDlg.SSTab tabPar 
      Height          =   8340
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   14711
      _Version        =   393216
      Style           =   1
      Tabs            =   2
      TabHeight       =   617
      WordWrap        =   0   'False
      TabCaption(0)   =   "ҽ���´�(&1)"
      TabPicture(0)   =   "frmInAdviceSetup.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbl����ҩ��"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbl����"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lblȱʡҩ��"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "vsfDrugStore"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "fra��Ժ���"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "fraҽ���´�"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "fra��Һ����"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "cbo����"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "fraPurMed"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "fra��Ч"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).ControlCount=   10
      TabCaption(1)   =   "ҽ������(&2)"
      TabPicture(1)   =   "frmInAdviceSetup.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "fra�����ջ�"
      Tab(1).Control(1)=   "fra��������"
      Tab(1).Control(2)=   "fraУ�Բ���"
      Tab(1).Control(3)=   "fraBat"
      Tab(1).Control(4)=   "fraBaby"
      Tab(1).Control(5)=   "fraAdvicePrint"
      Tab(1).Control(6)=   "fraBillPrint"
      Tab(1).Control(7)=   "fraBloodPrint"
      Tab(1).ControlCount=   8
      Begin VB.Frame fraBloodPrint 
         Caption         =   "��Ѫ���뵥��ӡģʽ"
         Height          =   855
         Left            =   -68040
         TabIndex        =   75
         Top             =   5205
         Width           =   2295
         Begin VB.OptionButton optBloodPrintType 
            Caption         =   "�¿�ʱ��ӡ"
            Height          =   180
            Index           =   0
            Left            =   150
            TabIndex        =   77
            Top             =   600
            Width           =   1440
         End
         Begin VB.OptionButton optBloodPrintType 
            Caption         =   "����ʱ��ӡ"
            Height          =   180
            Index           =   1
            Left            =   150
            TabIndex        =   76
            Top             =   285
            Value           =   -1  'True
            Width           =   1545
         End
      End
      Begin VB.Frame fra��Ч 
         Caption         =   "������Һ�������ĵ�ҽ����Ч"
         Height          =   1185
         Left            =   6600
         TabIndex        =   50
         Top             =   480
         Width           =   2655
         Begin VB.OptionButton opt��Ч 
            Caption         =   "����������"
            Height          =   180
            Index           =   2
            Left            =   240
            TabIndex        =   53
            Top             =   840
            Value           =   -1  'True
            Width           =   1200
         End
         Begin VB.OptionButton opt��Ч 
            Caption         =   "����"
            Height          =   180
            Index           =   1
            Left            =   240
            TabIndex        =   52
            Top             =   570
            Width           =   680
         End
         Begin VB.OptionButton opt��Ч 
            Caption         =   "����"
            Height          =   180
            Index           =   0
            Left            =   240
            TabIndex        =   51
            Top             =   300
            Width           =   680
         End
      End
      Begin VB.Frame fraBillPrint 
         Caption         =   "ҽ�����ͺ�,���Ƶ���"
         Height          =   1980
         Left            =   -68055
         TabIndex        =   66
         Top             =   6195
         Width           =   2295
         Begin VB.OptionButton optPrint 
            Caption         =   "�Զ���ӡ"
            Height          =   180
            Index           =   2
            Left            =   150
            TabIndex        =   69
            Top             =   1140
            Value           =   -1  'True
            Width           =   1080
         End
         Begin VB.OptionButton optPrint 
            Caption         =   "ѡ���Ƿ��ӡ"
            Height          =   180
            Index           =   1
            Left            =   150
            TabIndex        =   68
            Top             =   760
            Width           =   1440
         End
         Begin VB.OptionButton optPrint 
            Caption         =   "����ӡ"
            Height          =   180
            Index           =   0
            Left            =   150
            TabIndex        =   67
            Top             =   400
            Width           =   840
         End
      End
      Begin VB.Frame fraAdvicePrint 
         Caption         =   "ҽ������ӡģʽ"
         Height          =   885
         Left            =   -70800
         TabIndex        =   63
         Top             =   5205
         Width           =   2535
         Begin VB.OptionButton optPrintType 
            Caption         =   "У�Ժ��ӡ"
            Height          =   180
            Index           =   0
            Left            =   150
            TabIndex        =   65
            Top             =   285
            Value           =   -1  'True
            Width           =   1545
         End
         Begin VB.OptionButton optPrintType 
            Caption         =   "�¿�ʱ��ӡ"
            Height          =   180
            Index           =   1
            Left            =   150
            TabIndex        =   64
            Top             =   600
            Width           =   1440
         End
      End
      Begin VB.Frame fraPurMed 
         Caption         =   "����ҩ��ȱʡ��ҩĿ��"
         Height          =   1185
         Left            =   4320
         TabIndex        =   59
         Top             =   480
         Width           =   2205
         Begin VB.OptionButton optPurMed 
            Caption         =   "�´�ʱȷ��"
            Height          =   180
            Index           =   0
            Left            =   255
            TabIndex        =   78
            Top             =   270
            Width           =   1635
         End
         Begin VB.OptionButton optPurMed 
            Caption         =   "Ԥ��"
            Height          =   180
            Index           =   1
            Left            =   255
            TabIndex        =   61
            Top             =   585
            Width           =   680
         End
         Begin VB.OptionButton optPurMed 
            Caption         =   "����"
            Height          =   180
            Index           =   2
            Left            =   255
            TabIndex        =   60
            Top             =   870
            Value           =   -1  'True
            Width           =   680
         End
      End
      Begin VB.Frame fraBaby 
         Caption         =   "ҽ������ȱʡ��Χ(������)"
         Height          =   1200
         Left            =   -70800
         TabIndex        =   55
         Top             =   6975
         Width           =   2535
         Begin VB.OptionButton optBaby 
            Caption         =   "Ӥ��ҽ��"
            Height          =   180
            Index           =   2
            Left            =   150
            TabIndex        =   58
            Top             =   900
            Width           =   1440
         End
         Begin VB.OptionButton optBaby 
            Caption         =   "ȫ��ҽ��"
            Height          =   180
            Index           =   0
            Left            =   150
            TabIndex        =   57
            Top             =   285
            Value           =   -1  'True
            Width           =   1545
         End
         Begin VB.OptionButton optBaby 
            Caption         =   "����ҽ��"
            Height          =   180
            Index           =   1
            Left            =   150
            TabIndex        =   56
            Top             =   592
            Width           =   1440
         End
      End
      Begin VB.ComboBox cbo���� 
         ForeColor       =   &H80000012&
         Height          =   300
         Left            =   1425
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   500
         Width           =   2790
      End
      Begin VB.Frame fra��Һ���� 
         Caption         =   " ����Һ�������ķ�ҩ�Ĳ��˿��� "
         Height          =   2655
         Left            =   4320
         TabIndex        =   42
         Top             =   1800
         Width           =   4935
         Begin VB.CheckBox chkWithSelf 
            Caption         =   "�Ա�ҩ����ȡҩ����Ժ��ҩ�����͵���������"
            Height          =   195
            Left            =   210
            TabIndex        =   83
            Top             =   2370
            Width           =   4485
         End
         Begin VB.CheckBox chk����Ӫ�� 
            Caption         =   "�������Ĳ����յľ���Ӫ��ҽ���ڲ�������"
            Height          =   195
            Left            =   210
            TabIndex        =   79
            Top             =   2115
            Width           =   4485
         End
         Begin VB.CommandButton cmdAllDel 
            Caption         =   "ȫ��"
            Height          =   350
            Left            =   3720
            TabIndex        =   45
            Top             =   720
            Width           =   1100
         End
         Begin VB.CommandButton cmdAllSelect 
            Caption         =   "ȫѡ"
            Height          =   350
            Left            =   3720
            TabIndex        =   44
            Top             =   240
            Width           =   1100
         End
         Begin VB.ListBox lstDept 
            ForeColor       =   &H80000012&
            Height          =   1740
            IMEMode         =   3  'DISABLE
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   43
            Top             =   300
            Width           =   3465
         End
      End
      Begin VB.Frame fraBat 
         Caption         =   " ������������ "
         Height          =   690
         Left            =   -70800
         TabIndex        =   33
         Top             =   6195
         Width           =   2535
         Begin VB.CheckBox chkBat 
            Caption         =   "��ͣ/����"
            Height          =   195
            Index           =   1
            Left            =   1200
            TabIndex        =   35
            Top             =   360
            Width           =   1110
         End
         Begin VB.CheckBox chkBat 
            Caption         =   "У��"
            Height          =   195
            Index           =   0
            Left            =   240
            TabIndex        =   34
            Top             =   360
            Width           =   660
         End
      End
      Begin VB.Frame fraУ�Բ��� 
         Caption         =   " �Զ�У��ҽ���Ĳ��� "
         Height          =   4530
         Left            =   -70800
         TabIndex        =   29
         Top             =   480
         Width           =   5040
         Begin VB.ListBox lstУ�Բ��� 
            ForeColor       =   &H80000012&
            Height          =   4050
            Left            =   165
            Style           =   1  'Checkbox
            TabIndex        =   32
            Top             =   270
            Width           =   3510
         End
         Begin VB.CommandButton cmdУ�Բ���ALL 
            Caption         =   "ȫѡ"
            Height          =   350
            Left            =   3840
            TabIndex        =   31
            ToolTipText     =   "Ctrl+A"
            Top             =   240
            Width           =   1100
         End
         Begin VB.CommandButton cmdУ�Բ���Clear 
            Caption         =   "ȫ��"
            Height          =   350
            Left            =   3840
            TabIndex        =   30
            ToolTipText     =   "Ctrl+R"
            Top             =   720
            Width           =   1100
         End
      End
      Begin VB.Frame fra�������� 
         Caption         =   " ������������� "
         Height          =   4530
         Left            =   -74880
         TabIndex        =   21
         Top             =   480
         Width           =   3945
         Begin MSComCtl2.DTPicker dtpEnd 
            BeginProperty DataFormat 
               Type            =   0
               Format          =   "yyyy-M-d"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   2052
               SubFormatType   =   0
            EndProperty
            Height          =   300
            Left            =   2535
            TabIndex        =   85
            Top             =   4155
            Width           =   1095
            _ExtentX        =   1931
            _ExtentY        =   529
            _Version        =   393216
            Enabled         =   0   'False
            Format          =   97910786
            CurrentDate     =   42017
         End
         Begin VB.CheckBox chkInfo 
            Caption         =   "�����ڷ�ҩ���ͽ���ʱ��"
            Height          =   195
            Index           =   0
            Left            =   210
            TabIndex        =   84
            Top             =   4245
            Width           =   2325
         End
         Begin VB.CheckBox chkLimit 
            Caption         =   "ҩƷ�����ĸ�ҩ;�����ʹ����Խ���ʱ��Ϊ׼����"
            Height          =   420
            Left            =   210
            TabIndex        =   74
            Top             =   2865
            Width           =   3225
         End
         Begin VB.CheckBox chk�ر�ҽ�� 
            Caption         =   "������ɺ�ر�ҽ������"
            Height          =   195
            Left            =   210
            TabIndex        =   47
            Top             =   2130
            Width           =   3165
         End
         Begin VB.CheckBox chkShort 
            Caption         =   "����"
            Height          =   255
            Left            =   1200
            TabIndex        =   41
            Top             =   1155
            Width           =   975
         End
         Begin VB.CheckBox chkLong 
            Caption         =   "����"
            Height          =   255
            Left            =   480
            TabIndex        =   40
            Top             =   1155
            Width           =   975
         End
         Begin VB.CheckBox chkƤ�� 
            Caption         =   "��дƤ�Խ��ʱ��֤���"
            Height          =   195
            Left            =   210
            TabIndex        =   37
            Top             =   3945
            Width           =   2355
         End
         Begin VB.CheckBox chkУ��ǩ�� 
            Caption         =   "У�Ժ�ȷ��ֹͣʱʹ�õ���ǩ��"
            Height          =   195
            Left            =   210
            TabIndex        =   36
            Top             =   3645
            Width           =   3165
         End
         Begin VB.CheckBox chkҽ�� 
            Caption         =   "�����ҽ���´��ҽ�����к�������"
            Height          =   195
            Left            =   210
            TabIndex        =   28
            Top             =   3360
            Width           =   3180
         End
         Begin VB.CheckBox chk��ӡ 
            Caption         =   "У��,ȷ��ֹͣ,����ҽ������д�ӡ"
            Height          =   405
            Left            =   210
            TabIndex        =   27
            Top             =   525
            Width           =   3180
         End
         Begin VB.CheckBox chkУ�� 
            Caption         =   "�¿�ҽ�����Զ�У�ԼƼ�"
            Height          =   195
            Left            =   210
            TabIndex        =   26
            Top             =   330
            Width           =   3180
         End
         Begin VB.CheckBox chkִ�� 
            Caption         =   "����ʱ������ִ�е���Ŀ��Ϊ��ִ��"
            Height          =   195
            Left            =   210
            TabIndex        =   25
            Top             =   930
            Width           =   3180
         End
         Begin VB.CheckBox chkҽ������ 
            Caption         =   "����ʱ��ҽ�����˼����Ŀ�Ƿ�����"
            Height          =   195
            Left            =   210
            TabIndex        =   24
            Top             =   1830
            Value           =   1  'Checked
            Width           =   3180
         End
         Begin VB.CheckBox chkAutoVerify 
            Caption         =   "����У�Լ��ɷ���ҽ��"
            Height          =   180
            Left            =   210
            TabIndex        =   23
            Top             =   1530
            Width           =   3180
         End
         Begin VB.CheckBox chkTurnCheck 
            Caption         =   "����δУ��ҽ��������͵�ҽ��ʱ��ֹ����ת�ơ���Ժ��תԺ������ҽ��"
            Height          =   405
            Left            =   210
            TabIndex        =   22
            Top             =   2415
            Width           =   3180
         End
      End
      Begin VB.Frame fra�����ջ� 
         Caption         =   " �����ջ� "
         Height          =   2970
         Left            =   -74880
         TabIndex        =   16
         Top             =   5205
         Width           =   3945
         Begin VB.OptionButton optRoll 
            Caption         =   "��������"
            Height          =   255
            Index           =   1
            Left            =   2640
            TabIndex        =   73
            Top             =   300
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.OptionButton optRoll 
            Caption         =   "��������"
            Height          =   255
            Index           =   0
            Left            =   1440
            TabIndex        =   72
            Top             =   300
            Width           =   1095
         End
         Begin VB.CheckBox chk������� 
            Caption         =   "�����ջ�ʱ�Զ���˱���ִ�е���������"
            Height          =   195
            Left            =   210
            TabIndex        =   19
            Top             =   600
            Width           =   3680
         End
         Begin VB.CheckBox chkAutoRoll 
            Caption         =   "ȷ��ֹͣ���Զ�ִ�г����ջ�"
            Height          =   195
            Left            =   210
            TabIndex        =   18
            Top             =   900
            Width           =   3180
         End
         Begin VB.ListBox lst��ҩ���� 
            Columns         =   3
            ForeColor       =   &H80000012&
            Height          =   1320
            IMEMode         =   3  'DISABLE
            ItemData        =   "frmInAdviceSetup.frx":0044
            Left            =   210
            List            =   "frmInAdviceSetup.frx":0046
            Style           =   1  'Checkbox
            TabIndex        =   17
            Top             =   1500
            Width           =   3525
         End
         Begin VB.Label lblRoll 
            Caption         =   "�����ջ�ģʽ"
            Height          =   255
            Left            =   240
            TabIndex        =   71
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label lblSend 
            Caption         =   "���·�ҩ��ʽ����ҩһ����ҩ�Ͳ��ջ�"
            Height          =   255
            Left            =   240
            TabIndex        =   20
            Top             =   1230
            Width           =   3255
         End
      End
      Begin VB.Frame fraҽ���´� 
         Caption         =   " �������� "
         Height          =   3585
         Left            =   4320
         TabIndex        =   9
         Top             =   4575
         Width           =   4935
         Begin VB.CommandButton cmdBloodTip 
            Caption         =   "��Ѫ����ע����������"
            Height          =   350
            Left            =   225
            TabIndex        =   86
            Top             =   3180
            Width           =   2490
         End
         Begin VB.OptionButton optSTCheck 
            Caption         =   "�������Ľ��յ�ҩƷ"
            Height          =   255
            Index           =   1
            Left            =   2970
            TabIndex        =   82
            Top             =   1845
            Width           =   1920
         End
         Begin VB.OptionButton optSTCheck 
            Caption         =   "����ҩƷ"
            Height          =   255
            Index           =   0
            Left            =   1900
            TabIndex        =   81
            Top             =   1845
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.CheckBox chkST 
            Caption         =   "�Զ�����Ƥ�Բ����ݽ������ҽ������"
            Height          =   225
            Left            =   210
            TabIndex        =   62
            Top             =   1575
            Width           =   3555
         End
         Begin VB.CheckBox chk����ס����ҽ���´� 
            Caption         =   "���������ס�����´�ҽ��"
            Height          =   195
            Left            =   210
            TabIndex        =   54
            Top             =   2400
            Value           =   1  'Checked
            Width           =   2895
         End
         Begin VB.CheckBox chkͣ����� 
            Caption         =   $"frmInAdviceSetup.frx":0048
            Height          =   195
            Left            =   210
            TabIndex        =   46
            Top             =   2640
            Width           =   3180
         End
         Begin VB.CommandButton cmdAdviceSortSet 
            Caption         =   "�����������(&S)"
            Height          =   350
            Left            =   3120
            TabIndex        =   39
            Top             =   2800
            Width           =   1695
         End
         Begin VB.CheckBox chkAdviceSort 
            Caption         =   "����ҽ��ʱ�Զ�����"
            Height          =   255
            Left            =   210
            TabIndex        =   38
            Top             =   2880
            Width           =   1935
         End
         Begin VB.CheckBox chk����ҽ�� 
            Caption         =   "����ִ����ɺ�������´�����ҽ��"
            Height          =   195
            Left            =   210
            TabIndex        =   15
            Top             =   1350
            Width           =   3180
         End
         Begin VB.CheckBox chk���䵥�� 
            Caption         =   "�´�����ʱ�����뵥��"
            Height          =   195
            Left            =   210
            TabIndex        =   14
            Top             =   500
            Width           =   3180
         End
         Begin VB.CheckBox chk��Ժ��� 
            Caption         =   "�´��Ժҽ��ʱ����Ժ��ϵ���д"
            Height          =   195
            Left            =   210
            TabIndex        =   13
            Top             =   990
            Width           =   3180
         End
         Begin VB.CheckBox chkһ���� 
            Caption         =   "������ִ��Ƶ��ȱʡΪһ����"
            Height          =   195
            Left            =   210
            TabIndex        =   12
            Top             =   285
            Width           =   3180
         End
         Begin VB.CheckBox chk���� 
            Caption         =   "�´�ҩƷ����ʱ����ָ����ҩ����"
            Height          =   195
            Left            =   210
            TabIndex        =   11
            Top             =   750
            Width           =   3180
         End
         Begin VB.CheckBox chkStopNurseGrade 
            Caption         =   "������ֹͣ����ȼ�ҽ��"
            Height          =   195
            Left            =   210
            TabIndex        =   10
            ToolTipText     =   "������ʱ��ֻ��ͨ��ת�ơ���Ժ�����´��µ�ҽ����ֹͣ��ʿ�ȼ�"
            Top             =   2160
            Value           =   1  'Checked
            Width           =   3180
         End
         Begin VB.Label lblSTCheck 
            Caption         =   "����ҽ�������ͣ�"
            Height          =   255
            Left            =   480
            TabIndex        =   80
            Top             =   1845
            Width           =   1455
         End
      End
      Begin VB.Frame fra��Ժ��� 
         Height          =   1480
         Left            =   120
         TabIndex        =   4
         Top             =   6680
         Width           =   4095
         Begin VB.ListBox lst��Ժ��� 
            Columns         =   3
            Enabled         =   0   'False
            ForeColor       =   &H80000012&
            Height          =   1110
            IMEMode         =   3  'DISABLE
            Left            =   120
            Style           =   1  'Checkbox
            TabIndex        =   6
            Top             =   260
            Width           =   3900
         End
         Begin VB.CheckBox chk��Ժ��� 
            Caption         =   "�´���Щ����ҽ��ʱ����Ƿ���д���"
            Height          =   195
            Left            =   120
            TabIndex        =   5
            Top             =   0
            Width           =   3720
         End
      End
      Begin VSFlex8Ctl.VSFlexGrid vsfDrugStore 
         Height          =   4785
         Left            =   120
         TabIndex        =   7
         Top             =   1800
         Width           =   4095
         _cx             =   7223
         _cy             =   8440
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
         BackColorBkg    =   14737632
         BackColorAlternate=   -2147483643
         GridColor       =   -2147483633
         GridColorFixed  =   -2147483632
         TreeColor       =   -2147483632
         FloodColor      =   192
         SheetBorder     =   -2147483642
         FocusRect       =   1
         HighLight       =   1
         AllowSelection  =   0   'False
         AllowBigSelection=   -1  'True
         AllowUserResizing=   1
         SelectionMode   =   0
         GridLines       =   1
         GridLinesFixed  =   2
         GridLineWidth   =   1
         Rows            =   4
         Cols            =   4
         FixedRows       =   1
         FixedCols       =   1
         RowHeightMin    =   300
         RowHeightMax    =   0
         ColWidthMin     =   0
         ColWidthMax     =   0
         ExtendLastCol   =   -1  'True
         FormatString    =   $"frmInAdviceSetup.frx":0066
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
      Begin VB.Label lblȱʡҩ�� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ȱʡ�Ϳ���ҩ��"
         Height          =   180
         Left            =   120
         TabIndex        =   70
         Top             =   960
         Width           =   1260
      End
      Begin VB.Label lbl���� 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "ȱʡ���ϲ���"
         Height          =   180
         Left            =   120
         TabIndex        =   49
         Top             =   560
         Width           =   1080
      End
      Begin VB.Label lbl����ҩ�� 
         Caption         =   $"frmInAdviceSetup.frx":00EF
         Height          =   615
         Left            =   120
         TabIndex        =   8
         Top             =   1200
         Width           =   4095
      End
   End
   Begin VB.PictureBox picBottom 
      Align           =   2  'Align Bottom
      AutoRedraw      =   -1  'True
      BorderStyle     =   0  'None
      Height          =   530
      Left            =   0
      ScaleHeight     =   525
      ScaleWidth      =   9675
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   8505
      Width           =   9675
      Begin VB.CommandButton cmdOK 
         Caption         =   "ȷ��(&O)"
         Height          =   350
         Left            =   7320
         TabIndex        =   0
         Top             =   60
         Width           =   1100
      End
      Begin VB.CommandButton cmdCancel 
         Cancel          =   -1  'True
         Caption         =   "ȡ��(&C)"
         Height          =   350
         Left            =   8490
         TabIndex        =   1
         Top             =   60
         Width           =   1100
      End
   End
End
Attribute VB_Name = "frmInAdviceSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mMainPrivs As String
Public mint���� As Integer  '���ó���:0-ҽ��վ����,1-��ʿվ����,2-ҽ��վ����
Private Const VsPubBackColor = &HFAEADA
Private mblnTmp As Boolean

Private Enum mCtlID
    chk�����ڷ�ҩ���ͽ���ʱ�� = 0
End Enum

Private Sub chkAdviceSort_Click()
    cmdAdviceSortSet.Enabled = chkAdviceSort.value = 1 And chkAdviceSort.Enabled
End Sub

Private Sub chkInfo_Click(Index As Integer)
    Select Case Index
    Case chk�����ڷ�ҩ���ͽ���ʱ��
        dtpEnd.Enabled = chkInfo(Index).value = 1
    End Select
End Sub

Private Sub chkLong_Click()
    mblnTmp = True
    If chkShort.value = 0 Then chkִ��.value = chkLong.value
    mblnTmp = False
End Sub

Private Sub chkShort_Click()
    mblnTmp = True
    If chkLong.value = 0 Then chkִ��.value = chkShort.value
    mblnTmp = False
End Sub

Private Sub chkST_Click()
    If chkST.value Then
        optSTCheck(0).Enabled = True
        optSTCheck(1).Enabled = True
    Else
        optSTCheck(0).Enabled = False
        optSTCheck(1).Enabled = False
    End If
End Sub

Private Sub cmdBloodTip_Click()
    Dim strPar As String
    strPar = cmdBloodTip.Tag
    Call frmInputBox.InputBox(Me, "��Ѫ����ע������", "���ݣ�", 4000, 6, True, True, strPar)
    cmdBloodTip.Tag = strPar
End Sub

Private Sub optRoll_Click(Index As Integer)
    '��������ʱ����ʹ���Զ�������뵥
    If Index = 0 Then
        chk�������.Enabled = False
        chk�������.value = 0
    Else
        chk�������.Enabled = True
    End If
End Sub

Private Sub chk��Ժ���_Click()
    lst��Ժ���.Enabled = chk��Ժ���.value = 1 And lst��Ժ���.Tag = ""
End Sub

Private Sub chkУ��_Click()
    fraУ�Բ���.Enabled = chkУ��.value = 1
    cmdУ�Բ���ALL.Enabled = fraУ�Բ���.Enabled
    cmdУ�Բ���Clear.Enabled = fraУ�Բ���.Enabled
End Sub

Private Sub chkִ��_Click()
    If mblnTmp Then Exit Sub
    chkLong.value = chkִ��.value
    chkShort.value = chkִ��.value
End Sub

Private Sub cmdAdviceSortSet_Click()
    frmPathSetup.mbytFun = 1
    frmPathSetup.Show vbModal, Me
End Sub

Private Sub cmdAllDel_Click()
    Dim i As Long
    
    For i = 0 To lstDept.ListCount - 1
        lstDept.Selected(i) = False
    Next
End Sub

Private Sub cmdAllSelect_Click()
    Dim i As Long, Y As Long
    
    Y = lstDept.ListIndex
    For i = 0 To lstDept.ListCount - 1
        lstDept.Selected(i) = True
    Next
    lstDept.ListIndex = Y
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    Dim str��� As String, str���� As String, i As Long
    Dim strValue As String, bytType As Long
    Dim arr����ҩ��(3) As String, arrȱʡҩ��(3) As String, arrTmp() As String
    Dim str��Һ�������� As String, bln��Һȫѡ As Boolean
    Dim blnSetup As Boolean, blnSetPara As Boolean
    
     '������Ƿ�ָ����ȱʡҩ������Ϊ����û�в�������Ȩ�ޣ����������ǿ��Զ���ġ�
     
    If fra��Ժ���.Visible And chk��Ժ���.value = 1 Then
        For i = 0 To lst��Ժ���.ListCount - 1
            If lst��Ժ���.Selected(i) Then
                str��� = str��� & Chr(lst��Ժ���.ItemData(i))
            End If
        Next
        If str��� = "" Then
            MsgBox "������ѡ��һ��Ҫ�����Ժ��ϵ�ҽ�����", vbInformation, gstrSysName
            lst��Ժ���.SetFocus: Exit Sub
        End If
    End If
    If fraУ�Բ���.Visible And fraУ�Բ���.Enabled And chkУ��.value = 1 Then
        For i = 0 To lstУ�Բ���.ListCount - 1
            If lstУ�Բ���.Selected(i) Then
                str���� = str���� & "," & lstУ�Բ���.ItemData(i)
            End If
        Next
        str���� = Mid(str����, 2)
        If str���� = "" Then
            MsgBox "������ѡ��һ��Ҫ�Զ���ҽ������У�ԼƼ۵Ĳ�����", vbInformation, gstrSysName
            lstУ�Բ���.SetFocus: Exit Sub
        End If
    End If
    
    blnSetup = InStr(GetInsidePrivs(pסԺҽ���´�), ";ҽ��ѡ������;") > 0
    blnSetPara = InStr(GetInsidePrivs(pסԺҽ������), ";ҽ��ѡ������;") > 0
    
    Call zlDatabase.SetPara("����ȱʡһ����", chkһ����.value, glngSys, pסԺҽ���´�, blnSetup)
    Call zlDatabase.SetPara("ҽ��ִ������", chk����.value, glngSys, pסԺҽ���´�, blnSetup)
    Call zlDatabase.SetPara("�Զ�����Ƥ��", chkST.value, glngSys, pסԺҽ���´�, blnSetup)
    Call zlDatabase.SetPara("����Ƥ�Խ������ҽ����������", IIF(chkST.value = 1, IIF(optSTCheck(0).value, 0, 1), 0), glngSys, pסԺҽ���´�, blnSetup)
    Call zlDatabase.SetPara("���������뵥��", chk���䵥��.value, glngSys, pסԺҽ���´�, blnSetup)
    Call zlDatabase.SetPara("Ҫ��������Ժ���", str���, glngSys, pסԺҽ���´�, blnSetup)
    Call zlDatabase.SetPara("������ɺ��´�����ҽ��", chk����ҽ��.value, glngSys, pסԺҽ���´�, blnSetup)
    Call zlDatabase.SetPara("ҽ���Զ�����", chkAdviceSort.value, glngSys, pסԺҽ���´�, blnSetup)
    Call zlDatabase.SetPara("���������ס�����´�ҽ��", chk����ס����ҽ���´�.value, glngSys, pסԺҽ���´�, blnSetup)
    'ҽ������ӡģʽ
    Call zlDatabase.SetPara("ҽ������ӡģʽ", IIF(optPrintType(1).value, 1, 0), glngSys, pסԺҽ���´�, blnSetup)
    Call zlDatabase.SetPara("��Ѫ���뵥��ӡģʽ", IIF(optBloodPrintType(1).value, 1, 2), glngSys, pסԺҽ������, blnSetup)
    If mint���� <> 2 Then
        Call zlDatabase.SetPara("Ҫ�������Ժ���", chk��Ժ���.value, glngSys, pסԺҽ���´�, blnSetup)
        Call zlDatabase.SetPara("����ֹͣ����ȼ�", chkStopNurseGrade.value, glngSys, pסԺҽ���´�, blnSetup)
        Call zlDatabase.SetPara("ʵϰҽ��ֹͣҽ����Ҫ���", chkͣ�����.value, glngSys, pסԺҽ���´�, blnSetup)
    End If
    
    If chk�ر�ҽ��.Enabled = True Then
        Call zlDatabase.SetPara("������ɺ�ر�ҽ������", chk�ر�ҽ��.value, glngSys, pסԺҽ���´�, blnSetup)
    End If
    
    If mint���� = 1 Then
        If chkУ��.value = 0 Then
            Call zlDatabase.SetPara("�Զ����У�ԼƼ�", "", glngSys, pסԺҽ������, blnSetPara)
        ElseIf UBound(Split(str����, ",")) + 1 = lstУ�Բ���.ListCount Then
            Call zlDatabase.SetPara("�Զ����У�ԼƼ�", "*", glngSys, pסԺҽ������, blnSetPara)
        Else
            Call zlDatabase.SetPara("�Զ����У�ԼƼ�", str����, glngSys, pסԺҽ������, blnSetPara)
        End If
        
        Call zlDatabase.SetPara("����ִ���Զ����", chkLong.value & chkShort.value, glngSys, pסԺҽ������, blnSetPara)
        Call zlDatabase.SetPara("�Զ�����ҽ����ӡ", chk��ӡ.value, glngSys, pסԺҽ������, blnSetPara)
        Call zlDatabase.SetPara("ҽ��ҽ����������", chkҽ��.value, glngSys, pסԺҽ������, blnSetPara)
        Call zlDatabase.SetPara("Ƥ����֤���", chkƤ��.value, glngSys, pסԺҽ������, blnSetPara)
        Call zlDatabase.SetPara("�����ջز�����������", IIF(optRoll(0).value, 1, 0), glngSys, pסԺҽ������, blnSetPara)
        Call zlDatabase.SetPara("�����ջط��ñ����Զ����", chk�������.value, glngSys, pסԺҽ������, blnSetPara)
        Call zlDatabase.SetPara("У��ҽ������ǩ��", chkУ��ǩ��.value, glngSys, pסԺҽ������, blnSetPara)
        Call zlDatabase.SetPara("���ҽ������", chkҽ������.value, glngSys, pסԺҽ������, blnSetPara)
        Call zlDatabase.SetPara("����ǰ�Զ�У��", chkAutoVerify.value, glngSys, pסԺҽ������, blnSetPara)
        Call zlDatabase.SetPara("ֹͣ���Զ������ջ�", chkAutoRoll.value, glngSys, pסԺҽ������, blnSetPara)
        Call zlDatabase.SetPara("����ҽ������ǰ���δ��Чҽ��", chkTurnCheck.value, glngSys, pסԺҽ������, blnSetPara)
        Call zlDatabase.SetPara("ҩ���������ƽ���ʱ��", chkLimit.value, glngSys, pסԺҽ������, blnSetPara)
        
        '��������
        Call zlDatabase.SetPara("����ҽ��У��", chkBat(0).value, glngSys, pסԺҽ������, blnSetPara)
        Call zlDatabase.SetPara("����ҽ����ͣ", chkBat(1).value, glngSys, pסԺҽ������, blnSetPara)
        
        '��ҩ���Ͳ��ջ�
        strValue = ""
        For i = 0 To lst��ҩ����.ListCount - 1
            If lst��ҩ����.Selected(i) Then
                strValue = strValue & "," & ZLCommFun.GetNeedName(lst��ҩ����.List(i))
            End If
        Next
        strValue = Mid(strValue, 2)
        Call zlDatabase.SetPara("��ҩ���ջ�", strValue, glngSys, pסԺҽ������, blnSetPara)
        
        If chkInfo(chk�����ڷ�ҩ���ͽ���ʱ��).value = 1 Then
            strValue = "1|" & dtpEnd.value
        Else
            strValue = 0
        End If
        Call zlDatabase.SetPara("�����ڷ�ҩ���ͽ���ʱ��", strValue, glngSys, pסԺҽ������, blnSetPara)
    End If
    
    '����Һ�������ķ�ҩ�Ĳ��˿���
    bln��Һȫѡ = True
    If lstDept.Enabled Then
        For i = 0 To lstDept.ListCount - 1
            If lstDept.Selected(i) Then
                str��Һ�������� = str��Һ�������� & "," & lstDept.ItemData(i)
            Else
                bln��Һȫѡ = False
            End If
        Next
    End If
    
    If str��Һ�������� <> "" Then
        str��Һ�������� = Mid(str��Һ��������, 2)
    Else
        str��Һ�������� = ","
    End If
    If bln��Һȫѡ Then str��Һ�������� = "*"
    Call zlDatabase.SetPara("����Һ�������ķ�ҩ�Ĳ��˿���", str��Һ��������, glngSys, pסԺҽ���´�, blnSetup)
    
    '������Һ�������ĵ�ҽ����Ч
    If fra��Ч.Enabled Then
        For i = 0 To 2
            If opt��Ч(i).value Then
                Call zlDatabase.SetPara("������Һ�������ĵ�ҽ����Ч", i & "", glngSys, pסԺҽ���´�, blnSetup)
                Exit For
            End If
        Next
    End If
    Call zlDatabase.SetPara("�������Ĳ����յľ���Ӫ��ҽ���ڲ�������", chk����Ӫ��.value, glngSys, pסԺҽ���´�, blnSetup)
    Call zlDatabase.SetPara("��������ҩƷ�����͵���������", chkWithSelf.value, glngSys, pסԺҽ���´�, blnSetup)

    '����ҩ��ȱʡ��ҩĿ��
    For i = 0 To 2
        If optPurMed(i).value Then
            Call zlDatabase.SetPara("����ҩ��ȱʡ��ҩĿ��", i & "", glngSys, pסԺҽ���´�, blnSetup)
            Exit For
        End If
    Next
    
     'ҩ��
    With vsfDrugStore
        For i = .FixedRows To .Rows - 1
            Select Case .TextMatrix(i, .ColIndex("���"))
            Case "��ҩ��"
                bytType = 0
            Case "��ҩ��"
                bytType = 1
            Case "��ҩ��"
                bytType = 2
            End Select
            If .TextMatrix(i, .ColIndex("����")) <> 0 Then arr����ҩ��(bytType) = arr����ҩ��(bytType) & "," & .RowData(i)
            If .TextMatrix(i, .ColIndex("ȱʡ")) = "��" Then arrȱʡҩ��(bytType) = .RowData(i)
        Next
    End With
    arrTmp = Split("��ҩ��,��ҩ��,��ҩ��", ",")
    For bytType = 0 To UBound(arrTmp)
        Call zlDatabase.SetPara("סԺ����" & arrTmp(bytType), Mid(arr����ҩ��(bytType), 2), glngSys, pסԺҽ���´�, blnSetup)
        Call zlDatabase.SetPara("סԺȱʡ" & arrTmp(bytType), arrȱʡҩ��(bytType), glngSys, pסԺҽ���´�, blnSetup)
    Next
        
    Call zlDatabase.SetPara("סԺȱʡ���ϲ���", IIF(cbo����.ListIndex = 0, "0", cbo����.ItemData(cbo����.ListIndex)), glngSys, pסԺҽ���´�, blnSetup)
    Call zlDatabase.SetPara("��Ѫ����ע������", cmdBloodTip.Tag, glngSys, pסԺҽ���´�, blnSetup)
    
    '���ݴ�ӡ:0-����ӡ,1-�ֹ���ӡ,2-�Զ���ӡ
    If mint���� <> 2 Then
        Call zlDatabase.SetPara("סԺ���͵��ݴ�ӡ", IIF(optPrint(0).value, 0, IIF(optPrint(1).value, 1, 2)), glngSys, pסԺҽ������, blnSetPara)
        Call zlDatabase.SetPara("ҽ������Χ", IIF(optBaby(0).value, 0, IIF(optBaby(1).value, 1, 2)), glngSys, pסԺҽ������, blnSetPara)
    End If

    gblnOK = True
    Unload Me
End Sub

Private Sub cmdУ�Բ���ALL_Click()
    Dim i As Integer
    
    For i = 0 To lstУ�Բ���.ListCount - 1
        lstУ�Բ���.Selected(i) = True
    Next
    lstУ�Բ���.SetFocus
End Sub

Private Sub cmdУ�Բ���Clear_Click()
    Dim i As Integer
    
    For i = 0 To lstУ�Բ���.ListCount - 1
        lstУ�Բ���.Selected(i) = False
    Next
    lstУ�Բ���.SetFocus
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeyReturn Then
        '����checkbox���س�����ת�ƽ���
        If Not Me.ActiveControl Is vsfDrugStore Then
            Call ZLCommFun.PressKey(vbKeyTab)
        End If
    ElseIf KeyCode = vbKeyA And Shift = vbCtrlMask Then
        If cmdУ�Բ���ALL.Enabled And cmdУ�Բ���ALL.Visible Then Call cmdУ�Բ���ALL_Click
    ElseIf KeyCode = vbKeyR And Shift = vbCtrlMask Then
        If cmdУ�Բ���Clear.Enabled And cmdУ�Բ���Clear.Visible Then Call cmdУ�Բ���Clear_Click
    End If
End Sub

Private Sub Form_Load()
    Dim rsTmp As ADODB.Recordset
    Dim strSQL As String, strPar As String, i As Long
    Dim objControl As Control
    Dim bln�´����� As Boolean, bln�������� As Boolean
    Dim objctl As Object, arrTmp() As String
    Dim strDSIDs As String, strDefault As String, lngBackColor As Long, bytLockEdit As Byte
    Dim intType1 As Integer, intType2 As Integer, lngRow As Long
    Dim ctl As Control
    Dim strExcute As String
    
    On Error GoTo errH
    
    gblnOK = False
            
    If mint���� <> 1 Then
        fra��������.Enabled = False
        fra�����ջ�.Enabled = False
        fraBat.Enabled = False
        fraУ�Բ���.Enabled = False
        cmdУ�Բ���ALL.Enabled = False
        cmdУ�Բ���Clear.Enabled = False
        
        For Each ctl In Me.Controls
            If ctl.Container Is fra�������� Then
                ctl.Enabled = False
            ElseIf ctl.Container Is fra�����ջ� Then
                ctl.Enabled = False
            ElseIf ctl.Container Is fraBat Then
                ctl.Enabled = False
            ElseIf ctl.Container Is fraУ�Բ��� Then
                ctl.Enabled = False
            End If
        Next
        
                
        If mint���� = 2 Then    'ҽ��վ
            fra��Ժ���.Visible = False
            chk��Ժ���.Visible = False
            chkStopNurseGrade.Visible = False
            chkͣ�����.Visible = False
            
            tabPar.TabVisible(1) = False
        End If
    End If
    
    bln�´����� = InStr(GetInsidePrivs(pסԺҽ���´�), "ҽ��ѡ������") > 0
    bln�������� = InStr(GetInsidePrivs(pסԺҽ������), "ҽ��ѡ������") > 0
    
    If mint���� <> 0 Then
        chk�ر�ҽ��.Enabled = False
    Else
        fra��������.Enabled = True
        chk�ر�ҽ��.Enabled = True
        chk�ر�ҽ��.value = Val(zlDatabase.GetPara("������ɺ�ر�ҽ������", glngSys, pסԺҽ���´�, "0", Array(chk�ر�ҽ��), bln��������))
        cmdBloodTip.Tag = zlDatabase.GetPara("��Ѫ����ע������", glngSys, pסԺҽ���´�, , Array(cmdBloodTip), bln�´�����)
    End If
    
    'Ҫ��������Ժ���
    strPar = zlDatabase.GetPara("Ҫ��������Ժ���", glngSys, pסԺҽ���´�, , Array(chk��Ժ���, lst��Ժ���), bln�´�����)
    If Not chk��Ժ���.Enabled Then lst��Ժ���.Tag = "1" '�̶���ʶΪ������
    If strPar <> "" Then
        chk��Ժ���.value = 1
        Call chk��Ժ���_Click
    End If
    strSQL = "Select ����,���� From ������Ŀ��� Where ���� Not IN('4','5','6','7','8','9') Union ALL Select '5','ҩƷ' From Dual Order by ����"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    With lst��Ժ���
        Do While Not rsTmp.EOF
            .AddItem rsTmp!���� & "-" & rsTmp!����
            .ItemData(.NewIndex) = Asc(rsTmp!����)
            
            If strPar <> "" Then
                If InStr(strPar, Chr(.ItemData(.NewIndex))) > 0 Then
                    .Selected(.NewIndex) = True
                End If
            End If
            rsTmp.MoveNext
        Loop
        .ListIndex = 0
    End With
    
    '������Һ�������ĵ�ҽ����Ч
    strPar = zlDatabase.GetPara("������Һ�������ĵ�ҽ����Ч", glngSys, pסԺҽ���´�, "2", Array(opt��Ч(0), opt��Ч(1), opt��Ч(2)), bln�´�����)
    opt��Ч(Val(strPar)).value = True
    
    '����ҩ��ȱʡ��ҩĿ��
    strPar = zlDatabase.GetPara("����ҩ��ȱʡ��ҩĿ��", glngSys, pסԺҽ���´�, "1")
    If strPar = "3" Then strPar = "0"
    optPurMed(Val(strPar)).value = True
    
    '����Һ�������ķ�ҩ�Ĳ��˿���
    If gstr��Һ�������� <> "" Then
        strPar = zlDatabase.GetPara("����Һ�������ķ�ҩ�Ĳ��˿���", glngSys, pסԺҽ���´�, "*", Array(lstDept, cmdAllSelect, cmdAllDel), bln�´�����)
        strSQL = "select distinct ID,����,����" & _
                    " from ���ű� D,��������˵�� T" & _
                    " where D.ID=T.����ID and t.��������='�ٴ�' And t.������� in(2,3)" & _
                    "       and (D.����ʱ�� is null or D.����ʱ��=to_date('3000-01-01','YYYY-MM-DD'))" & _
                    " order by ����"
        
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        i = -1
        Do While Not rsTmp.EOF
            lstDept.AddItem rsTmp!���� & "-" & rsTmp!����
            lstDept.ItemData(lstDept.NewIndex) = rsTmp!ID
            If strPar = "*" Or InStr("," & strPar & ",", "," & rsTmp!ID & ",") > 0 Then
                lstDept.Selected(lstDept.NewIndex) = True
                If i = -1 Then i = lstDept.NewIndex
            End If
            rsTmp.MoveNext
        Loop
        If i <> -1 Then lstDept.ListIndex = i
        If lstDept.ListIndex = -1 And lstDept.ListCount > 0 Then lstDept.ListIndex = 0
        '�������Ĳ����յľ���Ӫ��ҽ���ڲ�������
        chk����Ӫ��.value = Val(zlDatabase.GetPara("�������Ĳ����յľ���Ӫ��ҽ���ڲ�������", glngSys, pסԺҽ���´�, "0", Array(chk����Ӫ��), bln�´�����))
        chkWithSelf.value = Val(zlDatabase.GetPara("��������ҩƷ�����͵���������", glngSys, pסԺҽ���´�, "0", Array(chkWithSelf), bln�´�����))
    Else
        lstDept.Enabled = False
        fra��Һ����.Enabled = False
        cmdAllSelect.Enabled = False
        cmdAllDel.Enabled = False
        chk����Ӫ��.Enabled = False
        chkWithSelf.Enabled = False
        lstDept.AddItem "û��������Һ�������ģ�"
        lstDept.ListIndex = -1
        fra��Ч.Enabled = False
        For i = 0 To 2
            opt��Ч(i).Enabled = False
        Next
    End If
    
    dtpEnd.value = "23:59:59"
    
    '�Զ�У�ԵĲ���
    If mint���� = 1 Then
        strPar = zlDatabase.GetPara("�Զ����У�ԼƼ�", glngSys, pסԺҽ������, , Array(chkУ��, lstУ�Բ���, fraУ�Բ���, cmdУ�Բ���ALL, cmdУ�Բ���Clear), bln��������)
        If strPar <> "" Then chkУ��.value = 1
        Call chkУ��_Click
        
        strSQL = _
            " Select Distinct A.ID,A.����,A.����" & _
            " From ���ű� A,��������˵�� B " & _
            " Where A.ID=B.����ID And B.������� in(2,3) And B.��������='����'" & _
            " And (A.����ʱ�� is NULL or Trunc(A.����ʱ��)=To_Date('3000-01-01','YYYY-MM-DD'))" & _
            " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
            " Order by A.����"
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        i = -1
        Do While Not rsTmp.EOF
            lstУ�Բ���.AddItem rsTmp!���� & "-" & rsTmp!����
            lstУ�Բ���.ItemData(lstУ�Բ���.NewIndex) = rsTmp!ID
            If strPar = "*" Or InStr("," & strPar & ",", "," & rsTmp!ID & ",") > 0 Then
                lstУ�Բ���.Selected(lstУ�Բ���.NewIndex) = True
                If i = -1 Then i = lstУ�Բ���.NewIndex
            End If
            rsTmp.MoveNext
        Loop
        If i <> -1 Then lstУ�Բ���.ListIndex = i
        If lstУ�Բ���.ListIndex = -1 And lstУ�Բ���.ListCount > 0 Then lstУ�Բ���.ListIndex = 0
        
    
        
        '���ջصķ�ҩ����
        strPar = zlDatabase.GetPara("��ҩ���ջ�", glngSys, pסԺҽ������, , Array(lst��ҩ����), bln��������)
        strSQL = "Select ����, ���� From ��ҩ���� Order by ����"
        Set rsTmp = New ADODB.Recordset
        Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
        i = -1
        Do While Not rsTmp.EOF
            lst��ҩ����.AddItem rsTmp!����
            If strPar <> "" Then
                If InStr("," & strPar & ",", "," & rsTmp!���� & ",") > 0 Then
                    lst��ҩ����.Selected(lst��ҩ����.NewIndex) = True
                    If i = -1 Then i = lst��ҩ����.NewIndex
                End If
            End If
            rsTmp.MoveNext
        Loop
        If i <> -1 Then lst��ҩ����.ListIndex = i
        If lst��ҩ����.ListIndex = -1 And lst��ҩ����.ListCount > 0 Then lst��ҩ����.ListIndex = 0
        
        chkTurnCheck.value = Val(zlDatabase.GetPara("����ҽ������ǰ���δ��Чҽ��", glngSys, pסԺҽ������, 0, Array(chkTurnCheck), bln��������))
        chkLimit.value = Val(zlDatabase.GetPara("ҩ���������ƽ���ʱ��", glngSys, pסԺҽ������, 0, Array(chkLimit), bln��������))
        strPar = zlDatabase.GetPara("�����ڷ�ҩ���ͽ���ʱ��", glngSys, pסԺҽ������, , Array(chkInfo(chk�����ڷ�ҩ���ͽ���ʱ��)), bln��������)
        If InStr(strPar, "|") = 0 Then
            chkInfo(chk�����ڷ�ҩ���ͽ���ʱ��).value = 0
        Else
            chkInfo(chk�����ڷ�ҩ���ͽ���ʱ��).value = Val(Split(strPar, "|")(0))
            If chkInfo(chk�����ڷ�ҩ���ͽ���ʱ��).value = 1 Then
                dtpEnd.value = Format(Split(strPar, "|")(1), "HH:MM:SS")
                dtpEnd.Enabled = True
            End If
        End If
    Else
        chkInfo(chk�����ڷ�ҩ���ͽ���ʱ��).Enabled = False
    End If
    
    chkһ����.value = Val(zlDatabase.GetPara("����ȱʡһ����", glngSys, pסԺҽ���´�, , Array(chkһ����), bln�´�����))
    chk����.value = Val(zlDatabase.GetPara("ҽ��ִ������", glngSys, pסԺҽ���´�, , Array(chk����), bln�´�����))
    chkST.value = Val(zlDatabase.GetPara("�Զ�����Ƥ��", glngSys, pסԺҽ���´�, , Array(chkST), bln�´�����))
    optSTCheck(Val(zlDatabase.GetPara("����Ƥ�Խ������ҽ����������", glngSys, pסԺҽ���´�, , Array(lblSTCheck, optSTCheck(0), optSTCheck(1)), bln�´�����))).value = True
    Call chkST_Click
    chk���䵥��.value = Val(zlDatabase.GetPara("���������뵥��", glngSys, pסԺҽ���´�, , Array(chk���䵥��), bln�´�����))
    chk��Ժ���.value = Val(zlDatabase.GetPara("Ҫ�������Ժ���", glngSys, pסԺҽ���´�, , Array(chk��Ժ���), bln�´�����))
    chkStopNurseGrade.value = Val(zlDatabase.GetPara("����ֹͣ����ȼ�", glngSys, pסԺҽ���´�, 1, Array(chkStopNurseGrade), bln�´�����))
    chkAdviceSort.value = Val(zlDatabase.GetPara("ҽ���Զ�����", glngSys, pסԺҽ���´�, 0, Array(chkAdviceSort, cmdAdviceSortSet), bln�´�����))
    chkͣ�����.value = Val(zlDatabase.GetPara("ʵϰҽ��ֹͣҽ����Ҫ���", glngSys, pסԺҽ���´�, 0, Array(chkͣ�����), bln�´�����))
    Call chkAdviceSort_Click
    chk����ס����ҽ���´�.value = Val(zlDatabase.GetPara("���������ס�����´�ҽ��", glngSys, pסԺҽ���´�, 1, Array(chk����ס����ҽ���´�), bln�´�����))
    chk����ҽ��.value = Val(zlDatabase.GetPara("������ɺ��´�����ҽ��", glngSys, pסԺҽ���´�, , Array(chk����ҽ��), bln�´�����))
    
    strExcute = zlDatabase.GetPara("����ִ���Զ����", glngSys, pסԺҽ������, , Array(chkִ��, chkLong, chkShort), bln��������)
    chkLong.value = Val(Mid(strExcute, 1, 1))
    chkShort.value = Val(Mid(strExcute, 2, 1))
    
    chk��ӡ.value = Val(zlDatabase.GetPara("�Զ�����ҽ����ӡ", glngSys, pסԺҽ������, , Array(chk��ӡ), bln��������))
    chkҽ��.value = Val(zlDatabase.GetPara("ҽ��ҽ����������", glngSys, pסԺҽ������, , Array(chkҽ��), bln��������))
    chkƤ��.value = Val(zlDatabase.GetPara("Ƥ����֤���", glngSys, pסԺҽ������, , Array(chkƤ��), bln��������))
    
    i = Val(zlDatabase.GetPara("�����ջز�����������", glngSys, pסԺҽ������, , Array(optRoll(0), optRoll(1)), bln��������))
    If i = 1 Then
        optRoll(0).value = True
    Else
        optRoll(1).value = True
    End If
    chk�������.value = Val(zlDatabase.GetPara("�����ջط��ñ����Զ����", glngSys, pסԺҽ������, , Array(chk�������), bln��������))
    Call optRoll_Click(IIF(optRoll(0).value, 0, 1))
    
    chkУ��ǩ��.value = Val(zlDatabase.GetPara("У��ҽ������ǩ��", glngSys, pסԺҽ������, , Array(chkУ��ǩ��), bln��������))
    chkҽ������.value = Val(zlDatabase.GetPara("���ҽ������", glngSys, pסԺҽ������, 1, Array(chkҽ������), bln��������))
    
    chkAutoVerify.value = Val(zlDatabase.GetPara("����ǰ�Զ�У��", glngSys, pסԺҽ������, , Array(chkAutoVerify), bln��������))
    chkAutoRoll.value = Val(zlDatabase.GetPara("ֹͣ���Զ������ջ�", glngSys, pסԺҽ������, , Array(chkAutoRoll), bln��������))
    
    '��������
    chkBat(0).value = Val(zlDatabase.GetPara("����ҽ��У��", glngSys, pסԺҽ������, , Array(chkBat(0)), bln��������))
    chkBat(1).value = Val(zlDatabase.GetPara("����ҽ����ͣ", glngSys, pסԺҽ������, , Array(chkBat(1)), bln��������))
    
    'ҽ������ӡģʽ
    If Val(zlDatabase.GetPara("ҽ������ӡģʽ", glngSys, pסԺҽ���´�, , Array(optPrintType(0), optPrintType(1)), bln�´�����)) <> 0 Then
        optPrintType(1).value = True
    Else
        optPrintType(0).value = True
    End If
    '��Ѫ���뵥��ӡģʽ
    If Val(zlDatabase.GetPara("��Ѫ���뵥��ӡģʽ", glngSys, pסԺҽ������, , Array(optBloodPrintType(0), optBloodPrintType(1)), bln�´�����)) <> 1 Then
        optBloodPrintType(0).value = True
    Else
        optBloodPrintType(1).value = True
    End If
    
    'ҩ���뷢�ϲ���
    strSQL = _
        "Select Distinct A.ID,A.����,A.����,B.�������� " & _
        " From ���ű� A,��������˵�� B " & _
        " Where (A.����ʱ��=TO_DATE('3000-01-01','YYYY-MM-DD') Or A.����ʱ�� is NULL)" & _
        " AND B.����ID=A.ID And B.������� IN(2,3) and B.�������� in('��ҩ��','��ҩ��','��ҩ��','���ϲ���')" & _
        " And (A.վ��='" & gstrNodeNo & "' Or A.վ�� is Null)" & _
        " Order by ��������,����"
    Set rsTmp = New ADODB.Recordset
    Call zlDatabase.OpenRecordset(rsTmp, strSQL, Me.Caption)
    
    With vsfDrugStore
        .Rows = .FixedRows
        .Editable = flexEDKbdMouse
        .MergeCol(.ColIndex("���")) = True
        .MergeCells = flexMergeFixedOnly
        
        rsTmp.Filter = "��������<>'���ϲ���'"
        If Not rsTmp.EOF Then
            .Rows = .FixedRows + rsTmp.RecordCount
            lngRow = .FixedRows
            arrTmp = Split("��ҩ��,��ҩ��,��ҩ��", ",")
            For i = 0 To UBound(arrTmp)
                rsTmp.Filter = "��������='" & arrTmp(i) & "'"
                strDefault = zlDatabase.GetPara("סԺȱʡ" & arrTmp(i), glngSys, pסԺҽ���´�, , , , intType1)
                strDSIDs = "," & zlDatabase.GetPara("סԺ����" & arrTmp(i), glngSys, pסԺҽ���´�, , , , intType2) & ","
                Do While Not rsTmp.EOF
                    .TextMatrix(lngRow, .ColIndex("���")) = arrTmp(i)
                    .TextMatrix(lngRow, .ColIndex("ҩ��")) = rsTmp!����
                    .RowData(lngRow) = Val(rsTmp!ID)
                    
                    If Val(rsTmp!ID) = Val(strDefault) Then
                        .TextMatrix(lngRow, .ColIndex("ȱʡ")) = "��"
                        .TextMatrix(lngRow, .ColIndex("����")) = -1   'true
                    Else
                        .TextMatrix(lngRow, .ColIndex("ȱʡ")) = ""
                        .TextMatrix(lngRow, .ColIndex("����")) = IIF(InStr(strDSIDs, "," & rsTmp!ID & ",") > 0, -1, 0)
                    End If
                    
                    'ȱʡ��Ԫ��
                    'intType-'���ز������ͣ�1.����ȫ��,2.˽��ȫ��,3.����ģ��,4.˽��ģ��,5.��������ģ��(����Ȩ����),6.����˽��ģ��,15.��������ģ��(Ҫ��Ȩ����)
                    bytLockEdit = 0
                    If InStr(1, ",1,3,15,", "," & intType1 & ",") > 0 Then
                        lngBackColor = IIF(bln�´�����, VsPubBackColor, &H8000000F)      '��Ȩ�޿���
                        bytLockEdit = IIF(bln�´�����, 0, 1)
                    ElseIf intType1 = 5 Then
                        lngBackColor = VsPubBackColor       '����ģ��,������Ȩ�޿���
                    Else
                        lngBackColor = &H80000005     '�����༭
                    End If
                    .Cell(flexcpBackColor, lngRow, .ColIndex("ȱʡ")) = lngBackColor
                    .Cell(flexcpData, lngRow, .ColIndex("ȱʡ")) = bytLockEdit
                     
                    '���õ�Ԫ��
                    bytLockEdit = 0
                    If InStr(1, ",1,3,15,", "," & intType2 & ",") > 0 Then
                        lngBackColor = IIF(bln�´�����, VsPubBackColor, &H8000000F)      '��Ȩ�޿���
                        bytLockEdit = IIF(bln�´�����, 0, 1)
                    ElseIf intType2 = 5 Then
                        lngBackColor = VsPubBackColor       '����ģ��,������Ȩ�޿���
                    Else
                        lngBackColor = &H80000005     '�����༭
                    End If
                    .Cell(flexcpBackColor, lngRow, .ColIndex("����")) = lngBackColor
                    .Cell(flexcpData, lngRow, .ColIndex("����")) = bytLockEdit
                    
                    lngRow = lngRow + 1
                    rsTmp.MoveNext
                Loop
                If lngRow < .Rows - 1 Then  '���ָ���
                    .Select lngRow, .FixedCols, lngRow, .Cols - 1
                    .CellBorder vbBlue, 0, 1, 0, 0, 0, 0
                End If
            Next
        End If
    End With
    
    cbo����.AddItem "�˹�ѡ��"
    rsTmp.Filter = "��������='���ϲ���'"
    Do While Not rsTmp.EOF
        cbo����.AddItem rsTmp!����
        cbo����.ItemData(cbo����.ListCount - 1) = rsTmp!ID
        rsTmp.MoveNext
    Loop
    strPar = zlDatabase.GetPara("סԺȱʡ���ϲ���", glngSys, pסԺҽ���´�, , Array(lbl����, cbo����), bln�´�����)
    zlControl.CboLocate cbo����, strPar, True
        
    
    '����δ����ǩ��ʱ������������
    If gintCA = 0 Or Mid(gstrESign, 2, 1) <> "1" Then
        chkУ��ǩ��.value = 0
        chkУ��ǩ��.Enabled = False
    End If
    
    '���ݴ�ӡ:0-����ӡ,1-�ֹ���ӡ,2-�Զ���ӡ
    optPrint(Val(zlDatabase.GetPara("סԺ���͵��ݴ�ӡ", glngSys, pסԺҽ������, "2", Array(optPrint(0), optPrint(1), optPrint(2)), bln��������))).value = True
    
    'ҽ������Χ
    optBaby(Val(zlDatabase.GetPara("ҽ������Χ", glngSys, pסԺҽ������, "0", Array(optBaby(0), optBaby(1), optBaby(2)), bln��������))).value = True
    
    Exit Sub
errH:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Sub

Private Sub Form_Resize()
    
    cmdCancel.Left = Me.ScaleLeft + Me.ScaleWidth - cmdCancel.Width - 200
    cmdOK.Left = cmdCancel.Left - cmdOK.Width - 200
End Sub

Private Sub Form_Unload(Cancel As Integer)
    mMainPrivs = ""
End Sub

Private Sub lstDept_KeyPress(KeyAscii As Integer)
    Dim i As Long
    
    If KeyAscii = vbKeyBack Then Exit Sub
    
    For i = 0 To lstDept.ListCount - 1
        If ZLCommFun.SpellCode(Mid(lstDept.List(i), InStr(lstDept.List(i), "-") + 1)) Like UCase(Chr(KeyAscii)) & "*" Then
            lstDept.ListIndex = i: Exit For
        End If
    Next
End Sub

Private Sub vsfDrugStore_StartEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    If Col = vsfDrugStore.ColIndex("����") Then
        Call Set����ҩ��(Row, True)
    ElseIf Col = vsfDrugStore.ColIndex("����") Then
        Call Setȱʡҩ��
    End If
    Cancel = True
End Sub

Private Sub vsfDrugStore_BeforeEdit(ByVal Row As Long, ByVal Col As Long, Cancel As Boolean)
    With vsfDrugStore
        Select Case Col
        Case .ColIndex("����")
            Cancel = Val(.Cell(flexcpData, Row, Col)) <> 0
        Case .ColIndex("ȱʡ")
            Cancel = Val(.Cell(flexcpData, Row, Col)) <> 0
        Case Else
            Cancel = True
            Exit Sub
        End Select
    End With
End Sub

Private Sub vsfDrugStore_DblClick()
    With vsfDrugStore
        If .MouseCol = .ColIndex("ȱʡ") Then
            Call Setȱʡҩ��
        ElseIf .MouseCol = .ColIndex("ҩ��") Then
            Call Set����ҩ��(.Row, True)
        ElseIf .MouseCol = .ColIndex("����") And .MouseRow = .FixedRows - 1 Then
            Dim i As Long
            For i = .FixedRows To .Rows - 1
                Call Set����ҩ��(i)
            Next
        End If
    End With
End Sub
Private Sub vsfDrugStore_KeyDown(KeyCode As Integer, Shift As Integer)
    If KeyCode = vbKeySpace Then
        If vsfDrugStore.Col = vsfDrugStore.ColIndex("ȱʡ") Then
            Call Setȱʡҩ��
        End If
    End If
End Sub

Private Sub Setȱʡҩ��()
'���ܣ����õ�ǰ�е�ȱʡҩ����ͬʱ������ͬ���͵������е�ȱʡҩ��
    Dim i As Long
    
    With vsfDrugStore
        If Val("" & .Cell(flexcpData, .Row, .ColIndex("ȱʡ"))) = 0 Then  '�ò��������޸ĵ������
            If .TextMatrix(.Row, .ColIndex("ȱʡ")) = "��" Then
                .TextMatrix(.Row, .ColIndex("ȱʡ")) = ""
            Else
                '��û����Ȩ���޸Ŀ���ʱ�ҿ���Ϊ0��false)ʱ����������ȱʡ
                If Not (Val(.TextMatrix(.Row, .ColIndex("����"))) = 0 And Val("" & .Cell(flexcpData, .Row, .ColIndex("����"))) = 1) Then
                    'ͬ����������ȡ��ȱʡ
                    For i = .FixedRows To .Rows - 1
                        If .TextMatrix(.Row, .ColIndex("���")) = .TextMatrix(i, .ColIndex("���")) Then
                            If .TextMatrix(i, .ColIndex("ȱʡ")) = "��" Then .TextMatrix(i, .ColIndex("ȱʡ")) = ""
                        End If
                    Next
                    .TextMatrix(.Row, .ColIndex("����")) = -1    '�Զ�����Ϊ����
                    .TextMatrix(.Row, .ColIndex("ȱʡ")) = "��"
                Else
                    MsgBox "���õ�ǰҩ��Ϊȱʡʱ����ͬʱ����ǰҩ������Ϊ���ã�" & vbNewLine & "��û���޸Ŀ���ҩ����Ȩ�ޡ�", vbInformation, gstrSysName
                End If
            End If
        Else
            MsgBox "��û���޸�ȱʡҩ����Ȩ�ޡ�", vbInformation, gstrSysName
        End If
    End With
End Sub

Private Sub Set����ҩ��(ByVal lngRow As Long, Optional ByVal blnAsk As Boolean = False)
'���ܣ����õ�ǰ�еĿ���ҩ����ͬʱ����ǰ�е�ȱʡҩ��

    With vsfDrugStore
        If Val("" & .Cell(flexcpData, lngRow, .ColIndex("����"))) = 0 Then   '�ò��������޸ĵ������
            If Val(.TextMatrix(lngRow, .ColIndex("����"))) = -1 Then
                '��ǰ���ҹ�ѡ����
                If Not (Val("" & .Cell(flexcpData, lngRow, .ColIndex("ȱʡ"))) = 1 And .TextMatrix(lngRow, .ColIndex("ȱʡ")) = "��") Then
                    .TextMatrix(lngRow, .ColIndex("����")) = 0
                    .TextMatrix(lngRow, .ColIndex("ȱʡ")) = ""
                Else
                    If blnAsk Then
                        MsgBox "ȡ����ǰҩ������ʱ����ͬʱȡ����ǰҩ��ȱʡ��" & vbNewLine & "��û���޸�ȱʡҩ����Ȩ�ޡ�", vbInformation, gstrSysName
                    End If
                End If
            Else
                .TextMatrix(lngRow, .ColIndex("����")) = -1    '�Զ�����Ϊ����
            End If
        Else
            If blnAsk Then
                MsgBox "��û���޸Ŀ���ҩ����Ȩ�ޡ�", vbInformation, gstrSysName
            End If
        End If
    End With
End Sub




