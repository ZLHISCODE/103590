VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmLisStationPara 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��������"
   ClientHeight    =   6975
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   5970
   Icon            =   "frmLisStationPara.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6975
   ScaleWidth      =   5970
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  '����������
   Begin MSComDlg.CommonDialog dlgFile 
      Left            =   5160
      Top             =   1020
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "����(&H)"
      Height          =   375
      Left            =   150
      TabIndex        =   6
      TabStop         =   0   'False
      Top             =   6495
      Width           =   1100
   End
   Begin TabDlg.SSTab tbs 
      Height          =   6315
      Left            =   90
      TabIndex        =   2
      Top             =   60
      Width           =   5715
      _ExtentX        =   10081
      _ExtentY        =   11139
      _Version        =   393216
      Style           =   1
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      TabCaption(0)   =   "&1.����"
      TabPicture(0)   =   "frmLisStationPara.frx":000C
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Frame6"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "Frame5"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Frame7"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).ControlCount=   3
      TabCaption(1)   =   "&2.�ļ���ȡ"
      TabPicture(1)   =   "frmLisStationPara.frx":0028
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtFile"
      Tab(1).Control(1)=   "cmdFile"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "cboDevice"
      Tab(1).Control(3)=   "optRange(0)"
      Tab(1).Control(4)=   "optRange(1)"
      Tab(1).Control(5)=   "cardSet"
      Tab(1).Control(6)=   "dtpStart"
      Tab(1).Control(7)=   "dtpEnd"
      Tab(1).Control(8)=   "lblNotify"
      Tab(1).Control(9)=   "Label10"
      Tab(1).Control(10)=   "Label12"
      Tab(1).Control(11)=   "Label13"
      Tab(1).ControlCount=   12
      TabCaption(2)   =   "&3.���Ҵ�ӡ"
      TabPicture(2)   =   "frmLisStationPara.frx":0044
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "cmdSelectAll"
      Tab(2).Control(1)=   "cmdClearAll"
      Tab(2).Control(2)=   "lvwDept"
      Tab(2).Control(3)=   "Label1"
      Tab(2).ControlCount=   4
      TabCaption(3)   =   "&4.����"
      TabPicture(3)   =   "frmLisStationPara.frx":0060
      Tab(3).ControlEnabled=   0   'False
      Tab(3).Control(0)=   "frmsign"
      Tab(3).Control(0).Enabled=   0   'False
      Tab(3).Control(1)=   "Frame11"
      Tab(3).Control(1).Enabled=   0   'False
      Tab(3).Control(2)=   "Frame10"
      Tab(3).Control(2).Enabled=   0   'False
      Tab(3).Control(3)=   "Frame9"
      Tab(3).Control(3).Enabled=   0   'False
      Tab(3).Control(4)=   "Frame8"
      Tab(3).Control(4).Enabled=   0   'False
      Tab(3).ControlCount=   5
      Begin VB.Frame frmsign 
         Caption         =   "ǩ������"
         Height          =   855
         Left            =   -74820
         TabIndex        =   76
         Top             =   2640
         Width           =   5280
         Begin VB.CheckBox checkSaveReprotSign 
            Caption         =   "���浥����ʱǩ��"
            Height          =   315
            Left            =   2970
            TabIndex        =   78
            Top             =   300
            Width           =   2025
         End
         Begin VB.CheckBox checkSaveInfoSign 
            Caption         =   "���յǼǱ���ʱǩ��"
            Height          =   315
            Left            =   210
            TabIndex        =   77
            Top             =   300
            Width           =   2025
         End
      End
      Begin VB.Frame Frame11 
         Height          =   525
         Left            =   -74820
         TabIndex        =   66
         Top             =   510
         Width           =   5280
         Begin VB.OptionButton opt���ﴦ�� 
            Caption         =   "��ʾ����"
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   0
            Left            =   1935
            TabIndex        =   69
            Top             =   180
            Value           =   -1  'True
            Width           =   1140
         End
         Begin VB.OptionButton opt���ﴦ�� 
            Caption         =   "�Զ�����"
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   1
            Left            =   3060
            TabIndex        =   68
            Top             =   180
            Width           =   1065
         End
         Begin VB.OptionButton opt���ﴦ�� 
            Caption         =   "������"
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   2
            Left            =   4275
            TabIndex        =   67
            Top             =   180
            Width           =   975
         End
         Begin VB.Label Label14 
            Caption         =   "���ﲡ����Ϣ��һ��ʱ"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   75
            TabIndex        =   70
            Top             =   210
            Width           =   2160
         End
      End
      Begin VB.Frame Frame10 
         Height          =   525
         Left            =   -74820
         TabIndex        =   61
         Top             =   1035
         Width           =   5280
         Begin VB.OptionButton optסԺ���� 
            Caption         =   "������"
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   2
            Left            =   4275
            TabIndex        =   64
            Top             =   180
            Width           =   975
         End
         Begin VB.OptionButton optסԺ���� 
            Caption         =   "�Զ�����"
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   1
            Left            =   3060
            TabIndex        =   63
            Top             =   180
            Width           =   1065
         End
         Begin VB.OptionButton optסԺ���� 
            Caption         =   "��ʾ����"
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   0
            Left            =   1935
            TabIndex        =   62
            Top             =   180
            Value           =   -1  'True
            Width           =   1140
         End
         Begin VB.Label Label11 
            Caption         =   "סԺ������Ϣ��һ��ʱ"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   75
            TabIndex        =   65
            Top             =   210
            Width           =   2160
         End
      End
      Begin VB.Frame Frame9 
         Height          =   525
         Left            =   -74820
         TabIndex        =   56
         Top             =   2085
         Width           =   5280
         Begin VB.OptionButton opt��촦�� 
            Caption         =   "��ʾ����"
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   0
            Left            =   1920
            TabIndex        =   59
            Top             =   165
            Value           =   -1  'True
            Width           =   1140
         End
         Begin VB.OptionButton opt��촦�� 
            Caption         =   "�Զ�����"
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   1
            Left            =   3045
            TabIndex        =   58
            Top             =   165
            Width           =   1065
         End
         Begin VB.OptionButton opt��촦�� 
            Caption         =   "������"
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   2
            Left            =   4260
            TabIndex        =   57
            Top             =   165
            Width           =   960
         End
         Begin VB.Label lbl������Ϣ 
            Caption         =   "��첡����Ϣ��һ��ʱ"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   60
            TabIndex        =   60
            Top             =   195
            Width           =   2160
         End
      End
      Begin VB.Frame Frame8 
         Height          =   525
         Left            =   -74820
         TabIndex        =   51
         Top             =   1575
         Width           =   5280
         Begin VB.OptionButton optԺ�⴦�� 
            Caption         =   "��ʾ����"
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   0
            Left            =   1935
            TabIndex        =   54
            Top             =   180
            Value           =   -1  'True
            Width           =   1140
         End
         Begin VB.OptionButton optԺ�⴦�� 
            Caption         =   "�Զ�����"
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   1
            Left            =   3060
            TabIndex        =   53
            Top             =   180
            Width           =   1065
         End
         Begin VB.OptionButton optԺ�⴦�� 
            Caption         =   "������"
            ForeColor       =   &H00C00000&
            Height          =   285
            Index           =   2
            Left            =   4275
            TabIndex        =   52
            Top             =   180
            Width           =   975
         End
         Begin VB.Label Label8 
            Caption         =   "Ժ�ⲡ����Ϣ��һ��ʱ"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   75
            TabIndex        =   55
            Top             =   210
            Width           =   2160
         End
      End
      Begin VB.CommandButton cmdSelectAll 
         Caption         =   "ȫѡ(&A)"
         Height          =   375
         Left            =   -70680
         TabIndex        =   48
         Top             =   840
         Width           =   1100
      End
      Begin VB.CommandButton cmdClearAll 
         Caption         =   "ȫ��(&L)"
         Height          =   375
         Left            =   -70680
         TabIndex        =   47
         Top             =   1440
         Width           =   1100
      End
      Begin VB.TextBox txtFile 
         Height          =   300
         Left            =   -74700
         TabIndex        =   41
         Top             =   1560
         Width           =   4725
      End
      Begin VB.CommandButton cmdFile 
         Caption         =   "&S"
         Height          =   300
         Left            =   -69990
         TabIndex        =   40
         TabStop         =   0   'False
         Top             =   1560
         Width           =   300
      End
      Begin VB.ComboBox cboDevice 
         Height          =   300
         Left            =   -73530
         Style           =   2  'Dropdown List
         TabIndex        =   39
         Top             =   1980
         Width           =   3855
      End
      Begin VB.OptionButton optRange 
         Caption         =   "ֻ��ȡ��������(&T)"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   0
         Left            =   -74700
         TabIndex        =   38
         Top             =   2460
         Value           =   -1  'True
         Width           =   1965
      End
      Begin VB.OptionButton optRange 
         Caption         =   "��ȡָ��ʱ������(&R)"
         ForeColor       =   &H00C00000&
         Height          =   180
         Index           =   1
         Left            =   -72600
         TabIndex        =   37
         Top             =   2460
         Width           =   2025
      End
      Begin VB.Frame cardSet 
         Caption         =   "�����ӿ�����"
         Height          =   810
         Left            =   -74670
         TabIndex        =   33
         Top             =   3240
         Width           =   4875
         Begin VB.CommandButton cmdIC 
            Caption         =   "IC������(I)"
            Height          =   375
            Left            =   3360
            TabIndex        =   35
            Top             =   285
            Width           =   1215
         End
         Begin VB.CommandButton cmdIdent 
            Caption         =   "�豸����(&S)"
            Height          =   390
            Left            =   330
            TabIndex        =   34
            Top             =   285
            Width           =   1260
         End
      End
      Begin VB.Frame Frame7 
         Caption         =   "�������ز���"
         Height          =   3810
         Left            =   120
         TabIndex        =   9
         Top             =   2370
         Width           =   5445
         Begin VB.CheckBox chkSampleType 
            Caption         =   "�ϴν�������ձ걾����"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   120
            TabIndex        =   79
            ToolTipText     =   "�ϴν�������ձ걾����"
            Top             =   2685
            Width           =   2640
         End
         Begin VB.CheckBox chkAutoAddItem 
            Caption         =   "�Զ����Ӽ�����Ŀ"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   3030
            TabIndex        =   75
            ToolTipText     =   "�Զ�����δ����ļ�����Ŀ"
            Top             =   2430
            Width           =   1965
         End
         Begin VB.CheckBox chkLoadLast 
            Caption         =   "�Ǽ�ʱ������һ��������Ŀ"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   120
            TabIndex        =   74
            Top             =   2430
            Width           =   2745
         End
         Begin VB.CheckBox chkLast 
            Caption         =   "����ʱ��ʾ�ϴγ�����"
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3030
            TabIndex        =   73
            Top             =   3450
            Width           =   2355
         End
         Begin VB.CheckBox chkOnlyMachine 
            Caption         =   "ֻ���յ�ǰ������Ŀ"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   3030
            TabIndex        =   72
            Top             =   2160
            Width           =   2265
         End
         Begin VB.CheckBox chkSkipRule 
            Caption         =   "��˺��Զ�������һ������걾"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   120
            TabIndex        =   71
            Top             =   2160
            Width           =   2835
         End
         Begin VB.CheckBox chkNotSend 
            Caption         =   "ʹ�ö����������"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   3030
            TabIndex        =   32
            Top             =   1890
            Width           =   2265
         End
         Begin VB.CheckBox chkItemNumber 
            Caption         =   "�ֹ���Ŀ����Ŀ�ۼӱ걾��"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   120
            TabIndex        =   31
            Top             =   1890
            Width           =   2685
         End
         Begin VB.CheckBox ChkCheckInNoItem 
            Caption         =   "�Ǽ�ʱ����Ҫ������Ŀ"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   3030
            TabIndex        =   28
            Top             =   1620
            Width           =   2265
         End
         Begin VB.CheckBox chkShowOption 
            Caption         =   "ֻ�ں��յǼ�ʱ��ʾ�ǼǴ���"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   120
            TabIndex        =   27
            Top             =   1620
            Width           =   2715
         End
         Begin VB.CheckBox chkNO 
            Caption         =   "���ϴ�����ı걾���ۼ�"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   3030
            TabIndex        =   26
            Top             =   1350
            Width           =   2355
         End
         Begin VB.CheckBox chkShowType 
            Caption         =   "����Ӧ�߶��Զ�������ʾ���"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   120
            TabIndex        =   25
            Top             =   1350
            Width           =   2685
         End
         Begin VB.CheckBox chkPatientType 
            Caption         =   "���еǼǲ��˱�ʶΪ����"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   3030
            TabIndex        =   23
            Top             =   1089
            Width           =   2355
         End
         Begin VB.CheckBox chkShowAll 
            Caption         =   "������������ʾ������Ŀ"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   120
            TabIndex        =   22
            Top             =   1089
            Width           =   2355
         End
         Begin VB.CheckBox ChkPrivacy 
            Caption         =   "���浥�Ƿ���ʾ��˽��Ŀ"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   3030
            TabIndex        =   21
            Top             =   816
            Width           =   2355
         End
         Begin VB.CheckBox chkNoRange 
            Caption         =   "����ʱ����ָ����ʱ�䷶Χ(&I)"
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   120
            TabIndex        =   17
            Top             =   3180
            Width           =   2745
         End
         Begin VB.CheckBox chkCheck 
            Caption         =   "����ʱ��ʾ�Ƿ��շ�(&N)"
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   3030
            TabIndex        =   16
            Top             =   3180
            Width           =   2355
         End
         Begin VB.CheckBox chkAutoRefresh 
            Caption         =   "�յ����������Զ�ˢ��(&A)"
            ForeColor       =   &H00C00000&
            Height          =   225
            Left            =   120
            TabIndex        =   15
            Top             =   270
            Width           =   2745
         End
         Begin VB.CheckBox chkComm 
            Caption         =   "����ʱ����˫��ͨ��(&D)"
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   120
            TabIndex        =   14
            Top             =   3450
            Width           =   2355
         End
         Begin VB.CheckBox chkSample 
            Caption         =   "�Ǽ�ʱ��ֱ�����벡����Ϣ"
            ForeColor       =   &H00C00000&
            Height          =   225
            Left            =   120
            TabIndex        =   13
            Top             =   816
            Width           =   2835
         End
         Begin VB.CheckBox chkEmerge 
            Caption         =   "�걾���ּ���/����(&E)"
            ForeColor       =   &H00C00000&
            Height          =   225
            Left            =   120
            TabIndex        =   12
            Top             =   543
            Width           =   2475
         End
         Begin VB.CheckBox chkPrint 
            Caption         =   "��˺��Զ���ӡ(&P)"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   3030
            TabIndex        =   11
            Top             =   270
            Width           =   1935
         End
         Begin VB.CheckBox chkCheckAll 
            Caption         =   "��������Ŀ����(&C)"
            ForeColor       =   &H00C00000&
            Height          =   180
            Left            =   3030
            TabIndex        =   10
            Top             =   543
            Width           =   2355
         End
         Begin VB.Label lblNotice 
            Caption         =   "ѡ������ѡ����ܻ�ʹ���չ��̱�����"
            ForeColor       =   &H000000FF&
            Height          =   255
            Left            =   90
            TabIndex        =   18
            Top             =   2940
            Width           =   4395
         End
      End
      Begin VB.Frame Frame5 
         Caption         =   "ʱ�䷶Χ"
         Height          =   1110
         Left            =   120
         TabIndex        =   3
         Top             =   375
         Width           =   5475
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   3
            Left            =   2085
            Style           =   2  'Dropdown List
            TabIndex        =   29
            ToolTipText     =   "�����ҵ���ǰ���˵����μ����ʱ�䷶Χ"
            Top             =   630
            Width           =   1920
         End
         Begin MSComCtl2.DTPicker DTPHisTory 
            Height          =   285
            Left            =   4080
            TabIndex        =   24
            Top             =   270
            Width           =   1335
            _ExtentX        =   2355
            _ExtentY        =   503
            _Version        =   393216
            Format          =   68812801
            CurrentDate     =   39475
         End
         Begin VB.ComboBox cbo 
            ForeColor       =   &H00000000&
            Height          =   300
            Index           =   2
            Left            =   2085
            Style           =   2  'Dropdown List
            TabIndex        =   19
            ToolTipText     =   "�����ҵ���ǰ���˵����μ����ʱ�䷶Χ"
            Top             =   270
            Width           =   1920
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "�걾������ɹ���(&2)"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   3
            Left            =   345
            TabIndex        =   30
            Top             =   690
            Width           =   1710
         End
         Begin VB.Label lbl 
            AutoSize        =   -1  'True
            Caption         =   "���μ��鷶Χ(&1)"
            ForeColor       =   &H00C00000&
            Height          =   180
            Index           =   2
            Left            =   705
            TabIndex        =   20
            Top             =   330
            Width           =   1350
         End
         Begin VB.Label Label9 
            Caption         =   "�ڼ��鼼ʦ����վ�еĴ����ա��ڼ����Լ�����ɵ�ʱ�䷶Χ�ֱ��������ý���������"
            Height          =   15
            Left            =   840
            TabIndex        =   4
            Top             =   240
            Visible         =   0   'False
            Width           =   4065
         End
         Begin VB.Image Image1 
            Height          =   480
            Left            =   135
            Picture         =   "frmLisStationPara.frx":007C
            Top             =   105
            Width           =   480
         End
      End
      Begin VB.Frame Frame6 
         Caption         =   "��ʷ�Ƚϲ���ʶ��ʽ"
         Height          =   765
         Left            =   105
         TabIndex        =   5
         Top             =   1530
         Width           =   5475
         Begin VB.OptionButton OptHistoryName 
            Caption         =   "��������"
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   2370
            TabIndex        =   8
            Top             =   330
            Width           =   1845
         End
         Begin VB.OptionButton optHistoryID 
            Caption         =   "����ID"
            ForeColor       =   &H00C00000&
            Height          =   285
            Left            =   840
            TabIndex        =   7
            Top             =   330
            Width           =   1455
         End
         Begin VB.Image Image2 
            Height          =   480
            Left            =   105
            Picture         =   "frmLisStationPara.frx":0946
            Top             =   210
            Width           =   480
         End
      End
      Begin MSComCtl2.DTPicker dtpStart 
         Height          =   285
         Left            =   -74670
         TabIndex        =   36
         Top             =   2790
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   68812803
         CurrentDate     =   38792
      End
      Begin MSComCtl2.DTPicker dtpEnd 
         Height          =   285
         Left            =   -72900
         TabIndex        =   42
         Top             =   2790
         Width           =   1425
         _ExtentX        =   2514
         _ExtentY        =   503
         _Version        =   393216
         Enabled         =   0   'False
         CustomFormat    =   "yyyy-MM-dd"
         Format          =   68812803
         CurrentDate     =   38792
      End
      Begin MSComctlLib.ListView lvwDept 
         Height          =   5625
         Left            =   -74850
         TabIndex        =   49
         Top             =   660
         Width           =   3855
         _ExtentX        =   6800
         _ExtentY        =   9922
         View            =   3
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         Checkboxes      =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "����"
            Object.Width           =   2277
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "����"
            Object.Width           =   4235
         EndProperty
      End
      Begin VB.Label Label1 
         Caption         =   "ֻ��ӡָ��������ҵı��浥"
         Height          =   195
         Left            =   -74820
         TabIndex        =   50
         Top             =   420
         Width           =   3525
      End
      Begin VB.Label lblNotify 
         AutoSize        =   -1  'True
         Caption         =   "    ĳЩ��������ֱ�ӴӴ��ڶ�ȡ���ݣ��������������ض��ļ�����ȡ���ݡ�"
         ForeColor       =   &H8000000D&
         Height          =   360
         Left            =   -74730
         TabIndex        =   46
         Top             =   660
         Width           =   5205
         WordWrap        =   -1  'True
      End
      Begin VB.Label Label10 
         Caption         =   "���������ļ�(&F)"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   -74730
         TabIndex        =   45
         Top             =   1290
         Width           =   1395
      End
      Begin VB.Label Label12 
         Caption         =   "��������(&Y)"
         ForeColor       =   &H00C00000&
         Height          =   225
         Left            =   -74700
         TabIndex        =   44
         Top             =   2040
         Width           =   1305
      End
      Begin VB.Label Label13 
         Caption         =   "��"
         Height          =   165
         Left            =   -73140
         TabIndex        =   43
         Top             =   2850
         Width           =   285
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��(&O)"
      Height          =   375
      Left            =   3480
      TabIndex        =   0
      Top             =   6495
      Width           =   1100
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "ȡ��(&C)"
      Height          =   375
      Left            =   4650
      TabIndex        =   1
      Top             =   6495
      Width           =   1100
   End
End
Attribute VB_Name = "frmLisStationPara"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private mblnOK As Boolean
Private mlngLoop As Long
Private mfrmMain As Object
Private mstrPrivs As String                                         'Ȩ��

Public Function ShowPara(ByVal frmMain As Object) As Boolean
    Dim rsTmp As New ADODB.Recordset
    Dim objCbo As ComboBox, lngҩ��ID As Long
    Dim strsql As String, strPar As String, i As Long
    Dim bln�������� As Boolean
    Dim strMachine As String
    Dim strDepts As String
    Dim lItem As ListItem
    Dim int������Ϣ���� As Integer
    
    mblnOK = False
    mstrPrivs = gstrPrivs
    
'    If InStr(mstrPrivs, "��������") <= 0 Then
'        Me.chkSample.Enabled = False
'        Me.chkPatientType.Enabled = False
'        Me.chkNotSend.Enabled = False
'    End If
    bln�������� = InStr(";" & mstrPrivs & ";", ";��������;")
    Set mfrmMain = frmMain
    '��ʼ��
    
    For mlngLoop = 2 To 2
        cbo(mlngLoop).AddItem "��  ��"
        cbo(mlngLoop).AddItem "��  ��"
        cbo(mlngLoop).AddItem "��  ��"
        cbo(mlngLoop).AddItem "��  ��"
        cbo(mlngLoop).AddItem "��  ��"
        cbo(mlngLoop).AddItem "������"
        cbo(mlngLoop).AddItem "��  ��"
        cbo(mlngLoop).AddItem "ǰ����"
        cbo(mlngLoop).AddItem "ǰһ��"
        cbo(mlngLoop).AddItem "ǰ����"
        cbo(mlngLoop).AddItem "ǰһ��"
        cbo(mlngLoop).AddItem "ǰ����"
        cbo(mlngLoop).AddItem "ǰ����"
        cbo(mlngLoop).AddItem "ǰ����"
        cbo(mlngLoop).AddItem "�Զ���"
    Next
    
    cbo(2).AddItem "ָ����ʼ����"
    
    cbo(3).AddItem "��  ��"
    cbo(3).AddItem "��  ��"
    cbo(3).AddItem "��  ��"
    cbo(3).AddItem "��  ��"
    cbo(3).AddItem "���ظ�"
    
    On Error Resume Next
    chkSample.Value = zlDatabase.GetPara("�Ǽ�ʱ��ֱ�����벡����Ϣ", 100, 1208, 0, Array(chkSample), bln��������)
    chkPrint.Value = zlDatabase.GetPara("��˴�ӡ", 100, 1208, 0, Array(chkPrint), bln��������)
    chkShowAll.Value = zlDatabase.GetPara("������������ʾ������Ŀ", 100, 1208, 0, Array(chkShowAll), bln��������)
    ChkPrivacy.Value = zlDatabase.GetPara("���浥�Ƿ���ʾ��˽��Ŀ", 100, 1208, 0, Array(ChkPrivacy), bln��������)
    chkPatientType.Value = zlDatabase.GetPara("���еǼǲ��˱�ʶΪ����", 100, 1208, 0, Array(chkPatientType), bln��������)
    chkNotSend.Value = zlDatabase.GetPara("ʹ�ö����������", 100, 1208, 0, Array(chkNotSend), bln��������)
    
    checkSaveInfoSign.Value = zlDatabase.GetPara("���յǼǱ���ʱǩ��", 100, 1208, 0, Array(checkSaveInfoSign), bln��������)
    checkSaveReprotSign.Value = zlDatabase.GetPara("���浥����ʱǩ��", 100, 1208, 0, Array(checkSaveReprotSign), bln��������)
    
    
    
    
    int������Ϣ���� = Val(zlDatabase.GetPara("���ﲡ����Ϣ��һ�µĴ���ʽ", 100, 1208, 1, Array(opt���ﴦ��(0), opt���ﴦ��(1), opt���ﴦ��(2)), bln��������))
    opt���ﴦ��(0).Value = int������Ϣ���� = 1
    opt���ﴦ��(1).Value = int������Ϣ���� = 2
    opt���ﴦ��(2).Value = int������Ϣ���� = 3
    
    int������Ϣ���� = Val(zlDatabase.GetPara("סԺ������Ϣ��һ�µĴ���ʽ", 100, 1208, 1, Array(optסԺ����(0), optסԺ����(1), optסԺ����(2)), bln��������))
    optסԺ����(0).Value = int������Ϣ���� = 1
    optסԺ����(1).Value = int������Ϣ���� = 2
    optסԺ����(2).Value = int������Ϣ���� = 3
    
    int������Ϣ���� = Val(zlDatabase.GetPara("Ժ�ⲡ����Ϣ��һ�µĴ���ʽ", 100, 1208, 1, Array(optԺ�⴦��(0), optԺ�⴦��(1), optԺ�⴦��(2)), bln��������))
    optԺ�⴦��(0).Value = int������Ϣ���� = 1
    optԺ�⴦��(1).Value = int������Ϣ���� = 2
    optԺ�⴦��(2).Value = int������Ϣ���� = 3
    
    int������Ϣ���� = Val(zlDatabase.GetPara("��첡����Ϣ��һ�µĴ���ʽ", 100, 1208, 1, Array(opt��촦��(0), opt��촦��(1), opt��촦��(2)), bln��������))
    opt��촦��(0).Value = int������Ϣ���� = 1
    opt��촦��(1).Value = int������Ϣ���� = 2
    opt��촦��(2).Value = int������Ϣ���� = 3
    
    
    cbo(2).Text = zlDatabase.GetPara("���μ��鷶Χ", 100, 1208, "��  ��", Array(cbo(2)), bln��������)
    cbo(3).Text = zlDatabase.GetPara("�걾������ɹ���", 100, 1208, "��  ��", Array(cbo(3)), bln��������)
    Me.DTPHisTory.Value = zlDatabase.GetPara("���μ��鷶Χָ����ʼ����", 100, 1208, Format(Now - 30, "yyyy-mm-dd"), Array(Me.DTPHisTory), bln��������)
    Me.DTPHisTory.Visible = (cbo(2).Text = "ָ����ʼ����")
    chkAutoRefresh.Value = zlDatabase.GetPara("�Զ�ˢ��", 100, 1208, 1, Array(chkAutoRefresh), bln��������)
    chkNoRange.Value = zlDatabase.GetPara("���պ���ʱ��", 100, 1208, 1, Array(chkNoRange), bln��������)
    chkCheck.Value = zlDatabase.GetPara("������ʾ�շ�", 100, 1208, 1, Array(chkCheck), bln��������)
    chkComm.Value = zlDatabase.GetPara("��������˫��", 100, 1208, 0, Array(chkComm), bln��������)
    chkEmerge.Value = zlDatabase.GetPara("����걾", 100, 1208, 0, Array(chkEmerge), bln��������)
    chkCheckAll.Value = zlDatabase.GetPara("��������Ŀ����", 100, 1208, 0, Array(chkCheckAll), bln��������)
    chkShowType.Value = zlDatabase.GetPara("����Ӧ��ʾ���", 100, 1208, 0, Array(chkShowType), bln��������)
    chkNO.Value = zlDatabase.GetPara("���ϴ�����ı걾���ۼ�", 100, 1208, 0, Array(chkNO), bln��������)
    chkShowOption.Value = zlDatabase.GetPara("ֻ�ں��յǼ�ʱ��ʾ�ǼǴ���", 100, 1208, 0, Array(chkShowOption), bln��������)
    ChkCheckInNoItem.Value = zlDatabase.GetPara("�Ǽ�ʱ����Ҫ������Ŀ", 100, 1208, 0, Array(ChkCheckInNoItem), bln��������)
    chkItemNumber.Value = zlDatabase.GetPara("�ֹ���Ŀ����Ŀ�ۼӱ걾��", 100, 1208, 0, Array(chkItemNumber), bln��������)
    chkSkipRule.Value = zlDatabase.GetPara("��˺�������һ������걾", 100, 1208, 0, Array(chkSkipRule), bln��������)
    chkOnlyMachine.Value = zlDatabase.GetPara("ֻ���յ�ǰ������Ŀ", 100, 1208, 0, Array(chkOnlyMachine), bln��������)
    chkLast.Value = zlDatabase.GetPara("����ʱ��ʾ�ϴγ�����", 100, 1208, 0, Array(chkLast), bln��������)
    chkLoadLast.Value = zlDatabase.GetPara("�Ǽ�ʱ������һ��������Ŀ", 100, 1208, 0, Array(chkLoadLast), bln��������)
    chkAutoAddItem.Value = zlDatabase.GetPara("�Զ����Ӽ�����Ŀ", 100, 1208, 1, Array(chkAutoAddItem), bln��������)
    chkSampleType.Value = zlDatabase.GetPara("�ϴν�������ձ걾����", 100, 1208, 0, Array(chkSampleType), bln��������)
    
    i = zlDatabase.GetPara("��ʷ����ʶ��", 100, 1208, 0, Array(optHistoryID), bln��������)
    i = zlDatabase.GetPara("��ʷ����ʶ��", 100, 1208, 0, Array(OptHistoryName), bln��������)
    If i = 0 Then
        Me.optHistoryID.Value = True
    Else
        Me.OptHistoryName.Value = True
    End If
    
    If cbo(0).ListIndex = -1 Then cbo(0).ListIndex = 0
    If cbo(1).ListIndex = -1 Then cbo(1).ListIndex = 0
    If cbo(2).ListIndex = -1 Then cbo(2).ListIndex = 0
    
    '��ʼ�ļ���ȡ����
    On Error GoTo DBError
    
    strsql = "Select " & gConst_��������_���� & " From �������� a"
    Set rsTmp = zlDatabase.OpenSQLRecord(strsql, Me.Caption)
    Me.cboDevice.Clear
    Do While Not rsTmp.EOF
        cboDevice.AddItem "(" & rsTmp("����") & ")" & rsTmp("����")
        cboDevice.ItemData(cboDevice.ListCount - 1) = rsTmp("ID")
        
        rsTmp.MoveNext
    Loop
    
    txtFile = zlDatabase.GetPara("���������ļ�", 100, 1208, "", Array(txtFile, cmdFile), bln��������)
    strMachine = zlDatabase.GetPara("�ļ���ȡ����", 100, 1208, "", Array(cboDevice), bln��������)

    On Error Resume Next
    If strMachine <> "" And cboDevice.Enabled = True Then
        cboDevice.ListIndex = GetComboxIndex(cboDevice, strMachine)
    End If
    i = zlDatabase.GetPara("�ļ���ȡ��Χ", 100, 1208, 0, Array(optRange), bln��������)
    optRange(i).Value = True
    If i = 0 Then '��ȡ����
        dtpStart = zlDatabase.Currentdate: dtpEnd = zlDatabase.Currentdate
    Else
        dtpStart = CDate(zlDatabase.GetPara("�ļ���ȡ��ʼ����", 100, 1208, zlDatabase.Currentdate, Array(dtpStart), bln��������))
        dtpEnd = CDate(zlDatabase.GetPara("�ļ���ȡ��������", 100, 1208, zlDatabase.Currentdate, Array(dtpEnd), bln��������))
    End If
    
    '������Щ������ҵı��浥���Դ�ӡ
    strDepts = zlDatabase.GetPara("ֻ��ָ�����ұ��浥", 100, 1208, "", Array(lvwDept), bln��������)
    gstrSql = "Select a.id,a.����,a.���� From ���ű� A, ��������˵�� B Where A.ID = B.����id And B.�������� In ('�ٴ�', '����','����','���') " & _
            " order by a.���� "
    Set rsTmp = zlDatabase.OpenSQLRecord(gstrSql, Me.Caption)
    With Me.lvwDept
        Do While Not rsTmp.EOF
            Set lItem = .ListItems.Add(1, "A" & rsTmp("id"), rsTmp("����"))
            lItem.SubItems(1) = rsTmp("����")
            If InStr("," & strDepts & ",", "," & rsTmp("id") & ",") > 0 Then
                lItem.Checked = True
            End If
            rsTmp.MoveNext
        Loop
    End With
    
    If strDepts = "" Then
        Call cmdSelectAll_Click
    End If
    
    
    Me.Show 1, frmMain
    
    ShowPara = mblnOK
    Exit Function
DBError:
    If ErrCenter() = 1 Then Resume
    Call SaveErrLog
End Function

Private Sub cbo_Click(Index As Integer)
    
    Me.DTPHisTory.Visible = (cbo(2).Text = "ָ����ʼ����")
End Sub

Private Sub cbo_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cboDevice_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo�ų�ҩ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo����ҩ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cbo����ҩ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cboס��ҩ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cboס��ҩ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub cboס��ҩ_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub


Private Sub chkActLog_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkFinish_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkAutoRefresh_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkCheck_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkComm_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkNoRange_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkPay_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkSample_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkShort_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkTime_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkҩ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub chkҩ��_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub CmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdClearAll_Click()
    Dim intLoop As Integer
    With Me.lvwDept
        For intLoop = 1 To .ListItems.Count
            .ListItems(intLoop).Checked = False
        Next
    End With
End Sub

Private Sub cmdFile_Click()
    On Error GoTo OpenError
    With dlgFile
        .CancelError = True
        .DialogTitle = "��ѡ�����������ļ�"
        .ShowOpen
        txtFile = .FileName
    End With
    zlCommFun.PressKey vbKeyTab
    Exit Sub
OpenError:
    txtFile.SetFocus
End Sub

Private Sub CmdHelp_Click()
    ShowHelp App.ProductName, Me.hWnd, Me.Name, Int((glngSys) / 100)
End Sub

Private Sub cmdIC_Click()
    Dim objIC As Object
    Set objIC = CreateObject("zlICCard.clsICCard")
    If Not objIC Is Nothing Then
        Call objIC.Set_Card
        Set objIC = Nothing
    End If
    
End Sub

Private Sub cmdIdent_Click()
    Call zlCommFun.DeviceSetup(Me, 100, 1101)
End Sub

Private Sub cmdOK_Click()
    Dim strPar As String, i As Long
    Dim intLoop As Integer
    Dim strDepts As String
    
    zlDatabase.SetPara "�Ǽ�ʱ��ֱ�����벡����Ϣ", chkSample.Value, 100, 1208
    zlDatabase.SetPara "��˴�ӡ", chkPrint.Value, 100, 1208
    zlDatabase.SetPara "������������ʾ������Ŀ", chkShowAll.Value, 100, 1208
    zlDatabase.SetPara "���浥�Ƿ���ʾ��˽��Ŀ", ChkPrivacy.Value, 100, 1208
    zlDatabase.SetPara "���еǼǲ��˱�ʶΪ����", chkPatientType.Value, 100, 1208
    zlDatabase.SetPara "ʹ�ö����������", chkNotSend.Value, 100, 1208
    zlDatabase.SetPara "��˺�������һ������걾", Me.chkSkipRule.Value, 100, 1208
    
    zlDatabase.SetPara "���յǼǱ���ʱǩ��", checkSaveInfoSign.Value, 100, 1208
    zlDatabase.SetPara "���浥����ʱǩ��", checkSaveReprotSign.Value, 100, 1208
        
    If opt���ﴦ��(0).Value = True Then zlDatabase.SetPara "���ﲡ����Ϣ��һ�µĴ���ʽ", 1, 100, 1208
    If opt���ﴦ��(1).Value = True Then zlDatabase.SetPara "���ﲡ����Ϣ��һ�µĴ���ʽ", 2, 100, 1208
    If opt���ﴦ��(2).Value = True Then zlDatabase.SetPara "���ﲡ����Ϣ��һ�µĴ���ʽ", 3, 100, 1208
    
    If optסԺ����(0).Value = True Then zlDatabase.SetPara "סԺ������Ϣ��һ�µĴ���ʽ", 1, 100, 1208
    If optסԺ����(1).Value = True Then zlDatabase.SetPara "סԺ������Ϣ��һ�µĴ���ʽ", 2, 100, 1208
    If optסԺ����(2).Value = True Then zlDatabase.SetPara "סԺ������Ϣ��һ�µĴ���ʽ", 3, 100, 1208
    
    If optԺ�⴦��(0).Value = True Then zlDatabase.SetPara "Ժ�ⲡ����Ϣ��һ�µĴ���ʽ", 1, 100, 1208
    If optԺ�⴦��(1).Value = True Then zlDatabase.SetPara "Ժ�ⲡ����Ϣ��һ�µĴ���ʽ", 2, 100, 1208
    If optԺ�⴦��(2).Value = True Then zlDatabase.SetPara "Ժ�ⲡ����Ϣ��һ�µĴ���ʽ", 3, 100, 1208
    
    If opt��촦��(0).Value = True Then zlDatabase.SetPara "��첡����Ϣ��һ�µĴ���ʽ", 1, 100, 1208
    If opt��촦��(1).Value = True Then zlDatabase.SetPara "��첡����Ϣ��һ�µĴ���ʽ", 2, 100, 1208
    If opt��촦��(2).Value = True Then zlDatabase.SetPara "��첡����Ϣ��һ�µĴ���ʽ", 3, 100, 1208
    
    '----------------------------------------------------------------------------------------
    zlDatabase.SetPara "���μ��鷶Χ", cbo(2).Text, 100, 1208
    zlDatabase.SetPara "�걾������ɹ���", cbo(3).Text, 100, 1208
    zlDatabase.SetPara "���μ��鷶Χָ����ʼ����", Format(Me.DTPHisTory.Value, "yyyy-mm-dd 00:00:00"), 100, 1208
    zlDatabase.SetPara "�Զ�ˢ��", chkAutoRefresh.Value, 100, 1208
    zlDatabase.SetPara "���պ���ʱ��", chkNoRange.Value, 100, 1208
    zlDatabase.SetPara "������ʾ�շ�", chkCheck.Value, 100, 1208
    zlDatabase.SetPara "��������˫��", chkComm.Value, 100, 1208
    zlDatabase.SetPara "����걾", chkEmerge.Value, 100, 1208
    zlDatabase.SetPara "��������Ŀ����", chkCheckAll.Value, 100, 1208
    zlDatabase.SetPara "��ʷ����ʶ��", IIf(Me.optHistoryID.Value, 0, 1), 100, 1208
    zlDatabase.SetPara "����Ӧ��ʾ���", chkShowType.Value, 100, 1208
    zlDatabase.SetPara "���ϴ�����ı걾���ۼ�", chkNO.Value, 100, 1208
    zlDatabase.SetPara "ֻ�ں��յǼ�ʱ��ʾ�ǼǴ���", chkShowOption.Value, 100, 1208
    zlDatabase.SetPara "�Ǽ�ʱ����Ҫ������Ŀ", ChkCheckInNoItem.Value, 100, 1208
    zlDatabase.SetPara "�ֹ���Ŀ����Ŀ�ۼӱ걾��", Me.chkItemNumber.Value, 100, 1208
    zlDatabase.SetPara "ֻ���յ�ǰ������Ŀ", Me.chkOnlyMachine.Value, 100, 1208
    zlDatabase.SetPara "����ʱ��ʾ�ϴγ�����", Me.chkLast.Value, 100, 1208
    zlDatabase.SetPara "�Ǽ�ʱ������һ��������Ŀ", Me.chkLoadLast.Value, 100, 1208
    zlDatabase.SetPara "�Զ����Ӽ�����Ŀ", Me.chkAutoAddItem.Value, 100, 1208
    zlDatabase.SetPara "�ϴν�������ձ걾����", Me.chkSampleType.Value, 100, 1208
    
    If Len(txtFile) > 0 Then
        zlDatabase.SetPara "���������ļ�", txtFile, 100, 1208
    End If
    If cboDevice.ListIndex > -1 Then
        zlDatabase.SetPara "�ļ���ȡ����", cboDevice.ItemData(cboDevice.ListIndex), 100, 1208
    End If
    zlDatabase.SetPara "�ļ���ȡ��Χ", IIf(optRange(0).Value, 0, 1), 100, 1208
    zlDatabase.SetPara "�ļ���ȡ��ʼ����", Format(dtpStart, "yyyy-MM-dd"), 100, 1208
    zlDatabase.SetPara "�ļ���ȡ��������", Format(dtpEnd, "yyyy-MM-dd"), 100, 1208
    '----------------------------------------------------------------------------------------
    With Me.lvwDept
        For intLoop = 1 To .ListItems.Count
            If .ListItems(intLoop).Checked = True Then
                strDepts = strDepts & "," & Mid(.ListItems(intLoop).Key, 2)
            End If
        Next
        Call zlDatabase.SetPara("ֻ��ָ�����ұ��浥", strDepts, 100, 1208)
    End With
    mblnOK = True
    
    Unload Me
End Sub

Private Sub cmdSelectAll_Click()
    Dim intLoop As Integer
    With Me.lvwDept
        For intLoop = 1 To .ListItems.Count
            .ListItems(intLoop).Checked = True
        Next
    End With
End Sub

Private Sub dtpEnd_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub dtpEnd_LostFocus()
    If dtpEnd < dtpStart Then dtpEnd = dtpStart
End Sub

Private Sub dtpStart_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub dtpStart_LostFocus()
    If dtpStart > dtpEnd Then dtpStart = dtpEnd
End Sub

Private Sub lst�շ����_ItemCheck(Item As Integer)
'    If lst�շ����.SelCount = 0 And Not lst�շ����.Selected(Item) Then
'        lst�շ����.Selected(Item) = True
'    End If
End Sub

Private Sub lst�շ����_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optRange_Click(Index As Integer)
    If Index = 0 Then
        dtpStart.Enabled = False
        dtpEnd.Enabled = False
    Else
        dtpStart.Enabled = True
        dtpEnd.Enabled = True
        zlCommFun.PressKey vbKeyTab
    End If
End Sub

Private Sub optRange_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub optҩƷ��λ_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Sub tbs_Click(PreviousTab As Integer)
    tbs.ZOrder 0
End Sub

Private Sub txt_GotFocus(Index As Integer)
'    zlControl.TxtSelAll txt(Index)
End Sub

Private Sub txt_KeyPress(Index As Integer, KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then
        tbs.Tab = 1
'        cbo����ҩ.SetFocus
    End If
End Sub

Private Sub txt_Validate(Index As Integer, Cancel As Boolean)
'    Cancel = Not StrIsValid(txt(Index).Text, txt(Index).MaxLength)
End Sub

Private Sub txtFile_GotFocus()
    With txtFile
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
End Sub

Private Sub txtFile_KeyPress(KeyAscii As Integer)
    If KeyAscii = vbKeyReturn Then zlCommFun.PressKey vbKeyTab
End Sub

Private Function GetComboxIndex(objCbo As ComboBox, ByVal SeekValue As Long) As Long
    Dim i As Long
    
    For i = 0 To objCbo.ListCount - 1
        If objCbo.ItemData(i) = SeekValue Then Exit For
    Next
    If i > objCbo.ListCount - 1 Then i = 0
    GetComboxIndex = i
End Function

